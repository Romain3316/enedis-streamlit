"""Lecture isolée des factures : aucun PDF ni résultat conservé globalement."""
import json
import os
from pathlib import Path
import re
import signal
import subprocess
import sys
import tempfile
import threading
import time

_READ_LOCK = threading.Lock()
MAX_BYTES = 10 * 1024 * 1024


def worker_failure(returncode, stderr):
    """Classify stderr without displaying document text, filenames or tracebacks."""
    detail = stderr.decode("utf-8", errors="replace")
    if "can't open file" in detail:
        reason = "Le module de lecture est introuvable sur le serveur. Le déploiement est à vérifier."
    elif "ModuleNotFoundError" in detail or "ImportError" in detail:
        reason = "Une dépendance du module de lecture ne se charge pas sur le serveur."
    elif "MemoryError" in detail or "Cannot allocate memory" in detail:
        reason = "Le processus de lecture n’a pas pu allouer la mémoire nécessaire."
    elif returncode == -9:
        reason = "Le serveur a arrêté le processus de lecture (SIGKILL). La cause, notamment une éventuelle limite mémoire, reste à vérifier dans les logs."
    elif returncode in (-11, -6):
        reason = "Le moteur de lecture a subi un arrêt natif inattendu sur le serveur."
    else:
        reason = "Le processus de lecture n’a pas terminé normalement."
    return f"{reason} Diagnostic OCR : sortie {returncode}."


def log_reader_state(stage):
    counters = {"stage": stage}
    for name in ("memory.current", "memory.max", "memory.events"):
        try:
            counters[name] = (Path("/sys/fs/cgroup") / name).read_text().strip()
        except OSError:
            pass
    print("[invoice_memory] " + json.dumps(counters), flush=True)


def read_invoice(data):
    suffix = (".pdf" if data.startswith(b"%PDF-") else
              ".png" if data.startswith(b"\x89PNG\r\n\x1a\n") else
              ".jpg" if data.startswith(b"\xff\xd8\xff") else None)
    if suffix is None or len(data) > MAX_BYTES:
        raise ValueError("Choisissez un PDF, JPG ou PNG de 10 Mo maximum.")
    if not _READ_LOCK.acquire(blocking=False):
        raise ValueError("Une facture est déjà en cours de lecture. Réessayez dans un instant.")
    try:
        with tempfile.TemporaryDirectory(prefix="invoice-") as directory:
            path = Path(directory) / ("source" + suffix)
            path.write_bytes(data)
            log_reader_state("before_read")
            started = time.monotonic()
            process = subprocess.Popen(
                [sys.executable, str(Path(__file__).resolve().with_name("invoice_worker.py")), str(path)],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                start_new_session=(os.name == "posix"),
                env={**os.environ, "OMP_THREAD_LIMIT": "1", "OMP_NUM_THREADS": "1",
                     "OPENBLAS_NUM_THREADS": "1", "MKL_NUM_THREADS": "1",
                     "PYTHONIOENCODING": "utf-8"},
            )
            try:
                output, stderr = process.communicate(timeout=90)
            except subprocess.TimeoutExpired:
                if os.name == "posix":
                    os.killpg(process.pid, signal.SIGKILL)
                else:
                    process.kill()
                process.communicate()
                raise ValueError("Lecture trop longue. Essayez un PDF plus court ou plus léger.")
            if process.returncode:
                log_reader_state("worker_failed")
                # Accept only fixed markers: never print raw OCR stderr or filenames.
                stages = re.findall(rb"(?m)^\[ocr_stage\] (worker_started|limits_set|image_decode|image_import|image_open|image_resize|image_orientation|image_grayscale|image_save|pdfinfo|pdftotext|pdftoppm|tesseract)$", stderr)
                stage = stages[-1].decode("ascii") if stages else "before_worker_marker"
                elapsed = round(time.monotonic() - started, 1)
                message = worker_failure(process.returncode, stderr)
                message += f" Étape : {stage} ; durée : {elapsed} s."
                print("[invoice_reader] " + message, flush=True)
                raise ValueError(message)
            if len(output) > 500_000:
                raise ValueError("Le texte extrait dépasse la limite du test (500 ko).")
            try:
                result = json.loads(output)
            except (ValueError, UnicodeError):
                raise ValueError("Réponse du moteur de lecture invalide. Diagnostic OCR : réponse JSON absente ou illisible.")
            if "error" in result:
                raise ValueError(result["error"])
            return {**result, **parse_invoice_text(result["text"])}
    finally:
        _READ_LOCK.release()


def parse_invoice_text(text):
    """Premier format testé : lignes Endesa. Ne jamais déduire les tarifs absents."""
    number = r"\d+(?:[ \u00a0\u202f]\d{3})*(?:[,.]\d+)?"
    date = r"\d{2}/\d{2}/\d{4}"
    rows = []
    for line in text.splitlines():
        match = re.search(
            rf"\b(HPE|HCE|HPH|HCH)\s+du\s+({date})\s+au\s+({date})\s+"
            rf"({number})\s*(MWh|kWh)\s+({number})\s*€\s*/\s*(MWh|kWh)",
            line, re.I,
        )
        if not match:
            continue
        series, start, end, quantity, unit, rate, rate_unit = match.groups()
        to_float = lambda v: float(re.sub(r"[ \u00a0\u202f]", "", v).replace(",", "."))
        rows.append(dict(series=series.upper(), start=start, end=end,
                         kwh=to_float(quantity) * (1000 if unit.lower() == "mwh" else 1),
                         price=to_float(rate) / (1000 if rate_unit.lower() == "mwh" else 1),
                         evidence=line.strip()))
    # EDF : conserver le lien explicite entre libellé, période, quantité et prix.
    # Ne pas assembler des colonnes OCR séparées par simple ordre d'apparition.
    if re.search(r"\bEDF\b", text, re.I):
        for line in text.splitlines():
            match = re.search(
                rf"Heures\s+(pleines|creuses)\s+(?:du\s+)?({date})\s+au\s+({date})\s+"
                rf"({number})\s*kWh\s+({number})\s*(c€|cEUR|cts|€|EUR)\s*/\s*kWh",
                line, re.I)
            if match:
                label, start, end, quantity, rate, unit = match.groups()
                convert = lambda v: float(re.sub(r"[ \u00a0\u202f]", "", v).replace(",", "."))
                rows.append(dict(series="HP" if label.lower() == "pleines" else "HC",
                                 start=start, end=end, kwh=convert(quantity),
                                 price=convert(rate) / (100 if unit.lower().startswith("c") else 1),
                                 evidence=line.strip()))
    # Plusieurs lignes d'une même plage peuvent signaler une régularisation ou un changement de prix.
    candidates = {r["series"]: r for r in rows
                  if sum(other["series"] == r["series"] for other in rows) == 1
                  and 0 <= r["price"] <= 5}
    return {"rows": rows, "candidates": candidates}
