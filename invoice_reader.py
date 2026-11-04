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

_READ_LOCK = threading.Lock()
MAX_BYTES = 10 * 1024 * 1024


def read_invoice(data):
    if not data.startswith(b"%PDF-") or len(data) > MAX_BYTES:
        raise ValueError("Choisissez un PDF de 10 Mo maximum.")
    if not _READ_LOCK.acquire(blocking=False):
        raise ValueError("Une facture est déjà en cours de lecture. Réessayez dans un instant.")
    try:
        with tempfile.TemporaryDirectory(prefix="invoice-") as directory:
            path = Path(directory) / "source.pdf"
            path.write_bytes(data)
            process = subprocess.Popen(
                [sys.executable, str(Path(__file__).with_name("invoice_worker.py")), str(path)],
                stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                start_new_session=(os.name == "posix"),
                env={**os.environ, "OMP_THREAD_LIMIT": "1"},
            )
            try:
                output, _ = process.communicate(timeout=90)
            except subprocess.TimeoutExpired:
                if os.name == "posix":
                    os.killpg(process.pid, signal.SIGKILL)
                else:
                    process.kill()
                process.communicate()
                raise ValueError("Lecture trop longue. Essayez un PDF plus court ou plus léger.")
            if process.returncode or len(output) > 500_000:
                raise ValueError("Lecture interrompue : fichier complexe, endommagé ou limite de ressources atteinte.")
            result = json.loads(output)
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
    # Plusieurs lignes d'une même plage peuvent signaler une régularisation ou un changement de prix.
    candidates = {r["series"]: r for r in rows
                  if sum(other["series"] == r["series"] for other in rows) == 1
                  and 0 <= r["price"] <= 5}
    return {"rows": rows, "candidates": candidates}
