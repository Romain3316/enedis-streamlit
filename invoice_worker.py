"""Processus court et borné pour Poppler/Tesseract, appelé uniquement sur demande."""
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys


def mark(stage):
    print("[ocr_stage] " + stage, file=sys.stderr, flush=True)


def command(args, timeout=20):
    tool = Path(args[0]).name
    mark(tool)
    try:
        return subprocess.run(args, check=True, capture_output=True, timeout=timeout).stdout.decode("utf-8", errors="replace")
    except subprocess.CalledProcessError as exc:
        raise ValueError(f"Le moteur {tool} a échoué (sortie {exc.returncode}). Vérifiez l’installation OCR et le document.") from None
    except subprocess.TimeoutExpired:
        raise ValueError(f"Le moteur {tool} a dépassé le délai de {timeout} secondes.") from None


def extract(path):
    if path.suffix.lower() in (".jpg", ".png"):
        if not shutil.which("tesseract"):
            raise ValueError("Le moteur OCR n'est pas disponible sur ce serveur.")
        mark("image_decode")
        from PIL import Image, ImageOps
        # Bound decoded memory before loading; EXIF orientation is handled locally.
        with Image.open(path) as original:
            if original.format not in ("JPEG", "PNG") or original.width * original.height > 20_000_000:
                raise ValueError("Photo trop grande : maximum 20 millions de pixels, en JPG ou PNG.")
            original.draft("RGB", (2400, 2400))
            with ImageOps.exif_transpose(original) as oriented:
                oriented.thumbnail((2400, 2400))
                with ImageOps.grayscale(oriented) as gray:
                    image_path = path.parent / "ocr.png"
                    gray.save(image_path)
        languages = command(["tesseract", "--list-langs"])
        language = "fra" if "fra" in languages.splitlines() else "eng"
        text = command(["tesseract", str(image_path), "stdout", "-l", language])
        if len(text) > 30_000:
            raise ValueError("Image trop dense pour cet espace de test.")
        return {"text": text, "ocr_pages": [1], "pages": 1}
    for tool in ("pdfinfo", "pdftotext"):
        if not shutil.which(tool):
            raise ValueError("Lecture PDF indisponible sur ce serveur (Poppler manquant).")
    info = command(["pdfinfo", str(path)])
    count = re.search(r"Pages:\s*(\d+)", info)
    if not count or not 1 <= int(count[1]) <= 6:
        raise ValueError("Le test accepte de 1 à 6 pages. Déposez une facture à la fois.")
    pages, scanned = [], []
    for page in range(1, int(count[1]) + 1):
        text = command(["pdftotext", "-f", str(page), "-l", str(page), "-layout", str(path), "-"])
        if len(re.sub(r"\s", "", text)) < 80:
            if not shutil.which("tesseract") or not shutil.which("pdftoppm"):
                raise ValueError("Ce PDF est scanné, mais le moteur OCR n'est pas disponible sur le serveur.")
            languages = command(["tesseract", "--list-langs"])
            language = "fra" if "fra" in languages.splitlines() else "eng"
            prefix = path.parent / "page"
            command(["pdftoppm", "-f", str(page), "-l", str(page), "-scale-to", "2000",
                     "-png", "-singlefile", str(path), str(prefix)])
            text = command(["tesseract", str(prefix) + ".png", "stdout", "-l", language])
            Path(str(prefix) + ".png").unlink(missing_ok=True)
            scanned.append(page)
        if len(text) > 30_000:
            raise ValueError("Page trop volumineuse pour cet espace de test.")
        pages.append(f"--- Page {page} ---\n{text}")
    return {"text": "\n".join(pages), "ocr_pages": scanned, "pages": len(pages)}


if __name__ == "__main__":
    mark("worker_started")
    try:
        if os.name == "posix":
            import resource
            resource.setrlimit(resource.RLIMIT_AS, (512 * 1024 ** 2, 512 * 1024 ** 2))
            resource.setrlimit(resource.RLIMIT_CPU, (60, 60))
            resource.setrlimit(resource.RLIMIT_FSIZE, (20 * 1024 ** 2, 20 * 1024 ** 2))
        mark("limits_set")
        result = extract(Path(sys.argv[1]))
    except ValueError as exc:
        result = {"error": str(exc)}
    except MemoryError:
        result = {"error": "Le module de lecture a atteint sa limite mémoire de travail. Diagnostic OCR : MemoryError."}
    except Exception as exc:
        result = {"error": f"Le document n’a pas pu être lu. Diagnostic OCR : {type(exc).__name__}."}
    print(json.dumps(result, ensure_ascii=False))
