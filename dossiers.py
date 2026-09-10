"""Portable JSON dossiers, with explicit types and no executable serialization."""

import base64
import binascii
import hashlib
import json
import math
from datetime import date, time
from io import BytesIO
from pathlib import PurePath

from dossier_schema import SETTINGS

FORMAT = "cma-analyse-energetique-draft"
MAX_FILE_BYTES = 50 * 1024 * 1024
MAX_DOSSIER_BYTES = 150 * 1024 * 1024
FILE_EXTENSIONS = {"curve": {".csv", ".xlsx", ".xls", ".txt"}, "pma": {".csv", ".xlsx", ".xls"}}


class SourceFile(BytesIO):
    def __init__(self, name, content):
        super().__init__(content)
        self.name = name


def validate_fields(fields):
    if not isinstance(fields, dict):
        raise ValueError("Les réglages du dossier sont invalides.")
    result = {}
    for key, value in fields.items():
        if key not in SETTINGS:
            raise ValueError(f"Champ de dossier non reconnu : {key}")
        spec = SETTINGS[key]
        kind = spec["type"]
        valid = True
        if kind == "date":
            try: value = date.fromisoformat(value) if isinstance(value, str) else value
            except ValueError: valid = False
            valid = valid and type(value) is date
        elif kind == "time":
            try: value = time.fromisoformat(value) if isinstance(value, str) else value
            except ValueError: valid = False
            valid = valid and type(value) is time and value.tzinfo is None
        elif kind in ("int", "float"):
            valid = type(value) in ((int,) if kind == "int" else (int, float))
            valid = valid and math.isfinite(value)
            valid = valid and spec.get("min", -math.inf) <= value <= spec.get("max", math.inf)
            if valid and kind == "float": value = float(value)
        elif kind == "bool": valid = type(value) is bool
        else: valid = isinstance(value, str) and len(value) <= 100000
        if not valid or ("choices" in spec and value not in spec["choices"]):
            raise ValueError(f"Valeur invalide pour {key}.")
        result[key] = value
    if result.get("start_date") and result.get("end_date") and result["start_date"] > result["end_date"]:
        raise ValueError("La date de début doit précéder la date de fin.")
    return result


def _json_default(value):
    if isinstance(value, (date, time)): return value.isoformat()
    raise TypeError(f"Type non pris en charge : {type(value).__name__}")


def _validate_sources(files):
    if not isinstance(files, dict) or set(files) - set(FILE_EXTENSIONS):
        raise ValueError("Liste des fichiers source invalide.")
    checked = {}
    for role, source in files.items():
        if source is None: continue
        if not isinstance(source, dict): raise ValueError("Fichier source invalide.")
        name, content = source.get("name"), source.get("content")
        if not isinstance(name, str) or len(name)>255 or '/' in name or '\\' in name or PurePath(name).suffix.lower() not in FILE_EXTENSIONS[role]:
            raise ValueError("Nom ou format de fichier source invalide.")
        if not isinstance(content, bytes) or len(content) > MAX_FILE_BYTES:
            raise ValueError("Chaque fichier source doit faire au maximum 50 Mo.")
        checked[role] = {"name": name, "content": content}
    return checked


def validate_solar_snapshot(snapshot):
    if snapshot is None: return None
    if not isinstance(snapshot, dict) or not isinstance(snapshot.get("parameters"), dict):
        raise ValueError("Profil solaire invalide.")
    bounds = {"latitude": (-90, 90), "longitude": (-180, 180), "tilt": (0, 90),
              "aspect": (-180, 180), "peak_power_kwp": (0.1, 1000), "losses_percent": (0, 35)}
    params = snapshot["parameters"]
    if set(params) != set(bounds): raise ValueError("Paramètres du profil solaire invalides.")
    for key, (low, high) in bounds.items():
        if type(params[key]) not in (int, float) or not math.isfinite(params[key]) or not low <= params[key] <= high:
            raise ValueError("Paramètres du profil solaire invalides.")
    rows = snapshot.get("rows")
    if not isinstance(rows, list) or not 1 <= len(rows) <= 8784: raise ValueError("Profil solaire vide ou trop volumineux.")
    seen = set()
    for row in rows:
        if not isinstance(row, dict) or set(row) != {"Mois", "Jour_mois", "Heure", "Irradiation_Wm2", "Production_PV_kW"}:
            raise ValueError("Colonnes du profil solaire invalides.")
        try:
            date(2024, row["Mois"], row["Jour_mois"])
            time(row["Heure"])
        except (TypeError, ValueError) as exc: raise ValueError("Date du profil solaire invalide.") from exc
        identity = (row["Mois"], row["Jour_mois"], row["Heure"])
        if identity in seen: raise ValueError("Doublon dans le profil solaire.")
        seen.add(identity)
        for key in ("Irradiation_Wm2", "Production_PV_kW"):
            value = row[key]
            if value is not None and (type(value) not in (int, float) or not math.isfinite(value) or value < 0):
                raise ValueError("Valeur du profil solaire invalide.")
    return {"parameters": params, "rows": rows}


def encode_dossier(fields, files, solar_snapshot=None):
    fields = validate_fields({**{k: s["default"] for k, s in SETTINGS.items() if "default" in s}, **fields})
    files = _validate_sources(files)
    encoded = {
        role: {"name": f["name"], "data": base64.b64encode(f["content"]).decode("ascii"),
               "sha256": hashlib.sha256(f["content"]).hexdigest()}
        for role, f in files.items()
    }
    payload = {"format": FORMAT, "version": 2, "fields": fields, "files": encoded,
               "solar_snapshot": validate_solar_snapshot(solar_snapshot)}
    raw = json.dumps(payload, ensure_ascii=False, allow_nan=False, default=_json_default).encode("utf-8")
    if len(raw) > MAX_DOSSIER_BYTES: raise ValueError("Le dossier dépasse la taille maximale de 150 Mo.")
    return raw


def decode_dossier(raw):
    if len(raw) > MAX_DOSSIER_BYTES: raise ValueError("Le dossier dépasse la taille maximale de 150 Mo.")
    try: payload = json.loads(raw.decode("utf-8-sig"))
    except (ValueError, UnicodeError) as exc: raise ValueError("Le fichier n'est pas un dossier JSON valide.") from exc
    if not isinstance(payload, dict) or payload.get("format") != FORMAT or type(payload.get("version")) is not int or payload["version"] not in (1, 2):
        raise ValueError("Format ou version de dossier non pris en charge.")
    fields = validate_fields(payload.get("fields"))
    files = {}
    if payload["version"] == 2:
        encoded = payload.get("files")
        if not isinstance(encoded, dict) or set(encoded) - set(FILE_EXTENSIONS): raise ValueError("Liste des fichiers source invalide.")
        for role, source in encoded.items():
            try:
                if not isinstance(source, dict) or not isinstance(source.get("data"), str): raise ValueError()
                content = base64.b64decode(source["data"], validate=True)
                if hashlib.sha256(content).hexdigest() != source.get("sha256"): raise ValueError()
            except (ValueError, binascii.Error) as exc: raise ValueError("Fichier source corrompu ou incomplet.") from exc
            files[role] = {"name": source.get("name"), "content": content}
        files = _validate_sources(files)
    return {"version": payload["version"], "fields": fields, "files": files,
            "solar_snapshot": validate_solar_snapshot(payload.get("solar_snapshot"))}


def source_from_record(record):
    return SourceFile(record["name"], record["content"]) if record else None
