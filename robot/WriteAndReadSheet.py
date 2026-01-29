import os
import csv
import hashlib
import json
import configparser
from pathlib import Path
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials
from typing import Optional


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ✅ Ahora se leen desde config.ini (sección [sheets])
# NOTA: se resuelve en runtime con base_dir (./robot)
DEFAULT_CREDENTIALS_NAME = "credentials.json"

# Columna C => índice 2 (0-based). Aquí la regla es por POSICIÓN C (como venimos haciendo)
PRESTACION_COL_INDEX = 2

# Separador tipo "celdas"
DELIMITER = ';'


def _load_sheets_config(config_path: str) -> tuple[str, str]:
    """
    Lee spreadsheet_id y sheet_name desde config.ini.
    Espera:
      [sheets]
      spreadsheet_id = ...
      sheet_name = ...
    """
    cp = configparser.ConfigParser()
    if not os.path.exists(config_path):
        raise FileNotFoundError(
            f"No encuentro config.ini en: {config_path}. "
            f"Asegúrate de tenerlo en ./robot/config.ini o pasa la ruta explícita."
        )

    cp.read(config_path, encoding="utf-8")

    if "sheets" not in cp:
        raise RuntimeError("Falta la sección [sheets] en config.ini.")

    spreadsheet_id = (cp.get("sheets", "spreadsheet_id", fallback="") or "").strip()
    sheet_name = (cp.get("sheets", "sheet_name", fallback="") or "").strip()

    if not spreadsheet_id:
        raise RuntimeError("Falta 'spreadsheet_id' en la sección [sheets] del config.ini.")
    if not sheet_name:
        raise RuntimeError("Falta 'sheet_name' en la sección [sheets] del config.ini.")

    return spreadsheet_id, sheet_name


def connect(credentials_path: str, config_path: Optional[str] = None):
    """
    Conecta a Google Sheets usando:
    - credentials_path: ruta a credentials.json
    - config_path (opcional): ruta a config.ini
      Si no se pasa, se asume que está en la MISMA carpeta del credentials.json (./robot/config.ini)
    """
    if not os.path.exists(credentials_path):
        raise FileNotFoundError(
            f"No encuentro {credentials_path}. Pon el JSON en ./robot/ o ajusta la ruta."
        )

    if config_path is None:
        config_path = str(Path(credentials_path).resolve().parent / "config.ini")

    spreadsheet_id, sheet_name = _load_sheets_config(config_path)

    creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
    client = gspread.authorize(creds)
    book = client.open_by_key(spreadsheet_id)
    return book.worksheet(sheet_name)


def normalize(s: str) -> str:
    return (s or "").strip()


def contains_mascota(s: str) -> bool:
    return "mascota" in (s or "").lower()


def is_target_row(prestacion_value: str) -> bool:
    """Fila objetivo: Columna "N Prestaciones" debe decir Pendiente."""
    v = normalize(prestacion_value).lower()
    return v == "pendiente"


def ensure_daily_folder(servicios_dir: str) -> str:
    """Crea ./robot/servicios/YYYY-MM-DD y devuelve el path."""
    today = datetime.now().strftime("%Y-%m-%d")
    folder = os.path.join(servicios_dir, today)
    os.makedirs(folder, exist_ok=True)
    return folder


def find_col_index(headers, candidates):
    """
    Busca un índice de columna por posibles nombres.
    candidates: lista de strings (nombres posibles)
    """
    norm_headers = [normalize(h).lower() for h in headers]
    for cand in candidates:
        c = cand.lower()
        if c in norm_headers:
            return norm_headers.index(c)
    return None


def make_mascota_id(row_values, salt="PISCO", used_ids=None) -> str:
    """
    Genera un M#### determinístico usando hash (0..9999).
    Evita duplicados dentro del CSV usando used_ids.
    """
    base = "|".join([normalize(x) for x in row_values])
    h = hashlib.sha256((salt + "|" + base).encode("utf-8")).hexdigest()

    # Primer intento
    n = int(h[:8], 16) % 10000
    candidate = f"M{n:04d}"

    if used_ids is None:
        return candidate

    if candidate not in used_ids:
        return candidate

    for i in range(1, 20000):
        hi = hashlib.sha256((h + f"|{i}").encode("utf-8")).hexdigest()
        n2 = int(hi[:8], 16) % 10000
        candidate2 = f"M{n2:04d}"
        if candidate2 not in used_ids:
            return candidate2

    for n3 in range(10000):
        candidate3 = f"M{n3:04d}"
        if candidate3 not in used_ids:
            return candidate3

    raise RuntimeError("No hay IDs disponibles M0000..M9999 (se agotaron).")


def export_filtered_to_csv(all_rows, out_csv_path):
    if not all_rows:
        raise RuntimeError("La hoja está vacía (no hay filas).")

    headers = all_rows[0]
    data_rows = all_rows[1:]

    idx_prestacion = find_col_index(headers, [
        "N° Prestacion",
        "N Prestacion",
        "N Prestaciones",
        "N° Prestaciones",
        "N Prestación",
        "N° Prestación",
        "No Prestaciones",
        "No Prestación",
        "Prestaciones",
        "Prestación",
    ])

    if idx_prestacion is None:
        idx_prestacion = PRESTACION_COL_INDEX
    if len(headers) <= idx_prestacion:
        raise RuntimeError(
            f"No existe la columna de Prestaciones (idx={idx_prestacion}) en la hoja. "
            "Revisa el nombre de la cabecera o ajusta PRESTACION_COL_INDEX."
        )

    idx_tipo = find_col_index(headers, ["TIPO", "Tipo"])
    idx_categoria = find_col_index(headers, ["Categoria", "Categoría"])
    idx_clasificacion = find_col_index(headers, ["Clasificacion", "Clasificación"])

    idx_cc_fallecido = find_col_index(headers, [
        "CC: Del Fallecido",
        "CC Del Fallecido",
        "CC Fallecido",
        "Cedula Fallecido",
        "Cédula Fallecido",
        "Documento Fallecido",
        "Documento del Fallecido",
    ])

    if idx_cc_fallecido is None:
        raise RuntimeError(
            "No encontré la columna de 'CC: Del Fallecido' (o equivalente) en cabeceras. "
            "Dime el nombre exacto como aparece en Google Sheets."
        )

    used_ids = set()
    filtered = []
    row_map = {}

    for i, row in enumerate(data_rows, start=2):
        gs_row = i

        if len(row) < len(headers):
            row = row + [""] * (len(headers) - len(row))

        prestacion_val = row[idx_prestacion] if len(row) > idx_prestacion else ""

        if not is_target_row(prestacion_val):
            continue

        # ✅ Regla: dejar la columna "N Prestaciones" vacía en el CSV
        row[idx_prestacion] = ""

        mascota_flag = False
        if idx_tipo is not None and contains_mascota(row[idx_tipo]):
            mascota_flag = True
        if idx_categoria is not None and contains_mascota(row[idx_categoria]):
            mascota_flag = True
        if idx_clasificacion is not None and contains_mascota(row[idx_clasificacion]):
            mascota_flag = True

        cc_val = normalize(row[idx_cc_fallecido])
        if mascota_flag and cc_val == "":
            mid = make_mascota_id(row, used_ids=used_ids)
            row[idx_cc_fallecido] = mid
            used_ids.add(mid)

        rid = hashlib.sha1("|".join([normalize(x) for x in row]).encode("utf-8")).hexdigest()
        row_map.setdefault(rid, []).append(gs_row)

        filtered.append(row)

    if not filtered:
        return 0

    with open(out_csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=DELIMITER)
        writer.writerow(headers)
        writer.writerows(filtered)

    map_path = str(Path(out_csv_path).with_suffix(Path(out_csv_path).suffix + ".map.json"))
    with open(map_path, "w", encoding="utf-8") as mf:
        json.dump(row_map, mf, ensure_ascii=False, indent=2)

    return len(filtered)


def generate_pendientes_csv(base_dir: Path | str) -> Optional[Path]:
    """Genera el CSV de pendientes en ./robot/servicios/YYYY-MM-DD/"""
    base_dir = Path(base_dir).resolve()
    credentials_path = str(base_dir / DEFAULT_CREDENTIALS_NAME)
    servicios_dir = str(base_dir / "servicios")

    # ✅ connect ahora lee spreadsheet_id y sheet_name desde ./robot/config.ini automáticamente
    ws = connect(credentials_path)
    all_rows = ws.get_all_values()

    daily_folder = ensure_daily_folder(servicios_dir=servicios_dir)
    hora = datetime.now().strftime("%H%M%S")
    filename = f"Prestacion_Pendiente_{hora}.csv"
    out_csv_path = os.path.join(daily_folder, filename)

    n = export_filtered_to_csv(all_rows, out_csv_path)

    if n == 0:
        return None

    return Path(out_csv_path).resolve()


def main():
    here = Path(__file__).resolve().parent
    generate_pendientes_csv(base_dir=here)


if __name__ == "__main__":
    main()
