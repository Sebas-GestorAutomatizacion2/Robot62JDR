import os
import csv
import hashlib
import json
from pathlib import Path
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials
from typing import Optional


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_ID = "1SIRomJ50ewG0XtD1sa7MiBkzcJB0mAfdW0EitBcAYnk"
SHEET_NAME = "Prestacion"

# NOTA: se resuelve en runtime con base_dir (./robot)
DEFAULT_CREDENTIALS_NAME = "credentials.json"

# Columna C => índice 2 (0-based). Aquí la regla es por POSICIÓN C (como venimos haciendo)
PRESTACION_COL_INDEX = 2

# Separador tipo "celdas"
DELIMITER = ';'


def connect(credentials_path: str):
    if not os.path.exists(credentials_path):
        raise FileNotFoundError(
            f"No encuentro {credentials_path}. Pon el JSON en ./robot/ o ajusta la ruta."
        )
    creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
    client = gspread.authorize(creds)
    book = client.open_by_key(SPREADSHEET_ID)
    return book.worksheet(SHEET_NAME)


def normalize(s: str) -> str:
    return (s or "").strip()


def contains_mascota(s: str) -> bool:
    return "mascota" in (s or "").lower()


def is_target_row(prestacion_value: str) -> bool:
    """Fila objetivo: Columna "N Prestaciones" (por defecto Col N) debe decir Pendiente."""
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

    # Resolver colisiones dentro del mismo archivo (muy raro, pero lo cubrimos)
    if candidate not in used_ids:
        return candidate

    # Rehash incremental
    for i in range(1, 20000):
        hi = hashlib.sha256((h + f"|{i}").encode("utf-8")).hexdigest()
        n2 = int(hi[:8], 16) % 10000
        candidate2 = f"M{n2:04d}"
        if candidate2 not in used_ids:
            return candidate2

    # Si llegara a pasar (casi imposible), fallback secuencial
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

    # Validación base (cabeceras)

    # Columna "N Prestaciones" (estado / prestación) por nombre, fallback por posición (Col N).
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


    # Columnas para detectar mascota
    idx_tipo = find_col_index(headers, ["TIPO", "Tipo"])
    idx_categoria = find_col_index(headers, ["Categoria", "Categoría"])
    idx_clasificacion = find_col_index(headers, ["Clasificacion", "Clasificación"])

    # Columna de cédula fallecido (puede variar el nombre; pongo varios candidatos)
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
        # Normalizar longitud
        if len(row) < len(headers):
            row = row + [""] * (len(headers) - len(row))

        prestacion_val = row[idx_prestacion] if len(row) > idx_prestacion else ""

        # Filtro principal (vacío o pendiente)
        if not is_target_row(prestacion_val):
            continue

        # ✅ Regla: dejar la columna "N Prestaciones" vacía en el CSV (PISCO la llenará con ERROR si aplica)
        row[idx_prestacion] = ""

        # Detectar mascota por columnas
        mascota_flag = False
        if idx_tipo is not None and contains_mascota(row[idx_tipo]):
            mascota_flag = True
        if idx_categoria is not None and contains_mascota(row[idx_categoria]):
            mascota_flag = True
        if idx_clasificacion is not None and contains_mascota(row[idx_clasificacion]):
            mascota_flag = True

        # Si es mascota y CC fallecido vacío => asignar M####.
        cc_val = normalize(row[idx_cc_fallecido])
        if mascota_flag and cc_val == "":
            mid = make_mascota_id(row, used_ids=used_ids)
            row[idx_cc_fallecido] = mid
            used_ids.add(mid)

        # Mapear esta fila del CSV a su fila real en Google Sheets (para actualizar luego sin tocar el CSV)
        rid = hashlib.sha1("|".join([normalize(x) for x in row]).encode("utf-8")).hexdigest()
        row_map.setdefault(rid, []).append(gs_row)

        filtered.append(row)

    # ✅ Si no hay filas pendientes, no crear CSV ni map y salir con 0
    if not filtered:
        return 0

    # Guardar CSV
    with open(out_csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=DELIMITER)
        writer.writerow(headers)
        writer.writerows(filtered)

    # Guardar mapa de filas (sidecar) para actualizaciones posteriores
    map_path = str(Path(out_csv_path).with_suffix(Path(out_csv_path).suffix + ".map.json"))
    with open(map_path, "w", encoding="utf-8") as mf:
        json.dump(row_map, mf, ensure_ascii=False, indent=2)

    return len(filtered)


def generate_pendientes_csv(base_dir: Path | str) -> Optional[Path]:
    """Genera el CSV de pendientes en ./robot/servicios/YYYY-MM-DD/

    Returns:
        Path absoluto del CSV creado, o None si no hay pendientes.
    """
    base_dir = Path(base_dir).resolve()
    credentials_path = str(base_dir / DEFAULT_CREDENTIALS_NAME)
    servicios_dir = str(base_dir / "servicios")

    ws = connect(credentials_path)
    all_rows = ws.get_all_values()

    daily_folder = ensure_daily_folder(servicios_dir=servicios_dir)
    hora = datetime.now().strftime("%H%M%S")
    filename = f"Prestacion_Pendiente_{hora}.csv"
    out_csv_path = os.path.join(daily_folder, filename)

    n = export_filtered_to_csv(all_rows, out_csv_path)

    if n == 0:
        # ✅ No hay pendientes -> no crear CSV ni ejecutar nada
        return None

    return Path(out_csv_path).resolve()


def main():
    # Permite ejecutar el módulo solo para generar el CSV
    here = Path(__file__).resolve().parent
    generate_pendientes_csv(base_dir=here)


if __name__ == "__main__":
    main()
