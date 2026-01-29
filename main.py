from __future__ import annotations

import csv
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional

from robot import PISCO
from robot import PISCO_CapturarServicios as PCS
from robot import WriteAndReadSheet as WARS


# ------------------------------------------------------------
# Logging
# ------------------------------------------------------------
def setup_logging(robot_dir: Path) -> None:
    logs_dir = robot_dir / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)

    log_file = logs_dir / "robot62.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )
    logging.getLogger("pywinauto").setLevel(logging.WARNING)


# ------------------------------------------------------------
# CSV helpers
# ------------------------------------------------------------
def _detect_delimiter(sample: str) -> str:
    return ";" if sample.count(";") >= sample.count(",") else ","


def read_csv_dicts(csv_path: Path) -> tuple[List[Dict[str, str]], List[str], str]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(4096)
        delim = _detect_delimiter(sample)
        f.seek(0)
        reader = csv.DictReader(f, delimiter=delim)
        rows = list(reader)
        headers = reader.fieldnames or []
    return rows, headers, delim


def write_csv_dicts(csv_path: Path, rows: List[Dict[str, str]], headers: List[str], delim: str) -> None:
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, delimiter=delim)
        w.writeheader()
        w.writerows(rows)


def find_header(headers: List[str], candidates: List[str]) -> Optional[str]:
    norm = {h.strip().lower(): h for h in headers}
    for c in candidates:
        key = c.strip().lower()
        if key in norm:
            return norm[key]
    for h in headers:
        hl = h.strip().lower()
        for c in candidates:
            if c.strip().lower() in hl:
                return h
    return None


def load_row_map(csv_path: Path) -> Dict[str, List[int]]:
    map_path = Path(str(csv_path) + ".map.json")
    if not map_path.exists():
        raise FileNotFoundError(f"No existe el mapa de filas: {map_path}")
    return json.loads(map_path.read_text(encoding="utf-8"))


def row_id_from_dict(row: Dict[str, str], headers: List[str]) -> str:
    import hashlib

    def norm(x: Optional[str]) -> str:
        return (x or "").strip()

    base = "|".join([norm(row.get(h, "")) for h in headers])
    return hashlib.sha1(base.encode("utf-8")).hexdigest()


# ------------------------------------------------------------
# Main orchestration
# ------------------------------------------------------------
def main() -> None:
    project_dir = Path(__file__).resolve().parent
    robot_dir = project_dir / "robot"
    config_path = robot_dir / "config.ini"

    setup_logging(robot_dir)
    logger = logging.getLogger("Robot62")

    def is_blank(v: str) -> bool:
        return (v or "").strip() == ""

    # Referencias para cerrar al final
    main_win = None
    mig_win = None

    try:
        # 1) Google Sheets -> CSV (solo "Pendiente" en la columna N° Prestacion)
        logger.info("1) Generando CSV desde Google Sheets (solo Pendiente)...")
        csv_path = WARS.generate_pendientes_csv(base_dir=robot_dir)

        if not csv_path:
            logger.info("No hay registros en 'Pendiente'. Finalizando sin ejecutar PISCO.")
            return

        logger.info("✅ CSV generado: %s", csv_path)

        # ------------------------------------------------------------
        # 1.1) Pre-validación:
        # - si falta CC: Del Fallecido => marcar en SHEETS "Falta CC fallecido" y SACAR del CSV
        # - si tiene CC => asegurar N° Prestacion VACÍO en CSV (para que PISCO lo cargue)
        # ------------------------------------------------------------
        rows0, headers0, delim0 = read_csv_dicts(Path(csv_path))

        col_prest0 = find_header(headers0, ["N° Prestacion", "N Prestaciones", "Prestaciones", "Prestación", "Prestacion"])
        if not col_prest0:
            raise RuntimeError(f"No encontré la columna de Prestación/N° Prestacion en el CSV. Headers={headers0}")

        col_cc0 = find_header(headers0, [
            "CC: Del Fallecido",
            "CC Del Fallecido",
            "CC Fallecido",
            "Cedula Fallecido",
            "Cédula Fallecido",
            "Documento Fallecido",
            "Documento del Fallecido",
        ])
        if not col_cc0:
            raise RuntimeError(f"No encontré la columna de cédula del fallecido en el CSV. Headers={headers0}")

        row_map0 = load_row_map(Path(csv_path))

        ws0 = WARS.connect(str(robot_dir / WARS.DEFAULT_CREDENTIALS_NAME))
        sheet_all0 = ws0.get_all_values()
        sheet_headers0 = sheet_all0[0] if sheet_all0 else []

        idx_prest_sheet0 = None
        for cand in ["N° Prestacion", "N° Prestaciones", "N Prestaciones", "Prestaciones", "Prestación", "Prestacion"]:
            if cand in sheet_headers0:
                idx_prest_sheet0 = sheet_headers0.index(cand) + 1
                break
        if idx_prest_sheet0 is None:
            for i, h in enumerate(sheet_headers0, start=1):
                if "prest" in (h or "").strip().lower():
                    idx_prest_sheet0 = i
                    break
        if idx_prest_sheet0 is None:
            raise RuntimeError("No encontré en Google Sheets la columna 'N° Prestacion' (o equivalente).")

        marca_cc = "Falta CC fallecido"
        updates0: List[tuple[int, int, str]] = []  # (row, col, value)
        valid_rows: List[Dict[str, str]] = []

        for r in rows0:
            cc = (r.get(col_cc0, "") or "").strip()
            rid = row_id_from_dict(r, headers0)  # rid antes de modificar fila

            if is_blank(cc):
                gs_rows = (row_map0.get(rid) or [])
                if gs_rows:
                    gs_row = gs_rows.pop(0)
                    updates0.append((gs_row, idx_prest_sheet0, marca_cc))
                continue

            # Si tiene CC => para PISCO el N° Prestacion debe ir vacío
            r[col_prest0] = ""
            valid_rows.append(r)

        write_csv_dicts(Path(csv_path), valid_rows, headers0, delim0)
        logger.info(
            "✅ CSV preparado para PISCO: filas_validas=%s | descartadas_sin_CC=%s",
            len(valid_rows),
            len(rows0) - len(valid_rows),
        )

        if updates0:
            min_row = min(r for r, _, _ in updates0)
            max_row = max(r for r, _, _ in updates0)

            cells = ws0.range(min_row, idx_prest_sheet0, max_row, idx_prest_sheet0)
            by_row = {r: v for r, _, v in updates0}

            for cell in cells:
                if cell.row in by_row:
                    cell.value = by_row[cell.row]

            ws0.update_cells(cells, value_input_option="USER_ENTERED")
            logger.info("✅ Google Sheets marcado 'Falta CC fallecido' en %s filas.", len(by_row))

        if not valid_rows:
            logger.info("No quedan filas con CC válido. Finalizando sin abrir PISCO.")
            return

        # ------------------------------------------------------------
        # 2) PISCO: cargar CSV + Guardar Masivo
        # ------------------------------------------------------------
        logger.info("2) Abriendo PISCO e iniciando sesión...")
        main_win = PISCO.open_and_login(config_path=str(config_path))

        logger.info("3) Abriendo Migración Servicios desde Excel...")
        mig_win = PISCO.open_migracion(main_win)

        logger.info("4) Cargando CSV en PISCO: %s", csv_path)
        res_carga = PISCO.cargar_csv(mig_win, str(csv_path))

        cargados = int(res_carga.get("cargados", 0) or 0)
        invalidos = int(res_carga.get("invalidos", 0) or 0)
        logger.info("Resultado carga: cargados=%s, invalidos=%s", cargados, invalidos)

        csv_to_use = Path(csv_path)

        if cargados == 0:
            logger.info("0 registros cargados. Cierro 'Datos' y cierro Migración para poder abrir Capturar Servicios.")
            try:
                PISCO.cerrar_datos_si_aparece(timeout=3.0)
            except Exception:
                pass
            try:
                PISCO.cerrar_ventana(mig_win, timeout=6.0)
                mig_win = None
            except Exception:
                pass
            try:
                main_win.set_focus()
            except Exception:
                pass
        else:
            logger.info("5) Guardar Masivo...")
            res_guardar = PISCO.guardar_masivo(mig_win)
            tiene_errores = bool(res_guardar.get("tiene_errores"))

            if tiene_errores:
                logger.info("6) Guardar Masivo terminó CON errores. Capturando errores desde ventana Datos...")
                res = PISCO.capturar_errores_desde_datos(str(csv_path))
                errores_out = Path(res.get("csv_salida") or "")
                if errores_out.exists():
                    logger.info("✅ CSV con errores: %s", errores_out)
                    csv_to_use = errores_out
            else:
                logger.info("6) Guardar Masivo terminó sin errores.")

        # ------------------------------------------------------------
        # 3) Capturar No Orden Servicio + actualizar Sheet
        # ------------------------------------------------------------
        logger.info("7) Preparando captura de No Orden Servicio...")
        rows, headers, delim = read_csv_dicts(csv_to_use)

        col_prest = find_header(headers, ["N° Prestacion", "N Prestaciones", "Prestaciones", "Prestación", "Prestacion"])
        if not col_prest:
            raise RuntimeError(f"No encontré la columna de Prestación/N° Prestacion en el CSV. Headers={headers}")

        col_cc = find_header(headers, [
            "CC: Del Fallecido",
            "CC Del Fallecido",
            "CC Fallecido",
            "Cedula Fallecido",
            "Cédula Fallecido",
            "Documento Fallecido",
            "Documento del Fallecido",
        ])
        if not col_cc:
            raise RuntimeError(f"No encontré la columna de cédula del fallecido en el CSV. Headers={headers}")

        def is_ok(r: Dict[str, str]) -> bool:
            v = (r.get(col_prest, "") or "").strip().lower()
            return v not in ("error", "falta cc fallecido", "cedula no registrada")

        ok_rows = [r for r in rows if is_ok(r) and not is_blank((r.get(col_cc, "") or "").strip())]

        if not ok_rows:
            logger.info("No hay filas OK para consultar No Orden Servicio.")
            return

        # Asegurar que Migración esté cerrada antes de abrir Capturar Servicios
        if mig_win is not None:
            try:
                PISCO.cerrar_ventana(mig_win, timeout=6.0)
            except Exception:
                pass
            mig_win = None

        try:
            main_win.set_focus()
        except Exception:
            pass

        logger.info("8) Abriendo menú: Archivo -> Capturar Servicios...")
        PCS.capturar_servicios_desde_menu(main_win)

        row_map = load_row_map(Path(csv_path))  # siempre el del CSV original

        ws = WARS.connect(str(robot_dir / WARS.DEFAULT_CREDENTIALS_NAME))
        sheet_all = ws.get_all_values()
        sheet_headers = sheet_all[0] if sheet_all else []

        idx_prest_sheet = None
        for cand in ["N° Prestacion", "N° Prestaciones", "N Prestaciones", "Prestaciones", "Prestación", "Prestacion"]:
            if cand in sheet_headers:
                idx_prest_sheet = sheet_headers.index(cand) + 1
                break
        if idx_prest_sheet is None:
            for i, h in enumerate(sheet_headers, start=1):
                if "prest" in (h or "").strip().lower():
                    idx_prest_sheet = i
                    break
        if idx_prest_sheet is None:
            raise RuntimeError("No encontré en Google Sheets la columna 'N° Prestacion' (o equivalente).")

        updates: List[tuple[int, int, str]] = []

        logger.info("9) Consultando No Orden Servicio para %s registros...", len(ok_rows))
        for r in ok_rows:
            cedula = (r.get(col_cc, "") or "").strip()
            if not cedula:
                continue

            rid0 = row_id_from_dict(r, headers)  # rid antes de cambiar N° Prestacion

            try:
                out = PCS.buscar_por_cedula_fallecido(main_win=main_win, cedula=cedula)
            except Exception as e:
                logger.warning("Cédula=%s -> fallo captura: %s", cedula, e)
                continue

            if not out.get("ok") and out.get("motivo") == "NO_ENCONTRADO":
                marca = "Cedula no registrada"
                r[col_prest] = marca

                gs_rows = (row_map.get(rid0) or [])
                if gs_rows:
                    gs_row = gs_rows.pop(0)
                    updates.append((gs_row, idx_prest_sheet, marca))
                else:
                    logger.warning("No pude mapear fila a Google Sheets (cedula=%s).", cedula)
                continue

            if not out.get("ok"):
                continue

            no_orden = (out.get("no_orden_servicio") or "").strip()
            if not no_orden:
                continue

            r[col_prest] = no_orden

            gs_rows = (row_map.get(rid0) or [])
            if gs_rows:
                gs_row = gs_rows.pop(0)
                updates.append((gs_row, idx_prest_sheet, no_orden))
            else:
                logger.warning("No pude mapear fila a Google Sheets (cedula=%s).", cedula)

        write_csv_dicts(csv_to_use, rows, headers, delim)
        logger.info("✅ CSV actualizado con No Orden Servicio: %s", csv_to_use)

        if updates:
            min_row = min(r for r, _, _ in updates)
            max_row = max(r for r, _, _ in updates)

            cells = ws.range(min_row, idx_prest_sheet, max_row, idx_prest_sheet)
            by_row = {r: v for r, _, v in updates}

            for cell in cells:
                if cell.row in by_row:
                    cell.value = by_row[cell.row]

            ws.update_cells(cells, value_input_option="USER_ENTERED")
            logger.info("✅ Google Sheets actualizado (col=%s) en %s filas.", idx_prest_sheet, len(by_row))
        else:
            logger.info("No hubo actualizaciones para Google Sheets.")

        logger.info("=== Robot62 finalizado OK ===")

    finally:
        # ------------------------------------------------------------
        # CIERRE FINAL (SIEMPRE)
        # ------------------------------------------------------------
        try:
            if mig_win is not None:
                try:
                    PISCO.cerrar_ventana(mig_win, timeout=4.0)
                except Exception:
                    pass
        except Exception:
            pass

        # Cierre suave del main
        try:
            if main_win is not None:
                try:
                    PISCO.cerrar_ventana(main_win, timeout=6.0)
                except Exception:
                    pass
        except Exception:
            pass

        # Cierre duro: taskkill por exe_path del config.ini (best-effort)
        try:
            cfg = PISCO.load_config(str(config_path))
            try:
                # función interna pero efectiva
                PISCO._kill_pisco_processes(cfg.exe_path)
            except Exception:
                pass
        except Exception:
            pass

        # Cerrar cualquier dialog colgado
        try:
            PISCO._close_any_dialogs(timeout=2.0)
        except Exception:
            pass


if __name__ == "__main__":
    main()
