# ==============================
# PISCO.py – versión estable
# ==============================

from __future__ import annotations
import time
import ctypes
import logging
import configparser
import pyperclip
from dataclasses import dataclass
from typing import Optional
from pathlib import Path

import win32gui
import win32con
import win32process

import re
import time
import win32gui
import win32con
import os


from pywinauto.timings import TimeoutError

from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
from pywinauto.timings import TimeoutError
from pywinauto.controls.hwndwrapper import HwndWrapper

from pywinauto import Desktop

logger = logging.getLogger("Robot62.PISCO")

import ctypes
from ctypes import wintypes
user32 = ctypes.windll.user32

LVM_FIRST = 0x1000
LVM_GETITEMCOUNT = LVM_FIRST + 4
LVM_GETCOLUMNWIDTH = LVM_FIRST + 29
LVM_GETHEADER = LVM_FIRST + 31
LVM_GETITEMTEXTW = LVM_FIRST + 115

HDM_FIRST = 0x1200
HDM_GETITEMCOUNT = HDM_FIRST + 0

class LVITEMW(ctypes.Structure):
    _fields_ = [
        ("mask", wintypes.UINT),
        ("iItem", wintypes.INT),
        ("iSubItem", wintypes.INT),
        ("state", wintypes.UINT),
        ("stateMask", wintypes.UINT),
        ("pszText", wintypes.LPWSTR),
        ("cchTextMax", wintypes.INT),
        ("iImage", wintypes.INT),
        ("lParam", wintypes.LPARAM),
    ]

LVIF_TEXT = 0x0001


def _close_any_dialogs(timeout: float = 2.0) -> int:
    """
    Cierra dialogs tipo MessageBox (#32770) visibles que a veces quedan colgados.
    Retorna cuántos intentó cerrar (best-effort).
    """
    t0 = time.time()
    closed = 0

    while time.time() - t0 < timeout:
        hwnd = win32gui.FindWindow("#32770", None)
        if not hwnd:
            break
        try:
            # Intentar cerrar suave
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            closed += 1
        except Exception:
            pass
        time.sleep(0.2)

    if closed:
        logger.warning("Preflight: cerré %s dialogs #32770 sueltos.", closed)

    return closed


def _kill_pisco_processes(exe_path: str) -> bool:
    """
    Mata procesos del ejecutable de PISCO por nombre, best-effort.
    Retorna True si intentó matar algo.
    """
    exe_name = Path(exe_path).name  # ej: PISCO.exe
    if not exe_name:
        return False

    try:
        # /T mata procesos hijos, /F forzado
        rc = os.system(f'taskkill /IM "{exe_name}" /T /F >NUL 2>&1')
        if rc == 0:
            logger.warning("Preflight: taskkill ejecutado para %s.", exe_name)
            time.sleep(1.0)
            return True
    except Exception as e:
        logger.warning("Preflight: no pude hacer taskkill: %s", e)

    return False


def cerrar_datos_si_aparece(timeout: float = 2.5) -> bool:
    """
    Si aparece la ventana 'Datos' (por registros inválidos), la cierra.
    Best-effort: no falla si no aparece.
    """
    t0 = time.time()
    desk = Desktop(backend="win32")
    while time.time() - t0 < timeout:
        try:
            w = desk.window(title_re=r"^Datos$")
            if w.exists(timeout=0.2) and w.is_visible():
                try:
                    cerrar_ventana(w.wrapper_object())
                except Exception:
                    try:
                        win32gui.PostMessage(w.handle, win32con.WM_CLOSE, 0, 0)
                    except Exception:
                        pass
                return True
        except Exception:
            pass
        time.sleep(0.15)
    return False


def _find_main_child_control(parent_hwnd: int) -> int:
    """
    Encuentra el control hijo más grande (normalmente el grid).
    """
    best = None
    best_area = 0

    def enum_child(hwnd, _):
        nonlocal best, best_area
        if not win32gui.IsWindowVisible(hwnd):
            return
        try:
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            area = max(0, r - l) * max(0, b - t)
            if area > best_area:
                best_area = area
                best = hwnd
        except Exception:
            pass

    win32gui.EnumChildWindows(parent_hwnd, enum_child, None)
    if not best:
        raise RuntimeError("No pude encontrar el control principal (grid) en Datos.")
    return best

GRID_CLASS_CANDIDATES = {
    # ListView / grids comunes
    "SysListView32",
    "MSFlexGridWndClass",
    "MshFlexGridWndClass",      # a veces
    "VSFlexGridWndClass",
    "VtListView",               # algunos terceros
    "TDBGrid",                  # algunos terceros
    # VB6 containers que pueden envolver otros
    "ThunderRT6UserControlDC",
    "ThunderRT6PictureBox",
}

def _enum_children(hwnd_parent: int) -> list[int]:
    out = []
    def cb(h, _):
        out.append(h)
    try:
        win32gui.EnumChildWindows(hwnd_parent, cb, None)
    except Exception:
        pass
    return out

def _dump_descendants(hwnd_parent: int, limit: int = 80):
    """
    Loggea clases/textos de descendientes para diagnóstico.
    """
    items = []
    seen = set()

    def walk(h, depth=0):
        if h in seen:
            return
        seen.add(h)
        try:
            cls = win32gui.GetClassName(h)
            txt = (win32gui.GetWindowText(h) or "").strip()
            l, t, r, b = win32gui.GetWindowRect(h)
            items.append((depth, h, cls, txt, (r-l, b-t)))
        except Exception:
            return

        for ch in _enum_children(h):
            walk(ch, depth + 1)

    walk(hwnd_parent, 0)

    # ordenar por área (más grande primero) para ver candidatos
    items.sort(key=lambda x: (x[4][0]*x[4][1]), reverse=True)

    logger.info("---- DUMP descendientes (top %s por area) ----", limit)
    for i, (depth, h, cls, txt, (w, hgt)) in enumerate(items[:limit], 1):
        logger.info("%02d) d=%d hwnd=%s cls=%s size=%sx%s txt='%s'",
                    i, depth, h, cls, w, hgt, txt[:60])
    logger.info("---- FIN DUMP ----")

def _find_grid_hwnd(hwnd_root: int) -> int | None:
    """
    Busca recursivamente un control que parezca grid real.
    Evita quedarse con el wrapper ATL:xxxx.
    """
    seen = set()
    best = None
    best_area = 0

    def area_of(h):
        try:
            l, t, r, b = win32gui.GetWindowRect(h)
            return max(0, r-l) * max(0, b-t)
        except Exception:
            return 0

    def walk(h):
        nonlocal best, best_area
        if h in seen:
            return
        seen.add(h)

        try:
            cls = win32gui.GetClassName(h)
        except Exception:
            return

        a = area_of(h)

        # si es candidato, lo consideramos
        if cls in GRID_CLASS_CANDIDATES or cls.startswith("ATL:"):
            # preferimos el MÁS GRANDE que sea candidato
            if a > best_area:
                best_area = a
                best = h

        # seguir bajando
        for ch in _enum_children(h):
            walk(ch)

    walk(hwnd_root)

    # Si lo “mejor” quedó siendo ATL:xxxx, intentamos buscar un hijo NO-ATL debajo de ese wrapper
    if best:
        try:
            cls_best = win32gui.GetClassName(best)
        except Exception:
            cls_best = ""

        if cls_best.startswith("ATL:"):
            # buscar el mejor candidato real debajo del ATL
            inner = None
            inner_area = 0
            for ch in _enum_children(best):
                try:
                    cls = win32gui.GetClassName(ch)
                except Exception:
                    continue
                a = area_of(ch)
                if (cls in GRID_CLASS_CANDIDATES) and a > inner_area:
                    inner_area = a
                    inner = ch
            if inner:
                return inner

    return best
#capturar_errores_desde_datos

def _listview_get_cols(hwnd_lv: int) -> int:
    hdr = user32.SendMessageW(hwnd_lv, LVM_GETHEADER, 0, 0)
    if not hdr:
        return 0
    return user32.SendMessageW(hdr, HDM_GETITEMCOUNT, 0, 0)

def _listview_get_text(hwnd_lv: int, row: int, col: int, maxlen: int = 512) -> str:
    buf = ctypes.create_unicode_buffer(maxlen)
    item = LVITEMW()
    item.mask = LVIF_TEXT
    item.iItem = row
    item.iSubItem = col
    item.pszText = ctypes.cast(buf, wintypes.LPWSTR)
    item.cchTextMax = maxlen
    user32.SendMessageW(hwnd_lv, LVM_GETITEMTEXTW, row, ctypes.byref(item))
    return buf.value

def _read_listview(hwnd_lv: int) -> list[list[str]]:
    n_rows = user32.SendMessageW(hwnd_lv, LVM_GETITEMCOUNT, 0, 0)
    n_cols = _listview_get_cols(hwnd_lv)
    if n_rows <= 0 or n_cols <= 0:
        return []

    data = []
    for r in range(n_rows):
        row = []
        for c in range(n_cols):
            row.append(_listview_get_text(hwnd_lv, r, c).strip())
        data.append(row)
    return data


# ==========================================================
# CONFIG
# ==========================================================

@dataclass
class PiscoConfig:
    exe_path: str
    login_title_re: str
    main_title_re: str
    usuario: str
    contrasena: str
    require_admin: bool = True
    main_load_timeout: int = 240


def load_config(path: str) -> PiscoConfig:
    cfg = configparser.ConfigParser()
    if not cfg.read(path, encoding="utf-8"):
        raise FileNotFoundError(path)

    return PiscoConfig(
        exe_path=cfg["app"]["exe_path"],
        login_title_re=cfg["app"]["title_re"],
        main_title_re=cfg["app"]["main_title_re"],
        usuario=cfg["login"]["usuario"],
        contrasena=cfg["login"]["contrasena"],
        require_admin=cfg.getboolean("app", "require_admin", fallback=True),
        main_load_timeout=cfg.getint("app", "main_load_timeout", fallback=240),
    )


def _is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


# ==========================================================
# LOGIN
# ==========================================================

def open_and_login(config_path: str):
    cfg = load_config(config_path)

    # --------------------------------------------------
    # 0) PRE-FLIGHT: limpiar estado viejo (SIN REINICIAR PC)
    # --------------------------------------------------
    _close_any_dialogs(timeout=2.0)
    _kill_pisco_processes(cfg.exe_path)
    _close_any_dialogs(timeout=2.0)

    # --------------------------------------------------
    # 1) Lanzar PISCO
    # --------------------------------------------------
    if cfg.require_admin and not _is_admin():
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", cfg.exe_path, None, None, 1
        )
    else:
        Application(backend="win32").start(f'"{cfg.exe_path}"')


    desk = Desktop(backend="win32")

    # --------------------------------------------------
    # 2) Detectar ventana de login (ROBUSTO)
    # --------------------------------------------------
    login_patterns = [
        cfg.login_title_re,
        r"(?i).*jardines.*renacer.*",
        r"(?i).*prueba.*jardines.*",
    ]

    login = None
    t0 = time.time()

    while time.time() - t0 < 90:
        for pat in login_patterns:
            try:
                w = desk.window(title_re=pat)
                if w.exists(timeout=0.3) and w.is_visible():
                    login = w.wrapper_object()
                    break
            except Exception:
                continue
        if login:
            break
        time.sleep(0.4)

    if not login:
        raise TimeoutError(
            "PISCO abrió, pero no pude detectar la ventana de login "
            f"(patrones probados: {login_patterns})"
        )

    logger.info("Ventana de login detectada: %s", login.window_text())

    # --------------------------------------------------
    # 3) Llenar credenciales
    # --------------------------------------------------
    login.set_focus()
    time.sleep(0.3)

    edits = login.children(class_name="Edit")
    if len(edits) >= 2:
        edits[0].set_edit_text(cfg.usuario)
        edits[1].set_edit_text(cfg.contrasena)
        send_keys("{ENTER}{ENTER}")
    else:
        # fallback teclado
        send_keys("^a{BACKSPACE}")
        send_keys(cfg.usuario)
        send_keys("{TAB}")
        send_keys("^a{BACKSPACE}")
        send_keys(cfg.contrasena)
        send_keys("{ENTER}{ENTER}")

    # --------------------------------------------------
    # 4) Esperar ventana principal
    # --------------------------------------------------
    t1 = time.time()
    while time.time() - t1 < cfg.main_load_timeout:
        try:
            main = desk.window(title_re=cfg.main_title_re)
            if main.exists(timeout=0.5) and main.is_visible():
                logger.info("Ventana principal detectada.")
                return main.wrapper_object()
        except Exception:
            pass
        time.sleep(0.5)

    raise TimeoutError("Login enviado, pero no apareció la ventana principal.")



# ==========================================================
# MIGRACIÓN DESDE EXCEL
# ==========================================================

def open_migracion(main_win):
    main_win.set_focus()
    main_win.menu_select("Operaciones->Migracion Servicios desde Excel")

    desk = Desktop(backend="win32")
    dlg = desk.window(title_re=r".*Migraci[oó]n\s+Servicios.*")
    dlg.wait("visible", timeout=60)
    return dlg.wrapper_object()


# ==========================================================
# VB6 – HELPERS DUROS
# ==========================================================

def _vb6_button(parent_hwnd: int, label: str) -> int | None:
    """
    Retorna HWND SOLO si el botón:
    - es ThunderRT6CommandButton
    - texto exacto
    - visible
    - habilitado
    """

    result = None

    def enum_child(hwnd, _):
        nonlocal result
        if result:
            return
        if win32gui.GetClassName(hwnd) != "ThunderRT6CommandButton":
            return
        if win32gui.GetWindowText(hwnd).strip().lower() != label.lower():
            return
        if not win32gui.IsWindowVisible(hwnd):
            return
        if not win32gui.IsWindowEnabled(hwnd):
            return
        result = hwnd

    win32gui.EnumChildWindows(parent_hwnd, enum_child, None)
    return result


def _bm_click(hwnd: int):
    win32gui.PostMessage(hwnd, win32con.BM_CLICK, 0, 0)


# ==========================================================
# CARGAR CSV
# ==========================================================

def cargar_csv(mig_win, csv_path: str):
    btn = _vb6_button(mig_win.handle, "Cargar Archivo")
    if not btn:
        raise RuntimeError("No se encontró 'Cargar Archivo'")

    _bm_click(btn)

    desk = Desktop(backend="win32")
    dlg = desk.window(class_name="#32770")
    dlg.wait("visible", timeout=30)
    dlg.set_focus()

    send_keys("%n")
    send_keys("^a{BACKSPACE}")
    send_keys(csv_path, with_spaces=True)
    send_keys("{ENTER}")

    return _wait_result_popup()




def _wait_result_popup(timeout: int = 60) -> dict:
    """
    Detecta el MessageBox VB6 de carga:
      'Archivo procesado correctamente.'
      'Registros cargados: X'
      'Registros inválidos: Y'
    Cierra el popup y retorna dict con contadores.
    """
    t0 = time.time()

    while time.time() - t0 < timeout:
        hwnd = win32gui.FindWindow("#32770", None)
        if hwnd:
            textos = []

            def enum_child(h, _):
                if win32gui.GetClassName(h) == "Static":
                    txt = (win32gui.GetWindowText(h) or "").strip()
                    if txt:
                        textos.append(txt)

            try:
                win32gui.EnumChildWindows(hwnd, enum_child, None)
            except Exception:
                textos = []

            full = "\n".join(textos)
            full_l = full.lower()

            if "archivo procesado" in full_l:
                # Parseo
                cargados = 0
                invalidos = 0

                m1 = re.search(r"registros\s+cargados\s*:\s*(\d+)", full_l)
                if m1:
                    cargados = int(m1.group(1))

                m2 = re.search(r"registros\s+inv[aá]lidos\s*:\s*(\d+)", full_l)
                if m2:
                    invalidos = int(m2.group(1))

                # Cerrar popup
                try:
                    btn = win32gui.FindWindowEx(hwnd, 0, "Button", None)
                    if btn:
                        win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)
                    else:
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                except Exception:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

                return {
                    "texto": full_l,
                    "cargados": cargados,
                    "invalidos": invalidos,
                }

        time.sleep(0.2)

    raise TimeoutError(
        "El CSV se cargó, pero no se pudo capturar el MessageBox VB6 "
        "con 'Registros cargados/inválidos'."
    )


# ==========================================================
# GUARDAR MASIVO – FIX CRÍTICO
# ==========================================================

def guardar_masivo(mig_win):
    """
    FUNCIÓN CRÍTICA.
    - NO reintenta
    - NO corrige
    - NO toca otros botones
    """

    btn = _vb6_button(mig_win.handle, "Guardar Masivo")
    if not btn:
        raise RuntimeError(
            "Guardar Masivo NO está disponible. "
            "El CSV no está cargado o la UI está inconsistente."
        )

    _bm_click(btn)

    confirm = _wait_confirmacion()
    _click_si(confirm)

    return _wait_proceso_finalizado(mig_win)


# ==========================================================
# POPUPS
# ==========================================================

def _wait_confirmacion(timeout: int = 20) -> HwndWrapper:
    t0 = time.time()

    while time.time() - t0 < timeout:
        for w in Desktop(backend="win32").windows(visible_only=True):
            try:
                txt = (w.window_text() or "").lower()
                if "confirmación" in txt:
                    logger.info("Popup de confirmación detectado.")
                    return w   # 👈 YA ES UN WRAPPER
            except Exception:
                continue

        time.sleep(0.2)

    raise TimeoutError("No apareció Confirmación")


def _click_si(win):
    win.set_focus()
    for c in win.descendants():
        if (c.window_text() or "").strip().lower() in ("si", "sí"):
            c.click_input()
            return
    send_keys("{ENTER}")


def _wait_proceso_finalizado(mig_win=None, timeout: int = 180) -> dict:
    """
    Captura el MessageBox VB6 final del Guardar Masivo:
    - 'Proceso finalizado con errores'
    - 'Proceso finalizado correctamente'
    """

    t0 = time.time()

    while time.time() - t0 < timeout:
        hwnd = win32gui.FindWindow("#32770", None)
        if hwnd:
            textos = []

            def enum_child(h, _):
                if win32gui.GetClassName(h) == "Static":
                    txt = win32gui.GetWindowText(h).strip()
                    if txt:
                        textos.append(txt)

            win32gui.EnumChildWindows(hwnd, enum_child, None)

            full = "\n".join(textos).lower()

            if "proceso finalizado" in full:
                logger.info("Popup final Guardar Masivo detectado.")

                tiene_errores = "con errores" in full

                # cerrar popup
                try:
                    btn = win32gui.FindWindowEx(hwnd, 0, "Button", None)
                    if btn:
                        win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)
                    else:
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                except Exception:
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

                return {
                    "texto": full,
                    "tiene_errores": tiene_errores
                }

        time.sleep(0.25)

    raise TimeoutError(
        "Guardar Masivo se ejecutó, pero no se pudo capturar "
        "el MessageBox VB6 final."
    )


# ==========================================================
# CSV – UTILIDADES (NO UI)
# ==========================================================

import csv
import random
import re


def crear_csv_muestra_2_sin_error(
    csv_path: str,
    n: int = 2,
) -> str:
    """
    Crea un CSV temporal con n filas cuyo campo de prestación
    NO sea 'ERROR'. Se permite campo vacío.

    Retorna la ruta del CSV generado.
    """

    p = Path(csv_path)
    if not p.exists():
        raise FileNotFoundError(f"No existe el CSV: {csv_path}")

    # Detectar delimitador
    with p.open("r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(2048)
        f.seek(0)
        delim = ";" if sample.count(";") >= sample.count(",") else ","
        reader = csv.DictReader(f, delimiter=delim)
        fieldnames = reader.fieldnames or []
        rows = list(reader)

    if not fieldnames:
        raise RuntimeError("El CSV no tiene encabezados.")

    # Detectar columna de prestación
    col_prest = None
    for h in fieldnames:
        nh = re.sub(r"\s+", " ", h.lower())
        if "prestacion" in nh:
            col_prest = h
            break

    if col_prest is None:
        raise RuntimeError(
            f"No se encontró columna de prestación en headers: {fieldnames}"
        )

    def is_ok(row):
        v = (row.get(col_prest, "") or "").strip().lower()
        return v != "error"

    ok_rows = [r for r in rows if is_ok(r)]

    if len(ok_rows) < n:
        raise RuntimeError(
            f"No hay suficientes filas sin ERROR. "
            f"Disponibles={len(ok_rows)}, requeridas={n}"
        )

    sample_rows = random.sample(ok_rows, n)

    import time

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    out_path = p.with_name(f"{p.stem}_MUESTRA_{n}_{timestamp}{p.suffix}")

    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, delimiter=delim)
        w.writeheader()
        w.writerows(sample_rows)

    logger.info("CSV muestra creado: %s", out_path)
    return str(out_path)



def capturar_errores_desde_datos(csv_path: str) -> dict:
    desk = Desktop(backend="win32")
    datos = desk.window(title_re=r"^Datos$")
    datos.wait("visible", timeout=30)
    datos.set_focus()
    time.sleep(0.4)

    hwnd_form = datos.handle

    # 1) Buscar grid real (recursivo)
    hwnd_grid = _find_grid_hwnd(hwnd_form)

    if not hwnd_grid:
        _dump_descendants(hwnd_form)
        raise RuntimeError("No pude encontrar ningún control tipo grid dentro de Datos.")

    cls = win32gui.GetClassName(hwnd_grid)
    logger.info("GRID detectado en Datos: hwnd=%s class=%s", hwnd_grid, cls)

    # 2) Si es ListView -> leer por mensajes
    rows = []
    if cls == "SysListView32":
        rows = _read_listview(hwnd_grid)

    # 3) Si NO es ListView -> fallback: copiar “como humano” desde el control correcto
    if not rows:
        # click dentro del grid para asegurarnos foco REAL
        try:
            l, t, r, b = win32gui.GetWindowRect(hwnd_grid)
            x = l + 50
            y = t + 60
            win32gui.SetForegroundWindow(hwnd_form)
            time.sleep(0.2)
            user32.SetCursorPos(x, y)
            user32.mouse_event(2, 0, 0, 0, 0)  # down
            user32.mouse_event(4, 0, 0, 0, 0)  # up
            time.sleep(0.2)
        except Exception:
            pass

        # limpiar clipboard y copiar
        try:
            pyperclip.copy("")
        except Exception:
            pass

        # intentos de selección/copia típicos
        send_keys("^a")
        time.sleep(0.15)
        send_keys("^c")
        time.sleep(0.35)

        raw = pyperclip.paste() or ""
        if not raw.strip():
            # segundo intento: a veces CTRL+A no funciona, probamos HOME + SHIFT+END, etc.
            send_keys("{HOME}")
            time.sleep(0.1)
            send_keys("+{END}")
            time.sleep(0.1)
            send_keys("^c")
            time.sleep(0.4)
            raw = pyperclip.paste() or ""

        if not raw.strip():
            _dump_descendants(hwnd_form)
            raise RuntimeError(
                f"No pude copiar contenido desde el grid. Clase detectada: {cls}. "
                "Revisar dump en log para ver controles internos."
            )

        # parse TSV
        filas = [ln.split("\t") for ln in raw.splitlines() if ln.strip()]
        # En tu screenshot: Fila | Identificacion | Nombre | Error
        # si la primera fila son headers, se omite
        if filas and ("error" in (filas[0][-1] or "").lower() or "ident" in (filas[0][1] or "").lower()):
            data = filas[1:]
        else:
            data = filas

        rows = data

    # 4) Extraer errores (col última o 4ta)
    errores = []
    for r in rows:
        if not r:
            continue
        if len(r) >= 4:
            errores.append(r[3].strip())
        else:
            errores.append(r[-1].strip())

    errores = [e for e in errores if e]

    if not errores:
        raise RuntimeError("Leí/copié la tabla pero no encontré textos de error.")

    # 5) Escribir errores en N° Prestacion del CSV
    p = Path(csv_path)
    with p.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        csv_rows = list(reader)
        headers = reader.fieldnames or []

    col_prest = next(h for h in headers if "prestacion" in h.lower())

    for row, err in zip(csv_rows, errores):
        row[col_prest] = err

    out = p.with_name(p.stem + "_ERRORES.csv")
    with out.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, delimiter=";")
        w.writeheader()
        w.writerows(csv_rows)

    logger.info("Errores extraídos y escritos en CSV: %s (n=%s)", out, len(errores))

    # cerrar ventana Datos
    win32gui.PostMessage(hwnd_form, win32con.WM_CLOSE, 0, 0)

    return {"csv_salida": str(out), "errores": len(errores)}



def _click_toolbar_copy(datos_hwnd: int):
    """
    Busca el botón 'Copiar' en la toolbar de la ventana Datos
    y lo presiona (VB6-safe).
    """

    btn_copy = None

    def enum_child(hwnd, _):
        nonlocal btn_copy
        cls = win32gui.GetClassName(hwnd)

        # ToolBarWindow32 es típico en VB6
        if cls == "ToolbarWindow32":
            # botón copiar suele ser el índice 3–6
            btn_copy = hwnd

    win32gui.EnumChildWindows(datos_hwnd, enum_child, None)

    if not btn_copy:
        raise RuntimeError("No se encontró Toolbar de la ventana Datos")

    # Simular click en botón copiar
    win32gui.SendMessage(btn_copy, win32con.WM_COMMAND, 0, 0)
    time.sleep(0.4)

def copiar_desde_menu_datos(datos_win):
    """
    Usa el menú VB6 de la ventana Datos:
    Archivo -> Copiar
    o
    Edición -> Copiar
    """

    datos_win.set_focus()
    time.sleep(0.3)

    try:
        # intenta Edición -> Copiar
        datos_win.menu_select("Edición->Copiar")
        return
    except Exception:
        pass

    try:
        # fallback Archivo -> Copiar
        datos_win.menu_select("Archivo->Copiar")
        return
    except Exception:
        pass

    raise RuntimeError("No se pudo ejecutar Copiar desde el menú de Datos")

def copiar_grid_datos(datos_win):
    """
    Copia el contenido del grid VB6 (MSFlexGrid / VSFlexGrid)
    activando el control correcto.
    """

    hwnd_form = datos_win.handle
    grid_hwnd = None

    def enum_child(hwnd, _):
        nonlocal grid_hwnd
        cls = win32gui.GetClassName(hwnd)

        if cls in (
            "MSFlexGridWndClass",
            "VSFlexGridWndClass",
            "ThunderRT6UserControlDC",
            "ThunderRT6PictureBox",
        ):
            grid_hwnd = hwnd

    win32gui.EnumChildWindows(hwnd_form, enum_child, None)

    if not grid_hwnd:
        raise RuntimeError("No se encontró el grid VB6 en la ventana Datos")

    # 👉 Activar grid
    win32gui.SetForegroundWindow(hwnd_form)
    win32gui.SetFocus(grid_hwnd)
    time.sleep(0.3)

    # 👉 Seleccionar todo + copiar
    send_keys("^a")
    time.sleep(0.1)
    send_keys("^c")
    time.sleep(0.3)

    raw = pyperclip.paste()
    if not raw.strip():
        raise RuntimeError("Clipboard vacío: el grid no copió datos")

    return raw

def cerrar_ventana(win, timeout: float = 8.0):
    """Cierra una ventana por WM_CLOSE y espera que desaparezca."""
    try:
        hwnd = win.handle if hasattr(win, "handle") else int(win)
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    except Exception:
        try:
            # fallback: Alt+F4 si es wrapper
            if hasattr(win, "set_focus"):
                win.set_focus()
                time.sleep(0.2)
            send_keys("%{F4}")
        except Exception:
            pass

    # esperar que cierre (best-effort)
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if not win32gui.IsWindow(hwnd):
                return True
        except Exception:
            return True
        time.sleep(0.2)
    return False

def cerrar_pisco(main_win=None, timeout: float = 10.0) -> bool:
    """
    Cierra PISCO completo.
    - Si main_win existe: intenta cerrarlo (Alt+F4) y aceptar confirmaciones si aparecen.
    - Luego hace taskkill best-effort por el exe configurado (si está en config.ini).
    """
    ok = False

    # Intento suave: cerrar ventana principal
    try:
        if main_win is not None:
            cerrar_ventana(main_win, timeout=timeout)
            ok = True
    except Exception:
        pass

    # Si queda algún popup suelto, intentar cerrarlo
    try:
        _close_any_dialogs(timeout=2.0)
    except Exception:
        pass

    # Intento fuerte: taskkill por exe_path (si se puede leer config)
    try:
        # Reutilizamos config.ini (mismo formato que load_config)
        # IMPORTANTE: ajusta la ruta si tu config.ini está en otro lado.
        cfg = None
        try:
            # si ya existe load_config arriba, úsalo
            # (no importa si falla)
            cfg = load_config(str(Path(__file__).resolve().parent / "config.ini"))
        except Exception:
            cfg = None

        if cfg and getattr(cfg, "exe_path", None):
            _kill_pisco_processes(cfg.exe_path)
            ok = True
    except Exception:
        pass

    return ok
