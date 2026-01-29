# robot/PISCO_CapturarServicios.py
# ==========================================
# PISCO_CapturarServicios.py – VB6/Win32 robusto
#
# Flujo:
#   1) Archivo -> Capturar Servicios
#   2) Popup "Mes a Visualizar Servicios" -> Aceptar
#   3) Esperar pantalla "Ordenes de Servicio" (combo + edit + lupa)
#   4) Seleccionar criterio en Combo DropDownList NO editable:
#         "Por Cedula del Fallecido"
#      (robusto: abrir dropdown y leer ComboLBox)
#   5) Escribir cédula en Edit
#   6) Click lupa
#   7) Si sale Error 13 / No coinciden los tipos -> cerrar y fallar
# ==========================================

from __future__ import annotations

import time
import logging
import ctypes


import unicodedata

import win32gui
import win32con
import win32api

from pywinauto import Desktop
from pywinauto.timings import TimeoutError
from pywinauto.keyboard import send_keys

logger = logging.getLogger("Robot62.PISCO.CapturarServicios")

user32 = ctypes.windll.user32

# -------------------------
# ComboBox messages
# -------------------------
CB_GETCOUNT = 0x0146
CB_SETCURSEL = 0x014E
CB_SHOWDROPDOWN = 0x014F

# -------------------------
# ListBox messages (ComboLBox)
# -------------------------
LB_GETCOUNT = 0x018B
LB_GETTEXTLEN = 0x018A
LB_GETTEXT = 0x0189
LB_SETCURSEL = 0x0186

# -------------------------
# Notificaciones WM_COMMAND
# -------------------------
CBN_SELCHANGE = 1
EN_CHANGE = 0x0300


# ----------------------------------------------------------
# Helpers Win32 base
# ----------------------------------------------------------
WM_SETTEXT = 0x000C
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E

import re

def _find_top_window_title_contains(substr: str, timeout: float = 10.0, poll: float = 0.2) -> int | None:
    """Busca una ventana top-level visible cuyo título contenga substr (case-insensitive)."""
    target = (substr or "").strip().lower()
    if not target:
        return None

    t0 = time.time()
    while time.time() - t0 < timeout:
        found = None

        def enum_windows(hwnd, _):
            nonlocal found
            if found:
                return
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                title = (win32gui.GetWindowText(hwnd) or "").strip().lower()
                if target in title:
                    found = hwnd
            except Exception:
                return

        win32gui.EnumWindows(enum_windows, None)
        if found:
            return found

        time.sleep(poll)

    return None


def _extract_contrato_nro_from_control_llamadas(timeout: float = 10.0) -> str | None:
    """
    Abre/espera la ventana 'Control de Llamadas / Novedades' y extrae el texto del campo 'Contrato Nro:'.
    Estrategia:
      A) Buscar el Static 'Contrato Nro:' y tomar el Edit más cercano a su derecha.
      B) Fallback: escoger un Edit que parezca contrato con patrón tipo '05-0791-26'.
    """
    hwnd = _find_top_window_title_contains("Control de Llamadas / Novedades", timeout=timeout, poll=0.2)
    if not hwnd:
        return None

    # 1) Enumerar hijos
    children = _all_descendants(hwnd)

    statics = []
    edits = []

    for h in children:
        try:
            if not win32gui.IsWindowVisible(h):
                continue
            cls = win32gui.GetClassName(h)
            txt = (win32gui.GetWindowText(h) or "").strip()

            l, t, r, b = win32gui.GetWindowRect(h)
            w, hgt = (r - l), (b - t)

            if cls == "Static" and txt:
                statics.append((h, txt, l, t, r, b))
            elif cls in ("Edit", "ThunderRT6TextBox", "ThunderRT6MaskedEdit", "ThunderRT6TextBox2"):
                # filtrar edits muy grandes tipo textarea
                if w <= 260 and hgt <= 40:
                    edits.append((h, l, t, r, b))
        except Exception:
            continue

    # 2) Buscar label "Contrato Nro"
    label = None
    for h, txt, l, t, r, b in statics:
        if "contrato" in txt.lower() and "nro" in txt.lower():
            label = (h, txt, l, t, r, b)
            break

    # 3) Si hallamos label, elegir el edit a la derecha, alineado en Y
    if label and edits:
        _, _, ll, lt, lr, lb = label

        best = None
        best_score = 10**9
        for eh, el, et, er, eb in edits:
            # debe estar a la derecha del label
            if el < lr - 5:
                continue
            y = abs(et - lt)
            xgap = abs(el - lr)
            score = y * 4 + xgap
            if score < best_score:
                best_score = score
                best = eh

        if best:
            val = _get_text(best)
            val = (val or "").strip()
            if val:
                return val

    # 4) Fallback: buscar un Edit con pinta de contrato (05-0791-26)
    # patrón flexible: 2 dígitos - 3 a 5 dígitos - 2 dígitos
    pat = re.compile(r"^\d{2}-\d{3,5}-\d{2}$")
    candidates = []
    for eh, el, et, er, eb in edits:
        v = (_get_text(eh) or "").strip()
        if not v:
            continue
        if pat.match(v):
            return v
        # guardar por si no matchea exacto pero tiene guiones
        if "-" in v and len(v) <= 12:
            candidates.append(v)

    if candidates:
        # devolver el “más parecido”
        return sorted(candidates, key=len)[0]

    return None


def _get_text(hwnd: int) -> str:
    """Lee texto real del control (Edit VB6) usando WM_GETTEXT."""
    try:
        ln = win32gui.SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
        if not ln or int(ln) <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(int(ln) + 1)
        win32gui.SendMessage(hwnd, WM_GETTEXT, int(ln) + 1, buf)
        return (buf.value or "").strip()
    except Exception:
        # fallback muy suave
        try:
            return (win32gui.GetWindowText(hwnd) or "").strip()
        except Exception:
            return ""

def _set_text_wm(hwnd_edit: int, value: str) -> bool:
    """Setea texto al Edit sin depender del teclado (VB6 friendly)."""
    try:
        win32gui.SendMessage(hwnd_edit, WM_SETTEXT, 0, str(value))
        time.sleep(0.05)
        _notify_parent_command_smart(hwnd_edit, EN_CHANGE)
        time.sleep(0.05)
        return True
    except Exception as e:
        logger.warning("WM_SETTEXT falló en edit=%s: %s", hwnd_edit, e)
        return False

def _type_cedula_robusto(hwnd_edit: int, cedula: str, retries: int = 3) -> None:
    """
    1) Click foco + limpiar + WM_SETTEXT
    2) Verifica leyendo texto; si no, fallback send_keys
    """
    for k in range(retries):
        try:
            # foco + click
            parent = win32gui.GetParent(hwnd_edit)
            if parent:
                try: win32gui.SetForegroundWindow(parent)
                except Exception: pass
            try: win32gui.SetFocus(hwnd_edit)
            except Exception: pass

            l, t, r, b = _rect(hwnd_edit)
            _click_at_screen(l + 10, (t + b) // 2)
            time.sleep(0.05)

            # limpiar
            send_keys("^a{BACKSPACE}")
            time.sleep(0.05)

            # set por WM
            _set_text_wm(hwnd_edit, cedula)

            # verificar
            txt = _get_text(hwnd_edit)
            if txt == str(cedula):
                logger.info("Cédula escrita OK en edit=%s", hwnd_edit)
                return

            # fallback: teclear directo
            send_keys(str(cedula), with_spaces=True)
            time.sleep(0.05)
            _notify_parent_command_smart(hwnd_edit, EN_CHANGE)

            txt2 = _get_text(hwnd_edit)
            if txt2 == str(cedula):
                logger.info("Cédula escrita OK (fallback teclado) edit=%s", hwnd_edit)
                return

            logger.warning("Intento %s: no quedó la cédula. edit_text='%s'", k + 1, txt2)
        except Exception as e:
            logger.warning("Intento %s: error escribiendo cédula: %s", k + 1, e)

        time.sleep(0.12)

    raise RuntimeError("No pude escribir la cédula en el campo correcto (Edit).")


def _dismiss_unexpected_mes_dialogs(timeout: float = 0.2) -> None:
    """
    Cierra cualquier diálogo modal 'Mes a Visualizar ...' que aparezca,
    excepto el de Servicios (porque ese ya lo manejas en capturar_servicios_desde_menu).
    """
    t0 = time.time()
    while time.time() - t0 < timeout:
        hwnd = _find_dialog_by_static_contains("Mes a Visualizar", timeout=0.05, poll=0.01)
        if not hwnd:
            return

        # Leer el texto completo del Static para saber cuál es
        texts = []
        def enum_child(ch, __):
            if win32gui.GetClassName(ch) == "Static":
                txt = (win32gui.GetWindowText(ch) or "").strip()
                if txt:
                    texts.append(txt)

        try:
            win32gui.EnumChildWindows(hwnd, enum_child, None)
        except Exception:
            return

        full = " ".join(texts).lower()

        # Si NO es el de servicios, cerrarlo (preferible Cancelar)
        if "servicios" not in full:
            if not _click_button_in_dialog(hwnd, "Cancelar"):
                _close_dialog_ok(hwnd)  # fallback
            time.sleep(0.05)
            continue

        # Si es el de servicios, no lo tocamos aquí
        return


def _enum_children(hwnd_parent: int) -> list[int]:
    out: list[int] = []

    def cb(h, _):
        out.append(h)

    try:
        win32gui.EnumChildWindows(hwnd_parent, cb, None)
    except Exception:
        pass
    return out


def _rect(hwnd: int):
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    return l, t, r, b


def _area(hwnd: int) -> int:
    l, t, r, b = _rect(hwnd)
    return max(0, r - l) * max(0, b - t)


def _get_ctrl_id(hwnd: int) -> int:
    try:
        return win32gui.GetDlgCtrlID(hwnd)
    except Exception:
        return 0


def _get_ancestor(hwnd: int, max_hops: int = 12) -> int:
    cur = hwnd
    for _ in range(max_hops):
        p = win32gui.GetParent(cur)
        if not p or p == cur:
            break
        cur = p
    return cur


def _notify_parent_command_smart(hwnd_ctrl: int, notify_code: int) -> None:
    parent = win32gui.GetParent(hwnd_ctrl)
    if not parent:
        return

    ctrl_id = _get_ctrl_id(hwnd_ctrl)
    wparam = (notify_code << 16) | (ctrl_id & 0xFFFF)

    try:
        win32gui.PostMessage(parent, win32con.WM_COMMAND, wparam, hwnd_ctrl)
    except Exception:
        pass

    anc = _get_ancestor(hwnd_ctrl)
    if anc and anc != parent:
        try:
            win32gui.PostMessage(anc, win32con.WM_COMMAND, wparam, hwnd_ctrl)
        except Exception:
            pass


def _click(hwnd_btn: int) -> None:
    try:
        win32gui.PostMessage(hwnd_btn, win32con.BM_CLICK, 0, 0)
    except Exception:
        pass


def _click_at_screen(x: int, y: int) -> None:
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


# ----------------------------------------------------------
# Popups (#32770) – Mes a Visualizar / Error 13
# ----------------------------------------------------------
def _find_dialog_by_static_contains(substr: str, timeout: float = 10.0, poll: float = 0.2) -> int | None:
    s = (substr or "").strip().lower()
    if not s:
        return None

    t0 = time.time()
    while time.time() - t0 < timeout:
        found = None

        def enum_windows(hwnd, _):
            nonlocal found
            if found:
                return
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                if win32gui.GetClassName(hwnd) != "#32770":
                    return

                texts = []

                def enum_child(ch, __):
                    if win32gui.GetClassName(ch) == "Static":
                        txt = (win32gui.GetWindowText(ch) or "").strip()
                        if txt:
                            texts.append(txt)

                win32gui.EnumChildWindows(hwnd, enum_child, None)
                full = " ".join(texts).lower()
                if s in full:
                    found = hwnd
            except Exception:
                return

        win32gui.EnumWindows(enum_windows, None)
        if found:
            return found

        time.sleep(poll)

    return None


def _click_button_in_dialog(hwnd_dlg: int, label: str) -> bool:
    target = (label or "").strip().lower()
    if not target:
        return False

    btn_hwnd = None

    def enum_child(hwnd, _):
        nonlocal btn_hwnd
        if btn_hwnd:
            return
        if win32gui.GetClassName(hwnd) != "Button":
            return
        txt = (win32gui.GetWindowText(hwnd) or "").strip().lower()
        if txt == target:
            btn_hwnd = hwnd

    try:
        win32gui.EnumChildWindows(hwnd_dlg, enum_child, None)
    except Exception:
        return False

    if not btn_hwnd:
        return False

    try:
        win32gui.PostMessage(btn_hwnd, win32con.BM_CLICK, 0, 0)
        return True
    except Exception:
        return False


def _close_dialog_ok(hwnd_dlg: int) -> None:
    if _click_button_in_dialog(hwnd_dlg, "Aceptar"):
        return
    try:
        win32gui.PostMessage(hwnd_dlg, win32con.WM_CLOSE, 0, 0)
    except Exception:
        pass


def _find_mes_servicios_dialog(timeout: float = 20.0) -> int | None:
    return _find_dialog_by_static_contains("Mes a Visualizar Servicios", timeout=timeout, poll=0.2)


def _find_error13_dialog(timeout: float = 2.5) -> int | None:
    hwnd = _find_dialog_by_static_contains("No coinciden los tipos", timeout=timeout, poll=0.1)
    if hwnd:
        return hwnd
    return _find_dialog_by_static_contains("Error '13'", timeout=timeout, poll=0.1)


# ----------------------------------------------------------
# Combo DropDownList NO editable – método robusto por ComboLBox
# ----------------------------------------------------------
def _wait_combolbox(timeout: float = 2.0, poll: float = 0.05) -> int | None:
    """
    Cuando un ComboBox despliega, Windows crea un ListBox top-level clase 'ComboLBox'.
    Esta función lo busca visible.
    """
    t0 = time.time()
    while time.time() - t0 < timeout:
        found = None

        def enum_windows(hwnd, _):
            nonlocal found
            if found:
                return
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                cls = win32gui.GetClassName(hwnd)
                if cls == "ComboLBox":
                    found = hwnd
            except Exception:
                return

        win32gui.EnumWindows(enum_windows, None)
        if found:
            return found

        time.sleep(poll)

    return None




def _listbox_items(hwnd_lb: int) -> list[str]:
    n = win32gui.SendMessage(hwnd_lb, LB_GETCOUNT, 0, 0)
    if not n or int(n) <= 0:
        return []

    items: list[str] = []
    for i in range(int(n)):
        ln = win32gui.SendMessage(hwnd_lb, LB_GETTEXTLEN, i, 0)
        if ln is None or int(ln) < 0:
            items.append("")
            continue

        buf = ctypes.create_unicode_buffer(int(ln) + 1)
        win32gui.SendMessage(hwnd_lb, LB_GETTEXT, i, buf)
        items.append((buf.value or "").strip())

    return items


def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    # quitar tildes/acentos: "cédula" -> "cedula"
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def _combo_select_second_option(hwnd_cb: int, timeout: float = 3.0) -> bool:
    """
    Regla de negocio: siempre elegir la 2da opción (index 1).
    Método robusto: abrir dropdown -> ubicar ComboLBox -> click coordenada del item #2.
    """
    # foco real al combo
    try:
        parent = win32gui.GetParent(hwnd_cb) or hwnd_cb
        win32gui.SetForegroundWindow(parent)
    except Exception:
        pass
    try:
        win32gui.SetFocus(hwnd_cb)
    except Exception:
        pass

    time.sleep(0.08)

    # abrir dropdown SOLO Win32
    try:
        win32gui.SendMessage(hwnd_cb, CB_SHOWDROPDOWN, 1, 0)
    except Exception:
        pass
    time.sleep(0.10)

    hwnd_lb = _wait_combolbox(timeout=timeout)
    if not hwnd_lb:
        # fallback click en flecha del combo
        try:
            l, t, r, b = _rect(hwnd_cb)
            _click_at_screen(r - 8, (t + b) // 2)
            time.sleep(0.10)
        except Exception:
            pass
        hwnd_lb = _wait_combolbox(timeout=timeout)

    if not hwnd_lb:
        logger.error("No apareció ComboLBox al desplegar el combo.")
        return False

    # asegurar que hay al menos 2 items
    n = win32gui.SendMessage(hwnd_lb, LB_GETCOUNT, 0, 0)
    if not n or int(n) < 2:
        logger.error("ComboLBox sin suficientes items. count=%s", n)
        send_keys("{ESC}")
        return False

    # click en el centro del 2do item (index=1)
    try:
        l, t, r, b = _rect(hwnd_lb)
        h = max(1, b - t)

        # estimación altura por item (si hay scroll, igual funciona bastante bien)
        item_h = max(14, int(h / max(2, int(n))))
        y = t + int(item_h * 1 + item_h / 2)   # item #2
        x = l + int((r - l) * 0.50)

        _click_at_screen(x, y)
        time.sleep(0.10)

    except Exception as e:
        logger.error("Fallo click en item #2 del ComboLBox: %s", e)
        send_keys("{ESC}")
        return False

    # cerrar dropdown sin disparar acciones (ENTER a veces activa "Buscar")
    try:
        win32gui.SendMessage(hwnd_cb, CB_SHOWDROPDOWN, 0, 0)
    except Exception:
        pass
    send_keys("{ESC}")  # respaldo
    time.sleep(0.10)

    # notificar VB6 (por si amarra lógica a CBN_SELCHANGE)
    _notify_parent_command_smart(hwnd_cb, CBN_SELCHANGE)
    time.sleep(0.12)

    return True


# ----------------------------------------------------------
# Escritura VB6-safe en Edit
# ----------------------------------------------------------
def _type_in_edit_like_human(hwnd_edit: int, value: str) -> None:
    parent = win32gui.GetParent(hwnd_edit)

    try:
        if parent:
            win32gui.SetForegroundWindow(parent)
    except Exception:
        pass

    try:
        win32gui.SetFocus(hwnd_edit)
    except Exception:
        pass

    # click dentro del edit
    try:
        l, t, r, b = _rect(hwnd_edit)
        _click_at_screen(l + 12, t + 10)
    except Exception:
        pass

    time.sleep(0.10)
    send_keys("^a{BACKSPACE}")
    time.sleep(0.05)
    send_keys(str(value), with_spaces=True)
    time.sleep(0.10)

    _notify_parent_command_smart(hwnd_edit, EN_CHANGE)
    time.sleep(0.12)


# ----------------------------------------------------------
# Detección de pantalla de búsqueda (combo + edit + lupa)
# ----------------------------------------------------------
def _find_controls_for_busqueda(hwnd_container: int):
    children = _enum_children(hwnd_container)

    combos = []
    edits = []
    buttons = []
    clickables = []  # cosas que podrían ser lupa aunque no sean Button

    for h in children:
        try:
            if not win32gui.IsWindowVisible(h):
                continue

            cls = win32gui.GetClassName(h)
            txt = (win32gui.GetWindowText(h) or "").strip()
            l, t, r, b = _rect(h)
            w, hgt = (r - l), (b - t)

            # Combo (tu dropdown)
            if cls in ("ComboBox", "ComboBoxEx32", "ThunderRT6ComboBox", "ThunderComboBox"):
                combos.append((h, w, hgt, l, t, r, b, txt, cls))

            # Edit (campo de texto / número)
            elif cls in ("Edit", "ThunderRT6TextBox"):
                edits.append((h, w, hgt, l, t, r, b, txt, cls))

            # Lupa a veces es Button normal
            elif cls in ("Button", "ThunderRT6CommandButton"):
                buttons.append((h, w, hgt, l, t, r, b, txt, cls))

            # Lupa a veces es imagen/toolbar/static
            elif cls in ("Static", "ToolbarWindow32", "ThunderRT6PictureBox"):
                clickables.append((h, w, hgt, l, t, r, b, txt, cls))

        except Exception:
            continue

    # ⚠️ Cambio clave: ya NO exigimos buttons
    if not combos or not edits:
        raise RuntimeError("No encontré combo y edit suficientes en el contenedor.")

    # Combo: el más ancho
    combo = max(combos, key=lambda x: x[1])[0]
    cl, ct, cr, cb = _rect(combo)

    # Edit: el más alineado con el combo en Y
    def ydist(e):
        _, _, _, _, et, _, _, _, _, _ = e
        return abs(et - ct)

    edit = min(edits, key=ydist)[0]

    # Intentar detectar lupa como botón primero
    lupa = None
    if buttons:
        def score_btn(btn):
            hwnd, w, hgt, l, t, r, b, txt, cls = btn
            dx = abs(l - cr)
            dy = abs(t - ct)
            squareness = abs(w - hgt)
            has_text_penalty = 200 if (txt.strip() != "") else 0
            return dx + dy + squareness + has_text_penalty

        lupa = min(buttons, key=score_btn)[0]

    # Si no hay Button, intentar con clickables (Static/Toolbar/etc.)
    if lupa is None and clickables:
        def score_clickable(it):
            hwnd, w, hgt, l, t, r, b, txt, cls = it
            dx = abs(l - cr)
            dy = abs(t - ct)
            squareness = abs(w - hgt)
            # la lupa suele ser “cuadrito” pequeño
            small_penalty = 0 if (w <= 40 and hgt <= 40) else 150
            return dx + dy + squareness + small_penalty

        lupa = min(clickables, key=score_clickable)[0]

    return combo, edit, lupa

def _is_in_top_bar(hwnd: int, main_hwnd: int, margin_top: int = 180) -> bool:
    """Filtra controles que están en la franja superior (donde está edit+combo+lupa)."""
    try:
        ml, mt, mr, mb = win32gui.GetWindowRect(main_hwnd)
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        return (t >= mt) and (t <= mt + margin_top)
    except Exception:
        return False


def _all_descendants(hwnd_parent: int) -> list[int]:
    out: list[int] = []
    def cb(h, _):
        out.append(h)
    try:
        win32gui.EnumChildWindows(hwnd_parent, cb, None)  # descendientes (no solo hijos directos)
    except Exception:
        pass
    return out


def _wait_busqueda_controls(main_win, timeout: float = 45.0, poll: float = 0.25):
    """
    Nuevo enfoque: buscar GLOBALMENTE dentro del main_win
    y armar el trío (edit + combo + lupa) por geometría.
    """
    t0 = time.time()
    last_err = None

    MAIN = main_win.handle

    while time.time() - t0 < timeout:
        try:
            hwnds = _all_descendants(MAIN)

            combos = []
            edits = []
            clickables = []

            for h in hwnds:
                try:
                    if not win32gui.IsWindowVisible(h):
                        continue
                    if not _is_in_top_bar(h, MAIN, margin_top=220):
                        continue

                    cls = win32gui.GetClassName(h)
                    txt = (win32gui.GetWindowText(h) or "").strip()
                    l, t, r, b = win32gui.GetWindowRect(h)
                    w, hgt = (r - l), (b - t)

                    # Combo (dropdown)
                    if cls in ("ComboBox", "ComboBoxEx32", "ThunderRT6ComboBox", "ThunderComboBox"):
                        # combos muy pequeños suelen ser basura, filtramos
                        if w >= 120 and hgt >= 18:
                            combos.append((h, l, t, r, b, w, hgt, txt, cls))

                    # Edit (campo texto)
                    elif cls in ("Edit", "ThunderRT6TextBox"):
                        # el edit de la barra suele ser ancho medio
                        if w >= 80 and hgt >= 18:
                            edits.append((h, l, t, r, b, w, hgt, txt, cls))

                    # Posible lupa (a veces no es Button)
                    elif cls in ("Button", "Static", "ToolbarWindow32", "ThunderRT6PictureBox", "ThunderRT6CommandButton"):
                        # la lupa es un cuadrito pequeño
                        if w <= 60 and hgt <= 60:
                            clickables.append((h, l, t, r, b, w, hgt, txt, cls))

                except Exception:
                    continue

            if not combos:
                last_err = RuntimeError("No encontré ningún Combo en la barra superior.")
                time.sleep(poll)
                continue

            # Elegir el combo "de criterio" = normalmente el más ancho de la barra
            combo = max(combos, key=lambda x: x[5])[0]
            cl, ct, cr, cb = win32gui.GetWindowRect(combo)

            # Elegir edit más cercano a la izquierda del combo y alineado en Y
            best_edit = None
            best_score = 10**9
            for e in edits:
                eh, el, et, er, eb, ew, ehgt, etxt, ecls = e
                # Queremos edit a la izquierda del combo y en la misma “fila”
                if er > cl:  # si el edit está pasando el combo, no es el que queremos
                    continue
                y = abs(et - ct)
                xgap = abs(cl - er)
                score = y * 3 + xgap
                if score < best_score:
                    best_score = score
                    best_edit = eh

            if not best_edit:
                last_err = RuntimeError("Encontré Combo, pero no hallé Edit candidato alineado a su izquierda.")
                time.sleep(poll)
                continue

            edit = best_edit

            # Lupa: lo más cercano al borde derecho del combo (si existe como control)
            lupa = None
            best_lupa_score = 10**9
            for it in clickables:
                hh, l, t, r, b, w, hgt, txt, cls = it
                dx = abs(l - cr)
                dy = abs(t - ct)
                score = dx + dy
                if score < best_lupa_score:
                    best_lupa_score = score
                    lupa = hh

            # ¡Listo!
            hwnd_container = MAIN  # ya no dependemos de contenedor especial
            return hwnd_container, combo, edit, lupa

        except Exception as e:
            last_err = e

        time.sleep(poll)

    raise TimeoutError(
        f"No apareció la pantalla de búsqueda (combo+edit+lupa) dentro del main_win. "
        f"Último error: {last_err}"
    )


# ----------------------------------------------------------
# API pública
# ----------------------------------------------------------
def abrir_capturar_servicios(main_win) -> None:
    main_win.set_focus()
    time.sleep(0.2)
    try:
        main_win.menu_select("Archivo->Capturar Servicios")
        return
    except Exception as e:
        logger.warning("menu_select falló en main_win: %s", e)

    desk = Desktop(backend="win32")
    w = desk.window(handle=main_win.handle)
    w.wait("visible", timeout=10)
    w.menu_select("Archivo->Capturar Servicios")


def aceptar_popup_mes_servicios(timeout: int = 20) -> dict:
    hwnd = _find_mes_servicios_dialog(timeout=timeout)
    if not hwnd:
        raise TimeoutError("No apareció el popup de 'Mes a Visualizar Servicios'.")

    if _click_button_in_dialog(hwnd, "Aceptar"):
        return {"ok": True}

    # fallback ENTER
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
    return {"ok": True, "fallback": "ENTER"}


def capturar_servicios_desde_menu(main_win, timeout_popup: int = 20) -> dict:
    abrir_capturar_servicios(main_win)

    # a veces demora en cargar el form y/o el popup
    time.sleep(0.8)

    res = aceptar_popup_mes_servicios(timeout=timeout_popup)
    logger.info("Popup mes servicios aceptado: %s", res)
    return res

def _find_busqueda_no_encontro(timeout: float = 0.8) -> int | None:
    # cubre: "No se encontro registro alguno bajo este criterio"
    hwnd = _find_dialog_by_static_contains("no se encontro registro", timeout=timeout, poll=0.05)
    if hwnd:
        return hwnd
    return _find_dialog_by_static_contains("no se encontr", timeout=timeout, poll=0.05)


def _close_busqueda_no_encontro():
    hwnd = _find_busqueda_no_encontro(timeout=0.8)
    if hwnd:
        _close_dialog_ok(hwnd)
        return True
    return False

def _click_lupa_relativo_al_combo(hwnd_combo: int) -> None:
    l, t, r, b = _rect(hwnd_combo)
    # en tus screenshots es el cuadrito inmediatamente a la derecha del combo
    x = r + 14
    y = (t + b) // 2
    _click_at_screen(x, y)
    time.sleep(0.08)


import re

# ----------------------------------------------------------
# NUEVO: cerrar "Control de Llamadas / Novedades"
# ----------------------------------------------------------
def _close_control_llamadas(timeout: float = 8.0) -> bool:
    hwnd = _find_top_window_title_contains("Control de Llamadas / Novedades", timeout=timeout, poll=0.2)
    if not hwnd:
        return False
    try:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    except Exception:
        pass

    # esperar a que cierre
    t0 = time.time()
    while time.time() - t0 < 3.0:
        if not win32gui.IsWindow(hwnd):
            return True
        time.sleep(0.1)

    return not win32gui.IsWindow(hwnd)


# ----------------------------------------------------------
# NUEVO: capturar "No Orden Servicio" del MAIN
# (busca label Static y toma el Edit más cercano a la derecha)
# ----------------------------------------------------------
def _extract_no_orden_servicio_from_main(main_win, timeout: float = 8.0) -> str | None:
    pat = re.compile(r"^\d{2}-\d{3,5}-\d{2}$")  # ej: 05-0791-26

    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            MAIN = main_win.handle
            kids = _all_descendants(MAIN)

            statics = []
            edits = []

            for h in kids:
                try:
                    if not win32gui.IsWindowVisible(h):
                        continue
                    cls = win32gui.GetClassName(h)
                    txt = (win32gui.GetWindowText(h) or "").strip()
                    l, t, r, b = win32gui.GetWindowRect(h)
                    w, hgt = (r - l), (b - t)

                    if cls == "Static" and txt:
                        statics.append((h, txt, l, t, r, b))
                    elif cls in ("Edit", "ThunderRT6TextBox", "ThunderRT6MaskedEdit"):
                        # el de "No Orden Servicio" es pequeño/mediano, no textarea gigante
                        if w <= 260 and hgt <= 40:
                            edits.append((h, l, t, r, b))
                except Exception:
                    continue

            # 1) Label exacto por texto
            label = None
            for _, txt, ll, lt, lr, lb in statics:
                low = txt.lower()
                if "orden servicio" in low:  # cubre "* No Orden Servicio:"
                    label = (ll, lt, lr, lb)
                    break

            # 2) Si hay label, escoger edit a la derecha
            if label and edits:
                ll, lt, lr, lb = label
                best = None
                best_score = 10**9

                for eh, el, et, er, eb in edits:
                    if el < lr - 5:
                        continue
                    y = abs(et - lt)
                    xgap = abs(el - lr)
                    score = y * 4 + xgap
                    if score < best_score:
                        best_score = score
                        best = eh

                if best:
                    v = (_get_text(best) or "").strip()
                    if v:
                        return v

            # 3) Fallback: algún edit con patrón "05-0791-26"
            for eh, *_ in edits:
                v = (_get_text(eh) or "").strip()
                if pat.match(v):
                    return v

        except Exception:
            pass

        time.sleep(0.2)

    return None


# ----------------------------------------------------------
# AJUSTE: flujo principal (sin el warning de cédula borrada)
# ----------------------------------------------------------
def buscar_por_cedula_fallecido(
    main_win,
    criterio_text: str = "Por Cedula del Fallecido",
    cedula: str = "8349505",
    timeout_form: int = 45,
) -> dict:
    main_win.set_focus()
    time.sleep(10)

    _dismiss_unexpected_mes_dialogs(timeout=0.8)

    # A) Ubicar barra (edit + combo + lupa)
    hwnd_container, combo, edit, _lupa = _wait_busqueda_controls(main_win, timeout=timeout_form)
    logger.info("BUSQUEDA A combo=%s edit=%s", combo, edit)

    # 1) escribir cédula
    _type_cedula_robusto(edit, cedula, retries=4)
    time.sleep(0.20)

    # 2) seleccionar 2da opción (Por Cédula...)
    ok = False
    for _ in range(3):
        ok = _combo_select_second_option(combo, timeout=6.0)
        if ok:
            break
        time.sleep(0.25)
    if not ok:
        raise RuntimeError("No pude seleccionar la 2da opción del combo (Por Cédula del Fallecido).")

    # 3) re-escribir cédula (VB6 a veces recalcula)
    time.sleep(0.25)
    _, combo2, edit2, _ = _wait_busqueda_controls(main_win, timeout=8.0)
    _type_cedula_robusto(edit2, cedula, retries=3)

    # 4) click lupa
    time.sleep(0.10)
    _click_lupa_relativo_al_combo(combo2)

    # ✅ PRIMERO: si aparece "No se encontró registro..." -> cerrar y salir
    time.sleep(0.35)
    if _close_busqueda_no_encontro():
        logger.warning("Búsqueda: NO encontró registro para cédula=%s", cedula)
        return {"ok": False, "motivo": "NO_ENCONTRADO", "cedula": cedula}

    # Error 13
    hwnd_err = _find_error13_dialog(timeout=1.2)
    if hwnd_err:
        _close_dialog_ok(hwnd_err)
        raise RuntimeError("PISCO Error 13: No coinciden los tipos (al buscar por cédula).")

    # D) esperar "Control de Llamadas / Novedades" (si aparece) y cerrarla
    _ = _find_top_window_title_contains("Control de Llamadas / Novedades", timeout=12.0, poll=0.2)
    _close_control_llamadas(timeout=2.0)

    # E) capturar "No Orden Servicio"
    no_orden = _extract_no_orden_servicio_from_main(main_win, timeout=8.0)
    if no_orden:
        logger.info("✅ No Orden Servicio capturado: %s", no_orden)
    else:
        logger.warning("No pude capturar 'No Orden Servicio' del formulario principal.")

    return {"ok": True, "criterio": "Por Cédula del Fallecido", "cedula": cedula, "no_orden_servicio": no_orden}

