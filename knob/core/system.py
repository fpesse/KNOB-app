# -*- coding: utf-8 -*-
"""Integración con Windows: instancia única, inicio automático, teclas, app en foco y barra de título."""
import ctypes
import os
import sys
import threading
import time
import winreg

import psutil

from knob.core.log import log

# --- instancia única ---------------------------------------------------------
_MUTEX_NAME = "Global\\KNOBAppMutex"
_ERROR_ALREADY_EXISTS = 183
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_mutex_handle = None


def acquire_single_instance() -> bool:
    global _mutex_handle
    _mutex_handle = _kernel32.CreateMutexW(None, False, _MUTEX_NAME)
    return ctypes.get_last_error() != _ERROR_ALREADY_EXISTS


def show_already_running() -> None:
    ctypes.windll.user32.MessageBoxW(0, "KNOB ya está en ejecución.", "KNOB", 0x30)


def set_app_user_model_id(app_id: str) -> None:
    """Hace que la barra de tareas muestre el ícono de KNOB y no el de Python."""
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


# --- inicio con Windows ------------------------------------------------------
APP_REG_NAME = "KNOBApp"
_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"


def startup_command() -> str:
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" --minimized'
    return f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}" --minimized'


def is_startup_enabled() -> bool:
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _RUN_KEY, 0, winreg.KEY_READ)
        winreg.QueryValueEx(key, APP_REG_NAME)
        winreg.CloseKey(key)
        return True
    except Exception:
        return False


def set_startup_enabled(enabled: bool) -> None:
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _RUN_KEY, 0, winreg.KEY_SET_VALUE)
        if enabled:
            winreg.SetValueEx(key, APP_REG_NAME, 0, winreg.REG_SZ, startup_command())
        else:
            try:
                winreg.DeleteValue(key, APP_REG_NAME)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except Exception as e:
        log(f"inicio con Windows: {e}")


def refresh_startup_path() -> None:
    """Si el inicio automático está activo, lo apunta al ejecutable actual."""
    if is_startup_enabled():
        set_startup_enabled(True)


# --- teclado -----------------------------------------------------------------
_KEYEVENTF_EXTENDEDKEY = 0x0001
_KEYEVENTF_KEYUP = 0x0002
_MODIFIERS = ("ctrl", "shift", "alt", "win")
VK_MAP = {
    "ctrl": 0x11, "shift": 0x10, "alt": 0x12, "win": 0x5B,
    "space": 0x20, "enter": 0x0D, "return": 0x0D, "esc": 0x1B, "escape": 0x1B,
    "tab": 0x09, "backspace": 0x08, "delete": 0x2E, "insert": 0x2D,
    "home": 0x24, "end": 0x23, "pageup": 0x21, "prior": 0x21,
    "pagedown": 0x22, "next": 0x22,
    "up": 0x26, "down": 0x28, "left": 0x25, "right": 0x27,
    "f1": 0x70, "f2": 0x71, "f3": 0x72, "f4": 0x73, "f5": 0x74, "f6": 0x75,
    "f7": 0x76, "f8": 0x77, "f9": 0x78, "f10": 0x79, "f11": 0x7A, "f12": 0x7B,
}
MEDIA_VK = {"play_pause": 0xB3, "next_track": 0xB0, "prev_track": 0xB1}


def parse_hotkey(combo: str):
    parts = [p.strip().lower() for p in (combo or "").split("+") if p.strip()]
    return [p for p in parts if p in _MODIFIERS], [p for p in parts if p not in _MODIFIERS]


def send_media_key(action: str) -> None:
    vk = MEDIA_VK.get(action)
    if vk is None:
        return
    ctypes.windll.user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY, 0)
    ctypes.windll.user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY | _KEYEVENTF_KEYUP, 0)


def _vk_for(key: str):
    if key in VK_MAP:
        return VK_MAP[key]
    if len(key) == 1:
        return ctypes.windll.user32.VkKeyScanW(ord(key)) & 0xFF
    return None


def send_hotkey(combo: str) -> None:
    modifiers, keys = parse_hotkey(combo)
    if not keys:
        return
    user32 = ctypes.windll.user32
    try:
        for mod in modifiers:
            user32.keybd_event(VK_MAP[mod], 0, 0, 0)
        for k in keys:
            vk = _vk_for(k)
            if vk:
                user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY, 0)
                user32.keybd_event(vk, 0, _KEYEVENTF_EXTENDEDKEY | _KEYEVENTF_KEYUP, 0)
        for mod in reversed(modifiers):
            user32.keybd_event(VK_MAP[mod], 0, _KEYEVENTF_KEYUP, 0)
    except Exception as e:
        log(f"send_hotkey({combo}): {e}")


# --- app en foco -------------------------------------------------------------
def strip_exe(name: str) -> str:
    return name[:-4] if name and name.lower().endswith(".exe") else (name or "")


class FocusTracker:
    """Hilo que cada 150 ms guarda el último exe en primer plano que no sea KNOB."""

    def __init__(self, interval_s: float = 0.15):
        self.current = ""
        self._interval = interval_s

    def start(self) -> None:
        threading.Thread(target=self._run, daemon=True, name="focus-tracker").start()

    def _run(self) -> None:
        own_pid = os.getpid()
        user32 = ctypes.windll.user32
        while True:
            try:
                hwnd = user32.GetForegroundWindow()
                if hwnd:
                    pid = ctypes.c_ulong()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                    if pid.value != own_pid:
                        exe = strip_exe(psutil.Process(pid.value).name()).capitalize()
                        if exe:
                            self.current = exe
            except Exception:
                pass
            time.sleep(self._interval)


# --- barra de título ---------------------------------------------------------
_DWMWA_USE_IMMERSIVE_DARK_MODE = 20
_DWMWA_BORDER_COLOR = 34
_DWMWA_CAPTION_COLOR = 35


def hex_to_colorref(hex_color: str) -> int:
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (b << 16) | (g << 8) | r


def style_title_bar(hwnd: int, caption_hex: str) -> None:
    """Barra de título oscura y del color del tema (el color solo aplica en Windows 11)."""
    try:
        dwm = ctypes.windll.dwmapi
        dark = ctypes.c_int(1)
        dwm.DwmSetWindowAttribute(hwnd, _DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(dark), ctypes.sizeof(dark))
        color = ctypes.c_uint(hex_to_colorref(caption_hex))
        dwm.DwmSetWindowAttribute(hwnd, _DWMWA_CAPTION_COLOR, ctypes.byref(color), ctypes.sizeof(color))
        dwm.DwmSetWindowAttribute(hwnd, _DWMWA_BORDER_COLOR, ctypes.byref(color), ctypes.sizeof(color))
    except Exception as e:
        log(f"style_title_bar: {e}")
