# -*- coding: utf-8 -*-
"""Rutas que funcionan igual desde el código fuente y desde el build de PyInstaller."""
import os
import sys


def is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def project_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def app_dir() -> str:
    """Carpeta del KNOB.exe instalado (o la raíz del proyecto en desarrollo)."""
    return os.path.dirname(sys.executable) if is_frozen() else project_root()


def resource_path(rel: str) -> str:
    """Archivos empaquetados (icon.png, fonts/, version.txt)."""
    base = getattr(sys, "_MEIPASS", None) if is_frozen() else project_root()
    return os.path.join(base or app_dir(), rel)
