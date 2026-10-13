# -*- coding: utf-8 -*-
"""Temas del estilo Glow, colores, tipografías y hoja de estilos global."""
from dataclasses import dataclass

from PySide6.QtGui import QColor, QFont, QFontDatabase

from knob.core.paths import resource_path

TEXT = "#F3EEFF"
GREEN, GREEN_TEXT = "#4ADE80", "#86EFAC"
AMBER, AMBER_TEXT = "#FBBF24", "#FCD34D"
RED, RED_TEXT = "#F87171", "#FCA5A5"
PANEL_BG = QColor(22, 16, 34, 247)

# Glifos de Segoe Fluent Icons (Win11) / Segoe MDL2 Assets (Win10)
GLYPH_VOLUME = "\ue767"
GLYPH_MIC = "\ue720"
GLYPH_FOCUS = "\ue7f4"
GLYPH_ADD = "\ue710"
GLYPH_SETTINGS = "\ue713"
GLYPH_LEFT = "\ue76b"
GLYPH_RIGHT = "\ue76c"
GLYPH_CLOSE = "\ue711"


@dataclass(frozen=True)
class Theme:
    name: str
    grad_a: str      # inicio del degradé (abajo / izquierda)
    grad_b: str      # fin del degradé
    halo: str        # halo radial del fondo
    bg: str          # fondo base y barra de título
    knob_face: str
    knob_ring: str
    pointer: str
    brand_a: str
    brand_b: str
    text_dim: str


THEMES = {
    "Nebula": Theme("Nebula", "#7C3AED", "#EC4899", "#2A1450", "#0C0A14", "#2B2238", "#4B3A6A", "#F0ABFC", "#C084FC", "#F472B6", "#B9AED6"),
    "Cobalt": Theme("Cobalt", "#2563EB", "#22D3EE", "#102A55", "#080B14", "#1D2638", "#33486A", "#7DD3FC", "#60A5FA", "#22D3EE", "#A9BEDB"),
    "Magma":  Theme("Magma",  "#DC2626", "#F97316", "#4A1410", "#120908", "#382222", "#6A3A33", "#FDBA74", "#F87171", "#FB923C", "#D9B3AE"),
    "Toxic":  Theme("Toxic",  "#059669", "#A3E635", "#0F3A26", "#08110C", "#1F3329", "#335A45", "#BEF264", "#34D399", "#A3E635", "#A8CDB8"),
    "Onyx":   Theme("Onyx",   "#737373", "#F5F5F5", "#2A2A2A", "#0C0C0C", "#2A2A2A", "#4A4A4A", "#FFFFFF", "#D4D4D4", "#FFFFFF", "#A3A3A3"),
    "Sunset": Theme("Sunset", "#DB2777", "#FBBF24", "#4A1636", "#120A0E", "#38222F", "#6A3A55", "#FDE68A", "#F472B6", "#FBBF24", "#D9B3C6"),
}
THEME_ORDER = ("Nebula", "Cobalt", "Magma", "Toxic", "Onyx", "Sunset")


def get_theme(name: str) -> Theme:
    return THEMES.get(name, THEMES["Nebula"])


def qcolor(hex_color: str, alpha: int = 255) -> QColor:
    c = QColor(hex_color)
    c.setAlpha(alpha)
    return c


def rgba_css(hex_color: str, alpha: int) -> str:
    c = QColor(hex_color)
    return f"rgba({c.red()},{c.green()},{c.blue()},{alpha})"


def ui_font(px: int, weight=QFont.Weight.Normal) -> QFont:
    f = QFont()
    f.setFamilies(["Segoe UI Variable Text", "Segoe UI"])
    f.setPixelSize(px)
    f.setWeight(weight)
    return f


def icon_font(px: int) -> QFont:
    f = QFont()
    f.setFamilies(["Segoe Fluent Icons", "Segoe MDL2 Assets"])
    f.setPixelSize(px)
    return f


_brand_family = None


def brand_font(px: int) -> QFont:
    """Fuente del logo (Amazing Views); si no se puede cargar, Segoe UI Black."""
    global _brand_family
    if _brand_family is None:
        fid = QFontDatabase.addApplicationFont(resource_path("fonts/AmazingViews.ttf"))
        families = QFontDatabase.applicationFontFamilies(fid) if fid >= 0 else []
        _brand_family = families[0] if families else ""
    f = QFont()
    if _brand_family:
        f.setFamilies([_brand_family])
    else:
        f.setFamilies(["Segoe UI Black", "Segoe UI"])
        f.setWeight(QFont.Weight.Black)
    f.setPixelSize(px)
    return f


def build_qss(t: Theme) -> str:
    """Hoja de estilos de los widgets estándar (los pintados a mano no la usan)."""
    sel = rgba_css(t.grad_a, 115)
    return f"""
    QWidget {{ color: {TEXT}; }}
    QLabel {{ background: transparent; }}
    QLabel#dim {{ color: {t.text_dim}; }}
    QLabel#section {{ color: {t.text_dim}; }}
    QLabel#error {{ color: {RED_TEXT}; }}
    QToolTip {{ background: #161022; color: {TEXT}; border: 1px solid rgba(255,255,255,26); padding: 4px 6px; }}
    QMenu {{ background: #161022; border: 1px solid rgba(255,255,255,26); border-radius: 8px; padding: 4px; }}
    QMenu::item {{ padding: 6px 22px 6px 12px; border-radius: 6px; }}
    QMenu::item:selected {{ background: {sel}; }}
    QMenu::indicator {{ width: 0px; }}
    QComboBox {{ background: rgba(255,255,255,15); border: 1px solid rgba(255,255,255,23); border-radius: 7px; padding: 5px 9px; }}
    QComboBox:hover {{ background: rgba(255,255,255,26); }}
    QComboBox::drop-down {{ border: none; width: 20px; }}
    QComboBox QAbstractItemView {{ background: #161022; border: 1px solid rgba(255,255,255,26); selection-background-color: {sel}; outline: none; padding: 4px; }}
    QLineEdit {{ background: rgba(255,255,255,15); border: 1px solid rgba(255,255,255,31); border-radius: 7px; padding: 7px 9px; selection-background-color: {sel}; }}
    QLineEdit:focus {{ border: 1px solid {t.brand_a}; }}
    QPushButton {{ background: rgba(255,255,255,18); border: 1px solid rgba(255,255,255,26); border-radius: 7px; padding: 6px 12px; }}
    QPushButton:hover {{ background: rgba(255,255,255,31); }}
    QPushButton:disabled {{ color: rgba(255,255,255,77); background: rgba(255,255,255,8); }}
    QPushButton#primary {{ background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {rgba_css(t.grad_a, 200)}, stop:1 {rgba_css(t.grad_b, 150)}); border: none; font-weight: 600; }}
    QPushButton#primary:hover {{ background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 {t.grad_a}, stop:1 {rgba_css(t.grad_b, 200)}); }}
    QPushButton#danger {{ color: {RED_TEXT}; }}
    QPushButton#link {{ background: transparent; border: none; color: {t.text_dim}; padding: 4px 2px; text-align: left; }}
    QPushButton#link:hover {{ color: {TEXT}; }}
    QPushButton#profile {{ padding: 4px 10px; font-size: 12px; }}
    QToolButton {{ background: rgba(255,255,255,15); border: 1px solid rgba(255,255,255,20); border-radius: 7px; }}
    QToolButton:hover {{ background: rgba(255,255,255,31); }}
    QScrollArea {{ background: transparent; border: none; }}
    QScrollArea > QWidget > QWidget {{ background: transparent; }}
    QScrollBar:vertical {{ background: transparent; width: 6px; margin: 0; }}
    QScrollBar::handle:vertical {{ background: rgba(255,255,255,38); border-radius: 3px; min-height: 24px; }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
    """
