# -*- coding: utf-8 -*-
"""Ícono de un destino: ícono real del exe, inicial de color, glifo de sistema o '+' (vacío)."""
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import QWidget

from knob.ui.theme import (GLYPH_ADD, GLYPH_FOCUS, GLYPH_MIC, GLYPH_VOLUME, get_theme, icon_font,
                           qcolor, ui_font)

_SYSTEM_GLYPHS = {"master": GLYPH_VOLUME, "mic": GLYPH_MIC, "focus": GLYPH_FOCUS}
_LETTER_COLORS = ("#D6293E", "#5865F2", "#1DB954", "#E5A50A", "#0EA5E9", "#A855F7", "#F97316", "#14B8A6")


def letter_color(name: str) -> str:
    return _LETTER_COLORS[sum(map(ord, name)) % len(_LETTER_COLORS)] if name else _LETTER_COLORS[0]


class IconBadge(QWidget):
    def __init__(self, size: int = 28, parent=None):
        super().__init__(parent)
        self._size = size
        self._name = ""
        self._pixmap = None
        self._kind = "unassigned"
        self._dimmed = False
        self._t = get_theme("Nebula")
        self.setFixedSize(size, size)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def set_content(self, name: str, pixmap, kind) -> None:
        self._name, self._pixmap, self._kind = name or "", pixmap, kind
        self.update()

    def set_dimmed(self, dimmed: bool) -> None:
        if dimmed != self._dimmed:
            self._dimmed = dimmed
            self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        if self._dimmed:
            p.setOpacity(0.4)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        radius = self._size * 0.28
        glyph_px = max(10, int(self._size * 0.52))
        if self._kind == "unassigned":
            p.setPen(QPen(QColor(255, 255, 255, 77), 1, Qt.PenStyle.DashLine))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRoundedRect(r, radius, radius)
            p.setPen(QColor(self._t.text_dim))
            p.setFont(icon_font(glyph_px - 2))
            p.drawText(r, Qt.AlignmentFlag.AlignCenter, GLYPH_ADD)
        elif self._kind in _SYSTEM_GLYPHS:
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(qcolor(self._t.grad_a, 90))
            p.drawRoundedRect(r, radius, radius)
            p.setPen(QColor("#FFFFFF"))
            p.setFont(icon_font(glyph_px))
            p.drawText(r, Qt.AlignmentFlag.AlignCenter, _SYSTEM_GLYPHS[self._kind])
        elif self._pixmap is not None and not self._pixmap.isNull():
            p.drawPixmap(r.toRect(), self._pixmap)
        else:
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(letter_color(self._name)))
            p.drawRoundedRect(r, radius, radius)
            p.setPen(QColor("#FFFFFF"))
            p.setFont(ui_font(int(self._size * 0.5), QFont.Weight.Bold))
            p.drawText(r, Qt.AlignmentFlag.AlignCenter, (self._name[:1] or "?").upper())
