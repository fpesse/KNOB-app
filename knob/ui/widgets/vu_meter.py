# -*- coding: utf-8 -*-
"""Medidor de nivel fino con retención de pico (verde / ámbar / rojo como v3.3)."""
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QPainter
from PySide6.QtWidgets import QWidget

from knob.ui.theme import AMBER, GREEN, RED
from knob.ui.widgets.fader import V_MARGIN

BAR_W = 3


class VuMeter(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._level = 0.0
        self._peak = 0.0
        self._active = True
        self.setFixedWidth(5)
        self.setMinimumHeight(100)

    def set_level(self, level: float, peak: float) -> None:
        if abs(level - self._level) > 0.002 or abs(peak - self._peak) > 0.002:
            self._level, self._peak = level, peak
            self.update()

    def set_active(self, active: bool) -> None:
        if active != self._active:
            self._active = active
            self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        x = (self.width() - BAR_W) / 2
        top, bottom = V_MARGIN, self.height() - V_MARGIN
        usable = bottom - top
        p.setBrush(QColor(255, 255, 255, 15))
        p.drawRoundedRect(QRectF(x, top, BAR_W, usable), 1.5, 1.5)
        if not self._active:
            return
        h = self._level * usable
        if h > 0.5:
            color = RED if self._level > 0.85 else AMBER if self._level > 0.70 else GREEN
            p.setBrush(QColor(color))
            p.drawRoundedRect(QRectF(x, bottom - h, BAR_W, h), 1.5, 1.5)
        if self._peak > 0.01:
            p.setBrush(QColor(255, 255, 255, 170))
            p.drawRect(QRectF(x, bottom - self._peak * usable - 1, BAR_W, 2))
