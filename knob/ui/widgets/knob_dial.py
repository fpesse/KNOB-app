# -*- coding: utf-8 -*-
"""Perilla estilo hardware: 5 marcas de escala, cara oscura y puntero luminoso (barrido de 270°)."""
import math

from PySide6.QtCore import QPointF, Qt
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import QWidget

from knob.ui.theme import qcolor

SIZE = 44
_TICKS = (-135.0, -67.5, 0.0, 67.5, 135.0)   # grados desde las 12, sentido horario


def _polar(cx, cy, r, deg):
    a = math.radians(deg)
    return QPointF(cx + r * math.sin(a), cy - r * math.cos(a))


class KnobDial(QWidget):
    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme
        self._v = 0.0
        self._muted = False
        self._active = True
        self.setFixedSize(SIZE, SIZE)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def set_value(self, v: float) -> None:
        v = max(0.0, min(1.0, v))
        if abs(v - self._v) > 0.001:
            self._v = v
            self.update()

    def set_muted(self, muted: bool) -> None:
        if muted != self._muted:
            self._muted = muted
            self.update()

    def set_active(self, active: bool) -> None:
        if active != self._active:
            self._active = active
            self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        cx, cy = self.width() / 2, self.height() / 2
        p.setPen(QPen(QColor(self._t.knob_ring), 2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        for deg in _TICKS:
            p.drawLine(_polar(cx, cy, 17.5, deg), _polar(cx, cy, 20.5, deg))
        p.setPen(QPen(QColor(self._t.knob_ring), 1))
        p.setBrush(QColor(self._t.knob_face))
        p.drawEllipse(QPointF(cx, cy), 13, 13)

        angle = -135.0 + self._v * 270.0
        start, end = _polar(cx, cy, 2.5, angle), _polar(cx, cy, 10.5, angle)
        lit = self._active and not self._muted
        if lit:
            p.setPen(QPen(qcolor(self._t.pointer, 70), 5, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
            p.drawLine(start, end)
            color = QColor(self._t.pointer)
        else:
            color = qcolor(self._t.text_dim, 120)
        p.setPen(QPen(color, 2.2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
        p.drawLine(start, end)
