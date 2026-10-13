# -*- coding: utf-8 -*-
"""Fader vertical de solo lectura: pista, relleno con degradé y brillo, tapa blanca."""
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QLinearGradient, QPainter
from PySide6.QtWidgets import QWidget

from knob.ui.theme import qcolor

TRACK_W = 8
CAP_W = 22
CAP_H = 10
V_MARGIN = CAP_H / 2 + 2      # VuMeter usa el mismo margen para alinearse


class Fader(QWidget):
    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme
        self._v = 0.0
        self._muted = False
        self._active = True
        self.setFixedWidth(28)
        self.setMinimumHeight(100)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def set_value(self, v: float) -> None:
        v = max(0.0, min(1.0, v))
        if abs(v - self._v) > 0.0005:
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
        p.setPen(Qt.PenStyle.NoPen)
        cx = self.width() / 2
        top, bottom = V_MARGIN, self.height() - V_MARGIN
        usable = bottom - top
        track = QRectF(cx - TRACK_W / 2, top, TRACK_W, usable)
        p.setBrush(QColor(255, 255, 255, 20))
        p.drawRoundedRect(track, TRACK_W / 2, TRACK_W / 2)

        y = bottom - self._v * usable
        lit = self._active and not self._muted
        if self._active:
            fill = QRectF(track.left(), y, TRACK_W, bottom - y)
            if fill.height() > 0.5:
                if lit:
                    for grow, alpha in ((5.0, 26), (2.5, 46)):          # brillo
                        p.setBrush(qcolor(self._t.grad_a, alpha))
                        p.drawRoundedRect(fill.adjusted(-grow, -grow, grow, grow),
                                          TRACK_W / 2 + grow, TRACK_W / 2 + grow)
                    g = QLinearGradient(0, bottom, 0, top)
                    g.setColorAt(0.0, QColor(self._t.grad_a))
                    g.setColorAt(1.0, QColor(self._t.grad_b))
                    p.setBrush(g)
                else:
                    p.setBrush(QColor(255, 255, 255, 45))
                p.drawRoundedRect(fill, TRACK_W / 2, TRACK_W / 2)

        cap = QRectF(cx - CAP_W / 2, y - CAP_H / 2, CAP_W, CAP_H)
        if lit:
            p.setBrush(qcolor(self._t.brand_a, 70))
            p.drawRoundedRect(cap.adjusted(-3, -3, 3, 3), 8, 8)
            p.setBrush(QColor("#FFFFFF"))
        else:
            p.setBrush(QColor(255, 255, 255, 90))
        p.drawRoundedRect(cap, CAP_H / 2, CAP_H / 2)
