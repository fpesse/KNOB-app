# -*- coding: utf-8 -*-
"""Pastilla de conexión: verde conectado, ámbar buscando (punto que late), rojo puerto ocupado."""
import math

from PySide6.QtCore import QPointF, QRectF, Qt, QTimer
from PySide6.QtGui import QColor, QFontMetrics, QPainter
from PySide6.QtWidgets import QWidget

from knob.ui.theme import AMBER, AMBER_TEXT, GREEN, GREEN_TEXT, RED, RED_TEXT, qcolor, ui_font

_STYLES = {"connected": (GREEN, GREEN_TEXT), "searching": (AMBER, AMBER_TEXT), "busy": (RED, RED_TEXT)}


class StatusPill(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._font = ui_font(11)
        self._state = ""
        self._text = ""
        self._phase = 0.0
        self._timer = QTimer(self)
        self._timer.setInterval(50)
        self._timer.timeout.connect(self._pulse)
        self.set_state("searching", "Buscando KNOB…")

    def state(self) -> str:
        return self._state

    def set_state(self, state: str, text: str) -> None:
        if state == self._state and text == self._text:
            return
        self._state, self._text = state, text
        fm = QFontMetrics(self._font)
        self.setFixedSize(fm.horizontalAdvance(text) + 30, 22)
        self.setToolTip(text)
        if state == "searching":
            self._timer.start()
        else:
            self._timer.stop()
        self.update()

    def _pulse(self) -> None:
        self._phase = (self._phase + 0.04) % 1.0
        self.update()

    def paintEvent(self, _):
        dot, text_color = _STYLES.get(self._state, _STYLES["searching"])
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(qcolor(dot, 32))
        p.drawRoundedRect(QRectF(self.rect()), 11, 11)
        alpha = 255
        if self._state == "searching":
            alpha = int(110 + 145 * (0.5 + 0.5 * math.cos(2 * math.pi * self._phase)))
        center = QPointF(12, self.height() / 2)
        p.setBrush(qcolor(dot, alpha // 4))
        p.drawEllipse(center, 5.5, 5.5)
        p.setBrush(qcolor(dot, alpha))
        p.drawEllipse(center, 3, 3)
        p.setPen(QColor(text_color))
        p.setFont(self._font)
        p.drawText(QRectF(21, 0, self.width() - 21, self.height()),
                   Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, self._text)
