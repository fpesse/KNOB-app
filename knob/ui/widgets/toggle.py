# -*- coding: utf-8 -*-
"""Interruptor on/off con el degradé del tema."""
from PySide6.QtCore import QPointF, QRectF, QSize, Qt
from PySide6.QtGui import QColor, QLinearGradient, QPainter
from PySide6.QtWidgets import QAbstractButton


class Toggle(QAbstractButton):
    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme
        self.setCheckable(True)
        self.setFixedSize(34, 19)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def sizeHint(self):
        return QSize(34, 19)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen)
        r = QRectF(self.rect())
        if self.isChecked():
            g = QLinearGradient(0, 0, r.width(), 0)
            g.setColorAt(0, QColor(self._t.grad_a))
            g.setColorAt(1, QColor(self._t.grad_b))
            p.setBrush(g)
        else:
            p.setBrush(QColor(255, 255, 255, 38))
        p.drawRoundedRect(r, r.height() / 2, r.height() / 2)
        knob_x = r.width() - r.height() / 2 if self.isChecked() else r.height() / 2
        p.setBrush(QColor("#FFFFFF"))
        p.drawEllipse(QPointF(knob_x, r.height() / 2), 7, 7)
