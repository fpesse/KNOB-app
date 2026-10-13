# -*- coding: utf-8 -*-
"""Superficies translúcidas del estilo Glow."""
from PySide6.QtCore import QPointF, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QPainter, QPen, QRadialGradient
from PySide6.QtWidgets import QWidget

from knob.ui.theme import qcolor


class GlassCard(QWidget):
    """Tarjeta translúcida con borde suave. Si es clickeable emite clicked."""
    clicked = Signal()

    def __init__(self, parent=None, clickable: bool = True):
        super().__init__(parent)
        self._clickable = clickable
        self._hover = False
        self._pressed = False
        self._dashed = False
        self._radius = 10
        if clickable:
            self.setCursor(Qt.CursorShape.PointingHandCursor)

    def set_dashed(self, dashed: bool) -> None:
        if dashed != self._dashed:
            self._dashed = dashed
            self.update()

    def enterEvent(self, e):
        self._hover = True
        self.update()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._hover = False
        self.update()
        super().leaveEvent(e)

    def mousePressEvent(self, e):
        if self._clickable and e.button() == Qt.MouseButton.LeftButton:
            self._pressed = True
            e.accept()
            return
        super().mousePressEvent(e)

    def mouseReleaseEvent(self, e):
        if self._pressed and e.button() == Qt.MouseButton.LeftButton:
            self._pressed = False
            if self.rect().contains(e.position().toPoint()):
                self.clicked.emit()
            e.accept()
            return
        super().mouseReleaseEvent(e)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        hover = self._hover and self._clickable
        if self._dashed:
            p.setPen(QPen(QColor(255, 255, 255, 60 if hover else 46), 1, Qt.PenStyle.DashLine))
            p.setBrush(QColor(255, 255, 255, 10) if hover else Qt.BrushStyle.NoBrush)
        else:
            p.setPen(QPen(QColor(255, 255, 255, 30 if hover else 18), 1))
            p.setBrush(QColor(255, 255, 255, 17 if hover else 10))
        p.drawRoundedRect(r, self._radius, self._radius)


class Backdrop(QWidget):
    """Fondo de la ventana: color base + halo radial arriba a la derecha."""

    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(self._t.bg))
        w = self.width()
        g = QRadialGradient(QPointF(w * 0.78, -20), w * 0.95)
        g.setColorAt(0.0, qcolor(self._t.halo, 255))
        g.setColorAt(0.55, qcolor(self._t.halo, 70))
        g.setColorAt(1.0, qcolor(self._t.bg, 0))
        p.fillRect(self.rect(), g)
