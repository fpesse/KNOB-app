# -*- coding: utf-8 -*-
"""Barra inferior: salida de audio predeterminada con flechas para cambiarla."""
from PySide6.QtCore import QRectF, Qt, QVariantAnimation, Signal
from PySide6.QtGui import QPainter, QPen
from PySide6.QtWidgets import QHBoxLayout, QLabel, QToolButton

from knob.ui.theme import GLYPH_LEFT, GLYPH_RIGHT, GLYPH_VOLUME, icon_font, qcolor, ui_font
from knob.ui.widgets.glass import GlassCard


class OutputBar(GlassCard):
    prev_clicked = Signal()
    next_clicked = Signal()

    def __init__(self, theme, parent=None):
        super().__init__(parent, clickable=False)
        self._t = theme
        self._flash = 0.0
        self.setFixedHeight(40)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 5, 6, 5)
        lay.setSpacing(8)
        icon = QLabel(GLYPH_VOLUME)
        icon.setFont(icon_font(14))
        icon.setObjectName("dim")
        lay.addWidget(icon)
        self.name_label = QLabel("")
        self.name_label.setFont(ui_font(12))
        lay.addWidget(self.name_label, 1)
        for glyph, signal, tip in ((GLYPH_LEFT, self.prev_clicked, "Salida anterior"),
                                   (GLYPH_RIGHT, self.next_clicked, "Salida siguiente")):
            b = QToolButton()
            b.setFont(icon_font(11))
            b.setText(glyph)
            b.setFixedSize(28, 28)
            b.setToolTip(tip)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.clicked.connect(signal)
            lay.addWidget(b)
        self._anim = QVariantAnimation(self)
        self._anim.setStartValue(1.0)
        self._anim.setEndValue(0.0)
        self._anim.setDuration(1500)
        self._anim.valueChanged.connect(self._on_flash)
        self.set_name(None)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def set_name(self, name, flash: bool = False) -> None:
        self.name_label.setText(name or "Sin dispositivo de salida")
        self.name_label.setToolTip(name or "")
        if flash:
            self._anim.stop()
            self._anim.start()

    def _on_flash(self, value) -> None:
        self._flash = float(value)
        self.update()

    def paintEvent(self, e):
        super().paintEvent(e)
        if self._flash <= 0.01:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(QPen(qcolor(self._t.brand_a, int(200 * self._flash)), 1.5))
        p.setBrush(qcolor(self._t.grad_a, int(50 * self._flash)))
        p.drawRoundedRect(QRectF(self.rect()).adjusted(1, 1, -1, -1), 10, 10)
