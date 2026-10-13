# -*- coding: utf-8 -*-
"""Logo 'KNOB' con la fuente Amazing Views rellena con el degradé del tema."""
import math

from PySide6.QtGui import QColor, QFontMetricsF, QLinearGradient, QPainter, QPainterPath
from PySide6.QtWidgets import QWidget

from knob.ui.theme import brand_font

TEXT = "KNOB"


class BrandLabel(QWidget):
    def __init__(self, theme, px: int = 26, parent=None):
        super().__init__(parent)
        self._t = theme
        self._font = brand_font(px)
        fm = QFontMetricsF(self._font)
        self._ascent = fm.ascent()
        self.setFixedSize(math.ceil(fm.horizontalAdvance(TEXT)) + 4, math.ceil(fm.height()))

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        path = QPainterPath()
        path.addText(2, self._ascent, self._font, TEXT)
        g = QLinearGradient(0, 0, self.width(), 0)
        g.setColorAt(0.0, QColor(self._t.brand_a))
        g.setColorAt(1.0, QColor(self._t.brand_b))
        p.fillPath(path, g)
