# -*- coding: utf-8 -*-
"""Capa que oscurece la ventana y muestra paneles (centrados o deslizados desde la derecha)."""
from PySide6.QtCore import QEasingCurve, QParallelAnimationGroup, QPoint, QPropertyAnimation, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QFont, QLinearGradient, QPainter, QPen
from PySide6.QtWidgets import QFrame, QGraphicsOpacityEffect, QHBoxLayout, QLabel, QWidget

from knob.ui.theme import PANEL_BG, qcolor, ui_font
from knob.ui.widgets.icon_badge import IconBadge

ANIM_MS = 190


def section_label(text: str) -> QLabel:
    lbl = QLabel(text.upper())
    lbl.setObjectName("section")
    f = ui_font(10, QFont.Weight.DemiBold)
    f.setLetterSpacing(QFont.SpacingType.AbsoluteSpacing, 1.0)
    lbl.setFont(f)
    lbl.setContentsMargins(0, 10, 0, 2)
    return lbl


class Panel(QFrame):
    """Fondo de un panel: superficie oscura con borde suave."""

    def __init__(self, parent=None, side: bool = False):
        super().__init__(parent)
        self._side = side

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        if self._side:
            r.adjust(0, -14, 14, 14)            # solo se ven las esquinas izquierdas
        p.setPen(QPen(QColor(255, 255, 255, 23), 1))
        p.setBrush(PANEL_BG)
        p.drawRoundedRect(r, 12, 12)


class OptionRow(QWidget):
    """Fila clickeable de una lista: ícono opcional, texto, pista y tilde si está elegida."""
    clicked = Signal()

    def __init__(self, text: str, theme, badge=None, selected: bool = False, hint: str = "", parent=None):
        super().__init__(parent)
        self.text = text
        self.selected = selected
        self._t = theme
        self._hover = False
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(34)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(9, 5, 9, 5)
        lay.setSpacing(9)
        if badge is not None:
            b = IconBadge(20)
            b.set_theme(theme)
            b.set_content(*badge)
            lay.addWidget(b)
        label = QLabel(text)
        label.setFont(ui_font(12, QFont.Weight.DemiBold if selected else QFont.Weight.Normal))
        lay.addWidget(label, 1)
        if hint:
            h = QLabel(hint)
            h.setObjectName("dim")
            h.setFont(ui_font(10))
            lay.addWidget(h)
        if selected:
            lay.addWidget(QLabel("✓"))

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def enterEvent(self, e):
        self._hover = True
        self.update()

    def leaveEvent(self, e):
        self._hover = False
        self.update()

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton and self.rect().contains(e.position().toPoint()):
            self.clicked.emit()

    def mousePressEvent(self, e):
        e.accept()

    def paintEvent(self, _):
        if not (self.selected or self._hover):
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        r = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        if self.selected:
            g = QLinearGradient(0, 0, r.width(), 0)
            g.setColorAt(0, qcolor(self._t.grad_a, 115))
            g.setColorAt(1, qcolor(self._t.grad_b, 64))
            p.setBrush(g)
            p.setPen(QPen(qcolor(self._t.brand_a, 115), 1))
        else:
            p.setBrush(QColor(255, 255, 255, 14))
            p.setPen(Qt.PenStyle.NoPen)
        p.drawRoundedRect(r, 8, 8)


class OverlayHost(QWidget):
    closed = Signal()

    def __init__(self, parent):
        super().__init__(parent)
        self._panel = None
        self._mode = None
        self._anim = None
        self._closing = False
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self._effect = QGraphicsOpacityEffect(self)
        self._effect.setOpacity(1.0)
        self.setGraphicsEffect(self._effect)
        self.hide()

    def is_open(self) -> bool:
        return self._panel is not None and not self._closing

    def panel(self):
        return self._panel

    def open_centered(self, panel, width: int = 340) -> None:
        self._open(panel, "center", width)

    def open_side(self, panel) -> None:
        self._open(panel, "side", int(self.parentWidget().width() * 0.72))

    def _open(self, panel, mode: str, width: int) -> None:
        was_visible = self.isVisible() and not self._closing
        self._stop_anim()
        self._closing = False
        self._dispose_panel()
        self.setGeometry(self.parentWidget().rect())
        panel.setParent(self)
        self._panel, self._mode = panel, mode
        panel.setFixedWidth(width)
        if mode == "center":
            h = min(panel.sizeHint().height(), self.height() - 60)
            panel.setFixedHeight(h)
            end = QPoint((self.width() - width) // 2, max(30, (self.height() - h) // 3))
            start = end + QPoint(0, 14)
        else:
            panel.setFixedHeight(self.height())
            end = QPoint(self.width() - width, 0)
            start = QPoint(self.width(), 0)
        panel.move(start)
        panel.show()
        self.show()
        self.raise_()
        self.setFocus()

        group = QParallelAnimationGroup(self)
        slide = QPropertyAnimation(panel, b"pos", self)
        slide.setDuration(ANIM_MS)
        slide.setStartValue(start)
        slide.setEndValue(end)
        slide.setEasingCurve(QEasingCurve.Type.OutCubic)
        group.addAnimation(slide)
        if not was_visible:
            fade = QPropertyAnimation(self._effect, b"opacity", self)
            fade.setDuration(ANIM_MS)
            fade.setStartValue(0.0)
            fade.setEndValue(1.0)
            group.addAnimation(fade)
        else:
            self._effect.setOpacity(1.0)
        self._anim = group
        group.start()

    def close_overlay(self) -> None:
        if self._panel is None or self._closing:
            return
        self._stop_anim()
        self._closing = True
        fade = QPropertyAnimation(self._effect, b"opacity", self)
        fade.setDuration(150)
        fade.setStartValue(self._effect.opacity())
        fade.setEndValue(0.0)
        fade.finished.connect(self._finish_close)
        self._anim = fade
        fade.start()

    def _finish_close(self) -> None:
        if not self._closing:
            return
        self._closing = False
        self._stop_anim()
        self.hide()
        self._effect.setOpacity(1.0)
        self._dispose_panel()
        self.closed.emit()

    def _stop_anim(self) -> None:
        if self._anim is not None:
            if self._closing:                 # solo la animación de cierre tiene 'finished' conectado
                self._anim.finished.disconnect()
            self._anim.stop()
            self._anim = None

    def _dispose_panel(self) -> None:
        if self._panel is not None:
            self._panel.hide()
            self._panel.deleteLater()
            self._panel = None

    def paintEvent(self, _):
        QPainter(self).fillRect(self.rect(), QColor(5, 3, 10, 140))

    def mousePressEvent(self, e):
        if self._panel is not None and not self._panel.geometry().contains(e.position().toPoint()):
            self.close_overlay()
        e.accept()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key.Key_Escape:
            self.close_overlay()
            return
        super().keyPressEvent(e)
