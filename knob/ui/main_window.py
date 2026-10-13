# -*- coding: utf-8 -*-
"""Ventana única de KNOB: barra superior, 3 canales, 2 encoders y salida de audio."""
from PySide6.QtCore import QPoint, QPointF, Qt, Signal
from PySide6.QtGui import QColor, QPainter
from PySide6.QtWidgets import QHBoxLayout, QMainWindow, QMenu, QPushButton, QToolButton, QVBoxLayout

from knob.core import system
from knob.ui.theme import GLYPH_SETTINGS, icon_font
from knob.ui.widgets.brand import BrandLabel
from knob.ui.widgets.channel_card import ChannelCard
from knob.ui.widgets.encoder_card import EncoderCard
from knob.ui.widgets.glass import Backdrop
from knob.ui.widgets.output_bar import OutputBar
from knob.ui.widgets.status_pill import StatusPill

WINDOW_W, WINDOW_H = 520, 620


class GearButton(QToolButton):
    """Engranaje de ajustes con un punto cuando hay actualización."""

    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme
        self._badge = False
        self.setFont(icon_font(14))
        self.setText(GLYPH_SETTINGS)
        self.setFixedSize(30, 30)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip("Ajustes")

    def set_theme(self, theme) -> None:
        self._t = theme
        self.update()

    def set_badge(self, on: bool) -> None:
        self._badge = on
        self.setToolTip("Ajustes · hay una actualización disponible" if on else "Ajustes")
        self.update()

    def paintEvent(self, e):
        super().paintEvent(e)
        if self._badge:
            p = QPainter(self)
            p.setRenderHint(QPainter.RenderHint.Antialiasing)
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(self._t.grad_b))
            p.drawEllipse(QPointF(self.width() - 6, 6), 3.5, 3.5)


class MainWindow(QMainWindow):
    channel_clicked = Signal(int)
    encoder_gear_clicked = Signal(int)
    settings_clicked = Signal()
    profile_selected = Signal(str)
    output_prev = Signal()
    output_next = Signal()
    close_requested = Signal()

    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._t = theme
        self.setWindowTitle("KNOB")
        self.setFixedSize(WINDOW_W, WINDOW_H)

        self.backdrop = Backdrop(theme)
        self.setCentralWidget(self.backdrop)
        root = QVBoxLayout(self.backdrop)
        root.setContentsMargins(16, 14, 16, 16)
        root.setSpacing(12)

        top = QHBoxLayout()
        top.setSpacing(10)
        self.brand = BrandLabel(theme)
        top.addWidget(self.brand, 0, Qt.AlignmentFlag.AlignVCenter)
        self.status = StatusPill()
        top.addWidget(self.status, 0, Qt.AlignmentFlag.AlignVCenter)
        top.addStretch(1)
        self.profile_btn = QPushButton("Perfil")
        self.profile_btn.setObjectName("profile")
        self.profile_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._profile_menu = QMenu(self)
        self.profile_btn.clicked.connect(self._show_profile_menu)
        top.addWidget(self.profile_btn, 0, Qt.AlignmentFlag.AlignVCenter)
        self.gear_btn = GearButton(theme)
        self.gear_btn.clicked.connect(self.settings_clicked)
        top.addWidget(self.gear_btn, 0, Qt.AlignmentFlag.AlignVCenter)
        root.addLayout(top)

        chans = QHBoxLayout()
        chans.setSpacing(10)
        self.channel_cards = []
        for i in range(3):
            card = ChannelCard(i, theme)
            card.clicked.connect(lambda i=i: self.channel_clicked.emit(i))
            chans.addWidget(card)
            self.channel_cards.append(card)
        root.addLayout(chans, 1)

        encs = QHBoxLayout()
        encs.setSpacing(10)
        self.encoder_cards = []
        for i in range(2):
            card = EncoderCard(i, theme)
            card.clicked.connect(lambda i=i: self.channel_clicked.emit(3 + i))
            card.gear_clicked.connect(lambda i=i: self.encoder_gear_clicked.emit(i))
            encs.addWidget(card)
            self.encoder_cards.append(card)
        root.addLayout(encs)

        self.output_bar = OutputBar(theme)
        self.output_bar.prev_clicked.connect(self.output_prev)
        self.output_bar.next_clicked.connect(self.output_next)
        root.addWidget(self.output_bar)

    # --- API para el controlador ---------------------------------------
    def card_for(self, ch: int):
        return (self.channel_cards + self.encoder_cards)[ch]

    def set_profiles(self, names, active: str) -> None:
        self.profile_btn.setText(f"Perfil: {active}  ▾")
        self._profile_menu.clear()
        for name in names:
            act = self._profile_menu.addAction(("✓  " if name == active else "     ") + name)
            act.triggered.connect(lambda _=False, n=name: self.profile_selected.emit(n))

    def set_update_badge(self, on: bool) -> None:
        self.gear_btn.set_badge(on)

    def apply_theme(self, theme) -> None:
        self._t = theme
        self.backdrop.set_theme(theme)
        self.brand.set_theme(theme)
        self.gear_btn.set_theme(theme)
        self.output_bar.set_theme(theme)
        for card in self.channel_cards + self.encoder_cards:
            card.set_theme(theme)
        self._style_title_bar()

    # --- internos ------------------------------------------------------
    def _show_profile_menu(self) -> None:
        self._profile_menu.exec(self.profile_btn.mapToGlobal(QPoint(0, self.profile_btn.height() + 4)))

    def _style_title_bar(self) -> None:
        if self.isVisible():
            system.style_title_bar(int(self.winId()), self._t.bg)

    def showEvent(self, e):
        super().showEvent(e)
        self._style_title_bar()

    def closeEvent(self, e):
        e.ignore()                     # el controlador decide: bandeja o salir
        self.close_requested.emit()
