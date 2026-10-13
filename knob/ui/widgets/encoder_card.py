# -*- coding: utf-8 -*-
"""Tarjeta de encoder: perilla, destino del giro y engranaje para las acciones del botón."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QFontMetrics
from PySide6.QtWidgets import QHBoxLayout, QLabel, QToolButton, QVBoxLayout

from knob.ui.theme import GLYPH_SETTINGS, RED_TEXT, icon_font, ui_font
from knob.ui.widgets.glass import GlassCard
from knob.ui.widgets.icon_badge import IconBadge
from knob.ui.widgets.knob_dial import KnobDial

CAPTIONS = ("Encoder izq.", "Encoder der.")
TITLE_MAX_PX = 120


class EncoderCard(GlassCard):
    gear_clicked = Signal()

    def __init__(self, index: int, theme, parent=None):
        super().__init__(parent)
        self.index = index
        self._t = theme
        self._assigned = False
        self._muted = False
        self._kind = None
        self.setToolTip("Click para elegir qué controla este encoder")

        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 9, 9, 9)
        lay.setSpacing(10)
        self.dial = KnobDial(theme)
        lay.addWidget(self.dial)

        col = QVBoxLayout()
        col.setSpacing(1)
        self.caption = QLabel(CAPTIONS[index])
        self.caption.setObjectName("dim")
        self.caption.setFont(ui_font(10))
        col.addWidget(self.caption)
        title_row = QHBoxLayout()
        title_row.setSpacing(5)
        self.badge = IconBadge(16)
        title_row.addWidget(self.badge)
        self.title_label = QLabel("Sin asignar")
        self.title_label.setFont(ui_font(13, QFont.Weight.DemiBold))
        title_row.addWidget(self.title_label, 1)
        col.addLayout(title_row)
        self.sub_label = QLabel("")
        self.sub_label.setObjectName("dim")
        self.sub_label.setFont(ui_font(10))
        self.sub_label.hide()
        col.addWidget(self.sub_label)
        lay.addLayout(col, 1)

        self.gear = QToolButton()
        self.gear.setFont(icon_font(13))
        self.gear.setText(GLYPH_SETTINGS)
        self.gear.setFixedSize(26, 26)
        self.gear.setCursor(Qt.CursorShape.PointingHandCursor)
        self.gear.setToolTip("Acciones del botón (click, doble click, mantener)")
        self.gear.clicked.connect(self.gear_clicked)
        lay.addWidget(self.gear, 0, Qt.AlignmentFlag.AlignTop)
        self.set_assignment(None, None, None)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.dial.set_theme(theme)
        self.badge.set_theme(theme)

    def set_assignment(self, title, pixmap, kind) -> None:
        self._assigned = title is not None
        self._kind = kind if self._assigned else None
        self.set_dashed(not self._assigned)
        self.dial.set_active(self._assigned)
        if not self._assigned:
            self.badge.set_content("", None, "unassigned")
            self.title_label.setText("Sin asignar")
            self.sub_label.hide()
            return
        self.badge.set_content(title, pixmap, kind)
        fm = QFontMetrics(self.title_label.font())
        self.title_label.setText(fm.elidedText(title, Qt.TextElideMode.ElideRight, TITLE_MAX_PX))
        self.title_label.setToolTip(title)
        if kind == "focus":
            self.sub_label.setText("→ ninguna")
            self.sub_label.show()
        else:
            self.sub_label.hide()

    def set_focus_target(self, name) -> None:
        if self._kind != "focus":
            return
        self.sub_label.setText(f"→ {name}" if name else "→ ninguna")
        self.sub_label.show()

    def set_value(self, raw: float) -> None:
        self.dial.set_value(raw / 1023.0)

    def set_muted(self, muted: bool) -> None:
        if muted == self._muted:
            return
        self._muted = muted
        self.dial.set_muted(muted)
        self.badge.set_dimmed(muted)
        self.title_label.setStyleSheet(f"color: {RED_TEXT};" if muted else "")
