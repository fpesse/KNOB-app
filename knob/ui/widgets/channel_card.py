# -*- coding: utf-8 -*-
"""Tarjeta de canal (faders 1-3): ícono, nombre, fader + VU y porcentaje."""
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QFontMetrics
from PySide6.QtWidgets import QHBoxLayout, QLabel, QVBoxLayout

from knob.ui.theme import RED_TEXT, ui_font
from knob.ui.widgets.fader import Fader
from knob.ui.widgets.glass import GlassCard
from knob.ui.widgets.icon_badge import IconBadge
from knob.ui.widgets.vu_meter import VuMeter

NAME_MAX_PX = 120


class ChannelCard(GlassCard):
    def __init__(self, index: int, theme, parent=None):
        super().__init__(parent)
        self.index = index
        self._t = theme
        self._assigned = False
        self._muted = False
        self._kind = None
        self._pct = None
        self.setToolTip("Click para elegir qué controla este fader")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 14, 8, 12)
        lay.setSpacing(6)
        self.badge = IconBadge(28)
        lay.addWidget(self.badge, 0, Qt.AlignmentFlag.AlignHCenter)

        self.name_label = QLabel("Sin asignar")
        self.name_label.setObjectName("dim")
        self.name_label.setFont(ui_font(12))
        self.name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self.name_label)

        self.sub_label = QLabel("")
        self.sub_label.setObjectName("dim")
        self.sub_label.setFont(ui_font(10))
        self.sub_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.sub_label.hide()
        lay.addWidget(self.sub_label)

        row = QHBoxLayout()
        row.setSpacing(3)
        row.addStretch(1)
        self.fader = Fader(theme)
        self.vu = VuMeter()
        row.addWidget(self.fader)
        row.addWidget(self.vu)
        row.addStretch(1)
        lay.addLayout(row, 1)

        self.value_label = QLabel("")
        self.value_label.setFont(ui_font(14, QFont.Weight.DemiBold))
        self.value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self.value_label)
        self.set_assignment(None, None, None)

    def set_theme(self, theme) -> None:
        self._t = theme
        self.fader.set_theme(theme)
        self.badge.set_theme(theme)

    def set_assignment(self, title, pixmap, kind) -> None:
        self._assigned = title is not None
        self._kind = kind if self._assigned else None
        self.set_dashed(not self._assigned)
        self.fader.set_active(self._assigned)
        self.vu.set_active(self._assigned)
        if not self._assigned:
            self.badge.set_content("", None, "unassigned")
            self.name_label.setText("Sin asignar")
            self.sub_label.setText("click para elegir")
            self.sub_label.show()
            self.value_label.setText("")
            self._pct = None
            return
        self.badge.set_content(title, pixmap, kind)
        fm = QFontMetrics(self.name_label.font())
        self.name_label.setText(fm.elidedText(title, Qt.TextElideMode.ElideRight, NAME_MAX_PX))
        self.name_label.setToolTip(title)
        if kind == "focus":
            self.sub_label.setText("→ ninguna")
            self.sub_label.show()
        else:
            self.sub_label.hide()
        self._pct = None
        self._render_value()

    def set_focus_target(self, name) -> None:
        if self._kind != "focus":
            return
        self.sub_label.setText(f"→ {name}" if name else "→ ninguna")
        self.sub_label.show()

    def set_value(self, raw: float) -> None:
        self.fader.set_value(raw / 1023.0)
        pct = round(raw / 1023.0 * 100)
        if pct != self._pct:
            self._pct = pct
            self._render_value()

    def set_muted(self, muted: bool) -> None:
        if muted == self._muted:
            return
        self._muted = muted
        self.fader.set_muted(muted)
        self.badge.set_dimmed(muted)
        self._render_value()

    def set_level(self, level: float, peak: float) -> None:
        self.vu.set_level(level if self._assigned else 0.0, peak if self._assigned else 0.0)

    def _render_value(self) -> None:
        if not self._assigned:
            return
        if self._muted:
            self.value_label.setText("MUTE")
            self.value_label.setStyleSheet(f"color: {RED_TEXT};")
        else:
            self.value_label.setText(f"{self._pct or 0}%")
            self.value_label.setStyleSheet("")
