# -*- coding: utf-8 -*-
"""Panel 'Canal N · ¿Qué controla?': destinos de sistema + apps con audio."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QLabel, QPushButton, QScrollArea, QVBoxLayout, QWidget

from knob.ui.overlays.base import OptionRow, Panel, section_label
from knob.ui.theme import ui_font

SYSTEM_OPTIONS = (("master", "Volumen general", ""),
                  ("mic", "Micrófono", ""),
                  ("focus", "App en foco", "la ventana activa"))


class AssignChannelPanel(Panel):
    chosen = Signal(object)          # str con el destino, o None para quitar la asignación

    def __init__(self, title: str, current, apps: list, icon_for, theme, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 14, 14, 12)
        lay.setSpacing(2)
        head = QLabel(f"{title} · ¿Qué controla?")
        head.setTextFormat(Qt.TextFormat.PlainText)
        head.setFont(ui_font(14, QFont.Weight.Bold))
        lay.addWidget(head)

        lay.addWidget(section_label("Sistema"))
        for key, label, hint in SYSTEM_OPTIONS:
            row = OptionRow(label, theme, badge=("", None, key), selected=(current == key), hint=hint)
            row.clicked.connect(lambda k=key: self.chosen.emit(k))
            lay.addWidget(row)

        lay.addWidget(section_label("Apps con audio ahora"))
        names = list(apps)
        if current and current not in ("master", "mic", "focus") and current not in names:
            names.insert(0, current)          # la app asignada aunque no esté sonando
        inner = QWidget()
        inner_lay = QVBoxLayout(inner)
        inner_lay.setContentsMargins(0, 0, 4, 0)
        inner_lay.setSpacing(2)
        for name in names:
            row = OptionRow(name, theme, badge=(name, icon_for(name), None), selected=(current == name))
            row.clicked.connect(lambda n=name: self.chosen.emit(n))
            inner_lay.addWidget(row)
        self.empty_label = QLabel("No hay apps reproduciendo audio.\nAbrí la app y volvé a intentar.")
        self.empty_label.setObjectName("dim")
        self.empty_label.setFont(ui_font(11))
        self.empty_label.setVisible(not names)
        inner_lay.addWidget(self.empty_label)
        inner_lay.addStretch(1)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(inner)
        scroll.setFixedHeight(min(220, max(48, 36 * len(names) + 8)))
        lay.addWidget(scroll)

        self.unassign_button = None
        if current is not None:
            self.unassign_button = QPushButton("Quitar asignación")
            self.unassign_button.setObjectName("link")
            self.unassign_button.clicked.connect(lambda: self.chosen.emit(None))
            lay.addWidget(self.unassign_button)
