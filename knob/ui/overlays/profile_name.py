# -*- coding: utf-8 -*-
"""Panel para crear o renombrar un perfil."""
from PySide6.QtCore import Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QHBoxLayout, QLabel, QLineEdit, QPushButton, QVBoxLayout

from knob.ui.overlays.base import Panel
from knob.ui.theme import ui_font


class ProfileNamePanel(Panel):
    submitted = Signal(str)
    cancelled = Signal()

    def __init__(self, title: str, initial: str, theme, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(8)
        head = QLabel(title)
        head.setFont(ui_font(14, QFont.Weight.Bold))
        lay.addWidget(head)
        self.name_edit = QLineEdit(initial)
        self.name_edit.setPlaceholderText("ej. Gaming, Trabajo…")
        self.name_edit.selectAll()
        self.name_edit.returnPressed.connect(self._submit)
        lay.addWidget(self.name_edit)
        self.error_label = QLabel("")
        self.error_label.setObjectName("error")
        self.error_label.setFont(ui_font(11))
        self.error_label.hide()
        lay.addWidget(self.error_label)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        cancel = QPushButton("Cancelar")
        cancel.clicked.connect(self.cancelled)
        buttons.addWidget(cancel)
        self.save_button = QPushButton("Guardar")
        self.save_button.setObjectName("primary")
        self.save_button.clicked.connect(self._submit)
        buttons.addWidget(self.save_button)
        lay.addLayout(buttons)

    def showEvent(self, e):
        super().showEvent(e)
        self.name_edit.setFocus()

    def _submit(self) -> None:
        name = self.name_edit.text().strip()
        if not name:
            self.show_error("El nombre no puede estar vacío")
            return
        self.submitted.emit(name)

    def show_error(self, msg: str) -> None:
        self.error_label.setText(msg)
        self.error_label.show()
