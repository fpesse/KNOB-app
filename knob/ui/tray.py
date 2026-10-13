# -*- coding: utf-8 -*-
"""Ícono en la bandeja del sistema (se muestra solo mientras la ventana está oculta, como v3.3)."""
from PySide6.QtCore import QObject, Signal
from PySide6.QtWidgets import QMenu, QSystemTrayIcon


class Tray(QObject):
    show_requested = Signal()
    quit_requested = Signal()

    def __init__(self, icon, parent=None):
        super().__init__(parent)
        self._tray = QSystemTrayIcon(icon)
        self._tray.setToolTip("KNOB")
        self._menu = QMenu()
        self._menu.addAction("Mostrar KNOB").triggered.connect(lambda: self.show_requested.emit())
        self._menu.addAction("Salir").triggered.connect(lambda: self.quit_requested.emit())
        self._tray.setContextMenu(self._menu)
        self._tray.activated.connect(self._on_activated)

    @staticmethod
    def available() -> bool:
        return QSystemTrayIcon.isSystemTrayAvailable()

    def _on_activated(self, reason) -> None:
        if reason in (QSystemTrayIcon.ActivationReason.Trigger,
                      QSystemTrayIcon.ActivationReason.DoubleClick):
            self.show_requested.emit()

    def show(self) -> None:
        self._tray.show()

    def hide(self) -> None:
        self._tray.hide()
