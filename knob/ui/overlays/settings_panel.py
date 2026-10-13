# -*- coding: utf-8 -*-
"""Panel de ajustes deslizable: perfiles, tema, sistema y versión."""
from PySide6.QtCore import QPointF, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QFont, QLinearGradient, QPainter, QPen
from PySide6.QtWidgets import (QAbstractButton, QHBoxLayout, QLabel, QPushButton, QScrollArea,
                               QToolButton, QVBoxLayout, QWidget)

from knob.ui.overlays.base import OptionRow, Panel, section_label
from knob.ui.theme import GLYPH_CLOSE, THEME_ORDER, THEMES, icon_font, ui_font
from knob.ui.widgets.toggle import Toggle

DELETE_CONFIRM_MS = 4000


class SwatchButton(QAbstractButton):
    def __init__(self, theme, parent=None):
        super().__init__(parent)
        self._swatch = theme
        self.selected = False
        self.setFixedSize(30, 30)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip(theme.name)

    def sizeHint(self):
        return QSize(30, 30)

    def set_selected(self, on: bool) -> None:
        self.selected = on
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        c = QPointF(self.width() / 2, self.height() / 2)
        g = QLinearGradient(0, 0, self.width(), self.height())
        g.setColorAt(0, QColor(self._swatch.grad_a))
        g.setColorAt(1, QColor(self._swatch.grad_b))
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(g)
        p.drawEllipse(c, 11, 11)
        if self.selected:
            p.setPen(QPen(QColor("#FFFFFF"), 2))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawEllipse(c, 14, 14)


class SettingsPanel(Panel):
    profile_switch = Signal(str)
    profile_new = Signal()
    profile_rename = Signal()
    profile_delete = Signal()
    theme_selected = Signal(str)
    startup_toggled = Signal(bool)
    tray_toggled = Signal(bool)
    update_clicked = Signal()
    close_requested = Signal()

    def __init__(self, theme, parent=None):
        super().__init__(parent, side=True)
        self._t = theme
        self._active = ""
        self._confirm_delete = False
        self._confirm_timer = QTimer(self)
        self._confirm_timer.setSingleShot(True)
        self._confirm_timer.setInterval(DELETE_CONFIRM_MS)
        self._confirm_timer.timeout.connect(self._reset_delete)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(4)

        head = QHBoxLayout()
        title = QLabel("Ajustes")
        title.setFont(ui_font(16, QFont.Weight.Bold))
        head.addWidget(title)
        head.addStretch(1)
        close = QToolButton()
        close.setFont(icon_font(10))
        close.setText(GLYPH_CLOSE)
        close.setFixedSize(26, 26)
        close.setCursor(Qt.CursorShape.PointingHandCursor)
        close.clicked.connect(self.close_requested)
        head.addWidget(close)
        lay.addLayout(head)

        lay.addWidget(section_label("Perfiles"))
        self._profiles_widget = QWidget()
        self._profiles_layout = QVBoxLayout(self._profiles_widget)
        self._profiles_layout.setContentsMargins(0, 0, 4, 0)
        self._profiles_layout.setSpacing(2)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self._profiles_widget)
        scroll.setMaximumHeight(150)
        self._profiles_scroll = scroll
        lay.addWidget(scroll)
        row = QHBoxLayout()
        row.setSpacing(6)
        self.btn_new = QPushButton("+ Nuevo")
        self.btn_rename = QPushButton("Renombrar")
        self.btn_delete = QPushButton("Borrar")
        self.btn_delete.setObjectName("danger")
        self.btn_new.clicked.connect(self.profile_new)
        self.btn_rename.clicked.connect(self.profile_rename)
        self.btn_delete.clicked.connect(self._on_delete)
        for b in (self.btn_new, self.btn_rename, self.btn_delete):
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            row.addWidget(b)
        row.addStretch(1)
        lay.addLayout(row)

        lay.addWidget(section_label("Tema"))
        sw_row = QHBoxLayout()
        sw_row.setSpacing(6)
        self.swatches = {}
        for name in THEME_ORDER:
            b = SwatchButton(THEMES[name])
            b.clicked.connect(lambda _=False, n=name: self.theme_selected.emit(n))
            sw_row.addWidget(b)
            self.swatches[name] = b
        sw_row.addStretch(1)
        lay.addLayout(sw_row)

        lay.addWidget(section_label("Sistema"))
        self.tg_startup = Toggle(theme)
        self.tg_tray = Toggle(theme)
        self.tg_startup.toggled.connect(self.startup_toggled)
        self.tg_tray.toggled.connect(self.tray_toggled)
        lay.addLayout(self._setting_row("Iniciar con Windows", self.tg_startup))
        lay.addLayout(self._setting_row("Al cerrar, ir a la bandeja", self.tg_tray))

        lay.addWidget(section_label("Acerca de"))
        about = QHBoxLayout()
        self.lbl_version = QLabel("")
        self.lbl_version.setFont(ui_font(12))
        about.addWidget(self.lbl_version)
        about.addStretch(1)
        self.btn_update = QPushButton("")
        self.btn_update.setObjectName("primary")
        self.btn_update.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_update.clicked.connect(self.update_clicked)
        self.btn_update.hide()
        about.addWidget(self.btn_update)
        lay.addLayout(about)
        lay.addStretch(1)
        self.set_theme(theme)

    @staticmethod
    def _setting_row(text: str, toggle) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setContentsMargins(0, 4, 0, 4)
        lbl = QLabel(text)
        lbl.setFont(ui_font(12))
        row.addWidget(lbl)
        row.addStretch(1)
        row.addWidget(toggle)
        return row

    # --- API ---------------------------------------------------------------
    def set_profiles(self, names, active: str) -> None:
        self._active = active
        while self._profiles_layout.count():
            item = self._profiles_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for name in names:
            r = OptionRow(name, self._t, selected=(name == active))
            r.clicked.connect(lambda n=name: self.profile_switch.emit(n))
            self._profiles_layout.addWidget(r)
        self._profiles_scroll.setFixedHeight(min(150, 36 * len(names) + 4))
        self.btn_delete.setEnabled(len(names) > 1)
        self._reset_delete()

    def set_theme(self, theme) -> None:
        self._t = theme
        self.tg_startup.set_theme(theme)
        self.tg_tray.set_theme(theme)
        for name, b in self.swatches.items():
            b.set_selected(name == theme.name)
        for r in self._profiles_widget.findChildren(OptionRow):
            r.set_theme(theme)

    def set_startup(self, on: bool) -> None:
        self.tg_startup.blockSignals(True)
        self.tg_startup.setChecked(on)
        self.tg_startup.blockSignals(False)

    def set_tray(self, on: bool) -> None:
        self.tg_tray.blockSignals(True)
        self.tg_tray.setChecked(on)
        self.tg_tray.blockSignals(False)

    def set_version(self, version: str) -> None:
        self.lbl_version.setText(f"KNOB {version}")

    def set_update_state(self, state: str, text: str) -> None:
        self.btn_update.setVisible(state != "none")
        self.btn_update.setText(text)
        self.btn_update.setEnabled(state == "available")

    # --- borrar con confirmación -----------------------------------------------
    def _on_delete(self) -> None:
        if not self._confirm_delete:
            self._confirm_delete = True
            self.btn_delete.setText(f"¿Borrar «{self._active.replace('&', '&&')}»?")
            self._confirm_timer.start()
            return
        self._reset_delete()
        self.profile_delete.emit()

    def _reset_delete(self) -> None:
        self._confirm_timer.stop()
        self._confirm_delete = False
        self.btn_delete.setText("Borrar")
