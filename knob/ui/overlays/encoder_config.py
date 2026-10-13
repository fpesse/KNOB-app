# -*- coding: utf-8 -*-
"""Panel de acciones del botón de un encoder (click, doble click, mantener y girar)."""
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont
from PySide6.QtWidgets import QComboBox, QHBoxLayout, QLabel, QPushButton, QVBoxLayout

from knob.ui.overlays.base import Panel, section_label
from knob.ui.theme import ui_font

ENCODER_CLICK_ACTIONS = [
    ("none", "Ninguna"),
    ("mute", "Mute / Unmute"),
    ("play_pause", "Play / Pausa"),
    ("next_track", "Siguiente pista"),
    ("prev_track", "Pista anterior"),
    ("vol_up", "Vol. general +5%"),
    ("vol_down", "Vol. general -5%"),
    ("hotkey", "Atajo de teclado"),
]
ENCODER_TURN_CHANNELS = [
    ("none", "Canal asignado (normal)"),
    ("master", "Volumen general"),
    ("mic", "Micrófono"),
    ("focus", "App en foco"),
    ("output_device", "Cambiar salida de audio"),
]
ENCODER_NAMES = ("Encoder izquierdo", "Encoder derecho")

_K = Qt.Key
_SPECIAL = {
    _K.Key_Space.value: "space", _K.Key_Return.value: "enter", _K.Key_Enter.value: "enter",
    _K.Key_Escape.value: "esc", _K.Key_Tab.value: "tab", _K.Key_Backspace.value: "backspace",
    _K.Key_Delete.value: "delete", _K.Key_Insert.value: "insert", _K.Key_Home.value: "home",
    _K.Key_End.value: "end", _K.Key_PageUp.value: "pageup", _K.Key_PageDown.value: "pagedown",
    _K.Key_Up.value: "up", _K.Key_Down.value: "down", _K.Key_Left.value: "left", _K.Key_Right.value: "right",
}
MODIFIER_KEYS = {_K.Key_Control.value, _K.Key_Shift.value, _K.Key_Alt.value, _K.Key_Meta.value,
                 _K.Key_AltGr.value, _K.Key_Super_L.value, _K.Key_Super_R.value}


def _int(k) -> int:
    return k.value if hasattr(k, "value") else int(k)


def key_name(key: int, text: str):
    """Nombre de tecla en el formato que entiende system.send_hotkey (igual que v3.3)."""
    if key in _SPECIAL:
        return _SPECIAL[key]
    if _K.Key_F1.value <= key <= _K.Key_F12.value:
        return f"f{key - _K.Key_F1.value + 1}"
    if 0x20 < key < 0x7F:
        return chr(key).lower()
    if text and len(text) == 1 and text.isprintable():
        return text.lower()
    return None


def build_combo(ctrl: bool, shift: bool, alt: bool, win: bool, key: int, text: str):
    if key in MODIFIER_KEYS:
        return None
    name = key_name(key, text)
    if not name:
        return None
    mods = [m for m, on in (("ctrl", ctrl), ("shift", shift), ("alt", alt), ("win", win)) if on]
    return "+".join(mods + [name])


class HotkeyButton(QPushButton):
    changed = Signal(str)

    def __init__(self, combo: str, parent=None):
        super().__init__(parent)
        self._combo = combo or ""
        self._recording = False
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clicked.connect(self._start)
        self._refresh()

    def combo(self) -> str:
        return self._combo

    def _refresh(self) -> None:
        if self._recording:
            self.setText("Presioná las teclas…")
        else:
            self.setText(self._combo.replace("+", " + ").title() if self._combo else "Grabar atajo…")

    def _start(self) -> None:
        self._recording = True
        self._refresh()
        self.setFocus()
        self.grabKeyboard()

    def _stop(self) -> None:
        if self._recording:
            self._recording = False
            self.releaseKeyboard()
            self._refresh()

    def keyPressEvent(self, e):
        if not self._recording:
            super().keyPressEvent(e)
            return
        mods = e.modifiers()
        if _int(e.key()) == _K.Key_Escape.value and not (
                mods & (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.ShiftModifier
                        | Qt.KeyboardModifier.AltModifier | Qt.KeyboardModifier.MetaModifier)):
            self._stop()                      # Esc solo: cancela la grabación
            return
        combo = build_combo(bool(mods & Qt.KeyboardModifier.ControlModifier),
                            bool(mods & Qt.KeyboardModifier.ShiftModifier),
                            bool(mods & Qt.KeyboardModifier.AltModifier),
                            bool(mods & Qt.KeyboardModifier.MetaModifier),
                            _int(e.key()), e.text())
        if combo:
            self._combo = combo
            self._stop()
            self.changed.emit(combo)

    def focusOutEvent(self, e):
        self._stop()
        super().focusOutEvent(e)


class EncoderConfigPanel(Panel):
    saved = Signal(dict)
    cancelled = Signal()

    def __init__(self, enc_idx: int, turn_label: str, entry: dict, theme, parent=None):
        super().__init__(parent)
        self.combos = {}
        self.hotkey_buttons = {}
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(4)
        head = QLabel(f"{ENCODER_NAMES[enc_idx]} · Botón")
        head.setFont(ui_font(14, QFont.Weight.Bold))
        lay.addWidget(head)
        sub = QLabel(f"Girar controla: <b>{turn_label}</b> (cambialo desde la tarjeta)")
        sub.setObjectName("dim")
        sub.setFont(ui_font(11))
        sub.setWordWrap(True)
        lay.addWidget(sub)

        for atype, label in (("click", "Click"), ("dblclick", "Doble click")):
            lay.addWidget(section_label(label))
            row = QHBoxLayout()
            row.setSpacing(6)
            combo = self._make_combo(ENCODER_CLICK_ACTIONS, entry.get(atype, "none"))
            row.addWidget(combo, 1)
            hk = HotkeyButton(entry.get(f"hotkey_keys_{atype}", ""))
            hk.setVisible(entry.get(atype) == "hotkey")
            row.addWidget(hk)
            combo.currentIndexChanged.connect(
                lambda _i, c=combo, h=hk: h.setVisible(c.currentData() == "hotkey"))
            lay.addLayout(row)
            self.combos[atype] = combo
            self.hotkey_buttons[atype] = hk

        lay.addWidget(section_label("Mantener y girar"))
        turn = self._make_combo(ENCODER_TURN_CHANNELS, entry.get("click_turn", "none"))
        lay.addWidget(turn)
        self.combos["click_turn"] = turn

        buttons = QHBoxLayout()
        buttons.setContentsMargins(0, 14, 0, 0)
        buttons.addStretch(1)
        cancel = QPushButton("Cancelar")
        cancel.clicked.connect(self.cancelled)
        buttons.addWidget(cancel)
        self.save_button = QPushButton("Guardar")
        self.save_button.setObjectName("primary")
        self.save_button.clicked.connect(self._save)
        buttons.addWidget(self.save_button)
        lay.addLayout(buttons)

    @staticmethod
    def _make_combo(options, current) -> QComboBox:
        combo = QComboBox()
        combo.setCursor(Qt.CursorShape.PointingHandCursor)
        for key, label in options:
            combo.addItem(label, key)
        idx = combo.findData(current)
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        return combo

    def _save(self) -> None:
        self.saved.emit({
            "click": self.combos["click"].currentData(),
            "dblclick": self.combos["dblclick"].currentData(),
            "click_turn": self.combos["click_turn"].currentData(),
            "hotkey_keys_click": self.hotkey_buttons["click"].combo(),
            "hotkey_keys_dblclick": self.hotkey_buttons["dblclick"].combo(),
        })
