# -*- coding: utf-8 -*-
from PySide6.QtCore import Qt


def theme():
    from knob.ui.theme import get_theme
    return get_theme("Nebula")


def test_build_combo():
    from knob.ui.overlays.encoder_config import build_combo, key_name
    K = Qt.Key
    assert build_combo(True, True, False, False, K.Key_M.value, "\x0d") == "ctrl+shift+m"
    assert build_combo(False, False, True, False, K.Key_F4.value, "") == "alt+f4"
    assert build_combo(False, False, False, True, K.Key_Space.value, " ") == "win+space"
    assert build_combo(True, False, False, False, K.Key_Control.value, "") is None
    assert key_name(K.Key_PageDown.value, "") == "pagedown"
    assert key_name(K.Key_5.value, "5") == "5"


def test_overlay_abre_y_cierra(qapp):
    from knob.ui.main_window import MainWindow
    from knob.ui.overlays.base import Panel
    w = MainWindow(theme())
    w.show()
    closed = []
    w.overlay.closed.connect(lambda: closed.append(1))
    w.overlay.open_centered(Panel(), 300)
    assert w.overlay.is_open()
    w.overlay.open_side(Panel(side=True))         # reemplaza el panel anterior
    assert w.overlay.is_open()
    w.overlay.close_overlay()
    qapp.processEvents()
    w.overlay._finish_close()                     # sin esperar la animación
    assert not w.overlay.is_open()
    assert closed == [1]


def test_asignar_canal_emite_eleccion(qapp):
    from knob.ui.overlays.assign_channel import AssignChannelPanel
    from knob.ui.overlays.base import OptionRow
    got = []
    p = AssignChannelPanel("Canal 1", "Discord", ["Spotify"], lambda name: None, theme())
    p.chosen.connect(got.append)
    rows = p.findChildren(OptionRow)
    texts = [r.text for r in rows]
    assert texts == ["Volumen general", "Micrófono", "App en foco", "Discord", "Spotify"]
    assert [r.selected for r in rows] == [False, False, False, True, False]
    rows[4].clicked.emit()
    p.unassign_button.click()
    assert got == ["Spotify", None]
    assert not p.grab().isNull()


def test_asignar_sin_apps_muestra_mensaje(qapp):
    from knob.ui.overlays.assign_channel import AssignChannelPanel
    p = AssignChannelPanel("Canal 2", None, [], lambda name: None, theme())
    assert p.empty_label.isVisibleTo(p)
    assert p.unassign_button is None


def test_configurar_encoder_guarda(qapp):
    from knob.ui.overlays.encoder_config import EncoderConfigPanel
    entry = {"click": "mute", "dblclick": "hotkey", "click_turn": "output_device",
             "hotkey_keys_click": "", "hotkey_keys_dblclick": "ctrl+m"}
    p = EncoderConfigPanel(0, "Volumen general", entry, theme())
    assert p.hotkey_buttons["dblclick"].isVisibleTo(p)
    assert not p.hotkey_buttons["click"].isVisibleTo(p)
    saved = []
    p.saved.connect(saved.append)
    p.combos["click"].setCurrentIndex(2)              # Play / Pausa
    p.save_button.click()
    assert saved == [{"click": "play_pause", "dblclick": "hotkey", "click_turn": "output_device",
                      "hotkey_keys_click": "", "hotkey_keys_dblclick": "ctrl+m"}]


def test_nombre_de_perfil(qapp):
    from knob.ui.overlays.profile_name import ProfileNamePanel
    p = ProfileNamePanel("Nuevo perfil", "", theme())
    got = []
    p.submitted.connect(got.append)
    p.name_edit.setText("  Gaming ")
    p.save_button.click()
    assert got == ["Gaming"]
    p.show_error("Ya existe un perfil con ese nombre")
    assert p.error_label.text() == "Ya existe un perfil con ese nombre"


def test_ajustes(qapp):
    from knob.ui.overlays.base import OptionRow
    from knob.ui.overlays.settings_panel import SettingsPanel
    p = SettingsPanel(theme())
    p.set_profiles(["Default", "Gaming"], "Gaming")
    p.set_startup(True)
    p.set_tray(False)
    p.set_version("v4.0.0")
    p.set_update_state("available", "↑ v4.0.1 disponible")
    got = []
    p.profile_switch.connect(lambda n: got.append(("switch", n)))
    p.profile_delete.connect(lambda: got.append(("delete",)))
    p.theme_selected.connect(lambda n: got.append(("theme", n)))
    p.tray_toggled.connect(lambda b: got.append(("tray", b)))
    p.update_clicked.connect(lambda: got.append(("update",)))
    p.findChildren(OptionRow)[0].clicked.emit()
    p.btn_delete.click()                              # primer click: pide confirmación
    assert ("delete",) not in got
    p.btn_delete.click()                              # segundo click: borra
    p.swatches["Cobalt"].click()
    p.tg_tray.click()
    p.btn_update.click()
    assert got == [("switch", "Default"), ("delete",), ("theme", "Cobalt"), ("tray", True), ("update",)]
    p.set_profiles(["Default"], "Default")
    assert not p.btn_delete.isEnabled()
    assert not p.grab().isNull()
