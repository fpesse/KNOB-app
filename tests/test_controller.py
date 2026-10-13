# -*- coding: utf-8 -*-
import pytest
from PySide6.QtCore import QTimer

from knob.core import config as cfg, updater


@pytest.fixture
def ctrl(qapp, monkeypatch, tmp_path):
    monkeypatch.setattr(cfg.Config, "save", lambda self: None)
    orig = cfg.Config.__init__
    monkeypatch.setattr(cfg.Config, "__init__", lambda self, path=None: orig(self, str(tmp_path / "c.json")))
    from knob.controller import Controller
    c = Controller(qapp)
    yield c
    c.window.hide()


def test_quit_no_reentra(ctrl, monkeypatch):
    calls = []
    monkeypatch.setattr(ctrl.serial, "stop", lambda: calls.append("stop"))
    monkeypatch.setattr(ctrl.tray, "show", lambda: calls.append("tray.show"))
    monkeypatch.setattr(ctrl.tray, "hide", lambda: calls.append("tray.hide"))
    monkeypatch.setattr(ctrl.app, "quit", lambda: (calls.append("app.quit"), ctrl.window.close()))
    ctrl.config.data["minimize_to_tray"] = True
    monkeypatch.setattr(type(ctrl.config), "minimize_to_tray", True, raising=False)
    ctrl.window.show()
    ctrl.quit()
    assert calls.count("stop") == 1
    assert "tray.hide" in calls and "tray.show" not in calls
    assert calls[-1] == "app.quit"
    n = len(calls)
    ctrl.quit()
    assert len(calls) == n


def test_instalador_falla_no_sale(ctrl, monkeypatch):
    def boom(path):
        raise OSError("uac")
    monkeypatch.setattr(updater, "launch_installer", boom)
    scheduled = []
    monkeypatch.setattr(QTimer, "singleShot", lambda *a: scheduled.append(a))
    ctrl._on_update_ready("x.exe")
    assert scheduled == []
    assert ctrl._update_state[0] == "error"


def test_cambiar_perfil_actualiza_motor_y_boton(ctrl):
    ctrl.config.create_profile("Juegos", {0: "Discord", 3: "master"}, [])
    ctrl.config.switch_profile("Default")
    ctrl._switch_profile("Juegos")
    assert ctrl.engine.channel_assignments[0] == "Discord"
    assert "Juegos" in ctrl.window.profile_btn.text()


def test_crear_perfil_desde_el_panel_y_borrar(ctrl):
    ctrl.engine.channel_assignments[1] = "Opera"
    ctrl._open_profile_name("new")
    panel = ctrl.window.overlay.panel()
    panel.name_edit.setText("Trabajo")
    panel.submitted.emit("Trabajo")
    assert ctrl.config.active_profile == "Trabajo"
    assert ctrl.config.channel_assignments()[1] == "Opera"
    assert "Trabajo" in ctrl.window.profile_btn.text()
    ctrl._delete_profile()
    assert ctrl.config.profile_names() == ["Default"]
    assert "Default" in ctrl.window.profile_btn.text()


def test_assign_persiste_y_refresca_tarjeta(ctrl):
    ctrl._assign(0, "master")
    assert ctrl.config.channel_assignments()[0] == "master"
    assert ctrl.window.card_for(0).name_label.toolTip() == "Volumen general"


def test_efectos_no_lanzan_si_el_backend_falla(ctrl, monkeypatch):
    from knob.core import system

    def boom(*a):
        raise RuntimeError("x")
    monkeypatch.setattr(ctrl.audio, "set_volume", boom)
    monkeypatch.setattr(system, "send_media_key", boom)
    ctrl.set_volume("master", 0.5)
    ctrl.media_key("play_pause")


def test_assign_con_fallo_de_guardado_conserva_memoria(ctrl, monkeypatch):
    def boom(self, *a):
        raise OSError("disco lleno")
    monkeypatch.setattr(cfg.Config, "set_channel_assignment", boom)
    ctrl._assign(2, "mic")
    assert ctrl.engine.channel_assignments[2] == "mic"


def test_estado_ocupado_texto_corto_y_tooltip(ctrl):
    ctrl._on_status("busy", "COM12")
    assert ctrl.window.status._text == "COM12 ocupado"
    assert "Serial Monitor de Arduino (COM12 ocupado)" in ctrl.window.status.toolTip()
