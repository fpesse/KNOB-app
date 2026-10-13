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
