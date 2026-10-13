# -*- coding: utf-8 -*-
from knob.core.audio import OutputDevices, is_system_app


def test_lista_negra_de_procesos_del_sistema():
    assert is_system_app("audiodg")
    assert is_system_app("MsEdgeWebView2")
    assert not is_system_app("Discord")


def test_cycle_sin_dispositivos_devuelve_none():
    assert OutputDevices().cycle(1) is None


def test_cycle_rota_indices(monkeypatch):
    od = OutputDevices()
    od.devices = [("id1", "Parlantes"), ("id2", "Auriculares")]
    od.index = 0
    monkeypatch.setattr(OutputDevices, "_set_default_async", staticmethod(lambda dev_id: None))
    assert od.cycle(1) == "Auriculares"
    assert od.cycle(1) == "Parlantes"
    assert od.cycle(-1) == "Auriculares"
    assert od.current_name() == "Auriculares"
