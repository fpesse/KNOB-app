# -*- coding: utf-8 -*-
import pytest

from knob.core.engine import (DEADBAND, Engine, MUTE_THRESHOLD, linear_to_vu, next_peak,
                              raw_to_level)
from tests.fakes import FakeClock, FakeEffects


def lv(raw):
    return round(raw_to_level(raw), 4)


def make(focus=""):
    fx, clock = FakeEffects(), FakeClock()
    eng = Engine(fx, lambda: focus, clock=clock)
    return eng, fx, clock


def frame(v0=512, v1=512, v2=512, v3=512, v4=512):
    return [v0, v1, v2, v3, v4]


def test_raw_to_level_ley_cuadratica_y_mute():
    assert raw_to_level(0) == 0.0
    assert raw_to_level(MUTE_THRESHOLD) == 0.0
    assert raw_to_level(1023) == 1.0
    assert raw_to_level(511.5) == pytest.approx(0.25)


def test_linear_to_vu_y_pico():
    assert linear_to_vu(0) == 0.0
    assert linear_to_vu(1.0) == 1.0
    assert linear_to_vu(0.001) == 0.0          # -60 dB, por debajo del piso de -30
    assert next_peak(0.5, 0.2) == 0.5
    assert next_peak(0.1, 0.5) == pytest.approx(0.488)


def test_no_aplica_volumen_antes_del_hardware():
    eng, fx, _ = make()
    eng.set_assignments({0: "Discord"}, [])
    assert eng.tick() is False
    assert fx.calls == []


def test_deadband_ignora_ruido():
    eng, _, _ = make()
    eng.on_frame(frame(v0=512 + DEADBAND - 1))
    assert eng.target_vals[0] == 512.0
    eng.on_frame(frame(v0=512 + DEADBAND))
    assert eng.target_vals[0] == 512.0 + DEADBAND


def test_tick_converge_y_aplica_volumen():
    eng, fx, _ = make()
    eng.set_assignments({0: "Discord", 1: "master"}, [])
    eng.on_frame(frame(v0=800, v1=100))
    assert eng.tick() is True
    assert eng.smooth_vals[0] == 800.0
    assert ("set_volume", "Discord", lv(800)) in fx.calls
    assert ("set_volume", "master", lv(100)) in fx.calls


def test_canal_sin_asignar_no_hace_nada():
    eng, fx, _ = make()
    eng.on_frame(frame(v2=900))
    eng.tick()
    assert fx.volumes() == []


def test_throttle_40ms_y_valor_pendiente_se_aplica():
    eng, fx, clock = make()
    eng.set_assignments({0: "Spotify"}, [])
    eng.on_frame(frame(v0=600))
    eng.tick()
    clock.t += 0.010
    eng.on_frame(frame(v0=700))
    eng.tick()                                   # dentro de la ventana de 40 ms
    assert ("set_volume", "Spotify", lv(700)) not in fx.calls
    clock.t += 0.040
    assert eng.tick() is False                   # ya no se mueve...
    assert ("set_volume", "Spotify", lv(700)) in fx.calls   # ...pero el pendiente se aplica


def test_is_muted():
    eng, _, _ = make()
    eng.set_assignments({0: "Discord"}, [])
    eng.on_frame(frame(v0=10, v1=10))
    eng.tick()
    assert eng.is_muted(0) is True
    assert eng.is_muted(1) is False              # sin asignar nunca está "muteado"


def test_prioridad_ultimo_que_se_movio_gana():
    eng, fx, _ = make(focus="Discord")
    eng.set_assignments({0: "Discord", 4: "focus"}, [])
    eng.apply_channel(0, 500)
    assert fx.volumes()[-1] == ("set_volume", "Discord", lv(500))
    eng.apply_channel(4, 300)                    # el foco se mueve: toma el control
    assert fx.volumes()[-1] == ("set_volume", "Discord", lv(300))
    n = len(fx.volumes())
    eng.apply_channel(0, 520)                    # directo se mueve 20 (<=40): no recupera
    assert len(fx.volumes()) == n
    eng.apply_channel(0, 560)                    # se mueve 60 (>40): recupera
    assert fx.volumes()[-1] == ("set_volume", "Discord", lv(560))
    n = len(fx.volumes())
    eng.apply_channel(4, 305)                    # el foco apenas se mueve (<=15): no toma
    assert len(fx.volumes()) == n


def test_foco_sin_ventana_activa_no_aplica():
    eng, fx, _ = make(focus="")
    eng.set_assignments({4: "focus"}, [])
    eng.apply_channel(4, 700)
    assert fx.volumes() == []
