# -*- coding: utf-8 -*-
from knob.core import protocol
from knob.core.engine import Engine, raw_to_level
from tests.fakes import FakeClock, FakeEffects


def lv(raw):
    return round(raw_to_level(raw), 4)


def make(channels, enc0):
    fx, clock = FakeEffects(), FakeClock()
    eng = Engine(fx, lambda: "", clock=clock)
    eng.set_assignments(channels, [enc0, {}])
    return eng, fx, clock


def settle(eng, clock, v3):
    """Manda un frame con el dial izquierdo en v3 y avanza un tick de 50 ms."""
    eng.on_frame([512, 512, 512, v3, 512])
    clock.t += 0.05
    eng.tick()


def test_click_play_pause_y_ninguna():
    eng, fx, _ = make({}, {"click": "play_pause", "dblclick": "none"})
    eng.on_button(0, protocol.BTN_CLICK)
    eng.on_button(0, protocol.BTN_DBLCLICK)
    assert fx.calls == [("media_key", "play_pause")]


def test_doble_click_atajo_de_teclado():
    eng, fx, _ = make({}, {"dblclick": "hotkey", "hotkey_keys_dblclick": "ctrl+shift+m"})
    eng.on_button(0, protocol.BTN_DBLCLICK)
    assert fx.calls == [("send_hotkey", "ctrl+shift+m")]


def test_atajo_vacio_no_hace_nada():
    eng, fx, _ = make({}, {"click": "hotkey", "hotkey_keys_click": ""})
    eng.on_button(0, protocol.BTN_CLICK)
    assert fx.calls == []


def test_vol_up_y_down():
    eng, fx, _ = make({}, {"click": "vol_up", "dblclick": "vol_down"})
    eng.on_button(0, protocol.BTN_CLICK)
    eng.on_button(0, protocol.BTN_DBLCLICK)
    assert fx.calls == [("step_master", 0.05), ("step_master", -0.05)]


def test_mute_guarda_y_restaura_e_ignora_hardware():
    eng, fx, clock = make({3: "master"}, {"click": "mute"})
    settle(eng, clock, 700)
    eng.on_button(0, protocol.BTN_CLICK)              # mute
    assert eng.dial_muted[0] is True
    assert eng.smooth_vals[3] == 0.0
    assert fx.calls[-1] == ("set_volume", "master", 0.0)
    settle(eng, clock, 800)                            # mover el encoder muteado no hace nada
    assert eng.target_vals[3] == 0.0
    assert eng.is_muted(3) is True
    eng.on_button(0, protocol.BTN_CLICK)              # unmute
    assert eng.dial_muted[0] is False
    assert eng.smooth_vals[3] == 700.0
    assert fx.calls[-1] == ("set_volume", "master", lv(700))


def test_mute_sin_asignacion_no_hace_nada():
    eng, fx, _ = make({}, {"click": "mute"})
    eng.on_button(0, protocol.BTN_CLICK)
    assert eng.dial_muted[0] is False
    assert fx.calls == []


def test_mantener_y_girar_comportamiento_v33():
    """v3.3: mientras se mantiene, el dial se congela y el destino de click_turn recibe
    el valor CONGELADO (no el giro). Al soltar se restaura y el próximo frame salta.
    Se conserva tal cual a propósito; ver la nota al final del plan."""
    eng, fx, clock = make({3: "master"}, {"click_turn": "mic"})
    settle(eng, clock, 600)
    eng.on_button(0, protocol.BTN_HELD_START)
    assert eng.hold_saved[0] == 600.0
    fx.calls.clear()
    settle(eng, clock, 900)
    assert eng.smooth_vals[3] == 600.0
    assert ("set_volume", "mic", lv(600)) in fx.calls
    assert ("set_volume", "master", lv(900)) not in fx.calls
    eng.on_button(0, protocol.BTN_HELD_END)
    assert eng.smooth_vals[3] == 600.0 and eng.target_vals[3] == 600.0
    assert eng.hold_saved[0] is None
    settle(eng, clock, 900)
    assert eng.smooth_vals[3] == 900.0


def test_mantener_sin_accion_mueve_normal():
    eng, fx, clock = make({3: "master"}, {"click_turn": "none"})
    settle(eng, clock, 600)
    eng.on_button(0, protocol.BTN_HELD_START)
    settle(eng, clock, 900)
    assert eng.smooth_vals[3] == 900.0
    assert ("set_volume", "master", lv(900)) in fx.calls


def test_mantener_y_girar_cambia_salida_con_antirrebote():
    eng, fx, clock = make({3: "master"}, {"click_turn": "output_device"})
    settle(eng, clock, 500)
    eng.on_button(0, protocol.BTN_HELD_START)
    fx.calls.clear()
    settle(eng, clock, 600)            # primera lectura: solo guarda referencia
    settle(eng, clock, 700)            # +100 (>35): siguiente dispositivo
    settle(eng, clock, 800)            # +100 pero dentro de 800 ms: ignorado
    clock.t += 1.0
    settle(eng, clock, 600)            # -200: dispositivo anterior
    assert [c for c in fx.calls if c[0] == "cycle_output"] == [("cycle_output", 1),
                                                                ("cycle_output", -1)]
    assert not any(c[0] == "set_volume" and c[1] == "master" for c in fx.calls)
    eng.on_button(0, protocol.BTN_HELD_END)
    assert eng.smooth_vals[3] == 500.0
