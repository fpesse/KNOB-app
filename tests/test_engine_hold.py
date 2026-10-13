# -*- coding: utf-8 -*-
"""'Mantener y girar' con firmware nuevo: el giro llega aparte (hold_turns) y mueve el
destino de forma relativa desde su volumen actual, sin tocar la perilla."""
from knob.core import protocol
from knob.core.engine import Engine, raw_to_level
from tests.fakes import FakeClock, FakeEffects


def lv(raw):
    return round(raw_to_level(raw), 4)


def make(enc0, levels=None, focus="", channels=None):
    fx, clock = FakeEffects(), FakeClock()
    fx.levels = levels or {}
    eng = Engine(fx, lambda: focus, clock=clock)
    eng.set_assignments(channels if channels is not None else {3: "master"}, [enc0, {}])
    return eng, fx, clock


def step(eng, clock, v3=600, turn=0):
    eng.on_frame([512, 512, 512, v3, 512], [turn, 0])
    clock.t += 0.05
    eng.tick()


def test_girar_mantenido_sube_el_destino_desde_su_volumen_actual():
    eng, fx, clock = make({"click_turn": "mic"}, levels={"mic": 0.25})   # 0.25 ≈ raw 511.5
    step(eng, clock)
    eng.on_button(0, protocol.BTN_HELD_START)
    fx.calls.clear()
    step(eng, clock, turn=100)
    assert fx.volumes() == [("set_volume", "mic", lv(611.5))]
    assert eng.smooth_vals[3] == 600.0                 # la perilla no se mueve
    step(eng, clock, turn=100)                         # sin giro nuevo: no repite
    assert len(fx.volumes()) == 1


def test_al_pasarse_del_tope_rebasa_y_responde_al_volver():
    eng, fx, clock = make({"click_turn": "mic"}, levels={"mic": 0.25})
    step(eng, clock)
    eng.on_button(0, protocol.BTN_HELD_START)
    fx.calls.clear()
    step(eng, clock, turn=-700)                        # 511.5 - 700 < 0 → 0
    step(eng, clock, turn=-600)                        # vuelve 100 desde el tope
    assert fx.volumes() == [("set_volume", "mic", 0.0), ("set_volume", "mic", lv(100))]


def test_soltar_no_hace_saltar_la_perilla_ni_toca_el_destino():
    eng, fx, clock = make({"click_turn": "mic"}, levels={"mic": 0.25})
    step(eng, clock)
    eng.on_button(0, protocol.BTN_HELD_START)
    step(eng, clock, turn=200)
    eng.on_button(0, protocol.BTN_HELD_END)
    fx.calls.clear()
    step(eng, clock, turn=0)
    assert eng.smooth_vals[3] == 600.0
    assert not any(c[1] == "mic" for c in fx.volumes())


def test_mientras_se_mantiene_no_se_aplica_el_canal_propio():
    eng, fx, clock = make({"click_turn": "master"}, levels={"master": 0.25})
    step(eng, clock)
    eng.on_button(0, protocol.BTN_HELD_START)
    fx.calls.clear()
    step(eng, clock, turn=100)
    step(eng, clock, turn=150)
    assert fx.volumes() == [("set_volume", "master", lv(611.5)), ("set_volume", "master", lv(661.5))]


def test_app_en_foco_usa_la_ventana_activa():
    eng, fx, clock = make({"click_turn": "focus"}, levels={"Discord": 1.0}, focus="Discord")
    step(eng, clock)
    eng.on_button(0, protocol.BTN_HELD_START)
    fx.calls.clear()
    step(eng, clock, turn=-100)
    assert fx.volumes() == [("set_volume", "Discord", lv(923))]


def test_sin_volumen_legible_usa_la_posicion_de_la_perilla():
    eng, fx, clock = make({"click_turn": "mic"}, levels={})
    step(eng, clock, v3=400)
    eng.on_button(0, protocol.BTN_HELD_START)
    fx.calls.clear()
    step(eng, clock, v3=400, turn=50)
    assert fx.volumes() == [("set_volume", "mic", lv(450))]


def test_cambiar_salida_usa_el_giro_separado():
    eng, fx, clock = make({"click_turn": "output_device"})
    step(eng, clock)
    eng.on_button(0, protocol.BTN_HELD_START)
    fx.calls.clear()
    step(eng, clock, turn=40)          # +40 (>35): siguiente
    step(eng, clock, turn=60)          # +20: nada
    clock.t += 1.0
    step(eng, clock, turn=0)           # -40 desde el último cambio (40): anterior
    assert [c for c in fx.calls if c[0] == "cycle_output"] == [("cycle_output", 1), ("cycle_output", -1)]
    assert eng.smooth_vals[3] == 600.0


def test_firmware_viejo_sigue_como_v33():
    eng, fx, clock = make({"click_turn": "mic"}, levels={"mic": 0.25})
    eng.on_frame([512, 512, 512, 600, 512])            # sin hold_turns
    clock.t += 0.05
    eng.tick()
    eng.on_button(0, protocol.BTN_HELD_START)
    assert eng.hold_saved[0] == 600.0
