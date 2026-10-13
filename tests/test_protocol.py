# -*- coding: utf-8 -*-
from knob.core import protocol
from knob.core.protocol import ButtonEdgeDetector, parse_frame


def test_frame_de_siete_valores():
    f = parse_frame("10|20|30|40|50|1|0\r\n")
    assert f.values == [10, 20, 30, 40, 50]
    assert f.buttons == [1, 0]


def test_frame_legado_de_cinco_valores():
    f = parse_frame("1|2|3|4|5")
    assert f.values == [1, 2, 3, 4, 5]
    assert f.buttons == [None, None]


def test_frames_invalidos():
    assert parse_frame("") is None
    assert parse_frame("1|2|3|4") is None
    assert parse_frame("{KNOB_LOG_INITIALIZED}") is None
    assert parse_frame("a|2|3|4|5") is None


def test_boton_no_numerico_queda_none():
    assert parse_frame("1|2|3|4|5|x|2").buttons == [None, 2]


def test_is_stream_line():
    assert protocol.is_stream_line("1|2|3|4|5")
    assert not protocol.is_stream_line(protocol.CMD_AWAIT_INIT)


def test_detector_dispara_solo_en_flanco():
    d = ButtonEdgeDetector()
    assert d.feed([0, 0]) == []
    assert d.feed([1, 0]) == [(0, protocol.BTN_CLICK)]
    assert d.feed([1, 0]) == []
    assert d.feed([0, 0]) == []
    assert d.feed([3, 3]) == [(0, 3), (1, 3)]
    assert d.feed([3, 3]) == []          # el firmware repite HELD mientras se mantiene
    assert d.feed([4, 0]) == [(0, 4)]
    assert d.feed([None, None]) == []


def test_frame_con_giro_mantenido():
    f = parse_frame("1|2|3|4|5|3|0|-120|0")
    assert f.buttons == [3, 0]
    assert f.hold_turns == [-120, 0]


def test_frame_sin_giro_mantenido_es_firmware_viejo():
    assert parse_frame("1|2|3|4|5|0|0").hold_turns is None
    assert parse_frame("1|2|3|4|5|0|0|x|1").hold_turns is None


def test_hold_command():
    assert protocol.hold_command([True, False]) == b"{KNOB_CMD_HOLD:10}\n"
    assert protocol.hold_command([False, True]) == b"{KNOB_CMD_HOLD:01}\n"
