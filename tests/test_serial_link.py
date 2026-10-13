# -*- coding: utf-8 -*-
from types import SimpleNamespace

import serial

from knob.core import protocol, serial_link
from knob.core.serial_link import candidate_ports, handshake


class FakeSerial:
    def __init__(self, lines):
        self._lines = list(lines)
        self.writes = []

    def readline(self):
        return (self._lines.pop(0) + "\n").encode() if self._lines else b""

    def write(self, data):
        self.writes.append(data)


class StepClock:
    """Avanza 0.3 s en cada consulta, para simular el paso del tiempo sin esperar."""
    def __init__(self):
        self.t = 0.0

    def __call__(self):
        self.t += 0.3
        return self.t


def test_handshake_completo():
    ser = FakeSerial([protocol.CMD_AWAIT_INIT, protocol.LOG_INITIALIZED])
    assert handshake(ser, clock=StepClock()) is True
    assert ser.writes == [protocol.CMD_START]


def test_handshake_con_arduino_ya_transmitiendo():
    ser = FakeSerial(["512|512|512|512|512|0|0"])
    assert handshake(ser, clock=StepClock()) is True
    assert ser.writes == []


def test_handshake_silencio_manda_sondas_y_falla():
    ser = FakeSerial([])
    assert handshake(ser, timeout_s=5.0, clock=StepClock()) is False
    assert len(ser.writes) >= 3
    assert set(ser.writes) == {protocol.CMD_START}


def port(device, vid=None, description=""):
    return SimpleNamespace(device=device, vid=vid, description=description)


def test_candidate_ports_filtra_y_prioriza_el_ultimo():
    ports = [port("COM1", None, "Communications Port"),
             port("COM5", 0x2341, "Arduino Micro"),
             port("COM7", None, "USB-SERIAL CH340"),
             port("COM13", 0x1B4F, "SparkFun Pro Micro")]
    assert candidate_ports(ports, "COM13") == ["COM13", "COM5", "COM7"]
    assert candidate_ports(ports, None) == ["COM5", "COM7", "COM13"]
    assert candidate_ports([port("COM1")], None) == []


class FakeLinkSerial:
    """Puerto falso para _search: el primero lanza SerialException en readline."""
    def __init__(self, device, fail=False):
        self.device = device
        self.fail = fail
        self.closed = False
        self.is_open = True
        self.timeout = None

    def readline(self):
        if self.fail:
            raise serial.SerialException("boom")
        return (protocol.LOG_INITIALIZED + "\n").encode()

    def write(self, data):
        pass

    def close(self):
        self.closed = True
        self.is_open = False


def test_search_contiene_excepcion_de_handshake_y_sigue_con_el_siguiente(monkeypatch):
    created = {}

    def fake_serial(dev, baud, timeout=None, write_timeout=None):
        created[dev] = FakeLinkSerial(dev, fail=(dev == "COM5"))
        return created[dev]

    monkeypatch.setattr(serial_link.serial, "Serial", fake_serial)
    monkeypatch.setattr(serial_link.serial.tools.list_ports, "comports", lambda: [
        SimpleNamespace(device="COM5", vid=0x2341, description="Arduino Micro"),
        SimpleNamespace(device="COM7", vid=0x1B4F, description="SparkFun Pro Micro"),
    ])
    monkeypatch.setattr(serial_link.time, "sleep", lambda s: None)

    link = serial_link.SerialLink(lambda: None)
    link._running = True
    statuses = []
    link.status_changed.connect(lambda s, d: statuses.append((s, d)))

    result = link._search()

    assert result is created["COM7"]
    assert created["COM5"].closed is True
    assert ("connected", "COM7") in statuses


def test_modos_de_mantener_se_envian_al_conectar_y_al_cambiar():
    from knob.core.serial_link import SerialLink

    class WriteSerial:
        def __init__(self):
            self.writes = []

        def write(self, data):
            self.writes.append(data)

    link = SerialLink(lambda: None)
    link.set_hold_modes([True, False])            # sin conexión: solo se guarda
    ser = WriteSerial()
    link._send_hold_modes(ser)                    # lo que hace _run al conectar
    link._ser = ser
    link.set_hold_modes([False, True])            # con conexión: se envía en el momento
    assert ser.writes == [b"{KNOB_CMD_HOLD:10}\n", b"{KNOB_CMD_HOLD:01}\n"]


def test_take_latest_frame_trae_el_giro():
    from knob.core import protocol
    from knob.core.serial_link import SerialLink
    link = SerialLink(lambda: None)
    link._latest = protocol.parse_frame("1|2|3|4|5|3|0|50|0")
    frame = link.take_latest_frame()
    assert frame.hold_turns == [50, 0]
    assert link.take_latest() is None
