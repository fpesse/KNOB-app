# -*- coding: utf-8 -*-
from types import SimpleNamespace

from knob.core import protocol
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
