# -*- coding: utf-8 -*-
"""Protocolo serial KNOB <-> app. Es el contrato con el firmware: no cambiar sin tocar ambos lados."""
from dataclasses import dataclass

BAUD_RATE = 115200
CMD_AWAIT_INIT = "{KNOB_CMD_AWAIT_INIT}"
CMD_START = b"{KNOB_CMD_START}\n"
LOG_INITIALIZED = "{KNOB_LOG_INITIALIZED}"
NUM_CHANNELS = 5

BTN_NONE = 0
BTN_CLICK = 1
BTN_DBLCLICK = 2
BTN_HELD_START = 3
BTN_HELD_END = 4


@dataclass
class Frame:
    values: list    # 5 valores analógicos 0-1023
    buttons: list   # [btnL, btnR]; None si el frame no los trae o no son numéricos


def parse_frame(line: str):
    """Parsea 'v0|v1|v2|v3|v4|btnL|btnR'. Acepta >= 5 valores (frames legados)."""
    partes = line.strip().split("|")
    if len(partes) < NUM_CHANNELS:
        return None
    try:
        values = [int(x) for x in partes[:NUM_CHANNELS]]
    except ValueError:
        return None
    buttons = []
    for col in (5, 6):
        if len(partes) > col:
            try:
                buttons.append(int(partes[col]))
            except ValueError:
                buttons.append(None)
        else:
            buttons.append(None)
    return Frame(values, buttons)


def is_stream_line(line: str) -> bool:
    return line.count("|") >= 4


class ButtonEdgeDetector:
    """El firmware repite el código mientras dura el evento; solo se dispara cuando cambia."""

    def __init__(self):
        self._prev = [BTN_NONE, BTN_NONE]

    def feed(self, buttons):
        events = []
        for idx, ev in enumerate(buttons[:2]):
            if ev is None:
                continue
            if ev != self._prev[idx]:
                self._prev[idx] = ev
                if ev > 0:
                    events.append((idx, ev))
        return events
