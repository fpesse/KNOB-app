# -*- coding: utf-8 -*-
"""Dobles de prueba para el motor."""


class FakeEffects:
    def __init__(self):
        self.calls = []

    def set_volume(self, target, level):
        self.calls.append(("set_volume", target, round(level, 4)))

    def media_key(self, action):
        self.calls.append(("media_key", action))

    def send_hotkey(self, combo):
        self.calls.append(("send_hotkey", combo))

    def step_master(self, delta):
        self.calls.append(("step_master", delta))

    def cycle_output(self, direction):
        self.calls.append(("cycle_output", direction))

    def volumes(self):
        return [c for c in self.calls if c[0] == "set_volume"]


class FakeClock:
    def __init__(self, t=100.0):
        self.t = t

    def __call__(self):
        return self.t
