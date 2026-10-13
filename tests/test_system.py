# -*- coding: utf-8 -*-
from knob.core import system, updater


def test_parse_hotkey():
    assert system.parse_hotkey("ctrl+shift+m") == (["ctrl", "shift"], ["m"])
    assert system.parse_hotkey(" Alt + F4 ") == (["alt"], ["f4"])
    assert system.parse_hotkey("") == ([], [])


def test_strip_exe():
    assert system.strip_exe("Discord.exe") == "Discord"
    assert system.strip_exe("spotify.EXE") == "spotify"
    assert system.strip_exe("opera") == "opera"
    assert system.strip_exe("") == ""


def test_hex_to_colorref():
    assert system.hex_to_colorref("#0C0A14") == 0x00140A0C
    assert system.hex_to_colorref("FF0000") == 0x000000FF


def test_parse_version():
    assert updater.parse_version("v4.0.0") == (4, 0, 0)
    assert updater.parse_version("v3.10.0") > updater.parse_version("v3.9.9")
    assert updater.parse_version("basura") == (0,)


def test_read_version_lee_version_txt():
    assert updater.read_version().startswith("v")
