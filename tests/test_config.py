# -*- coding: utf-8 -*-
import json

import pytest

from knob.core.config import Config, ProfileError


def write(path, data):
    path.write_text(json.dumps(data), encoding="utf-8")


def test_archivo_inexistente_usa_defaults_sin_crear_archivo(tmp_path):
    p = tmp_path / "config.json"
    c = Config(str(p))
    assert c.profile_names() == ["Default"]
    assert c.active_profile == "Default"
    assert c.theme == "Nebula"
    assert c.channel_assignments() == {0: None, 1: None, 2: None, 3: None, 4: None}
    assert c.encoder_assignments()[0]["click"] == "none"
    assert c.minimize_to_tray is False
    assert not p.exists()


def test_migra_formato_plano_a_perfiles(tmp_path):
    p = tmp_path / "config.json"
    write(p, {"channel_assignments": {"0": "Discord", "3": "master"}, "theme": "Cobalt"})
    c = Config(str(p))
    assert c.profile_names() == ["Default"]
    assert c.channel_assignments()[0] == "Discord"
    assert c.channel_assignments()[3] == "master"
    assert c.theme == "Cobalt"


@pytest.mark.parametrize("viejo,nuevo", [("Arctic", "Onyx"), ("Steel", "Onyx"),
                                          ("Raro", "Nebula"), ("Magma", "Magma")])
def test_migracion_de_temas(tmp_path, viejo, nuevo):
    p = tmp_path / "config.json"
    write(p, {"theme": viejo})
    assert Config(str(p)).theme == nuevo


def test_config_corrupto_no_se_pisa(tmp_path):
    p = tmp_path / "config.json"
    p.write_text("{no es json", encoding="utf-8")
    c = Config(str(p))
    assert c.profile_names() == ["Default"]
    assert p.read_text(encoding="utf-8") == "{no es json"


def test_asignacion_persiste(tmp_path):
    p = tmp_path / "config.json"
    Config(str(p)).set_channel_assignment(1, "Spotify")
    c = Config(str(p))
    assert c.channel_assignments()[1] == "Spotify"
    assert json.loads(p.read_text(encoding="utf-8"))["profiles"]["Default"]["channel_assignments"]["1"] == "Spotify"


def test_encoders_se_normalizan(tmp_path):
    p = tmp_path / "config.json"
    write(p, {"profiles": {"Default": {"encoder_assignments": [{"click": "mute"}]}},
              "active_profile": "Default"})
    encs = Config(str(p)).encoder_assignments()
    assert len(encs) == 2
    assert encs[0]["click"] == "mute"
    assert encs[0]["hotkey_keys_click"] == ""
    assert encs[1]["click_turn"] == "none"


def test_crear_renombrar_y_borrar_perfiles(tmp_path):
    p = tmp_path / "config.json"
    c = Config(str(p))
    c.set_channel_assignment(0, "Opera")
    c.create_profile("Gaming", c.channel_assignments(), c.encoder_assignments())
    assert c.active_profile == "Gaming"
    assert c.channel_assignments()[0] == "Opera"
    c.set_channel_assignment(0, "Discord")
    c.switch_profile("Default")
    assert c.channel_assignments()[0] == "Opera"
    assert c.rename_profile("Gaming", "Juegos") == "Juegos"
    assert c.profile_names() == ["Default", "Juegos"]
    c.switch_profile("Juegos")
    assert c.channel_assignments()[0] == "Discord"
    assert c.delete_profile("Juegos") is True
    assert c.active_profile == "Default"
    assert Config(str(p)).profile_names() == ["Default"]


def test_renombrar_activo_actualiza_activo(tmp_path):
    c = Config(str(tmp_path / "config.json"))
    c.rename_profile("Default", "Casa")
    assert c.active_profile == "Casa"


def test_no_se_borra_el_ultimo_perfil(tmp_path):
    c = Config(str(tmp_path / "config.json"))
    assert c.delete_profile("Default") is False
    assert c.profile_names() == ["Default"]


def test_nombres_de_perfil_invalidos(tmp_path):
    c = Config(str(tmp_path / "config.json"))
    with pytest.raises(ProfileError):
        c.create_profile("   ", {}, [])
    with pytest.raises(ProfileError):
        c.create_profile("Default", {}, [])
    c.create_profile("Otro", {}, [])
    with pytest.raises(ProfileError):
        c.rename_profile("Otro", "Default")
    with pytest.raises(ProfileError):
        c.switch_profile("No existe")


def test_preferencias_y_exe_paths(tmp_path):
    p = tmp_path / "config.json"
    c = Config(str(p))
    c.theme = "Toxic"
    c.minimize_to_tray = True
    c.last_port = "COM13"
    c.remember_exe_path("Discord", r"C:\Discord\Discord.exe")
    c2 = Config(str(p))
    assert (c2.theme, c2.minimize_to_tray, c2.last_port) == ("Toxic", True, "COM13")
    assert c2.exe_path("discord") == r"C:\Discord\Discord.exe"
    with pytest.raises(ValueError):
        c2.theme = "Arctic"


def test_nombres_no_ascii_se_guardan_como_ascii(tmp_path):
    p = tmp_path / "config.json"
    c = Config(str(p))
    c.create_profile("Música", {}, [])
    c.remember_exe_path("Spotify", r"C:\Música\Spotify.exe")
    raw = p.read_bytes()
    raw.decode("ascii")   # v3.3 lee sin encoding: el archivo debe ser ASCII puro
    c2 = Config(str(p))
    assert c2.active_profile == "Música"
    assert "Música" in c2.profile_names()
    assert c2.exe_path("spotify") == r"C:\Música\Spotify.exe"


def test_nombre_demasiado_largo_se_rechaza(tmp_path):
    c = Config(str(tmp_path / "config.json"))
    with pytest.raises(ProfileError):
        c.create_profile("x" * 25, {}, [])
    assert c.create_profile("x" * 24, {}, []) == "x" * 24


def test_migrate_tolera_perfiles_y_asignaciones_no_dict(tmp_path):
    p = tmp_path / "config.json"
    write(p, {"profiles": {"A": "basura", "B": {"channel_assignments": [1, 2]}}, "active_profile": "A"})
    c = Config(str(p))
    assert c.channel_assignments() == {i: None for i in range(5)}
    c.switch_profile("B")
    assert c.channel_assignments() == {i: None for i in range(5)}
    p2 = tmp_path / "plano.json"
    write(p2, {"channel_assignments": ["x"]})
    assert Config(str(p2)).channel_assignments()[0] is None
