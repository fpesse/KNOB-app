# -*- coding: utf-8 -*-
"""Vista previa de la interfaz con datos falsos animados (sin hardware ni audio).

Uso:
  python -m knob.ui.preview [Tema] [--screenshot ruta.png]
"""
import math
import sys

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from knob.ui.main_window import MainWindow
from knob.ui.theme import build_qss, get_theme, ui_font


def build_demo_window(theme):
    w = MainWindow(theme)
    w.apply_theme(theme)
    w.set_profiles(["Default", "Gaming", "Trabajo"], "Gaming")
    w.status.set_state("connected", "Conectado")
    w.output_bar.set_name("Realtek Digital Output (Realtek(R) Audio)")
    w.card_for(0).set_assignment("Opera", None, None)
    w.card_for(1).set_assignment("Discord", None, None)
    w.card_for(2).set_assignment("Spotify", None, None)
    w.card_for(3).set_assignment("Volumen general", None, "master")
    w.card_for(4).set_assignment("App en foco", None, "focus")
    w.card_for(4).set_focus_target("Chrome")
    for ch, raw in enumerate((490, 860, 870, 715, 330)):
        w.card_for(ch).set_value(raw)
    for i, lvl in enumerate((0.45, 0.72, 0.6)):
        w.channel_cards[i].set_level(lvl, lvl + 0.1)
    return w


def main(argv):
    args = [a for a in argv[1:] if not a.startswith("--")]
    shot = argv[argv.index("--screenshot") + 1] if "--screenshot" in argv else None
    if shot in args:
        args.remove(shot)
    app = QApplication(argv)
    app.setStyle("Fusion")
    app.setFont(ui_font(13))
    theme = get_theme(args[0] if args else "Nebula")
    app.setStyleSheet(build_qss(theme))
    w = build_demo_window(theme)
    w.close_requested.connect(app.quit)
    w.show()

    t = {"v": 0.0}

    def animate():
        t["v"] += 0.05
        for i in range(3):
            raw = 512 + 400 * math.sin(t["v"] + i)
            w.card_for(i).set_value(raw)
            lvl = 0.5 + 0.4 * math.sin(t["v"] * 3 + i * 2)
            w.channel_cards[i].set_level(lvl, min(1.0, lvl + 0.08))
        w.card_for(3).set_value(512 + 400 * math.sin(t["v"] * 0.7))

    if shot:
        def capture():
            w.grab().save(shot)
            print(f"Captura guardada en {shot}")
            app.quit()
        QTimer.singleShot(600, capture)
    else:
        timer = QTimer(w)
        timer.timeout.connect(animate)
        timer.start(33)
    return app.exec()


if __name__ == "__main__":
    sys.exit(main(sys.argv))
