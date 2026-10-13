# Rediseño de KNOB App en PySide6 (v4.0.0)

Fecha: 2026-10-13 · Rama: `rediseno-qt` · Base: `a7acd7e` (v3.3.0, customtkinter)

## Objetivo

Reemplazar la interfaz customtkinter de `app.py` por una interfaz PySide6 con estética profesional (estilo "Glow"), **sin cambiar el comportamiento**: serial, audio, encoders, perfiles, bandeja, inicio con Windows y actualizador funcionan igual que en v3.3.0.

Criterio de éxito: con el KNOB físico conectado, todas las pruebas de la lista manual (sección 9) pasan, y un usuario de v3.3.0 actualiza y conserva sus perfiles, asignaciones y tema sin hacer nada.

## Fuera de alcance

- App de Mac (`D:\Proyecto\app-mac`).
- Firmware y protocolo serial (handshake `{KNOB_CMD_AWAIT_INIT}` → `{KNOB_CMD_START}` → `{KNOB_LOG_INITIALIZED}`, frame `v0|v1|v2|v3|v4|btnL|btnR`, 115200 baudios).
- Cambios de comportamiento de audio. Las funciones se mueven, no se rediseñan.
- Faders controlables con el mouse: siguen siendo de solo lectura.

## 1. Dirección visual

**Estilo Glow** (como SteelSeries GG): fondo casi negro con un halo radial de color arriba a la derecha, superficies translúcidas (`rgba(255,255,255,.04)` con borde `.07`), degradés de dos colores en lo activo y brillo sutil (glow) en el fader y en la perilla.

**Encoders con estilo "perilla física"** (tomado de la dirección Consola): un círculo oscuro con un borde fino, 5 marcas de escala alrededor y un puntero luminoso del color del tema que rota 270° (de 7 a 5 en punto).

**Íconos reales de las apps** extraídos del `.exe` (como hoy, con `SHGetFileInfoW`), mostrados en 28 px dentro de cada tarjeta de canal. Son centrales en el diseño. El respaldo es una inicial sobre un círculo cuando no hay ícono.

**Tipografía:** el logo "KNOB" usa la fuente Amazing Views (`fonts/AmazingViews.ttf`, cargada con `QFontDatabase`) rellena con el degradé del tema. El resto usa Segoe UI Variable, con Segoe UI como respaldo. Los números del fader usan cifras tabulares.

**Barra de título nativa** de Windows, teñida con el color de fondo del tema mediante `DwmSetWindowAttribute`:
- `DWMWA_USE_IMMERSIVE_DARK_MODE` = 20.
- `DWMWA_CAPTION_COLOR` = 35.
- `DWMWA_BORDER_COLOR` = 34.

En Windows 10, donde el color de la barra no se aplica, queda el modo oscuro.

### Temas

Cada tema define `grad_a`, `grad_b` (degradé), `halo` (color del halo), `bg` (fondo base), `knob_face`, `knob_ring`, `pointer` y `text_brand_a/b`.

| Tema | Degradé | Notas |
|---|---|---|
| Nebula | `#7C3AED` → `#EC4899` | por defecto |
| Cobalt | `#2563EB` → `#22D3EE` | |
| Magma | `#DC2626` → `#F97316` | |
| Onyx | `#737373` → `#F5F5F5` | monocromo |
| Toxic | `#059669` → `#A3E635` | nuevo |
| Sunset | `#DB2777` → `#FBBF24` | nuevo |

Arctic y Steel se eliminan. Si `config.json` tiene `"theme": "Arctic"` o `"Steel"`, se migra a `"Onyx"`; cualquier otro nombre desconocido pasa a `"Nebula"`.

Los colores de estado no dependen del tema:
- Verde `#4ADE80` (conectado y medidor).
- Ámbar `#FBBF24` (buscando).
- Rojo `#F87171` (error y MUTE).

## 2. Layout: una sola pantalla

La ventana es de **520 × 620 px**, de tamaño fijo y con la escala DPI de Qt (se elimina el truco `SetProcessDpiAwareness(0)`, porque Qt escala bien con DPI alto).

```
┌───────────────────────────────────────────────┐
│ KNOB  ● Conectado        [Perfil: Gaming ▾] ⚙ │  barra superior
│ ┌─────────┐ ┌─────────┐ ┌─────────┐           │
│ │ [icono] │ │ [icono] │ │ [icono] │           │  3 tarjetas de canal
│ │  Opera  │ │ Discord │ │ Spotify │           │  (click = asignar)
│ │  ┃ ▌    │ │  ┃ ▌    │ │  ┃ ▌    │           │  fader + medidor VU
│ │  48%    │ │  84%    │ │  85%    │           │
│ └─────────┘ └─────────┘ └─────────┘           │
│ ┌──────────────────┐ ┌──────────────────┐     │  2 tarjetas de encoder
│ │ (perilla) Enc izq│ │ (perilla) Enc der│     │  click = asignar canal 3/4
│ │  Vol. general  ⚙ │ │  App en foco   ⚙ │     │  ⚙ = acciones del botón
│ └──────────────────┘ └──────────────────┘     │
│ 🔊 Realtek Digital Output             ‹  ›    │  salida de audio
└───────────────────────────────────────────────┘
```

- **Barra superior:** logo, pastilla de conexión, menú de perfil (cambio rápido) y ⚙ para abrir Ajustes.
- **Tarjeta de canal (0–2):** ícono, nombre, fader vertical (pista de 8 px, relleno con degradé y glow, tapa blanca de 22×10 px), medidor VU fino de 3 px al lado (verde, con retención de pico) y el valor en %. Un click en cualquier parte de la tarjeta abre "Asignar canal".
- **Tarjeta de encoder (3–4):** perilla, la etiqueta "Encoder izq./der." y el nombre del destino. Un click en la tarjeta abre "Asignar canal" para el canal 3/4. El ⚙ chico abre "Configurar encoder".
- **Barra de salida de audio:** nombre del dispositivo predeterminado actual y flechas ‹ › para cambiarlo con el mouse (función nueva que reutiliza `_cycle_output_device`).

## 3. Superposiciones (dentro de la misma ventana)

Todas se abren sobre la ventana principal oscureciendo el fondo (`rgba(5,3,10,.55)`), con fade y desplazamiento de 150–200 ms. Se cierran con Esc o con un click afuera. No se abren ventanas sueltas.

**Asignar canal:**
- Título "Canal N · ¿Qué controla?".
- Sección *Sistema*: Volumen general, Micrófono y App en foco.
- Sección *Apps con audio ahora*: la lista de `get_audio_apps()` con íconos reales.
- La opción actual aparece resaltada con el degradé.
- Al final, un enlace "Quitar asignación".
- Un click asigna, guarda y cierra.

**Configurar encoder:**
- Muestra solo de lectura "Girar controla: X".
- Desplegables para *Click*, *Doble click* (con las opciones de `ENCODER_CLICK_ACTIONS`) y *Mantener y girar* (con las de `ENCODER_TURN_CHANNELS`).
- Si se elige "Atajo de teclado", aparece al lado un botón para grabar la combinación. Al hacer click en él se capturan los modificadores y la tecla con `keyPressEvent`, y se guardan en el mismo formato de hoy (`ctrl+shift+m`).
- Botones Cancelar / Guardar.

**Panel de ajustes** (se desliza desde la derecha y ocupa ~72% del ancho):
- *Perfiles*: lista con el perfil activo marcado, y botones + Nuevo, **Renombrar** (nuevo) y Borrar. Borrar se desactiva cuando queda un solo perfil y pide una confirmación en línea.
- *Tema*: 6 círculos con el degradé de cada tema; el activo lleva un anillo blanco. El cambio es en vivo.
- *Sistema*: interruptores "Iniciar con Windows" y "Al cerrar, ir a la bandeja".
- *Acerca de*: la versión. Si hay actualización, aparece el botón "↑ vX disponible", que al tocarlo muestra el progreso de descarga ("Descargando 42%") y luego lanza el instalador, igual que hoy.

**Nuevo perfil / Renombrar:** una tarjeta pequeña con un campo de texto, Guardar y Cancelar. Enter confirma. Los nombres vacíos o duplicados se rechazan con un mensaje en línea.

## 4. Estados

**Pastilla de conexión:**
- Verde **Conectado**.
- Ámbar **Buscando KNOB…**, con un punto que late (animación de opacidad).
- Rojo **COMx ocupado: cerrá el Serial Monitor**.

**Estados de la tarjeta de canal:**
- **Muteado** (valor ≤ `MUTE_THRESHOLD`, o encoder muteado con el botón): ícono al 40% de opacidad, fader gris sin glow y etiqueta roja "MUTE" en lugar del %.
- **Sin asignar**: borde punteado, un "+" en lugar del ícono, el texto "Sin asignar" y el fader apagado.
- **App en foco**: debajo del nombre muestra "→ Chrome" (el `_last_focus_exe` actual), actualizado cada ~300 ms.

**Cambio de salida:** la barra inferior muestra el nuevo nombre y brilla con el color del tema durante 1,5 s.

**Lista de apps vacía:** "No hay apps reproduciendo audio. Abrí la app y volvé a intentar".

## 5. Arquitectura del código

```
main.py                 entrada: instancia única, QApplication, arranque
knob/
  core/
    config.py           load/save de %APPDATA%\KNOB\config.json, perfiles, migración de temas
    protocol.py         parseo de frames y eventos de botón (funciones puras)
    serial_link.py      QThread: búsqueda de puertos, handshake, lectura → señales
    engine.py           suavizado EMA, deadband, mute, prioridad foco, hold/click_turn (puro, sin Qt)
    audio.py            pycaw: caché de sesiones, volumen, medidores, mute, dispositivos de salida
    icons.py            extracción de íconos de .exe → QPixmap, caché, respaldo con inicial
    system.py           inicio con Windows (registro), teclas multimedia/atajos, rastreo de app en foco, DWM
    updater.py          consulta a GitHub Releases, descarga, lanzamiento del instalador
  ui/
    theme.py            definición de temas + aplicación (paleta, QSS mínimo)
    main_window.py      ventana, barra superior, grilla de tarjetas, barra de salida
    widgets/
      fader.py          QWidget con paintEvent (QPainter con antialiasing)
      knob_dial.py      perilla (QPainter)
      vu_meter.py       medidor con retención de pico
      channel_card.py   tarjeta de canal (ícono, nombre, fader, VU, %, estados)
      encoder_card.py   tarjeta de encoder
      status_pill.py
    overlays/
      overlay_host.py   capa que oscurece y anima la entrada/salida
      assign_channel.py
      encoder_config.py
      settings_panel.py
      profile_name.py
    tray.py             QSystemTrayIcon (Mostrar KNOB / Salir)
tests/
  test_protocol.py  test_engine.py  test_config.py
```

**Reglas:**
- `core/` no importa nada de `ui/`. `engine.py` y `protocol.py` no importan Qt ni pycaw, así se pueden testear con pytest sin hardware.
- **Hilos:**
  - `serial_link` corre en un `QThread` y emite `frame_received(list[int])`, `button_event(enc_idx, code)` y `status_changed(state, detail)`.
  - El caché de sesiones de audio y el rastreo de foco siguen en hilos daemon, como hoy. La UI los lee con un `QTimer`.
  - **Ningún hilo toca widgets.** Todo pasa por señales o por timers del hilo principal.
- **Bucles del hilo principal** (equivalentes a los `root.after` de hoy):
  - `QTimer` de 16 ms: el engine avanza el suavizado → se actualizan faders y perillas → se aplica el volumen limitado a 40 ms (`VOL_RATE_MS`).
  - `QTimer` de 33 ms: medidores VU.
- **Lógica migrada sin cambios** desde `app.py` (constantes incluidas):
  - `SMOOTH_ALPHA`, `DEADBAND`, `MUTE_THRESHOLD`, `raw_to_level` (ley cuadrática), `VOL_RATE_MS`.
  - Prioridad "el último que se movió gana" (`_channel_priority`, `_focus_took_at`, umbral de 40).
  - Mute del encoder con guardado y restauración, hold + click_turn con congelamiento de la perilla, y cambio de salida al girar con umbral de 35 y antirrebote de 800 ms.
  - `_send_hotkey`, teclas multimedia, `_set_default_output_device` (IPolicyConfig), lista negra de procesos del sistema y `_linear_to_vu` (piso de -30 dB).
- **Instancia única:** se mantiene el mutex `Global\KNOBAppMutex`. Si ya hay una instancia, se avisa con un mensaje (como hoy).
- **Bandeja:** `QSystemTrayIcon` reemplaza a pystray. Doble click muestra la ventana. `--minimized` arranca oculto en la bandeja. Cerrar con la bandeja activada oculta la ventana; si no, sale cerrando el serial.
- **Log:** `knob_debug.log` sigue en la carpeta del exe con el mismo formato `_log`.

### Formato de config.json

Se mantiene sin cambios: `theme`, `minimize_to_tray`, `last_port`, `profiles{nombre: {channel_assignments, encoder_assignments}}`, `active_profile` y `exe_paths`. La única migración es la de los temas eliminados. Renombrar un perfil mueve la clave y actualiza `active_profile` si corresponde.

## 6. Manejo de errores

- Las fallas de pycaw o COM en llamadas de volumen invalidan el caché y se reintentan en el próximo ciclo, como hoy. Nunca muestran un diálogo.
- Si el puerto está ocupado, la pastilla pasa a rojo y la búsqueda sigue cada 2 s.
- Si el config.json está corrupto o no existe, se arranca con los valores por defecto sin pisar el archivo hasta el primer cambio, y el error se registra en el log.
- Si la extracción de un ícono falla, se usa la inicial y se reintenta cada 10 s (comportamiento actual).
- Si el actualizador no tiene red o falla la descarga, falla en silencio. Si falla la descarga, el botón muestra "Error al descargar".

## 7. Empaquetado y entrega

- Nuevas dependencias: `PySide6-Essentials` (solo QtCore, QtGui y QtWidgets) y `pytest` (solo para desarrollo). Se quitan `customtkinter` y `pystray` del build.
- `KNOB.spec`:
  - Entrada `main.py`.
  - Modo **onedir** (arranque instantáneo).
  - Datas: `icon.png`, `fonts/` y `version.txt`.
  - Excluir los módulos Qt que no se usan (QtWebEngine, Qml, Quick, Multimedia, etc.) para reducir el tamaño.
- `KNOB_App_Installer.iss`:
  - Copiar `dist\KNOB\*` de forma recursiva.
  - Mantener AppId y ruta para que actualice sobre la v3.3.0.
  - El nuevo `KNOB.exe` (onedir, con `_internal\` al lado) sobrescribe al de un solo archivo en la misma ruta, así los accesos directos y la clave de inicio con Windows siguen funcionando.
- `release.py` no cambia: el instalador sigue llamándose `KNOB_App_Installer.exe`, así el actualizador de la v3.3.0 lo encuentra.
- Versión: **v4.0.0**.
- `app.py` queda en la rama como referencia hasta que se valide la v4. Se elimina en el último paso del plan.

## 8. Tests automáticos (pytest)

- `test_protocol`:
  - Frames de 5 y 7 valores.
  - Basura y líneas vacías.
  - Detección de cambios en los eventos de botón (solo dispara cuando cambia el valor).
- `test_engine`:
  - Deadband (cambios menores a 10 se ignoran).
  - Convergencia del EMA.
  - Mute al umbral 15.
  - `raw_to_level`.
  - Prioridad foco/directo: el foco toma el control al moverse más de 15, y el canal directo lo recupera al moverse más de 40.
  - Mute/unmute del encoder con restauración del valor.
  - Hold + click_turn: congela la perilla y restaura al soltar.
  - Cambio de salida por giro (umbral de 35 y antirrebote).
- `test_config`:
  - Migración del formato plano a perfiles.
  - Migración de Arctic/Steel.
  - Crear, renombrar y borrar perfiles (no se puede borrar el último).
  - Config corrupto.

## 9. Lista de pruebas manuales (con el KNOB)

1. Arrancar con el KNOB conectado: la pastilla pasa de "Buscando" a "Conectado" en menos de 3 s.
2. Mover cada fader: la UI sigue suavemente y cambia el volumen de la app correcta (comparar con el mezclador de Windows).
3. Medidores VU: se mueven con el audio y escalan con la posición del fader.
4. Encoders: giro, click, doble click, mantener y girar (incluido el cambio de salida de audio).
5. Atajo de teclado grabado y ejecutado.
6. Desenchufar y reenchufar: reconecta solo.
7. Abrir el Serial Monitor de Arduino: aparece la pastilla roja de puerto ocupado.
8. Cambiar de perfil y de tema; crear, renombrar y borrar un perfil; reiniciar la app y verificar que todo persiste.
9. "App en foco" muestra la app activa y respeta la prioridad del último que se movió.
10. Bandeja, `--minimized`, iniciar con Windows y segunda instancia bloqueada.
11. Actualizar sobre una instalación v3.3.0 existente: se conservan perfiles y tema.
12. Escala de Windows al 100%, 125% y 150%: nada se corta.
