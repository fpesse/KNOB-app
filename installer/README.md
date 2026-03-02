# KNOB Audio Mixer – Instalador

Este directorio contiene los archivos necesarios para generar un instalador de Windows usando **Inno Setup** y firmar tanto el ejecutable como el instalador final.

## Requisitos previos

1. **PyInstaller**: asegúrate de haber generado la carpeta `dist/` ejecutando:
   ```powershell
   pyinstaller audio_mixer_app.spec
   ```
2. **Inno Setup** (https://jrsoftware.org/isinfo.php). Instala el compilador y agrega `ISCC.exe` al `PATH` o usa la ruta completa.
3. (Opcional pero recomendado) **Certificado de firma de código** y el SDK de Windows para usar `signtool.exe`.

## Generar el instalador

1. Desde la raíz del proyecto:
   ```powershell
   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\KNOB_installer.iss
   ```
   El instalador resultante quedará en `installer/output/KNOB_Setup.exe`.
2. El script copia **todo** el contenido de `dist/` dentro de `{app}`. Asegúrate de que `dist/` contenga el ejecutable y los recursos necesarios.

## Firmar los binarios (opcional, recomendado)

1. Firma el ejecutable antes de empaquetarlo (ajusta la ruta del certificado y contraseña):
   ```powershell
   signtool sign /fd SHA256 /td SHA256 /tr http://timestamp.digicert.com /a dist\audio_mixer_app.exe
   ```
2. Tras generar `KNOB_Setup.exe`, fírmalo también:
   ```powershell
   signtool sign /fd SHA256 /td SHA256 /tr http://timestamp.digicert.com /a installer\output\KNOB_Setup.exe
   ```

## Personalización rápida

- Actualiza metadatos (nombre, versión, URL) editando `installer/KNOB_installer.iss`.
- Para añadir archivos adicionales (p.ej. manual PDF), colócalos en `dist/` o añade líneas en la sección `[Files]`.
- Si cambias el icono principal, ajusta `SetupIconFile` y asegúrate de incluir el nuevo `.ico`.

Con estos pasos tendrás un instalador tradicional que crea accesos directos y puede ser firmado para evitar advertencias de SmartScreen.
