"""
Uso: python release.py v3.2.0 ghp_TOKEN
Hace todo: actualiza version.txt, compila, crea release en GitHub y sube el instalador.
"""
import sys, os, re, subprocess, json, urllib.request

if len(sys.argv) < 3:
    print("Uso: python release.py v3.2.0 ghp_TOKEN")
    sys.exit(1)

VERSION = sys.argv[1].strip()
TOKEN   = sys.argv[2].strip()

if not VERSION.startswith("v"):
    VERSION = "v" + VERSION

ROOT = os.path.dirname(os.path.abspath(__file__))

# 1. Actualizar version.txt y AppVersion del instalador
print(f"[1/4] Actualizando version.txt e instalador -> {VERSION}")
with open(os.path.join(ROOT, "version.txt"), "w", encoding="utf-8") as f:
    f.write(VERSION + "\n")
iss_path = os.path.join(ROOT, "KNOB_App_Installer.iss")
with open(iss_path, encoding="utf-8") as f:
    iss, n = re.subn(r"(?m)^AppVersion=.*$", f"AppVersion={VERSION.lstrip('v')}", f.read())
if n == 0:
    print("ERROR: no se encontro AppVersion en KNOB_App_Installer.iss"); sys.exit(1)
with open(iss_path, "w", encoding="utf-8") as f:
    f.write(iss)

# 2. Compilar exe
print("[2/4] Compilando exe...")
r = subprocess.run(
    [os.path.join(ROOT, "venv/Scripts/pyinstaller.exe"), "--clean", "-y", "KNOB.spec"],
    cwd=ROOT
)
if r.returncode != 0:
    print("ERROR: PyInstaller fallo"); sys.exit(1)

# 3. Compilar instalador
print("[3/4] Compilando instalador...")
r = subprocess.run(
    [r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe", "KNOB_App_Installer.iss"],
    cwd=ROOT
)
if r.returncode != 0:
    print("ERROR: Inno Setup fallo"); sys.exit(1)

# 4. Crear release y subir instalador
print("[4/4] Subiendo release a GitHub...")
headers_json = {"Authorization": f"token {TOKEN}", "Content-Type": "application/json"}

print(f"\nDescribe los cambios de {VERSION} (Enter en linea vacia para terminar):")
lines = []
while True:
    line = input("  > ")
    if line == "":
        break
    lines.append(line)
changelog = "\n".join(f"- {l}" for l in lines) if lines else "_Sin descripcion_"

body = (
    f"## Novedades {VERSION}\n\n"
    f"{changelog}\n\n"
    "---\n\n"
    "## Instalacion\n\n"
    "1. Descarga KNOB_App_Installer.exe\n"
    "2. Ejecutalo y segui el instalador\n"
    "3. Conecta tu KNOB por USB\n"
    "4. La app detecta el dispositivo automaticamente"
)
data = json.dumps({"tag_name": VERSION, "name": f"KNOB App {VERSION}",
                   "body": body, "draft": False, "prerelease": False}).encode()
req = urllib.request.Request(
    "https://api.github.com/repos/fpesse/KNOB-app/releases",
    data=data, headers=headers_json
)
with urllib.request.urlopen(req) as resp:
    release = json.loads(resp.read())
release_id = release["id"]
print(f"    Release creado: ID {release_id}")

installer_path = os.path.join(ROOT, "Output", "KNOB_App_Installer.exe")
with open(installer_path, "rb") as f:
    payload = f.read()

upload_url = f"https://uploads.github.com/repos/fpesse/KNOB-app/releases/{release_id}/assets?name=KNOB_App_Installer.exe"
req2 = urllib.request.Request(
    upload_url, data=payload,
    headers={"Authorization": f"token {TOKEN}", "Content-Type": "application/octet-stream"}
)
with urllib.request.urlopen(req2, timeout=120) as resp:
    asset = json.loads(resp.read())
print(f"    Instalador subido: {asset['browser_download_url']}")
print(f"\nListo! https://github.com/fpesse/KNOB-app/releases/tag/{VERSION}")
