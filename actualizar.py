"""
actualizar.py — Script todo-en-uno para actualizar el tablero BRE-B
====================================================================
Uso:
    python actualizar.py

Hace las tres cosas en secuencia:
  1. Genera el nuevo tablero HTML desde el Excel
  2. Hace git add + commit con fecha automática
  3. Hace git push → GitHub Pages se actualiza en ~30 segundos
"""

import subprocess
import sys
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).parent

def run(cmd, check=True):
    print(f"  $ {cmd}")
    result = subprocess.run(cmd, shell=True, cwd=BASE,
                            capture_output=True, text=True)
    if result.stdout.strip():
        print(f"    {result.stdout.strip()}")
    if result.returncode != 0 and check:
        print(f"  ❌  Error: {result.stderr.strip()}")
        sys.exit(1)
    return result

print("=" * 55)
print("  Tablero BRE-B · Febor — Actualización automática")
print("=" * 55)
print()

# ── Paso 1: Generar HTML ──────────────────────────────────────────────────────
print("1️⃣   Generando tablero desde Excel…")
result = subprocess.run([sys.executable, "generar.py"], cwd=BASE,
                        capture_output=False)
if result.returncode != 0:
    print("❌  Error al generar el tablero. Revisa el Excel.")
    sys.exit(1)
print()

# ── Paso 2: Git add + commit ──────────────────────────────────────────────────
print("2️⃣   Registrando cambios en Git…")
fecha_msg = datetime.now().strftime("%Y-%m-%d %H:%M")
run("git add docs/index.html datos/datos.xlsx")
run(f'git commit -m "Actualizar datos BRE-B · {fecha_msg}"')
print()

# ── Paso 3: Git push ─────────────────────────────────────────────────────────
print("3️⃣   Publicando en GitHub Pages…")
run("git push")
print()

print("=" * 55)
print("  ✅  ¡Listo! El tablero se actualizará en ~30 segundos.")
print()

# Leer URL del remote
remote = subprocess.run("git remote get-url origin", shell=True, cwd=BASE,
                        capture_output=True, text=True).stdout.strip()
if "github.com" in remote:
    # Transformar https://github.com/user/repo.git  →  https://user.github.io/repo/
    import re
    m = re.search(r"github\.com[:/]([^/]+)/([^/.]+)", remote)
    if m:
        user, repo = m.group(1), m.group(2)
        print(f"  🔗  URL pública: https://{user}.github.io/{repo}/")
print("=" * 55)
