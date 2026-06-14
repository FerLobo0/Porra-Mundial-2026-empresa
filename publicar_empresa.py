# -*- coding: utf-8 -*-
"""
publicar.py
Regenera la web desde el Excel y la publica en GitHub Pages.

Uso diario:
  1. Actualiza el Excel maestro
  2. Ejecuta este script (desde Spyder o desde el terminal)
  3. En ~30 segundos la web estará actualizada en internet
"""

import subprocess
import sys
import os
import glob
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIGURACIÓN — edita esto si es necesario
# ─────────────────────────────────────────────

MENSAJE_COMMIT = "Actualización de resultados"
RAMA           = "main"

# ─────────────────────────────────────────────


def run(comando, descripcion):
    """Ejecuta un comando de shell, aborta si falla."""
    print(f"  → {descripcion}...")
    resultado = subprocess.run(
        comando,
        shell=True,
        capture_output=True,
        text=True,
        encoding="utf-8"
    )
    if resultado.returncode != 0:
        print(f"\n❌ ERROR en: {descripcion}")
        print(resultado.stderr or resultado.stdout)
        sys.exit(1)
    return resultado.stdout.strip()


def hay_cambios():
    """True si hay archivos HTML modificados respecto al último commit."""
    resultado = subprocess.run(
        "git status --porcelain",
        shell=True,
        capture_output=True,
        text=True,
        encoding="utf-8"
    )
    # Solo nos interesan cambios en HTML
    lineas = [l for l in resultado.stdout.splitlines() if l.strip().endswith(".html")]
    return bool(lineas)


def main():
    print("=" * 52)
    print("   PUBLICADOR — Porra Mundial 2026")
    print("=" * 52)

    # Situarse en el directorio del script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    print(f"\nCarpeta: {script_dir}\n")

    # ── PASO 1: regenerar HTML desde el Excel ──────────
    print("[ 1/3 ] Generando HTML desde el Excel...")
    resultado = subprocess.run(
        f'"{sys.executable}" generar_web.py',
        shell=True,
        capture_output=True,
        text=True,
        encoding="utf-8"
    )
    if resultado.returncode != 0:
        print("❌ Error al ejecutar generar_web.py:")
        print(resultado.stderr)
        sys.exit(1)
    # Mostrar la salida de generar_web.py indentada
    for linea in resultado.stdout.splitlines():
        print(f"   {linea}")

    # ── PASO 2: comprobar si hay cambios reales ────────
    print("\n[ 2/3 ] Comprobando cambios...")
    if not hay_cambios():
        print("   ℹ️  Los HTML son idénticos a la versión publicada.")
        print("      No se sube nada. La web ya está al día.")
        return

    # ── PASO 3: subir a GitHub ─────────────────────────
    print("\n[ 3/3 ] Subiendo a GitHub Pages...")

    # Añadir todos los .html que existan (incluye participantes nuevos automáticamente)
    archivos_html = glob.glob("*.html")
    if not archivos_html:
        print("❌ No se encontraron archivos .html para subir.")
        sys.exit(1)

    archivos_str = " ".join(f'"{f}"' for f in archivos_html)
    run(f"git add {archivos_str}", f"Preparando {len(archivos_html)} archivos HTML")

    fecha_hora = datetime.now().strftime("%d/%m/%Y %H:%M")
    mensaje = f"{MENSAJE_COMMIT} — {fecha_hora}"
    run(f'git commit -m "{mensaje}"', "Creando commit")
    run(f"git push origin {RAMA}", "Subiendo a GitHub")

    print("\n" + "=" * 52)
    print("   ✅  ¡Web actualizada correctamente!")
    print(f"   🕐  {fecha_hora}")
    print("=" * 52)
    print("\n   Tu URL de GitHub Pages:")
    # Intentar leer el remote para mostrar la URL esperada
    remote = subprocess.run("git remote get-url origin", shell=True,
                            capture_output=True, text=True).stdout.strip()
    if "github.com" in remote:
        partes = remote.replace(".git", "").replace("https://github.com/", "").split("/")
        if len(partes) == 2:
            usuario, repo = partes
            print(f"   https://{usuario}.github.io/{repo}/\n")


if __name__ == "__main__":
    main()
