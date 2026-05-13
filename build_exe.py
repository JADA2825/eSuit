"""
build_exe.py — Construye eSuit.exe con PyInstaller.

Uso:
    python build_exe.py             # modo carpeta (recomendado, arranca rápido)
    python build_exe.py --onefile   # un solo .exe (más lento al arrancar)

Salida:
    dist/eSuit/eSuit.exe   (modo onedir)
    dist/eSuit.exe         (modo onefile)
"""
import sys
import shutil
import subprocess
from pathlib import Path

ROOT = Path(__file__).parent
DIST = ROOT / "dist"
BUILD = ROOT / "build"


def limpiar_anterior():
    for p in (DIST, BUILD):
        if p.exists():
            print(f"Eliminando {p}...")
            shutil.rmtree(p, ignore_errors=True)
    for spec in ROOT.glob("*.spec"):
        try:
            spec.unlink()
        except OSError:
            pass


def construir(onefile: bool = False):
    sep = ";" if sys.platform.startswith("win") else ":"
    archivos_proyecto = [
        "app.py", "calculos.py", "reporte.py",
        "reporte_docx.py", "auth.py", "cuadro_cargas.py",
    ]
    add_data_args = []
    for f in archivos_proyecto:
        if (ROOT / f).exists():
            add_data_args += ["--add-data", f"{f}{sep}."]

    modo = "--onefile" if onefile else "--onedir"
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=eSuit",
        modo,
        "--console",        # ventana de consola visible (logs + cierre limpio)
        "--noconfirm",
        "--clean",
        *add_data_args,
        # Streamlit y dependencias
        "--collect-all=streamlit",
        "--collect-all=altair",
        "--collect-data=streamlit",
        "--copy-metadata=streamlit",
        # PDF, Word, Excel
        "--collect-all=reportlab",
        "--collect-all=docx",      # python-docx
        "--collect-all=openpyxl",
        "--collect-all=pandas",
        # Hidden imports
        "--hidden-import=streamlit.web.cli",
        "--hidden-import=streamlit.runtime.scriptrunner.magic_funcs",
        "--hidden-import=streamlit.runtime.caching",
        "--hidden-import=streamlit.runtime.scriptrunner_utils.script_run_context",
        "--hidden-import=reportlab.lib",
        "--hidden-import=reportlab.platypus",
        "--hidden-import=reportlab.pdfbase.ttfonts",
        "--hidden-import=docx",
        "--hidden-import=openpyxl",
        # Entry point
        "launcher.py",
    ]
    print("=" * 60)
    print("Comando PyInstaller:")
    print("  " + " ".join(cmd))
    print("=" * 60)

    res = subprocess.run(cmd, cwd=ROOT)
    if res.returncode != 0:
        print(f"\n[ERROR] Build FALLÓ con código {res.returncode}")
        sys.exit(res.returncode)

    if onefile:
        exe = DIST / "eSuit.exe"
    else:
        exe = DIST / "eSuit" / "eSuit.exe"

    if exe.exists():
        size_mb = exe.stat().st_size / 1024 / 1024
        print(f"\n[OK] Build OK")
        print(f"   Ejecutable: {exe}  ({size_mb:.1f} MB)")
        if not onefile:
            carpeta = DIST / "eSuit"
            print(f"   Distribución completa: {carpeta}")
            print(f"   (Toda la carpeta debe distribuirse junta)")
    else:
        print(f"\n[WARN] Build terminó pero {exe} no existe.")


if __name__ == "__main__":
    onefile = "--onefile" in sys.argv
    limpiar_anterior()
    construir(onefile=onefile)
