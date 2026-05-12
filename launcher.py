"""
launcher.py — Punto de entrada de eSuit.

Arranca el servidor Streamlit en localhost y abre el navegador automáticamente.
Funciona tanto al ejecutar `python launcher.py` como cuando está empaquetado
con PyInstaller (`eSuit.exe`).
"""
import os
import sys
import threading
import time
import socket
import webbrowser
from pathlib import Path


def _puerto_libre(preferido: int = 8501) -> int:
    """Devuelve un puerto libre, intentando el preferido primero."""
    for puerto in (preferido, 8502, 8503, 8504, 8505, 0):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", puerto))
                return s.getsockname()[1]
            except OSError:
                continue
    return preferido


def _abrir_navegador(url: str, delay: float = 2.0) -> None:
    time.sleep(delay)
    try:
        webbrowser.open(url)
    except Exception:
        pass


def main() -> None:
    # Localización de app.py — funciona tanto en dev como empaquetado
    if getattr(sys, "frozen", False):
        # PyInstaller — los assets viven en sys._MEIPASS
        base = Path(getattr(sys, "_MEIPASS", os.path.dirname(sys.executable)))
    else:
        base = Path(__file__).parent

    app_path = base / "app.py"
    if not app_path.exists():
        print(f"ERROR: no se encontró app.py en {base}")
        sys.exit(1)

    puerto = _puerto_libre(8501)
    url = f"http://localhost:{puerto}"
    print(f"╔═══════════════════════════════════════════════╗")
    print(f"║   eSuit — Cálculo eléctrico profesional       ║")
    print(f"║   Abriendo en {url:<31} ║")
    print(f"╚═══════════════════════════════════════════════╝")

    # Lanza el navegador después de unos segundos en un hilo aparte
    threading.Thread(target=_abrir_navegador,
                      args=(url, 2.0), daemon=True).start()

    # Configura argv y arranca Streamlit
    from streamlit.web import cli as stcli
    sys.argv = [
        "streamlit", "run", str(app_path),
        f"--server.port={puerto}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
        "--server.fileWatcherType=none",
    ]
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
