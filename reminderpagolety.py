"""
reminderpagolety.py
-------------------
Punto de entrada de la aplicación ReminderMailletYQuincena.
Este archivo es intencionalmente mínimo: solo ajusta el path y arranca la GUI.

Para compilar con PyInstaller:
    pyinstaller --noconfirm --clean --distpath . --workpath build/pyinstaller reminder.spec

Para ejecutar en modo desarrollo:
    python reminderpagolety.py
"""
import sys
import os

# Asegurar que el directorio raíz del proyecto esté en sys.path.
# Esto permite que los imports 'from src.xxx' funcionen tanto al
# correr el script directamente como desde el ejecutable compilado.
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if _BASE_DIR not in sys.path:
    sys.path.insert(0, _BASE_DIR)

from src.gui.main_window import run_app  # noqa: E402

if __name__ == "__main__":
    run_app()
