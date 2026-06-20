# -*- mode: python ; coding: utf-8 -*-
#
# reminder.spec — Configuración de compilación PyInstaller
# Para compilar: pyinstaller reminder.spec
# Resultado:     dist/reminderpagolety.exe

a = Analysis(
    ['reminderpagolety.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Archivos de traducción (i18n) — se copian a src/i18n/ dentro del exe
        ('src/i18n/es.json', 'src/i18n'),
        ('src/i18n/en.json', 'src/i18n'),
        # Icono (también disponible en runtime si la app lo necesita)
        ('reminderagua.ico', '.'),
    ],
    # Módulos que PyInstaller no detecta automáticamente por el uso dinámico de COM
    hiddenimports=[
        'win32com',
        'win32com.client',
        'win32com.server',
        'pythoncom',
        'pywintypes',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='reminderpagolety',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                  # Sin ventana de consola (aplicación de escritorio)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='reminderagua.ico',        # Icono del ejecutable
)
