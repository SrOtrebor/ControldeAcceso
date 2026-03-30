# -*- mode: python ; coding: utf-8 -*-
# =============================================================================
# ControlAcceso.spec — Configuración de empaquetado con PyInstaller
# =============================================================================
# Genera un único archivo .exe que incluye templates, estáticos y
# todas las dependencias necesarias (Flask, Pandas, python-dotenv, etc.)
#
# Uso: pyinstaller ControlAcceso.spec
#
# IMPORTANTE: El archivo .env debe colocarse manualmente junto al .exe
# resultante antes de ejecutarlo (no se incluye en el paquete por seguridad).
# =============================================================================

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Incluir carpetas de templates y estáticos dentro del paquete
        ('templates', 'templates'),
        ('static',    'static'),
    ],
    hiddenimports=[
        # Módulos internos de pandas que PyInstaller no detecta automáticamente
        'pandas._libs.tslibs.base',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.timestamps',
        'pandas._libs.tslibs.offsets',
        'pandas._libs.tslibs.conversion',
        'pandas._libs.tslibs.timezones',
        'pandas._libs.interval',
        'pandas._libs.hashtable',
        'pandas._libs.missing',
        # Motores de escritura Excel
        'openpyxl',
        'xlsxwriter',
        # Carga de .env
        'dotenv',
        # Ventana nativa (pywebview)
        'webview',
        'webview.platforms.winforms',
        'clr',
        'pythonnet',
        'System',
        'System.Windows.Forms',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ControlAcceso',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                # Sin ventana de consola (app de escritorio)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,                    # Agregar ruta a un .ico si se desea
)
