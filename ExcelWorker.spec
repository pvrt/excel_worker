# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('credentials.json', '.'), ('token.json', '.')] if __import__('os').path.exists('credentials.json') else [],
    hiddenimports=[
        'googleapiclient.discovery',
        'googleapiclient.http',
        'google.oauth2.credentials',
        'google_auth_oauthlib.flow',
        'google.auth.transport.requests',
        'reportlab.lib.colors',
        'reportlab.lib.pagesizes',
        'reportlab.lib.styles',
        'reportlab.platypus',
        'openpyxl',
        'PIL',
        'pikepdf',
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
    name='ExcelWorker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
