# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app_gradio.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Código\\sad_auto\\auto_venv\\Lib\\site-packages\\safehttpx', 'safehttpx'), ('C:\\Código\\sad_auto\\auto_venv\\Lib\\site-packages\\groovy', 'groovy'), ('C:\\Código\\sad_auto\\auto_venv\\Lib\\site-packages\\gradio', 'gradio')],
    hiddenimports=[],
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
    name='Automacao_IPCA',
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
d