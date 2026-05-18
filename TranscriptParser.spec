# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['OCR_v2_2.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('zju_logo.png', '.'),
        ('dnd_icon.png', '.'),
    ],
    hiddenimports=[
        
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
    a.zipfiles,
    a.datas,
    [],
    name='TranscriptParser',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,                             
    upx=False,                               
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                          # no console window (windowed app)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='zju_logo.ico'
)

# Strip unused stdlib modules — each saves ~0.1-0.3s startup on slow machines
        #'unittest', 'email', 'html', 'http', 'urllib', 'xmlrpc',
        #'pydoc', 'doctest', 'difflib', 'pickle', 'shelve',
        #'distutils', 'setuptools', 'pkg_resources',
        #'tkinter.test', 'test',