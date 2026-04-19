# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates/*.docx', 'templates'),
        ('favicon.png', '.'),
        ('config.py', '.'),
    ],
    hiddenimports=[
        'PIL.Image',
        'docxtpl',
        'docx',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        '_tkinter',
        'tkinter',
        'PIL._tkinter_finder',
    ],
    cipher=block_cipher,
    noarchive=True,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PGT',
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon='favicon.png',
    version_file='version.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name='PGT',
)
