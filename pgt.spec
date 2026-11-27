# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

block_cipher = None

heif_datas, heif_binaries, heif_hiddenimports = collect_all('pi_heif')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=heif_binaries,
    datas=heif_datas + [
        ('templates/Template-[0-9]*.docx', 'templates'),
        ('favicon.png', '.'),
        ('tests/fixtures/sample.heic', 'diagnostics'),
    ],
    hiddenimports=[
        'PIL.Image',
        'docxtpl',
        'docx',
        'PyQt6.QtNetwork',
    ] + heif_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        '_tkinter',
        'tkinter',
        'PIL._tkinter_finder',
        'PyQt6.QtPdf',
        'PyQt6.QtPdfWidgets',
        'PIL.AvifImagePlugin',
    ],
    cipher=block_cipher,
    noarchive=True,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# PyQt hooks may collect optional QtPdf DLLs even when their Python modules are excluded.
# No code in this application uses QtPdf; preserve QtNetwork and software OpenGL.
a.binaries = [
    entry for entry in a.binaries
    if not entry[0].lower().endswith(('qt6pdf.dll', 'qt6pdfwidgets.dll'))
]
a.datas = [
    entry for entry in a.datas
    if not entry[0].lower().endswith(('qt6pdf.dll', 'qt6pdfwidgets.dll'))
]

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PGT',
    debug=False,
    strip=False,
    upx=False,
    console=True,
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
