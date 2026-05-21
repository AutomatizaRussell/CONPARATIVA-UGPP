# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['app.py'],
    pathex=['convertidor'],          # PyInstaller busca módulos aquí también
    binaries=[],
    datas=[],
    hiddenimports=[
        # Scripts del proyecto (importados dinámicamente en app.py)
        'factura',
        'convertidor',
        'excel',
        'excel_a_json',
        'poblar_excel',
        # pdfplumber y sus dependencias
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        'pdfminer.pdfpage',
        'pdfminer.converter',
        'pdfminer.pdfdocument',
        'pdfminer.pdfinterp',
        'pdfminer.pdfparser',
        'pdfminer.pdftypes',
        'pdfminer.pdfdevice',
        'pdfminer.psparser',
        'pdfminer.utils',
        'pdfminer.cmapdb',
        'pdfminer.encodingdb',
        'pdfminer.image',
        # pandas
        'pandas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.skiplist',
        # openpyxl
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.styles.fonts',
        'openpyxl.styles.fills',
        'openpyxl.styles.borders',
        'openpyxl.styles.alignment',
        'openpyxl.utils',
        'openpyxl.utils.cell',
        'et_xmlfile',
        # otros
        'PIL',
        'PIL.Image',
        'pypdfium2',
        'cryptography',
        'cryptography.hazmat.primitives.ciphers',
        'cryptography.hazmat.primitives.ciphers.algorithms',
        'cryptography.hazmat.backends',
        'cryptography.hazmat.backends.openssl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ComparativaUGPP',
    debug=False,
    strip=False,
    upx=True,
    console=False,          # Sin ventana de consola (app gráfica)
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ComparativaUGPP',
)
