# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['Integrated_Tool_ByTAX.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config.txt', '.'),
        ('favicon.ico', '.'),
        ('Bldbuy_Recon_ByTAX.py', '.'),
        ('Product_Classification_Tool_ByTAX.py', '.'),
    ],
    hiddenimports=[
        'Bldbuy_Recon_ByTAX',
        'Product_Classification_Tool_ByTAX',
        'numpy',
        'tkinter.constants',
        'datetime',
        'warnings',
        'glob',
        'threading',
        'shutil',
        'logging',
        'pandas',
        'openpyxl',
        'xlrd',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    cipher=block_cipher,
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='供应商对账工具集',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # 禁用UPX压缩以减少误报
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='favicon.ico',
)