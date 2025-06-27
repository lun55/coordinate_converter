# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['ConvertApp.py'],
    pathex=['.'],  # 当前目录
    binaries=[],
    datas=[],  # 如果有图标等资源文件，可以添加在这里
    hiddenimports=[
        'coordinateConverter',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.sas',
        'pandas._libs',
        'openpyxl',
        'PyQt5.QtPrintSupport'  # 如果使用了打印功能需要添加
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    noarchive=False,
    cipher=block_cipher,
    optimize=1,  # 使用优化级别1
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ConvertApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 启用UPX压缩
    upx_exclude=[],  # 可以排除某些不需要压缩的dll
    runtime_tmpdir=None,
    console=False,  # 因为是GUI应用
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon= None  # 确保这个图标文件存在
)