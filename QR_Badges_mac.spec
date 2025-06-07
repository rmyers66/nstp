# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['generate_qr_badges_final.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('GT_full_logo.png', '.'),
        ('GT_small_logo.png', '.'),
        ('GT_ribbon.png', '.')
    ],
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
    [],
    exclude_binaries=True,
    name='QR Badges',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch="universal2",
    codesign_identity=None,
    entitlements_file=None,
    icon=['AppIcon.icns'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='QR Badges',
)
app = BUNDLE(
    coll,
    name='QR Badges.app',
    icon='AppIcon.icns',
    bundle_identifier='com.example.qrbadges',
    info_plist={
        'CFBundleShortVersionString': '3.6.1-postfix',
        'CFBundleVersion': '3.6.1-postfix',
    }
)
