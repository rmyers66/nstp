# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis([
    'generate_qr_badges_final.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='QR_Badges',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False,
          icon=None)
coll = COLLECT(exe,
               a.binaries,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='QR_Badges')
