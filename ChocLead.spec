# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['ChocLead.py'],
             pathex=['.\\plugin\\paddleocr', '.\\plugin\\paddle\\libs', '.\\plugin\\scipy\\.libs'],
             binaries=[('.\\plugin\\paddle\\libs', '.')],
             datas=[('.\\Lbt', '.\\Lbt'), ('.\\locale', '.\\locale'), ('.\\Lbt\\General\\lib_custom\\SunKeyboard.lib\\DD94687.64.dll', '.\\Lbt\\General\\lib_custom\\SunKeyboard.lib\\'), ('.\\plugin\\paddleocr\\ppocr\\utils\\ppocr_keys_v1.txt', '.\\ppocr\\utils\\'), ('.\\TecleadContract.docx', '.')],
             hiddenimports=['ChocLead'],
             hookspath=[],
             runtime_hooks=[],
             excludes=['matplotlib'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='ChocLead',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='ChocLead',
               console=False , version='versionInfo.txt', icon='robot.ico')