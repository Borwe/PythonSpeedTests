# -*- mode: python -*-

block_cipher = None


a = Analysis(['serena_speeds_run.py'],
             pathex=['/home/brian/Documents/Python/PythonSpeedTests/dist'],
             binaries=[],
             datas=[],
             hiddenimports=['speedtest'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          name='serena_speeds_run',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
