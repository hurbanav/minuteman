block_cipher = None

a = Analysis(
    ['C:\\Users\\bi\\Documents\\BotCase\\Scripts\\XML_Upload\\toolbox\\New_Botcase_App.py'],
    pathex=['C:\\Users\\bi\\Documents\\BotCase\\Scripts\\XML_Upload',
            'C:\\Users\\bi\\Documents\\BotCase\\Scripts\\XML_Upload\\toolbox'],
    binaries=[],
    datas=[
        ('C:\\Users\\bi\\AppData\\Local\\anaconda_true\\Lib\\site-packages\\openpyxl\\', 'openpyxl'),
	('C:\\Users\\bi\\Documents\\BotCase\\Scripts\\XML_Upload\\toolbox', 'toolbox')
    ],
    hiddenimports=['openpyxl', 'upload_xml_v3', 'move_archives', 'toolbox'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='Botcase Solutions',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          icon='C:\\Users\\bi\\Documents\\BotCase\\Dash\\icon.ico'
          ) 