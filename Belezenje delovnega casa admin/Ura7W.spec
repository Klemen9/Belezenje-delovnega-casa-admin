# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['Ura7W.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Main application icon
        ('icon.ico', '.'),
        # Toolbar icons
        ('search.png', '.'),
        ('export.png', '.'),
        ('import.png', '.'),
        ('settings.png', '.'),
        ('refresh.png', '.'),
        ('manual.png', '.'),
        # Context menu and functionality icons
        ('calendar.png', '.'),
        ('lostcard.png', '.'),
        ('archive.png', '.'),
        ('trashcan.png', '.'),
        ('save.png', '.'),
    ],
    hiddenimports=[
        # PyQt6 modules that might not be detected automatically
        'PyQt6.QtCore',
        'PyQt6.QtGui', 
        'PyQt6.QtWidgets',
        'PyQt6.sip',
        # Pandas and data processing
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timezones',
        'pandas._libs.window.aggregations',
        'pandas._libs.window.indexers',
        'pandas.plotting._matplotlib',
        # SMB connection
        'smb.SMBConnection',
        'smb.smb_structs',
        'smb.smb2_structs', 
        'smb.smb2_connection',
        'smb.base',
        # Keyboard input handling
        'keyboard._canonical_names',
        'keyboard._generic',
        'keyboard._winkeyboard',
        # SQLite3 (usually included but ensure compatibility)
        'sqlite3',
        # Configuration parser
        'configparser',
        # Threading and JSON
        'threading',
        'json',
        'io',
        're',
        # Date and time
        'datetime',
        # System modules
        'sys',
        'os',
        'tempfile',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude unnecessary modules to reduce size
        'matplotlib.tests',
        'pandas.tests',
        'numpy.tests',
        'PyQt6.QtNetwork',
        'PyQt6.QtWebEngine',
        'PyQt6.QtWebEngineWidgets',
        'PyQt6.QtPrintSupport',
        'PyQt6.QtOpenGL',
        'PyQt6.QtSql',
        'PyQt6.QtDesigner',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Belezenje delovnega casa admin',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Belezenje delovnega casa'
) 