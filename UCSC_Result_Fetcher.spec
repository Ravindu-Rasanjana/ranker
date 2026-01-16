# -*- mode: python ; coding: utf-8 -*-

import sys
block_cipher = None

a = Analysis(
    ['Fetcher_ultimate.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('credits.csv', '.'),  # Include the credits.csv file
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'openpyxl.cell._writer',
        'requests',
        'bs4',
        'smtplib',
        'email',
        'email.mime',
        'email.mime.text',
        'email.mime.multipart',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='UCSC Result Fetcher',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Set to False for GUI app (no terminal window)
    disable_windowed_traceback=False,
    argv_emulation=False,  # IMPORTANT: Set to False to avoid macOS compatibility issues
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='UCSC Result Fetcher',
)

app = BUNDLE(
    coll,
    name='UCSC Result Fetcher.app',
    icon='AppIcon.icns',  # Graduation cap icon
    bundle_identifier='com.ucsc.resultfetcher',
    info_plist={
        'CFBundleName': 'UCSC Result Fetcher',
        'CFBundleDisplayName': 'UCSC Result Fetcher',
        'CFBundleVersion': '2.0.0',
        'CFBundleShortVersionString': '2.0.0',
        'NSHighResolutionCapable': 'True',
        'LSMinimumSystemVersion': '10.13.0',
        'LSEnvironment': {
            'TK_SILENCE_DEPRECATION': '1',
        },
    },
)
