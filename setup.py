"""
py2app build script for UCSC Result Fetcher
Run with: python3 setup.py py2app
"""

from setuptools import setup

APP = ['Fetcher_ultimate.py']
DATA_FILES = [
    ('', ['credits.csv']),  # Include credits.csv at the root of the app
]

OPTIONS = {
    'argv_emulation': False,  # Critical: disable to avoid Tk compatibility issues
    'iconfile': None,  # Add .icns file path here if you have one
    'plist': {
        'CFBundleName': 'UCSC Result Fetcher',
        'CFBundleDisplayName': 'UCSC Result Fetcher',
        'CFBundleIdentifier': 'com.ucsc.resultfetcher',
        'CFBundleVersion': '2.0.0',
        'CFBundleShortVersionString': '2.0.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.13.0',
    },
    'packages': [
        'pandas',
        'openpyxl',
        'requests',
        'bs4',
        'certifi',
        'charset_normalizer',
        'urllib3',
        'idna',
        'soupsieve',
    ],
    'includes': [
        'tkinter',
        'tkinter.ttk',
        'tkinter.scrolledtext',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'email.mime.text',
        'email.mime.multipart',
    ],
    'excludes': ['matplotlib', 'scipy', 'numpy.testing'],  # Exclude large unused packages
}

setup(
    app=APP,
    name='UCSC Result Fetcher',
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
