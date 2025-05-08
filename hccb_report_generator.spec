# -*- mode: python ; coding: utf-8 -*-
import os
import shutil
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Define output paths
DIST_PATH = os.path.abspath('dist')
RELEASE_PATH = os.path.abspath('release')
EXE_NAME = 'HCCB Excel Report & Summary Generator'

a = Analysis(
    ['main_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('report_template.xlsx', '.'),
        ('summary_report_template.xlsx', '.'),
        ('app_icon.ico', '.'),
        ('README.txt', '.'),
    ],
    hiddenimports=['pandas', 'openpyxl', 'tkinter'],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=EXE_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False for no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico'
)

# Post-build: Create release folder and copy necessary files
def create_release_package():
    print("\n=== Creating Release Package ===")
    
    # Create release directory if it doesn't exist
    if not os.path.exists(RELEASE_PATH):
        os.makedirs(RELEASE_PATH)
        print(f"Created release directory: {RELEASE_PATH}")
    
    # Copy the executable
    exe_src = os.path.join(DIST_PATH, f"{EXE_NAME}.exe")
    exe_dst = os.path.join(RELEASE_PATH, f"{EXE_NAME}.exe")
    if os.path.exists(exe_src):
        shutil.copy2(exe_src, exe_dst)
        print(f"Copied: {exe_src} -> {exe_dst}")
    else:
        print(f"WARNING: Executable not found at {exe_src}")
    
    # Copy template files and README
    files_to_copy = [
        'report_template.xlsx',
        'summary_report_template.xlsx',
        'README.txt'
    ]
    
    for file in files_to_copy:
        if os.path.exists(file):
            dst = os.path.join(RELEASE_PATH, file)
            shutil.copy2(file, dst)
            print(f"Copied: {file} -> {dst}")
        else:
            print(f"WARNING: File not found: {file}")
    
    print("\nRelease package created successfully!")
    print(f"Location: {RELEASE_PATH}")
    print("===========================")

# Run the post-build function
create_release_package() 