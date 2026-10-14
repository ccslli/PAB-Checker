# -*- mode: python ; coding: utf-8 -*-
# PyInstaller recipe for "PAB Checker.exe" (the window version).
# Build with build_exe.bat, or:  py -3.10 -m PyInstaller --noconfirm --clean pab_checker.spec
#
# Same as the plain command line build, but leaves out what the window never uses,
# which is most of the file size:
#   - numpy / pandas / matplotlib etc. (xlwings links to them if they're installed)
#   - Qt's software OpenGL driver, translations, and unused image/style plugins

EXCLUDES = [
    "numpy", "pandas", "matplotlib", "scipy", "PIL", "tkinter",
    "IPython", "jupyter_client", "notebook", "pytest",
]

# Qt files dropped from the bundle (matched against their path inside the .exe)
DROP = (
    "opengl32sw.dll",              # software OpenGL fallback, ~20 MB
    "/translations/",              # Qt's other-language texts
    "/plugins/styles/",            # the window uses the built-in Fusion style
    "/plugins/iconengines/",
    "/plugins/generic/",
    "/plugins/networkinformation/",
    "/plugins/tls/",
    "qt6pdf", "qt6qml", "qt6quick", "qt6virtualkeyboard",
)
KEEP_IMAGE_PLUGINS = ("qico",)     # .ico support for the DT icon (PNG is built in)


def keep(entry):
    path = entry[0].replace("\\", "/").lower()
    if any(d in path for d in DROP):
        return False
    if "/plugins/imageformats/" in path:
        return any(k in path for k in KEEP_IMAGE_PLUGINS)
    return True


import os
HERE = SPECPATH                   # folder of this .spec file (set by PyInstaller)

a = Analysis(
    [os.path.join(HERE, "pab_checker_gui.py")],
    pathex=[HERE],
    binaries=[],
    datas=[(os.path.join(HERE, "branding"), "branding")],
    hiddenimports=["branding.dt_icon_data"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
)
a.binaries = [b for b in a.binaries if keep(b)]
a.datas = [d for d in a.datas if keep(d)]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="PAB Checker",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    icon=os.path.join(HERE, "branding", "dt_icon.ico"),
)
