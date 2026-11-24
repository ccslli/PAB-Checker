# -*- mode: python ; coding: utf-8 -*-
# PyInstaller recipe for "PAB Checker.exe" (the window version).
# Build with build_exe.bat, or:  py -3.10 -m PyInstaller --noconfirm --clean pab_checker.spec
#
# Same as the plain command line build, but leaves out what the window never uses,
# which is most of the file size:
#   - numpy / pandas / matplotlib etc. (xlwings links to them if they're installed)
#   - Qt's software OpenGL driver, translations, and unused image/style plugins

EXCLUDES = [
    "numpy", "pandas", "matplotlib", "scipy", "PIL",   # (tkinter stays: the splash screen uses Tcl/Tk)
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
    hiddenimports=["branding.dt_icon_data", "scan_vlan_check", "openpyxl"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
)
a.binaries = [b for b in a.binaries if keep(b)]
a.datas = [d for d in a.datas if keep(d)]

pyz = PYZ(a.pure)

# Start-up splash: shown straight away while the one-file .exe unpacks itself;
# the window closes it once it's open (pyi_splash in pab_checker_gui.py).
# "Loading..." with a 0-100% bar instead of each file's name, as in Design Team
# Tools: branding/splash_progress.tcl is added to PyInstaller's splash script,
# with the number of files to unpack.
class ProgressSplash(Splash):
    total_files = 0

    def generate_script(self):
        script = super().generate_script()
        with open(os.path.join(HERE, "branding", "splash_progress.tcl"), encoding="utf-8") as fh:
            script += "\n" + fh.read().replace("@TOTAL@", str(self.total_files))
        with open(self.script_name, "w", encoding="utf-8") as fh:     # keep build\..._script.tcl accurate
            fh.write(script)
        return script


ProgressSplash.total_files = len({entry[0] for entry in a.binaries + a.datas})   # one splash update per file

splash = ProgressSplash(
    os.path.join(HERE, "branding", "splash_pab_checker.png"),
    binaries=a.binaries,
    datas=a.datas,
    text_pos=(22, 226),
    text_size=9,
    # Tcl needs braces around a font name with a space - without them the splash
    # script stops at the font and leaves an empty "tk" window instead of the splash
    text_font="{Segoe UI}",
    text_color="#9AA7B4",
    text_default="Loading...",
    minify_script=True,
)

exe = EXE(
    pyz,
    a.scripts,
    splash,
    splash.binaries,
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
