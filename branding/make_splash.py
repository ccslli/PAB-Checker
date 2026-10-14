#!/usr/bin/env python3
"""
make_splash.py - builds a Design Team start-up splash image for an app.

    pip install pillow
    python make_splash.py "PAB Checker"

Writes splash_<app_name>.png in this folder.  PyInstaller shows it while a
one-file .exe unpacks (see pab_checker.spec), with its own status text in the
strip along the bottom.  Run it on Windows to get the Bahnschrift title font.
"""

import os
import sys

from PIL import Image, ImageDraw, ImageFont

import make_dt_icon as icon

W, H = 480, 240
BACKGROUND = "#161D22"          # a shade darker than the icon tile, so the tile stands out
TITLE = "#E4E5E7"
SUBTITLE = icon.GLOW            # team cyan
RULE = "#2A3A42"
HERE = os.path.dirname(os.path.abspath(__file__))
TEXT_STRIP_Y = 204              # PyInstaller's status text goes below this line (see the .spec)


def font(names, size):
    for name in names:
        for path in (name, os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", name)):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                pass
    return ImageFont.load_default()


def main(app_name):
    im = Image.new("RGB", (W, H), BACKGROUND)
    d = ImageDraw.Draw(im)
    logo = icon.master(True).resize((128, 128), Image.LANCZOS)
    im.paste(logo, (32, 40), logo)
    title = font(["bahnschrift.ttf", "segoeuib.ttf", "DejaVuSans-Bold.ttf"], 34)
    sub = font(["segoeui.ttf", "DejaVuSans.ttf"], 15)
    d.text((184, 70), app_name, fill=TITLE, font=title)
    d.text((186, 120), "Design Team", fill=SUBTITLE, font=sub)
    d.line((0, TEXT_STRIP_Y, W, TEXT_STRIP_Y), fill=RULE, width=1)
    out = os.path.join(HERE, "splash_%s.png" % app_name.lower().replace(" ", "_"))
    im.save(out)
    print("Splash written to", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "PAB Checker")
