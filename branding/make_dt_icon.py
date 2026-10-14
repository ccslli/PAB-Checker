#!/usr/bin/env python3
"""
make_dt_icon.py - builds the Design Team "DT" icon files in this folder.

    pip install pillow
    python make_dt_icon.py

Writes:
    dt_icon.ico        Windows icon, 16-256 px (use for windows, shortcuts, .exe)
    dt_icon_1024.png   large master (slides, docs, web)
    dt_icon_256.png    medium PNG

Colours come from the team logo: dark slate-teal background, cyan letters.
The letters are drawn as shapes (no font needed), so the icon is identical
on every machine.  Change the values below and re-run to adjust.
"""

import os

from PIL import Image, ImageDraw, ImageFilter

BACKGROUND = "#1E2A31"      # dark slate-teal
LETTERS = "#3CC9CB"         # cyan
GLOW = "#3CC9CB"            # soft halo behind the letters (large sizes only)

S = 1024                    # master size; smaller sizes are scaled from it
ICO_SIZES = [16, 20, 24, 32, 40, 48, 64, 128, 256]
HERE = os.path.dirname(os.path.abspath(__file__))


def letters_mask(scale=0.86):
    """White 'DT' on black: straight-stemmed D, heavy T."""
    m = Image.new("L", (S, S), 0)
    d = ImageDraw.Draw(m)
    h, stroke = 470 * scale, 128 * scale
    dw, tw, gap = 370 * scale, 380 * scale, 44 * scale
    x0 = (S - (dw + gap + tw)) / 2
    y0 = (S - h) / 2
    r = min(h / 2, dw)
    d.rounded_rectangle((x0, y0, x0 + dw, y0 + h), r, fill=255, corners=(False, True, True, False))
    d.rounded_rectangle((x0 + stroke, y0 + stroke, x0 + dw - stroke, y0 + h - stroke),
                        max(1, r - stroke), fill=0, corners=(False, True, True, False))
    tx = x0 + dw + gap
    d.rectangle((tx, y0, tx + tw, y0 + stroke), fill=255)
    d.rectangle((tx + tw / 2 - stroke / 2, y0, tx + tw / 2 + stroke / 2, y0 + h), fill=255)
    return m


def master(with_glow=True):
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(im).rounded_rectangle((40, 40, S - 40, S - 40), 210, fill=BACKGROUND)
    mask = letters_mask()
    if with_glow:
        halo = mask.filter(ImageFilter.GaussianBlur(28)).point(lambda v: int(v * 0.6))
        layer = Image.new("RGBA", (S, S), GLOW)
        layer.putalpha(halo)
        im = Image.alpha_composite(im, layer)
    im.paste(Image.new("RGBA", (S, S), LETTERS), (0, 0), mask)
    return im


def main():
    glow, flat = master(True), master(False)
    # the glow turns to mush below 48 px, so the small sizes use the flat version
    sized = [(flat if sz < 48 else glow).resize((sz, sz), Image.LANCZOS) for sz in ICO_SIZES]
    sized[-1].save(os.path.join(HERE, "dt_icon.ico"), format="ICO",
                   sizes=[(sz, sz) for sz in ICO_SIZES], append_images=sized[:-1])
    glow.save(os.path.join(HERE, "dt_icon_1024.png"))
    glow.resize((256, 256), Image.LANCZOS).save(os.path.join(HERE, "dt_icon_256.png"))
    print("Icon files written to", HERE)


if __name__ == "__main__":
    main()
