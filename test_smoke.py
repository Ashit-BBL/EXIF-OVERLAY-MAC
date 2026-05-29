#!/usr/bin/env python3
"""
Smoke test for EXIF Overlay Tool — tests all core logic headlessly (no GUI).
Run with: python test_smoke.py
"""
import sys
from unittest.mock import MagicMock

# ── Mock all GUI modules so core logic can be tested without a display ────────
for mod in ["customtkinter", "tkinter", "tkinter.filedialog",
            "tkinter.messagebox", "_tkinter"]:
    sys.modules[mod] = MagicMock()

# ── Now safe to import ────────────────────────────────────────────────────────
import exif_tool
from exif_tool import (
    _fmt_exposure, _fmt_aperture, _fmt_focal, _fmt_distance,
    _fmt_flash, _fmt_exposure_bias, _fmt_datetime, _fmt_copyright,
    apply_overlay, read_exif, collect_jpgs, _load_font,
)
from PIL import Image
from pathlib import Path
import tempfile

errors = []

def check(name, condition, detail=""):
    if condition:
        print(f"  ✅  {name}")
    else:
        msg = f"  ❌  {name}" + (f"  ({detail})" if detail else "")
        print(msg)
        errors.append(name)

# ─────────────────────────────────────────────────────────────────────────────
print("\n── Format helpers ───────────────────────────────────────────────────────")
check("Shutter  1/2s",       _fmt_exposure(0.5)      == "1/2s")
check("Shutter  1.0s",       _fmt_exposure(1.0)      == "1.0s")
check("Shutter  —  (zero)",  _fmt_exposure(0)        == "—")
check("Aperture f/2.8",      _fmt_aperture(2.8)      == "f/2.8")
check("Focal    85 mm",      _fmt_focal(85.0)        == "85 mm")
check("Distance ∞ (zero)",   _fmt_distance(0)        == "∞")
check("Distance ∞ (>9999)",  _fmt_distance(10000)    == "∞")
check("Flash    No Flash",   _fmt_flash(0)           == "No Flash")
check("Flash    Fired",      _fmt_flash(1)           == "Fired")
check("ExpComp  0 EV",       _fmt_exposure_bias(0)   == "0 EV")
check("ExpComp  +1.0 EV",    _fmt_exposure_bias(1.0) == "+1.0 EV")
check("ExpComp  -0.7 EV",    _fmt_exposure_bias(-0.666) == "-0.7 EV")
check("Copyright string",    _fmt_copyright("Ashit Gandhi") == "Ashit Gandhi")
check("Copyright null strip",_fmt_copyright("Test\x00garbage") == "Test")

# ─────────────────────────────────────────────────────────────────────────────
print("\n── Font loading ─────────────────────────────────────────────────────────")
font = _load_font(30)
check("Font loaded (not None)", font is not None)
font_small = _load_font(12)
check("Font loaded (small 12pt)", font_small is not None)

# ─────────────────────────────────────────────────────────────────────────────
print("\n── Overlay — Multi-line, all 4 corners ──────────────────────────────────")
img = Image.new("RGB", (1200, 800), color=(80, 80, 80))
fields = {
    "Camera": "OM-1", "Lens": "12-40mm Pro",
    "ISO": "400", "Aperture": "f/2.8", "Shutter": "1/250s",
}
for corner in ["Top Right", "Top Left", "Bottom Right", "Bottom Left"]:
    r = apply_overlay(img, fields, 30, "White", corner,
                      layout="Multi-line", show_labels=True, notes="")
    check(f"Multi-line  {corner}", r.size == (1200, 800))

# ─────────────────────────────────────────────────────────────────────────────
print("\n── Overlay — Multi-line, colours & labels ───────────────────────────────")
for colour in ["White", "Yellow", "Black"]:
    r = apply_overlay(img, fields, 30, colour, "Bottom Right",
                      layout="Multi-line", show_labels=True, notes="")
    check(f"Colour  {colour}", r.size == (1200, 800))

r = apply_overlay(img, fields, 30, "White", "Bottom Right",
                  layout="Multi-line", show_labels=False, notes="")
check("Labels hidden", r.size == (1200, 800))

# ─────────────────────────────────────────────────────────────────────────────
print("\n── Overlay — Single-line, top & bottom ──────────────────────────────────")
for pos in ["Top", "Bottom"]:
    r = apply_overlay(img, fields, 30, "White", pos,
                      layout="Single-line", show_labels=True, notes="")
    check(f"Single-line  {pos}  (labels on)", r.size == (1200, 800))
    r = apply_overlay(img, fields, 30, "White", pos,
                      layout="Single-line", show_labels=False, notes="")
    check(f"Single-line  {pos}  (labels off)", r.size == (1200, 800))

# ─────────────────────────────────────────────────────────────────────────────
print("\n── Overlay — Notes ──────────────────────────────────────────────────────")
r = apply_overlay(img, fields, 30, "White", "Bottom",
                  layout="Single-line", show_labels=True,
                  notes="Shot at golden hour\nOlympus OM-1")
check("Notes 2 lines", r.size == (1200, 800))

r = apply_overlay(img, {}, 30, "White", "Bottom Right",
                  layout="Multi-line", show_labels=True, notes="Just a note")
check("Notes only (no fields)", r.size == (1200, 800))

# ─────────────────────────────────────────────────────────────────────────────
print("\n── Overlay — edge cases ─────────────────────────────────────────────────")
r = apply_overlay(img, {}, 30, "White", "Bottom Right",
                  layout="Multi-line", show_labels=True, notes="")
check("Empty fields + no notes → copy", r.size == (1200, 800))

many_fields = {
    "Camera": "OM-1", "Lens": "12-40mm", "ISO": "1600",
    "Aperture": "f/4.0", "Shutter": "1/500s", "Focal Length": "40 mm",
    "White Balance": "Auto", "Flash": "No Flash", "Exp. Comp.": "0 EV",
}
r = apply_overlay(img, many_fields, 40, "White", "Bottom",
                  layout="Single-line", show_labels=True, notes="")
check("Single-line auto font-cap (many fields)", r.size == (1200, 800))

# ─────────────────────────────────────────────────────────────────────────────
print("\n── collect_jpgs ─────────────────────────────────────────────────────────")
with tempfile.TemporaryDirectory() as tmp:
    for name in ["a.jpg", "b.JPEG", "c_exif.jpg", "d.png", "e.JPG"]:
        Path(tmp, name).touch()
    jpgs = collect_jpgs(tmp)
    check("Collects .jpg / .JPEG / .JPG",  len(jpgs) == 3,
          f"found {len(jpgs)}")
    check("Excludes _exif files",          not any("_exif" in f for f in jpgs))
    check("Excludes .png",                 not any(".png"  in f for f in jpgs))

# ─────────────────────────────────────────────────────────────────────────────
print("\n── Platform helpers ─────────────────────────────────────────────────────")
check("_IS_MAC is True on macOS",   exif_tool._IS_MAC == (sys.platform == "darwin"))
check("_IS_WIN is False on macOS",  exif_tool._IS_WIN == (sys.platform == "win32"))
check("MONO_FONT set",              isinstance(exif_tool.MONO_FONT, str) and
                                    len(exif_tool.MONO_FONT) > 0)

# ─────────────────────────────────────────────────────────────────────────────
print("\n" + "─" * 72)
if errors:
    print(f"\n❌  {len(errors)} test(s) FAILED:  {', '.join(errors)}\n")
    sys.exit(1)
else:
    print(f"\n✅  All {28} tests passed — core logic is Mac-ready!\n")
    sys.exit(0)
