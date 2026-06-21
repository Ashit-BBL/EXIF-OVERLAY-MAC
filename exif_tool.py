"""
EXIF Overlay Tool  v1.4
Burn EXIF data onto photos with full customisation.
Created by Ashit Gandhi — June 2026
Requires: pip install customtkinter pillow
"""

import json
import os
import sys
import threading

# Hide the console window immediately on Windows (Nuitka onefile bootstrap creates one)
if sys.platform == "win32":
    import ctypes as _ctypes
    _hwnd = _ctypes.windll.kernel32.GetConsoleWindow()
    if _hwnd:
        _ctypes.windll.user32.ShowWindow(_hwnd, 0)   # SW_HIDE
from datetime import datetime
from fractions import Fraction
from pathlib import Path
import tkinter as _tk
from tkinter import filedialog, messagebox

try:
    import customtkinter as ctk
    from PIL import Image, ImageDraw, ImageFont, ImageTk
    import PIL.ExifTags
except ImportError as e:
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("Missing library",
        f"{e}\n\nOpen PowerShell and run:\n  pip install customtkinter pillow")
    sys.exit(1)

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("dark-blue")

C_BG       = "#e8e2d0"
C_CARD     = "#d4cdb5"
C_HEADER   = "#38332a"
C_ACCENT   = "#7d7256"
C_HOVER    = "#625848"
C_LABEL    = "#6b614e"
C_TEXT     = "#2e2a20"
C_DIM      = "#a89e84"
C_SEP      = "#bcb49c"
C_PROG     = "#8a7e62"
C_HDR_TEXT = "#ffffff"
C_HDR_SUB  = "#cfc9b4"

TEXT_COLORS = {
    "White":  (255, 255, 255),
    "Yellow": (255, 220,  50),
    "Black":  (  0,   0,   0),
}
CORNER_OPTIONS = ["Top Right", "Top Left", "Bottom Right", "Bottom Left"]
STRIP_OPTIONS  = ["Top", "Bottom", "Caption"]

EXIF_FIELD_LABELS = [
    "Camera", "Lens", "Shutter", "Aperture",
    "ISO", "White Balance", "Focal Length", "Subject Dist.",
    "Flash", "Exp. Comp.", "Metering", "Exp. Program",
    "Date & Time", "Copyright", "GPS",
]

_APP_DIR = Path(sys.argv[0]).parent if not sys.argv[0].endswith('.py') else Path(__file__).parent
LOGO_FILE   = _APP_DIR / "ashitg - new2 - White.png"
CONFIG_FILE = _APP_DIR / "config.json"
RENDER_MAX = 2000
DISPLAY_W  = 1100
DISPLAY_H  = 650

# Font size slider is calibrated to this longest-edge reference (px).
# Each image's font is scaled by (its longest edge / FONT_REF_LONG_EDGE) so
# portrait and landscape frames from the same camera look visually identical.
FONT_REF_LONG_EDGE = 6000

_IS_MAC = sys.platform == "darwin"
_IS_WIN = sys.platform == "win32"
MONO_FONT = "Menlo" if _IS_MAC else ("Consolas" if _IS_WIN else "DejaVu Sans Mono")


# ─────────────────────────────────────────────────────────────────────────────
#  System font discovery
# ─────────────────────────────────────────────────────────────────────────────

_SYSTEM_FONTS_CACHE = None   # {display_name: absolute_path}

def _scan_system_fonts():
    """Walk OS font directories and return {stem_name: path} for all TTF/OTF files."""
    dirs = []
    if sys.platform == "win32":
        windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
        dirs.append(windir / "Fonts")
        local = os.environ.get("LOCALAPPDATA", "")
        if local:
            dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    elif sys.platform == "darwin":
        dirs = [
            Path("/Library/Fonts"),
            Path("/System/Library/Fonts"),
            Path.home() / "Library" / "Fonts",
        ]
    else:
        dirs = [Path("/usr/share/fonts"), Path.home() / ".fonts"]

    fonts = {}
    for d in dirs:
        if not d.exists():
            continue
        for f in sorted(d.rglob("*")):
            if f.suffix.lower() in (".ttf", ".otf") and f.is_file():
                if f.stem not in fonts:           # first path wins
                    fonts[f.stem] = str(f)
    return fonts


def get_system_fonts():
    """Return cached {name: path} dict of all system fonts."""
    global _SYSTEM_FONTS_CACHE
    if _SYSTEM_FONTS_CACHE is None:
        _SYSTEM_FONTS_CACHE = _scan_system_fonts()
    return _SYSTEM_FONTS_CACHE


# ─────────────────────────────────────────────────────────────────────────────
#  EXIF format helpers
# ─────────────────────────────────────────────────────────────────────────────

_WB_LABELS = {0: "Auto", 1: "Manual"}
_METERING_LABELS = {
    0: "Unknown", 1: "Average", 2: "Centre Weighted",
    3: "Spot", 4: "Multi Spot", 5: "Matrix", 6: "Partial", 255: "Other",
}
_EXPOSURE_PROG = {
    0: "Not Defined", 1: "Manual", 2: "Auto",
    3: "Aperture Priority", 4: "Shutter Priority",
    5: "Creative", 6: "Action", 7: "Portrait", 8: "Landscape",
}

def _rational_to_float(v):
    try:    return float(v)
    except TypeError: pass
    if isinstance(v, tuple) and len(v) == 2:
        n, d = v; return n / d if d else 0.0
    return 0.0

def _fmt_exposure(v):
    f = _rational_to_float(v)
    if f == 0: return "—"
    if f >= 1: return f"{f:.1f}s"
    return f"1/{Fraction(f).limit_denominator(8000).denominator}s"

def _fmt_aperture(v):
    f = _rational_to_float(v); return f"f/{f:.1f}" if f else "—"

def _fmt_focal(v):
    f = _rational_to_float(v); return f"{f:.0f} mm" if f else "—"

def _fmt_distance(v):
    f = _rational_to_float(v)
    if f == 0 or f > 9999: return "∞"
    return f"{f:.2f} m"

def _fmt_flash(v):
    val   = int(v)
    fired = bool(val & 0x01)
    mode  = (val >> 3) & 0x03
    modes = {0: "", 1: " · On", 2: " · Off", 3: " · Auto"}
    return ("Fired" if fired else "No Flash") + modes.get(mode, "")

def _fmt_exposure_bias(v):
    f = _rational_to_float(v)
    if f == 0: return "0 EV"
    return f"{f:+.1f} EV"

def _fmt_datetime(v):
    try:
        dt = datetime.strptime(str(v).strip(), "%Y:%m:%d %H:%M:%S")
        return dt.strftime("%d %b %Y   %H:%M")
    except Exception:
        return str(v)

def _fmt_copyright(v):
    if isinstance(v, bytes):
        try:    v = v.decode("utf-8")
        except: v = v.decode("latin-1", errors="replace")
    elif isinstance(v, str):
        try:    v = v.encode("latin-1").decode("utf-8")
        except: pass
    return str(v).split("\x00")[0].strip() or None

def _fmt_gps(gps_ifd):
    try:
        lat_vals = gps_ifd.get(2)
        lat_ref  = gps_ifd.get(1, "N")
        lon_vals = gps_ifd.get(4)
        lon_ref  = gps_ifd.get(3, "E")
        if not (lat_vals and lon_vals): return None
        def dms(t):
            return (_rational_to_float(t[0])
                    + _rational_to_float(t[1]) / 60
                    + _rational_to_float(t[2]) / 3600)
        return f"{dms(lat_vals):.5f}°{lat_ref}  {dms(lon_vals):.5f}°{lon_ref}"
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
#  EXIF reader
# ─────────────────────────────────────────────────────────────────────────────

def read_exif(image):
    try:    exif = image.getexif()
    except: return {}
    if not exif: return {}
    raw = dict(exif)
    try:    raw.update(exif.get_ifd(0x8769))
    except: pass
    t = {PIL.ExifTags.TAGS.get(k, k): v for k, v in raw.items()}
    gps_ifd = {}
    try:    gps_ifd = exif.get_ifd(0x8825)
    except: pass
    def get(*names):
        for n in names:
            if n in t and t[n] not in (None, "", b""): return t[n]
        return None
    fd = {}
    m = get("Model")
    if m: fd["Camera"] = str(m).strip()
    ln = get("LensModel", "LensSpecification")
    if ln:
        if isinstance(ln, str): fd["Lens"] = ln.strip()
        else:
            try:
                a, b = _rational_to_float(ln[0]), _rational_to_float(ln[1])
                fd["Lens"] = f"{a:.0f}–{b:.0f} mm" if a != b else f"{a:.0f} mm"
            except: pass
    e  = get("ExposureTime")
    if e  is not None: fd["Shutter"]       = _fmt_exposure(e)
    fn = get("FNumber")
    if fn is not None: fd["Aperture"]      = _fmt_aperture(fn)
    i  = get("ISOSpeedRatings")
    if i  is not None: fd["ISO"]           = str(i)
    wb = get("WhiteBalance")
    if wb is not None: fd["White Balance"] = _WB_LABELS.get(int(wb), str(wb))
    fc = get("FocalLength")
    if fc is not None: fd["Focal Length"]  = _fmt_focal(fc)
    d  = get("SubjectDistance")
    if d  is not None: fd["Subject Dist."] = _fmt_distance(d)
    fl = get("Flash")
    if fl is not None: fd["Flash"]         = _fmt_flash(fl)
    bias = get("ExposureBiasValue")
    if bias is not None: fd["Exp. Comp."]  = _fmt_exposure_bias(bias)
    mm = get("MeteringMode")
    if mm is not None: fd["Metering"]      = _METERING_LABELS.get(int(mm), str(mm))
    ep = get("ExposureProgram")
    if ep is not None: fd["Exp. Program"]  = _EXPOSURE_PROG.get(int(ep), str(ep))
    dt = get("DateTimeOriginal", "DateTime")
    if dt: fd["Date & Time"]               = _fmt_datetime(dt)
    cp = get("Copyright")
    if cp:
        cps = _fmt_copyright(cp)
        if cps: fd["Copyright"]            = cps
    if gps_ifd:
        gps = _fmt_gps(gps_ifd)
        if gps: fd["GPS"]                  = gps
    return fd


def _load_font(size, bold=False, family=None):
    """Load a PIL font. family can be a system font stem name or None for Arial."""
    sys_fonts = get_system_fonts()

    if family and family != "Arial (Default)":
        # 1. Exact name in scanned system fonts dict
        if family in sys_fonts:
            try: return ImageFont.truetype(sys_fonts[family], size)
            except (IOError, OSError): pass
        # 2. Bold variant suffixes
        if bold:
            for suffix in ("bd", " Bold", "B", "-Bold"):
                candidate = family + suffix
                if candidate in sys_fonts:
                    try: return ImageFont.truetype(sys_fonts[candidate], size)
                    except (IOError, OSError): pass
        # 3. Try treating family as a direct path / PIL-resolvable name
        try: return ImageFont.truetype(family, size)
        except (IOError, OSError): pass

    # Default: Arial / system fallback
    if bold:
        candidates = ["arialbd.ttf", "Arial Bold.ttf", "arial.ttf", "Arial.ttf"]
    else:
        candidates = ["arial.ttf", "Arial.ttf", "arialbd.ttf", "consola.ttf", "cour.ttf"]
    for n in candidates:
        try: return ImageFont.truetype(n, size)
        except (IOError, OSError): pass
    return ImageFont.load_default()


# ─────────────────────────────────────────────────────────────────────────────
#  Core overlay
# ─────────────────────────────────────────────────────────────────────────────

def _max_font_for_single_line(parts, max_w, requested,
                               sep="   |   ", min_size=12, max_lines=3,
                               font_family=None):
    _draw = ImageDraw.Draw(Image.new("RGB", (1, 1)))

    def fits(fs):
        f   = _load_font(fs, family=font_family)
        rem = list(parts)
        n   = 0
        while rem:
            cur = []
            while rem:
                if _draw.textbbox((0, 0), sep.join(cur + [rem[0]]), font=f)[2] <= max_w:
                    cur.append(rem.pop(0))
                else:
                    if not cur:
                        return False
                    break
            n += 1
            if n > max_lines:
                return False
        return True

    if fits(requested):
        return requested
    lo, hi = min_size, requested - 1
    while lo < hi:
        mid = (lo + hi + 1) // 2
        if fits(mid): lo = mid
        else:         hi = mid - 1
    return lo


def apply_overlay(img, fields, font_size, color_name, corner,
                  layout="Multi-line", show_labels=True, notes="",
                  font_family=None,
                  padding=48, line_spacing=10):
    note_lines = [l for l in notes.strip().split("\n")[:2] if l.strip()]
    if not fields and not note_lines:
        return img.copy()

    out   = img.copy().convert("RGB")
    font  = _load_font(font_size)                          # EXIF data — always Arial
    color = TEXT_COLORS.get(color_name, (255, 255, 255))
    draw  = ImageDraw.Draw(out)
    dummy = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    w, h  = out.size
    lh    = dummy.textbbox((0, 0), "Ag", font=font)[3]

    # ── Single-line mode ──────────────────────────────────────────────────────
    if layout == "Single-line":
        SEP    = "   |   "
        max_w  = w - 2 * padding
        font_b = _load_font(font_size, bold=True, family=font_family)
        lh_b   = dummy.textbbox((0, 0), "Ag", font=font_b)[3]

        if fields:
            if show_labels:
                parts = [f"{k} : {v}" for k, v in fields.items()]
            else:
                parts = list(fields.values())

            # Caption mode: bar height is dynamic so no need to shrink the font —
            # just let text wrap to as many lines as required.
            if corner != "Caption":
                capped = _max_font_for_single_line(parts, max_w, font_size,
                                                   sep=SEP, font_family=None)
                if capped != font_size:
                    font   = _load_font(capped)                            # EXIF — always Arial
                    font_b = _load_font(capped, bold=True, family=font_family)
                    lh     = dummy.textbbox((0, 0), "Ag", font=font)[3]
                    lh_b   = dummy.textbbox((0, 0), "Ag", font=font_b)[3]

            full_tw = dummy.textbbox((0, 0), SEP.join(parts), font=font)[2]
            if full_tw <= max_w or len(parts) <= 1:
                exif_lines = [SEP.join(parts)]
            else:
                exif_lines = []
                remaining  = list(parts)
                # Caption mode: unlimited lines — bar grows to fit.
                # Other modes: cap at 3 lines to stay within image bounds.
                max_wrap = None if corner == "Caption" else 3
                while remaining:
                    current = []
                    while remaining:
                        candidate = SEP.join(current + [remaining[0]])
                        if dummy.textbbox((0, 0), candidate, font=font)[2] <= max_w:
                            current.append(remaining.pop(0))
                        else:
                            if not current:
                                current.append(remaining.pop(0))
                            break
                    # On the last allowed line (non-Caption), dump remaining fields
                    if max_wrap and len(exif_lines) == max_wrap - 1 and remaining:
                        current.extend(remaining)
                        remaining = []
                    exif_lines.append(SEP.join(current))
        else:
            exif_lines = []

        # ── Caption: black bar added below image ──────────────────────────────
        if corner == "Caption":
            all_cap = note_lines + exif_lines
            if not all_cap:
                return out

            note_h = sum(lh_b + line_spacing for _ in note_lines)
            exif_h = sum(lh  + line_spacing for _ in exif_lines)
            bar_h  = note_h + exif_h - line_spacing + 2 * padding

            cap_img = Image.new("RGB", (w, h + bar_h), color=(0, 0, 0))
            cap_img.paste(out, (0, 0))
            draw_c  = ImageDraw.Draw(cap_img)

            y = h + padding
            for i, line_text in enumerate(all_cap):
                fnt    = font_b if i < len(note_lines) else font
                cur_lh = lh_b   if i < len(note_lines) else lh
                tw     = dummy.textbbox((0, 0), line_text, font=fnt)[2]
                tx     = max(padding, (w - tw) // 2)
                draw_c.text((tx, y), line_text, font=fnt, fill=(255, 255, 255))
                y     += cur_lh + line_spacing

            return cap_img

        # ── Top / Bottom overlay ──────────────────────────────────────────────
        else:
            all_lines  = note_lines + exif_lines
            note_total = sum(lh_b + line_spacing for _ in note_lines)
            exif_total = sum(lh  + line_spacing for _ in exif_lines)
            total_h    = (note_total + exif_total - line_spacing) if all_lines else 0

            ty = padding if "Top" in corner else h - total_h - padding
            y  = ty
            for i, line_text in enumerate(all_lines):
                fnt    = font_b if i < len(note_lines) else font
                cur_lh = lh_b   if i < len(note_lines) else lh
                tw     = dummy.textbbox((0, 0), line_text, font=fnt)[2]
                tx     = max(padding, (w - tw) // 2)
                draw.text((tx, y), line_text, font=fnt, fill=color)
                y     += cur_lh + line_spacing

    # ── Multi-line mode ───────────────────────────────────────────────────────
    else:
        GAP    = max(6, font_size // 4)
        lbls   = list(fields.keys())
        vals   = list(fields.values())
        font_b = _load_font(font_size, bold=True, family=font_family)
        lh_b   = dummy.textbbox((0, 0), "Ag", font=font_b)[3]

        if show_labels and lbls:
            mlw = max(dummy.textbbox((0, 0), k, font=font)[2] for k in lbls)
            cw  = dummy.textbbox((0, 0), ":", font=font)[2]
            mvw = max(dummy.textbbox((0, 0), v, font=font)[2] for v in vals) if vals else 0
            bw  = mlw + GAP + cw + GAP + mvw
        else:
            mlw = cw = 0
            mvw = max(dummy.textbbox((0, 0), v, font=font)[2] for v in vals) if vals else 0
            bw  = mvw

        if note_lines:
            nw = max(dummy.textbbox((0, 0), nl, font=font_b)[2] for nl in note_lines)
            bw = max(bw, nw)

        note_h = len(note_lines) * (lh_b + line_spacing)
        exif_h = len(lbls)       * (lh   + line_spacing)
        gap_h  = line_spacing if note_lines and lbls else 0
        bh     = (note_h + gap_h + exif_h - line_spacing) if (note_lines or lbls) else 0

        if   corner == "Top Right":    tx, ty = w - bw - padding, padding
        elif corner == "Top Left":     tx, ty = padding,           padding
        elif corner == "Bottom Right": tx, ty = w - bw - padding,  h - bh - padding
        else:                          tx, ty = padding,            h - bh - padding

        y = ty

        for nl in note_lines:
            draw.text((tx, y), nl, font=font_b, fill=color)
            y += lh_b + line_spacing

        if note_lines and lbls:
            y += line_spacing

        if show_labels and lbls:
            cx = tx + mlw + GAP
            vx = cx + cw  + GAP
            for lb, vl in zip(lbls, vals):
                draw.text((tx, y), lb,  font=font, fill=color)
                draw.text((cx, y), ":", font=font, fill=color)
                draw.text((vx, y), vl,  font=font, fill=color)
                y += lh + line_spacing
        else:
            for vl in vals:
                draw.text((tx, y), vl, font=font, fill=color)
                y += lh + line_spacing

    return out


def burn_exif(src_path, active_fields, font_size, color_name, corner,
              layout="Multi-line", show_labels=True, notes="",
              font_family=None, dest_folder=None):
    p    = Path(src_path).resolve()
    img  = Image.open(p)
    orig_exif = img.info.get("exif", b"")
    img  = img.convert("RGB")
    all_f = read_exif(img)
    f    = {k: v for k, v in all_f.items() if k in active_fields}

    # Normalise font size by longest edge so portrait vs landscape frames
    # from the same camera get visually identical text weight.
    long_edge = max(img.width, img.height)
    adj_font  = max(8, round(font_size * long_edge / FONT_REF_LONG_EDGE))

    res  = apply_overlay(img, f, adj_font, color_name, corner,
                         layout=layout, show_labels=show_labels, notes=notes,
                         font_family=font_family)
    dest = Path(dest_folder) if dest_folder else p.parent
    out  = dest / (p.stem + "_exif.jpg")
    save_kw = {"quality": 95}
    if orig_exif and res.size == img.size:
        save_kw["exif"] = orig_exif
    res.save(out, "JPEG", **save_kw)
    return str(out)


def collect_jpgs(folder):
    exts = {".jpg", ".jpeg"}
    return sorted(str(f) for f in Path(folder).iterdir()
                  if f.suffix.lower() in exts and not f.stem.endswith("_exif"))


# ─────────────────────────────────────────────────────────────────────────────
#  Folder picker popup
# ─────────────────────────────────────────────────────────────────────────────

class FolderPicker(ctk.CTkToplevel):
    def __init__(self, parent, paths, on_confirm):
        super().__init__(parent)
        self.title("Select Photos to Process")
        self.geometry("500x600")
        self.resizable(False, False)
        self.configure(fg_color=C_BG)
        self.grab_set(); self.lift(); self.focus_force()
        self._paths, self._on_confirm = paths, on_confirm
        self._vars = {p: ctk.BooleanVar(value=True) for p in paths}
        self._build()

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color=C_HEADER, corner_radius=0, height=54)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text=f"📁  {Path(self._paths[0]).parent.name}",
                     font=ctk.CTkFont("Arial", 14, "bold"),
                     text_color=C_HDR_TEXT, fg_color="transparent"
                     ).place(relx=0.5, rely=0.5, anchor="center")
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=14, pady=(10, 4))
        self._count_lbl = ctk.CTkLabel(top, text="",
                                        font=ctk.CTkFont("Arial", 12),
                                        text_color=C_LABEL, fg_color="transparent")
        self._count_lbl.pack(side="left")
        ctk.CTkButton(top, text="None", width=64, height=28,
                      font=ctk.CTkFont("Arial", 11), fg_color=C_SEP,
                      hover_color=C_CARD, text_color=C_TEXT, corner_radius=6,
                      command=lambda: self._toggle(False)).pack(side="right", padx=(4,0))
        ctk.CTkButton(top, text="All", width=64, height=28,
                      font=ctk.CTkFont("Arial", 11), fg_color=C_ACCENT,
                      hover_color=C_HOVER, text_color="#ffffff", corner_radius=6,
                      command=lambda: self._toggle(True)).pack(side="right")
        ctk.CTkFrame(self, height=1, fg_color=C_SEP).pack(fill="x", padx=14)
        scroll = ctk.CTkScrollableFrame(self, fg_color=C_CARD, corner_radius=8,
                                         scrollbar_button_color=C_ACCENT,
                                         scrollbar_button_hover_color=C_HOVER)
        scroll.pack(fill="both", expand=True, padx=14, pady=8)
        for p in self._paths:
            size_kb = Path(p).stat().st_size // 1024
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkCheckBox(row, text=Path(p).name, variable=self._vars[p],
                             font=ctk.CTkFont("Arial", 12), text_color=C_TEXT,
                             fg_color=C_ACCENT, hover_color=C_HOVER,
                             checkmark_color="#ffffff", checkbox_width=18,
                             checkbox_height=18, corner_radius=4,
                             command=self._update_count).pack(side="left", anchor="w", padx=(6,0))
            ctk.CTkLabel(row, text=f"{size_kb} KB",
                         font=ctk.CTkFont("Arial", 10),
                         text_color=C_DIM, fg_color="transparent").pack(side="right", padx=10)
        self._update_count()
        ctk.CTkFrame(self, height=1, fg_color=C_SEP).pack(fill="x", padx=14)
        ctk.CTkButton(self, text="Queue Selected  →",
                      font=ctk.CTkFont("Arial", 13, "bold"), height=44,
                      corner_radius=0, fg_color=C_ACCENT, hover_color=C_HOVER,
                      text_color="#ffffff", command=self._confirm).pack(fill="x")

    def _update_count(self):
        n = sum(v.get() for v in self._vars.values())
        self._count_lbl.configure(text=f"{n} of {len(self._vars)} photos selected")

    def _toggle(self, state):
        for v in self._vars.values(): v.set(state)
        self._update_count()

    def _confirm(self):
        sel = [p for p, v in self._vars.items() if v.get()]
        if not sel:
            messagebox.showwarning("Nothing selected",
                                   "Select at least one photo.", parent=self)
            return
        self.destroy(); self._on_confirm(sel)


# ─────────────────────────────────────────────────────────────────────────────
#  Main App
# ─────────────────────────────────────────────────────────────────────────────

class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("EXIF Overlay Tool  v1.4")
        self.configure(fg_color=C_BG)
        self.resizable(True, True)
        self.minsize(1000, 680)

        # ── state ─────────────────────────────────────────────────────────────
        self._files          = []
        self._thumb_ref      = None
        self._logo_ref       = None
        self._render_base    = None
        self._preview_exif   = {}
        self._preview_timer  = None
        self._orig_size      = (1, 1)
        self._current_idx    = 0
        self._photo_settings = {}
        self._dest_folder    = None

        self._field_vars      = {f: ctk.BooleanVar(value=True) for f in EXIF_FIELD_LABELS}
        self._layout_var      = ctk.StringVar(value="Multi-line")
        self._strip_pos_var   = ctk.StringVar(value="Bottom")
        self._show_labels_var = ctk.BooleanVar(value=True)
        self._font_family_var = ctk.StringVar(value="Arial (Default)")

        self._build_ui()
        self._apply_config()           # restore saved defaults before first use

        # traces that trigger live preview
        for var in self._field_vars.values():
            var.trace_add("write", self._schedule_preview)
        self._corner_var.trace_add("write", self._schedule_preview)
        self._layout_var.trace_add("write", self._on_layout_change)
        self._strip_pos_var.trace_add("write", self._schedule_preview)
        self._show_labels_var.trace_add("write", self._schedule_preview)
        self._font_family_var.trace_add("write", self._schedule_preview)

        # Save config 1 second after any global setting changes
        for var in (self._corner_var, self._layout_var, self._strip_pos_var,
                    self._show_labels_var, self._font_family_var, self._font_val,
                    self._color_var):
            var.trace_add("write", self._schedule_config_save)

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(0, lambda: self.state("zoomed"))

    # ══════════════════════════════════════════════════════════════════════════
    #  UI BUILD
    # ══════════════════════════════════════════════════════════════════════════

    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, corner_radius=0, fg_color=C_HEADER, height=80)
        hdr.pack(fill="x"); hdr.pack_propagate(False)

        if LOGO_FILE.exists():
            raw = Image.open(LOGO_FILE).convert("RGBA")
            lh = 56; lw = int(raw.width * lh / raw.height)
            raw = raw.resize((lw, lh), Image.LANCZOS)
            self._logo_ref = ctk.CTkImage(light_image=raw, dark_image=raw, size=(lw, lh))
            ctk.CTkLabel(hdr, image=self._logo_ref, text="",
                         fg_color="transparent").place(x=16, rely=0.5, anchor="w")
            tx = lw + 28
        else:
            tx = 16

        ctk.CTkLabel(hdr, text="EXIF Overlay Tool",
                     font=ctk.CTkFont("Arial", 20, "bold"),
                     text_color=C_HDR_TEXT, fg_color="transparent"
                     ).place(x=tx, rely=0.30, anchor="w")
        ctk.CTkLabel(hdr, text="v 1.4",
                     font=ctk.CTkFont("Arial", 11),
                     text_color=C_HDR_SUB, fg_color="transparent"
                     ).place(x=tx + 178, rely=0.30, anchor="w")
        ctk.CTkLabel(hdr, text="Stamp camera data onto your photos",
                     font=ctk.CTkFont("Arial", 11),
                     text_color=C_HDR_SUB, fg_color="transparent"
                     ).place(x=tx, rely=0.68, anchor="w")

        # ── 3-column body ─────────────────────────────────────────────────────
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=14, pady=10)

        left = ctk.CTkFrame(body, corner_radius=10, fg_color=C_CARD, width=220)
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)
        self._build_left_panel(left)

        right = ctk.CTkFrame(body, corner_radius=10, fg_color=C_CARD, width=265)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)
        self._build_right_panel(right)

        centre = ctk.CTkFrame(body, fg_color="transparent")
        centre.pack(side="left", fill="both", expand=True)
        self._build_centre_panel(centre)

        # ── Action buttons ────────────────────────────────────────────────────
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=14, pady=(0, 4))
        dim = dict(font=ctk.CTkFont("Arial", 13, "bold"), height=44,
                   corner_radius=10, fg_color=C_CARD,
                   hover_color="#c8c0a8", text_color=C_TEXT)

        self._btn_files = ctk.CTkButton(btn_frame, text="📂  Select Photo(s)",
                                         **dim, command=self._select_files)
        self._btn_files.pack(side="left", expand=True, fill="x", padx=(0, 5))

        self._btn_folder = ctk.CTkButton(btn_frame, text="📁  Select Folder",
                                          **dim, command=self._select_folder)
        self._btn_folder.pack(side="left", expand=True, fill="x", padx=(0, 5))

        self._btn_process = ctk.CTkButton(btn_frame, text="✨  Process",
                                           font=ctk.CTkFont("Arial", 13, "bold"),
                                           height=44, corner_radius=10,
                                           fg_color=C_ACCENT, hover_color=C_HOVER,
                                           text_color="#ffffff",
                                           text_color_disabled="#aaaaaa",
                                           state="disabled", command=self._run)
        self._btn_process.pack(side="left", expand=True, fill="x")

        self._save_info = ctk.CTkLabel(self, text="",
                     font=ctk.CTkFont("Arial", 11),
                     text_color=C_LABEL, fg_color="transparent", anchor="w")
        self._save_info.pack(fill="x", padx=18, pady=(4, 0))

        self._status = ctk.CTkLabel(self,
                     text="Select photo(s) or a folder to begin.",
                     font=ctk.CTkFont("Arial", 11),
                     text_color=C_DIM, fg_color="transparent", anchor="w")
        self._status.pack(fill="x", padx=18, pady=(1, 0))

        ctk.CTkLabel(self,
                     text="Created by Ashit Gandhi   •   Ver 1.4   •   June 2026",
                     font=ctk.CTkFont("Arial", 10),
                     text_color=C_DIM, fg_color="transparent").pack(pady=(2, 8))

    # ── Left panel ────────────────────────────────────────────────────────────

    def _build_left_panel(self, parent):
        ctk.CTkLabel(parent, text="EXIF FIELDS",
                     font=ctk.CTkFont("Arial", 11, "bold"),
                     text_color=C_LABEL, fg_color="transparent"
                     ).pack(pady=(12, 4), padx=12, anchor="w")

        mini = ctk.CTkFrame(parent, fg_color="transparent")
        mini.pack(fill="x", padx=10, pady=(0, 6))
        ctk.CTkButton(mini, text="All", width=60, height=24,
                      font=ctk.CTkFont("Arial", 11), fg_color=C_ACCENT,
                      hover_color=C_HOVER, text_color="#ffffff", corner_radius=6,
                      command=lambda: self._toggle_all(True)).pack(side="left", padx=(0, 4))
        ctk.CTkButton(mini, text="None", width=60, height=24,
                      font=ctk.CTkFont("Arial", 11), fg_color=C_SEP,
                      hover_color="#c8c0a8", text_color=C_TEXT, corner_radius=6,
                      command=lambda: self._toggle_all(False)).pack(side="left")

        ctk.CTkFrame(parent, height=1, fg_color=C_SEP).pack(fill="x", padx=10, pady=(0, 4))

        original_fields = EXIF_FIELD_LABELS[:8]
        new_fields      = EXIF_FIELD_LABELS[8:]

        for field in original_fields:
            ctk.CTkCheckBox(parent, text=field,
                             variable=self._field_vars[field],
                             font=ctk.CTkFont("Arial", 12), text_color=C_TEXT,
                             fg_color=C_ACCENT, hover_color=C_HOVER,
                             checkmark_color="#ffffff",
                             checkbox_width=18, checkbox_height=18, corner_radius=4
                             ).pack(anchor="w", padx=12, pady=2)

        ctk.CTkFrame(parent, height=1, fg_color=C_SEP).pack(fill="x", padx=10, pady=6)
        ctk.CTkLabel(parent, text="ADDITIONAL",
                     font=ctk.CTkFont("Arial", 9, "bold"),
                     text_color=C_DIM, fg_color="transparent"
                     ).pack(anchor="w", padx=12, pady=(0, 2))

        for field in new_fields:
            ctk.CTkCheckBox(parent, text=field,
                             variable=self._field_vars[field],
                             font=ctk.CTkFont("Arial", 12), text_color=C_TEXT,
                             fg_color=C_ACCENT, hover_color=C_HOVER,
                             checkmark_color="#ffffff",
                             checkbox_width=18, checkbox_height=18, corner_radius=4
                             ).pack(anchor="w", padx=12, pady=2)

    # ── Centre panel ──────────────────────────────────────────────────────────

    def _build_centre_panel(self, parent):
        self._thumb_frame = ctk.CTkFrame(parent, corner_radius=10, fg_color=C_CARD)
        self._thumb_frame.pack(fill="both", expand=True)

        self._thumb_lbl = ctk.CTkLabel(self._thumb_frame,
                                        text="No photo selected",
                                        font=ctk.CTkFont("Arial", 13),
                                        text_color=C_DIM, fg_color="transparent")
        self._thumb_lbl.place(relx=0.5, rely=0.5, anchor="center")

        self._preview_badge = ctk.CTkLabel(self._thumb_frame,
                                            text=" LIVE PREVIEW ",
                                            font=ctk.CTkFont("Arial", 9, "bold"),
                                            text_color=C_HDR_TEXT,
                                            fg_color=C_ACCENT, corner_radius=4)

        self._btn_close_sel = ctk.CTkButton(self._thumb_frame, text="✕",
                                             font=ctk.CTkFont("Arial", 11, "bold"),
                                             width=28, height=28, corner_radius=6,
                                             fg_color=C_HEADER, hover_color="#5a5040",
                                             text_color=C_HDR_TEXT,
                                             command=self._close_selection)

        self._nav_frame = ctk.CTkFrame(parent, fg_color=C_CARD,
                                        corner_radius=8, height=36)
        self._nav_frame.pack_propagate(False)

        self._btn_prev = ctk.CTkButton(self._nav_frame, text="◀  Prev",
                                        width=90, height=26,
                                        font=ctk.CTkFont("Arial", 11, "bold"),
                                        fg_color=C_ACCENT, hover_color=C_HOVER,
                                        text_color="#ffffff", corner_radius=6,
                                        command=self._nav_prev)
        self._btn_prev.place(x=6, rely=0.5, anchor="w")

        self._nav_label = ctk.CTkLabel(self._nav_frame, text="",
                                        font=ctk.CTkFont("Arial", 11, "bold"),
                                        text_color=C_LABEL, fg_color="transparent")
        self._nav_label.place(relx=0.5, rely=0.5, anchor="center")

        self._btn_next = ctk.CTkButton(self._nav_frame, text="Next  ▶",
                                        width=90, height=26,
                                        font=ctk.CTkFont("Arial", 11, "bold"),
                                        fg_color=C_ACCENT, hover_color=C_HOVER,
                                        text_color="#ffffff", corner_radius=6,
                                        command=self._nav_next)
        self._btn_next.place(relx=1.0, x=-6, rely=0.5, anchor="e")

        self._btn_apply_all = ctk.CTkButton(self._nav_frame, text="Apply to All",
                                             width=90, height=22,
                                             font=ctk.CTkFont("Arial", 10),
                                             fg_color=C_SEP, hover_color=C_CARD,
                                             text_color=C_TEXT, corner_radius=6,
                                             command=self._apply_to_all)
        self._btn_apply_all.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(parent, text="EXIF DATA",
                     font=ctk.CTkFont("Arial", 9, "bold"),
                     text_color=C_LABEL, fg_color="transparent"
                     ).pack(anchor="w", pady=(4, 1))

        self._exif_grid = ctk.CTkFrame(parent, fg_color=C_CARD, corner_radius=8)
        self._exif_grid.pack(fill="x")

        self._progress = ctk.CTkProgressBar(parent, height=8, corner_radius=4,
                                             fg_color=C_CARD, progress_color=C_PROG)
        self._progress.set(0); self._progress.pack_forget()

    # ── Right panel ───────────────────────────────────────────────────────────

    def _build_right_panel(self, parent):
        def sep(): ctk.CTkFrame(parent, height=1, fg_color=C_SEP).pack(fill="x", padx=12, pady=3)
        def lbl(t): ctk.CTkLabel(parent, text=t,
                     font=ctk.CTkFont("Arial", 10, "bold"),
                     text_color=C_LABEL, fg_color="transparent"
                     ).pack(padx=14, anchor="w")

        # ── Layout mode ───────────────────────────────────────────────────────
        ctk.CTkFrame(parent, height=6, fg_color="transparent").pack()
        lbl("LAYOUT MODE")
        self._layout_seg = ctk.CTkSegmentedButton(parent,
                               values=["Multi-line", "Single-line"],
                               variable=self._layout_var,
                               font=ctk.CTkFont("Arial", 11),
                               fg_color=C_SEP, selected_color=C_ACCENT,
                               selected_hover_color=C_HOVER,
                               unselected_color=C_SEP, unselected_hover_color=C_CARD,
                               text_color=C_TEXT, text_color_disabled=C_DIM,
                               command=self._on_layout_btn)
        self._layout_seg.pack(fill="x", padx=12, pady=(2, 0))
        self._on_layout_btn(self._layout_var.get())
        sep()

        # ── Position ──────────────────────────────────────────────────────────
        lbl("POSITION")
        self._corner_var = ctk.StringVar(value="Top Right")

        self._pos_multi_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self._pos_multi_frame.pack(fill="x", padx=4, pady=(2, 0))
        for c in CORNER_OPTIONS:
            ctk.CTkRadioButton(self._pos_multi_frame, text=c,
                               variable=self._corner_var, value=c,
                               font=ctk.CTkFont("Arial", 11), text_color=C_TEXT,
                               fg_color=C_ACCENT, hover_color=C_HOVER,
                               radiobutton_width=15, radiobutton_height=15
                               ).pack(anchor="w", padx=14, pady=0)

        STRIP_LABELS = {
            "Top":     "Top  (overlay)",
            "Bottom":  "Bottom  (overlay)",
            "Caption": "Bottom  (caption)",
        }
        self._pos_single_frame = ctk.CTkFrame(parent, fg_color="transparent")
        for s in STRIP_OPTIONS:
            ctk.CTkRadioButton(self._pos_single_frame, text=STRIP_LABELS[s],
                               variable=self._strip_pos_var, value=s,
                               font=ctk.CTkFont("Arial", 11), text_color=C_TEXT,
                               fg_color=C_ACCENT, hover_color=C_HOVER,
                               radiobutton_width=15, radiobutton_height=15
                               ).pack(anchor="w", padx=14, pady=1)
        sep()

        # ── Show labels toggle ────────────────────────────────────────────────
        lbl("TEXT FORMAT")
        lbl_row = ctk.CTkFrame(parent, fg_color="transparent")
        lbl_row.pack(fill="x", padx=12, pady=(2, 0))
        ctk.CTkLabel(lbl_row, text="Show field labels",
                     font=ctk.CTkFont("Arial", 12), text_color=C_TEXT,
                     fg_color="transparent").pack(side="left")
        ctk.CTkSwitch(lbl_row, text="", variable=self._show_labels_var,
                      width=44, fg_color=C_SEP, progress_color=C_ACCENT,
                      button_color=C_LABEL, button_hover_color=C_HOVER
                      ).pack(side="right")
        sep()

        # ── Font size ─────────────────────────────────────────────────────────
        lbl("FONT SIZE")
        srow = ctk.CTkFrame(parent, fg_color="transparent")
        srow.pack(fill="x", padx=12, pady=(2, 0))
        self._font_val = ctk.IntVar(value=40)
        self._font_lbl = ctk.CTkLabel(srow, text="40 pt",
                                       font=ctk.CTkFont("Arial", 12),
                                       text_color=C_TEXT, fg_color="transparent", width=44)
        self._font_lbl.pack(side="right")
        ctk.CTkSlider(srow, from_=30, to=120, variable=self._font_val,
                      fg_color=C_SEP, progress_color=C_ACCENT,
                      button_color=C_LABEL, button_hover_color="#a89060",
                      command=self._on_slider).pack(side="left", fill="x", expand=True)
        sep()

        # ── Notes font (regional language support) ────────────────────────────
        lbl("NOTES FONT")
        # Search entry
        self._font_search_var = ctk.StringVar()
        ctk.CTkEntry(parent,
                     textvariable=self._font_search_var,
                     placeholder_text="Search fonts…",
                     font=ctk.CTkFont("Arial", 11),
                     fg_color=C_BG, text_color=C_TEXT,
                     border_color=C_SEP, border_width=1,
                     height=28
                     ).pack(fill="x", padx=12, pady=(2, 2))
        self._font_search_var.trace_add("write", self._on_font_search)

        # Native Listbox — supports mouse-wheel scrolling and live filtering
        self._all_font_names = ["Arial (Default)"] + sorted(get_system_fonts().keys())
        lb_frame = _tk.Frame(parent, bg=C_CARD)
        lb_frame.pack(fill="x", padx=12, pady=(0, 0))
        sb = _tk.Scrollbar(lb_frame, orient="vertical")
        self._font_listbox = _tk.Listbox(
            lb_frame, height=4,
            yscrollcommand=sb.set,
            font=("Arial", 11),
            bg=C_BG, fg=C_TEXT,
            selectbackground=C_ACCENT, selectforeground="#ffffff",
            activestyle="none",
            borderwidth=0, highlightthickness=1,
            highlightbackground=C_SEP, highlightcolor=C_ACCENT,
            relief="flat")
        sb.config(command=self._font_listbox.yview)
        self._font_listbox.pack(side="left", fill="x", expand=True)
        sb.pack(side="right", fill="y")
        self._font_listbox.bind("<<ListboxSelect>>", self._on_font_listbox_select)
        self._font_listbox.bind("<MouseWheel>",
            lambda e: self._font_listbox.yview_scroll(-1*(e.delta//120), "units"))
        self._populate_font_listbox("")
        sep()

        # ── Text colour ───────────────────────────────────────────────────────
        lbl("TEXT COLOUR")
        self._color_var = ctk.StringVar(value="White")
        self._color_seg = ctk.CTkSegmentedButton(parent,
                               values=["White", "Yellow", "Black"],
                               variable=self._color_var,
                               font=ctk.CTkFont("Arial", 11),
                               fg_color=C_SEP, selected_color=C_ACCENT,
                               selected_hover_color=C_HOVER,
                               unselected_color=C_SEP, unselected_hover_color=C_CARD,
                               text_color=C_TEXT, text_color_disabled=C_DIM,
                               command=self._on_color_select)
        self._color_seg.pack(fill="x", padx=12, pady=(2, 0))
        self._on_color_select("White")
        sep()

        # ── Notes (per photo) ─────────────────────────────────────────────────
        lbl("NOTES  (per photo · max 2 lines)")
        self._notes_box = ctk.CTkTextbox(parent, height=44,
                                          font=ctk.CTkFont("Arial", 12),
                                          fg_color=C_BG, text_color=C_TEXT,
                                          corner_radius=6,
                                          scrollbar_button_color=C_CARD,
                                          scrollbar_button_hover_color=C_CARD)
        self._notes_box.pack(fill="x", padx=12, pady=(2, 0))
        self._notes_box.bind("<KeyRelease>", self._on_notes_change)
        sep()

        # ── Destination folder ────────────────────────────────────────────────
        lbl("DESTINATION FOLDER")
        ctk.CTkButton(parent, text="📁  Browse…", height=30,
                      font=ctk.CTkFont("Arial", 11),
                      fg_color=C_SEP, hover_color=C_CARD, text_color=C_TEXT,
                      corner_radius=6, command=self._pick_dest_folder
                      ).pack(fill="x", padx=12, pady=(2, 2))
        self._dest_lbl = ctk.CTkLabel(parent, text="Same folder as original",
                                       font=ctk.CTkFont("Arial", 10),
                                       text_color=C_DIM, fg_color="transparent",
                                       wraplength=220, anchor="w")
        self._dest_lbl.pack(fill="x", padx=14)

    # ══════════════════════════════════════════════════════════════════════════
    #  Config persistence
    # ══════════════════════════════════════════════════════════════════════════

    def _apply_config(self):
        try:
            cfg = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return
        if "font_size" in cfg:
            self._font_val.set(cfg["font_size"])
            self._font_lbl.configure(text=f"{cfg['font_size']} pt")
        if "color" in cfg:
            self._color_var.set(cfg["color"])
            self._on_color_select(cfg["color"])
        if "layout" in cfg:
            self._layout_var.set(cfg["layout"])
            self._on_layout_btn(cfg["layout"])
            self._on_layout_change()   # trace not yet wired at startup — swap frames manually
        if "corner" in cfg:
            self._corner_var.set(cfg["corner"])
        if "strip_pos" in cfg:
            self._strip_pos_var.set(cfg["strip_pos"])
        if "show_labels" in cfg:
            self._show_labels_var.set(cfg["show_labels"])
        if "font_family" in cfg:
            self._font_family_var.set(cfg["font_family"])
            self._font_selecting = True
            self._font_search_var.set(cfg["font_family"])
            self._font_selecting = False
            self._populate_font_listbox("")
        if "fields" in cfg:
            for f, v in cfg["fields"].items():
                if f in self._field_vars:
                    self._field_vars[f].set(v)

    def _save_config(self):
        cfg = {
            "font_size":   int(self._font_val.get()),
            "color":       self._color_var.get(),
            "layout":      self._layout_var.get(),
            "corner":      self._corner_var.get(),
            "strip_pos":   self._strip_pos_var.get(),
            "show_labels": self._show_labels_var.get(),
            "font_family": self._font_family_var.get(),
            "fields":      {f: v.get() for f, v in self._field_vars.items()},
        }
        try:
            CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        except Exception:
            pass

    _config_timer = None

    def _schedule_config_save(self, *_):
        if self._config_timer is not None:
            self.after_cancel(self._config_timer)
        self._config_timer = self.after(1000, self._save_config)

    def _on_close(self):
        self._save_config()
        self.destroy()

    # ══════════════════════════════════════════════════════════════════════════
    #  Event handlers
    # ══════════════════════════════════════════════════════════════════════════

    def _on_layout_btn(self, selected):
        for name, btn in self._layout_seg._buttons_dict.items():
            btn.configure(text_color="#ffffff" if name == selected else C_TEXT)

    def _on_layout_change(self, *_):
        if not hasattr(self, "_pos_single_frame"):
            return
        if self._layout_var.get() == "Single-line":
            self._pos_multi_frame.pack_forget()
            self._pos_single_frame.pack(fill="x", padx=4, pady=(4, 0),
                                         after=self._layout_seg)
        else:
            self._pos_single_frame.pack_forget()
            self._pos_multi_frame.pack(fill="x", padx=4, pady=(4, 0),
                                        after=self._layout_seg)
        self._schedule_preview()

    def _on_slider(self, val):
        self._font_lbl.configure(text=f"{int(val)} pt")
        self._schedule_preview()

    def _on_color_select(self, selected):
        for name, btn in self._color_seg._buttons_dict.items():
            btn.configure(text_color="#ffffff" if name == selected else C_TEXT)
        self._schedule_preview()

    def _populate_font_listbox(self, query=""):
        self._font_listbox.delete(0, "end")
        names = (self._all_font_names if not query
                 else ["Arial (Default)"] + [
                     n for n in self._all_font_names[1:] if query in n.lower()])
        for name in names:
            self._font_listbox.insert("end", name)
        current = self._font_family_var.get()
        for i, name in enumerate(names):
            if name == current:
                self._font_listbox.selection_clear(0, "end")
                self._font_listbox.selection_set(i)
                self._font_listbox.see(i)
                break

    _font_selecting = False   # guard against search↔select feedback loop

    def _on_font_search(self, *_):
        if self._font_selecting:
            return
        self._populate_font_listbox(self._font_search_var.get().lower().strip())

    def _on_font_listbox_select(self, event):
        sel = self._font_listbox.curselection()
        if not sel:
            return
        name = self._font_listbox.get(sel[0])
        self._font_family_var.set(name)
        # Show selected name in the search box without re-filtering the list
        self._font_selecting = True
        self._font_search_var.set(name)
        self._font_selecting = False
        self._schedule_preview()

    def _on_notes_change(self, *_):
        content = self._notes_box.get("1.0", "end-1c")
        lines = content.split("\n")
        if len(lines) > 2:
            self._notes_box.delete("1.0", "end")
            self._notes_box.insert("1.0", "\n".join(lines[:2]))
        self._schedule_preview()

    def _pick_dest_folder(self):
        folder = filedialog.askdirectory(title="Select destination folder")
        if not folder:
            return
        self._dest_folder = folder
        display = folder if len(folder) <= 34 else "…" + folder[-32:]
        self._dest_lbl.configure(text=display, text_color=C_LABEL)
        self._update_save_info()

    def _toggle_all(self, state):
        for v in self._field_vars.values(): v.set(state)

    # ══════════════════════════════════════════════════════════════════════════
    #  File selection
    # ══════════════════════════════════════════════════════════════════════════

    def _select_files(self):
        paths = filedialog.askopenfilenames(
            title="Select JPG photo(s)",
            filetypes=[("JPEG images", "*.jpg *.jpeg *.JPG *.JPEG"), ("All files", "*.*")]
        )
        if paths: self._load(list(paths))

    def _select_folder(self):
        folder = filedialog.askdirectory(title="Select folder with JPG photos")
        if not folder: return
        paths = collect_jpgs(folder)
        if not paths:
            messagebox.showinfo("No photos", "No JPG files found in that folder.")
            return
        FolderPicker(self, paths, on_confirm=self._load)

    def _load(self, paths):
        self._files = paths; self._current_idx = 0; self._photo_settings = {}
        if len(paths) > 1:
            self._nav_frame.pack(fill="x", pady=(6, 0))
        else:
            self._nav_frame.pack_forget()
        self._show_photo(0)
        self._btn_process.configure(state="normal")
        self._btn_close_sel.place(relx=1.0, x=-6, y=6, anchor="ne")

    def _close_selection(self):
        if self._preview_timer is not None:
            self.after_cancel(self._preview_timer); self._preview_timer = None
        self._files = []; self._current_idx = 0; self._photo_settings = {}
        self._render_base = None; self._preview_exif = {}
        self._nav_frame.pack_forget()
        self._thumb_lbl.place_forget()
        self._preview_badge.place_forget()
        self._notes_box.delete("1.0", "end")
        for w in self._exif_grid.winfo_children(): w.destroy()
        self._btn_process.configure(state="disabled")
        self._btn_close_sel.place_forget()
        self._save_info.configure(text="")
        self._status.configure(text="Select photo(s) or a folder to begin.")

    # ══════════════════════════════════════════════════════════════════════════
    #  Navigation
    # ══════════════════════════════════════════════════════════════════════════

    def _show_photo(self, idx):
        self._current_idx = idx
        path  = self._files[idx]
        total = len(self._files)
        if path in self._photo_settings:
            s = self._photo_settings[path]
            for f, v in self._field_vars.items(): v.set(s["fields"].get(f, True))
            self._font_val.set(s["font_size"])
            self._font_lbl.configure(text=f"{s['font_size']} pt")
            self._color_var.set(s["color"])
            self._on_color_select(s["color"])
            self._corner_var.set(s["corner"])
            fam = s.get("font_family", "Arial (Default)")
            self._font_family_var.set(fam)
            self._font_selecting = True
            self._font_search_var.set(fam)
            self._font_selecting = False
            self._populate_font_listbox("")
            # Per-photo notes: restore each photo's own label text
            self._notes_box.delete("1.0", "end")
            self._notes_box.insert("1.0", s.get("notes", ""))
        self._nav_label.configure(text=f"Photo {idx+1} of {total}")
        self._btn_prev.configure(state="normal" if idx > 0       else "disabled")
        self._btn_next.configure(state="normal" if idx < total-1 else "disabled")
        self._load_render_base(path)
        self._show_exif_panel(path)
        self._status.configure(text=f"Photo {idx+1} of {total}  —  {Path(path).name}")
        self._update_save_info()

    def _update_save_info(self):
        if not self._files: return
        stem = Path(self._files[self._current_idx]).stem
        if self._dest_folder:
            dest_str = str(Path(self._dest_folder) / f"{stem}_exif.jpg")
        else:
            dest_str = f"same folder as original  →  {stem}_exif.jpg"
        self._save_info.configure(text=f"📁  Saved as:  {dest_str}")

    def _save_current_settings(self):
        if not self._files: return
        path = self._files[self._current_idx]
        self._photo_settings[path] = {
            "fields":      {f: v.get() for f, v in self._field_vars.items()},
            "font_size":   int(self._font_val.get()),
            "color":       self._color_var.get(),
            "corner":      self._corner_var.get(),
            "font_family": self._font_family_var.get(),
            "notes":       self._notes_box.get("1.0", "end-1c"),
        }

    def _nav_prev(self):
        self._save_current_settings(); self._show_photo(self._current_idx - 1)

    def _nav_next(self):
        self._save_current_settings(); self._show_photo(self._current_idx + 1)

    def _apply_to_all(self):
        """Copy EXIF fields, font size/family, colour, and position to all photos.
        Notes are intentionally excluded — species/location labels differ per photo."""
        self._save_current_settings()
        base = self._photo_settings.get(self._files[self._current_idx])
        if not base: return
        shared_keys = ("fields", "font_size", "color", "corner", "font_family")
        shared = {k: base[k] for k in shared_keys if k in base}
        for p in self._files:
            existing_notes = self._photo_settings.get(p, {}).get("notes", "")
            self._photo_settings[p] = {**shared, "notes": existing_notes}
        self._status.configure(
            text="✓  Settings applied to all photos  (notes kept per-photo).")

    # ══════════════════════════════════════════════════════════════════════════
    #  Live preview
    # ══════════════════════════════════════════════════════════════════════════

    def _load_render_base(self, path):
        try:
            img = Image.open(path).convert("RGB")
            self._preview_exif = read_exif(img)
            self._orig_size = (img.width, img.height)
            rb = img.copy()
            rb.thumbnail((RENDER_MAX, RENDER_MAX), Image.LANCZOS)
            self._render_base = rb
            self._render_preview()
        except Exception:
            self._render_base = None; self._preview_exif = {}
            self._thumb_lbl.configure(image=None, text="(preview unavailable)")

    def _schedule_preview(self, *_):
        if self._preview_timer is not None:
            self.after_cancel(self._preview_timer)
        self._preview_timer = self.after(180, self._render_preview)

    def _render_preview(self):
        if self._render_base is None: return
        active = [f for f, v in self._field_vars.items() if v.get()]
        fields = {k: v for k, v in self._preview_exif.items() if k in active}

        # thumbnail → original scale factor
        scale = self._render_base.width / self._orig_size[0]
        # apply same longest-edge normalisation as burn_exif so preview matches output
        long_edge  = max(self._orig_size)
        ref_scale  = long_edge / FONT_REF_LONG_EDGE
        prev_font  = max(6, round(int(self._font_val.get()) * ref_scale * scale))
        prev_pad   = max(4, round(48 * scale))
        prev_spc   = max(2, round(10 * scale))

        layout      = self._layout_var.get()
        corner      = (self._corner_var.get() if layout == "Multi-line"
                       else self._strip_pos_var.get())
        notes       = self._notes_box.get("1.0", "end-1c")
        show_labels = self._show_labels_var.get()
        font_family = self._font_family_var.get()

        result = apply_overlay(self._render_base, fields, prev_font,
                               self._color_var.get(), corner,
                               layout=layout, show_labels=show_labels,
                               notes=notes, font_family=font_family,
                               padding=prev_pad, line_spacing=prev_spc)
        display = result.copy()
        fw = max(400, self._thumb_frame.winfo_width()  - 20)
        fh = max(300, self._thumb_frame.winfo_height() - 20)
        display.thumbnail((fw, fh), Image.LANCZOS)
        self._thumb_ref = ctk.CTkImage(light_image=display, dark_image=display,
                                        size=display.size)
        self._thumb_lbl.configure(image=self._thumb_ref, text="")
        self._thumb_lbl.place(relx=0.5, rely=0.5, anchor="center")
        self._preview_badge.place(x=6, y=6)

    # ══════════════════════════════════════════════════════════════════════════
    #  EXIF data panel
    # ══════════════════════════════════════════════════════════════════════════

    def _show_exif_panel(self, path):
        for w in self._exif_grid.winfo_children():
            w.destroy()

        try:    fields = read_exif(Image.open(path))
        except: fields = {}

        if not fields:
            _tk.Label(self._exif_grid, text="  No EXIF data found in this file.",
                      font=("Consolas", 10), fg=C_DIM, bg=C_CARD, pady=4
                      ).pack(anchor="w", padx=8)
            return

        items  = list(fields.items())
        n      = len(items)
        rows_n = (n + 2) // 3
        cols   = [items[i*rows_n:(i+1)*rows_n] for i in range(3)]

        wrapper = _tk.Frame(self._exif_grid, bg=C_CARD)
        wrapper.pack(fill="x", padx=6, pady=(4, 4))

        def build_col(parent, entries):
            for k, v in entries:
                active     = self._field_vars.get(k, ctk.BooleanVar(value=False)).get()
                tick_color = C_ACCENT if active else C_DIM
                tick       = "+" if active else "-"
                row = _tk.Frame(parent, bg=C_CARD)
                row.pack(fill="x")
                _tk.Label(row, text=tick, font=("Consolas", 10, "bold"),
                          fg=tick_color, bg=C_CARD, pady=0, padx=0, width=2
                          ).pack(side="left", padx=(4, 2))
                _tk.Label(row, text=f"{k:<14}", font=("Consolas", 10),
                          fg=C_LABEL, bg=C_CARD, pady=0
                          ).pack(side="left")
                _tk.Label(row, text=v, font=("Consolas", 10),
                          fg=C_TEXT, bg=C_CARD, pady=0
                          ).pack(side="left", padx=(2, 8))

        for ci, col_items in enumerate(cols):
            if not col_items:
                continue
            col_frame = _tk.Frame(wrapper, bg=C_CARD)
            col_frame.pack(side="left", fill="both", expand=True)
            build_col(col_frame, col_items)
            if ci < 2 and any(cols[ci+1:]):
                _tk.Frame(wrapper, width=1, bg=C_SEP
                          ).pack(side="left", fill="y", padx=3)

    # ══════════════════════════════════════════════════════════════════════════
    #  Processing
    # ══════════════════════════════════════════════════════════════════════════

    def _run(self):
        self._save_current_settings()
        default_active  = [f for f, v in self._field_vars.items() if v.get()]
        default_font    = int(self._font_val.get())
        default_color   = self._color_var.get()
        default_corner  = self._corner_var.get()
        default_notes   = self._notes_box.get("1.0", "end-1c")
        default_family  = self._font_family_var.get()
        layout          = self._layout_var.get()
        show_labels     = self._show_labels_var.get()
        dest_folder     = self._dest_folder

        tasks = []
        for path in self._files:
            if path in self._photo_settings:
                s      = self._photo_settings[path]
                active = [f for f, on in s["fields"].items() if on]
                corner = (s["corner"] if layout == "Multi-line"
                          else self._strip_pos_var.get())
                tasks.append((path,
                               active or default_active,
                               s["font_size"],
                               s["color"],
                               corner,
                               s.get("font_family", default_family),
                               s.get("notes", "")))
            else:
                corner = (default_corner if layout == "Multi-line"
                          else self._strip_pos_var.get())
                tasks.append((path, default_active, default_font,
                               default_color, corner, default_family, default_notes))

        if not any(t[1] for t in tasks):
            messagebox.showwarning("No fields", "Tick at least one EXIF field.")
            return

        self._set_busy(True)
        self._progress.pack(fill="x", pady=(6, 0))
        self._progress.set(0)
        threading.Thread(
            target=self._worker,
            args=(tasks, layout, show_labels, dest_folder),
            daemon=True
        ).start()

    def _worker(self, tasks, layout, show_labels, dest_folder):
        results, errors = [], []
        total = len(tasks)
        for i, (path, active, font_sz, color, corner, font_family, notes) in enumerate(tasks, 1):
            self.after(0, self._status.configure,
                       {"text": f"Processing {i} / {total}  —  {Path(path).name}"})
            self.after(0, self._progress.set, i / total)
            try:
                results.append(
                    burn_exif(path, active, font_sz, color, corner,
                              layout=layout, show_labels=show_labels,
                              notes=notes, font_family=font_family,
                              dest_folder=dest_folder)
                )
            except Exception as e:
                errors.append(f"{Path(path).name}: {e}")
        self.after(0, self._finish, results, errors)

    def _finish(self, results, errors):
        self._set_busy(False)
        self._progress.pack_forget(); self._progress.set(0)
        if errors: messagebox.showerror("Errors", "\n".join(errors))
        if results:
            folder = str(Path(results[0]).parent)
            n = len(results)
            self._status.configure(text=f"✓  Saved {n} file{'s' if n>1 else ''}  →  {folder}")
            if messagebox.askyesno("Done!", f"Done!  {n} file{'s' if n>1 else ''} saved.\n\nOpen folder?"):
                os.startfile(folder)
        else:
            self._status.configure(text="No files were processed.")

    def _set_busy(self, busy):
        s = "disabled" if busy else "normal"
        self._btn_process.configure(state=s,
                                    text="Processing…" if busy else "✨  Process")
        self._btn_files.configure(state=s)
        self._btn_folder.configure(state=s)
        if busy: self._btn_close_sel.place_forget()
        elif self._files: self._btn_close_sel.place(relx=1.0, x=-6, y=6, anchor="ne")


if __name__ == "__main__":
    files_from_cli = [f for f in sys.argv[1:] if os.path.isfile(f)]
    app = App()
    if files_from_cli:
        app.after(800, lambda: app._load(files_from_cli))
    app.mainloop()
