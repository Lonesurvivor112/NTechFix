#!/usr/bin/env python3
"""
Generate and Print Snipe-IT Asset Tag labels.
Supports arbitrary label sizes (e.g. 60x80mm, 40x30mm).

Modes: print (USB/driver), preview, raw (ESC/POS), aml (Labelife LPAPI),
       bt (Bluetooth to Phomemo M220).

Examples:
  python AssetTag.py EC-12345 --mode preview
  python AssetTag.py EC-12345 --mode preview --width 40 --height 30
  python AssetTag.py EC-12345 --mode print
  python AssetTag.py A0028,A0022,A0293,A0160 --mode print --width 40 --height 30
  python AssetTag.py A0028 A0022 A0160 --mode print
  python AssetTag.py EC-12345 --mode aml --out-file label.aml
  python AssetTag.py EC-12345 --mode bt --bt-mac 68:70:18:68:AD:0E

Helper commands (NEW):
  python AssetTag.py --check                  # report installed deps / fonts / token / printer
  python AssetTag.py --list-printers          # list Windows printers and exit
  python AssetTag.py A0028 --printer "M220"   # override the printer name at runtime
  python AssetTag.py A0028 --yes              # auto-install missing deps without prompting
  python AssetTag.py A0028 --no-deps-check    # skip the dependency check

===============================================================================
 IMPORTANT: CHANGING LABEL SIZES (read this before switching label stock!)
===============================================================================

 The --width and --height flags control the IMAGE the script renders, but the
 M220 printer driver has its OWN paper size setting that MUST match. If they
 disagree, the print will be squished, shifted, or clipped.

 To switch label sizes (e.g. from 60x80mm to 40x30mm):

   STEP 1 - Change the DRIVER paper size:
     1. Open Windows Settings > Bluetooth & devices > Printers & scanners
     2. Click "M220 Printer" > Printing preferences
     3. Set paper/media size to match your labels (e.g. "40x30mm" or create
        a custom size with Width=40, Height=30)
     4. Click OK / Apply

   STEP 2 - Change the SCRIPT dimensions:
     python AssetTag.py B0024 --mode print --width 40 --height 30

 Both steps are required. The driver tells the printer how much paper to
 feed; the script tells Pillow how many pixels to render. If only one is
 changed, you get misaligned prints.

 TIP: Use --mode preview first to verify the layout looks correct before
      wasting a label.
 TIP: Use --save-png debug.png to save the exact image being sent to the printer.
 TIP: If --mode print gives bad results despite correct driver settings,
      try --mode raw which bypasses the driver entirely.

===============================================================================

Environment variables:
  SNIPE_TOKEN  - Your Snipe-IT API key (or paste into the config below)
"""
# ========== CONFIG ==========
SNIPE_URL = "https://snipe.domain.com"

import os

SNIPE_TOKEN = os.environ.get(
    "SNIPE_TOKEN",
    ""  # $env:SNIPE_TOKEN = "" Paste your token here as fallback, or set the SNIPE_TOKEN env var
)

# Default label size in mm (overridden via --width / --height)
DEFAULT_WIDTH_MM  = 60
DEFAULT_HEIGHT_MM = 80

# Default render DPI (print/preview/AML). BT mode auto-switches to 203 for M220.
DEFAULT_DPI = 203

# Font paths - update for your OS (these are Windows defaults)
FONT_PATH_REGULAR = r"C:\Windows\Fonts\arial.ttf"
FONT_PATH_BOLD    = r"C:\Windows\Fonts\arialbd.ttf"

# Windows printer name for USB/driver printing ("" = default printer)
PRINTER_NAME = "M220 Printer"

# Debug: draw red border + margin guides in preview mode
DEBUG_OVERLAY = False

# ========== IMPORTS (stdlib) ==========
import sys
import io
import base64
import time
import argparse
import subprocess
import textwrap

# ============================================================================
# DEPENDENCY BOOTSTRAP  (stdlib only - MUST run before the third-party imports
# below, because importing a missing package would crash before any checker
# could run). Prompts to pip-install whatever this run needs.
# ============================================================================
import importlib
import importlib.util

_CORE_DEPS = [
    ("requests", "requests"),
    ("PIL",      "Pillow"),
    ("qrcode",   "qrcode"),
]
_MODE_DEPS = {
    "preview": [("matplotlib", "matplotlib")],
    "print":   [("win32print", "pywin32")],
    "raw":     [("win32print", "pywin32")],
    # phomemo_printer is often installed from source, not PyPI -> best-effort.
    "bt":      [("phomemo_printer", "phomemo_printer")],
    "aml":     [],
}
_MIN_PYTHON = (3, 8)


def _argv_value(flag, default=None):
    """Read --flag VALUE or --flag=VALUE straight from sys.argv (pre-argparse)."""
    for i, a in enumerate(sys.argv):
        if a == flag and i + 1 < len(sys.argv):
            return sys.argv[i + 1]
        if a.startswith(flag + "="):
            return a.split("=", 1)[1]
    return default


def _has_flag(*names):
    return any(n in sys.argv for n in names)


def _module_present(import_name):
    """True if a module can be imported, without importing it."""
    try:
        return importlib.util.find_spec(import_name) is not None
    except (ImportError, ValueError):
        return False


def _pip_install(pip_name):
    """pip-install a package into the CURRENT interpreter (sys.executable)."""
    print(f"[deps] Installing: {pip_name} ...")
    cmd = [sys.executable, "-m", "pip", "install", "--upgrade", pip_name]
    try:
        subprocess.check_call(cmd)
        importlib.invalidate_caches()
        return True
    except Exception as e:
        print(f"[deps] pip failed for {pip_name}: {e}")
        return False


def _bootstrap_dependencies():
    """Check (and optionally install) the pip packages this run needs."""
    if _has_flag("--no-deps-check"):
        return

    if sys.version_info < _MIN_PYTHON:
        print(f"[deps] WARNING: Python {_MIN_PYTHON[0]}.{_MIN_PYTHON[1]}+ "
              f"recommended; you have {sys.version.split()[0]}")

    check_only = _has_flag("--check")
    mode = (_argv_value("--mode", "print") or "print").strip()
    assume_yes = _has_flag("-y", "--yes", "--ensure-deps", "--check")

    # For --check we only force-install the CORE libs (so the script can run
    # and report); mode-specific gaps are reported, not silently installed.
    wanted = list(_CORE_DEPS) if check_only else list(_CORE_DEPS) + _MODE_DEPS.get(mode, [])
    optional = set(_MODE_DEPS.get(mode, []))
    missing = [(imp, pip) for imp, pip in wanted if not _module_present(imp)]
    if not missing:
        return

    print("[deps] Missing package(s): " + ", ".join(pip for _, pip in missing))
    interactive = bool(getattr(sys, "stdin", None)) and sys.stdin.isatty()
    if not assume_yes and interactive:
        try:
            resp = input("[deps] Install them now with pip? [Y/n] ").strip().lower()
        except EOFError:
            resp = "y"
        if resp in ("n", "no"):
            print("[deps] Skipping install - the script may fail for this mode.")
            return
    elif not assume_yes and not interactive:
        print("[deps] Non-interactive session; attempting install automatically.")

    for imp, pip in missing:
        ok = _pip_install(pip)
        if not ok:
            if (imp, pip) in optional:
                print(f"[deps] '{pip}' is optional for mode '{mode}'; continuing.")
            else:
                print(f"[deps] '{pip}' is REQUIRED. Install manually:\n"
                      f"       {sys.executable} -m pip install {pip}")


_bootstrap_dependencies()
# ============================ END BOOTSTRAP ==================================

# ========== IMPORTS (third-party) ==========
import requests
from PIL import Image, ImageDraw, ImageFont
import qrcode

# Windows-only printing (safe to fail on other platforms)
try:
    import win32print
    import win32ui
    from PIL import ImageWin
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# Matplotlib for preview
try:
    import matplotlib.pyplot as plt
    HAS_MPL = True
except ImportError:
    HAS_MPL = False

# Optional Phomemo BT library
try:
    from phomemo_printer.ESCPOS_printer import Printer
    HAS_PHOMEMO_LIB = True
except Exception:
    HAS_PHOMEMO_LIB = False


# ========== UTILS ==========

def load_font(path, size):
    """Load a TrueType font, fall back to Pillow default."""
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()


def fetch_asset(tag):
    """Fetch asset details from Snipe-IT by asset tag."""
    if not SNIPE_TOKEN:
        raise RuntimeError(
            "SNIPE_TOKEN is not set. Set the SNIPE_TOKEN environment variable "
            "or paste your API key into the script config."
        )
    url = (
        f"{SNIPE_URL.rstrip('/')}/api/v1/hardware/bytag/"
        f"{requests.utils.quote(tag)}?deleted=false"
    )
    headers = {
        "Authorization": f"Bearer {SNIPE_TOKEN}",
        "Accept": "application/json",
    }
    r = requests.get(url, headers=headers, timeout=30)
    try:
        data = r.json()
    except Exception:
        data = {"error": r.text}
    if r.status_code >= 400 or "id" not in data:
        raise RuntimeError(f"Error fetching asset: HTTP {r.status_code} -> {data}")
    return data


def normalize_purchase_date(v):
    """Extract a readable date string from Snipe-IT's purchase_date field."""
    if isinstance(v, str):
        return v
    if isinstance(v, dict):
        return v.get("formatted") or v.get("date") or ""
    return ""


def make_qr(url):
    """Generate a QR code PIL image for the given URL."""
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M)
    qr.add_data(url)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white").convert("L")


# ========== RENDERING ==========

def draw_label(name, purchase, serial, asset_tag, asset_id,
               width_mm, height_mm, dpi):
    """Render a centered label image at the given DPI for any label size."""

    # --- Canvas ---
    W = int((width_mm / 25.4) * dpi)
    H = int((height_mm / 25.4) * dpi)
    img = Image.new("L", (W, H), 255)
    d = ImageDraw.Draw(img)

    # --- Safe margins (5% each side) ---
    MARGIN_X = int(W * 0.05)
    MARGIN_Y = int(H * 0.05)
    usable_w = W - 2 * MARGIN_X
    usable_h = H - 2 * MARGIN_Y

    # --- Dynamic font sizing ---
    ref = min(W, H)
    label_font_size = max(8, int(ref * 0.055))
    tag_font_size   = max(9, int(ref * 0.065))

    label_font = load_font(FONT_PATH_BOLD, label_font_size)
    tag_font   = load_font(FONT_PATH_BOLD, tag_font_size)

    # --- Build text lines ---
    lines = [
        ("Name: ",     name,      label_font),
        ("Purchase: ", purchase,  label_font),
        ("Serial: ",   serial,    label_font),
        ("Tag: ",      asset_tag, tag_font),
    ]

    def shorten(text, font, max_w):
        """Truncate text with ellipsis if it exceeds max_w pixels."""
        if d.textlength(text, font=font) <= max_w:
            return text
        for i in range(len(text), 0, -1):
            if d.textlength(text[:i] + "…", font=font) <= max_w:
                return text[:i] + "…"
        return text

    # Measure each line and compute the max width
    rendered_lines = []
    max_text_w = 0
    line_heights = []

    for prefix, value, font in lines:
        full = prefix + shorten(value, font, usable_w - int(d.textlength(prefix, font=font)))
        tw = int(d.textlength(full, font=font))
        th = font.getbbox("Ag")[3]
        leading = int(th * 0.35)
        rendered_lines.append((full, font, tw, th))
        line_heights.append(th + leading)
        max_text_w = max(max_text_w, tw)

    total_text_h = sum(line_heights)

    # --- QR sizing ---
    asset_url = f"{SNIPE_URL.rstrip('/')}/hardware/{asset_id}"
    qr_img = make_qr(asset_url)

    qr_gap = int(min(W, H) * 0.04)
    available_for_qr = usable_h - total_text_h - qr_gap

    qr_size = max(int(min(W, H) * 0.12), min(available_for_qr, usable_w))
    qr_size = max(qr_size, 1)

    if qr_size > available_for_qr:
        qr_size = max(int(available_for_qr * 0.90), int(min(W, H) * 0.10))
    qr_size = max(qr_size, 1)

    qr_img = qr_img.resize((qr_size, qr_size), Image.NEAREST)

    # --- Compute total content block ---
    content_w = max(max_text_w, qr_size)
    content_h = total_text_h + qr_gap + qr_size

    # --- Center the block on the canvas ---
    block_x = max(MARGIN_X, (W - content_w) // 2)
    block_y = max(MARGIN_Y, (H - content_h) // 2)

    # --- Draw text lines ---
    y = block_y
    for i, (text, font, tw, th) in enumerate(rendered_lines):
        d.text((block_x, y), text, font=font, fill=0)
        y += line_heights[i]

    # --- Draw QR code (left-aligned with text block) ---
    qr_y = y + qr_gap
    if qr_y + qr_size > H - MARGIN_Y:
        qr_y = H - MARGIN_Y - qr_size
    img.paste(qr_img, (block_x, max(0, qr_y)))

    # --- Debug overlay ---
    if DEBUG_OVERLAY:
        d.rectangle((0, 0, W - 1, H - 1), outline=0, width=2)
        d.line([(MARGIN_X, 0), (MARGIN_X, H - 1)], fill=0, width=1)
        d.line([(W - MARGIN_X, 0), (W - MARGIN_X, H - 1)], fill=0, width=1)
        d.line([(0, MARGIN_Y), (W - 1, MARGIN_Y)], fill=0, width=1)
        d.line([(0, H - MARGIN_Y), (W - 1, H - MARGIN_Y)], fill=0, width=1)

    return img


# ========== PREVIEW ==========

def preview_label(img, width_mm, height_mm):
    """Show a matplotlib preview at physical label size."""
    if not HAS_MPL:
        raise RuntimeError("matplotlib is required for preview mode (pip install matplotlib)")
    fig_w = width_mm / 25.4
    fig_h = height_mm / 25.4
    plt.figure(figsize=(max(fig_w, 2), max(fig_h, 2)))
    plt.imshow(img, cmap="gray")
    plt.axis("off")
    plt.title(f"Preview - {width_mm}x{height_mm} mm")
    plt.tight_layout()
    plt.show()


# ========== PRINTER PREFLIGHT ==========

def list_installed_printers():
    """Return a list of installed printer names (local + connections)."""
    if not HAS_WIN32:
        return []
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    names = []
    for p in win32print.EnumPrinters(flags):
        if isinstance(p, dict):
            names.append(p.get("pPrinterName") or p.get("Name"))
        else:
            names.append(p[2])  # level-1 tuple: (flags, desc, name, comment)
    return [n for n in names if n]


def resolve_printer_name(preferred):
    """
    Find the real installed printer to use.
    Order: exact -> case-insensitive -> fuzzy keyword (m220/phomemo) -> default.
    Returns (name, note) or (None, message).
    """
    printers = list_installed_printers()
    if not printers:
        return None, "No printers are installed (or pywin32 can't see them)."

    if preferred:
        for n in printers:
            if n == preferred:
                return n, "exact match"
        for n in printers:
            if n.lower() == preferred.lower():
                return n, f"case-insensitive match: {n}"

    for kw in ("m220", "phomemo"):
        for n in printers:
            if kw in n.lower():
                return n, f"fuzzy match on '{kw}': {n}"

    try:
        d = win32print.GetDefaultPrinter()
        if d:
            return d, f"using system default: {d}"
    except Exception:
        pass

    return None, "no match; installed: " + ", ".join(printers)


def printer_health(name):
    """Return a human-readable status string for a printer."""
    if not HAS_WIN32:
        return "pywin32 not available"
    try:
        h = win32print.OpenPrinter(name)
        try:
            info = win32print.GetPrinter(h, 2)
        finally:
            win32print.ClosePrinter(h)
    except Exception as e:
        return f"could NOT open '{name}': {e}"

    status = info.get("Status", 0)
    attrs = info.get("Attributes", 0)
    jobs = info.get("cJobs", 0)
    bits = {
        0x00000001: "PAUSED",
        0x00000002: "ERROR",
        0x00000080: "OFFLINE",
        0x00000400: "PAPER_OUT",
        0x00000800: "PAPER_PROBLEM",
        0x00200000: "DOOR_OPEN",
    }
    flags = [label for bit, label in bits.items() if status & bit]
    if attrs & 0x00000400:  # PRINTER_ATTRIBUTE_WORK_OFFLINE
        flags.append("SET_OFFLINE")
    state = ", ".join(flags) if flags else "ready"
    return f"status={state}; queued jobs={jobs}"


def preflight_printer(preferred):
    """
    Resolve + sanity-check the printer before printing. Returns the real
    printer name to use, or raises RuntimeError with an actionable message.
    """
    name, note = resolve_printer_name(preferred)
    if not name:
        installed = list_installed_printers()
        raise RuntimeError(
            f"Printer '{preferred}' was not found.\n"
            f"  Installed printers: {installed or '(none)'}\n"
            f"  Fix: pass --printer \"Exact Name\" (or edit PRINTER_NAME), or\n"
            f"  install/add the M220 driver first (see --list-printers)."
        )
    if name != preferred:
        print(f"[warn] Printer '{preferred}' not found exactly; {note}")
    health = printer_health(name)
    print(f"[info] Printer health: {health}")
    if "could NOT open" in health or "OFFLINE" in health or "ERROR" in health:
        print("[warn] Printer may be offline/errored - check it's powered on, "
              "paired, and not paused in the print queue.")
    return name


# ========== PRINT (Multiple strategies) ==========

def _set_devmode_paper(printer, width_mm, height_mm):
    """Configure the driver's DEVMODE to match our label size. Returns hDC or None."""
    try:
        import win32con
        hPrinter = win32print.OpenPrinter(printer)
        try:
            devmode = win32print.GetPrinter(hPrinter, 2)["pDevMode"]
            devmode.PaperSize = 256  # DMPAPER_USER / custom
            devmode.PaperWidth = width_mm * 10   # DEVMODE uses tenths of mm
            devmode.PaperLength = height_mm * 10
            devmode.Orientation = 1  # Portrait
            devmode.Fields |= (0x00000002 |  # DM_PAPERSIZE
                               0x00000008 |  # DM_PAPERLENGTH
                               0x00000004 |  # DM_PAPERWIDTH
                               0x00000001)   # DM_ORIENTATION
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer)
            hDC.ResetDC(devmode)
            return hDC
        finally:
            win32print.ClosePrinter(hPrinter)
    except Exception as e:
        print(f"[warn] DEVMODE configuration failed: {e}")
        return None


def print_image(img, width_mm, height_mm):
    """Print via Windows GDI with custom DEVMODE paper size."""
    if not HAS_WIN32:
        raise RuntimeError("win32print/win32ui required for print mode (install pywin32)")

    # Resolve + health-check the printer (handles wrong/renamed PRINTER_NAME).
    printer = preflight_printer(PRINTER_NAME)

    hDC = _set_devmode_paper(printer, width_mm, height_mm)
    if hDC:
        print(f"[info] DEVMODE set to {width_mm}x{height_mm} mm custom paper")
    else:
        print(f"[warn] Using default DC (label may not be sized correctly)")
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer)

    printer_w  = hDC.GetDeviceCaps(8)   # HORZRES
    printer_h  = hDC.GetDeviceCaps(10)  # VERTRES
    offset_x   = hDC.GetDeviceCaps(112) # PHYSICALOFFSETX
    offset_y   = hDC.GetDeviceCaps(113) # PHYSICALOFFSETY
    phys_w     = hDC.GetDeviceCaps(110) # PHYSICALWIDTH
    phys_h     = hDC.GetDeviceCaps(111) # PHYSICALHEIGHT
    dpi_x      = hDC.GetDeviceCaps(88)  # LOGPIXELSX
    dpi_y      = hDC.GetDeviceCaps(90)  # LOGPIXELSY

    img_w, img_h = img.size

    print(f"[info] Printer : {printer}")
    print(f"[info] Driver DPI    : {dpi_x} x {dpi_y}")
    print(f"[info] Physical paper: {phys_w} x {phys_h} px "
          f"({phys_w/max(dpi_x,1)*25.4:.1f} x {phys_h/max(dpi_y,1)*25.4:.1f} mm)")
    print(f"[info] Printable area: {printer_w} x {printer_h} px, "
          f"offset: ({offset_x}, {offset_y})")
    print(f"[info] Image size    : {img_w} x {img_h} px")

    hDC.StartDoc("Asset Tag Print")
    hDC.StartPage()

    dib = ImageWin.Dib(img)
    dib.draw(hDC.GetHandleOutput(), (0, 0, printer_w, printer_h))

    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()
    print(f"[ok] Printed to {printer} - mapped to {printer_w}x{printer_h} printable area")


def print_image_raw(img, width_mm, height_mm):
    """Bypass the driver: send raw ESC/POS image data via the spooler in RAW mode."""
    if not HAS_WIN32:
        raise RuntimeError("win32print required for raw print mode")

    # Resolve + health-check the printer (handles wrong/renamed PRINTER_NAME).
    printer_name = preflight_printer(PRINTER_NAME)

    bw = img.convert("1")
    width_px, height_px = bw.size

    PRINT_HEAD_DOTS = 384
    if width_px != PRINT_HEAD_DOTS:
        ratio = PRINT_HEAD_DOTS / width_px
        new_h = int(height_px * ratio)
        bw = bw.resize((PRINT_HEAD_DOTS, new_h), Image.LANCZOS)
        width_px, height_px = bw.size

    print(f"[info] RAW mode: {width_px}x{height_px} px, head width={PRINT_HEAD_DOTS} dots")

    bytes_per_row = width_px // 8
    pixels = bw.load()

    esc_init = b'\x1b\x40'        # ESC @  - initialize printer
    esc_linespace = b'\x1b\x33\x00'  # ESC 3 0 - line spacing 0

    xL = bytes_per_row & 0xFF
    xH = (bytes_per_row >> 8) & 0xFF
    yL = height_px & 0xFF
    yH = (height_px >> 8) & 0xFF

    raster_header = bytes([0x1d, 0x76, 0x30, 0x00, xL, xH, yL, yH])

    img_bytes = bytearray()
    for row_y in range(height_px):
        for byte_x in range(bytes_per_row):
            byte_val = 0
            for bit in range(8):
                px = byte_x * 8 + bit
                if px < width_px:
                    if pixels[px, row_y] == 0:  # 0=black in PIL "1"
                        byte_val |= (0x80 >> bit)
            img_bytes.append(byte_val)

    esc_feed = b'\x1b\x64\x03'  # ESC d 3 - feed 3 lines
    raw_data = esc_init + esc_linespace + raster_header + bytes(img_bytes) + esc_feed

    print(f"[info] Sending {len(raw_data)} bytes RAW to {printer_name}")

    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(hPrinter, 1, ("Asset Tag", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)
        win32print.WritePrinter(hPrinter, raw_data)
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)

    print(f"[ok] RAW printed to {printer_name}")


# ========== AML (Labelife LPAPI) ==========

def png_bytes(img):
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def make_lpapi_aml(img, width_mm, height_mm, dpi, label_name="No Title 1",
                   paper_name="White-6080", platform="win",
                   x_mm=0.0, y_mm=0.0, w_mm=None, h_mm=None,
                   orientation=0):
    if w_mm is None:
        w_mm = width_mm
    if h_mm is None:
        h_mm = height_mm

    display_mm = f"{width_mm}mm x {height_mm}mm"
    display_in = f"{width_mm / 25.4:.2f}\" x {height_mm / 25.4:.2f}\""
    save_time_ms = int(time.time() * 1000)
    content_b64 = base64.b64encode(png_bytes(img)).decode("ascii")

    common_page = f"""\
      <isPrintHorizontal>0</isPrintHorizontal>
      <labelHeight>{height_mm}</labelHeight>
      <labelWidth>{width_mm}</labelWidth>
      <paperName>{paper_name}</paperName>
      <validBoundsX>0</validBoundsX>
      <validBoundsY>0</validBoundsY>
      <validBoundsWidth>{width_mm}</validBoundsWidth>
      <validBoundsHeight>{height_mm}</validBoundsHeight>
      <paperType>0</paperType>
      <paperDesc />
      <paperBackground>#FFFFFFFF</paperBackground>
      <paperForeground>#FF000000</paperForeground>
      <DisplaySize_mm>{display_mm}</DisplaySize_mm>
      <DisplaySize_in>{display_in}</DisplaySize_in>
      <isAutoHeight>0</isAutoHeight>
      <isRotate180>0</isRotate180>
      <leftBlank>0</leftBlank>
      <rightBlank>0</rightBlank>
      <topBlank>0</topBlank>
      <bottomBlank>0</bottomBlank>
      <isCustomSize>0</isCustomSize>
      <isBannerMode>0</isBannerMode>
      <flagType>0</flagType>
      <tailDirection>0</tailDirection>
      <tailLength>0</tailLength>
      <hasTail>0</hasTail>
      <paperCategory>0</paperCategory>
      <tag />
      <squareBackingPaper>0</squareBackingPaper>
      <columnCount>0</columnCount>
      <marginLeft>0</marginLeft>
      <marginRight>0</marginRight>
      <paddingWidth>0</paddingWidth>
      <isCustomTemplateSize>0</isCustomTemplateSize>
      <templateHeight>0</templateHeight>
      <templateWidth>0</templateWidth>"""

    xml = f"""<?xml version="1.0" encoding="utf-8"?>
<LPAPI version="2.1">
{common_page}
  <labelName>{label_name}</labelName>
  <Platform>{platform}</Platform>
  <saveTime>{save_time_ms}</saveTime>
  <contents>
    <WdPage>
{common_page}
      <masksToBoundsType>1</masksToBoundsType>
      <borderDisplay>0</borderDisplay>
      <lineType>0</lineType>
      <borderWidth>0.35277778</borderWidth>
      <borderColor>#FF000000</borderColor>
      <selectColor>#FF000000</selectColor>
      <isBorderTemplate>0</isBorderTemplate>
      <lockMovement>0</lockMovement>
      <useBgImage>0</useBgImage>
      <group />
      <contents>
        <Image>
          <objectId>{os.urandom(16).hex()}</objectId>
          <x>{x_mm}</x>
          <y>{y_mm}</y>
          <width>{w_mm}</width>
          <height>{h_mm}</height>
          <borderDisplay>0</borderDisplay>
          <lineType>0</lineType>
          <borderHeight>0.35277778</borderHeight>
          <borderColor>#FF000000</borderColor>
          <textColor>#FF000000</textColor>
          <bgAlpha>0</bgAlpha>
          <bgColor>#00000000</bgColor>
          <orientation>{orientation}</orientation>
          <lockMovement>0</lockMovement>
          <isTemplate>0</isTemplate>
          <antiColor>0</antiColor>
          <selectColor>#FF000000</selectColor>
          <isRedBlack>0</isRedBlack>
          <isRatioScale>1</isRatioScale>
          <content>{content_b64}</content>
        </Image>
      </contents>
    </WdPage>
  </contents>
</LPAPI>
"""
    return xml


def export_aml(img, out_file, width_mm, height_mm, dpi, **kwargs):
    aml_xml = make_lpapi_aml(img, width_mm, height_mm, dpi, **kwargs)
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(aml_xml)


# ========== BLUETOOTH (Phomemo M220) ==========

def bt_print_image(img, bt_mac, bt_channel, use_cli=False):
    """Print a PIL image to a Phomemo printer via Bluetooth RFCOMM."""
    tmp_png = os.path.abspath(f"_phomemo_{int(time.time())}.png")
    img.save(tmp_png, "PNG")
    try:
        if not use_cli and HAS_PHOMEMO_LIB:
            printer = Printer(bluetooth_address=bt_mac, channel=int(bt_channel))
            if hasattr(printer, "print_image"):
                printer.print_image(tmp_png)
            elif hasattr(printer, "print_file"):
                printer.print_file(tmp_png)
            else:
                raise RuntimeError("phomemo_printer module lacks print_image/print_file")
            printer.close()
            return True
        else:
            cmd = [
                sys.executable, "-m", "phomemo_printer",
                "-a", bt_mac, "-c", str(bt_channel), "-i", tmp_png,
            ]
            r = subprocess.run(cmd, capture_output=True, text=True)
            if r.returncode != 0:
                if str(bt_channel) == "1":
                    cmd[cmd.index("1")] = "6"
                    r2 = subprocess.run(cmd, capture_output=True, text=True)
                    if r2.returncode == 0:
                        print("[info] RFCOMM channel 1 failed; channel 6 succeeded.")
                        return True
                raise RuntimeError(f"phomemo_printer CLI failed: {r.stderr or r.stdout}")
            return True
    finally:
        if os.path.exists(tmp_png):
            try:
                os.remove(tmp_png)
            except OSError:
                pass


# ========== PREREQUISITE CHECKER ==========

def check_prerequisites(mode):
    """
    Print a report of everything the script needs for `mode`.
    Returns True if all REQUIRED items are satisfied, else False.
    """
    print("=" * 60)
    print(f" Prerequisite check  (mode: {mode})")
    print("=" * 60)
    ok = True

    pyver = sys.version.split()[0]
    py_ok = sys.version_info >= _MIN_PYTHON
    print(f"[{'ok ' if py_ok else 'FAIL'}] Python {pyver}  ({sys.executable})")
    if not py_ok:
        ok = False

    if mode in ("print", "raw") and not sys.platform.startswith("win"):
        print(f"[warn] mode '{mode}' uses Windows printing APIs; OS is '{sys.platform}'")

    print("-" * 60)
    print(" Python packages:")
    for imp, pip in _CORE_DEPS:
        present = _module_present(imp)
        print(f"  [{'ok ' if present else 'MISS'}] {pip:<14} (import {imp})"
              f"{'' if present else '   <- REQUIRED'}")
        if not present:
            ok = False
    for imp, pip in _MODE_DEPS.get(mode, []):
        present = _module_present(imp)
        tag = "" if present else f"   <- needed for --mode {mode}"
        print(f"  [{'ok ' if present else 'MISS'}] {pip:<14} (import {imp}){tag}")
        if not present and imp == "win32print":
            ok = False

    print("-" * 60)
    print(" Fonts (missing = falls back to Pillow default):")
    for label, p in (("regular", FONT_PATH_REGULAR), ("bold", FONT_PATH_BOLD)):
        exists = os.path.exists(p)
        print(f"  [{'ok ' if exists else 'warn'}] {label:<8} {p}")

    print("-" * 60)
    has_token = bool(SNIPE_TOKEN)
    print(f"  [{'ok ' if has_token else 'warn'}] SNIPE_TOKEN "
          f"{'is set' if has_token else 'is NOT set (set env var or edit config)'}")
    print(f"  [info] SNIPE_URL    = {SNIPE_URL}")

    print("-" * 60)
    print(" Printer:")
    if not HAS_WIN32:
        print("  [warn] pywin32 not available - cannot inspect printers")
    else:
        printers = list_installed_printers()
        print(f"  [info] installed: {printers or '(none)'}")
        name, note = resolve_printer_name(PRINTER_NAME)
        if name:
            print(f"  [ok ] resolves '{PRINTER_NAME}' -> '{name}' ({note})")
            print(f"  [info] {printer_health(name)}")
        else:
            print(f"  [FAIL] '{PRINTER_NAME}' not found: {note}")
            if mode in ("print", "raw"):
                ok = False

    print("=" * 60)
    print(f" Result: {'ALL REQUIRED ITEMS OK' if ok else 'MISSING REQUIRED ITEMS'}")
    print("=" * 60)
    return ok


# ========== BATCH HELPERS ==========

def parse_tags(values):
    """
    Flatten the asset-tag argument(s) into a clean, ordered, de-duplicated list.
    Accepts space-separated args, comma-separated strings, and semicolons.
    Duplicates (case-insensitive) are dropped; original order is kept.
    """
    out, seen = [], set()
    for item in values:
        for piece in str(item).replace(";", ",").split(","):
            t = piece.strip()
            if t and t.lower() not in seen:
                seen.add(t.lower())
                out.append(t)
    return out


def render_and_output(tag, args, total=1):
    """
    Fetch ONE asset, render its label, and emit it according to --mode.
    Raises on failure so the batch loop can record it and keep going.
    With more than one label, AML/PNG outputs get a per-tag suffix.
    """
    asset = fetch_asset(tag)
    width_mm, height_mm = args.width, args.height

    img = draw_label(
        name=asset.get("name", ""),
        purchase=normalize_purchase_date(asset.get("purchase_date")),
        serial=asset.get("serial", ""),
        asset_tag=asset.get("asset_tag", tag),
        asset_id=asset.get("id"),
        width_mm=width_mm,
        height_mm=height_mm,
        dpi=args.dpi,
    )

    if args.save_png:
        base, ext = os.path.splitext(args.save_png)
        out_png = args.save_png if total == 1 else f"{base}_{tag}{ext or '.png'}"
        img.save(out_png, "PNG")
        print(f"[info] Saved {out_png} ({img.size[0]}x{img.size[1]} px @ {args.dpi} DPI)")

    if args.mode == "preview":
        preview_label(img, width_mm, height_mm)

    elif args.mode == "aml":
        base, ext = os.path.splitext(args.out_file)
        out_file = args.out_file if total == 1 else f"{base}_{tag}{ext or '.aml'}"
        paper = f"White-{width_mm:02d}{height_mm:02d}"
        export_aml(
            img, out_file=out_file, width_mm=width_mm, height_mm=height_mm,
            dpi=args.dpi, label_name=f"Asset {asset.get('asset_tag', tag)}",
            paper_name=paper, platform="win",
            x_mm=0.0, y_mm=0.0, w_mm=width_mm, h_mm=height_mm, orientation=0,
        )
        print(f"[ok] Saved AML to {out_file} ({width_mm}x{height_mm} mm @ {args.dpi} DPI)")

    elif args.mode == "bt":
        if not bt_print_image(img.convert("L"), args.bt_mac, args.bt_channel,
                              use_cli=args.bt_use_cli):
            raise RuntimeError("Bluetooth print returned failure")
        print(f"[ok] BT printed {tag} to {args.bt_mac} (channel {args.bt_channel})")

    elif args.mode == "raw":
        print_image_raw(img, width_mm, height_mm)

    else:
        print_image(img, width_mm, height_mm)

    return True


# ========== MAIN ==========

def main():
    parser = argparse.ArgumentParser(
        description="Generate & Print Snipe-IT Asset Tag(s). Modes: print | raw | preview | aml | bt"
    )
    parser.add_argument(
        "ASSET_TAG", nargs="*",
        help="One or more asset tags. Space- or comma-separated, e.g. "
             "'A0028 A0022' or 'A0028,A0022,A0293,A0160'. "
             "(Not required with --check or --list-printers.)",
    )
    parser.add_argument(
        "--mode", choices=["print", "raw", "preview", "aml", "bt"], default="print",
        help="Output mode (default: print)",
    )
    parser.add_argument("--width", type=int, default=DEFAULT_WIDTH_MM,
                        help=f"Label width in mm (default: {DEFAULT_WIDTH_MM})")
    parser.add_argument("--height", type=int, default=DEFAULT_HEIGHT_MM,
                        help=f"Label height in mm (default: {DEFAULT_HEIGHT_MM})")
    parser.add_argument("--dpi", type=int, default=DEFAULT_DPI,
                        help=f"Render DPI (default: {DEFAULT_DPI})")
    parser.add_argument("--out-file", default="label.aml",
                        help="Output path for --mode aml (per-tag suffixed in a batch)")
    parser.add_argument("--delay", type=float, default=1.0,
                        help="Seconds between prints in a batch (default: 1.0; 0 disables)")
    parser.add_argument("--printer", default=None,
                        help="Exact Windows printer name (overrides PRINTER_NAME)")
    # Bluetooth options
    parser.add_argument("--bt-mac", default="68:70:18:68:AD:0E", help="Bluetooth MAC of the M220")
    parser.add_argument("--bt-channel", type=int, default=1,
                        help="RFCOMM channel (auto-retries 6 on failure)")
    parser.add_argument("--bt-use-cli", action="store_true",
                        help="Force CLI path instead of module for BT printing")
    parser.add_argument("--debug", action="store_true", help="Show margin guides in preview")
    parser.add_argument("--save-png", default="",
                        help="Save the rendered label image(s) to PNG for inspection")
    # Prerequisite / dependency helpers
    parser.add_argument("--check", action="store_true",
                        help="Run the prerequisite check (packages, fonts, token, printer) and exit")
    parser.add_argument("--list-printers", action="store_true",
                        help="List installed Windows printers and exit")
    parser.add_argument("--yes", "-y", action="store_true",
                        help="Auto-install any missing pip packages without prompting")
    parser.add_argument("--no-deps-check", action="store_true",
                        help="Skip the automatic dependency check at startup")
    parser.add_argument("--ensure-deps", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()

    global DEBUG_OVERLAY, PRINTER_NAME
    if args.debug:
        DEBUG_OVERLAY = True
    if args.printer:
        PRINTER_NAME = args.printer

    # --list-printers: show what Windows has and exit.
    if args.list_printers:
        if not HAS_WIN32:
            print("pywin32 not available - cannot list printers.")
        else:
            default = win32print.GetDefaultPrinter()
            names = list_installed_printers()
            if not names:
                print("(no printers found)")
            for n in names:
                print(("* " if n == default else "  ") + n)
            print("\n(* = system default)")
        raise SystemExit(0)

    # --check: report prerequisites and exit.
    if args.check:
        ok = check_prerequisites(args.mode)
        raise SystemExit(0 if ok else 1)

    # BT mode: force M220 native DPI
    if args.mode == "bt":
        args.dpi = 203

    # Collect and validate the tag list.
    tags = parse_tags(args.ASSET_TAG)
    if not tags:
        parser.error("at least one ASSET_TAG is required (or use --check / --list-printers)")

    print(f"[batch] {len(tags)} label(s): {', '.join(tags)}")
    failures = []
    for i, tag in enumerate(tags, 1):
        print(f"\n[{i}/{len(tags)}] {tag}")
        try:
            render_and_output(tag, args, total=len(tags))
        except Exception as e:
            print(f"[FAIL] {tag}: {e}")
            failures.append((tag, str(e)))
        else:
            if args.mode in ("print", "raw", "bt") and i < len(tags):
                time.sleep(args.delay)

    ok_count = len(tags) - len(failures)
    print(f"\n[batch] Done - {ok_count} ok, {len(failures)} failed.")
    if failures:
        for t, msg in failures:
            print(f"   - {t}: {msg}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
