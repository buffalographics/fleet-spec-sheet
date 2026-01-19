# Fleet Spec Sheet Generator (Buffalo Graphics) — V1

Generates a one-page PDF “spec sheet” for fleet vehicle production + install.

The sheet includes:
- Header bar
- Job details table (Customer, Vehicle, blank Unit # for handwriting)
- Boxed proof/mockup image
- Print Manifest table (Thumbnail, Name, File Name, Qty, Done)
- Install Sign-Off section (blank fields)

Built to match Buffalo Graphics’ daily workflow and formatting standards.

---

## What it does

### Inputs
- A **proof/mockup image** (PNG/JPG/TIF)
- A **PRINT folder** (or PRINT.zip in CLI mode) containing print-ready PDFs
- HudsonNY font (defaults to `assets/HudsonNY.ttf`)

### Output
- A single PDF saved to the folder you choose  
  - Default save location and initial folder are based on the proof’s folder
  - Filename auto-suggested:
    `YYYY-MM-DD__CUSTOMER__VEHICLE__Spec_Sheet.pdf`

---

## Rules / Behavior

- Only **PDF files** are included in the manifest (non-PDF ignored, including `.vw`)
- Folders matching are ignored anywhere in the PRINT tree:
  - `unit number`, `unit numbers`, `unit #`, `unit no` (case-insensitive)
- Qty parsing supports:
  - `QTY2`, `QTY_2`, `QTY-2`, `QTY 2`, `x2`, `2x`
- Thumbnails:
  - Never skew (aspect preserved)
  - Centered in the thumbnail cell
- Fonts:
  - HudsonNY used everywhere
- Unit # is intentionally blank for handwritten entry after printing

---

## Install

### Create virtual environment + install dependencies
```bash
cd /Users/printstation/Downloads/fleet-spec-sheet
python3 -m venv .venv
source .venv/bin/activate
pip install pymupdf pillow reportlab
```

---

## Usage (Recommended: GUI)

### GUI workflow
```bash
source .venv/bin/activate
python main.py --gui
```

GUI order:
1. DETAILS (single dialog with Customer + Vehicle)
2. Select PROOF image
3. Select PRINT folder (opens in the proof’s folder)
4. Font prompt only if `assets/HudsonNY.ttf` is missing
5. Save PDF (defaults to proof’s folder with auto filename)

---

## Usage (CLI)

You can run without the GUI if you want to script it.

```bash
source .venv/bin/activate

python main.py \
  --print-dir "/path/to/PRINT" \
  --proof "/path/to/PROOF.png" \
  --out "/path/to/output/Spec_Sheet.pdf" \
  --customer "AM Corp (American Materials Corporation)" \
  --vehicle "MIXER TRUCK"
```

Font defaults to `assets/HudsonNY.ttf`. If you store it elsewhere:
```bash
python main.py --font "/path/to/HudsonNY.ttf" ...
```

---

## Project Structure

Recommended:
```
fleet-spec-sheet/
  main.py
  assets/
    HudsonNY.ttf
  README.md
```

---

## Troubleshooting

### “tkinter is not available”
You’re running a Python build without Tk support.

Quick check:
```bash
python3 -c "import tkinter; print('ok')"
```

If it fails, use a Python that includes Tk (commonly the system Python on macOS, or a python.org installer build). Alternatively, run CLI mode.

### “No PDFs found in PRINT after filtering”
- Confirm the PRINT folder contains PDFs
- Confirm they aren’t all inside ignored “unit number” folders
- Confirm the PDFs open normally

---

## Roadmap (optional)
- Remember last Customer/Vehicle (settings persistence)
- Autosave mode (no Save dialog)
- macOS double-click launcher (.app)
- Multi-page manifest when PRINT list is long

---

## Version

**V1** — Initial production-ready release  
- Locked layout and formatting  
- GUI-driven daily workflow  
- Manual unit number (handwritten)  
- Stable ruleset for PRINT parsing  
