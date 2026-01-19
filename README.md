# Fleet Spec Sheet Generator (Buffalo Graphics)

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