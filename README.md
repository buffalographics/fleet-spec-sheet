# Fleet Spec Sheet Generator

This project generates fleet vehicle instruction/spec sheets from:
- Approved proof image
- PRINT.zip containing print PDFs

## Font Handling
HudsonNY.ttf is stored locally in:
```
assets/fonts/HudsonNY.ttf
```
No prompts or config files required.

## Setup
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run
```bash
python -m fleet_spec_sheet.cli --proof proof.jpg --print-zip PRINT.zip --out out/spec_sheet.pdf
```
