#!/usr/bin/env python3
"""
Fleet Spec Sheet Generator (Buffalo Graphics) — CLI + GUI picker

GUI mode:
  python main.py --gui
  (opens file pickers for PRINT folder and proof image, then prompts for Customer and Vehicle; Unit left blank for handwriting)

CLI mode:
  python main.py --print-dir PRINT --proof PROOF.png --font HudsonNY.ttf --out out.pdf

Rules:
- Ignore non-PDF files (e.g. .vw)
- Ignore folders named like: "unit number", "unit numbers", "unit #", "unit no" (case-insensitive)
- Qty parsing supports: QTY2, QTY_2, QTY-2, QTY 2, x2, 2x
- Thumbnails: centered, never skewed (aspect preserved)
- HudsonNY everywhere
- No title line (only header bar)
"""

from __future__ import annotations

import argparse
import datetime
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from os.path import commonpath
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image as PILImage

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Image, Spacer


IGNORE_DIR_RE = re.compile(r"unit\s*(number|numbers|#|no\.?)", re.IGNORECASE)

QTY_PATTERNS = [
    r"(?:^|[_\-\s])qty(?:[_\-\s]*)(\d+)(?:[_\-\s\.]|$)",  # qty_2 / qty-2 / qty 2 / qty2
    r"(?:^|[_\-\s])x(\d+)(?:[_\-\s\.]|$)",                # x2
    r"(?:^|[_\-\s])(\d+)x(?:[_\-\s\.]|$)",                # 2x
]


@dataclass
class Item:
    pdf_path: Path
    rel_path: str
    qty: int
    thumb_path: Path


def parse_qty(filename: str) -> int:
    for pat in QTY_PATTERNS:
        m = re.search(pat, filename, flags=re.IGNORECASE)
        if m:
            try:
                q = int(m.group(1))
                return q if q > 0 else 1
            except ValueError:
                pass
    return 1


def safe_filename(s: str) -> str:
    # Keep it filesystem-safe across macOS/Windows
    s = re.sub(r"[^\w\s\-]+", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace(" ", "_")
    return s[:120] if s else "Spec_Sheet"


def is_ignored_by_folder(p: Path) -> bool:
    return any(IGNORE_DIR_RE.search(parent.name) for parent in p.parents)


def extract_print_source(print_zip: Path | None, print_dir: Path | None, work_dir: Path) -> Path:
    target = work_dir / "PRINT_SRC"
    target.mkdir(parents=True, exist_ok=True)

    if print_zip:
        with zipfile.ZipFile(print_zip, "r") as z:
            z.extractall(target)
        return target

    if print_dir:
        if not print_dir.exists():
            raise FileNotFoundError(f"PRINT folder not found: {print_dir}")
        shutil.copytree(print_dir, target, dirs_exist_ok=True)
        return target

    raise ValueError("Either --print-zip or --print-dir is required.")


def collect_pdf_items(print_root: Path, thumbs_dir: Path) -> list[Item]:
    pdfs: list[Path] = []

    for p in print_root.rglob("*"):
        if not p.is_file():
            continue
        if "__MACOSX" in p.parts or p.name.startswith("._") or p.name == ".DS_Store":
            continue
        if p.suffix.lower() != ".pdf":
            continue
        if is_ignored_by_folder(p):
            continue
        try:
            d = fitz.open(p)
            ok = d.page_count > 0
            d.close()
            if ok:
                pdfs.append(p)
        except Exception:
            continue

    if not pdfs:
        return []

    rel_root = Path(commonpath([str(p.parent) for p in pdfs]))

    items: list[Item] = []
    for i, p in enumerate(sorted(pdfs, key=lambda x: str(x).lower()), 1):
        rel = str(p.relative_to(rel_root)).replace("\\", "/")
        q = parse_qty(p.name)

        d = fitz.open(p)
        pix = d.load_page(0).get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        d.close()

        thumb_path = thumbs_dir / f"thumb_{i}.png"
        pix.save(thumb_path)

        items.append(Item(pdf_path=p, rel_path=rel, qty=q, thumb_path=thumb_path))

    items.sort(key=lambda it: it.rel_path.lower())
    return items


def build_pdf(
    out_pdf: Path,
    proof_image: Path,
    font_path: Path,
    customer: str,
    vehicle_type: str,
    items: list[Item],
) -> None:
    pdfmetrics.registerFont(TTFont("HudsonNY", str(font_path)))

    section = ParagraphStyle("section", fontName="HudsonNY", fontSize=11)
    hdr = ParagraphStyle("hdr", fontName="HudsonNY", fontSize=9)
    body = ParagraphStyle("body", fontName="HudsonNY", fontSize=9, leading=10)

    doc = SimpleDocTemplate(
        str(out_pdf),
        pagesize=letter,
        leftMargin=0.5 * inch,
        rightMargin=0.5 * inch,
        topMargin=0.45 * inch,
        bottomMargin=0.45 * inch,
    )

    story: list = []

    bar = Table([["BUFFALO GRAPHICS COMPANY"]], [7.5 * inch])
    bar.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.black),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), "HudsonNY"),
                ("FONTSIZE", (0, 0), (-1, -1), 14),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story += [bar, Spacer(1, 8)]

    # Job details table (kept within 7.5" header width)
    info = Table(
        [
            ["CUSTOMER:", customer, "", "", "", ""],
            ["VEHICLE:", vehicle_type, "", "", "VEHICLE UNIT #:", ""],
        ],
        # Total = 7.5 inches
        [1.2 * inch, 4.4 * inch, 0.0 * inch, 0.0 * inch, 1.3 * inch, 0.6 * inch],
    )
    info.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("FONTNAME", (0, 0), (-1, -1), "HudsonNY"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                # Span customer and vehicle values across the middle columns
                ("SPAN", (1, 0), (5, 0)),
                ("SPAN", (1, 1), (3, 1)),
                # Keep labels left-aligned
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("ALIGN", (4, 1), (4, 1), "LEFT"),
            ]
        )
    )
    story.append(info)

    story.append(Spacer(1, 10))
    pil = PILImage.open(proof_image)
    w, h = pil.size
    proof_img = Image(str(proof_image), width=7.2 * inch, height=h * (7.2 * inch / w))
    proof_box = Table([[proof_img]], [7.5 * inch])
    proof_box.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    story.append(proof_box)
    story.append(Spacer(1, 12))

    story.append(Paragraph("<u>PRINT MANIFEST</u>", section))

    BOX_W = 0.95 * inch
    BOX_H = 0.70 * inch

    rows = [
        [
            Paragraph("THUMBNAIL", hdr),
            Paragraph("NAME", hdr),
            Paragraph("FILE NAME", hdr),
            Paragraph("QTY", hdr),
            Paragraph("DONE", hdr),
        ]
    ]

    for it in items:
        im = PILImage.open(it.thumb_path)
        tw, th = im.size
        s = min(BOX_W / tw, BOX_H / th)
        thumb = Image(str(it.thumb_path), width=tw * s, height=th * s)
        thumb.hAlign = "CENTER"

        display_name = Path(it.rel_path).name.replace("_", " ").replace(".pdf", "").upper()

        rows.append(
            [
                thumb,
                Paragraph(display_name, body),
                Paragraph(it.rel_path.upper(), body),
                Paragraph(str(it.qty), body),
                Paragraph("", body),
            ]
        )

    manifest = Table(
        rows,
        [1.4 * inch, 1.9 * inch, 2.6 * inch, 0.7 * inch, 0.9 * inch],
        rowHeights=[0.32 * inch] + [0.85 * inch] * (len(rows) - 1),
    )
    manifest.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("FONTNAME", (0, 0), (-1, -1), "HudsonNY"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ALIGN", (0, 1), (0, -1), "CENTER"),
                ("VALIGN", (0, 1), (0, -1), "MIDDLE"),
                ("VALIGN", (1, 1), (3, -1), "TOP"),
                ("ALIGN", (3, 1), (3, -1), "CENTER"),
            ]
        )
    )
    story.append(manifest)

    story.append(Spacer(1, 14))
    story.append(Paragraph("<u>INSTALL SIGN-OFF</u>", section))

    sign = Table(
        [
            ["INSTALLED BY:", "", "PHOTOS TAKEN:", "YES   NO", "DATE:", ""],
            ["ISSUES / DAMAGE NOTES:", "", "", "", "", ""],
        ],
        [1.3 * inch, 2.2 * inch, 1.1 * inch, 0.9 * inch, 0.7 * inch, 1.3 * inch],
        rowHeights=[0.35 * inch, 1.2 * inch],
    )
    sign.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("SPAN", (1, 1), (5, 1)),
                ("FONTNAME", (0, 0), (-1, -1), "HudsonNY"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("VALIGN", (1, 1), (5, 1), "TOP"),
            ]
        )
    )
    story.append(sign)

    doc.build(story)


def run_gui_defaults(args: argparse.Namespace) -> argparse.Namespace:
    """
    Prompts for missing inputs with native file pickers and simple dialogs.
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as e:
        raise RuntimeError(
            "tkinter is not available in this Python install. Use CLI mode or install a Python build with tkinter."
        ) from e

    root = tk.Tk()
    root.withdraw()
    root.update()

    # DETAILS
    dlg = tk.Toplevel(root)
    dlg.title("Spec Sheet Info")
    dlg.resizable(False, False)

    frm = tk.Frame(dlg, padx=12, pady=12)
    frm.pack(fill="both", expand=True)

    tk.Label(frm, text="Customer:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=(0, 8))
    customer_var = tk.StringVar(value=args.customer or "")
    customer_entry = tk.Entry(frm, width=42, textvariable=customer_var)
    customer_entry.grid(row=0, column=1, sticky="w", pady=(0, 8))

    tk.Label(frm, text="Vehicle:").grid(row=1, column=0, sticky="e", padx=(0, 8))
    vehicle_var = tk.StringVar(value=args.vehicle or "MIXER")
    vehicle_entry = tk.Entry(frm, width=42, textvariable=vehicle_var)
    vehicle_entry.grid(row=1, column=1, sticky="w")

    btns = tk.Frame(frm, pady=12)
    btns.grid(row=2, column=0, columnspan=2, sticky="e")

    result = {"ok": False}

    def _ok():
        result["ok"] = True
        dlg.destroy()

    def _cancel():
        dlg.destroy()

    tk.Button(btns, text="Cancel", width=10, command=_cancel).pack(side="right", padx=(8, 0))
    tk.Button(btns, text="OK", width=10, command=_ok).pack(side="right")

    dlg.bind("<Return>", lambda _e: _ok())
    dlg.bind("<Escape>", lambda _e: _cancel())

    dlg.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() // 2) - (dlg.winfo_width() // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (dlg.winfo_height() // 2)
    dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    customer_entry.focus_set()
    dlg.grab_set()
    root.wait_window(dlg)

    if not result["ok"]:
        raise SystemExit(1)

    args.customer = customer_var.get().strip() or ""
    args.vehicle = vehicle_var.get().strip() or "MIXER"

    # PROOF
    proof = filedialog.askopenfilename(
        title="Select mockup/proof image",
        filetypes=[("Images", "*.png *.jpg *.jpeg *.tif *.tiff"), ("All files", "*.*")],
    )
    if not proof:
        raise SystemExit(1)
    args.proof = Path(proof)

    # PRINT FOLDER (default to folder containing the selected proof)
    initial_print_dir = str(Path(args.proof).parent)
    print_dir = filedialog.askdirectory(title="Select PRINT folder", initialdir=initial_print_dir)
    if not print_dir:
        raise SystemExit(1)
    args.print_dir = Path(print_dir)
    args.print_zip = None

    # Font
    if args.font and Path(args.font).exists():
        args.font = Path(args.font)
    else:
        font = filedialog.askopenfilename(
            title="Select HudsonNY font file",
            filetypes=[("Font files", "*.ttf *.otf"), ("All files", "*.*")],
        )
        if not font:
            raise SystemExit(1)
        args.font = Path(font)

    # OUTPUT
    today = datetime.date.today().isoformat()
    suggested = f"{today}__{safe_filename(args.customer)}__{safe_filename(args.vehicle)}__Spec_Sheet.pdf"
    initial_save_dir = Path(args.proof).parent
    out = filedialog.asksaveasfilename(
        title="Save spec sheet PDF as...",
        initialdir=str(initial_save_dir),
        initialfile=suggested,
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
    )
    if not out:
        raise SystemExit(1)
    args.out = Path(out)

    root.destroy()
    return args


def main() -> int:
    ap = argparse.ArgumentParser(description="Generate fleet instruction/spec sheet PDF (CLI or GUI).")
    ap.add_argument("--gui", action="store_true", help="Use file-picker GUI (no CLI paths needed).")

    src = ap.add_mutually_exclusive_group(required=False)
    src.add_argument("--print-zip", type=Path, help="Path to PRINT.zip")
    src.add_argument("--print-dir", type=Path, help="Path to PRINT folder")

    ap.add_argument("--proof", type=Path, help="Proof image (png/jpg)")
    # If you keep the font in assets/HudsonNY.ttf, you won't be prompted each run.
    ap.add_argument("--font", type=Path, default=Path("assets/HudsonNY.ttf"), help="HudsonNY.ttf or .otf")
    ap.add_argument("--out", type=Path, help="Output PDF path")

    ap.add_argument("--customer", default="", help="Customer name")
    ap.add_argument("--vehicle", default="MIXER TRUCK", help="Vehicle type label")

    args = ap.parse_args()

    if args.gui:
        args = run_gui_defaults(args)
    else:
        # CLI required fields
        if not (args.print_zip or args.print_dir):
            ap.error("CLI mode requires --print-zip or --print-dir (or use --gui).")
        if not args.proof:
            ap.error("CLI mode requires --proof (or use --gui).")
        if not args.font:
            ap.error("CLI mode requires --font (or use --gui).")
        if not args.out:
            ap.error("CLI mode requires --out (or use --gui).")

    # Validate files
    if args.print_zip and not args.print_zip.exists():
        raise FileNotFoundError(f"PRINT.zip not found: {args.print_zip}")
    if args.print_dir and not args.print_dir.exists():
        raise FileNotFoundError(f"PRINT folder not found: {args.print_dir}")
    if not args.proof.exists():
        raise FileNotFoundError(f"Proof not found: {args.proof}")
    if not args.font.exists():
        raise FileNotFoundError(f"Font not found: {args.font}")

    with tempfile.TemporaryDirectory(prefix="fleet_spec_") as td:
        work = Path(td)
        thumbs_dir = work / "thumbs"
        thumbs_dir.mkdir(parents=True, exist_ok=True)

        print_root = extract_print_source(args.print_zip, args.print_dir, work)
        items = collect_pdf_items(print_root, thumbs_dir)
        if not items:
            raise RuntimeError("No PDFs found in PRINT after filtering.")

        args.out.parent.mkdir(parents=True, exist_ok=True)

        build_pdf(
            out_pdf=args.out,
            proof_image=args.proof,
            font_path=args.font,
            customer=args.customer,
            vehicle_type=args.vehicle,
            items=items,
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())