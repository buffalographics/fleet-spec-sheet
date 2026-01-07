from pathlib import Path
import re, shutil, zipfile
from os.path import commonpath
import fitz
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Image, Spacer

FONT_PATH = Path(__file__).resolve().parents[2] / "assets" / "fonts" / "HudsonNY.ttf"

def parse_qty(name):
    for pat in [r"qty(\d+)", r"x(\d+)", r"(\d+)x"]:
        m = re.search(pat, name, re.IGNORECASE)
        if m:
            return int(m.group(1))
    return 1

def generate_spec_sheet(proof, print_zip, out_pdf, customer, job_no, unit_no, vehicle):
    pdfmetrics.registerFont(TTFont("HudsonNY", str(FONT_PATH)))

    work = out_pdf.parent / (out_pdf.stem + "_work")
    if work.exists():
        shutil.rmtree(work)
    unzip_dir = work / "unzipped"
    thumb_dir = work / "thumbs"
    unzip_dir.mkdir(parents=True)
    thumb_dir.mkdir(parents=True)

    with zipfile.ZipFile(print_zip) as z:
        z.extractall(unzip_dir)

    pdfs = [p for p in unzip_dir.rglob("*.pdf")]
    root = Path(commonpath([str(p.parent) for p in pdfs]))

    items = []
    for p in pdfs:
        rel = str(p.relative_to(root)).replace("\\", "/")
        items.append((p, rel, parse_qty(p.name)))

    thumbs = []
    for i,(pdf,_,_) in enumerate(items,1):
        d = fitz.open(pdf)
        pix = d.load_page(0).get_pixmap(matrix=fitz.Matrix(2,2))
        tp = thumb_dir / f"t{i}.png"
        pix.save(tp)
        d.close()
        thumbs.append(tp)

    styles = getSampleStyleSheet()
    title = ParagraphStyle("t", fontName="HudsonNY", fontSize=18)
    section = ParagraphStyle("s", fontName="HudsonNY", fontSize=11)
    body = ParagraphStyle("b", fontSize=9)

    doc = SimpleDocTemplate(str(out_pdf), pagesize=letter,
        leftMargin=0.5*inch, rightMargin=0.5*inch,
        topMargin=0.5*inch, bottomMargin=0.5*inch)

    story = []
    bar = Table([["BUFFALO GRAPHICS COMPANY"]],[7.5*inch])
    bar.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.black),
                             ("TEXTCOLOR",(0,0),(-1,-1),colors.white),
                             ("FONTNAME",(0,0),(-1,-1),"HudsonNY"),
                             ("FONTSIZE",(0,0),(-1,-1),14)]))
    story.append(bar)
    story.append(Paragraph("FLEET VEHICLE INSTRUCTION / SPEC SHEET (V1)", title))
    story.append(Paragraph("INSTALL REFERENCE — NOT TO SCALE", section))
    pil = PILImage.open(proof)
    w,h = pil.size
    story.append(Image(str(proof), 7.5*inch, h*(7.5*inch/w)))
    doc.build(story)
