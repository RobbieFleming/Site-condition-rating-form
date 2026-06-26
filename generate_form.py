#!/usr/bin/env python3
"""Generate NZGTTM Site Assurance Form as a Word .docx document."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── PALETTE ──────────────────────────────────────────────────────────────────
NAVY     = "00274D"
ORANGE   = "E07728"
RED_HDR  = "7A0E0E"
RED      = "C0392B"
L_BLUE   = "E8F0FA"
GRAY     = "F5F6F8"
WHITE    = "FFFFFF"
BORDER   = "C0C6D4"
TEXT     = "1C1C1E"
MUTED    = "555555"

CB = "☐"   # unchecked box — check by hand or replace with ☑ digitally

# ── HELPERS ──────────────────────────────────────────────────────────────────

def set_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_table_borders(table, color=BORDER, sz=4):
    tbl = table._tbl
    tblPr = tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:color"), color)
        b.append(el)
    tblPr.append(b)

def no_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:color"), "auto")
        b.append(el)
    tblPr.append(b)

def cell_margins(cell, top=0.05, bottom=0.05, left=0.12, right=0.12):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(int(val * 567)))   # twips
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)

def fmt(para, space_before=0, space_after=0, align=WD_ALIGN_PARAGRAPH.LEFT):
    para.paragraph_format.space_before = Pt(space_before)
    para.paragraph_format.space_after  = Pt(space_after)
    para.alignment = align

def rn(para, text, bold=False, size=8, color=TEXT, italic=False, font="Arial", fg=None):
    r = para.add_run(text)
    r.bold   = bold
    r.italic = italic
    r.font.name = font
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor.from_string(fg if fg is not None else color)
    return r

def write_cell(cell, text, bold=False, size=8, fg=TEXT, bg=None,
               align=WD_ALIGN_PARAGRAPH.LEFT, valign=WD_ALIGN_VERTICAL.CENTER,
               italic=False):
    if bg:
        set_bg(cell, bg)
    cell.vertical_alignment = valign
    p = cell.paragraphs[0]
    fmt(p, align=align)
    rn(p, text, bold=bold, size=size, color=fg, italic=italic)
    cell_margins(cell)

def section_header(table_or_doc, letter, title, badge_bg=ORANGE, header_bg=NAVY,
                   note=None, full_width_cm=18.4):
    """Add a section header row or standalone header table."""
    t = table_or_doc.add_table(rows=1, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    no_borders(t)
    row = t.rows[0]
    # Badge cell
    row.cells[0].merge(row.cells[0])
    set_bg(row.cells[0], badge_bg)
    p0 = row.cells[0].paragraphs[0]
    fmt(p0, align=WD_ALIGN_PARAGRAPH.CENTER)
    rn(p0, letter, bold=True, size=9, fg=WHITE)
    row.cells[0].width = Cm(0.65)
    row.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Title cell
    set_bg(row.cells[1], header_bg)
    p1 = row.cells[1].paragraphs[0]
    fmt(p1)
    rn(p1, title.upper(), bold=True, size=8.5, fg=WHITE)
    row.cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell_margins(row.cells[1], left=0.15)

    # Note cell (right-aligned hint)
    set_bg(row.cells[2], header_bg)
    row.cells[2].width = Cm(4.5)
    p2 = row.cells[2].paragraphs[0]
    fmt(p2, align=WD_ALIGN_PARAGRAPH.RIGHT)
    if note:
        rn(p2, note, size=6.5, fg="AACCEE", italic=True)
    row.cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cell_margins(row.cells[2], right=0.15)
    return t

def checklist_table(doc_or_cell, items, col1_label="Checklist Item",
                    yes_label="YES", no_label="NO", na_label="N/A",
                    header_bg=L_BLUE, yes_color=TEXT, full_width=True):
    """Create a Yes/No/N/A checklist table and add to doc_or_cell."""
    t = doc_or_cell.add_table(rows=1 + len(items), cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_borders(t)

    # Header row
    hdr = t.rows[0]
    hdrs = [(col1_label, WD_ALIGN_PARAGRAPH.LEFT),
            (yes_label,  WD_ALIGN_PARAGRAPH.CENTER),
            (no_label,   WD_ALIGN_PARAGRAPH.CENTER),
            (na_label,   WD_ALIGN_PARAGRAPH.CENTER)]
    for i, (lbl, al) in enumerate(hdrs):
        set_bg(hdr.cells[i], header_bg)
        p = hdr.cells[i].paragraphs[0]
        fmt(p, align=al)
        fc = RED if (i == 1 and yes_label == "YES\n(Risk Present)") else NAVY
        rn(p, lbl, bold=True, size=7, fg=fc)
        cell_margins(hdr.cells[i])
        hdr.cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Data rows
    for r_idx, item in enumerate(items):
        row = t.rows[r_idx + 1]
        bg = WHITE if r_idx % 2 == 0 else "FAFBFD"
        # Item cell
        set_bg(row.cells[0], bg)
        p = row.cells[0].paragraphs[0]
        fmt(p)
        rn(p, item, size=7.5, fg=TEXT)
        cell_margins(row.cells[0], left=0.15)
        row.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        # CB cells
        for ci in (1, 2, 3):
            set_bg(row.cells[ci], bg)
            p = row.cells[ci].paragraphs[0]
            fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
            rn(p, CB, size=9, fg="333333", font="Segoe UI Symbol")
            cell_margins(row.cells[ci], left=0.05, right=0.05)
            row.cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Column widths
    check_w = Cm(1.2)
    for row in t.rows:
        for ci in (1, 2, 3):
            row.cells[ci].width = check_w
    return t

def spacer(doc, height_pt=4):
    p = doc.add_paragraph()
    fmt(p, space_before=0, space_after=0)
    p.paragraph_format.line_spacing = Pt(height_pt)

def line_field(para, label, line_len_cm=4.5, label_size=7, bold_label=True):
    rn(para, label + "  ", bold=bold_label, size=label_size, fg=MUTED)
    rn(para, "_" * int(line_len_cm * 5.5), size=7, fg="888888")
    rn(para, "    ", size=7, fg=TEXT)

def compliance_row(doc):
    """Add the 4-option compliance classification row."""
    t = doc.add_table(rows=1, cols=5)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    no_borders(t)
    set_bg(t.rows[0].cells[0], GRAY)
    p0 = t.rows[0].cells[0].paragraphs[0]
    fmt(p0, align=WD_ALIGN_PARAGRAPH.LEFT)
    rn(p0, "Section Assessment: ", bold=True, size=7, fg=MUTED)
    t.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    t.rows[0].cells[0].width = Cm(3.5)

    opts = [
        ("Compliant",              "1B6B3A"),
        ("Minor Non-Conformance",  "7B5800"),
        ("Major Non-Conformance",  "B84700"),
        ("Critical Risk / Stop Work", RED),
    ]
    for i, (lbl, col) in enumerate(opts, start=1):
        set_bg(t.rows[0].cells[i], GRAY)
        p = t.rows[0].cells[i].paragraphs[0]
        fmt(p)
        rn(p, CB + "  ", size=9, fg="333333", font="Segoe UI Symbol")
        rn(p, lbl, bold=True, size=7, fg=col)
        cell_margins(t.rows[0].cells[i], left=0.1)
        t.rows[0].cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    return t


# ── BUILD DOCUMENT ───────────────────────────────────────────────────────────

doc = Document()

# Page setup — A4, tight margins
sec = doc.sections[0]
sec.page_width    = Cm(21.0)
sec.page_height   = Cm(29.7)
sec.left_margin   = Cm(1.3)
sec.right_margin  = Cm(1.3)
sec.top_margin    = Cm(1.3)
sec.bottom_margin = Cm(1.3)

# Default style
style = doc.styles["Normal"]
style.font.name = "Arial"
style.font.size = Pt(8)
style.paragraph_format.space_before = Pt(0)
style.paragraph_format.space_after  = Pt(0)

FULL_W = Cm(18.4)   # usable width

# ── HEADER ───────────────────────────────────────────────────────────────────
hdr_t = doc.add_table(rows=1, cols=3)
hdr_t.alignment = WD_TABLE_ALIGNMENT.LEFT
no_borders(hdr_t)

# Logo cell
lc = hdr_t.rows[0].cells[0]
lc.width = Cm(3.8)
set_bg(lc, NAVY)
lc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
cell_margins(lc, left=0.25, top=0.1, bottom=0.1)
p = lc.paragraphs[0]
fmt(p)
rn(p, "NZGTTM", bold=True, size=16, fg=WHITE)
p2 = lc.add_paragraph()
fmt(p2)
rn(p2, "New Zealand Guide to TTM", size=5.5, fg=ORANGE)

# Title cell
tc = hdr_t.rows[0].cells[1]
set_bg(tc, NAVY)
tc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
cell_margins(tc, left=0.25, top=0.1, bottom=0.1)
p = tc.paragraphs[0]
fmt(p)
rn(p, "Site Assurance and Audit Form", bold=True, size=13, fg=WHITE)
p2 = tc.add_paragraph()
fmt(p2)
rn(p2, "Risk-Based Assessment  ·  Safe System Approach  ·  Replaces CoPTTM Site Condition Rating (SCR) Form", size=6.5, fg="90B4D8")
p3 = tc.add_paragraph()
fmt(p3)
rn(p3, "For use by: TMO  ·  STMS  ·  Corridor Managers  ·  RCA Auditors", size=6, fg="7090B0", italic=True)

# Ref cell
rc = hdr_t.rows[0].cells[2]
rc.width = Cm(2.6)
set_bg(rc, ORANGE)
rc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
cell_margins(rc, top=0.1, bottom=0.1)
p = rc.paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p, "Form Ref", size=5.5, fg="FFFFFF")
p2 = rc.add_paragraph()
fmt(p2, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p2, "SAF-1.0", bold=True, size=12, fg=WHITE)
p3 = rc.add_paragraph()
fmt(p3, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p3, "NZGTTM", size=5.5, fg="FFFFFF")

# Safe System strip
spacer(doc, 2)
strip = doc.add_table(rows=1, cols=1)
no_borders(strip)
set_bg(strip.rows[0].cells[0], "1A3A6B")
p = strip.rows[0].cells[0].paragraphs[0]
fmt(p)
cell_margins(strip.rows[0].cells[0], left=0.2, top=0.05, bottom=0.05)
rn(p, "SAFE SYSTEM PRINCIPLES  ", bold=True, size=6, fg="90B4D8")
for item in ["Safe Roads & Roadsides", "Safe Speeds", "Safe Vehicles",
             "Safe People", "Vulnerable Road Users", "Continuous Improvement"]:
    rn(p, f"·  {item}  ", size=6, fg="AABBCC")

spacer(doc, 3)

# ── SECTION A: SITE DETAILS ───────────────────────────────────────────────────
section_header(doc, "A", "Site Details", note="Complete all fields before commencing audit")
spacer(doc, 1)

det = doc.add_table(rows=3, cols=4)
det.alignment = WD_TABLE_ALIGNMENT.LEFT
set_table_borders(det)
det.rows[0].cells[0].width = Cm(4.6)

# Row 1
r0 = det.rows[0]
for cell, lbl, width in [
    (r0.cells[0], "RCA:", 4.6),
    (r0.cells[1], "Project Name:", 4.6),
    (r0.cells[2], "TGS / TMP No.:", 4.6),
    (r0.cells[3], "Date / Time:", 4.6),
]:
    cell_margins(cell, left=0.12, top=0.08, bottom=0.08)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    fmt(p)
    rn(p, lbl + "  ", bold=True, size=6.5, fg=MUTED)

# Row 2
r1 = det.rows[1]
r1.cells[0].merge(r1.cells[1])
for cell, lbl in [
    (r1.cells[0], "Road Name / Location:"),
    (r1.cells[2], "STMS Name:"),
    (r1.cells[3], "Auditor Name:"),
]:
    cell_margins(cell, left=0.12, top=0.08, bottom=0.08)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    fmt(p)
    rn(p, lbl + "  ", bold=True, size=6.5, fg=MUTED)

# Row 3
r2 = det.rows[2]
for cell, lbl in [
    (r2.cells[0], "GPS Coordinates:"),
    (r2.cells[1], "Company / Organisation:"),
    (r2.cells[2], "Contact Number:"),
    (r2.cells[3], "Weather Conditions:"),
]:
    cell_margins(cell, left=0.12, top=0.08, bottom=0.08)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    fmt(p)
    rn(p, lbl + "  ", bold=True, size=6.5, fg=MUTED)

spacer(doc, 4)

# ── SECTIONS B–G  (2-column outer table) ─────────────────────────────────────
outer = doc.add_table(rows=1, cols=2)
outer.alignment = WD_TABLE_ALIGNMENT.LEFT
no_borders(outer)
outer.rows[0].cells[0].width = Cm(9.1)
outer.rows[0].cells[1].width = Cm(9.1)

left_cell  = outer.rows[0].cells[0]
right_cell = outer.rows[0].cells[1]

for cell in (left_cell, right_cell):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    cell_margins(cell, left=0, right=0, top=0, bottom=0)

# ── LEFT COLUMN: B, D, F ──────────────────────────────────────────────────────

# Section B — Immediate Safety Risks
sh_b = left_cell.add_table(rows=1, cols=3)
no_borders(sh_b)
set_bg(sh_b.rows[0].cells[0], RED)
p = sh_b.rows[0].cells[0].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p, "B", bold=True, size=9, fg=WHITE)
sh_b.rows[0].cells[0].width = Cm(0.65)
sh_b.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_b.rows[0].cells[1], RED_HDR)
p = sh_b.rows[0].cells[1].paragraphs[0]
fmt(p)
rn(p, "IMMEDIATE SAFETY RISKS", bold=True, size=8.5, fg=WHITE)
cell_margins(sh_b.rows[0].cells[1], left=0.15)
sh_b.rows[0].cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_b.rows[0].cells[2], RED_HDR)
sh_b.rows[0].cells[2].width = Cm(3.0)
p = sh_b.rows[0].cells[2].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.RIGHT)
rn(p, "YES = Risk Present ⚠", size=6.5, fg="FFCCCC", italic=True)
cell_margins(sh_b.rows[0].cells[2], right=0.1)
sh_b.rows[0].cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

# Alert bar
alert_t = left_cell.add_table(rows=1, cols=1)
no_borders(alert_t)
set_bg(alert_t.rows[0].cells[0], "FFF3F3")
p = alert_t.rows[0].cells[0].paragraphs[0]
fmt(p)
rn(p, "⚠  Any YES response requires immediate corrective action — consider Stop Work order",
   bold=True, size=6.5, fg="8B0000")
cell_margins(alert_t.rows[0].cells[0], left=0.12, top=0.05, bottom=0.05)

b_items = [
    "Workers exposed to live traffic without adequate protection",
    "Vehicle intrusion risk into the work area",
    "Inadequate physical separation — workers and traffic",
    "Unsafe pedestrian route at or through the worksite",
    "Unsafe cyclist route or accommodation at the worksite",
    "Visibility issues — signs or work area not clearly visible",
    "Queue management issues — queues beyond safe sight distance",
    "Plant or equipment operating unsafely near live traffic",
    "Emergency access compromised — routes blocked or impeded",
]
checklist_table(left_cell, b_items,
                col1_label="Safety Risk Item",
                yes_label="YES\n(Risk Present)",
                yes_color=RED)

# Section D — Vulnerable Road Users
spacer_p = left_cell.add_paragraph()
fmt(spacer_p, space_before=3)

sh_d = left_cell.add_table(rows=1, cols=3)
no_borders(sh_d)
set_bg(sh_d.rows[0].cells[0], ORANGE)
p = sh_d.rows[0].cells[0].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p, "D", bold=True, size=9, fg=WHITE)
sh_d.rows[0].cells[0].width = Cm(0.65)
sh_d.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_d.rows[0].cells[1], NAVY)
p = sh_d.rows[0].cells[1].paragraphs[0]
fmt(p)
rn(p, "VULNERABLE ROAD USERS", bold=True, size=8.5, fg=WHITE)
cell_margins(sh_d.rows[0].cells[1], left=0.15)
sh_d.rows[0].cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_d.rows[0].cells[2], NAVY)
sh_d.rows[0].cells[2].width = Cm(3.0)
p = sh_d.rows[0].cells[2].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.RIGHT)
rn(p, "YES = Compliant", size=6.5, fg="AACCEE", italic=True)
cell_margins(sh_d.rows[0].cells[2], right=0.1)
sh_d.rows[0].cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

d_items = [
    "Continuous pedestrian route provided through/around site",
    "Pedestrian route accessible — adequate width, surface and lit",
    "Cyclist accommodation provided and appropriate for speed",
    "Pedestrian crossing facilities adequate and clearly signed",
    "Mobility-impaired users considered — kerb ramps, tactile pads",
]
checklist_table(left_cell, d_items)

# Section F — Worksite Safety
spacer_p = left_cell.add_paragraph()
fmt(spacer_p, space_before=3)

sh_f = left_cell.add_table(rows=1, cols=3)
no_borders(sh_f)
set_bg(sh_f.rows[0].cells[0], ORANGE)
p = sh_f.rows[0].cells[0].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p, "F", bold=True, size=9, fg=WHITE)
sh_f.rows[0].cells[0].width = Cm(0.65)
sh_f.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_f.rows[0].cells[1], NAVY)
p = sh_f.rows[0].cells[1].paragraphs[0]
fmt(p)
rn(p, "WORKSITE SAFETY", bold=True, size=8.5, fg=WHITE)
cell_margins(sh_f.rows[0].cells[1], left=0.15)
sh_f.rows[0].cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_f.rows[0].cells[2], NAVY)
sh_f.rows[0].cells[2].width = Cm(3.0)
p = sh_f.rows[0].cells[2].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.RIGHT)
rn(p, "YES = Compliant", size=6.5, fg="AACCEE", italic=True)
cell_margins(sh_f.rows[0].cells[2], right=0.1)
sh_f.rows[0].cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

f_items = [
    "Crew briefed — TM plan, hazards and emergency procedures",
    "All personnel wearing correct PPE (hi-viz, footwear, head)",
    "Plant and equipment operating safely — exclusion zones set",
    "Materials and equipment secured — no risk into traffic lanes",
    "Work area housekeeping acceptable — debris managed",
    "Emergency procedures understood by all site personnel",
]
checklist_table(left_cell, f_items)

# ── RIGHT COLUMN: C, E, G ─────────────────────────────────────────────────────
cell_margins(right_cell, left=0.2, right=0, top=0, bottom=0)

# Section C — TM Layout
sh_c = right_cell.add_table(rows=1, cols=3)
no_borders(sh_c)
set_bg(sh_c.rows[0].cells[0], ORANGE)
p = sh_c.rows[0].cells[0].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p, "C", bold=True, size=9, fg=WHITE)
sh_c.rows[0].cells[0].width = Cm(0.65)
sh_c.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_c.rows[0].cells[1], NAVY)
p = sh_c.rows[0].cells[1].paragraphs[0]
fmt(p)
rn(p, "TRAFFIC MANAGEMENT LAYOUT", bold=True, size=8.5, fg=WHITE)
cell_margins(sh_c.rows[0].cells[1], left=0.15)
sh_c.rows[0].cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_c.rows[0].cells[2], NAVY)
sh_c.rows[0].cells[2].width = Cm(3.0)
p = sh_c.rows[0].cells[2].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.RIGHT)
rn(p, "YES = Compliant", size=6.5, fg="AACCEE", italic=True)
cell_margins(sh_c.rows[0].cells[2], right=0.1)
sh_c.rows[0].cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

c_items = [
    "Site layout matches the approved TGS / TMP",
    "Signs correctly positioned per approved plan",
    "Signs visible, legible, clean and unobstructed",
    "Delineation (cones, drums, barriers) adequate and spaced",
    "Tapers appropriate for speed environment and site conditions",
    "Work area clearly defined and separated from traffic",
    "Traffic lanes clearly guided through the worksite",
    "Temporary speed management appropriate and compliant",
    "Access points controlled and appropriately managed",
]
checklist_table(right_cell, c_items)

# Section E — Traffic Operations
spacer_p = right_cell.add_paragraph()
fmt(spacer_p, space_before=3)

sh_e = right_cell.add_table(rows=1, cols=3)
no_borders(sh_e)
set_bg(sh_e.rows[0].cells[0], ORANGE)
p = sh_e.rows[0].cells[0].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p, "E", bold=True, size=9, fg=WHITE)
sh_e.rows[0].cells[0].width = Cm(0.65)
sh_e.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_e.rows[0].cells[1], NAVY)
p = sh_e.rows[0].cells[1].paragraphs[0]
fmt(p)
rn(p, "TRAFFIC OPERATIONS", bold=True, size=8.5, fg=WHITE)
cell_margins(sh_e.rows[0].cells[1], left=0.15)
sh_e.rows[0].cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_e.rows[0].cells[2], NAVY)
sh_e.rows[0].cells[2].width = Cm(3.0)
p = sh_e.rows[0].cells[2].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.RIGHT)
rn(p, "YES = Compliant", size=6.5, fg="AACCEE", italic=True)
cell_margins(sh_e.rows[0].cells[2], right=0.1)
sh_e.rows[0].cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

e_items = [
    "Traffic flow operating safely through the worksite",
    "Driver behaviour appropriate — no unsafe manoeuvres observed",
    "Traffic delays are within acceptable limits for site/road type",
    "Queue lengths being managed effectively",
    "No evidence of driver confusion at merge or lane change points",
    "Sight distances adequate — approach, through and exit zones",
]
checklist_table(right_cell, e_items)

# Section G — Documentation
spacer_p = right_cell.add_paragraph()
fmt(spacer_p, space_before=3)

sh_g = right_cell.add_table(rows=1, cols=3)
no_borders(sh_g)
set_bg(sh_g.rows[0].cells[0], ORANGE)
p = sh_g.rows[0].cells[0].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p, "G", bold=True, size=9, fg=WHITE)
sh_g.rows[0].cells[0].width = Cm(0.65)
sh_g.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_g.rows[0].cells[1], NAVY)
p = sh_g.rows[0].cells[1].paragraphs[0]
fmt(p)
rn(p, "DOCUMENTATION", bold=True, size=8.5, fg=WHITE)
cell_margins(sh_g.rows[0].cells[1], left=0.15)
sh_g.rows[0].cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_g.rows[0].cells[2], NAVY)
sh_g.rows[0].cells[2].width = Cm(3.0)
p = sh_g.rows[0].cells[2].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.RIGHT)
rn(p, "YES = Available & Current", size=6.5, fg="AACCEE", italic=True)
cell_margins(sh_g.rows[0].cells[2], right=0.1)
sh_g.rows[0].cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

g_items = [
    "Current approved TGS / TMP available at the worksite",
    "All required permits available and current",
    "RCA approvals and consent conditions available on site",
    "Daily TM records and site logs completed and up to date",
    "Changes to TM layout documented and approved",
    "Current risk assessment available and reflects site conditions",
]
checklist_table(right_cell, g_items)

spacer(doc, 4)

# ── SECTION H: NON-CONFORMANCES ───────────────────────────────────────────────
section_header(doc, "H", "Non-Conformances Identified",
               note="Complete for every NO response in Sections B – G")
spacer(doc, 1)

nc = doc.add_table(rows=6, cols=7)
nc.alignment = WD_TABLE_ALIGNMENT.LEFT
set_table_borders(nc)

# Header
nc_hdrs = ["#", "Description of Non-Conformance",
           "Risk Level\n(Low/Med/High/Critical)",
           "Corrective Action Required",
           "Responsible Person", "Due Date", "Closed ✓"]
nc_widths = [Cm(0.7), Cm(5.2), Cm(2.3), Cm(4.5), Cm(2.8), Cm(1.6), Cm(1.3)]
hdr_row = nc.rows[0]
for i, (h, w) in enumerate(zip(nc_hdrs, nc_widths)):
    set_bg(hdr_row.cells[i], L_BLUE)
    hdr_row.cells[i].width = w
    p = hdr_row.cells[i].paragraphs[0]
    fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
    rn(p, h, bold=True, size=6.5, fg=NAVY)
    cell_margins(hdr_row.cells[i], top=0.08, bottom=0.08)
    hdr_row.cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

# Data rows
for r_idx in range(1, 6):
    row = nc.rows[r_idx]
    bg = WHITE if r_idx % 2 != 0 else "FAFBFD"
    set_bg(row.cells[0], GRAY)
    p = row.cells[0].paragraphs[0]
    fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
    rn(p, str(r_idx), bold=True, size=8, fg=TEXT)
    row.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for ci in range(1, 7):
        set_bg(row.cells[ci], bg)
        row.cells[ci].width = nc_widths[ci]
    # Set row height
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trH = OxmlElement("w:trHeight")
    trH.set(qn("w:val"), "567")   # ~1cm
    trPr.append(trH)

spacer(doc, 4)

# ── SECTIONS I + J (side by side) ────────────────────────────────────────────
ij = doc.add_table(rows=1, cols=2)
ij.alignment = WD_TABLE_ALIGNMENT.LEFT
no_borders(ij)
ij.rows[0].cells[0].width = Cm(9.0)
ij.rows[0].cells[1].width = Cm(9.4)
icell = ij.rows[0].cells[0]
jcell = ij.rows[0].cells[1]
for c in (icell, jcell):
    c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    cell_margins(c, left=0, right=0, top=0, bottom=0)
cell_margins(jcell, left=0.2)

# Section I — Overall Assessment
sh_i = icell.add_table(rows=1, cols=3)
no_borders(sh_i)
set_bg(sh_i.rows[0].cells[0], ORANGE)
p = sh_i.rows[0].cells[0].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p, "I", bold=True, size=9, fg=WHITE)
sh_i.rows[0].cells[0].width = Cm(0.65)
sh_i.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_i.rows[0].cells[1], NAVY)
p = sh_i.rows[0].cells[1].paragraphs[0]
fmt(p)
rn(p, "OVERALL SITE ASSESSMENT", bold=True, size=8.5, fg=WHITE)
cell_margins(sh_i.rows[0].cells[1], left=0.15)
sh_i.rows[0].cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_i.rows[0].cells[2], NAVY)
sh_i.rows[0].cells[2].width = Cm(1.5)
icell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

# Rating options
rating_t = icell.add_table(rows=4, cols=1)
no_borders(rating_t)
ratings = [
    ("Compliant", "1B5E20", "F0FFF4", "2E7D32"),
    ("Minor Non-Conformance", "7B5800", "FFFDE7", "F9A825"),
    ("Major Non-Conformance", "B71C1C", "FFF8F0", "E64A19"),
    ("Critical Risk / Stop Work", "8B0000", "FFF5F5", RED),
]
for i, (lbl, txt_col, bg_col, border_col) in enumerate(ratings):
    row = rating_t.rows[i]
    set_bg(row.cells[0], bg_col)
    p = row.cells[0].paragraphs[0]
    fmt(p)
    rn(p, CB + "  ", size=10, fg="333333", font="Segoe UI Symbol")
    rn(p, lbl, bold=True, size=8, fg=txt_col)
    cell_margins(row.cells[0], left=0.15, top=0.06, bottom=0.06)

p_lbl = icell.add_paragraph()
fmt(p_lbl, space_before=3)
rn(p_lbl, "AUDITOR COMMENTS / IMMEDIATE ACTIONS TAKEN:", bold=True, size=6.5, fg=NAVY)

# Comment lines
comments_t = icell.add_table(rows=4, cols=1)
set_table_borders(comments_t, color=BORDER, sz=2)
for row in comments_t.rows:
    set_bg(row.cells[0], WHITE)
    row.cells[0].paragraphs[0].paragraph_format.space_before = Pt(0)
    row.cells[0].paragraphs[0].paragraph_format.space_after  = Pt(0)
    cell_margins(row.cells[0], top=0.06, bottom=0.06, left=0.12)
    trPr = row._tr.get_or_add_trPr()
    trH  = OxmlElement("w:trHeight")
    trH.set(qn("w:val"), "425")
    trPr.append(trH)

p_ri = icell.add_paragraph()
fmt(p_ri, space_before=3)
rn(p_ri, "Re-inspection required:  ", bold=True, size=7, fg=NAVY)
rn(p_ri, CB + " Yes    ", size=9, fg="333333", font="Segoe UI Symbol")
rn(p_ri, CB + " No     ", size=9, fg="333333", font="Segoe UI Symbol")
rn(p_ri, "By: ________________________", size=7, fg=MUTED)

# Section J — Sign-offs
sh_j = jcell.add_table(rows=1, cols=3)
no_borders(sh_j)
set_bg(sh_j.rows[0].cells[0], ORANGE)
p = sh_j.rows[0].cells[0].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p, "J", bold=True, size=9, fg=WHITE)
sh_j.rows[0].cells[0].width = Cm(0.65)
sh_j.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_j.rows[0].cells[1], NAVY)
p = sh_j.rows[0].cells[1].paragraphs[0]
fmt(p)
rn(p, "SIGN-OFFS", bold=True, size=8.5, fg=WHITE)
cell_margins(sh_j.rows[0].cells[1], left=0.15)
sh_j.rows[0].cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
set_bg(sh_j.rows[0].cells[2], NAVY)
sh_j.rows[0].cells[2].width = Cm(1.5)

so_t = jcell.add_table(rows=1, cols=2)
no_borders(so_t)
for ci, (title, fields) in enumerate([
    ("Auditor / Inspector",     ["Name:", "Role / Title:", "Company:", "Date:"]),
    ("STMS / Site Representative", ["Name:", "STMS Cert No.:", "Company:", "Date:"]),
]):
    c = so_t.rows[0].cells[ci]
    c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    cell_margins(c, left=0.1, right=0.1)
    p_title = c.paragraphs[0]
    fmt(p_title)
    rn(p_title, title.upper(), bold=True, size=7.5, fg=NAVY)
    # Underline via border hack: add a bottom-bordered paragraph
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    pPr = p_title._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "8")
    bot.set(qn("w:color"), NAVY)
    pBdr.append(bot)
    pPr.append(pBdr)

    for field in fields:
        pf = c.add_paragraph()
        fmt(pf, space_before=1)
        rn(pf, field + "  ", bold=True, size=6.5, fg=MUTED)
        rn(pf, "_" * 28, size=7, fg="999999")

    # Signature box (simulated with a bordered table)
    sig_t = c.add_table(rows=1, cols=1)
    set_table_borders(sig_t, color="AAAAAA", sz=4)
    set_bg(sig_t.rows[0].cells[0], WHITE)
    p_sig = sig_t.rows[0].cells[0].paragraphs[0]
    fmt(p_sig)
    rn(p_sig, "Signature", size=6, fg="BBBBBB", italic=True)
    cell_margins(sig_t.rows[0].cells[0], left=0.1, top=0.05)
    trPr = sig_t.rows[0]._tr.get_or_add_trPr()
    trH  = OxmlElement("w:trHeight")
    trH.set(qn("w:val"), "992")   # ~1.75cm
    trPr.append(trH)

spacer(doc, 4)

# ── LEGEND BAR ────────────────────────────────────────────────────────────────
leg_t = doc.add_table(rows=1, cols=1)
no_borders(leg_t)
set_bg(leg_t.rows[0].cells[0], GRAY)
p = leg_t.rows[0].cells[0].paragraphs[0]
fmt(p)
cell_margins(leg_t.rows[0].cells[0], left=0.2, top=0.08, bottom=0.08)
rn(p, "Classification:  ", bold=True, size=6.5, fg=NAVY)
for lbl, col, suffix in [
    ("Compliant", "1B5E20", "  all requirements met    "),
    ("Minor NC",  "7B5800", "  low risk, rectify within agreed timeframe    "),
    ("Major NC",  "B84700", "  significant risk, immediate management action    "),
    ("Critical / Stop Work", RED, "  imminent danger, work must cease until controlled"),
]:
    rn(p, lbl, bold=True, size=6.5, fg=col)
    rn(p, suffix, size=6.5, fg=MUTED)

spacer(doc, 3)

# ── FOOTER LINE ───────────────────────────────────────────────────────────────
ft = doc.add_table(rows=1, cols=3)
no_borders(ft)
set_bg(ft.rows[0].cells[0], NAVY)
set_bg(ft.rows[0].cells[1], NAVY)
set_bg(ft.rows[0].cells[2], ORANGE)
ft.rows[0].cells[0].width = Cm(10)
ft.rows[0].cells[2].width = Cm(3.8)
p = ft.rows[0].cells[0].paragraphs[0]
fmt(p)
cell_margins(ft.rows[0].cells[0], left=0.2, top=0.07, bottom=0.07)
rn(p, "NZGTTM Site Assurance Form  ·  Risk-Based  ·  Safe System  ·  Continuous Improvement", size=6, fg="88AACC")
p = ft.rows[0].cells[1].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
rn(p, "Retain on site — forward to TMO and RCA within 24 hours of completion", size=6, fg="88AACC", italic=True)
p = ft.rows[0].cells[2].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER)
cell_margins(ft.rows[0].cells[2], top=0.07, bottom=0.07)
rn(p, "NZGTTM Site Assurance Form - Version 1.0", bold=True, size=6, fg=WHITE)

# ── SAVE ─────────────────────────────────────────────────────────────────────
out_path = "/home/user/Site-condition-rating-form/NZGTTM-Site-Assurance-Form-v1.0.docx"
doc.save(out_path)
print(f"Saved: {out_path}")
