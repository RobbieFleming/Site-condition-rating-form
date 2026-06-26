#!/usr/bin/env python3
"""Generate NZGTTM Site Assurance Form — single A4 page Word document."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── PALETTE ──────────────────────────────────────────────────────────────────
NAVY    = "00274D"
ORANGE  = "E07728"
RED_HDR = "7A0E0E"
RED     = "C0392B"
L_BLUE  = "E8F0FA"
GRAY    = "F5F6F8"
WHITE   = "FFFFFF"
BORDER  = "C5CAD5"
TEXT    = "1C1C1E"
MUTED   = "555555"

CB = "☐"   # Checkbox — replace with ☑ when completing digitally, or tick when printed

# ── HELPERS ──────────────────────────────────────────────────────────────────

def set_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    for s in tcPr.findall(qn("w:shd")):
        tcPr.remove(s)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def set_table_borders(table, color=BORDER, sz=4):
    tbl  = table._tbl
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
    tbl  = table._tbl
    tblPr = tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:color"), "auto")
        b.append(el)
    tblPr.append(b)

def pad(cell, top=0.03, bottom=0.03, left=0.1, right=0.1):
    tcPr = cell._tc.get_or_add_tcPr()
    mar  = OxmlElement("w:tcMar")
    for side, v in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(int(v * 567)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)

def row_h(row, cm, exact=False):
    trPr = row._tr.get_or_add_trPr()
    for x in trPr.findall(qn("w:trHeight")):
        trPr.remove(x)
    el = OxmlElement("w:trHeight")
    el.set(qn("w:val"), str(int(cm * 567)))
    el.set(qn("w:hRule"), "exact" if exact else "atLeast")
    trPr.append(el)

def fmt(p, before=0, after=0, align=WD_ALIGN_PARAGRAPH.LEFT, spacing=None):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    p.alignment = align
    if spacing:
        p.paragraph_format.line_spacing = Pt(spacing)

def rn(p, text, bold=False, size=7, color=TEXT, italic=False, font="Arial"):
    r = p.add_run(text)
    r.bold   = bold
    r.italic = italic
    r.font.name = font
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor.from_string(color)
    return r

def write(cell, text, bold=False, size=7, fg=TEXT, bg=None,
          align=WD_ALIGN_PARAGRAPH.LEFT, va=WD_ALIGN_VERTICAL.CENTER,
          italic=False):
    if bg:
        set_bg(cell, bg)
    cell.vertical_alignment = va
    p = cell.paragraphs[0]
    fmt(p, align=align, spacing=7)
    rn(p, text, bold=bold, size=size, color=fg, italic=italic)
    pad(cell)

def gap(container, pt=2):
    """Add a tiny spacer paragraph."""
    p = container.add_paragraph()
    fmt(p)
    p.paragraph_format.line_spacing = Pt(pt)

# ── SECTION HEADER BUILDER ────────────────────────────────────────────────────

def sec_hdr(container, letter, title, badge=ORANGE, bg=NAVY, note=None, badge_red=False):
    t = container.add_table(rows=1, cols=3)
    no_borders(t)
    r = t.rows[0]
    # Badge
    set_bg(r.cells[0], RED if badge_red else badge)
    r.cells[0].width = Cm(0.58)
    r.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = r.cells[0].paragraphs[0]
    fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=7)
    rn(p, letter, bold=True, size=8, color=WHITE)
    pad(r.cells[0], top=0.05, bottom=0.05, left=0.03, right=0.03)
    # Title
    set_bg(r.cells[1], RED_HDR if badge_red else bg)
    r.cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = r.cells[1].paragraphs[0]
    fmt(p, spacing=7)
    rn(p, title.upper(), bold=True, size=7.5, color=WHITE)
    pad(r.cells[1], left=0.12, top=0.05, bottom=0.05)
    # Note
    set_bg(r.cells[2], RED_HDR if badge_red else bg)
    r.cells[2].width = Cm(3.2)
    r.cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = r.cells[2].paragraphs[0]
    fmt(p, align=WD_ALIGN_PARAGRAPH.RIGHT, spacing=7)
    if note:
        rn(p, note, size=6, color="FFCCCC" if badge_red else "AACCEE", italic=True)
    pad(r.cells[2], right=0.1, top=0.05, bottom=0.05)
    return t

# ── CHECKLIST TABLE BUILDER ───────────────────────────────────────────────────

CHECK_W = Cm(1.0)   # Width of each YES / NO / N/A column

def chk_tbl(container, items, yes_lbl="YES", hdr_bg=L_BLUE, risk_mode=False):
    t = container.add_table(rows=1 + len(items), cols=4)
    set_table_borders(t, sz=3)
    # Header row
    h = t.rows[0]
    row_h(h, 0.42, exact=True)
    hdrs = [("Checklist Item", WD_ALIGN_PARAGRAPH.LEFT),
            (yes_lbl,         WD_ALIGN_PARAGRAPH.CENTER),
            ("NO",            WD_ALIGN_PARAGRAPH.CENTER),
            ("N/A",           WD_ALIGN_PARAGRAPH.CENTER)]
    for i, (lbl, al) in enumerate(hdrs):
        set_bg(h.cells[i], hdr_bg)
        h.cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = h.cells[i].paragraphs[0]
        fmt(p, align=al, spacing=6)
        fc = RED if (risk_mode and i == 1) else NAVY
        rn(p, lbl, bold=True, size=6, color=fc)
        pad(h.cells[i], top=0.03, bottom=0.03, left=0.08, right=0.03)
        if i > 0:
            h.cells[i].width = CHECK_W
    # Data rows
    for ri, item in enumerate(items):
        row = t.rows[ri + 1]
        row_h(row, 0.38, exact=True)
        bg = WHITE if ri % 2 == 0 else "FAFBFD"
        set_bg(row.cells[0], bg)
        row.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = row.cells[0].paragraphs[0]
        fmt(p, spacing=6.5)
        rn(p, item, size=6.5, color=TEXT)
        pad(row.cells[0], top=0.02, bottom=0.02, left=0.1, right=0.05)
        for ci in (1, 2, 3):
            set_bg(row.cells[ci], bg)
            row.cells[ci].width = CHECK_W
            row.cells[ci].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = row.cells[ci].paragraphs[0]
            fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=6.5)
            rn(p, CB, size=9, color="333333", font="Segoe UI Symbol")
            pad(row.cells[ci], top=0.02, bottom=0.02, left=0.02, right=0.02)
    return t

# ── BUILD DOCUMENT ────────────────────────────────────────────────────────────

doc = Document()
sec = doc.sections[0]
sec.page_width    = Cm(21.0)
sec.page_height   = Cm(29.7)
M = 0.8
sec.left_margin   = Cm(M)
sec.right_margin  = Cm(M)
sec.top_margin    = Cm(M)
sec.bottom_margin = Cm(M)

style = doc.styles["Normal"]
style.font.name = "Arial"
style.font.size = Pt(7)
style.paragraph_format.space_before = Pt(0)
style.paragraph_format.space_after  = Pt(0)

FULL = Cm(21 - 2 * M)   # 19.4 cm usable width
HALF = Cm((21 - 2 * M - 0.25) / 2)   # ~9.575 cm per column

# ── HEADER ───────────────────────────────────────────────────────────────────
hdr = doc.add_table(rows=1, cols=3)
no_borders(hdr)
r = hdr.rows[0]
row_h(r, 1.2, exact=True)

# Logo
set_bg(r.cells[0], NAVY)
r.cells[0].width = Cm(3.5)
r.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
pad(r.cells[0], left=0.2, top=0.1, bottom=0.1)
p = r.cells[0].paragraphs[0]
fmt(p, spacing=14)
rn(p, "NZGTTM", bold=True, size=15, color=WHITE)
p2 = r.cells[0].add_paragraph()
fmt(p2, spacing=6)
rn(p2, "New Zealand Guide to TTM", size=5.5, color=ORANGE)

# Title
set_bg(r.cells[1], NAVY)
r.cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
pad(r.cells[1], left=0.2, top=0.1, bottom=0.1)
p = r.cells[1].paragraphs[0]
fmt(p, spacing=11)
rn(p, "Site Assurance and Audit Form", bold=True, size=12, color=WHITE)
p2 = r.cells[1].add_paragraph()
fmt(p2, spacing=6)
rn(p2, "Risk-Based Assessment  ·  Safe System Approach  ·  Replaces CoPTTM SCR Form", size=6, color="90B4D8")
p3 = r.cells[1].add_paragraph()
fmt(p3, spacing=6)
rn(p3, "For use by: TMO  ·  STMS  ·  Corridor Managers  ·  RCA Auditors", size=5.5, color="7090B0", italic=True)

# Ref box
set_bg(r.cells[2], ORANGE)
r.cells[2].width = Cm(2.5)
r.cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
pad(r.cells[2], top=0.1, bottom=0.1)
p = r.cells[2].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=6)
rn(p, "Form Ref", size=5.5, color=WHITE)
p2 = r.cells[2].add_paragraph()
fmt(p2, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=11)
rn(p2, "SAF-1.0", bold=True, size=11, color=WHITE)
p3 = r.cells[2].add_paragraph()
fmt(p3, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=6)
rn(p3, "NZGTTM", size=5.5, color=WHITE)

gap(doc, 2)

# ── SECTION A — SITE DETAILS ─────────────────────────────────────────────────
sec_hdr(doc, "A", "Site Details", note="Complete all fields before commencing audit")

det = doc.add_table(rows=3, cols=4)
set_table_borders(det, sz=3)
row_h(det.rows[0], 0.45, exact=True)
row_h(det.rows[1], 0.45, exact=True)
row_h(det.rows[2], 0.45, exact=True)

all_labels = [
    ["RCA:",              "Project Name:",           "TGS / TMP No.:",    "Date / Time:"],
    ["Road / Location:",  "STMS Name:",              "Contact Number:",    "Auditor Name:"],
    ["GPS Coordinates:",  "Company / Organisation:", "Weather:",           ""],
]
for ri, labels in enumerate(all_labels):
    row = det.rows[ri]
    for ci, lbl in enumerate(labels):
        cell = row.cells[ci]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        fmt(p, spacing=7)
        rn(p, lbl, bold=True, size=6, color=MUTED)
        pad(cell, top=0.02, bottom=0.02, left=0.1)

gap(doc, 2)

# ── SECTIONS B–G  (2-column grid) ────────────────────────────────────────────
outer = doc.add_table(rows=1, cols=3)
no_borders(outer)
outer.rows[0].cells[0].width = HALF
outer.rows[0].cells[1].width = Cm(0.25)   # gap column
outer.rows[0].cells[2].width = HALF
lc = outer.rows[0].cells[0]   # left
gc = outer.rows[0].cells[1]   # gap
rc = outer.rows[0].cells[2]   # right
for c in (lc, gc, rc):
    c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    pad(c, top=0, bottom=0, left=0, right=0)
set_bg(gc, WHITE)

# ── LEFT COLUMN: B, D, F ──────────────────────────────────────────────────────

sec_hdr(lc, "B", "Immediate Safety Risks",
        badge_red=True, note="⚠ YES = Risk Present")

# Alert bar
ab = lc.add_table(rows=1, cols=1)
no_borders(ab)
set_bg(ab.rows[0].cells[0], "FFF0F0")
pad(ab.rows[0].cells[0], left=0.1, top=0.04, bottom=0.04)
p = ab.rows[0].cells[0].paragraphs[0]
fmt(p, spacing=6)
rn(p, "⚠  Any YES requires immediate corrective action — consider Stop Work", bold=True, size=6, color="8B0000")

chk_tbl(lc, [
    "Workers exposed to live traffic",
    "Vehicle intrusion into work area",
    "Inadequate worker separation",
    "Unsafe pedestrian route",
    "Unsafe cyclist route/accommodation",
    "Visibility / sign obstruction",
    "Queue management issues",
    "Plant operating unsafely near traffic",
    "Emergency access compromised",
], yes_lbl="YES", risk_mode=True)

gap(lc, 2)
sec_hdr(lc, "D", "Vulnerable Road Users", note="YES = Compliant")
chk_tbl(lc, [
    "Pedestrian route provided",
    "Pedestrian route accessible",
    "Cyclist accommodation provided",
    "Crossing facilities adequate",
    "Mobility-impaired users considered",
])

gap(lc, 2)
sec_hdr(lc, "F", "Worksite Safety", note="YES = Compliant")
chk_tbl(lc, [
    "Crew briefed — TM, hazards, emergency",
    "All personnel wearing correct PPE",
    "Plant operating safely, exclusion zones",
    "Materials and equipment secured",
    "Work area housekeeping acceptable",
    "Emergency procedures understood",
])

# ── RIGHT COLUMN: C, E, G ─────────────────────────────────────────────────────

sec_hdr(rc, "C", "Traffic Management Layout", note="YES = Compliant")
chk_tbl(rc, [
    "Site matches approved TGS/TMP",
    "Signs correctly positioned",
    "Signs visible, legible and clean",
    "Delineation adequate and spaced",
    "Tapers appropriate for speed",
    "Work area clearly defined",
    "Traffic lanes clearly guided",
    "Temporary speed management compliant",
    "Access points controlled",
])

gap(rc, 2)
sec_hdr(rc, "E", "Traffic Operations", note="YES = Compliant")
chk_tbl(rc, [
    "Traffic flow operating safely",
    "Driver behaviour appropriate",
    "Delays within acceptable limits",
    "Queue lengths managed effectively",
    "No driver confusion at merge points",
    "Sight distances adequate",
])

gap(rc, 2)
sec_hdr(rc, "G", "Documentation", note="YES = Available & Current")
chk_tbl(rc, [
    "Current TGS/TMP available on site",
    "Permits available and current",
    "RCA approvals available on site",
    "Daily TM records up to date",
    "TM layout changes documented",
    "Risk assessment available on site",
])

gap(doc, 2)

# ── SECTION H — NON-CONFORMANCES ─────────────────────────────────────────────
sec_hdr(doc, "H", "Non-Conformances Identified",
        note="Complete for every NO response above")

nc = doc.add_table(rows=6, cols=7)
set_table_borders(nc, sz=3)
NC_W = [Cm(0.6), Cm(4.8), Cm(2.2), Cm(4.5), Cm(2.9), Cm(1.7), Cm(1.5)]
NC_H = ["#", "Description of Non-Conformance",
        "Risk Level\n(Low / Med / High / Critical)",
        "Corrective Action Required",
        "Responsible Person", "Due Date", "Closed ✓"]

hrow = nc.rows[0]
row_h(hrow, 0.5, exact=True)
for i, (h, w) in enumerate(zip(NC_H, NC_W)):
    hrow.cells[i].width = w
    set_bg(hrow.cells[i], L_BLUE)
    hrow.cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = hrow.cells[i].paragraphs[0]
    fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=6)
    rn(p, h, bold=True, size=6, color=NAVY)
    pad(hrow.cells[i], top=0.03, bottom=0.03, left=0.06, right=0.03)

for ri in range(1, 6):
    row = nc.rows[ri]
    row_h(row, 0.72, exact=True)
    bg = WHITE if ri % 2 != 0 else "FAFBFD"
    set_bg(row.cells[0], GRAY)
    p = row.cells[0].paragraphs[0]
    fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=7)
    rn(p, str(ri), bold=True, size=7, color=TEXT)
    row.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    pad(row.cells[0], top=0.03, bottom=0.03, left=0.03, right=0.03)
    for ci in range(1, 7):
        set_bg(row.cells[ci], bg)
        row.cells[ci].width = NC_W[ci]

gap(doc, 2)

# ── I + J SIDE BY SIDE ───────────────────────────────────────────────────────
ij = doc.add_table(rows=1, cols=3)
no_borders(ij)
ij.rows[0].cells[0].width = HALF
ij.rows[0].cells[1].width = Cm(0.25)
ij.rows[0].cells[2].width = HALF
ic = ij.rows[0].cells[0]
jc = ij.rows[0].cells[2]
for c in (ic, ij.rows[0].cells[1], jc):
    c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    pad(c, top=0, bottom=0, left=0, right=0)
set_bg(ij.rows[0].cells[1], WHITE)

# Section I — Overall Assessment
sec_hdr(ic, "I", "Overall Site Assessment")

rating_t = ic.add_table(rows=4, cols=1)
no_borders(rating_t)
ratings = [
    ("Compliant",              "1B5E20", "F0FFF4"),
    ("Minor Non-Conformance",  "7B5800", "FFFDE7"),
    ("Major Non-Conformance",  "B71C1C", "FFF8F0"),
    ("Critical Risk / Stop Work", "8B0000", "FFF5F5"),
]
for i, (lbl, col, bg) in enumerate(ratings):
    row = rating_t.rows[i]
    row_h(row, 0.5, exact=True)
    set_bg(row.cells[0], bg)
    p = row.cells[0].paragraphs[0]
    fmt(p, spacing=7)
    rn(p, CB + "  ", size=9, color="333333", font="Segoe UI Symbol")
    rn(p, lbl, bold=True, size=7.5, color=col)
    pad(row.cells[0], left=0.12, top=0.04, bottom=0.04)

p_lbl = ic.add_paragraph()
fmt(p_lbl, before=2, spacing=6)
rn(p_lbl, "AUDITOR COMMENTS / IMMEDIATE ACTIONS TAKEN:", bold=True, size=6, color=NAVY)

ct = ic.add_table(rows=4, cols=1)
set_table_borders(ct, color=BORDER, sz=2)
for row in ct.rows:
    row_h(row, 0.45, exact=True)
    set_bg(row.cells[0], WHITE)
    pad(row.cells[0], top=0.02, bottom=0.02, left=0.1)
    fmt(row.cells[0].paragraphs[0], spacing=6)

p_ri = ic.add_paragraph()
fmt(p_ri, before=2, spacing=6)
rn(p_ri, "Re-inspection required: ", bold=True, size=6.5, color=NAVY)
rn(p_ri, CB + " Yes   ", size=9, color="333333", font="Segoe UI Symbol")
rn(p_ri, CB + " No   ", size=9, color="333333", font="Segoe UI Symbol")
rn(p_ri, "By: _______________________", size=6.5, color=MUTED)

# Section J — Sign-offs
sec_hdr(jc, "J", "Sign-Offs",
        note="Both parties must sign to confirm accuracy")

so = jc.add_table(rows=1, cols=2)
no_borders(so)
for ci, (title, fields_so) in enumerate([
    ("Auditor / Inspector",      ["Name:", "Role / Title:", "Company:", "Date:"]),
    ("STMS / Site Representative", ["Name:", "STMS Cert No.:", "Company:", "Date:"]),
]):
    c = so.rows[0].cells[ci]
    c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    pad(c, left=0.08, right=0.08, top=0, bottom=0)

    p_t = c.paragraphs[0]
    fmt(p_t, spacing=7)
    rn(p_t, title.upper(), bold=True, size=7, color=NAVY)
    pPr = p_t._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single"); bot.set(qn("w:sz"), "6"); bot.set(qn("w:color"), NAVY)
    pBdr.append(bot)
    pPr.append(pBdr)

    for fld in fields_so:
        pf = c.add_paragraph()
        fmt(pf, before=1, spacing=6.5)
        rn(pf, fld + "  ", bold=True, size=6, color=MUTED)
        rn(pf, "_" * 24, size=6.5, color="AAAAAA")

    sig_t = c.add_table(rows=1, cols=1)
    set_table_borders(sig_t, color="AAAAAA", sz=4)
    row_h(sig_t.rows[0], 1.3, exact=True)
    set_bg(sig_t.rows[0].cells[0], WHITE)
    p_s = sig_t.rows[0].cells[0].paragraphs[0]
    fmt(p_s, spacing=6)
    rn(p_s, "Signature", size=6, color="CCCCCC", italic=True)
    pad(sig_t.rows[0].cells[0], left=0.1, top=0.05)

gap(doc, 2)

# ── LEGEND ────────────────────────────────────────────────────────────────────
leg = doc.add_table(rows=1, cols=1)
no_borders(leg)
set_bg(leg.rows[0].cells[0], GRAY)
row_h(leg.rows[0], 0.38, exact=True)
pad(leg.rows[0].cells[0], left=0.15, top=0.04, bottom=0.04)
leg.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
p = leg.rows[0].cells[0].paragraphs[0]
fmt(p, spacing=6)
rn(p, "Classification:  ", bold=True, size=6, color=NAVY)
for lbl, col, sfx in [
    ("Compliant", "1B5E20", "  all requirements met    "),
    ("Minor NC",  "7B5800", "  low risk, rectify within timeframe    "),
    ("Major NC",  "B84700", "  significant risk, immediate action    "),
    ("Critical / Stop Work", RED, "  imminent danger, work must cease"),
]:
    rn(p, lbl, bold=True, size=6, color=col)
    rn(p, sfx, size=6, color=MUTED)

gap(doc, 2)

# ── FOOTER ────────────────────────────────────────────────────────────────────
ft = doc.add_table(rows=1, cols=3)
no_borders(ft)
row_h(ft.rows[0], 0.38, exact=True)
set_bg(ft.rows[0].cells[0], NAVY); ft.rows[0].cells[0].width = Cm(9.5)
set_bg(ft.rows[0].cells[1], NAVY)
set_bg(ft.rows[0].cells[2], ORANGE); ft.rows[0].cells[2].width = Cm(4.2)
for c in ft.rows[0].cells:
    c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
p = ft.rows[0].cells[0].paragraphs[0]
fmt(p, spacing=6); pad(ft.rows[0].cells[0], left=0.15, top=0.04, bottom=0.04)
rn(p, "NZGTTM Site Assurance Form  ·  Risk-Based  ·  Safe System  ·  Continuous Improvement", size=5.5, color="88AACC")
p = ft.rows[0].cells[1].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=6)
rn(p, "Forward to TMO and RCA within 24 hours of completion", size=5.5, color="88AACC", italic=True)
p = ft.rows[0].cells[2].paragraphs[0]
fmt(p, align=WD_ALIGN_PARAGRAPH.CENTER, spacing=6); pad(ft.rows[0].cells[2], top=0.04, bottom=0.04)
rn(p, "NZGTTM Site Assurance Form - Version 1.0", bold=True, size=5.5, color=WHITE)

# ── SAVE ─────────────────────────────────────────────────────────────────────
OUT = "/home/user/Site-condition-rating-form/NZGTTM-Site-Assurance-Form-v1.0.docx"
doc.save(OUT)
print(f"Saved: {OUT}")
