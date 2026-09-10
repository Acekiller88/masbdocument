#!/usr/bin/env python3
"""
Rebuild MIMOS Academy / MIMOS Solutions Purchase Order as a Word template
matching contoh po.PDF as closely as Word layout allows.

Fonts (substitutes; original PDF uses Helvetica / Courier Type1):
  Helvetica        -> Arial
  Helvetica-Bold   -> Arial Bold
  Helvetica-Oblique-> Arial Italic
  Courier          -> Courier New
  Courier-Bold     -> Courier New Bold

Geometry is taken from PDF drawings (pt on A4 595 x 842).
"""
from __future__ import annotations

import json
import shutil
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentClass
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
from docx.shared import Cm, Emu, Inches, Pt, RGBColor, Twips
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent
FORENSIC = ROOT / "_forensic"
LOGO_SRC = ROOT / "MIMOS_Solutions_LOGO_from_PO.png"
if not LOGO_SRC.exists():
    LOGO_SRC = FORENSIC / "images" / "xref9.png"
SPANS_DIR = FORENSIC
OUT_DOCX = ROOT / "MIMOS_Academy_Blank_Purchase_Order_Template.docx"

# --- measured geometry (PDF points) ---
PAGE_W = 595.27
PAGE_H = 841.89
MARGIN_L = 28.35
MARGIN_R = 28.32
MARGIN_T = 28.35
MARGIN_B = 8.50
CONTENT_W = 538.60  # 566.95 - 28.35

PURPLE = "801070"          # measured banner fill RGB(128,16,112)
PURPLE_TOTAL = "8F1B7E"    # measured TOTAL bar RGB(143,27,126)
BLACK = "000000"
WHITE = "FFFFFF"

# header split
LEFT_W = 283.45
META_W = 255.15
META_COL = 85.05

# vendor row under supplier
VEND_W, TEL_W, FAX_W = 70.85, 99.20, 113.40

# deliver nested
DEL_W, EXP_W = 170.10, 85.05

# PR row
PR_W, BUY_W, RFQ_W = 141.75, 212.60, 184.25

# items
ITEM_W, DESC_W, QTY_W, REST_W = 34.00, 158.75, 82.20, 263.65
UNIT_W, PER_W, UP_W, TP_W = 28.35, 31.20, 93.55, 110.55
CHK_W, APPR_W = 93.55, 170.10  # 93.55+170.10 = 263.65

H_HEADER = 90.80          # 119.15-28.35
H_PO_BANNER = 14.15
H_PO_META = 31.20
H_PO_BANNER2 = 14.15
H_SUPPLIER_BLOCK = 147.40  # 266.55-119.15
H_SUPPLIER = 119.05
H_VENDOR = 28.35
H_DELIVER = 48.20
H_ATTN = 28.35
H_DELIVERY = 36.85
H_PAYMENT = 34.00
H_GAP_PR = 17.00
H_PR = 28.35
H_SUPPLY_BANNER = 14.15
H_ITEM_HDR = 28.35
H_ITEM_BODY = 263.60      # 354.45-618.05
H_PRICE_BODY = 178.60     # 354.45-533.05
H_TOTAL = 28.35
H_APPROVAL = 56.70
H_GAP_ACK = 3.95
H_ACK = 146.40
H_GAP_SIGN = 48.80
H_SIGN = 16.20


def tw(pt: float) -> int:
    return int(round(pt * 20))


def emu_pt(pt: float) -> int:
    return int(round(pt * 12700))


def RGB(hex_color: str) -> RGBColor:
    hex_color = hex_color.lstrip("#")
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))


# ---------------------------------------------------------------------------
# Low-level OOXML helpers
# ---------------------------------------------------------------------------
def _get_or_add(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def set_run_font(run, name: str, size_pt: float, *, bold=False, italic=False, color: str | None = None):
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size_pt)
    run.font.name = name
    rPr = run._element.get_or_add_rPr()
    rFonts = _get_or_add(rPr, "w:rFonts")
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:cs"), name)
    rFonts.set(qn("w:eastAsia"), name)
    if color:
        run.font.color.rgb = RGB(color)
        c = _get_or_add(rPr, "w:color")
        c.set(qn("w:val"), color)


def set_para(p: Paragraph, *, align="left", before=0, after=0, line_pt=None, exact=True, keep_together=False):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.widow_control = False
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    p.alignment = align_map.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    pPr = p._p.get_or_add_pPr()
    spacing = _get_or_add(pPr, "w:spacing")
    spacing.set(qn("w:before"), str(tw(before)))
    spacing.set(qn("w:after"), str(tw(after)))
    if line_pt is not None:
        spacing.set(qn("w:line"), str(tw(line_pt)))
        spacing.set(qn("w:lineRule"), "exact" if exact else "auto")
        pf.line_spacing = Pt(line_pt)
    ind = _get_or_add(pPr, "w:ind")
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:right"), "0")
    ind.set(qn("w:firstLine"), "0")
    jc = _get_or_add(pPr, "w:jc")
    jc.set(qn("w:val"), align)
    # no extra auto space
    for tag in ("w:contextualSpacing",):
        el = OxmlElement(tag)
        el.set(qn("w:val"), "true")
        pPr.append(el)


def squeeze_para(p: Paragraph, line_pt=1.0):
    set_para(p, align="left", before=0, after=0, line_pt=line_pt, exact=True)


def set_cell_shading(cell: _Cell, fill: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = _get_or_add(tcPr, "w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_cell_valign(cell: _Cell, valign: str):
    mapping = {
        "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
        "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
        "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
    }
    cell.vertical_alignment = mapping.get(valign, WD_CELL_VERTICAL_ALIGNMENT.TOP)


def set_cell_margins(cell: _Cell, top=40, bottom=40, left=60, right=40):
    """Margins in twips (20 twips = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = _get_or_add(tcPr, "w:tcMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = _get_or_add(tcMar, f"w:{name}")
        el.set(qn("w:w"), str(int(val)))
        el.set(qn("w:type"), "dxa")


def set_cell_border(cell: _Cell, **sides):
    """sides: top/left/bottom/right/insideH/insideV -> (sz, color) or 'nil'."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = _get_or_add(tcPr, "w:tcBorders")
    for edge, spec in sides.items():
        el = _get_or_add(tcBorders, f"w:{edge}")
        if spec is None or spec == "nil":
            el.set(qn("w:val"), "nil")
            el.set(qn("w:sz"), "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
        else:
            sz, color = spec
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(sz))  # eighths of a point; 4 = 0.5pt
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)


def no_cell_borders(cell: _Cell):
    set_cell_border(cell, top="nil", left="nil", bottom="nil", right="nil")


def set_tbl_borders(table: Table, sz=4, color="000000", edges=None):
    tblPr = table._tbl.tblPr
    borders = _get_or_add(tblPr, "w:tblBorders")
    use = ("top", "left", "bottom", "right", "insideH", "insideV") if edges is None else edges
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = _get_or_add(borders, f"w:{edge}")
        if edge in use:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)
        else:
            el.set(qn("w:val"), "nil")
            el.set(qn("w:sz"), "0")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")


def no_tbl_borders(table: Table):
    set_tbl_borders(table, edges=())


def set_tbl_width(table: Table, width_pt: float):
    tblPr = table._tbl.tblPr
    tblW = _get_or_add(tblPr, "w:tblW")
    tblW.set(qn("w:w"), str(tw(width_pt)))
    tblW.set(qn("w:type"), "dxa")
    layout = _get_or_add(tblPr, "w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    jc = _get_or_add(tblPr, "w:jc")
    jc.set(qn("w:val"), "left")
    # indent 0
    tblInd = _get_or_add(tblPr, "w:tblInd")
    tblInd.set(qn("w:w"), "0")
    tblInd.set(qn("w:type"), "dxa")
    # cell spacing 0
    tblCellSpacing = _get_or_add(tblPr, "w:tblCellSpacing")
    tblCellSpacing.set(qn("w:w"), "0")
    tblCellSpacing.set(qn("w:type"), "dxa")
    table.autofit = False
    table.allow_autofit = False


def set_tbl_cell_mar(table: Table, top=40, bottom=40, left=60, right=40):
    tblPr = table._tbl.tblPr
    mar = _get_or_add(tblPr, "w:tblCellMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = _get_or_add(mar, f"w:{name}")
        el.set(qn("w:w"), str(int(val)))
        el.set(qn("w:type"), "dxa")


def set_col_widths(table: Table, widths_pt: list[float]):
    set_tbl_width(table, sum(widths_pt))
    tbl = table._tbl
    tblPr = tbl.tblPr
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tblPr.addnext(grid)
    else:
        for child in list(grid):
            grid.remove(child)
    for wpt in widths_pt:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(tw(wpt)))
        grid.append(gc)
    for row in table.rows:
        cells = row.cells
        # only set width on the first cell of each grid slot; merged cells handled by gridSpan
        for i, wpt in enumerate(widths_pt):
            if i >= len(cells):
                break
            tc = cells[i]._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = _get_or_add(tcPr, "w:tcW")
            tcW.set(qn("w:w"), str(tw(wpt)))
            tcW.set(qn("w:type"), "dxa")


def set_row_height(row, height_pt: float, rule="exact"):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trH = _get_or_add(trPr, "w:trHeight")
    trH.set(qn("w:val"), str(tw(height_pt)))
    trH.set(qn("w:hRule"), rule)


def prevent_row_split(row):
    trPr = row._tr.get_or_add_trPr()
    cs = _get_or_add(trPr, "w:cantSplit")
    cs.set(qn("w:val"), "true")


def clear_tc(tc):
    for child in list(tc):
        if child.tag != qn("w:tcPr"):
            tc.remove(child)


def new_p(line_pt=9.0) -> OxmlElement:
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), str(tw(line_pt)))
    sp.set(qn("w:lineRule"), "exact")
    pPr.append(sp)
    p.append(pPr)
    return p


def add_run_el(p, text, font, size_pt, *, bold=False, italic=False, color=None):
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)
    rFonts.set(qn("w:cs"), font)
    rFonts.set(qn("w:eastAsia"), font)
    rPr.append(rFonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(round(size_pt * 2))))
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(round(size_pt * 2))))
    rPr.append(szCs)
    if bold:
        rPr.append(OxmlElement("w:b"))
        rPr.append(OxmlElement("w:bCs"))
    if italic:
        rPr.append(OxmlElement("w:i"))
        rPr.append(OxmlElement("w:iCs"))
    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rPr.append(c)
    r.append(rPr)
    t = OxmlElement("w:t")
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    return r


def write_runs(cell: _Cell, paragraphs, *, valign="top", margins=(20, 20, 40, 40)):
    """
    paragraphs: list of dicts:
      {align, line, before, after, runs:[(text,font,size,bold,italic,color), ...]}
    """
    tc = cell._tc
    clear_tc(tc)
    top, bottom, left, right = margins
    set_cell_margins(cell, top=top, bottom=bottom, left=left, right=right)
    set_cell_valign(cell, valign)
    if not paragraphs:
        tc.append(new_p(1))
        return
    for spec in paragraphs:
        line = spec.get("line", 9.0)
        p = new_p(line)
        pPr = p.find(qn("w:pPr"))
        sp = pPr.find(qn("w:spacing"))
        sp.set(qn("w:before"), str(tw(spec.get("before", 0))))
        sp.set(qn("w:after"), str(tw(spec.get("after", 0))))
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), spec.get("align", "left"))
        pPr.append(jc)
        for run in spec.get("runs", []):
            text, font, size = run[0], run[1], run[2]
            bold = run[3] if len(run) > 3 else False
            italic = run[4] if len(run) > 4 else False
            color = run[5] if len(run) > 5 else None
            add_run_el(p, text, font, size, bold=bold, italic=italic, color=color)
        tc.append(p)


def P(text, font="Arial", size=8, *, bold=False, italic=False, color=None,
      align="left", line=None, before=0, after=0):
    if line is None:
        line = size + 1
    return {
        "align": align,
        "line": line,
        "before": before,
        "after": after,
        "runs": [(text, font, size, bold, italic, color)],
    }


def PM(runs, *, align="left", line=9, before=0, after=0):
    return {"align": align, "line": line, "before": before, "after": after, "runs": runs}


def cell_add_table(cell: _Cell, rows: int, cols: int) -> Table:
    tc = cell._tc
    clear_tc(tc)
    tbl = cell.add_table(rows, cols)
    # trailing collapsed paragraph so Word is happy (0-height)
    p = new_p(0)
    sp = p.find(qn("w:pPr")).find(qn("w:spacing"))
    sp.set(qn("w:line"), "0")
    tc.append(p)
    return tbl


def apply_inner_grid(table: Table):
    """Draw only interior grid lines via cell borders (outer edges stay nil
    so a parent cell/table provides the outer box once)."""
    nrows = len(table.rows)
    ncols = len(table.columns)
    for r in range(nrows):
        for c in range(ncols):
            set_cell_border(
                table.cell(r, c),
                top="nil",
                left="nil",
                bottom=(4, "000000") if r < nrows - 1 else "nil",
                right=(4, "000000") if c < ncols - 1 else "nil",
            )


def add_table(parent, rows, cols, widths_pt, *, borders=True, cell_mar=(20, 20, 40, 40), nested=False) -> Table:
    if isinstance(parent, DocumentClass):
        table = parent.add_table(rows, cols)
    else:
        table = cell_add_table(parent, rows, cols)
    set_col_widths(table, widths_pt)
    set_tbl_cell_mar(table, top=cell_mar[0], bottom=cell_mar[1], left=cell_mar[2], right=cell_mar[3])
    if nested:
        set_tbl_borders(table, sz=4, color="000000", edges=("insideH", "insideV"))
        apply_inner_grid(table)
    elif borders:
        set_tbl_borders(table, sz=4, color="000000")
    else:
        no_tbl_borders(table)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    return table


def spacer(doc: Document, height_pt: float, width_pt: float = CONTENT_W):
    t = add_table(doc, 1, 1, [width_pt], borders=False, cell_mar=(0, 0, 0, 0))
    set_row_height(t.rows[0], height_pt, "exact")
    write_runs(t.cell(0, 0), [], valign="top", margins=(0, 0, 0, 0))
    no_cell_borders(t.cell(0, 0))
    squeeze_following(doc)
    return t


def squeeze_following(doc: Document):
    """Remove empty body paragraphs inserted between tables so exact row
    heights control vertical rhythm. Keep page-break paragraphs."""
    body = doc.element.body
    for child in list(body):
        if child.tag != qn("w:p"):
            continue
        if child.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br") is not None:
            continue
        if child.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing") is not None:
            continue
        texts = [t.text or "" for t in child.iter(qn("w:t"))]
        if "".join(texts).strip():
            continue
        body.remove(child)


def delete_trailing_empty_paragraphs(doc: Document):
    body = doc.element.body
    children = list(body)
    # keep sectPr
    for child in reversed(children):
        if child.tag == qn("w:sectPr"):
            continue
        if child.tag == qn("w:p"):
            texts = [t.text or "" for t in child.iter(qn("w:t"))]
            if not "".join(texts).strip() and child.find(qn("w:r")) is None or (
                not "".join(texts).strip() and child.find(".//" + qn("w:drawing")) is None
                and child.find(".//" + qn("w:br")) is None
            ):
                # don't delete if it is the only thing; squeeze instead
                sp_el = child.find(qn("w:pPr"))
                if sp_el is None:
                    continue
        break


def page_break(doc: Document):
    p = doc.add_paragraph()
    squeeze_para(p, 1)
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)


def set_normal_style(doc: Document):
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(8)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing = 1.0
    rPr = style.element.get_or_add_rPr()
    rFonts = _get_or_add(rPr, "w:rFonts")
    rFonts.set(qn("w:ascii"), "Arial")
    rFonts.set(qn("w:hAnsi"), "Arial")
    rFonts.set(qn("w:cs"), "Arial")
    rFonts.set(qn("w:eastAsia"), "Arial")
    # document defaults
    styles = doc.styles.element
    docDefaults = styles.find(qn("w:docDefaults"))
    if docDefaults is None:
        docDefaults = OxmlElement("w:docDefaults")
        styles.insert(0, docDefaults)
    rPrDefault = _get_or_add(docDefaults, "w:rPrDefault")
    rPr = _get_or_add(rPrDefault, "w:rPr")
    rFonts = _get_or_add(rPr, "w:rFonts")
    rFonts.set(qn("w:ascii"), "Arial")
    rFonts.set(qn("w:hAnsi"), "Arial")
    rFonts.set(qn("w:cs"), "Arial")
    rFonts.set(qn("w:eastAsia"), "Arial")
    sz = _get_or_add(rPr, "w:sz")
    sz.set(qn("w:val"), "16")
    pPrDefault = _get_or_add(docDefaults, "w:pPrDefault")
    pPr = _get_or_add(pPrDefault, "w:pPr")
    sp = _get_or_add(pPr, "w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), "240")
    sp.set(qn("w:lineRule"), "auto")


def configure_section(section, *, top=MARGIN_T, bottom=MARGIN_B, left=MARGIN_L, right=MARGIN_R):
    section.page_width = Emu(emu_pt(PAGE_W))
    section.page_height = Emu(emu_pt(PAGE_H))
    section.left_margin = Emu(emu_pt(left))
    section.right_margin = Emu(emu_pt(right))
    section.top_margin = Emu(emu_pt(top))
    section.bottom_margin = Emu(emu_pt(bottom))
    section.header_distance = Emu(emu_pt(0))
    section.footer_distance = Emu(emu_pt(0))
    section.different_first_page_header_footer = False
    # Gutter 0 already default
    pgMar = section._sectPr.find(qn("w:pgMar"))
    if pgMar is not None:
        pgMar.set(qn("w:header"), "0")
        pgMar.set(qn("w:footer"), "0")
        pgMar.set(qn("w:gutter"), "0")


# ---------------------------------------------------------------------------
# Content builders
# ---------------------------------------------------------------------------
def add_logo(cell: _Cell, logo_path: Path):
    tc = cell._tc
    clear_tc(tc)
    set_cell_margins(cell, top=0, bottom=0, left=0, right=20)
    set_cell_valign(cell, "top")
    p = new_p(10)
    pPr = p.find(qn("w:pPr"))
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "left")
    pPr.append(jc)
    tc.append(p)
    # use python-docx picture via a temporary paragraph object
    para = Paragraph(p, cell)
    run = para.add_run()
    run.add_picture(str(logo_path), width=Pt(120), height=Pt(44.65))
    # formerly known as
    p2 = new_p(5)
    p2Pr = p2.find(qn("w:pPr"))
    sp = p2Pr.find(qn("w:spacing"))
    sp.set(qn("w:before"), str(tw(3)))
    add_run_el(
        p2,
        "(formerly known as MIMOS Technology Solutions Sdn. Bhd.)",
        "Arial",
        4,
        italic=True,
        color="404040",
    )
    tc.append(p2)


def build_header(doc: Document, page_label: str, logo_path: Path):
    outer = add_table(doc, 1, 2, [LEFT_W, META_W], borders=False, cell_mar=(0, 0, 0, 0))
    set_row_height(outer.rows[0], H_HEADER, "exact")
    left, right = outer.cell(0, 0), outer.cell(0, 1)
    no_cell_borders(left)
    no_cell_borders(right)

    # left: logo + company
    inner = add_table(left, 1, 2, [130.0, LEFT_W - 130.0], borders=False, cell_mar=(0, 0, 0, 0))
    set_row_height(inner.rows[0], H_HEADER, "exact")
    no_cell_borders(inner.cell(0, 0))
    no_cell_borders(inner.cell(0, 1))
    add_logo(inner.cell(0, 0), logo_path)

    company = [
        P("MIMOS SOLUTIONS SDN. BHD.", "Arial", 8, bold=True, line=9),
        P("201101029324 (957459-K)", "Arial", 8, bold=True, line=9),
        P("SST No. : W10-1808-32000747", "Arial", 8, line=9),
        P("Technology Park Malaysia", "Arial", 8, line=9),
        P("57000 Kuala Lumpur", "Arial", 8, line=9),
        P("Malaysia", "Arial", 8, line=9),
        P("Tel  No:+60 3 8995 5000", "Arial", 8, line=9),
        P("http://www.mimossolutions.my", "Arial", 8, line=9),
    ]
    write_runs(inner.cell(0, 1), company, valign="top", margins=(20, 0, 40, 0))

    # right: PO meta 3-col
    meta = add_table(right, 4, 3, [META_COL, META_COL, META_COL], borders=True, cell_mar=(20, 20, 20, 20))
    set_row_height(meta.rows[0], H_PO_BANNER, "exact")
    set_row_height(meta.rows[1], H_PO_META, "exact")
    set_row_height(meta.rows[2], H_PO_META, "exact")
    set_row_height(meta.rows[3], H_PO_BANNER2, "exact")

    # banner row merge
    a = meta.cell(0, 0).merge(meta.cell(0, 2))
    set_cell_shading(a, PURPLE)
    write_runs(
        a,
        [P("PURCHASE ORDER", "Arial", 10, bold=True, color=WHITE, align="center", line=12)],
        valign="center",
        margins=(0, 0, 20, 20),
    )

    # row 1 labels + values
    write_runs(
        meta.cell(1, 0),
        [
            P("Purchase Order No.", "Arial", 8, align="center", line=9),
            P("7000000306", "Courier New", 9, align="center", line=12, before=2),
        ],
        valign="top",
        margins=(40, 20, 20, 20),
    )
    write_runs(
        meta.cell(1, 1),
        [
            P("Purchase Order", "Arial", 8, align="center", line=9),
            P("Release No.", "Arial", 8, align="center", line=9),
            P("N/A", "Courier New", 9, align="center", line=11),
        ],
        valign="top",
        margins=(20, 10, 20, 20),
    )
    write_runs(
        meta.cell(1, 2),
        [
            P("Order Date", "Arial", 8, align="center", line=9),
            P("25.05.2026", "Courier New", 9, align="center", line=12, before=8),
        ],
        valign="top",
        margins=(40, 20, 20, 20),
    )

    write_runs(
        meta.cell(2, 0),
        [
            P("Contract No.", "Arial", 8, align="center", line=9),
            P("N/A", "Courier New", 9, align="center", line=12, before=6),
        ],
        valign="top",
        margins=(40, 20, 20, 20),
    )
    write_runs(
        meta.cell(2, 1),
        [
            P("Revision No.", "Arial", 8, align="center", line=9),
            P("N/A", "Courier New", 9, align="center", line=12, before=6),
        ],
        valign="top",
        margins=(40, 20, 20, 20),
    )
    write_runs(
        meta.cell(2, 2),
        [
            P("Page No", "Arial", 8, align="center", line=9),
            P(page_label, "Courier New", 9, align="center", line=12, before=6),
        ],
        valign="top",
        margins=(40, 20, 20, 20),
    )

    b = meta.cell(3, 0).merge(meta.cell(3, 2))
    set_cell_shading(b, PURPLE)
    write_runs(
        b,
        [
            P(
                "THIS PURCHASE ORDER NO. MUST APPEAR ON ALL INVOICES, PACKING LISTS,",
                "Arial",
                5.5,
                bold=True,
                color=WHITE,
                align="center",
                line=6.5,
            ),
            P(
                "CARTONS AND CORRESPONDENCE RELATED TO THIS ORDER",
                "Arial",
                5.5,
                bold=True,
                color=WHITE,
                align="center",
                line=6.5,
            ),
        ],
        valign="center",
        margins=(10, 10, 20, 20),
    )
    squeeze_following(doc)
    return outer


def build_supplier_block(doc: Document):
    outer = add_table(doc, 1, 2, [LEFT_W, META_W], borders=True, cell_mar=(0, 0, 0, 0))
    set_row_height(outer.rows[0], H_SUPPLIER_BLOCK, "exact")

    # LEFT: 3-col so vendor/tel/fax share the grid; row 0 is merged supplier
    left = add_table(
        outer.cell(0, 0), 2, 3, [VEND_W, TEL_W, FAX_W], borders=True, cell_mar=(20, 20, 80, 40), nested=True
    )
    set_row_height(left.rows[0], H_SUPPLIER, "exact")
    set_row_height(left.rows[1], H_VENDOR, "exact")
    sup = left.cell(0, 0).merge(left.cell(0, 2))
    apply_inner_grid(left)
    write_runs(
        sup,
        [
            P("Supplier", "Arial", 8, align="left", line=10, before=2),
            P("MIMOS SERVICES SDN BHD   ", "Courier New", 9, line=11, before=8),
            P("SST No. : W10-1808-21023186", "Courier New", 9, line=11),
            P("MIMOS Berhad,", "Courier New", 9, line=11),
            P("Taman Teknologi Malaysia,", "Courier New", 9, line=11),
            P("57000 Kuala Lumpur", "Courier New", 9, line=11),
            P("Kuala Lumpur", "Courier New", 9, line=11),
            P("Malaysia", "Courier New", 9, line=11),
            PM(
                [
                    ("Attn To:  ", "Arial", 8, False, False, None),
                    ("Mohamad Fauzi", "Courier New", 9, False, False, None),
                ],
                line=12,
                before=8,
            ),
        ],
        valign="top",
        margins=(40, 40, 140, 40),
    )
    write_runs(
        left.cell(1, 0),
        [
            P("Vendor No", "Arial", 8, align="center", line=9),
            P("20009041", "Courier New", 9, align="center", line=11),
        ],
        valign="center",
        margins=(20, 20, 20, 20),
    )
    write_runs(
        left.cell(1, 1),
        [
            P("Tel. No", "Arial", 8, align="center", line=9),
            P("60389955000", "Courier New", 9, align="center", line=11),
        ],
        valign="center",
        margins=(20, 20, 20, 20),
    )
    write_runs(
        left.cell(1, 2),
        [
            P("Fax No", "Arial", 8, align="center", line=9),
            P("", "Courier New", 9, align="center", line=11),
        ],
        valign="center",
        margins=(20, 20, 20, 20),
    )

    # RIGHT nested: deliver / attn / delivery / payment
    right = add_table(
        outer.cell(0, 1),
        4,
        2,
        [DEL_W, EXP_W],
        borders=True,
        cell_mar=(20, 20, 40, 40),
        nested=True,
    )
    set_row_height(right.rows[0], H_DELIVER, "exact")
    set_row_height(right.rows[1], H_ATTN, "exact")
    set_row_height(right.rows[2], H_DELIVERY, "exact")
    set_row_height(right.rows[3], H_PAYMENT, "exact")

    write_runs(
        right.cell(0, 0),
        [
            P("Deliver To", "Arial", 8, align="center", line=10),
            P("MIMOS Solutions Sdn Bhd", "Courier New", 8, align="left", line=10, before=2),
            P("Technology Park Malaysia", "Courier New", 8, align="left", line=10),
            P("57000 Kuala Lumpur.", "Courier New", 8, align="left", line=10),
        ],
        valign="top",
        margins=(40, 20, 80, 40),
    )
    write_runs(
        right.cell(0, 1),
        [
            P("Expected Delivery", "Arial", 8, align="center", line=10),
            P(" 24.06.2026", "Courier New", 9, align="center", line=11),
            P("Completion Work", "Arial", 8, align="center", line=10, before=2),
            P(" 24.06.2026", "Courier New", 9, align="center", line=11),
        ],
        valign="top",
        margins=(40, 20, 20, 20),
    )

    write_runs(
        right.cell(1, 0),
        [
            P("Attn to (Requestor)", "Arial", 8, align="center", line=10),
            P("Saszwani", "Courier New", 9, align="center", line=11),
        ],
        valign="center",
        margins=(20, 20, 20, 20),
    )
    write_runs(
        right.cell(1, 1),
        [
            P("Tel. or Ext. no.", "Arial", 8, align="center", line=10),
            P("-", "Courier New", 9, align="center", line=11),
        ],
        valign="center",
        margins=(20, 20, 20, 20),
    )

    deliv = right.cell(2, 0).merge(right.cell(2, 1))
    write_runs(
        deliv,
        [
            P("Delivery Terms", "Arial", 8, align="center", line=10),
            P("Service Completion at MIMOS Solutions", "Courier New", 9, align="center", line=11),
        ],
        valign="center",
        margins=(20, 20, 40, 40),
    )

    pay = right.cell(3, 0).merge(right.cell(3, 1))
    apply_inner_grid(right)
    write_runs(
        pay,
        [
            P("Payment Terms", "Arial", 8, align="center", line=10),
            P("30 Days Net After Receipt of Undisputed", "Courier New", 9, align="center", line=11),
            P("Invoice", "Courier New", 9, align="center", line=11),
        ],
        valign="center",
        margins=(20, 20, 40, 40),
    )
    squeeze_following(doc)


def build_pr_row(doc: Document):
    t = add_table(doc, 1, 3, [PR_W, BUY_W, RFQ_W], borders=True, cell_mar=(20, 20, 40, 40))
    set_row_height(t.rows[0], H_PR, "exact")
    write_runs(
        t.cell(0, 0),
        [
            P("Purchase Requisition(PR) No.", "Arial", 8, align="center", line=10),
            P("6000000481", "Courier New", 9, align="center", line=11),
        ],
        valign="center",
        margins=(20, 20, 20, 20),
    )
    write_runs(
        t.cell(0, 1),
        [
            P("Buyers Ref /  Tel. or Ext No.", "Arial", 8, align="center", line=10),
            P("MiTecSolu Proc.Dep / 56009/55923", "Courier New", 9, align="center", line=11),
        ],
        valign="center",
        margins=(20, 20, 20, 20),
    )
    write_runs(
        t.cell(0, 2),
        [
            P(" Refer to RFQ / ESS / SebutHarga / Tender No", "Arial", 8, align="center", line=10),
            P("34512", "Courier New", 10, align="center", line=12),
        ],
        valign="center",
        margins=(20, 20, 20, 20),
    )
    squeeze_following(doc)


def build_supply_banner(doc: Document):
    t = add_table(doc, 1, 1, [CONTENT_W], borders=True, cell_mar=(0, 0, 20, 20))
    set_row_height(t.rows[0], H_SUPPLY_BANNER, "exact")
    set_cell_shading(t.cell(0, 0), PURPLE)
    write_runs(
        t.cell(0, 0),
        [
            P(
                "PLEASE SUPPLY AND DELIVER PROMPTLY THE GOODS AND SERVICES AS SHOWN BELOW IN ACCORDANCE WITH THE TERMS AND CONDITIONS OF PURCHASE",
                "Arial",
                5.5,
                bold=True,
                color=WHITE,
                align="center",
                line=6.5,
            ),
            P(
                "SHOWN ON THE FACE AND REVERSE SIDE OF THIS PURCHASE ORDER AND WHERE APPLICABLE APPENDENCIES WHICH ARE MADE PART OF THIS PURCHASE ORDER",
                "Arial",
                5.5,
                bold=True,
                color=WHITE,
                align="center",
                line=6.5,
            ),
        ],
        valign="center",
        margins=(8, 8, 20, 20),
    )
    squeeze_following(doc)


def _item_header_cell(cell, lines):
    write_runs(cell, lines, valign="center", margins=(20, 10, 20, 20))


def build_items_table(doc: Document, page: int):
    """
    Outer 4-col: Item | Description | Quantity | Rest
    Rest nested:
      header 4-col Unit/Per/UP/TP
      body 4-col
      total 4-col (unit+per empty / TOTAL / amount-or-continue)
      approval 2-col
    Item/Desc/Qty vertically merged across body+total+approval.
    """
    outer = add_table(
        doc, 2, 4, [ITEM_W, DESC_W, QTY_W, REST_W], borders=True, cell_mar=(20, 20, 20, 20)
    )
    set_row_height(outer.rows[0], H_ITEM_HDR, "exact")
    set_row_height(outer.rows[1], H_ITEM_BODY, "exact")

    # ----- header -----
    _item_header_cell(
        outer.cell(0, 0),
        [
            P("Item", "Arial", 8, align="center", line=9),
            P("No", "Arial", 8, align="center", line=9),
        ],
    )
    _item_header_cell(
        outer.cell(0, 1),
        [P("Description of Goods and/or Services", "Arial", 8, align="center", line=10)],
    )
    _item_header_cell(
        outer.cell(0, 2),
        [P("Quantity", "Arial", 8, align="center", line=10)],
    )

    rest_hdr = add_table(
        outer.cell(0, 3),
        1,
        4,
        [UNIT_W, PER_W, UP_W, TP_W],
        borders=True,
        cell_mar=(10, 10, 10, 10),
        nested=True,
    )
    set_row_height(rest_hdr.rows[0], H_ITEM_HDR, "exact")
    _item_header_cell(rest_hdr.cell(0, 0), [P("Unit", "Arial", 8, align="center", line=9)])
    _item_header_cell(rest_hdr.cell(0, 1), [P("Per", "Arial", 8, align="center", line=9)])
    write_runs(
        rest_hdr.cell(0, 2),
        [
            P("Unit Price", "Arial", 8, align="center", line=9),
            PM(
                [
                    ("(Currency  ", "Arial", 8, False, False, None),
                    ("MYR", "Courier New", 10, False, False, None),
                    (")", "Arial", 8, False, False, None),
                ],
                align="center",
                line=11,
            ),
        ],
        valign="center",
        margins=(10, 10, 10, 10),
    )
    write_runs(
        rest_hdr.cell(0, 3),
        [
            P("Total Price", "Arial", 8, align="center", line=9),
            PM(
                [
                    ("(Currency  ", "Arial", 8, False, False, None),
                    ("MYR", "Courier New", 10, False, False, None),
                    (")", "Arial", 8, False, False, None),
                ],
                align="center",
                line=11,
            ),
        ],
        valign="center",
        margins=(10, 10, 10, 10),
    )

    # ----- body (one tall row) with nested rest -----
    if page == 1:
        write_runs(
            outer.cell(1, 0),
            [P("010", "Courier New", 8, align="center", line=10, before=8)],
            valign="top",
            margins=(40, 20, 20, 20),
        )
        desc_lines = [
            "7500936   - Oth-Training",
            "Programme Title: AI Training",
            "for Work Efficiency (maximum",
            "participant: 10 pax)",
            "Package inclusive of:",
            "- Training Materials",
            "- Certificate of Completion.",
            "",
            "Note :",
            "1. Reference on this PO is made",
            "   to your RFQ ref.:",
            "   MSSB-01/2026 dated 5 May",
            "   2026.",
            "",
            "2. The delivery date for the",
            "   items shall refer to the",
            "   delivery period in your",
            "   quotation or earlier.",
        ]
        paras = []
        for i, ln in enumerate(desc_lines):
            sz = 9 if (ln.startswith("Note") or ln.startswith("1.") or ln.startswith("2.") or ln.startswith("   ")) else 8
            if ln == "":
                paras.append(P(" ", "Courier New", 8, line=8))
            else:
                paras.append(P(ln, "Courier New", sz, line=sz + 1, before=0))
        write_runs(outer.cell(1, 1), paras, valign="top", margins=(40, 20, 40, 20))
        write_runs(
            outer.cell(1, 2),
            [P("1", "Courier New", 8, align="center", line=10, before=8)],
            valign="top",
            margins=(40, 20, 20, 20),
        )
    else:
        write_runs(outer.cell(1, 0), [P("", "Courier New", 8, line=10)], valign="top", margins=(40, 20, 20, 20))
        desc_lines = [
            "3. Supplier is reminded to",
            "   adhere to the terms and",
            "   condition stipulated in this",
            "   Purchase Order. Kindly always",
            "   refer to the PO terms and",
            "   conditions.",
        ]
        write_runs(
            outer.cell(1, 1),
            [P(ln, "Courier New", 9, line=10) for ln in desc_lines],
            valign="top",
            margins=(40, 20, 40, 20),
        )
        write_runs(outer.cell(1, 2), [P("", "Courier New", 8, line=10)], valign="top", margins=(40, 20, 20, 20))

    rest = add_table(
        outer.cell(1, 3),
        3,
        4,
        [UNIT_W, PER_W, UP_W, TP_W],
        borders=True,
        cell_mar=(20, 20, 20, 20),
        nested=True,
    )
    set_row_height(rest.rows[0], H_PRICE_BODY, "exact")
    set_row_height(rest.rows[1], H_TOTAL, "exact")
    set_row_height(rest.rows[2], H_APPROVAL, "exact")

    if page == 1:
        write_runs(
            rest.cell(0, 0),
            [P("lot", "Courier New", 8, align="center", line=10, before=8)],
            valign="top",
            margins=(40, 20, 10, 10),
        )
        write_runs(
            rest.cell(0, 1),
            [P("1", "Courier New", 8, align="center", line=10, before=8)],
            valign="top",
            margins=(40, 20, 10, 10),
        )
        write_runs(
            rest.cell(0, 2),
            [P("5,000.00", "Courier New", 8, align="right", line=10, before=8)],
            valign="top",
            margins=(40, 20, 20, 60),
        )
        write_runs(
            rest.cell(0, 3),
            [P("5,000.00", "Courier New", 8, align="right", line=10, before=8)],
            valign="top",
            margins=(40, 20, 20, 60),
        )
        # Unit/Per continue through the TOTAL row (no extra horizontal rule)
        rest.cell(0, 0).merge(rest.cell(1, 0))
        rest.cell(0, 1).merge(rest.cell(1, 1))
        apply_inner_grid(rest)
        # TOTAL / Continue Next Page
        set_cell_shading(rest.cell(1, 2), PURPLE_TOTAL)
        write_runs(
            rest.cell(1, 2),
            [P("TOTAL", "Arial", 14, bold=True, color=WHITE, align="center", line=16)],
            valign="center",
            margins=(0, 0, 10, 10),
        )
        write_runs(
            rest.cell(1, 3),
            [
                P("Continue Next", "Courier New", 9, align="center", line=11),
                P("Page", "Courier New", 9, align="center", line=11),
            ],
            valign="center",
            margins=(20, 20, 20, 20),
        )
    else:
        write_runs(rest.cell(0, 0), [], valign="top", margins=(0, 0, 0, 0))
        write_runs(rest.cell(0, 1), [], valign="top", margins=(0, 0, 0, 0))
        write_runs(
            rest.cell(0, 2),
            [
                P("Sub Total", "Courier New", 9, align="left", line=11),
                P("Service Tax", "Courier New", 9, align="left", line=11),
                P("Sales Tax", "Courier New", 9, align="left", line=11),
            ],
            valign="bottom",
            margins=(20, 40, 40, 20),
        )
        write_runs(
            rest.cell(0, 3),
            [
                P("5,000.00", "Courier New", 9, align="right", line=11),
                P("400.00", "Courier New", 9, align="right", line=11),
                P("0.00", "Courier New", 9, align="right", line=11),
            ],
            valign="bottom",
            margins=(20, 40, 20, 40),
        )
        rest.cell(0, 0).merge(rest.cell(1, 0))
        rest.cell(0, 1).merge(rest.cell(1, 1))
        apply_inner_grid(rest)
        set_cell_shading(rest.cell(1, 2), PURPLE_TOTAL)
        write_runs(
            rest.cell(1, 2),
            [P("TOTAL", "Arial", 14, bold=True, color=WHITE, align="center", line=16)],
            valign="center",
            margins=(0, 0, 10, 10),
        )
        write_runs(
            rest.cell(1, 3),
            [P("5,400.00", "Courier New", 9, align="right", line=12)],
            valign="center",
            margins=(20, 20, 20, 40),
        )

    # approval row: rebuild as 2-col nested covering the 4-col grid
    # Merge all 4 cells of row 2 then insert 2-col table
    appr_host = rest.cell(2, 0).merge(rest.cell(2, 3))
    appr = add_table(appr_host, 1, 2, [CHK_W, APPR_W], borders=True, cell_mar=(20, 20, 20, 20))
    set_row_height(appr.rows[0], H_APPROVAL, "exact")
    write_runs(
        appr.cell(0, 0),
        [
            P("Checked by Buyer", "Arial", 8, align="center", line=10),
            P("ROSSIDAWATI MANSOR", "Courier New", 9, align="center", line=11, before=4),
        ],
        valign="top",
        margins=(40, 20, 20, 20),
    )
    write_runs(
        appr.cell(0, 1),
        [
            P("Approved By", "Arial", 8, align="center", line=10),
            P("WAN MUHAMMAD WAN UMAR", "Courier New", 9, align="center", line=11, before=2),
            P("Head of Procurement", "Courier New", 9, align="center", line=11),
        ],
        valign="top",
        margins=(40, 20, 20, 20),
    )
    squeeze_following(doc)


ACK_LINES = [
    ("1. ACKNOWLEDGEMENT COPY.", True),
    (
        "MIMOS Group of Companies (MIMOS) requires acknowledgment and ot this purchase order (PO) via email within three (3) working days",
        False,
    ),
    (
        "for acceptance/rejection MIMOS' offer. However, in the event Supplier fails to provide MIMOS with a email acknowledgement within",
        False,
    ),
    (
        "three  (3) working days ARO and fails to provide MIMOS with a email notice of rejections ofthe order or specific modification of",
        False,
    ),
    (
        "a  particular  term  &  condition within that three(3) working days, MIMOS and Supplier hereby agree that all prices and all the",
        False,
    ),
    (
        "terms & conditions as expressed herein and its attachment/appendices, if any, shall apply in full.",
        False,
    ),
    ("", False),
    ("2. RECIPIENT", True),
    (
        "Copies  of  this  purchase  order  (PO) will also be extended to Requestor,MIMOS' Group Finance and MIMOS' Group Procurement for",
        False,
    ),
    ("their acknowledgment and filing purpose.", False),
    ("", False),
    ("3.COMPULSORY DOCUMENTS UPON SUBMISSION OF PAYMENT.", True),
    (
        "You  are required to attach this Purchase Order (PO) with the authorized personnel details, signature and company stamp together",
        False,
    ),
    (
        "with  the  related  supporting  documents  i.e.  Delivery  Order  (DO) with serial numbers of goods delivered clearly indicated,",
        False,
    ),
    (
        "Certificate of Completion and undisputed invoice to MIMOS Group Finance.",
        False,
    ),
    ("", False),
    (
        "4.This is a computer generated document and does not require a signature by MIMOS.",
        True,
    ),
]


def build_ack(doc: Document):
    t = add_table(doc, 1, 1, [CONTENT_W], borders=False, cell_mar=(0, 0, 0, 0))
    set_row_height(t.rows[0], H_ACK, "exact")
    no_cell_borders(t.cell(0, 0))
    paras = []
    for text, bold in ACK_LINES:
        if text == "":
            paras.append(P(" ", "Courier New", 7, line=8))
        else:
            paras.append(P(text, "Courier New", 7, bold=bold, line=8.2))
    write_runs(t.cell(0, 0), paras, valign="top", margins=(0, 0, 0, 0))
    squeeze_following(doc)


def build_signature(doc: Document):
    t = add_table(doc, 1, 3, [179.53, 179.53, 179.54], borders=False, cell_mar=(0, 0, 0, 0))
    set_row_height(t.rows[0], H_SIGN, "exact")
    for c in range(3):
        no_cell_borders(t.cell(0, c))
    write_runs(
        t.cell(0, 0),
        [
            P("................................", "Arial", 8, line=9),
            P("Name of personnel", "Courier New", 7, line=8),
        ],
        valign="top",
        margins=(0, 0, 0, 0),
    )
    write_runs(
        t.cell(0, 1),
        [
            P(".................................................", "Arial", 8, align="center", line=9),
            P("Authorized Signature & Date", "Courier New", 7, align="center", line=8),
        ],
        valign="top",
        margins=(0, 0, 0, 0),
    )
    write_runs(
        t.cell(0, 2),
        [
            P("  .........................", "Arial", 8, align="right", line=9),
            P(" Company Stamp", "Courier New", 7, align="right", line=8),
        ],
        valign="top",
        margins=(0, 0, 0, 0),
    )
    squeeze_following(doc)


def build_po_page(doc: Document, page: int, logo_path: Path):
    label = "1/02" if page == 1 else "2/02"
    build_header(doc, label, logo_path)
    build_supplier_block(doc)
    spacer(doc, H_GAP_PR)
    build_pr_row(doc)
    build_supply_banner(doc)
    build_items_table(doc, page)
    spacer(doc, H_GAP_ACK)
    build_ack(doc)
    spacer(doc, H_GAP_SIGN)
    build_signature(doc)


# ---------------------------------------------------------------------------
# Terms & Conditions (pages 3-6) — exact PDF line content
# ---------------------------------------------------------------------------
def load_terms_pages():
    pages = {}
    for n in (3, 4, 5, 6):
        path = SPANS_DIR / f"page-{n}-spans.json"
        spans = json.loads(path.read_text())
        items = [s for s in spans if "text" in s]
        items.sort(key=lambda s: (round(s["bbox"][1], 1), s["bbox"][0]))
        lines = []
        cur_y = None
        buf = []
        for s in items:
            y = round(s["bbox"][1], 1)
            rec = {
                "text": s["text"],
                "font": s["font"],
                "size": s["size"],
                "bold": "Bold" in s["font"],
                "y": y,
                "x": s["bbox"][0],
            }
            if cur_y is None:
                cur_y = y
            if abs(y - cur_y) > 2.0:
                buf.sort(key=lambda r: r["x"])
                lines.append({"y": cur_y, "runs": buf})
                buf = [rec]
                cur_y = y
            else:
                buf.append(rec)
        if buf:
            buf.sort(key=lambda r: r["x"])
            lines.append({"y": cur_y, "runs": buf})
        pages[n] = lines
    return pages


def build_terms(doc: Document, pages):
    # new section with slightly different top margin so title sits at ~28.7pt
    # We keep same section and just page-break; title is repeated in body.
    first = True
    for n in (3, 4, 5, 6):
        if not first:
            page_break(doc)
        first = False
        lines = pages[n]
        # title
        p = doc.add_paragraph()
        set_para(p, align="center", before=0, after=16, line_pt=14, exact=True)
        run = p.add_run("TERMS & CONDITIONS OF PURCHASE ORDER")
        set_run_font(run, "Arial", 12, bold=True)

        body_lines = [ln for ln in lines if "TERMS & CONDITIONS" not in "".join(r["text"] for r in ln["runs"])]
        prev_y = None
        for ln in body_lines:
            y = ln["y"]
            extra = 0
            if prev_y is not None:
                gap = y - prev_y
                if gap > 12:
                    extra = max(0, gap - 9)
            prev_y = y
            p = doc.add_paragraph()
            set_para(p, align="left", before=extra, after=0, line_pt=9, exact=True)
            # keep double spaces
            for r in ln["runs"]:
                text = r["text"]
                # skip empty bold spacer that is just a single space between number and title if we keep it
                font = "Courier New"
                run = p.add_run(text)
                set_run_font(run, font, 8, bold=r["bold"])
                # preserve spaces
                if text.startswith(" ") or text.endswith(" ") or "  " in text:
                    run._element.find(qn("w:t")).set(qn("xml:space"), "preserve")


def add_section_break(doc: Document):
    """Insert next-page section break by splitting sectPr — used if needed."""
    p = doc.add_paragraph()
    squeeze_para(p, 1)
    pPr = p._p.get_or_add_pPr()
    sect = OxmlElement("w:sectPr")
    # copy current section later; simpler to use page break
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)


def build():
    if not LOGO_SRC.exists():
        raise SystemExit(f"Logo not found: {LOGO_SRC}")
    logo_copy = FORENSIC / "MIMOS_Solutions_LOGO.png"
    if not logo_copy.exists():
        shutil.copy2(LOGO_SRC, logo_copy)

    doc = Document()
    set_normal_style(doc)
    configure_section(doc.sections[0])

    # remove the default empty paragraph
    body = doc.element.body
    for child in list(body):
        if child.tag == qn("w:p"):
            body.remove(child)

    build_po_page(doc, 1, LOGO_SRC)
    page_break(doc)
    build_po_page(doc, 2, LOGO_SRC)
    page_break(doc)

    terms = load_terms_pages()
    build_terms(doc, terms)

    # compatibility
    settings = doc.settings.element
    compat = _get_or_add(settings, "w:compat")
    compatSetting = OxmlElement("w:compatSetting")
    compatSetting.set(qn("w:name"), "compatibilityMode")
    compatSetting.set(qn("w:uri"), "http://schemas.microsoft.com/office/word")
    compatSetting.set(qn("w:val"), "15")
    compat.append(compatSetting)

    OUT_DOCX.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUT_DOCX))
    print(f"Wrote {OUT_DOCX}")
    return OUT_DOCX


if __name__ == "__main__":
    build()
