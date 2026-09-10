#!/usr/bin/env python3
"""Render the generated DOCX to page PNGs using stored table geometry.

This is not Microsoft Word's layout engine. It honours the explicit
column widths, exact row heights, shading, nested tables, fonts and
page breaks written into the DOCX so we can visually QA against the
reference PDF. Word-opened output will be very close when Arial and
Courier New are available.
"""
from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree
from PIL import Image, ImageDraw, ImageFont

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}
W = "{%s}" % NS["w"]

ROOT = Path(__file__).resolve().parent
DOCX = ROOT / "MIMOS_Academy_Blank_Purchase_Order_Template.docx"
OUT_DIR = ROOT / "06_DOCX_RENDER"

# 150 dpi to match 01_REFERENCE_RENDER
DPI = 150
SCALE = DPI / 72.0  # px per pt
PAGE_W_PT = 595.27
PAGE_H_PT = 841.89

FONT_DIR = Path("/usr/share/fonts/truetype/dejavu")
FONTS = {
    ("Arial", False, False): FONT_DIR / "DejaVuSans.ttf",
    ("Arial", True, False): FONT_DIR / "DejaVuSans-Bold.ttf",
    ("Arial", False, True): FONT_DIR / "DejaVuSans.ttf",
    ("Courier New", False, False): FONT_DIR / "DejaVuSansMono.ttf",
    ("Courier New", True, False): FONT_DIR / "DejaVuSansMono-Bold.ttf",
    ("Courier New", False, True): FONT_DIR / "DejaVuSansMono.ttf",
}

_font_cache: dict = {}


def font(name: str, size_pt: float, bold=False, italic=False) -> ImageFont.FreeTypeFont:
    key = (name, bool(bold), bool(italic), round(size_pt, 2))
    if key in _font_cache:
        return _font_cache[key]
    # map unknown to Arial/Courier
    family = "Courier New" if "Courier" in (name or "") else "Arial"
    path = FONTS.get((family, bool(bold), bool(italic))) or FONTS[("Arial", False, False)]
    px = max(1, size_pt * SCALE * 0.96)
    f = ImageFont.truetype(str(path), size=int(round(px)))
    _font_cache[key] = f
    return f


def twips_to_pt(v) -> float:
    try:
        return float(v) / 20.0
    except Exception:
        return 0.0


def emu_to_pt(v) -> float:
    try:
        return float(v) / 12700.0
    except Exception:
        return 0.0


def child(el, tag):
    return el.find(f"w:{tag}", NS) if el is not None else None


def children(el, tag):
    if el is None:
        return []
    return el.findall(f"w:{tag}", NS)


def val(el, attr="val"):
    if el is None:
        return None
    return el.get(f"{{{NS['w']}}}{attr}")


def hex_to_rgb(h):
    h = (h or "000000").lstrip("#")
    if len(h) != 6:
        return (0, 0, 0)
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


class DocxRenderer:
    def __init__(self, path: Path):
        self.zf = zipfile.ZipFile(path)
        self.doc = etree.fromstring(self.zf.read("word/document.xml"))
        self.rels = self._load_rels()
        self.images = {}
        self.pages: list[Image.Image] = []
        self.im = None
        self.dr = None
        self.y = 0.0
        self.margin_l = 28.35
        self.margin_t = 28.35
        self.margin_r = 28.32
        self.margin_b = 8.5
        self._load_sect()

    def _load_rels(self):
        rels = {}
        try:
            xml = etree.fromstring(self.zf.read("word/_rels/document.xml.rels"))
        except KeyError:
            return rels
        ns = {"pr": "http://schemas.openxmlformats.org/package/2006/relationships"}
        for rel in xml.findall("pr:Relationship", ns):
            rels[rel.get("Id")] = rel.get("Target")
        return rels

    def _load_image(self, rid):
        if rid in self.images:
            return self.images[rid]
        target = self.rels.get(rid)
        if not target:
            return None
        path = "word/" + target.lstrip("/")
        if target.startswith(".."):
            path = target.replace("../", "")
        try:
            data = self.zf.read(path)
            im = Image.open(BytesIO(data)).convert("RGBA")
            self.images[rid] = im
            return im
        except Exception:
            return None

    def _load_sect(self):
        sect = self.doc.find(".//w:sectPr", NS)
        if sect is None:
            return
        pgMar = child(sect, "pgMar")
        if pgMar is not None:
            self.margin_t = twips_to_pt(pgMar.get(f"{{{NS['w']}}}top"))
            self.margin_b = twips_to_pt(pgMar.get(f"{{{NS['w']}}}bottom"))
            self.margin_l = twips_to_pt(pgMar.get(f"{{{NS['w']}}}left"))
            self.margin_r = twips_to_pt(pgMar.get(f"{{{NS['w']}}}right"))

    def new_page(self):
        w = int(round(PAGE_W_PT * SCALE))
        h = int(round(PAGE_H_PT * SCALE))
        self.im = Image.new("RGB", (w, h), (255, 255, 255))
        self.dr = ImageDraw.Draw(self.im)
        self.pages.append(self.im)
        self.y = self.margin_t

    def px(self, pt):
        return pt * SCALE

    def draw_rect(self, x, y, w, h, fill=None, outline=None, width=1):
        if w <= 0 or h <= 0:
            return
        x0, y0 = self.px(x), self.px(y)
        x1, y1 = self.px(x + w), self.px(y + h)
        kw = {}
        if fill:
            kw["fill"] = fill
        if outline:
            kw["outline"] = outline
            kw["width"] = max(1, int(round(width)))
        if kw:
            self.dr.rectangle([x0, y0, x1, y1], **kw)

    def draw_line(self, x0, y0, x1, y1, color=(0, 0, 0), width=1):
        self.dr.line(
            [self.px(x0), self.px(y0), self.px(x1), self.px(y1)],
            fill=color,
            width=max(1, int(round(width))),
        )

    def _border_spec(self, el):
        if el is None:
            return None
        v = val(el) or el.get(f"{{{NS['w']}}}val")
        if v in (None, "nil", "none"):
            return None
        sz = float(el.get(f"{{{NS['w']}}}sz") or 4) / 8.0
        col = el.get(f"{{{NS['w']}}}color") or "000000"
        if col == "auto":
            col = "000000"
        return (hex_to_rgb(col), sz)

    def parse_borders(self, tcPr, tblBorders, *, row_i=0, col_i=0, n_rows=1, n_cols=1):
        """Return dict edge -> (color, sz_pt) or None if nil."""
        result = {}
        src = child(tcPr, "tcBorders") if tcPr is not None else None
        insideH = self._border_spec(child(tblBorders, "insideH") if tblBorders is not None else None)
        insideV = self._border_spec(child(tblBorders, "insideV") if tblBorders is not None else None)
        for edge in ("top", "left", "bottom", "right"):
            el = child(src, edge) if src is not None else None
            if el is not None:
                # explicit cell border (including nil) wins
                result[edge] = self._border_spec(el)
                continue
            spec = None
            if tblBorders is not None:
                spec = self._border_spec(child(tblBorders, edge))
            if spec is None:
                if edge == "bottom" and insideH and row_i < n_rows - 1:
                    spec = insideH
                elif edge == "top" and insideH and row_i > 0:
                    spec = insideH
                elif edge == "right" and insideV and col_i < n_cols - 1:
                    spec = insideV
                elif edge == "left" and insideV and col_i > 0:
                    spec = insideV
            result[edge] = spec
        return result

    def cell_fill(self, tcPr):
        if tcPr is None:
            return None
        shd = child(tcPr, "shd")
        if shd is None:
            return None
        fill = shd.get(f"{{{NS['w']}}}fill")
        if not fill or fill in ("auto", "FFFFFF"):
            return None
        return hex_to_rgb(fill)

    def cell_margins(self, tcPr, tblMar):
        def read(src, name, default):
            el = child(src, name) if src is not None else None
            if el is None:
                return default
            return twips_to_pt(el.get(f"{{{NS['w']}}}w"))

        tcMar = child(tcPr, "tcMar") if tcPr is not None else None
        src = tcMar if tcMar is not None else tblMar
        return {
            "top": read(src, "top", 1.0),
            "bottom": read(src, "bottom", 1.0),
            "left": read(src, "left", 2.0),
            "right": read(src, "right", 2.0),
        }

    def run_props(self, r):
        rPr = child(r, "rPr")
        name = "Arial"
        size = 8.0
        bold = False
        italic = False
        color = (0, 0, 0)
        if rPr is not None:
            rFonts = child(rPr, "rFonts")
            if rFonts is not None:
                name = rFonts.get(f"{{{NS['w']}}}ascii") or name
            sz = child(rPr, "sz")
            if sz is not None:
                size = float(val(sz) or 16) / 2.0
            bold = child(rPr, "b") is not None
            italic = child(rPr, "i") is not None
            c = child(rPr, "color")
            if c is not None and val(c) and val(c) != "auto":
                color = hex_to_rgb(val(c))
        return name, size, bold, italic, color

    def para_align(self, p):
        pPr = child(p, "pPr")
        if pPr is None:
            return "left"
        jc = child(pPr, "jc")
        return val(jc) or "left"

    def para_spacing(self, p):
        pPr = child(p, "pPr")
        before = after = 0.0
        line = None
        if pPr is not None:
            sp = child(pPr, "spacing")
            if sp is not None:
                before = twips_to_pt(sp.get(f"{{{NS['w']}}}before") or 0)
                after = twips_to_pt(sp.get(f"{{{NS['w']}}}after") or 0)
                ln = sp.get(f"{{{NS['w']}}}line")
                rule = sp.get(f"{{{NS['w']}}}lineRule")
                if ln is not None and rule == "exact":
                    line = twips_to_pt(ln)
        return before, after, line

    def draw_text_line(self, runs, x, y, w, align="left"):
        """Draw a sequence of runs on one line inside [x, x+w] at top y (pt)."""
        pieces = []
        total = 0.0
        for name, size, bold, italic, color, text in runs:
            if not text:
                continue
            f = font(name, size, bold, italic)
            bbox = f.getbbox(text)
            tw = (bbox[2] - bbox[0]) / SCALE
            pieces.append((f, color, text, tw, size))
            total += tw
        if align == "center":
            cx = x + max(0, (w - total) / 2)
        elif align == "right":
            cx = x + max(0, w - total)
        else:
            cx = x
        for f, color, text, tw, size in pieces:
            # y is top of line box; draw at baseline ~ 0.8 * size
            self.dr.text(
                (self.px(cx), self.px(y)),
                text,
                font=f,
                fill=color,
            )
            cx += tw

    def iter_para_runs(self, p):
        runs_out = []
        has_break = False
        has_drawing = False
        drawing_rid = None
        drawing_w = drawing_h = 0
        for r in children(p, "r"):
            # page break?
            for br in r.findall("w:br", NS):
                if br.get(f"{{{NS['w']}}}type") == "page":
                    has_break = True
            drawing = r.find(".//w:drawing", NS)
            if drawing is None:
                drawing = r.find(".//{%s}drawing" % NS["w"])
            blip = r.find(".//a:blip", NS)
            if blip is not None:
                has_drawing = True
                drawing_rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                ext = r.find(".//wp:extent", NS)
                if ext is not None:
                    drawing_w = emu_to_pt(ext.get("cx"))
                    drawing_h = emu_to_pt(ext.get("cy"))
            texts = []
            for t in r.findall("w:t", NS):
                texts.append(t.text or "")
            name, size, bold, italic, color = self.run_props(r)
            if texts:
                runs_out.append((name, size, bold, italic, color, "".join(texts)))
        return runs_out, has_break, has_drawing, drawing_rid, drawing_w, drawing_h

    def render_paragraphs_in_box(self, paras, x, y, w, h, valign="top"):
        """Render w:p elements inside a cell box. Returns nothing (clipped by box)."""
        # measure total height
        measured = []
        for p in paras:
            runs, has_break, has_drawing, rid, dw, dh = self.iter_para_runs(p)
            before, after, line = self.para_spacing(p)
            align = self.para_align(p)
            if has_drawing and dh:
                ph = before + dh + after
            else:
                ph = before + (line if line else 9) + after
            measured.append((p, runs, has_drawing, rid, dw, dh, before, after, line, align, ph))
        total_h = sum(m[-1] for m in measured)
        if valign == "center":
            cy = y + max(0, (h - total_h) / 2)
        elif valign == "bottom":
            cy = y + max(0, h - total_h)
        else:
            cy = y
        clip_bottom = y + h
        for p, runs, has_drawing, rid, dw, dh, before, after, line, align, ph in measured:
            cy += before
            if has_drawing and rid:
                im = self._load_image(rid)
                if im is not None and dw > 0 and dh > 0:
                    box = [int(self.px(x)), int(self.px(cy)), int(self.px(x + dw)), int(self.px(cy + dh))]
                    try:
                        resized = im.resize((max(1, box[2] - box[0]), max(1, box[3] - box[1])), Image.Resampling.LANCZOS)
                        self.im.paste(resized, (box[0], box[1]), resized if resized.mode == "RGBA" else None)
                    except Exception:
                        pass
                cy += dh + after
                continue
            line_h = line if line else 9
            if runs and cy < clip_bottom:
                self.draw_text_line(runs, x, cy, w, align)
            cy += line_h + after
            if cy > clip_bottom + 2:
                break

    def grid_widths(self, tbl):
        grid = child(tbl, "tblGrid")
        widths = []
        if grid is not None:
            for gc in children(grid, "gridCol"):
                widths.append(twips_to_pt(gc.get(f"{{{NS['w']}}}w")))
        return widths

    def row_height(self, tr, default=12.0):
        trPr = child(tr, "trPr")
        if trPr is None:
            return default
        trH = child(trPr, "trHeight")
        if trH is None:
            return default
        return twips_to_pt(trH.get(f"{{{NS['w']}}}val") or 0) or default

    def tbl_margins(self, tbl):
        tblPr = child(tbl, "tblPr")
        mar = child(tblPr, "tblCellMar") if tblPr is not None else None
        return mar

    def tbl_borders(self, tbl):
        tblPr = child(tbl, "tblPr")
        return child(tblPr, "tblBorders") if tblPr is not None else None

    def tc_grid_span(self, tc):
        tcPr = child(tc, "tcPr")
        if tcPr is None:
            return 1
        gs = child(tcPr, "gridSpan")
        if gs is None:
            return 1
        try:
            return int(val(gs) or 1)
        except Exception:
            return 1

    def tc_vmerge(self, tc):
        tcPr = child(tc, "tcPr")
        if tcPr is None:
            return None
        vm = child(tcPr, "vMerge")
        if vm is None:
            return None
        v = val(vm)
        return v or "continue"

    def tc_valign(self, tc):
        tcPr = child(tc, "tcPr")
        if tcPr is None:
            return "top"
        va = child(tcPr, "vAlign")
        return val(va) or "top"

    def render_table(self, tbl, origin_x, origin_y) -> float:
        widths = self.grid_widths(tbl)
        if not widths:
            return 0.0
        tbl_w = sum(widths)
        tblMar = self.tbl_margins(tbl)
        tblBorders = self.tbl_borders(tbl)
        y = origin_y
        rows = children(tbl, "tr")
        n_rows = len(rows)
        for row_i, tr in enumerate(rows):
            rh = self.row_height(tr)
            tcs = children(tr, "tc")
            col = 0
            x = origin_x
            n_cols = len(widths)
            for tc in tcs:
                span = self.tc_grid_span(tc)
                cw = sum(widths[col : col + span]) if col < len(widths) else 40
                tcPr = child(tc, "tcPr")
                fill = self.cell_fill(tcPr)
                if fill:
                    self.draw_rect(x, y, cw, rh, fill=fill)
                nested = children(tc, "tbl")
                paras = children(tc, "p")
                mar = self.cell_margins(tcPr, tblMar)
                ix = x + mar["left"]
                iy = y + mar["top"]
                iw = max(1, cw - mar["left"] - mar["right"])
                ih = max(1, rh - mar["top"] - mar["bottom"])
                if nested:
                    for nt in nested:
                        self.render_table(nt, x, y)
                else:
                    self.render_paragraphs_in_box(paras, ix, iy, iw, ih, valign=self.tc_valign(tc))
                b = self.parse_borders(
                    tcPr, tblBorders, row_i=row_i, col_i=col, n_rows=n_rows, n_cols=n_cols
                )
                if b.get("top"):
                    colr, sz = b["top"]
                    self.draw_line(x, y, x + cw, y, colr, max(1, sz * SCALE))
                if b.get("bottom"):
                    colr, sz = b["bottom"]
                    self.draw_line(x, y + rh, x + cw, y + rh, colr, max(1, sz * SCALE))
                if b.get("left"):
                    colr, sz = b["left"]
                    self.draw_line(x, y, x, y + rh, colr, max(1, sz * SCALE))
                if b.get("right"):
                    colr, sz = b["right"]
                    self.draw_line(x + cw, y, x + cw, y + rh, colr, max(1, sz * SCALE))
                x += cw
                col += span
            y += rh
        return y - origin_y

    def para_is_page_break(self, p):
        for r in children(p, "r"):
            for br in r.findall("w:br", NS):
                if br.get(f"{{{NS['w']}}}type") == "page":
                    return True
        return False

    def para_is_empty(self, p):
        texts = [t.text or "" for t in p.findall(".//w:t", NS)]
        return not "".join(texts).strip() and not self.para_is_page_break(p) and p.find(".//w:drawing", NS) is None

    def render_body_para(self, p):
        if self.para_is_page_break(p):
            self.new_page()
            return
        runs, has_break, has_drawing, rid, dw, dh, = (*self.iter_para_runs(p)[:],)
        # unpack properly
        runs, has_break, has_drawing, rid, dw, dh = self.iter_para_runs(p)
        if has_break:
            self.new_page()
            return
        before, after, line = self.para_spacing(p)
        align = self.para_align(p)
        self.y += before
        content_x = self.margin_l
        content_w = PAGE_W_PT - self.margin_l - self.margin_r
        if has_drawing and dh:
            im = self._load_image(rid)
            if im is not None:
                box = [
                    int(self.px(content_x)),
                    int(self.px(self.y)),
                    int(self.px(content_x + dw)),
                    int(self.px(self.y + dh)),
                ]
                resized = im.resize((max(1, box[2] - box[0]), max(1, box[3] - box[1])), Image.Resampling.LANCZOS)
                self.im.paste(resized, (box[0], box[1]), resized if resized.mode == "RGBA" else None)
            self.y += dh + after
            return
        if runs:
            self.draw_text_line(runs, content_x, self.y, content_w, align)
        self.y += (line if line else 9) + after

    def render(self):
        self.new_page()
        body = self.doc.find("w:body", NS)
        for el in list(body):
            tag = etree.QName(el).localname
            if tag == "tbl":
                used = self.render_table(el, self.margin_l, self.y)
                self.y += used
            elif tag == "p":
                self.render_body_para(el)
            elif tag == "sectPr":
                pass
        return self.pages


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    r = DocxRenderer(DOCX)
    pages = r.render()
    for i, im in enumerate(pages, 1):
        p = OUT_DIR / f"page-{i}.png"
        im.save(p, "PNG", optimize=True)
        print(f"wrote {p} {im.size}")
    print(f"total pages: {len(pages)}")


if __name__ == "__main__":
    main()
