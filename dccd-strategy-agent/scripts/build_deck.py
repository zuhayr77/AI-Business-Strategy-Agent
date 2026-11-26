"""Build a DCCD deck in the exact 46-882 lecture style from a JSON spec.

usage: python build_deck.py spec.json output.pptx [template.pptx]

Fixed type scale on every slide (same text type = same size everywhere):
  title 28 bold (red stage word + black text), question bullets 20 bold (Wingdings square),
  answer bullets 16 (round), key takeaways 16 red bold italic with a square bullet,
  table header 12 bold, table body 9-12 (only tables adapt to fit).
Line spacing 115% and the same paragraph spacing everywhere.
Every key takeaway sits one blank line below the content above it.
The script prints WARNING lines if text would overflow; shorten the spec and rebuild.
"""
import json
import math
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

# ---------- fixed style ----------
red = RGBColor(0xC0, 0x00, 0x00)
table_red = RGBColor(0xAF, 0x23, 0x17)
black = RGBColor(0, 0, 0)
grey = RGBColor(0x59, 0x59, 0x59)
light = RGBColor(0xF2, 0xF2, 0xF2)
white = RGBColor(0xFF, 0xFF, 0xFF)
hl_row = RGBColor(0xFB, 0xE3, 0xE1)
dot_grey = RGBColor(0x40, 0x4A, 0x5A)

title_pt = 28
q_pt = 20            # question / section bullets (square)
a_pt = 16            # answer bullets (round)
takeaway_pt = 16     # key takeaways and competitive advantage (red bold italic)
head_pt = 12         # table header
line_spacing = 1.15
space_before_q = 6   # points before each question bullet
space_before_a = 2   # points before each answer bullet
gap = 0.3            # one blank line (inches) between content and the takeaway
bottom_limit = 7.3   # inches; nothing should run past this
left_margin = 0.31
full_width = 12.7

here = os.path.dirname(os.path.abspath(__file__))
default_template = os.path.join(here, "..", "assets", "template.pptx")
warnings = []


# ---------- measurement ----------

def text_lines(text, width_in, size):
    clean = str(text).replace("**", "")
    per_line = max(8, int(width_in * 72 / (size * 0.47)))
    lines, cur = 1, 0
    for w in clean.split():
        add = len(w) + (1 if cur else 0)
        if cur + add > per_line:
            lines += 1
            cur = len(w)
        else:
            cur += add
    return lines


def item_height(level, text, width_in):
    if level == 0:
        size, indent, before = q_pt, 0.375, space_before_q
    elif level == "takeaway":
        size, indent, before = takeaway_pt, 0.375, space_before_q
    else:
        size, indent, before = a_pt, 0.875, space_before_a
    lines = text_lines(text, width_in - indent - 0.2, size)
    return lines * size * line_spacing * 1.2 / 72 + before / 72


def bullets_height(items, width_in):
    return sum(item_height(lv, t, width_in) for lv, t in items) + 0.1


def cell_height(text, width_in, size):
    return text_lines(text, width_in - 0.14, size) * size * 1.2 / 72 + 0.1


# ---------- helpers ----------

def drop_all_slides_but_first(prs):
    sld_ids = prs.slides._sldIdLst
    for sld_id in list(sld_ids)[1:]:
        prs.part.drop_rel(sld_id.get(qn("r:id")))
        sld_ids.remove(sld_id)


def content_layout(prs):
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name == "Title and Content":
                return layout
    return prs.slide_layouts[1]


small_words = {"a", "an", "the", "and", "but", "or", "nor", "for", "of", "in", "on", "at", "to",
               "by", "with", "vs", "vs.", "as", "per", "via"}


def title_case(text):
    """Capitalize every word except short connecting words (unless first, or after a colon)."""
    out = []
    after_colon = True
    for word in text.split(" "):
        core = word.strip("()?:,")
        low = core.lower()
        if not core:
            out.append(word)
            continue
        if core.isupper() and len(core) > 1:          # keep acronyms (VRI, DCCD, WTP)
            new = word
        elif low in small_words and not after_colon:
            new = word.replace(core, low, 1)
        elif "'" in core:                               # possessives: Nike's
            new = word.replace(core, core[0].upper() + core[1:], 1)
        else:
            new = word.replace(core, core[0].upper() + core[1:], 1)
        out.append(new)
        after_colon = word.endswith(":")
    return " ".join(out)


def new_content_slide(prs, stage, rest):
    slide = prs.slides.add_slide(content_layout(prs))
    for shape in list(slide.placeholders):
        if shape.placeholder_format.idx != 0:
            shape._element.getparent().remove(shape._element)
    title = slide.shapes.title
    title.left, title.top = Emu(279399), Emu(260080)
    title.width, title.height = Emu(10847779), Emu(668610)
    tf = title.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    for text, col in ((f"{stage}: ", red), (title_case(rest), black)):
        r = p.add_run()
        r.text = text
        r.font.size = Pt(title_pt)
        r.font.bold = True
        r.font.name = "Calibri"
        r.font.color.rgb = col
    return slide


def add_run(p, text, size, bold=False, italic=False, color=None):
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = "Calibri"
    r.font.color.rgb = color if color is not None else black
    return r


def add_rich(p, text, size, bold=False, italic=False, color=None):
    for i, part in enumerate(str(text).split("**")):
        if part:
            add_run(p, part, size, bold=bold or (i % 2 == 1), italic=italic, color=color)


def _bullet(p, level):
    pPr = p._p.get_or_add_pPr()
    for tag in ("a:lnSpc", "a:spcBef", "a:buFont", "a:buChar", "a:buNone"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    ln = etree.SubElement(pPr, qn("a:lnSpc"))
    etree.SubElement(ln, qn("a:spcPct")).set("val", str(int(line_spacing * 100000)))
    sb = etree.SubElement(pPr, qn("a:spcBef"))
    before = space_before_a if level == 1 else space_before_q
    etree.SubElement(sb, qn("a:spcPts")).set("val", str(before * 100))
    if level in (0, "takeaway"):
        pPr.set("marL", "342900")
        pPr.set("indent", "-342900")
        bf = etree.SubElement(pPr, qn("a:buFont"))
        bf.set("typeface", "Wingdings")
        bf.set("pitchFamily", "2")
        bf.set("charset", "2")
        etree.SubElement(pPr, qn("a:buChar")).set("char", "q")
    else:
        pPr.set("marL", "800100")
        pPr.set("lvl", "1")
        pPr.set("indent", "-342900")
        etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
        etree.SubElement(pPr, qn("a:buChar")).set("char", "•")


def bullets(slide, left, top, width, items):
    """items: (0, question) / (1, answer) / ('takeaway', text). Returns bottom (inches)."""
    h = bullets_height(items, width)
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_top = Inches(0)
    for i, (level, text) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        _bullet(p, level)
        if level == 0:
            add_rich(p, text, q_pt, bold=True)
        elif level == 1:
            add_rich(p, text, a_pt)
        else:
            add_rich(p, text, takeaway_pt, bold=True, italic=True, color=red)
    return top + h


def as_list(x):
    if not x:
        return []
    return x if isinstance(x, list) else [x]


def takeaway_lines(entries):
    return [("takeaway", f"{label}: {text}" if label else text) for label, text in entries]


def write_takeaways(slide, top, width, entries, name):
    """Red bold italic takeaways (lecture style), one blank line below the content."""
    if not entries:
        return top
    items = takeaway_lines(entries)
    bottom = bullets(slide, left_margin, top + gap, width, items)
    if bottom > bottom_limit:
        warnings.append(f"WARNING [{name}]: takeaway ends at {bottom:.2f}in (limit {bottom_limit}); shorten content above it")
    return bottom


def takeaways(obj, key="takeaway", label="Key takeaway"):
    return [(label, t) for t in as_list(obj.get(key))][:2]


def table(slide, left, top, col_widths, header, rows, max_bottom, name=""):
    """rows: list of (values, highlight). Body font shrinks 12 -> 9 until the table fits."""
    for body_pt in (13, 12, 11, 10):
        head_h = max(cell_height(h, w, head_pt) for h, w in zip(header, col_widths))
        heights = [max(cell_height(v, w, body_pt) for v, w in zip(vals, col_widths)) for vals, _ in rows]
        if top + head_h + sum(heights) <= max_bottom:
            break
    else:
        warnings.append(f"WARNING [{name}]: table too long even at 9pt; shorten cells")
    shape = slide.shapes.add_table(len(rows) + 1, len(header), Inches(left), Inches(top),
                                   Inches(sum(col_widths)), Inches(head_h + sum(heights)))
    tbl = shape.table
    style = shape._element.graphic.graphicData.tbl.tblPr.find(qn("a:tableStyleId"))
    if style is not None:
        style.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"
    for i, w in enumerate(col_widths):
        tbl.columns[i].width = Inches(w)
    tbl.rows[0].height = Inches(head_h)
    for r, h in enumerate(heights, start=1):
        tbl.rows[r].height = Inches(h)

    def fill(cell, text, size, bold, color, bg):
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
        cell.margin_left = cell.margin_right = Inches(0.07)
        cell.margin_top = cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        tf.clear()
        add_rich(tf.paragraphs[0], text, size, bold=bold, color=color)

    for c, h in enumerate(header):
        fill(tbl.cell(0, c), h, head_pt, True, white, table_red)
    for r, (vals, hl) in enumerate(rows, start=1):
        bg = hl_row if hl else (white if r % 2 == 1 else light)
        for c, v in enumerate(vals):
            fill(tbl.cell(r, c), v, body_pt, c == 0 or hl, red if (hl and c == 0) else black, bg)
    return top + head_h + sum(heights)


def reserve_for(entries, width):
    return bullets_height(takeaway_lines(entries), width) + gap if entries else 0


# ---------- positioning map ----------

def positioning_map(slide, m, left, top, w, h):
    """Lecture position test: y = increased WTP (attribute customers choose on),
    x = low cost (falls to the right), concave competitive frontier."""
    ax = left + 1.35
    ay_top = top + 0.35
    ay_bot = top + h - 0.55
    ax_right = left + w - 0.15
    for (x1, y1, x2, y2) in ((ax, ay_top, ax, ay_bot), (ax, ay_bot, ax_right, ay_bot)):
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        ln.line.color.rgb = black
        ln.line.width = Pt(1.5)
    yl = slide.shapes.add_textbox(Inches(left - 0.05), Inches(ay_top - 0.05), Inches(1.38), Inches(1.6))
    tf = yl.text_frame
    tf.word_wrap = True
    add_run(tf.paragraphs[0], "Increased WTP", 12, bold=True)
    add_run(tf.add_paragraph(), m.get("y_label", "Perceived quality"), 11, color=grey)
    xl = slide.shapes.add_textbox(Inches(ax_right - 3.4), Inches(ay_bot + 0.05), Inches(3.4), Inches(0.5))
    tf = xl.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    add_run(p, "Low cost  ", 12, bold=True)
    add_run(p, m.get("x_label", ""), 11, color=grey)
    fx0, fy0 = ax + 0.15, ay_top + 0.05
    fx1, fy1 = ax_right - 0.35, ay_bot - 0.05
    pts = [(fx0 + (fx1 - fx0) * math.sin(math.pi / 2 * k / 40), fy1 - (fy1 - fy0) * math.cos(math.pi / 2 * k / 40))
           for k in range(41)]
    fb = slide.shapes.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]), scale=1.0)
    fb.add_line_segments([(Inches(x), Inches(y)) for x, y in pts[1:]], close=False)
    curve = fb.convert_to_shape()
    curve.fill.background()
    curve.line.color.rgb = black
    curve.line.width = Pt(1.75)
    fl = slide.shapes.add_textbox(Inches(fx0 + (fx1 - fx0) * 0.66), Inches(fy0 - 0.12), Inches(1.2), Inches(0.5))
    tf = fl.text_frame
    tf.word_wrap = True
    add_run(tf.paragraphs[0], "Competitive", 11, bold=True)
    add_run(tf.add_paragraph(), "Frontier", 11, bold=True)
    pw, ph = fx1 - ax, ay_bot - fy0
    for f in m["firms"]:
        cx = ax + 0.05 + pw * (float(f["x"]) / 10.0)
        cy = ay_bot - ph * (float(f["y"]) / 10.0)
        if cx <= fx0:
            y_curve = fy0
        elif cx >= fx1:
            y_curve = fy1
        else:
            y_curve = fy1 - (fy1 - fy0) * math.cos(math.asin((cx - fx0) / (fx1 - fx0)))
        if f.get("frontier") or cy < y_curve:
            cy = y_curve
        focal = f.get("focal", False)
        d = 0.24 if focal else 0.17
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - d / 2), Inches(cy - d / 2), Inches(d), Inches(d))
        dot.fill.solid()
        dot.fill.fore_color.rgb = red if focal else dot_grey
        dot.line.fill.background()
        right_edge = cx > fx1 - 0.9
        lw = 1.8
        lx = cx - d / 2 - lw if right_edge else cx + d / 2
        lab = slide.shapes.add_textbox(Inches(lx), Inches(cy - 0.18), Inches(lw), Inches(0.35))
        lab.text_frame.word_wrap = False
        pp = lab.text_frame.paragraphs[0]
        if right_edge:
            pp.alignment = PP_ALIGN.RIGHT
        add_run(pp, f["name"], 12 if focal else 11, bold=focal, color=red if focal else black)


# ---------- slides ----------

def title_slide(prs, spec):
    for shape in prs.slides[0].shapes:
        if not shape.has_text_frame:
            continue
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                if r.text.strip().startswith("Company X"):
                    r.text = f"{spec['company']}:"
                if r.text.strip().startswith("Group:"):
                    if spec.get("author"):
                        r.text = f"By: {spec['author']}"
                    elif spec.get("group"):
                        r.text = f"Group: {spec['group']}"


def statements_slide(prs, spec):
    slide = new_content_slide(prs, "DCCD", f"{spec['company']}'s strategy in four statements")
    st = spec["statements"]
    items = []
    for key, label in (("define", "Define"), ("create", "Create"), ("capture", "Capture"), ("deliver", "Deliver")):
        items += [(0, label), (1, st[key])]
    bottom = bullets(slide, left_margin, 0.95, 12.6, items)
    entries = []
    if spec.get("competitive_advantage"):
        entries.append(("Competitive advantage", spec["competitive_advantage"]))
    if spec.get("one_line"):
        entries.append(("Key takeaway", spec["one_line"]))
    write_takeaways(slide, bottom, 12.6, entries, "statements")


def chain_slide(prs, spec):
    """Lecture slide 3: 'The DCCD chain: each stage feeds the next', filled for this firm."""
    ch = spec["chain"]
    slide = new_content_slide(prs, "The DCCD Chain", "each stage feeds the next")
    quads = [
        ("Define", "define", 0.31, 1.09, RGBColor(0xFF, 0xC0, 0x00), RGBColor(0xFF, 0xF2, 0xCC), black),
        ("Create", "create", 6.59, 1.09, RGBColor(0xFF, 0xFF, 0x00), RGBColor(0xEF, 0xFD, 0xA4), black),
        ("Capture", "capture", 0.31, 4.23, RGBColor(0xFF, 0x00, 0x00), RGBColor(0xFF, 0xE5, 0xEA), white),
        ("Deliver", "deliver", 6.59, 4.23, RGBColor(0x00, 0x70, 0xC0), RGBColor(0xDD, 0xEB, 0xF7), white),
    ]
    for label, key, x, y, hdr_fill, body_fill, hdr_text in quads:
        hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(6.12), Inches(0.36))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = hdr_fill
        hdr.line.color.rgb = black
        hdr.line.width = Pt(0.75)
        hdr.shadow.inherit = False
        tf = hdr.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_run(p, label, 16, bold=True, color=hdr_text)
        body = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y + 0.36), Inches(6.12), Inches(2.61))
        body.fill.solid()
        body.fill.fore_color.rgb = body_fill
        body.line.color.rgb = black
        body.line.width = Pt(0.75)
        body.shadow.inherit = False
        tf = body.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_top = tf.margin_bottom = Inches(0.08)
        for i, text in enumerate(ch[key]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", "285750")
            pPr.set("indent", "-285750")
            sb = etree.SubElement(pPr, qn("a:spcBef"))
            etree.SubElement(sb, qn("a:spcPts")).set("val", "600")
            bf = etree.SubElement(pPr, qn("a:buFont"))
            bf.set("typeface", "Arial")
            etree.SubElement(pPr, qn("a:buChar")).set("char", "•")
            add_rich(p, text, a_pt)
        need = sum(text_lines(t, 5.7, a_pt) for t in ch[key]) * a_pt * 1.2 / 72 + len(ch[key]) * 0.09 + 0.16
        if need > 2.61:
            warnings.append(f"WARNING [chain {label}]: text needs {need:.2f}in of 2.61in; shorten it")


def define_slide(prs, spec):
    d = spec["define"]
    slide = new_content_slide(prs, "Define", "who, what, and the problem solved")
    items = [(0, "Who is the customer?")] + [(1, t) for t in d["who"]]
    items += [(0, "What is the offer?")] + [(1, t) for t in d["what"]]
    items += [(0, "What problem does it solve?")] + [(1, t) for t in d["problem"]]
    bottom = bullets(slide, left_margin, 0.95, 12.6, items)
    write_takeaways(slide, bottom, 12.6, takeaways(d), "define")


def create_slides(prs, spec):
    c = spec["create"]
    slide = new_content_slide(prs, "Create", "why will customers pay more than it costs us?")
    items = [(0, "Willingness to pay vs. cost")] + [(1, t) for t in c["wtp_vs_cost"]]
    items += [(0, "Competitive positioning map")] + [(1, t) for t in c["map_reading"]]
    bottom = bullets(slide, left_margin, 0.95, 7.5, items)
    positioning_map(slide, c["map"], 7.9, 1.0, 5.25, 5.6)
    cap = slide.shapes.add_textbox(Inches(8.6), Inches(6.85), Inches(4.5), Inches(0.35))
    p = cap.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_run(p, "Illustrative positions, not measured", 10, italic=True, color=grey)
    write_takeaways(slide, bottom, 7.5, takeaways(c), "create map")

    slide = new_content_slide(prs, "Create", "can the resources behind the position hold it?")
    bottom = bullets(slide, left_margin, 0.95, 12.6, [(0, "VRI: Valuable, Rare, Costly to Imitate (Inimitable)?")])
    rows = [([r["resource"], r["type"], r["v"], r["r"], r["i"], r["verdict"]], bool(r.get("highlight")))
            for r in c["vri"]]
    entries = []
    if spec.get("competitive_advantage"):
        entries.append(("Competitive advantage", spec["competitive_advantage"]))
    entries = (entries + takeaways(c, "vri_takeaway"))[:2]
    bottom = table(slide, left_margin, bottom + 0.1, [2.6, 1.35, 2.25, 2.0, 2.4, 2.1],
                   ["Resource / capability", "Type", "V", "R", "I", "Verdict"], rows,
                   bottom_limit - reserve_for(entries, 12.6) - 0.4, "VRI")
    end = write_takeaways(slide, bottom, 12.6, entries, "VRI")
    note = slide.shapes.add_textbox(Inches(left_margin), Inches(end + 0.05), Inches(12.6), Inches(0.35))
    add_run(note.text_frame.paragraphs[0], "O (organized to exploit?) is saved for Deliver.", 14, color=grey)


def capture_slide(prs, spec):
    c = spec["capture"]
    slide = new_content_slide(prs, "Capture", "how much of the value stays with the firm?")
    rows = [([f["force"], f["level"], f["logic"], f["firm_power"]], bool(f.get("highlight"))) for f in c["forces"]]
    entries = takeaways(c)
    bottom = table(slide, left_margin, 1.0, [1.75, 1.0, 6.05, 3.9],
                   ["Force", "Power level", "What drives it, and what it does to margin",
                    "Does the firm have power over it, relative to competitors?"],
                   rows, bottom_limit - reserve_for(entries, 12.6), "capture")
    write_takeaways(slide, bottom, 12.6, entries, "capture")


def deliver_slides(prs, spec):
    """Deliver is the O: are the activities aligned with each other and with Define, Create, Capture?"""
    d = spec["deliver"]
    # slide 1: the Deliver statement and the activity system, linked to D, C, C
    slide = new_content_slide(prs, "Deliver", "which activities must the firm do well?")
    top_items = [(0, "Deliver statement"), (1, d["statement"])]
    bottom = bullets(slide, left_margin, 0.95, 12.6, top_items)
    rows = [([a["activity"], a["define_link"], a["create_link"], a["capture_link"]], bool(a.get("highlight")))
            for a in d["activities"]]
    entries = takeaways(d, "activities_takeaway")
    bottom = table(slide, left_margin, bottom + 0.12, [2.7, 3.2, 3.4, 3.4],
                   ["Key activity", "Delivers Define (who / what it serves)", "Reinforces Create (WTP or cost)",
                    "Reinforces Capture (force it defends against)"],
                   rows, bottom_limit - reserve_for(entries, 12.6), "deliver activities")
    write_takeaways(slide, bottom, 12.6, entries, "deliver activities")

    # slide 2: fit, trade-offs, misalignment, judgment
    slide = new_content_slide(prs, "Deliver", "do the activities fit and reinforce each other?")
    items = [(0, "Choices that reinforce each other")] + [(1, t) for t in as_list(d["reinforcement"])]
    items += [(0, "What the firm chooses not to do")] + [(1, t) for t in d["not_do"]]
    if d.get("misaligned"):
        items += [(0, "Activities that work against the position")] + [(1, t) for t in d["misaligned"]]
    if d.get("judgment"):
        items += [(0, "Judgment: organized to execute?")] + [(1, t) for t in as_list(d["judgment"])]
    bottom = bullets(slide, left_margin, 0.95, 12.6, items)
    write_takeaways(slide, bottom, 12.6, takeaways(d), "deliver fit")


def build(spec, out, template=default_template):
    prs = Presentation(template)
    drop_all_slides_but_first(prs)
    title_slide(prs, spec)
    statements_slide(prs, spec)
    if spec.get("chain"):
        chain_slide(prs, spec)
    define_slide(prs, spec)
    create_slides(prs, spec)
    capture_slide(prs, spec)
    deliver_slides(prs, spec)
    prs.save(out)
    for w in warnings:
        print(w)
    return out


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as fh:
        spec = json.load(fh)
    print(build(spec, sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else default_template))
