"""Build a DCCD strategy deck from a JSON spec.

usage: python build_deck.py spec.json output.pptx

Design: consulting content layouts (band tables, grey cards, dashed item lists,
zebra data tables, stat tiles, Lato type) from lcp_deck.py, reskinned in the
Tepper red palette, with the Tepper tartan sidebar and "T" logo. No wordmark and
no course line. Everything is native and editable. The script prints WARNING
lines for any layout problem; shorten the spec text it names and rebuild.
"""
import datetime
import json
import math
import os
import sys

from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt, Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lcp_deck as lc

skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
brand_dir = os.path.join(skill_dir, "assets", "brand")


def hexc(h):
    return RGBColor.from_string(h)


# ---------- Tepper red palette, mapped onto the content library ----------
red = hexc("C41230")      # Carnegie Red: titles, rails, bullets, highlights
rose = hexc("EF3A47")     # Scots Rose
skibo = hexc("941120")    # Skibo Red: table headers, card titles
deep = hexc("5E0A18")     # darkest red
darkest = hexc("3D0710")
iron = hexc("6D6E71")     # Iron Gray
tint = hexc("FBECEE")     # highlighted rows and tiles
blush = hexc("E8B4BC")    # dashed separators
panel = hexc("F3F3F3")
ink = hexc("1C1C1C")
gray = hexc("7F7F7F")
white = hexc("FFFFFF")
stage_fill = {"define": skibo, "create": red, "capture": rose, "deliver": iron}

remap = {
    str(lc.blue): red, str(lc.sky): rose, str(lc.navy): skibo, str(lc.deep): deep,
    str(lc.mid): red, str(lc.soft): rose, str(lc.accent): red, str(lc.pale): tint,
    str(lc.dash_blue): blush, str(lc.ice): tint,
}
for name in ("blue", "sky", "navy", "deep", "mid", "soft", "accent", "pale", "dash_blue", "ice"):
    setattr(lc, name, remap[str(getattr(lc, name))])
lc.band_fills = [skibo, red, rose, iron, deep]
for fn in list(vars(lc).values()):
    if callable(fn) and getattr(fn, "__defaults__", None):
        fn.__defaults__ = tuple(remap.get(str(v), v) if isinstance(v, RGBColor) else v
                                for v in fn.__defaults__)

# ---------- grid ----------
slide_w, slide_h = 13.333, 7.5
sidebar_w = 0.70
left, right = 1.05, 12.40
width = right - left
margin = 0.35            # top of the action title, from the slide edge
gap = 0.27               # one vertical and horizontal rhythm for every block
title_y, title_h = margin, 0.55
rule_y = title_y + title_h + 0.10             # 1.00
body_top = rule_y + gap                       # 1.27
mark_x, mark_y, mark_h = left, 6.675, 0.55    # CMU Tepper wordmark: same size and place as on the cover
mark_w = mark_h * 760 / 139                   # 3.01, keeps the png's aspect ratio
logo_x, logo_y, logo_s = 12.35, 0.15, 0.90    # Tepper T, top right
num_right = logo_x + logo_s                   # page number right edge lines up with the T
foot_y, foot_h = mark_y, mark_h               # footer row: wordmark, source, page number
source_x = mark_x + mark_w + 0.30
bar_h = 0.78                                  # key takeaway and competitive advantage bars
takeaway_y = foot_y - gap - bar_h             # 5.63
body_bottom = takeaway_y - gap                # 5.36

stage_label = {"define": "Who it serves", "create": "How it wins", "capture": "Who keeps the value",
               "deliver": "How it delivers"}

default_titles = {
    "statements": "{c}'s Strategy in Four Parts",
    "chain": "How Each Part of {c}'s Strategy Builds on the Next",
    "define": "Who {c} Serves and Why They Choose It",
    "create": "Why Customers Choose {c} Over the Alternatives",
    "edge": "What Rivals Can and Cannot Copy",
}


def as_list(v):
    return v if isinstance(v, list) else [v]


def first(v):
    return as_list(v)[0]


class StrategyDeck(lc.Deck):
    def __init__(self):
        super().__init__()
        self.warnings = []

    def cover(self, company, by):
        s = self._new("cover")
        lc.gradient_rect(s, 0, 0, slide_w, slide_h, deep, darkest)
        s.shapes.add_picture(os.path.join(brand_dir, "tartan-strip.png"), 0, 0, Inches(sidebar_w), Inches(slide_h))
        name = company.upper()
        name_size = 72 if lc._text_w(name, 72, True) <= 11.6 else 54
        lc.text(s, 1.25, 1.99, 11.8, 1.15, name, name_size, white, bold=True, line=1.0)
        lc.text(s, 1.35, 3.40, 11.0, 0.5, "STRATEGY ANALYSIS", 26, blush, bold=True, line=1.0)
        lc.hline(s, 1.35, 4.25, 6.0, rose, 2.5)
        lc.text(s, 1.35, 4.75, 11.0, 0.4, by, 18, white, line=1.0)
        lc.text(s, 1.35, 5.30, 11.0, 0.35, datetime.date.today().strftime("%B %Y"), 15, blush, line=1.0)
        band_y = 6.40
        lc.rect(s, sidebar_w, band_y, slide_w - sidebar_w, slide_h - band_y, fill=white)
        lc.picture_fit(s, os.path.join(brand_dir, "cmu-wordmark.png"), mark_x, mark_y, mark_w, mark_h, align="l")
        self.meta[s.slide_id]["n_furniture"] = len(s.shapes)
        return s

    def page(self, key, title, src):
        s = self._new("page", title=title)
        s.shapes.add_picture(os.path.join(brand_dir, "tartan-strip.png"), 0, 0, Inches(sidebar_w), Inches(slide_h))
        lc.picture_fit(s, os.path.join(brand_dir, "tepper-T.png"), logo_x, logo_y, logo_s, logo_s, align="r")
        size = 26 if lc._text_w(title, 26, True) <= width - 0.8 else 22
        lc.text(s, left, title_y, width - 0.8, title_h, title, size, red, bold=True, anchor="m", line=1.0, wrap=False)
        lc.hline(s, left, rule_y, right, lc.rule, 0.75)
        lc.picture_fit(s, os.path.join(brand_dir, "cmu-wordmark.png"), mark_x, mark_y, mark_w, mark_h, align="l")
        lc.text(s, source_x, foot_y, right - source_x, foot_h, src, 9, gray, italic=True, anchor="m", line=1.0)
        lc.text(s, num_right - 0.6, foot_y, 0.6, foot_h, str(len(self.prs.slides)), 10, gray, align="r", anchor="m",
                line=1.0)
        if lc._text_w(title, 22, True) > width - 0.8:
            self.warnings.append(f"slide {len(self.prs.slides)}: title too long: '{title}'")
        self.meta[s.slide_id]["n_furniture"] = len(s.shapes)
        return s


# ---------- components ----------

def bar(s, y, label, content, x=left, w=width, h=bar_h):
    """Tinted bar with a red rail: small red caps label, then the bold message on a new line.
    Used for both the key takeaway and the competitive advantage."""
    lc.rect(s, x, y, w, h, fill=tint)
    lc.rect(s, x, y, 0.07, h, fill=red)
    tb, tf = lc.tbox(s, x + 0.30, y + 0.06, w - 0.50, h - 0.12, anchor="m")
    lc.para(tf, label.upper(), 9.5, red, bold=True, first=True, after=3, line=1.0)
    lc.para(tf, content, 12.5, skibo, bold=True, after=0, line=1.05)


def takeaway(s, content):
    bar(s, takeaway_y, "Key takeaway", content)


def card(s, x, y, w, h, title, color, items, size=12):
    """Grey card with a coloured top rule, title centred between the rule and the dashed line, square bullets."""
    lc.rect(s, x, y, w, h, fill=panel)
    lc.rect(s, x, y, w, 0.06, fill=color)
    lc.text(s, x + 0.22, y + 0.06, w - 0.44, 0.48, title, 16, color, bold=True, anchor="m", line=1.0)
    lc.hline(s, x + 0.22, y + 0.54, x + w - 0.22, lc.rule, 0.75, "dash")
    lc.bullets(s, x + 0.22, y + 0.68, w - 0.44, h - 0.80, items, size=size, gap=5)


def table_height(gf):
    tbl = gf.table
    total = 0.0
    for row in tbl.rows:
        need = Emu(row.height).inches
        for ci, c in enumerate(row.cells):
            cw = Emu(tbl.columns[ci].width).inches - 0.16
            nh, _ = lc._tf_need(c.text_frame, cw)
            need = max(need, nh + 0.12)
        total += need
    return total


def fitted_table(s, y, data, col_w, max_bottom, highlight_rows=(), row_h=0.5, sizes=(11.5, 11, 10.5, 10, 9.5, 9)):
    """Largest font whose zebra table fits above max_bottom. Highlighted rows get the
    red tint and a bold red first column. Returns (graphic frame, bottom)."""
    gf = None
    for fs in sizes:
        if gf is not None:
            gf._element.getparent().remove(gf._element)
        gf = lc.data_table(s, left, y, width, data, col_w, row_h=row_h, size=fs, head_size=fs + 0.5, dash_lists=False)
        gf.table.rows[0].height = Inches(0.42)
        bottom = y + table_height(gf)
        if bottom <= max_bottom:
            break
    tbl = gf.table
    # stretch rows evenly so the table ends exactly at max_bottom (even spacing on every slide)
    needs = []
    for row in tbl.rows:
        need = Emu(row.height).inches
        for ci, c in enumerate(row.cells):
            nh, _ = lc._tf_need(c.text_frame, Emu(tbl.columns[ci].width).inches - 0.16)
            need = max(need, nh + 0.12)
        needs.append(need)
    extra = max(0.0, (max_bottom - y) - sum(needs)) / (len(needs) - 1)
    for ri, row in enumerate(tbl.rows):
        row.height = Inches(needs[ri] + (extra if ri else 0))
    bottom = y + sum(needs) + extra * (len(needs) - 1)
    for ri in highlight_rows:
        for ci in range(len(col_w)):
            c = tbl.cell(ri, ci)
            c.fill.solid()
            c.fill.fore_color.rgb = tint
            for p in c.text_frame.paragraphs:
                for r in p.runs:
                    r.font.bold = True
                    if ci == 0:
                        r.font.color.rgb = red
    return gf, bottom


# ---------- slides ----------

def statements_slide(d, spec, src):
    s = d.page("statements", spec["titles"]["statements"], src)
    st = spec["statements"]
    keys = ("define", "create", "capture", "deliver")
    n = len(keys)
    row_h = (body_bottom - bar_h - gap - body_top) / n
    label_w = 2.05
    for i, k in enumerate(keys):
        yy = body_top + i * row_h
        box = lc.rect(s, left, yy, label_w, row_h - 0.04, fill=lc.band_fills[i])
        lc.shape_text(box, stage_label[k], 15, white, align="c", anchor="m")
        if i % 2 == 1:
            lc.rect(s, left + label_w + 0.02, yy, width - label_w - 0.02, row_h - 0.04, fill=hexc("F9F9F9"))
        lc.text(s, left + label_w + 0.25, yy + 0.04, width - label_w - 0.40, row_h - 0.12, st[k], 12, ink,
                anchor="m", line=1.15)
        lc.hline(s, left + label_w, yy + row_h - 0.02, right, blush, 1.0, "dash")
    bar(s, body_bottom - bar_h, "Competitive advantage", spec["competitive_advantage"])
    takeaway(s, spec["one_line"])


def chain_slide(d, spec, src):
    s = d.page("chain", spec["titles"]["chain"], src)
    ch = spec["chain"]
    cw = (width - gap) / 2
    ch_h = (body_bottom - body_top - gap) / 2
    for i, key in enumerate(("define", "create", "capture", "deliver")):
        x = left + (i % 2) * (cw + gap)
        y = body_top + (i // 2) * (ch_h + gap)
        card(s, x, y, cw, ch_h, stage_label[key], stage_fill[key], ch[key], size=11.5)
    takeaway(s, spec.get("chain_takeaway") or spec["one_line"])


def define_slide(d, spec, src):
    df = spec["define"]
    s = d.page("define", spec["titles"]["define"], src)
    cw = (width - 2 * gap) / 3
    cols = [("Who is the customer?", df["who"]), ("What do they get?", df["what"]),
            ("What problem does it solve?", df["problem"])]
    y = body_top
    stats = df.get("stats", [])[:3]
    if stats:
        sh = 1.05
        for i, st in enumerate(stats):
            x = left + i * (cw + gap)
            lc.rect(s, x, y, cw, sh, fill=tint)
            lc.text(s, x, y + 0.10, cw, 0.55, st["value"], 30, red, bold=True, align="c", anchor="m", line=1.0)
            lc.text(s, x + 0.15, y + 0.66, cw - 0.30, 0.30, st["label"], 11, ink, align="c", anchor="t", line=1.0)
        y += sh + gap
    head_h = 0.50
    for i, (head, items) in enumerate(cols):
        x = left + i * (cw + gap)
        hb = lc.rect(s, x, y, cw, head_h, fill=lc.band_fills[i])
        lc.shape_text(hb, head, 14, white, bold=True, align="c", anchor="m")
        lc.rect(s, x, y + head_h, cw, body_bottom - y - head_h, fill=panel)
        items = as_list(items)
        ih = (body_bottom - y - head_h) / len(items)
        for j, it in enumerate(items):
            iy = y + head_h + j * ih
            lc.text(s, x + 0.22, iy + 0.06, cw - 0.44, ih - 0.12, it, 12, ink, anchor="m", line=1.15)
            if j < len(items) - 1:
                lc.hline(s, x + 0.22, iy + ih, x + cw - 0.22, lc.rule, 0.75, "dash")
    takeaway(s, first(df["takeaway"]))


def frontier_y(xv):
    xv = min(max(xv, 0.0), 10.0)
    return 9.6 * math.sqrt(max(0.0, 1 - (xv / 10.4) ** 2))


def positioning_map(s, m, x0, y0, w, h):
    lc.text(s, x0, y0, w, 0.30, "Competitive positioning map", 14, ink, bold=True, line=1.0)
    lc.text(s, x0, y0 + 0.30, w, 0.26, f"{m.get('y_label', '')} vs. cost {m.get('x_label', '')}", 10.5, gray,
            line=1.0)
    lc.rect(s, x0, y0 + 0.66, w, h - 0.66, fill=panel)
    ax, ay_top, ay_bot, ax_r = x0 + 0.55, y0 + 1.05, y0 + h - 0.45, x0 + w - 0.20
    lc.line(s, ax, ay_top, ax, ay_bot, iron, 1.25)
    lc.line(s, ax, ay_bot, ax_r, ay_bot, iron, 1.25)
    lab = lc.text(s, x0 - 0.92, y0 + 2.45, 2.30, 0.26, "Customers value it more  >", 9.5, iron, bold=True,
                  align="c", line=1.0)
    lab.rotation = 270
    lc.text(s, ax, ay_bot + 0.07, ax_r - ax, 0.24, "Lower cost  >", 9.5, iron, bold=True, align="r", line=1.0)
    pw, ph = ax_r - ax - 0.30, ay_bot - ay_top - 0.10

    def pos(xv, yv):
        return ax + 0.12 + pw * xv / 10.0, ay_bot - ph * yv / 10.0

    pts = [pos(k / 4.0, frontier_y(k / 4.0)) for k in range(0, 41)]
    fb = s.shapes.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]), scale=1.0)
    fb.add_line_segments([(Inches(px), Inches(py)) for px, py in pts[1:]], close=False)
    curve = fb.convert_to_shape()
    curve.fill.background()
    curve.line.color.rgb = rose
    curve.line.width = Pt(1.75)
    lc._dash(curve.line, "dash")
    st = curve._element.find(lc.qn("p:style"))
    if st is not None:
        curve._element.remove(st)
    fx, fy = pos(6.6, frontier_y(6.6))
    lc.text(s, fx + 0.12, fy - 0.50, 1.4, 0.24, "Best value on offer", 9, rose, italic=True, line=1.0)
    for f in m["firms"]:
        xv, yv = float(f["x"]), float(f["y"])
        if f.get("frontier") or yv > frontier_y(xv):
            yv = frontier_y(xv)
        cx, cy = pos(xv, yv)
        focal = f.get("focal", False)
        dd = 0.26 if focal else 0.17
        lc.rect(s, cx - dd / 2, cy - dd / 2, dd, dd, fill=red if focal else skibo,
                line=white, lw=1.0, shape=lc.MSO_SHAPE.OVAL)
        right_edge = cx > ax_r - 1.55
        lw_ = 1.50
        lx = cx - dd / 2 - lw_ - 0.05 if right_edge else cx + dd / 2 + 0.05
        lc.text(s, lx, cy - 0.13, lw_, 0.26, f["name"], 10.5 if focal else 9.5, red if focal else ink,
                bold=focal, align="r" if right_edge else "l", line=1.0)


def create_slide(d, spec, src):
    cr = spec["create"]
    s = d.page("create", spec["titles"]["create"], src + "; map positions are illustrative, not measured")
    lw = 5.55
    lc.header_strip(s, left, body_top, lw, "Why customers choose it, and what it costs", h=0.48, size=14)
    lc.bullets(s, left + 0.12, body_top + 0.62, lw - 0.15, 1.80, cr["wtp_vs_cost"], size=11.5, gap=5)
    lc.header_strip(s, left, body_top + 2.55, lw, "Reading the map", h=0.48, size=14)
    lc.bullets(s, left + 0.12, body_top + 3.17, lw - 0.15, body_bottom - body_top - 3.17, cr["map_reading"],
               size=11.5, gap=3)
    positioning_map(s, cr["map"], left + lw + gap, body_top, width - lw - gap, body_bottom - body_top)
    takeaway(s, first(cr["takeaway"]))


def edge_slide(d, spec, src):
    """One table: what the firm has or does, why it matters, whether rivals can copy it,
    the threat it guards against, and how long the edge lasts."""
    s = d.page("edge", spec["titles"]["edge"], src)
    rows = spec.get("advantage_table")
    if rows:
        data = [["What the company has or does", "Why it matters", "Can rivals copy it?",
                 "Threat it guards against", "How long it lasts"]]
        keys = ("strength", "why", "copy", "threat", "lasts")
        col_w = [2.45, 2.45, 2.25, 2.25, 1.95]
    else:
        rows = [{"strength": r["resource"], "why": r["v"], "copy": r["i"], "lasts": r["verdict"],
                 "highlight": r.get("highlight")} for r in spec["create"]["vri"]]
        data = [["What the company has or does", "Why it matters", "Can rivals copy it?", "How long it lasts"]]
        keys = ("strength", "why", "copy", "lasts")
        col_w = [3.0, 3.0, 3.0, 2.35]
    hl = []
    for i, r in enumerate(rows, start=1):
        data.append([r[k] for k in keys])
        if r.get("highlight"):
            hl.append(i)
    fitted_table(s, body_top, data, col_w, body_bottom, hl, row_h=0.55)
    edge = spec.get("edge_takeaway") or first(spec["create"].get("vri_takeaway", spec["one_line"]))
    takeaway(s, edge)


def check(d):
    """Body content must stay inside the content box; text must fit; tables must not grow past the takeaway."""
    problems = list(d.warnings)
    for idx, s in enumerate(d.prs.slides, start=1):
        meta = d.meta.get(s.slide_id, {})
        if meta.get("kind") != "page":
            continue
        for sh in list(s.shapes)[meta["n_furniture"]:]:
            if sh.left is None or sh.rotation not in (0, 0.0):
                continue
            l, t, w, h = lc._box(sh)
            if l + w > right + 0.03 or l < left - 0.03:
                problems.append(f"slide {idx}: '{sh.name}' leaves the content column")
            if t + h > foot_y - 0.05 and t < foot_y:
                problems.append(f"slide {idx}: '{sh.name}' ends at {t + h:.2f}, into the source line")
            if sh.has_table:
                bottom = t + table_height(sh)
                if bottom > body_bottom + 0.02:
                    problems.append(f"slide {idx}: table grows to {bottom:.2f}in and hits the takeaway bar")
            elif sh.has_text_frame and sh.text_frame.text.strip():
                tf = sh.text_frame
                ml = Emu(tf.margin_left or 0).inches + Emu(tf.margin_right or 0).inches
                mt = Emu(tf.margin_top or 0).inches + Emu(tf.margin_bottom or 0).inches
                need, _ = lc._tf_need(tf, w - ml, tf.word_wrap is not False)
                if tf.word_wrap is not False and need > (h - mt) * 1.12 + 0.05:
                    problems.append(f"slide {idx}: text likely overflows: '{tf.text[:50]}'")
            if sh.has_text_frame and ("—" in sh.text_frame.text or "–" in sh.text_frame.text):
                problems.append(f"slide {idx}: dash character in '{sh.text_frame.text[:40]}'")
    return problems


def build(spec_path, out_path):
    spec = json.load(open(spec_path))
    c = spec["company"]
    titles = {k: v.format(c=c) for k, v in default_titles.items()}
    titles.update(spec.get("titles", {}))
    spec["titles"] = titles
    src = "Source: " + spec.get("source", "company filings and reputable press; figures marked est. are estimates")
    d = StrategyDeck()
    d.cover(c, f"By {spec['author']}" if spec.get("author") else "")
    statements_slide(d, spec, src)
    chain_slide(d, spec, src)
    define_slide(d, spec, src)
    create_slide(d, spec, src)
    edge_slide(d, spec, src)
    problems = check(d)
    for p in problems:
        print("WARNING:", p)
    d.save(out_path)
    print(out_path)
    return problems


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit("usage: python build_deck.py spec.json output.pptx")
    build(sys.argv[1], sys.argv[2])
