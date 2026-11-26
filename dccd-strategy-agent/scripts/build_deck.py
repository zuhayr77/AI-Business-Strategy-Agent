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
margin = 0.35            # top of the eyebrow line = bottom of the source line, from the slide edge
gap = 0.27               # one vertical and horizontal rhythm for every block
rule_y = 1.18
body_top = rule_y + gap                       # 1.45
source_h = 0.22
source_y = slide_h - margin - source_h        # 6.93
takeaway_h = 0.50
takeaway_y = source_y - gap - takeaway_h      # 6.16
body_bottom = takeaway_y - gap                # 5.89

questions = {
    "statements": ("DCCD", "The strategy in four statements"),
    "chain": ("DCCD CHAIN", "Each stage feeds the next"),
    "define": ("DEFINE", "Who is the customer, what is the offer, and what problem does it solve?"),
    "create": ("CREATE", "Why will customers pay more than it costs us?"),
    "vri": ("CREATE", "Are the resources behind the position valuable, rare and costly to imitate?"),
    "capture": ("CAPTURE", "How much of the value stays with the firm? Industry forces, then firm power"),
    "activities": ("DELIVER", "Which activities must the firm do well?"),
    "fit": ("DELIVER", "Do the activities fit and reinforce each other?"),
}
eyebrow_fill = {"DCCD": red, "DCCD CHAIN": red, "DEFINE": stage_fill["define"], "CREATE": stage_fill["create"],
                "CAPTURE": stage_fill["capture"], "DELIVER": stage_fill["deliver"]}
default_titles = {
    "statements": "{c}'s Strategy in Four Statements",
    "chain": "Each Stage of {c}'s Strategy Feeds the Next",
    "define": "Who, What, and the Problem Solved",
    "create": "Why Customers Pay More Than It Costs",
    "vri": "Can the Resources Behind the Position Hold It?",
    "capture": "How Much of the Value Stays With {c}",
    "activities": "Which Activities {c} Must Do Well",
    "fit": "Do {c}'s Activities Fit and Reinforce Each Other?",
}


def as_list(v):
    return v if isinstance(v, list) else [v]


def first(v):
    return as_list(v)[0]


class StrategyDeck(lc.Deck):
    def __init__(self):
        super().__init__()
        self.warnings = []

    def cover(self, company, subtitle, by):
        s = self._new("cover")
        lc.gradient_rect(s, 0, 0, slide_w, slide_h, deep, darkest)
        s.shapes.add_picture(os.path.join(brand_dir, "tartan-strip.png"), 0, 0, Inches(sidebar_w), Inches(slide_h))
        lc.text(s, 1.35, 1.55, 10.5, 1.0, company.upper(), 40, white, bold=True, anchor="b", line=1.0)
        lc.text(s, 1.35, 2.70, 10.5, 0.5, "STRATEGY ANALYSIS", 16, blush, bold=True, line=1.0)
        lc.text(s, 1.35, 3.25, 10.5, 0.5, subtitle, 20, rose, line=1.0)
        lc.text(s, 1.35, 4.10, 10.5, 0.4, by, 14, white, line=1.0)
        lc.text(s, 1.35, 4.50, 10.5, 0.4, datetime.date.today().strftime("%B %Y"), 12, blush, line=1.0)
        lc.hline(s, 1.35, 5.05, 4.2, rose, 2.0)
        lc.rect(s, sidebar_w, 6.15, slide_w - sidebar_w, slide_h - 6.15, fill=white)
        lc.picture_fit(s, os.path.join(brand_dir, "tepper-T.png"), 1.35, 6.38, 0.85, 0.85, align="l")
        self.meta[s.slide_id]["n_furniture"] = len(s.shapes)
        return s

    def page(self, key, title, src):
        s = self._new("page", title=title)
        s.shapes.add_picture(os.path.join(brand_dir, "tartan-strip.png"), 0, 0, Inches(sidebar_w), Inches(slide_h))
        lc.picture_fit(s, os.path.join(brand_dir, "tepper-T.png"), 12.48, margin, 0.60, 0.60, align="r")
        stage, question = questions[key]
        eb = lc.text(s, left, margin, width - 0.6, 0.24, f"**{stage}**   {question}", 10.5, gray, line=1.0)
        eb.text_frame.paragraphs[0].runs[0].font.color.rgb = eyebrow_fill[stage]
        size = 26 if lc._text_w(title, 26, True) <= width - 0.4 else 22
        lc.text(s, left, 0.62, width - 0.4, 0.46, title, size, red, bold=True, anchor="m", line=1.0, wrap=False)
        lc.hline(s, left, rule_y, right, lc.rule, 0.75)
        lc.text(s, left, source_y, width - 1.0, source_h, src, 9, gray, italic=True, anchor="b", line=1.0)
        lc.text(s, right - 0.6, source_y, 0.6, source_h, str(len(self.prs.slides)), 9, gray, align="r", anchor="b",
                line=1.0)
        if lc._text_w(title, 22, True) > width - 0.4:
            self.warnings.append(f"slide {len(self.prs.slides)}: title too long: '{title}'")
        self.meta[s.slide_id]["n_furniture"] = len(s.shapes)
        return s


# ---------- components ----------

def takeaway(s, content):
    """Key takeaway bar that closes every slide: tinted, red rail, bold dark red text."""
    lc.rect(s, left, takeaway_y, width, takeaway_h, fill=tint)
    lc.rect(s, left, takeaway_y, 0.07, takeaway_h, fill=red)
    lc.text(s, left + 0.30, takeaway_y + 0.04, width - 0.5, takeaway_h - 0.08,
            f"[[Key takeaway:]]  {content}", 12.5, skibo, bold=True, anchor="m", line=1.0)


def tile(s, x, y, w, h, label, content, size=12):
    """Stat-tile style panel: grey, red left rail, small red caps label, body text."""
    lc.rect(s, x, y, w, h, fill=panel)
    lc.rect(s, x, y, 0.07, h, fill=red)
    lc.text(s, x + 0.28, y + 0.10, w - 0.45, 0.24, label.upper(), 9.5, red, bold=True, line=1.0)
    lc.text(s, x + 0.28, y + 0.36, w - 0.45, h - 0.44, content, size, ink, line=1.1)


def card(s, x, y, w, h, title, color, items, size=12):
    """Grey card with a coloured top rule, stage-coloured title and square bullets."""
    lc.rect(s, x, y, w, h, fill=panel)
    lc.rect(s, x, y, w, 0.06, fill=color)
    lc.text(s, x + 0.22, y + 0.18, w - 0.44, 0.38, title, 17, color, bold=True, line=1.0)
    lc.hline(s, x + 0.22, y + 0.62, x + w - 0.22, lc.rule, 0.75, "dash")
    lc.bullets(s, x + 0.22, y + 0.76, w - 0.44, h - 0.90, items, size=size, gap=6)


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
    rows = [(k.capitalize(), None, [st[k]]) for k in ("define", "create", "capture", "deliver")]
    tile_h = 0.80
    lc.band_table(s, left, body_top, width, rows, label_w=1.75, h=body_bottom - tile_h - gap - body_top + 0.02,
                  size=12, label_size=15)
    tile(s, left, body_bottom - tile_h, width, tile_h, "Competitive advantage", spec["competitive_advantage"])
    takeaway(s, spec["one_line"])


def chain_slide(d, spec, src):
    s = d.page("chain", spec["titles"]["chain"], src)
    ch = spec["chain"]
    cw = (width - gap) / 2
    ch_h = (body_bottom - body_top - gap) / 2
    for i, key in enumerate(("define", "create", "capture", "deliver")):
        x = left + (i % 2) * (cw + gap)
        y = body_top + (i // 2) * (ch_h + gap)
        card(s, x, y, cw, ch_h, key.capitalize(), stage_fill[key], ch[key], size=11.5)
    takeaway(s, spec.get("chain_takeaway") or spec["one_line"])


def define_slide(d, spec, src):
    df = spec["define"]
    s = d.page("define", spec["titles"]["define"], src)
    cw = (width - 2 * gap) / 3
    cols = [("Who is the customer?", df["who"]), ("What is the offer?", df["what"]),
            ("What problem does it solve?", df["problem"])]
    for i, (head, items) in enumerate(cols):
        x = left + i * (cw + gap)
        lc.rect(s, x, body_top, cw, 0.06, fill=lc.band_fills[i])
        lc.header_strip(s, x, body_top + 0.06, cw, head, h=0.55, size=15)
        lc.dashed_items(s, x, body_top + 0.61 + gap, cw, items, size=12.5, h=body_bottom - body_top - 0.61 - gap)
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
    lab = lc.text(s, x0 - 0.92, y0 + 2.45, 2.30, 0.26, "Higher willingness to pay  >", 9.5, iron, bold=True,
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
    lc.text(s, fx + 0.12, fy - 0.50, 1.4, 0.24, "Competitive frontier", 9, rose, italic=True, line=1.0)
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
    lc.header_strip(s, left, body_top, lw, "Willingness to pay vs. cost", h=0.48, size=14)
    lc.bullets(s, left + 0.12, body_top + 0.62, lw - 0.15, 2.05, cr["wtp_vs_cost"], size=11.5, gap=6)
    lc.header_strip(s, left, body_top + 2.80, lw, "Reading the map", h=0.48, size=14)
    lc.bullets(s, left + 0.12, body_top + 3.42, lw - 0.15, body_bottom - body_top - 3.42, cr["map_reading"],
               size=11.5, gap=3)
    positioning_map(s, cr["map"], left + lw + gap, body_top, width - lw - gap, body_bottom - body_top)
    takeaway(s, first(cr["takeaway"]))


def vri_slide(d, spec, src):
    cr = spec["create"]
    s = d.page("vri", spec["titles"]["vri"], src)
    data = [["Resource / capability", "Type", "Valuable?", "Rare?", "Costly to imitate?", "Verdict"]]
    hl = []
    for i, r in enumerate(cr["vri"], start=1):
        data.append([r["resource"], r["type"], r["v"], r["r"], r["i"], r["verdict"]])
        if r.get("highlight"):
            hl.append(i)
    tile_h = 0.80
    fitted_table(s, body_top, data, [2.5, 1.35, 2.05, 1.95, 2.15, 1.35], body_bottom - tile_h - gap, hl, row_h=0.62)
    tile(s, left, body_bottom - tile_h, width, tile_h, "Competitive advantage", spec["competitive_advantage"])
    takeaway(s, first(cr["vri_takeaway"]))


def capture_slide(d, spec, src):
    cp = spec["capture"]
    s = d.page("capture", spec["titles"]["capture"], src)
    data = [["Force", "Power", "What drives it, and what it does to margin", "Does the firm push back, vs. rivals?"]]
    hl = []
    for i, f in enumerate(cp["forces"], start=1):
        data.append([f["force"], f["level"], f["logic"], f["firm_power"]])
        if f.get("highlight"):
            hl.append(i)
    gf, _ = fitted_table(s, body_top, data, [1.9, 0.85, 5.05, 3.55], body_bottom, hl, row_h=0.6)
    for ri in range(1, len(data)):
        c = gf.table.cell(ri, 1)
        for p in c.text_frame.paragraphs:
            for r in p.runs:
                r.font.bold = True
                r.font.color.rgb = red if data[ri][1].lower().startswith("high") else (
                    iron if data[ri][1].lower().startswith("low") else ink)
    takeaway(s, first(cp["takeaway"]))


def activities_slide(d, spec, src):
    dv = spec["deliver"]
    s = d.page("activities", spec["titles"]["activities"], src)
    tile(s, left, body_top, width, 0.95, "Deliver statement", dv["statement"], size=11.5)
    data = [["Key activity", "Delivers Define (who it serves)", "Reinforces Create (WTP or cost)",
             "Reinforces Capture (force it defends)"]]
    hl = []
    for i, a in enumerate(dv["activities"], start=1):
        data.append([a["activity"], a["define_link"], a["create_link"], a["capture_link"]])
        if a.get("highlight"):
            hl.append(i)
    fitted_table(s, body_top + 0.95 + gap, data, [2.75, 2.85, 2.85, 2.9], body_bottom, hl, row_h=0.56)
    takeaway(s, first(dv["activities_takeaway"]))


def fit_slide(d, spec, src):
    dv = spec["deliver"]
    s = d.page("fit", spec["titles"]["fit"], src)
    rows = [
        ("Reinforcing choices", "how activities feed each other", as_list(dv["reinforcement"])),
        ("Trade-offs", "what the firm won't do", dv["not_do"]),
        ("Misaligned", "what works against the position", as_list(dv.get("misaligned", ["None material"]))),
        ("Organized?", "judgment", as_list(dv.get("judgment", []))),
    ]
    lc.band_table(s, left, body_top, width, rows, label_w=2.1, h=body_bottom - body_top + 0.02, size=12, label_size=15)
    takeaway(s, first(dv["takeaway"]))


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
            if t + h > source_y - 0.05 and t < source_y:
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
    d.cover(c, "Define, Create, Capture and Deliver", f"By {spec['author']}" if spec.get("author") else "")
    statements_slide(d, spec, src)
    chain_slide(d, spec, src)
    define_slide(d, spec, src)
    create_slide(d, spec, src)
    vri_slide(d, spec, src)
    capture_slide(d, spec, src)
    activities_slide(d, spec, src)
    fit_slide(d, spec, src)
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
