"""Build a DCCD strategy deck in the Carnegie Mellon (Tepper) design system from a JSON spec.

usage: python build_deck.py spec.json output.pptx

Design comes from tepper_deck.py (Open Sans, Carnegie Red palette, tartan sidebar,
red action titles that state the finding, summary bars, native shapes and tables).
Differences from the stock library: no CMU wordmark on any slide and no course line.
The script runs lint() and prints every problem; shorten the spec text it names and rebuild.
"""
import json
import math
import os
import sys

from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import tepper_deck as td
from tepper_deck import (Deck, red, rose, skibo, iron, light, tint, white, black, title_panel, title_gray,
                         rect, oval, text, tbox, para, bullets, hline, line, picture_fit, slide_number,
                         summary_bar, callout, table, option_cards, section_label, lint, _tf_need)

skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
brand_dir = os.path.join(skill_dir, "assets", "brand")
stage_fill = {"define": red, "create": rose, "capture": skibo, "deliver": iron}

questions = {
    "statements": "DCCD: the strategy in four statements",
    "chain": "DCCD chain: each stage feeds the next",
    "define": "Define: who is the customer, what is the offer, and what problem does it solve?",
    "create": "Create: why will customers pay more than it costs us?",
    "vri": "Create: are the resources behind the position valuable, rare and costly to imitate?",
    "capture": "Capture: how much of the value stays with the firm? Industry forces, then firm power",
    "activities": "Deliver: which activities must the firm do well? Each linked to Define, Create, Capture",
    "fit": "Deliver: do the activities fit and reinforce each other?",
}
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


class StrategyDeck(Deck):
    """Tepper design without the CMU wordmark or a course line."""

    def title_slide(self, title, org, names):
        s = self._new("title")
        bgp = self._a("title_background")
        if bgp:
            s.shapes.add_picture(bgp, 0, 0, Inches(td.slide_w), Inches(td.slide_h))
        else:
            rect(s, 0, 0, td.slide_w, td.slide_h, fill=td.title_navy)
        rect(s, 5.44, 0, td.slide_w - 5.44, td.slide_h, fill=title_panel)
        text(s, 6.32, 2.50, 6.20, 1.31, title, 32 if len(title) <= 34 else 26, iron, anchor="m")
        hline(s, 6.32, 4.10, 12.54, color=title_gray, lw=1.0)
        text(s, 6.38, 4.40, 6.10, 0.37, org, 15, iron, bold=True, anchor="m")
        text(s, 6.38, 4.87, 6.40, 0.71, names, 13, iron)
        self._mark_furniture(s)
        return s

    def content(self, title, source=None, sub=None, sidebar=None):
        kind = sidebar or self.sidebar
        s = self._new("content", sidebar=kind, has_sub=bool(sub), has_source=bool(source))
        left = self._sidebar(s, kind)
        self.meta[s.slide_id]["left"] = left
        tp = self._a("t_logo")
        if tp:
            picture_fit(s, tp, *td.logo_box)
        n = len(title)
        size = 24 if n <= 56 else (21 if n <= 66 else 19)
        _, htf = tbox(s, left, td.title_y, 13.20 - left, 0.65, wrap=False, inset=0.1)
        para(htf, title, size, red, bold=True, first=True, after=0)
        hline(s, left - 0.05, td.rule_y, 13.10, color=td.rule_gray, lw=0.75)
        if sub:
            text(s, left, td.sub_y, td.content_r - left, 0.30, sub, 12, iron, italic=True)
        if source:
            _, stf = tbox(s, left, td.source_y, td.content_r - left, 0.37, wrap=False, inset=0.1)
            para(stf, source, 9, iron, italic=True, first=True, after=0)
        slide_number(s, len(self.prs.slides))
        self._mark_furniture(s)
        return s


# ---------- helpers ----------

def as_list(v):
    return v if isinstance(v, list) else [v]


def first(v):
    return as_list(v)[0]


def table_height(gt):
    """Estimated rendered height of a native table (rows grow to fit text)."""
    total = 0.0
    for row in gt.rows:
        need = row.height / 914400
        for ci, c in enumerate(row.cells):
            cw = gt.columns[ci].width / 914400 - 0.16
            nh, _ = _tf_need(c.text_frame, cw)
            need = max(need, nh + 0.10)
        total += need
    return total


def fitted_table(s, x, y, w, data, col_w, max_bottom, highlight_rows=(), row_min=0.3, sizes=(11, 10.5, 10, 9.5, 9)):
    """Largest font (10.5 down to 9) whose table fits above max_bottom. Highlighted rows get
    the red tint, bold text and a red first column. Returns (table, bottom)."""
    frame = None
    for fs in sizes:
        if frame is not None:
            frame._element.getparent().remove(frame._element)
        gt = table(s, x, y, w, 0.4 * len(data), data, col_w, row_h=[0.40] + [row_min] * (len(data) - 1),
                   fsize=fs, align=["l"] * len(col_w))
        frame = s.shapes[-1]
        for ri in range(1, len(data)):
            for r in gt.cell(ri, 0).text_frame.paragraphs[0].runs:
                r.font.bold = True
        for ri in highlight_rows:
            for ci in range(len(col_w)):
                c = gt.cell(ri, ci)
                c.fill.solid()
                c.fill.fore_color.rgb = tint
                for r in c.text_frame.paragraphs[0].runs:
                    r.font.bold = True
                    if ci == 0:
                        r.font.color.rgb = red
        bottom = y + table_height(gt)
        if bottom <= max_bottom:
            break
    return gt, bottom


def rail_panel(s, x, y, w, h, content, size=11):
    """Tinted panel with a red left rail: used for the competitive advantage and statements."""
    rect(s, x, y, w, h, fill=tint)
    rect(s, x, y, 0.055, h, fill=red)
    text(s, x + 0.25, y + 0.06, w - 0.40, h - 0.12, content, size, black, anchor="m")


def stage_card(s, x, y, w, header, fill, items, body_h, size=11.5):
    rect(s, x, y, w, 0.48, fill=fill)
    text(s, x + 0.20, y + 0.05, w - 0.35, 0.38, header, 13, white, bold=True, anchor="m")
    rect(s, x, y + 0.48, w, body_h, fill=light)
    bullets(s, x + 0.22, y + 0.62, w - 0.44, body_h - 0.24, items, size=size, gap=6)


def bar(s, x, y, w, content, i):
    summary_bar(s, x, y, w, content, fill=red if i % 2 == 0 else iron)


# ---------- slides ----------

def statements_slide(d, spec, src):
    s = d.content(spec["titles"]["statements"], source=src, sub=questions["statements"])
    x, w = d.left(s), d.width(s)
    st = spec["statements"]
    y = d.top(s)
    for key in ("define", "create", "capture", "deliver"):
        chip = rect(s, x, y, 1.45, 0.76, fill=stage_fill[key])
        td.shape_text(chip, key.capitalize(), 13, white, bold=True)
        rect(s, x + 1.45, y, w - 1.45, 0.76, fill=light)
        text(s, x + 1.65, y + 0.04, w - 1.85, 0.68, st[key], 11, black, anchor="m")
        y += 0.84
    rail_panel(s, x, y + 0.04, w, 0.58, f"**Competitive advantage:** {spec['competitive_advantage']}", 11)
    bar(s, x, 5.85, w, spec["one_line"], 0)


def chain_slide(d, spec, src):
    s = d.content(spec["titles"]["chain"], source=src, sub=questions["chain"])
    x, w = d.left(s), d.width(s)
    ch = spec["chain"]
    cw = (w - 0.25) / 2
    y0 = d.top(s)
    for i, key in enumerate(("define", "create", "capture", "deliver")):
        cx = x + (i % 2) * (cw + 0.25)
        cy = y0 + (i // 2) * 2.12
        stage_card(s, cx, cy, cw, key.capitalize(), stage_fill[key], ch[key], 1.47)
    bar(s, x, 5.85, w, spec.get("chain_takeaway") or spec["one_line"], 1)


def define_slide(d, spec, src):
    df = spec["define"]
    s = d.content(spec["titles"]["define"], source=src, sub=questions["define"])
    x, w = d.left(s), d.width(s)
    option_cards(s, x, d.top(s), w, [
        ("people", "Who is the customer?", red, df["who"]),
        ("target", "What is the offer?", rose, df["what"]),
        ("gauge", "What problem does it solve?", skibo, df["problem"]),
    ], body_h=3.35, size=12.5)
    bar(s, x, 5.85, w, first(df["takeaway"]), 0)


def frontier_y(xv):
    """Concave competitive frontier on a 0 to 10 grid (x = 10 is lowest cost)."""
    xv = min(max(xv, 0.0), 10.0)
    return 9.6 * math.sqrt(max(0.0, 1 - (xv / 10.4) ** 2))


def positioning_map(s, m, x0, y0, w, h):
    ax, ay_top, ay_bot, ax_r = x0 + 0.30, y0 + 0.45, y0 + h - 0.50, x0 + w - 0.10
    line(s, ax, ay_top, ax, ay_bot, color=iron, lw=1.5)
    line(s, ax, ay_bot, ax_r, ay_bot, color=iron, lw=1.5)
    text(s, x0, y0, w, 0.25, "**Higher willingness to pay**", 10.5, black)
    text(s, x0, y0 + 0.22, w, 0.22, m.get("y_label", ""), 9.5, iron, italic=True)
    text(s, ax, ay_bot + 0.06, ax_r - ax, 0.22, "**Lower cost** " + m.get("x_label", ""), 10, black, align="r")
    pw, ph = ax_r - ax - 0.25, ay_bot - ay_top

    def pos(xv, yv):
        return ax + 0.10 + pw * xv / 10.0, ay_bot - ph * yv / 10.0

    pts = [pos(k / 4.0, frontier_y(k / 4.0)) for k in range(0, 41)]
    fb = s.shapes.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]), scale=1.0)
    fb.add_line_segments([(Inches(px), Inches(py)) for px, py in pts[1:]], close=False)
    curve = fb.convert_to_shape()
    curve.fill.background()
    curve.line.color.rgb = red
    curve.line.width = Pt(1.75)
    curve.shadow.inherit = False
    fx, fy = pos(6.2, frontier_y(6.2))
    text(s, fx + 0.15, fy - 0.55, 1.3, 0.40, "Competitive frontier", 9, red, italic=True)
    for f in m["firms"]:
        xv, yv = float(f["x"]), float(f["y"])
        if f.get("frontier") or yv > frontier_y(xv):
            yv = frontier_y(xv)
        cx, cy = pos(xv, yv)
        focal = f.get("focal", False)
        dd = 0.24 if focal else 0.17
        oval(s, cx - dd / 2, cy - dd / 2, dd, dd, fill=red if focal else iron, lw=0)
        right_edge = cx > ax_r - 1.6
        lw_ = 1.55
        lx = cx - dd / 2 - lw_ - 0.04 if right_edge else cx + dd / 2 + 0.04
        text(s, lx, cy - 0.14, lw_, 0.28, f["name"], 10 if focal else 9.5, red if focal else black,
             bold=focal, align="r" if right_edge else "l")


def create_slide(d, spec, src):
    cr = spec["create"]
    s = d.content(spec["titles"]["create"], source=src + "; map positions are illustrative, not measured",
                  sub=questions["create"])
    x, w = d.left(s), d.width(s)
    lw = 5.65
    y = d.top(s)
    section_label(s, x, y, lw, "Willingness to pay vs. cost", color=red)
    bullets(s, x, y + 0.36, lw, 2.30, cr["wtp_vs_cost"], size=11.5, gap=6)
    section_label(s, x, y + 2.70, lw, "Competitive positioning map", color=red)
    bullets(s, x, y + 3.06, lw, 1.05, cr["map_reading"], size=11.5, gap=4)
    positioning_map(s, cr["map"], x + lw + 0.35, y, w - lw - 0.35, 4.10)
    bar(s, x, 5.85, w, first(cr["takeaway"]), 1)


def vri_slide(d, spec, src):
    cr = spec["create"]
    s = d.content(spec["titles"]["vri"], source=src, sub=questions["vri"])
    x, w = d.left(s), d.width(s)
    data = [["Resource / capability", "Type", "Valuable?", "Rare?", "Costly to imitate?", "Verdict"]]
    hl = []
    for i, r in enumerate(cr["vri"], start=1):
        data.append([r["resource"], r["type"], r["v"], r["r"], r["i"], r["verdict"]])
        if r.get("highlight"):
            hl.append(i)
    _, bottom = fitted_table(s, x, d.top(s), w, data, [2.45, 1.30, 2.05, 1.95, 2.10, 1.40], 4.95, hl, row_min=0.62)
    y = max(bottom + 0.20, 4.60)
    rail_panel(s, x, y, w, 0.62, f"**Competitive advantage:** {spec['competitive_advantage']}", 11)
    bar(s, x, 5.85, w, first(cr["vri_takeaway"]), 0)


def capture_slide(d, spec, src):
    cp = spec["capture"]
    s = d.content(spec["titles"]["capture"], source=src, sub=questions["capture"])
    x, w = d.left(s), d.width(s)
    data = [["Force", "Power", "What drives it, and what it does to margin", "Does the firm push back, vs. rivals?"]]
    hl = []
    for i, f in enumerate(cp["forces"], start=1):
        data.append([f["force"], f["level"], f["logic"], f["firm_power"]])
        if f.get("highlight"):
            hl.append(i)
    gt, _ = fitted_table(s, x, d.top(s), w, data, [1.85, 0.85, 5.05, 3.50], 5.70, hl, row_min=0.55)
    for ri in range(1, len(data)):
        for r in gt.cell(ri, 1).text_frame.paragraphs[0].runs:
            r.font.bold = True
            r.font.color.rgb = red if data[ri][1].lower().startswith("high") else black
    bar(s, x, 5.85, w, first(cp["takeaway"]), 1)


def activities_slide(d, spec, src):
    dv = spec["deliver"]
    s = d.content(spec["titles"]["activities"], source=src, sub=questions["activities"])
    x, w = d.left(s), d.width(s)
    y = d.top(s)
    rail_panel(s, x, y, w, 0.82, f"**Deliver statement:** {dv['statement']}", 10.5)
    data = [["Key activity", "Delivers Define (who it serves)", "Reinforces Create (WTP or cost)",
             "Reinforces Capture (force it defends)"]]
    hl = []
    for i, a in enumerate(dv["activities"], start=1):
        data.append([a["activity"], a["define_link"], a["create_link"], a["capture_link"]])
        if a.get("highlight"):
            hl.append(i)
    fitted_table(s, x, y + 1.00, w, data, [2.70, 2.85, 2.85, 2.85], 5.70, hl, row_min=0.52)
    bar(s, x, 5.85, w, first(dv["activities_takeaway"]), 0)


def fit_slide(d, spec, src):
    dv = spec["deliver"]
    s = d.content(spec["titles"]["fit"], source=src, sub=questions["fit"])
    x, w = d.left(s), d.width(s)
    cw = (w - 0.25) / 2
    y0 = d.top(s)
    cards = [
        ("Choices that reinforce each other", red, as_list(dv["reinforcement"])),
        ("What the firm chooses not to do", rose, dv["not_do"]),
        ("Activities that work against the position", skibo, as_list(dv.get("misaligned", ["None material"]))),
        ("Judgment: organized to execute?", iron, as_list(dv.get("judgment", []))),
    ]
    for i, (head, fill, items) in enumerate(cards):
        cx = x + (i % 2) * (cw + 0.25)
        cy = y0 + (i // 2) * 2.12
        stage_card(s, cx, cy, cw, head, fill, items, 1.47)
    bar(s, x, 5.85, w, first(dv["takeaway"]), 1)


def build(spec_path, out_path):
    spec = json.load(open(spec_path))
    c = spec["company"]
    titles = {k: v.format(c=c) for k, v in default_titles.items()}
    titles.update(spec.get("titles", {}))
    spec["titles"] = titles
    src = "Source: " + spec.get("source", "company filings and reputable press; figures marked est. are estimates")
    d = StrategyDeck(assets={
        "tartan_strip": os.path.join(brand_dir, "tartan-strip.png"),
        "t_logo": os.path.join(brand_dir, "tepper-T.png"),
        "title_background": os.path.join(brand_dir, "title-background.png"),
    })
    by = f"By {spec['author']}" if spec.get("author") else ""
    d.title_slide(f"{c}: Strategy Analysis", "Define, Create, Capture and Deliver", by)
    statements_slide(d, spec, src)
    chain_slide(d, spec, src)
    define_slide(d, spec, src)
    create_slide(d, spec, src)
    vri_slide(d, spec, src)
    capture_slide(d, spec, src)
    activities_slide(d, spec, src)
    fit_slide(d, spec, src)
    problems = lint(d, verbose=False)
    for p in problems:
        print("WARNING:", p)
    d.save(out_path)
    print(out_path)
    return problems


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit("usage: python build_deck.py spec.json output.pptx")
    build(sys.argv[1], sys.argv[2])
