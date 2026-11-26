"""LightCastle Partners (LCP) deck library.

Rebuilds the LightCastle client-deck look (learned from the "Meta SBGA Trainer
Onboarding Deck") as native, editable python-pptx shapes, tables and charts.

Usage:
    import sys; sys.path.insert(0, "<skill>/scripts")
    from lcp_deck import *
    d = Deck(client_logo="/path/meta.png")
    d.cover("Programme Name In Caps", "Subtitle", "June 2026")
    d.toc(["Overview", "Approach", "Timeline"])
    s = d.content("Programme Overview")
    bullets(s, 0.6, 1.5, 11.8, 1.2, ["Point one", "Point **two**"])
    d.closing()
    d.save("out.pptx"); lint(d); render("out.pptx")

Text markup in every helper: **bold dark**, [[bold blue key phrase]].
All variable names are lowercase snake_case on purpose.
"""
import copy
import glob
import os
import re
import shutil
import subprocess

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

# ---------------------------------------------------------------- brand ----
here = os.path.dirname(os.path.abspath(__file__))
asset_dir = os.path.join(os.path.dirname(here), "brand-assets")

font = "Lato"


def _c(h):
    return RGBColor.from_string(h)


blue = _c("0473BA")        # primary LCP blue: slide titles, TOC chips, strip top
sky = _c("429DE5")         # light blue on navy: cover subtitle, "Thank You", accents
navy = _c("073763")        # table headers, card titles, TOC labels
deep = _c("061B39")        # strip bottom, darkest navy
mid = _c("0B5394")         # bullet squares, second band colour
soft = _c("3D85C6")        # third band colour
accent = _c("3C78D8")      # [[key phrase]] highlight in body text
gantt_fill = _c("2E75B6")  # filled gantt cells
gantt_head = _c("0F243E")  # gantt header row
gantt_group = _c("244062") # gantt section rows
zebra_blue = _c("DCE6F2")  # gantt zebra rows
panel = _c("F3F3F3")       # light grey cards and zebra rows
pale = _c("EEF5FD")        # pale blue highlight cells
ice = _c("CFE2F3")         # thin title rule (alt) and light fills
dash_blue = _c("9FC5E8")   # dashed row separators
rule = _c("AEAEAF")        # title rule, dotted TOC leaders, dashed card separators
ink = _c("1C1C1C")         # body text
gray = _c("7F7F7F")        # source lines, captions
white = _c("FFFFFF")
off_white = _c("F2F2F2")
band_fills = [navy, mid, soft, blue, sky]
chart_colors = ["00A2FF", "073163", "3C78D8", "073763", "9FC5E8", "A3A3A3"]
neon = _c("00A2FF")         # highlighted / above-average / good data
dark = _c("073163")         # bad data and everything else

# ---------------------------------------------------------------- grid -----
slide_w, slide_h = 13.333, 7.5
strip_x, strip_w = 12.91, 0.425
content_l, content_r = 0.6, 12.45
title_y, title_h = 0.50, 0.55
rule_y = 1.13
body_y = 1.45
body_max = 7.05
body_max_src = 6.80
source_y = 6.88

_mark = re.compile(r"(\*\*.+?\*\*|\[\[.+?\]\])")
_align = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT, "j": PP_ALIGN.JUSTIFY}
_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


# ----------------------------------------------------------- text basics ----
def runs(p, text, size, color, bold=False, italic=False, highlight=None):
    """Append runs honouring **bold dark** and [[bold blue]] markup."""
    for seg in _mark.split(str(text)):
        if not seg:
            continue
        c, b = color, bold
        if seg.startswith("**") and seg.endswith("**"):
            seg, b = seg[2:-2], True
        elif seg.startswith("[[") and seg.endswith("]]"):
            seg, b, c = seg[2:-2], True, (highlight or accent)
        r = p.add_run()
        r.text = seg
        f = r.font
        f.name = font
        f.size = Pt(size)
        f.bold = b
        f.italic = italic
        f.color.rgb = c
    return p


def tbox(slide, l, t, w, h, anchor="t", wrap=True, inset=0.0):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = _anchor.get(anchor, anchor)
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(inset)
    return tb, tf


def para(tf, text, size=12, color=ink, bold=False, first=False, after=4, before=0,
         align="l", italic=False, line=1.15):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = _align.get(align, align)
    p.space_after = Pt(after)
    p.space_before = Pt(before)
    if line:
        p.line_spacing = line
    return runs(p, text, size, color, bold, italic)


def text(slide, l, t, w, h, content, size=12, color=ink, bold=False, italic=False,
         align="l", anchor="t", after=4, line=1.15, wrap=True):
    """One-call text box; content is a string or list of paragraph strings."""
    tb, tf = tbox(slide, l, t, w, h, anchor=anchor, wrap=wrap)
    items = content if isinstance(content, (list, tuple)) else [content]
    for i, c in enumerate(items):
        para(tf, c, size, color, bold, first=(i == 0), after=(after if i < len(items) - 1 else 0),
             align=align, italic=italic, line=line)
    return tb


def _bullet_ppr(p, char="■", color=mid, indent=0.19, size_pct=None):
    ppr = p._p.get_or_add_pPr()
    ppr.set("marL", str(int(Inches(indent))))
    ppr.set("indent", str(-int(Inches(indent))))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buFont", "a:buSzPct"):
        for el in ppr.findall(qn(tag)):
            ppr.remove(el)
    buclr = etree.SubElement(ppr, qn("a:buClr"))
    etree.SubElement(buclr, qn("a:srgbClr")).set("val", str(color))
    if size_pct:
        etree.SubElement(ppr, qn("a:buSzPct")).set("val", str(int(size_pct * 1000)))
    etree.SubElement(ppr, qn("a:buFont")).set("typeface", font)
    etree.SubElement(ppr, qn("a:buChar")).set("char", char)


def bullets(slide, x, y, w, h, items, size=12, color=ink, gap=4, marker="■",
            marker_color=mid, line=1.15):
    """Square-bullet list, the LCP default. items: strings, or (text, level) tuples."""
    tb, tf = tbox(slide, x, y, w, h)
    for i, it in enumerate(items):
        txt, lvl = (it if isinstance(it, tuple) else (it, 0))
        p = para(tf, txt, size, color, first=(i == 0), after=(gap if i < len(items) - 1 else 0), line=line)
        if lvl:
            _bullet_ppr(p, "-", ink, indent=0.19 + 0.25 * lvl)
            p._p.get_or_add_pPr().set("marL", str(int(Inches(0.19 + 0.25 * lvl))))
        else:
            _bullet_ppr(p, marker, marker_color, size_pct=85)
    return tb


def dash_list(slide, x, y, w, h, items, size=12, color=ink, gap=2):
    """Hyphen list used inside LCP tables and cards ("-  item")."""
    tb, tf = tbox(slide, x, y, w, h)
    for i, it in enumerate(items):
        p = para(tf, it, size, color, first=(i == 0), after=(gap if i < len(items) - 1 else 0))
        _bullet_ppr(p, "-", color, indent=0.17)
    return tb


# ---------------------------------------------------------- shape basics ----
def _dash(sh_line, dash):
    if dash:
        sh_line.dash_style = {"dash": MSO_LINE_DASH_STYLE.DASH, "dot": MSO_LINE_DASH_STYLE.ROUND_DOT,
                              "sysdash": MSO_LINE_DASH_STYLE.SQUARE_DOT}.get(dash, dash)


def rect(slide, l, t, w, h, fill=None, line=None, lw=1.0, dash=None, shape=MSO_SHAPE.RECTANGLE):
    sh = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
        _dash(sh.line, dash)
    st = sh._element.find(qn("p:style"))
    if st is not None:
        sh._element.remove(st)   # no theme shadow / effects
    return sh


def gradient_rect(slide, l, t, w, h, top=blue, bottom=deep, angle=90):
    """Rectangle with a linear gradient (angle 90 = top to bottom)."""
    sh = rect(slide, l, t, w, h, fill=top)
    sppr = sh._element.spPr
    for el in sppr.findall(qn("a:solidFill")):
        sppr.remove(el)
    grad = etree.fromstring(
        '<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1">'
        f'<a:gsLst><a:gs pos="0"><a:srgbClr val="{top}"/></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="{bottom}"/></a:gs></a:gsLst>'
        f'<a:lin ang="{int(angle * 60000)}" scaled="0"/></a:gradFill>')
    sppr.insert(list(sppr).index(sppr.find(qn("a:prstGeom"))) + 1, grad)
    return sh


def shape_text(sh, content, size=12, color=white, bold=False, align="c", anchor="m", inset=0.06):
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = _anchor.get(anchor, anchor)
    tf.margin_left = tf.margin_right = Inches(inset)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    items = content if isinstance(content, (list, tuple)) else [content]
    for i, c in enumerate(items):
        para(tf, c, size, color, bold, first=(i == 0), after=0, align=align, line=1.0)
    return sh


def line(slide, x1, y1, x2, y2, color=rule, lw=0.75, dash=None):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = color
    ln.line.width = Pt(lw)
    _dash(ln.line, dash)
    st = ln._element.find(qn("p:style"))
    if st is not None:
        ln._element.remove(st)
    return ln


def hline(slide, x1, y, x2, color=rule, lw=0.75, dash=None):
    return line(slide, x1, y, x2, y, color, lw, dash)


def picture_fit(slide, path, l, t, w, h, align="c"):
    """Place an image inside a box, keeping aspect ratio."""
    from PIL import Image
    iw, ih = Image.open(path).size
    r = min(w / iw, h / ih)
    pw, ph = iw * r, ih * r
    x = l + {"l": 0, "c": (w - pw) / 2, "r": w - pw}[align]
    y = t + (h - ph) / 2
    return slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(pw), Inches(ph))


def source(slide, txt, y=source_y):
    """Bottom-left source line, 9 pt grey."""
    return text(slide, content_l, y, 11.7, 0.22, txt, size=9, color=gray, line=1.0)


# ------------------------------------------------------------------ deck ----
class Deck:
    def __init__(self, client_logo=None, client_logo_white=None, year=2026, assets=None):
        self.prs = Presentation()
        self.prs.slide_width = Inches(slide_w)
        self.prs.slide_height = Inches(slide_h)
        self.layout = self.prs.slide_layouts[6]
        self.a = {
            "logo": os.path.join(asset_dir, "lcp_logo.png"),
            "logo_white": os.path.join(asset_dir, "lcp_logo_white.png"),
            "icon_white": os.path.join(asset_dir, "lcp_icon_white.png"),
            "icon_blue": os.path.join(asset_dir, "lcp_icon_blue.png"),
            "cover_bg": os.path.join(asset_dir, "cover_bg.png"),
            "closing_bg": os.path.join(asset_dir, "closing_bg.png"),
            "toc_bg": os.path.join(asset_dir, "toc_bg.png"),
            "light_bg": os.path.join(asset_dir, "light_bg.png"),
        }
        if assets:
            self.a.update(assets)
        self.client_logo = client_logo
        self.client_logo_white = client_logo_white
        self.year = year
        self.meta = {}
        self.toc_items = None
        self._set_theme_font()

    def _set_theme_font(self):
        master = self.prs.slide_masters[0]
        theme_part = [r.target_part for r in master.part.rels.values() if r.reltype.endswith("/theme")][0]
        root = etree.fromstring(theme_part.blob)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        for tag in ("a:majorFont", "a:minorFont"):
            for el in root.findall(f".//{tag}/a:latin", ns):
                el.set("typeface", font)
        theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _new(self, kind, **info):
        s = self.prs.slides.add_slide(self.layout)
        self.meta[s.slide_id] = dict(kind=kind, **info)
        return s

    def _bg(self, s, path):
        s.shapes.add_picture(path, 0, 0, Inches(slide_w), Inches(slide_h))

    def _strip(self, s):
        """Right vertical blue strip: icon, rotated copyright, page number."""
        n = len(self.prs.slides)
        gradient_rect(s, strip_x, 0, strip_w, slide_h, blue, deep)
        s.shapes.add_picture(self.a["icon_white"], Inches(12.98), Inches(0.15), Inches(0.29), Inches(0.29))
        cp = text(s, 11.66, 5.22, 2.92, 0.24, f"LightCastle Partners © {self.year}", size=8,
                  color=white, align="c", anchor="m", line=1.0)
        cp.rotation = 270
        text(s, strip_x, 6.98, strip_w, 0.35, str(n), size=8, color=white, align="c", anchor="m", line=1.0)

    def _logos_band(self, s, band_y, logo_y, logo_w, client_box):
        rect(s, 0, band_y, slide_w, slide_h - band_y, fill=white)
        s.shapes.add_picture(self.a["logo"], Inches(logo_y[0]), Inches(logo_y[1]), Inches(logo_w))
        if self.client_logo:
            picture_fit(s, self.client_logo, *client_box, align="r")

    # -- slide types --------------------------------------------------------
    def cover(self, title, subtitle=None, date=None):
        """Navy tiled background, white CAPS title, sky subtitle, white logo band."""
        s = self._new("cover")
        self._bg(s, self.a["cover_bg"])
        title = title.upper()
        n = _est_lines(title, 34, 10.2, True)
        size = 34 if n <= 2 else 28
        th = max(1, _est_lines(title, size, 10.2, True)) * size * 1.15 / 72
        text(s, 0.93, 0.95, 10.4, th + 0.1, title, size=size, color=white, bold=True, line=1.05)
        y = 0.95 + th + 0.35
        if subtitle:
            text(s, 0.93, y, 11.4, 0.5, subtitle, size=24, color=sky, line=1.0)
            y += 0.95
        if date:
            text(s, 0.93, y, 4.0, 0.4, date, size=16, color=white, line=1.0)
        rect(s, 0, 5.71, slide_w, 0.07, fill=sky)
        self._logos_band(s, 5.77, (1.0, 6.37), 2.48, (9.9, 6.19, 2.2, 0.9))
        return s

    def toc(self, items, pages=None, start_page=None):
        """Table of contents. items: list of section labels. pages default to the
        slide numbers you pass, or are guessed as consecutive from this slide + 1."""
        s = self._new("toc")
        self._bg(s, self.a["toc_bg"])
        tb, tf = tbox(s, 0.6, 5.62, 4.42, 1.25)
        para(tf, "Table of", 24, sky, first=True, after=0, line=1.0)
        para(tf, "Contents", 40, white, bold=True, after=0, line=1.0)
        rect(s, 3.35, 6.62, 0.91, 0.03, fill=sky)
        rect(s, 0, 7.42, 5.04, 0.08, fill=sky)
        start = start_page or len(self.prs.slides) + 1
        pages = pages or [start + i for i in range(len(items))]
        step = min(0.65, 5.2 / max(1, len(items)))
        for i, (lab, pg) in enumerate(zip(items, pages)):
            y = 1.57 + i * step
            chip = rect(s, 4.91, y, 0.59, 0.32, fill=blue)
            shape_text(chip, f"{i + 1}.", 12, white, bold=True, align="r", inset=0.08)
            lw = min(5.4, _text_w(lab, 12, True) + 0.12)
            text(s, 5.72, y - 0.04, lw + 0.1, 0.4, lab, 12, navy, bold=True, anchor="m", line=1.0, wrap=False)
            hline(s, 5.72 + lw + 0.12, y + 0.16, 11.85, rule, 0.75, "dash")
            text(s, 11.96, y - 0.04, 0.67, 0.4, f"{int(pg):02d}", 12, _c("002060"), bold=True, anchor="m", line=1.0)
        self._strip(s)
        return s

    def divider(self, number, title, subtitle=None):
        """Section divider on the light tiled background."""
        s = self._new("divider")
        self._bg(s, self.a["light_bg"])
        chip = rect(s, 0.93, 2.7, 1.0, 1.0, fill=blue)
        shape_text(chip, f"{int(number):02d}", 32, white, bold=True)
        text(s, 2.2, 2.62, 10.0, 0.75, title, 36, navy, bold=True, anchor="m", line=1.0)
        rect(s, 2.2, 3.45, 0.91, 0.04, fill=sky)
        if subtitle:
            text(s, 2.2, 3.62, 10.0, 0.5, subtitle, 16, ink, line=1.1)
        self._strip(s)
        return s

    def content(self, title, source_text=None, rule_color=rule):
        """Standard white content slide: blue title, thin rule, right strip."""
        s = self._new("content", has_source=bool(source_text), title=title)
        text(s, content_l, title_y, content_r - content_l, title_h, title, 30, blue, bold=True,
             anchor="b", line=1.0, wrap=False)
        hline(s, 0.54, rule_y, content_r, rule_color, 0.75)
        self._strip(s)
        if source_text:
            source(s, source_text)
        self.meta[s.slide_id]["n_furniture"] = len(s.shapes)
        return s

    def closing(self, headline="Thank You", offices=None, contact=None):
        """Navy closing slide with office addresses and white logo band."""
        s = self._new("closing")
        self._bg(s, self.a["closing_bg"])
        text(s, 0.69, 0.50, 4.42, 0.8, headline, 40, sky, bold=True, anchor="m", line=1.0, wrap=False)
        rect(s, 0.69 + _text_w(headline, 40, True) + 0.25, 1.05, 0.91, 0.03, fill=sky)
        offices = offices if offices is not None else default_offices
        xs = [0.69, 3.64, 6.41]
        for (lab, addr), x in zip(offices[:3], xs):
            tb, tf = tbox(s, x, 5.26, 2.64, 0.95)
            para(tf, lab, 10, sky, first=True, after=0, line=1.15)
            para(tf, addr, 10, white, after=0, line=1.15)
        contact = contact if contact is not None else default_contact
        if contact:
            tb, tf = tbox(s, 9.57, 5.26, 3.3, 0.95)
            for i, (k, v) in enumerate(contact):
                p = para(tf, f"{k}: ", 10, white, first=(i == 0), after=0, line=1.15)
                runs(p, v, 10, sky if k in ("Web", "Other Platform") else white)
        self._logos_band(s, 6.49, (0.65, 6.80), 1.84, (10.9, 6.64, 1.8, 0.71))
        return s

    def save(self, path):
        self.prs.save(path)
        return path


default_offices = [
    ("Bangladesh Office:", "Navana Osman @ Link, Plot 214/D, 3rd Floor, Tejgaon Gulshan link Road Dhaka 1208, Bangladesh."),
    ("Singapore Office:", "36 Robinson Road, #20-01 City House, Singapore 068877"),
    ("USA Office:", "3921 W ST NW, #3, Washington DC, 20007, USA."),
]
default_contact = [
    ("Email", "info@lightcastlepartners.com"),
    ("Web", "www.lightcastlepartners.com"),
    ("Other Platform", "www.investbangladesh.co"),
]


# ------------------------------------------------------------ components ----
def card(slide, x, y, w, h, title, body, size=12, title_size=20, fill=panel, gap=8):
    """Grey panel with navy bold heading and body paragraphs (slide 3 left card)."""
    rect(slide, x, y, w, h, fill=fill)
    th = _est_lines(title, title_size, w - 0.36, True) * title_size * 1.2 / 72
    text(slide, x + 0.18, y + 0.15, w - 0.36, th + 0.05, title, title_size, navy, bold=True, line=1.15)
    paras = body if isinstance(body, (list, tuple)) else [body]
    text(slide, x + 0.18, y + 0.15 + th + 0.22, w - 0.36, h - th - 0.5, paras, size, ink, after=gap, line=1.15)


def header_strip(slide, x, y, w, title, h=0.55, size=20, fill=panel):
    """Grey title bar with navy bold text (slide 3 'Three Core Objectives')."""
    r = rect(slide, x, y, w, h, fill=fill)
    shape_text(r, title, size, navy, bold=True, align="l", inset=0.12)
    return r


def dashed_items(slide, x, y, w, items, size=12, item_h=None, h=None, sep=rule):
    """Square-bullet items separated by long dashed rules (slide 3 right column).
    items: '[[Label:]] description' strings."""
    n = len(items)
    item_h = item_h or ((h / n) if h else 1.0)
    for i, it in enumerate(items):
        yy = y + i * item_h
        bullets(slide, x + 0.15, yy + 0.12, w - 0.3, item_h - 0.24, [it], size=size)
        hline(slide, x + 0.1, yy + item_h - 0.03, x + w - 0.1, sep, 1.0, "dash")


def panel_box(slide, x, y, w, h, title, items, size=12, title_size=18):
    """'What this means for you' callout: grey panel, navy title, bullets, navy base bar."""
    rect(slide, x, y, w, h, fill=panel)
    text(slide, x + 0.12, y + 0.16, w - 0.24, 0.4, title, title_size, navy, bold=True, line=1.0)
    bullets(slide, x + 0.18, y + 0.75, w - 0.4, h - 1.25, items, size=size, gap=10)
    rect(slide, x + 0.35, y + h - 0.48, w - 0.7, 0.03, fill=_c("15386A"))


def kpi_row(slide, x, y, w, items, h=1.35, gap=0.25, num_size=32, fill=panel):
    """Row of stat tiles: (value, caption). Value in LCP blue, caption in ink."""
    n = len(items)
    cw = (w - gap * (n - 1)) / n
    for i, (val, cap) in enumerate(items):
        xx = x + i * (cw + gap)
        rect(slide, xx, y, cw, h, fill=fill)
        rect(slide, xx, y, 0.06, h, fill=blue)
        text(slide, xx + 0.22, y + 0.12, cw - 0.32, 0.6, val, num_size, blue, bold=True, anchor="m", line=1.0)
        text(slide, xx + 0.22, y + 0.12 + 0.62, cw - 0.32, h - 0.82, cap, 11, ink, line=1.1)


def icon_steps(slide, x, y, w, steps, h=1.6, fills=None):
    """Horizontal numbered steps: blue square number chips over short text."""
    n = len(steps)
    gap = 0.25
    cw = (w - gap * (n - 1)) / n
    for i, st in enumerate(steps):
        head, body = st if isinstance(st, tuple) else (st, "")
        xx = x + i * (cw + gap)
        f = (fills or band_fills)[i % len(fills or band_fills)]
        top = rect(slide, xx, y, cw, 0.62, fill=f)
        shape_text(top, f"{i + 1}. {head}", 14, white, bold=True, align="l", inset=0.14)
        rect(slide, xx, y + 0.62, cw, h - 0.62, fill=panel)
        if body:
            text(slide, xx + 0.14, y + 0.74, cw - 0.28, h - 0.84, body, 11, ink, line=1.15)


# ---------------------------------------------------------------- tables ----
_no_style = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"   # "No Style, No Grid"


def _cell_border(cell, side, color=None, w=1.0, dash=None):
    """side in L R T B. color None = no line."""
    tcpr = cell._tc.get_or_add_tcPr()
    tag = qn(f"a:ln{side}")
    for el in tcpr.findall(tag):
        tcpr.remove(el)
    ln = etree.SubElement(tcpr, tag)
    if color is None:
        ln.set("w", "0")
        etree.SubElement(ln, qn("a:noFill"))
    else:
        ln.set("w", str(int(Pt(w))))
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr")).set("val", str(color))
        if dash:
            etree.SubElement(ln, qn("a:prstDash")).set("val", dash)
    # schema order: lnL lnR lnT lnB ... then fill; move fills after borders
    for f in ("a:solidFill", "a:noFill"):
        for el in tcpr.findall(qn(f)):
            tcpr.remove(el)
            tcpr.append(el)


def _cell(cell, content, size=12, color=ink, bold=False, fill=None, align="l", anchor="m",
          dash_items=False, inset=0.08):
    if fill is None:
        cell.fill.background()
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    cell.margin_left = cell.margin_right = Inches(inset)
    cell.margin_top = cell.margin_bottom = Inches(0.05)
    cell.vertical_anchor = _anchor.get(anchor, anchor)
    tf = cell.text_frame
    tf.word_wrap = True
    items = content if isinstance(content, (list, tuple)) else [content]
    for i, c in enumerate(items):
        p = para(tf, c, size, color, bold, first=(i == 0), after=0, align=align, line=1.1)
        if dash_items:
            _bullet_ppr(p, "-", color, indent=0.17)


def _new_table(slide, x, y, w, rows, cols, col_w, row_h):
    gf = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(row_h * rows))
    tbl = gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    tblpr = tbl._tbl.tblPr
    sid = tblpr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblpr, qn("a:tableStyleId"))
    sid.text = _no_style
    tot = sum(col_w)
    for i, cw in enumerate(col_w):
        tbl.columns[i].width = Inches(w * cw / tot)
    for r in tbl.rows:
        r.height = Inches(row_h)
    return gf, tbl


def data_table(slide, x, y, w, data, col_w=None, row_h=0.55, size=12, head_size=15,
               first_col_bold=True, zebra=True, dash_lists=True):
    """LCP data table (slide 4): navy header, white bold header text, zebra F3F3F3,
    bold first column. A cell given as a list renders as a "-" list."""
    rows, cols = len(data), len(data[0])
    gf, tbl = _new_table(slide, x, y, w, rows, cols, col_w or [1] * cols, row_h)
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            c = tbl.cell(ri, ci)
            if ri == 0:
                _cell(c, val, head_size, off_white, True, navy, "c")
                for sd in "LRT":
                    _cell_border(c, sd, white, 2)
                _cell_border(c, "B", sky, 2)
            else:
                f = panel if (zebra and ri % 2 == 0) else None
                is_list = isinstance(val, (list, tuple)) and dash_lists
                _cell(c, val, size, ink, first_col_bold and ci == 0, f, "l", dash_items=is_list)
                for sd in "LR":
                    _cell_border(c, sd, white, 2)
                _cell_border(c, "T", None)
                _cell_border(c, "B", None)
    return gf


def compare_table(slide, x, y, w, header, rows, col_w=None, row_h=0.6, size=12, label_size=14,
                  highlight_rows=(1, 3)):
    """Comparison grid (slide 6): no fills in the header, navy bold header text,
    bold row labels, alternating pale-blue body cells, thin grey row rules."""
    data = [header] + rows
    n_r, n_c = len(data), len(header)
    gf, tbl = _new_table(slide, x, y, w, n_r, n_c, col_w or [1] * n_c, row_h)
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            c = tbl.cell(ri, ci)
            if ri == 0:
                _cell(c, val, label_size, _c("061B39"), True, None, "l")
            elif ci == 0:
                _cell(c, val, label_size, ink, True, None, "l")
            else:
                f = pale if ri in highlight_rows else None
                _cell(c, val, size, _c("1E1E1E"), False, f, "l")
            for sd in "LRT":
                _cell_border(c, sd, None)
            _cell_border(c, "B", _c("E5E5E5"), 0.75)
    return gf


def band_table(slide, x, y, w, rows, label_w=1.75, row_h=None, h=None, size=12, label_size=15):
    """Coloured label column + bullet column with dashed separators (slide 5).
    rows: (label, sublabel, [items]). Label fills step navy, mid, soft."""
    n = len(rows)
    row_h = row_h or ((h / n) if h else 1.5)
    for i, (lab, sub, items) in enumerate(rows):
        yy = y + i * row_h
        box = rect(slide, x, yy, label_w, row_h - 0.04, fill=band_fills[i % 3])
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(tf, lab, label_size, white, first=True, after=0, align="c", line=1.0)
        if sub:
            para(tf, sub, 11, white, after=0, align="c", italic=True, line=1.0)
        if i % 2 == 1:
            rect(slide, x + label_w + 0.02, yy, w - label_w - 0.02, row_h - 0.04, fill=_c("F9F9F9"))
        bullets(slide, x + label_w + 0.2, yy + 0.12, w - label_w - 0.35, row_h - 0.25, items, size=size, gap=1)
        hline(slide, x + label_w, yy + row_h - 0.02, x + w, dash_blue, 1.0, "dash")


def gantt(slide, x, y, w, periods, groups, label_w=3.2, row_h=0.215, head_h=0.25, size=10):
    """Timeline table (slide 7). groups: [(group_name, [(activity, [period indexes])])].
    Dark header, section rows in steel navy, filled cells in blue, light blue zebra."""
    n_rows = 1 + sum(1 + len(acts) for _, acts in groups)
    cols = 1 + len(periods)
    col_w = [label_w] + [(w - label_w) / len(periods)] * len(periods)
    gf, tbl = _new_table(slide, x, y, w, n_rows, cols, col_w, row_h)
    tbl.rows[0].height = Inches(head_h)
    for ci, lab in enumerate(["Activity"] + list(periods)):
        c = tbl.cell(0, ci)
        _cell(c, lab, size + 2, white, True, gantt_head, "c")
    r = 1
    for gname, acts in groups:
        c0 = tbl.cell(r, 0)
        c0.merge(tbl.cell(r, cols - 1))
        _cell(c0, gname.upper(), size, white, True, gantt_group, "l")
        r += 1
        for ai, (act, filled) in enumerate(acts):
            base = zebra_blue if ai % 2 == 1 else white
            _cell(tbl.cell(r, 0), act, size, _c("1A1A1A"), False, base, "l")
            for pi in range(len(periods)):
                _cell(tbl.cell(r, pi + 1), "", size, ink, False, gantt_fill if pi in filled else base)
            r += 1
    for row in tbl.rows:
        for c in row.cells:
            for sd in "LRTB":
                _cell_border(c, sd, _c("BFBFBF"), 0.5)
    return gf


# ---------------------------------------------------------------- charts ----
def _style_chart(chart, size=11, legend=True, gridlines=True):
    chart.font.name = font
    chart.font.size = Pt(size)
    chart.font.color.rgb = ink
    chart.has_title = False
    chart.has_legend = legend
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(size)
    try:
        va = chart.value_axis
        va.has_major_gridlines = gridlines
        if gridlines:
            va.major_gridlines.format.line.color.rgb = _c("E7E7E7")
            va.major_gridlines.format.line.width = Pt(0.75)
        va.format.line.fill.background()
        va.tick_labels.font.color.rgb = _c("A3A3A3")
        va.tick_labels.font.size = Pt(size - 1)
        ca = chart.category_axis
        ca.format.line.color.rgb = _c("333333")
        ca.tick_labels.font.size = Pt(size)
        from pptx.enum.chart import XL_TICK_MARK
        ca.major_tick_mark = XL_TICK_MARK.NONE
        va.major_tick_mark = XL_TICK_MARK.NONE
    except Exception:
        pass


def bar_chart(slide, x, y, w, h, cats, series, colors=None, horizontal=False, stacked=False,
              labels=True, fmt='0', legend=None, highlight=None, size=11, gap=60):
    """Native editable bar/column chart in the LCP (Datawrapper-like) palette.
    series: {"name": [values]}. highlight: set of category indexes coloured neon
    blue (good / above average), others dark navy (single-series only)."""
    cd = CategoryChartData()
    cd.categories = cats
    for k, v in series.items():
        cd.add_series(k, v)
    if horizontal:
        kind = XL_CHART_TYPE.BAR_STACKED if stacked else XL_CHART_TYPE.BAR_CLUSTERED
    else:
        kind = XL_CHART_TYPE.COLUMN_STACKED if stacked else XL_CHART_TYPE.COLUMN_CLUSTERED
    gf = slide.shapes.add_chart(kind, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    ch = gf.chart
    legend = (len(series) > 1) if legend is None else legend
    _style_chart(ch, size, legend)
    plot = ch.plots[0]
    plot.gap_width = gap
    if stacked:
        plot.overlap = 100
    cols = colors or chart_colors
    for i, s in enumerate(plot.series):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = _c(cols[i % len(cols)])
        if highlight is not None and len(series) == 1:
            for pi in range(len(cats)):
                pt = s.points[pi]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = neon if pi in highlight else dark
    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format = fmt
        dl.number_format_is_linked = False
        dl.font.size = Pt(size - 1)
        dl.font.color.rgb = white if stacked else ink
        dl.position = XL_LABEL_POSITION.CENTER if stacked else XL_LABEL_POSITION.OUTSIDE_END
    if horizontal:
        ch.category_axis.reverse_order = True
    return gf


def line_chart(slide, x, y, w, h, cats, series, colors=None, labels=False, fmt='0', size=11,
               legend=None, width=2.5, markers=True):
    """Native editable line chart in the LCP palette."""
    cd = CategoryChartData()
    cd.categories = cats
    for k, v in series.items():
        cd.add_series(k, v)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS if markers else XL_CHART_TYPE.LINE,
                                Inches(x), Inches(y), Inches(w), Inches(h), cd)
    ch = gf.chart
    _style_chart(ch, size, (len(series) > 1) if legend is None else legend)
    cols = colors or chart_colors
    for i, s in enumerate(ch.plots[0].series):
        c = _c(cols[i % len(cols)])
        s.format.line.color.rgb = c
        s.format.line.width = Pt(width)
        s.smooth = False
        if markers:
            from pptx.enum.chart import XL_MARKER_STYLE
            s.marker.style = XL_MARKER_STYLE.CIRCLE
            s.marker.size = 7
            s.marker.format.fill.solid()
            s.marker.format.fill.fore_color.rgb = c
            s.marker.format.line.color.rgb = c
    if labels:
        p = ch.plots[0]
        p.has_data_labels = True
        p.data_labels.number_format = fmt
        p.data_labels.number_format_is_linked = False
        p.data_labels.position = XL_LABEL_POSITION.ABOVE
        p.data_labels.font.size = Pt(size - 1)
    return gf


def chart_header(slide, x, y, w, title, subtitle=None):
    """Datawrapper-style chart heading used above LCP charts: bold title, plain subtitle."""
    text(slide, x, y, w, 0.35, title, 15, ink, bold=True, line=1.0)
    if subtitle:
        text(slide, x, y + 0.36, w, 0.3, subtitle, 11, _c("595959"), line=1.0)


# ------------------------------------------------------------ measurement ---
_font_cache = {}


def _pil_font(size, bold):
    from PIL import ImageFont
    key = (round(size * 4), bold)
    if key in _font_cache:
        return _font_cache[key]
    cands = [os.path.expanduser("~/Library/Fonts/" + ("Lato-Bold.ttf" if bold else "Lato-Regular.ttf")),
             "/Library/Fonts/" + ("Lato-Bold.ttf" if bold else "Lato-Regular.ttf")]
    f = None
    for c in cands:
        if os.path.exists(c):
            f = ImageFont.truetype(c, int(size * 4))
            break
    _font_cache[key] = f
    return f


def _text_w(txt, size_pt, bold=False):
    """Width in inches of one line of Lato text."""
    f = _pil_font(size_pt, bold)
    if f is None:
        return len(txt) * size_pt / 72.0 * (0.54 if bold else 0.50)
    return f.getlength(txt) / 4 / 72.0


def _est_lines(txt, size_pt, width_in, bold=False):
    if width_in <= 0:
        return 99
    lines = 0
    for seg in str(txt).split("\n"):
        words = seg.split(" ")
        cur = ""
        n = 1
        for wd in words:
            trial = (cur + " " + wd) if cur else wd
            if cur and _text_w(trial, size_pt, bold) > width_in:
                n += 1
                cur = wd
            else:
                cur = trial
        lines += n
    return lines


def _tf_need(tf, width_in, wrap=True):
    hgt = 0.0
    widest = 0.0
    for p in tf.paragraphs:
        txt = "".join(r.text for r in p.runs)
        sizes = [r.font.size.pt for r in p.runs if r.font.size] or [12]
        sz = max(sizes)
        bold = any(r.font.bold for r in p.runs)
        ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
        indent = 0.19 if p._p.pPr is not None and p._p.pPr.find(qn("a:buChar")) is not None else 0
        n = _est_lines(txt, sz, width_in - indent, bold) if wrap else txt.count("\n") + 1
        hgt += n * sz * 1.2 * ls / 72.0
        if p.space_after is not None:
            hgt += p.space_after.pt / 72.0
        widest = max(widest, _text_w(txt, sz, bold) + indent)
    if tf.paragraphs and tf.paragraphs[-1].space_after is not None:
        hgt -= tf.paragraphs[-1].space_after.pt / 72.0
    return hgt, widest


def _box(sh):
    return (Emu(sh.left).inches, Emu(sh.top).inches, Emu(sh.width).inches, Emu(sh.height).inches)


# ------------------------------------------------------------- QA loop ------
def lint(deck, verbose=True):
    """Static checks: off-slide shapes, body content in the title zone or the
    right strip, text overflow, one-line titles that are too long, and banned
    em/en dash characters. Returns a list of problem strings."""
    problems = []
    for idx, s in enumerate(deck.prs.slides, start=1):
        meta = deck.meta.get(s.slide_id, {})
        kind = meta.get("kind")
        shapes = list(s.shapes)
        nf = meta.get("n_furniture", 0)
        for sh in shapes:
            if sh.left is None:
                continue
            l, t, w, h = _box(sh)
            if sh.rotation in (0, 0.0) and (l < -0.01 or t < -0.01 or l + w > slide_w + 0.02 or t + h > slide_h + 0.02):
                problems.append(f"slide {idx}: '{sh.name}' runs off the slide")
            txts = []
            if sh.has_text_frame:
                txts.append(sh.text_frame.text)
            if sh.has_table:
                txts += [c.text for r in sh.table.rows for c in r.cells]
            for t_ in txts:
                if "—" in t_ or "–" in t_:
                    problems.append(f"slide {idx}: dash character in '{t_[:50]}'")
        if kind == "content":
            title = meta.get("title", "")
            if _text_w(title, 30, True) > content_r - content_l:
                problems.append(f"slide {idx}: title too long for one line at 30 pt: '{title}'")
            lim = body_max_src if meta.get("has_source") else body_max
            for sh in shapes[nf:]:
                if sh.left is None or sh.rotation not in (0, 0.0):
                    continue
                l, t, w, h = _box(sh)
                tag = f"slide {idx}: '{sh.name}'"
                if l + w > strip_x - 0.1:
                    problems.append(f"{tag} runs into the right strip (ends {l + w:.2f})")
                if t < rule_y + 0.08:
                    problems.append(f"{tag} starts in the title zone ({t:.2f})")
                if t + h > lim + 0.03:
                    problems.append(f"{tag} ends at {t + h:.2f}, below the body limit {lim:.2f}")
        for sh in shapes:
            if not sh.has_text_frame or sh.left is None or not sh.text_frame.text.strip():
                continue
            if sh.rotation not in (0, 0.0):
                continue
            tf = sh.text_frame
            l, t, w, h = _box(sh)
            ml = Emu(tf.margin_left or 0).inches + Emu(tf.margin_right or 0).inches
            mt = Emu(tf.margin_top or 0).inches + Emu(tf.margin_bottom or 0).inches
            need_h, widest = _tf_need(tf, w - ml, tf.word_wrap is not False)
            if tf.word_wrap is False and l + widest > slide_w - 0.05:
                problems.append(f"slide {idx}: one-line text runs past the slide: '{tf.text[:40]}'")
            elif tf.word_wrap is not False and need_h > (h - mt) * 1.12 + 0.05:
                problems.append(f"slide {idx}: text likely overflows ({need_h:.2f}in needed, {h - mt:.2f}in box): "
                                f"'{tf.text[:50]}'")
        for sh in shapes:
            if not sh.has_table:
                continue
            l, t, w, h = _box(sh)
            est = 0.0
            for row in sh.table.rows:
                need = Emu(row.height).inches
                for ci, c in enumerate(row.cells):
                    if c.is_spanned or c.is_merge_origin:
                        continue
                    cw = Emu(sh.table.columns[ci].width).inches - 0.16
                    nh, _ = _tf_need(c.text_frame, cw)
                    need = max(need, nh + 0.10)
                est += need
            lim = body_max_src if meta.get("has_source") else body_max
            if t + est > lim + 0.03:
                problems.append(f"slide {idx}: table will grow to {t + est:.2f}in, past {lim:.2f}in")
    if verbose:
        print(f"lint: {len(problems)} problem(s)")
        for p in problems:
            print("  -", p)
    return problems


def _office():
    office = shutil.which("soffice") or shutil.which("libreoffice")
    for cand in ["/Applications/LibreOffice.app/Contents/MacOS/soffice", "/opt/homebrew/bin/soffice"]:
        if not office and os.path.exists(cand):
            office = cand
    return office


def render(path, out_dir=None, dpi=80):
    """Convert a pptx or docx to PDF with LibreOffice, then one PNG per page.
    Returns (pdf_path, [png paths])."""
    out_dir = out_dir or os.path.join(os.path.dirname(os.path.abspath(path)), "render")
    os.makedirs(out_dir, exist_ok=True)
    office = _office()
    if not office:
        print("render: LibreOffice not found; rely on lint() only")
        return None, []
    subprocess.run([office, "--headless", "--convert-to", "pdf", "--outdir", out_dir, path],
                   check=True, capture_output=True, timeout=300)
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    stem = os.path.join(out_dir, "page")
    for f in glob.glob(stem + "-*.png"):
        os.remove(f)
    pngs = []
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf, stem], check=True)
        pngs = sorted(glob.glob(stem + "-*.png"))
    return pdf, pngs
