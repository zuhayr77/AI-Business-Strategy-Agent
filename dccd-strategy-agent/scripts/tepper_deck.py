"""
tepper_deck.py : standalone Tepper / Carnegie Mellon slide library (python-pptx)

Builds 16:9 Tepper-branded decks from a blank presentation. No donor template
is needed. Brand images (tartan strip, Tepper T, CMU wordmark, title background)
come from a local asset folder; if one is missing, a native-shape fallback is
drawn so the build never fails.

Every chart, table and diagram is native PowerPoint (shapes, tables, text
boxes, connectors). Nothing is rasterised, so every number and label stays
editable after import into Google Slides or PowerPoint.

Text markup accepted by every text argument:
    **text**   bold, same colour
    [[text]]   bold Carnegie Red highlight
    "\n"       line break inside a paragraph
"""
import copy, glob, math, os, re, shutil, subprocess
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn

# ───────────────────────────── palette ─────────────────────────────
red         = RGBColor(0xC4, 0x12, 0x30)   # Carnegie Red, primary
rose        = RGBColor(0xEF, 0x3A, 0x47)   # Scots Rose, second series
skibo       = RGBColor(0x94, 0x11, 0x20)   # Skibo Red, third series
iron        = RGBColor(0x6D, 0x6E, 0x71)   # Iron Gray, neutral series / secondary text
steel       = RGBColor(0xE0, 0xE0, 0xE0)   # Steel Gray, grid lines / connectors
light       = RGBColor(0xF2, 0xF2, 0xF2)   # panel fill, zebra rows
tint        = RGBColor(0xFB, 0xEC, 0xEE)   # pale red panel fill
teal        = RGBColor(0x00, 0x8F, 0x91)   # tartan accent (Teal Thread)
blue        = RGBColor(0x04, 0x36, 0x73)   # tartan accent (Blue Thread)
sky         = RGBColor(0x00, 0x7B, 0xC0)   # tartan accent (Sky Blue)
green       = RGBColor(0x00, 0x9F, 0x6B)   # tartan accent (Green Thread)
gold        = RGBColor(0xFD, 0xB5, 0x15)   # tartan accent (Gold Thread)
black       = RGBColor(0x00, 0x00, 0x00)
white       = RGBColor(0xFF, 0xFF, 0xFF)
rule_gray   = RGBColor(0xAE, 0xAE, 0xAF)   # rule under the action title
title_gray  = RGBColor(0xA5, 0xA5, 0xA5)   # rule on the title slide
title_navy  = RGBColor(0x0B, 0x1B, 0x2E)   # title-slide background base
title_panel = RGBColor(0xF2, 0xF2, 0xF2)   # title-slide right panel
divider_bg  = RGBColor(0x5E, 0x0A, 0x18)   # section divider background
closing_bg  = RGBColor(0x3D, 0x07, 0x10)   # closing slide background
closing_sub = RGBColor(0xE8, 0xB4, 0xBC)   # closing tagline
closing_hd  = RGBColor(0xF0, 0x9A, 0xA5)   # closing column headers
venn_right  = RGBColor(0xE8, 0xD8, 0xDA)   # right venn circle fill
font        = "Open Sans"
serif_font  = "Georgia"                    # only used by the text fallback for the CMU wordmark

# ───────────────────────────── geometry (inches) ─────────────────────────────
slide_w, slide_h = 13.333, 7.5
tartan_w   = 0.70     # tartan sidebar width
m_l        = 1.05     # content left edge with tartan sidebar
gradient_w = 1.24     # gradient sidebar width
gm_l       = 1.55     # content left edge with gradient sidebar
content_r  = 12.30    # content right edge (both sidebars)
content_w  = content_r - m_l      # 11.25
gcontent_w = content_r - gm_l     # 10.75
title_y    = 0.40     # action title top
rule_y     = 1.15     # grey rule under the title
sub_y      = 1.22     # italic sub-heading
body_y_sub = 1.62     # first body row when a sub-heading is present
body_y     = 1.35     # first body row with no sub-heading
source_y   = 6.47     # source line top
body_max_src = 6.40   # body must end above this when there is a source line
body_max   = 6.80     # body must end above this with no source line
logo_box   = (12.30, 0.40, 0.84, 0.84)   # Tepper T, top right
wordmark_y, wordmark_w, wordmark_h = 6.90, 2.52, 0.46

_ns_a = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


# ───────────────────────────── assets ─────────────────────────────
asset_rules = {
    # key: (name keywords, expected aspect = width / height, tolerance factor)
    "tartan_strip":     (["tartan", "plaid", "strip", "sidebar"], 0.093, 4.0),
    "t_logo":           (["tepper-t", "tepper_t", "t-logo", "t_logo", "tlogo", "logo-t", "monogram", "tepper-logo", "tepper_logo"], 1.0, 1.6),
    "wordmark":         (["wordmark", "lockup", "cmu-tepper", "cmu_tepper", "signature", "tepper-school", "tepper_school"], 5.5, 2.5),
    "wordmark_white":   (["white", "reverse", "knockout", "stacked", "inverse"], 1.58, 2.5),
    "title_background": (["background", "title-bg", "title_bg", "cover", "lattice", "bg"], 1.78, 1.3),
}


def _img_size(path):
    try:
        from PIL import Image
        with Image.open(path) as im:
            return im.size
    except Exception:
        return None


def _is_mostly_white(path):
    """True for a reversed (white) logo: transparent background, near-white ink."""
    try:
        from PIL import Image
        with Image.open(path) as im:
            im = im.convert("RGBA"); im.thumbnail((200, 200))
            allpx = list(im.getdata())
            px = [p for p in allpx if p[3] > 128]
            if not px or len(px) / len(allpx) > 0.97:   # no transparency: not a reversed logo
                return False
            return sum((p[0] + p[1] + p[2]) / 3 for p in px) / len(px) > 215
    except Exception:
        return False


def _red_share(path):
    """Share of inked (non-white, opaque) pixels that are red, 0..1."""
    try:
        from PIL import Image
        with Image.open(path) as im:
            im = im.convert("RGBA"); im.thumbnail((120, 120))
            ink = [p for p in im.getdata() if p[3] > 128 and min(p[:3]) < 200]
            if not ink: return 0.0
            return sum(1 for p in ink if p[0] > 150 and p[1] < 110 and p[2] < 120) / len(ink)
    except Exception:
        return 0.0


def find_assets(folder, verbose=True):
    """Scan a folder (recursively) for brand images and classify them.
    Returns {key: path or None, 'icons_dir': path or None, 'other_logos': [paths]}.
    Classification uses file-name keywords first, then aspect ratio and colour.
    ALWAYS look at the matched images afterwards and override wrong matches."""
    folder = os.path.expanduser(folder)
    exts = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff")
    files = [p for p in glob.glob(os.path.join(folder, "**", "*"), recursive=True)
             if p.lower().endswith(exts)]
    found = {k: None for k in asset_rules}
    icon_like = [p for p in files
                 if re.match(r"^(" + "|".join(icon_names) + r")-(red|white)\.png$", os.path.basename(p).lower())]
    found["icons_dir"] = os.path.dirname(icon_like[0]) if icon_like else None
    pool = [p for p in files if p not in icon_like]
    taken = set()
    # pass 1: names
    for key, (words, _, _) in asset_rules.items():
        best = None
        for p in pool:
            if p in taken: continue
            name = os.path.basename(p).lower()
            if key == "wordmark" and _is_mostly_white(p): continue
            if key == "wordmark_white" and not (any(w in name for w in words) and
                                                ("wordmark" in name or "cmu" in name or "carnegie" in name)):
                continue
            if any(w in name for w in words):
                best = p; break
        if best: found[key] = best; taken.add(best)
    # pass 2: shape and colour for anything still missing
    for key, (_, aspect, tol) in asset_rules.items():
        if found[key]: continue
        for p in pool:
            if p in taken: continue
            sz = _img_size(p)
            if not sz: continue
            a = sz[0] / sz[1]
            if not (aspect / tol <= a <= aspect * tol): continue
            if key == "wordmark_white" and not _is_mostly_white(p): continue
            if key == "wordmark" and _is_mostly_white(p): continue
            if key == "t_logo" and _red_share(p) < 0.6: continue
            found[key] = p; taken.add(p); break
    found["other_logos"] = [p for p in pool if p not in taken]
    if verbose:
        print("asset scan of", folder)
        for k in asset_rules:
            print(f"  {k:17s} -> {found[k] or 'MISSING (native fallback will be drawn)'}")
        print(f"  {'icons_dir':17s} -> {found['icons_dir'] or 'none (icons are generated with matplotlib)'}")
        print(f"  other_logos       -> {len(found['other_logos'])} file(s)")
        for p in found["other_logos"]: print("      ", p)
    return found


# ───────────────────────────── primitives ─────────────────────────────
_mark = re.compile(r'(\*\*[^*]+\*\*|\[\[[^\]]+\]\])')
_align = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT, "j": PP_ALIGN.JUSTIFY,
          None: None}
_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}


def _al(a):
    return _align.get(a, a)


def _an(a):
    return _anchor.get(a, a)


def runs(p, text, size, color, bold=False, italic=False, highlight=None):
    """Append runs to a paragraph, honouring **bold** and [[red highlight]] markup."""
    for seg in _mark.split(str(text)):
        if not seg: continue
        c, b = color, bold
        if seg.startswith("**") and seg.endswith("**"):
            seg, b = seg[2:-2], True
        elif seg.startswith("[[") and seg.endswith("]]"):
            seg, b, c = seg[2:-2], True, (highlight or red)
        r = p.add_run(); r.text = seg
        f = r.font; f.name = font; f.size = Pt(size)
        f.bold = b; f.italic = italic; f.color.rgb = c
    return p


def tbox(slide, l, t, w, h, anchor="t", wrap=True, inset=0.0, align=None):
    """Empty text box with zero (or given) insets. Returns (shape, text_frame)."""
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap; tf.vertical_anchor = _an(anchor)
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(inset)
    if align: tf.paragraphs[0].alignment = _al(align)
    return tb, tf


def para(tf, text, size=12, color=black, bold=False, first=False, after=5, before=0,
         align="l", italic=False, line=None):
    """Add a paragraph (first=True reuses the empty first paragraph)."""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = _al(align)
    p.space_after = Pt(after); p.space_before = Pt(before)
    if line: p.line_spacing = line
    return runs(p, text, size, color, bold, italic)


def text(slide, l, t, w, h, content, size=12, color=black, bold=False, italic=False,
         align="l", anchor="t", after=5, line=None, wrap=True):
    """One-call text box. content is a string or a list of strings (one paragraph each)."""
    tb, tf = tbox(slide, l, t, w, h, anchor=anchor, wrap=wrap)
    items = content if isinstance(content, (list, tuple)) else [content]
    for i, c in enumerate(items):
        para(tf, c, size, color, bold, first=(i == 0), after=(after if i < len(items) - 1 else 0),
             align=align, italic=italic, line=line)
    return tb


def rect(slide, l, t, w, h, fill=None, line=None, lw=1.0, dash=None, shape=MSO_SHAPE.RECTANGLE):
    sh = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill is None: sh.fill.background()
    else: sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None: sh.line.fill.background()
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(lw)
        if dash: sh.line.dash_style = dash
    sh.shadow.inherit = False
    sh.text_frame.word_wrap = True
    return sh


def oval(slide, l, t, w, h, fill=None, line=None, lw=1.5):
    return rect(slide, l, t, w, h, fill=fill, line=line, lw=lw, shape=MSO_SHAPE.OVAL)


def shape_text(sh, content, size, color, bold=False, align="c", anchor="m", italic=False):
    """Write text inside an autoshape (rectangle, oval, chip)."""
    tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = _an(anchor)
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]; p.alignment = _al(align)
    return runs(p, content, size, color, bold, italic)


def line(slide, x1, y1, x2, y2, color=steel, lw=1.0, dash=None):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color; c.line.width = Pt(lw)
    if dash: c.line.dash_style = dash
    return c


def hline(slide, x1, y, x2, color=steel, lw=1.0, dash=None):
    return line(slide, x1, y, x2, y, color, lw, dash)


def vline(slide, x, y1, y2, color=steel, lw=1.0, dash=None):
    return line(slide, x, y1, x, y2, color, lw, dash)


def picture_fit(slide, path, l, t, w, h, align="c"):
    """Insert an image scaled to fit inside (l, t, w, h) with its aspect ratio kept."""
    pic = slide.shapes.add_picture(path, Inches(l), Inches(t))
    sc = min(Inches(w) / pic.width, Inches(h) / pic.height)
    pw, ph = int(pic.width * sc), int(pic.height * sc)
    pic.width, pic.height = pw, ph
    off_x = {"l": 0, "c": (Inches(w) - pw) / 2, "r": Inches(w) - pw}[align]
    pic.left = int(Inches(l) + off_x); pic.top = int(Inches(t) + (Inches(h) - ph) / 2)
    return pic


def circle_photo(slide, path, l, t, d):
    """Square-crop an image and mask it to a circle."""
    pic = slide.shapes.add_picture(path, Inches(l), Inches(t))
    w, h = pic.width, pic.height
    if w > h:
        c = (1 - h / w) / 2; pic.crop_left = c; pic.crop_right = c
    elif h > w:
        c = (1 - w / h) / 2; pic.crop_top = c; pic.crop_bottom = c
    pic.width = pic.height = Inches(d)
    pic.left, pic.top = Inches(l), Inches(t)
    pic.auto_shape_type = MSO_SHAPE.OVAL
    return pic


# ───────────────────────────── icons ─────────────────────────────
icon_names = ["gear", "check", "doc", "shield", "gauge", "people", "growth", "target", "clock", "scale"]
_icon_cache = os.path.join(os.getcwd(), "tepper_icons")


def make_icons(folder=_icon_cache):
    """Draw the 10 line icons (red and white PNGs) with matplotlib. Returns folder or None."""
    try:
        import matplotlib; matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.patches import Circle, Polygon, Wedge, Rectangle
        import numpy as np
    except Exception:
        return None
    os.makedirs(folder, exist_ok=True)
    lw = 2.6

    def ax_():
        fig, ax = plt.subplots(figsize=(1.4, 1.4), dpi=200)
        ax.set_xlim(0, 10); ax.set_ylim(0, 10); ax.axis("off"); ax.set_aspect("equal")
        fig.patch.set_alpha(0)
        return fig, ax

    def draw(name, c):
        fig, ax = ax_()
        if name == "gear":
            for i in range(8):
                a = 2 * np.pi * i / 8
                ax.plot([5 + 3.5 * np.cos(a), 5 + 4.4 * np.cos(a)], [5 + 3.5 * np.sin(a), 5 + 4.4 * np.sin(a)],
                        color=c, lw=lw, solid_capstyle="round")
            ax.add_patch(Circle((5, 5), 3.5, fill=False, edgecolor=c, lw=lw))
            ax.add_patch(Circle((5, 5), 1.6, fill=False, edgecolor=c, lw=lw))
        elif name == "check":
            ax.add_patch(Circle((5, 5), 4.0, fill=False, edgecolor=c, lw=lw))
            ax.plot([2.9, 4.4, 7.3], [5.0, 3.2, 6.8], color=c, lw=lw, solid_capstyle="round", solid_joinstyle="round")
        elif name == "doc":
            ax.add_patch(Polygon([(2.3, 1.2), (6.0, 1.2), (7.7, 2.9), (7.7, 8.8), (2.3, 8.8)], closed=True,
                                 fill=False, edgecolor=c, lw=lw, joinstyle="round"))
            ax.plot([6.0, 6.0, 7.7], [1.2, 2.9, 2.9], color=c, lw=lw)
            for y in (5.0, 6.1, 7.2): ax.plot([3.4, 6.6], [y, y], color=c, lw=lw * 0.75, solid_capstyle="round")
        elif name == "shield":
            ax.add_patch(Polygon([(5, 9.0), (8.3, 7.4), (8.3, 3.6), (5, 0.7), (1.7, 3.6), (1.7, 7.4)], closed=True,
                                 fill=False, edgecolor=c, lw=lw, joinstyle="round"))
            ax.plot([3.3, 4.5, 6.9], [4.7, 3.1, 6.6], color=c, lw=lw, solid_capstyle="round", solid_joinstyle="round")
        elif name == "gauge":
            ax.add_patch(Wedge((5, 4.3), 4.0, 20, 160, width=0.0, fill=False, edgecolor=c, lw=lw))
            a = np.radians(70)
            ax.plot([5, 5 + 3.1 * np.cos(a)], [4.3, 4.3 + 3.1 * np.sin(a)], color=c, lw=lw, solid_capstyle="round")
            ax.add_patch(Circle((5, 4.3), 0.45, facecolor=c, edgecolor="none"))
        elif name == "people":
            ax.add_patch(Circle((5, 7.1), 1.5, fill=False, edgecolor=c, lw=lw))
            ax.add_patch(Wedge((5, 1.0), 3.4, 20, 160, width=0.0, fill=False, edgecolor=c, lw=lw))
        elif name == "growth":
            for x, h in [(2.4, 2.2), (4.2, 3.6), (6.0, 5.0)]:
                ax.add_patch(Rectangle((x, 1.4), 1.2, h, fill=False, edgecolor=c, lw=lw))
            ax.annotate("", xy=(8.6, 8.2), xytext=(2.6, 5.4),
                        arrowprops=dict(arrowstyle="-|>", color=c, lw=lw, mutation_scale=17))
        elif name == "target":
            for r in (4.0, 2.6, 1.2): ax.add_patch(Circle((5, 5), r, fill=False, edgecolor=c, lw=lw))
            ax.add_patch(Circle((5, 5), 0.4, facecolor=c, edgecolor="none"))
        elif name == "clock":
            ax.add_patch(Circle((5, 5), 4.0, fill=False, edgecolor=c, lw=lw))
            ax.plot([5, 5], [5, 7.6], color=c, lw=lw, solid_capstyle="round")
            ax.plot([5, 7.0], [5, 5], color=c, lw=lw, solid_capstyle="round")
        elif name == "scale":
            ax.plot([5, 5], [1.6, 8.2], color=c, lw=lw, solid_capstyle="round")
            ax.plot([1.6, 8.4], [7.4, 7.4], color=c, lw=lw, solid_capstyle="round")
            for cx in (1.6, 8.4):
                ax.add_patch(Wedge((cx, 7.0), 1.5, 200, 340, width=0.0, fill=False, edgecolor=c, lw=lw))
            ax.plot([3.2, 6.8], [1.6, 1.6], color=c, lw=lw, solid_capstyle="round")
        return fig

    for n in icon_names:
        for cname, hexc in (("red", "#C41230"), ("white", "#FFFFFF")):
            p = os.path.join(folder, f"{n}-{cname}.png")
            if not os.path.exists(p):
                fig = draw(n, hexc); fig.savefig(p, transparent=True); plt.close(fig)
    return folder


def icon(slide, name, l, t, size=0.42, color="red", icons_dir=None):
    """Place a line icon. name: a path to any image, or one of icon_names.
    color: 'red' (on light backgrounds) or 'white' (on red / gray / dark fills)."""
    if os.path.exists(str(name)):
        return slide.shapes.add_picture(name, Inches(l), Inches(t), Inches(size), Inches(size))
    for d in [icons_dir, _icon_cache]:
        if d:
            p = os.path.join(d, f"{name}-{color}.png")
            if os.path.exists(p):
                return slide.shapes.add_picture(p, Inches(l), Inches(t), Inches(size), Inches(size))
    d = make_icons()
    if d:
        return slide.shapes.add_picture(os.path.join(d, f"{name}-{color}.png"),
                                        Inches(l), Inches(t), Inches(size), Inches(size))
    # last resort: native outlined circle with the initial letter
    o = oval(slide, l, t, size, size, line=(red if color == "red" else white), lw=1.5)
    shape_text(o, str(name)[:1].upper(), size * 22, red if color == "red" else white, bold=True)
    return o


# ───────────────────────────── furniture ─────────────────────────────
def gradient_sidebar(slide, width=gradient_w):
    """Diagonal Carnegie Red to Blue Thread strip with white triangle outlines
    (BaseCamp Program Overview deck, slide 4 onwards). Native shapes only."""
    bg = rect(slide, 0, 0, width, slide_h)
    bg.fill.gradient()
    stops = bg.fill.gradient_stops
    stops[0].position = 0.0; stops[0].color.rgb = red
    stops[1].position = 1.0; stops[1].color.rgb = blue
    # fill.gradient_angle is unreliable (setting 105 wrote 255 degrees), and on a
    # tall narrow strip any angle far from vertical collapses the blend into a
    # sliver, so the <a:lin ang> attribute is written directly (60000ths of a degree).
    bg._element.spPr.find(f'.//{_ns_a}lin').set('ang', str(100 * 60000))
    bg.line.fill.background()
    for tx, ty, ts, rot in [(0.10, 0.35, 0.30, -8), (0.30, 1.55, 0.22, 14), (0.06, 2.55, 0.34, -4),
                            (0.28, 3.55, 0.24, 10), (0.08, 4.55, 0.30, -10), (0.30, 5.55, 0.22, 16),
                            (0.10, 6.55, 0.28, -6)]:
        t = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(tx), Inches(ty), Inches(ts), Inches(ts))
        t.rotation = rot; t.fill.background()
        t.line.color.rgb = white; t.line.width = Pt(1.0); t.shadow.inherit = False


def native_tartan(slide, width=tartan_w):
    """Fallback tartan strip drawn with rectangles (used only if no tartan image)."""
    rect(slide, 0, 0, width, slide_h, fill=blue)
    rect(slide, width * 0.60, 0, width * 0.40, slide_h, fill=RGBColor(0x00, 0x6B, 0x5A))
    y = 0.0
    pattern = [(0.95, None), (0.10, red), (0.55, None), (0.05, gold), (0.40, None), (0.12, red),
               (0.30, RGBColor(0x0A, 0x2A, 0x4A))]
    while y < slide_h:
        for hgt, col in pattern:
            if col is not None and y < slide_h:
                rect(slide, 0, y, width, min(hgt, slide_h - y), fill=col)
            y += hgt
    vline(slide, width * 0.27, 0, slide_h, color=red, lw=1.2, dash=MSO_LINE_DASH_STYLE.DASH)


def native_wordmark(slide, l, t, color_top=red, color_bottom=black):
    _, tf = tbox(slide, l, t - 0.04, 3.2, 0.52)
    p = tf.paragraphs[0]; p.space_after = Pt(0)
    r = p.add_run(); r.text = "Carnegie Mellon University"
    r.font.name = serif_font; r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = color_top
    p2 = tf.add_paragraph()
    r2 = p2.add_run(); r2.text = "Tepper School of Business"
    r2.font.name = font; r2.font.size = Pt(14); r2.font.color.rgb = color_bottom


def slide_number(slide, n, l=12.30, t=6.95, w=0.80, h=0.30, color=iron):
    """Live slide-number field (updates if slides are reordered)."""
    _, tf = tbox(slide, l, t, w, h, anchor="m")
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    fld = etree.SubElement(p._p, qn('a:fld'))
    fld.set('id', '{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}'); fld.set('type', 'slidenum')
    rpr = etree.SubElement(fld, qn('a:rPr')); rpr.set('lang', 'en-US'); rpr.set('sz', '1100')
    sf = etree.SubElement(rpr, qn('a:solidFill')); clr = etree.SubElement(sf, qn('a:srgbClr'))
    clr.set('val', str(color))
    lat = etree.SubElement(rpr, qn('a:latin')); lat.set('typeface', font)
    t_ = etree.SubElement(fld, qn('a:t')); t_.text = str(n)


# ───────────────────────────── deck ─────────────────────────────
class Deck:
    """d = Deck(asset_dir="/path/to/logos")   or   Deck(assets={...})
    sidebar: default sidebar for content slides, 'tartan' or 'gradient'."""

    def __init__(self, asset_dir=None, assets=None, sidebar="tartan"):
        self.prs = Presentation()
        self.prs.slide_width = Inches(slide_w); self.prs.slide_height = Inches(slide_h)
        self.blank = self.prs.slide_layouts[6]
        self.assets = {k: None for k in asset_rules}; self.assets["icons_dir"] = None
        self.assets["other_logos"] = []
        if asset_dir: self.assets.update(find_assets(asset_dir))
        if assets: self.assets.update(assets)
        self.sidebar = sidebar
        self.meta = {}          # slide_id -> layout info used by lint()
        self._set_theme_font()

    # theme fonts -> Open Sans, so any text without an explicit font still matches
    def _set_theme_font(self):
        from pptx.opc.constants import RELATIONSHIP_TYPE as rt
        theme = self.prs.slide_master.part.part_related_by(rt.THEME)
        root = etree.fromstring(theme.blob)
        for tag in ("majorFont", "minorFont"):
            for el in root.iter(f"{_ns_a}{tag}"):
                el.find(f"{_ns_a}latin").set("typeface", font)
        theme._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def _a(self, key):
        p = self.assets.get(key)
        return p if p and os.path.exists(p) else None

    def _new(self, kind, **info):
        s = self.prs.slides.add_slide(self.blank)
        self.meta[s.slide_id] = dict(kind=kind, **info)
        return s

    def _mark_furniture(self, s):
        self.meta[s.slide_id]["n_furniture"] = len(s.shapes)

    def _sidebar(self, s, kind):
        if kind == "gradient":
            gradient_sidebar(s); return gm_l
        p = self._a("tartan_strip")
        if p: s.shapes.add_picture(p, 0, 0, Inches(tartan_w), Inches(slide_h))
        else: native_tartan(s)
        return m_l

    # ── title slide ──
    def title_slide(self, title, org, names, partner_logos=None):
        """Cover: lattice background left, grey panel right. org = event / course line
        (bold), names = presenters line. partner_logos = image paths shown bottom-right."""
        s = self._new("title")
        bgp = self._a("title_background")
        if bgp: s.shapes.add_picture(bgp, 0, 0, Inches(slide_w), Inches(slide_h))
        else:
            rect(s, 0, 0, slide_w, slide_h, fill=title_navy)
            cols = [red, green, teal, gold, sky, rose]
            for i in range(26):
                c = cols[i % len(cols)]
                x0 = -2.0 + i * 0.32
                line(s, x0, 0, x0 + 3.2, slide_h, color=c, lw=0.9)
                line(s, x0 + 3.0, 0, x0 - 1.0, slide_h, color=c, lw=0.9)
            rect(s, 4.2, 0, 1.30, slide_h, fill=title_navy)
        rect(s, 5.44, 0, slide_w - 5.44, slide_h, fill=title_panel)
        wm = self._a("wordmark_white")
        if wm: picture_fit(s, wm, 2.85, 2.80, 2.27, 1.44, align="l")
        else:
            _, tf = tbox(s, 2.85, 2.70, 2.5, 1.6, anchor="m")
            for i, w_ in enumerate(["Carnegie", "Mellon", "University"]):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                r = p.add_run(); r.text = w_
                r.font.name = serif_font; r.font.bold = True; r.font.size = Pt(30); r.font.color.rgb = white
        text(s, 6.32, 2.50, 6.20, 1.31, title, 32 if len(title) <= 34 else 26, iron, anchor="m")
        hline(s, 6.32, 4.10, 12.54, color=title_gray, lw=1.0)
        text(s, 6.38, 4.40, 6.10, 0.37, org, 15, iron, bold=True, anchor="m")
        text(s, 6.38, 4.87, 6.40, 0.71, names, 13, iron)
        if partner_logos:
            x = 6.38
            for p in partner_logos:
                pic = picture_fit(s, p, x, 5.95, 1.6, 0.55, align="l")
                x = Emu(pic.left + pic.width).inches + 0.35
        slide_number(s, len(self.prs.slides), l=12.48, t=6.93)
        self._mark_furniture(s)
        return s

    # ── content slide ──
    def content(self, title, source=None, sub=None, sidebar=None):
        """Standard slide: sidebar, Tepper T, action title, grey rule, optional italic
        sub-heading, optional source line, CMU wordmark, slide number.
        Returns the slide. Body coordinates: see deck.left(s), deck.width(s), deck.top(s), deck.bottom(s)."""
        kind = sidebar or self.sidebar
        s = self._new("content", sidebar=kind, has_sub=bool(sub), has_source=bool(source))
        left = self._sidebar(s, kind)
        self.meta[s.slide_id]["left"] = left
        tp = self._a("t_logo")
        if tp: picture_fit(s, tp, *logo_box)
        wp = self._a("wordmark")
        if wp: picture_fit(s, wp, left, wordmark_y, wordmark_w, wordmark_h, align="l")
        else: native_wordmark(s, left, wordmark_y)
        n = len(title)
        size = 24 if n <= 56 else (21 if n <= 66 else 19)
        hb, htf = tbox(s, left, title_y, 13.20 - left, 0.65, wrap=False, inset=0.1)
        para(htf, title, size, red, bold=True, first=True, after=0)
        hline(s, left - 0.05, rule_y, 13.10, color=rule_gray, lw=0.75)
        if sub:
            text(s, left, sub_y, content_r - left, 0.30, sub, 12, iron, italic=True)
        if source:
            sb, stf = tbox(s, left, source_y, content_r - left, 0.37, wrap=False, inset=0.1)
            para(stf, source, 9, iron, italic=True, first=True, after=0)
        slide_number(s, len(self.prs.slides))
        self._mark_furniture(s)
        return s

    # body-geometry helpers for a content slide
    def left(self, s):   return self.meta[s.slide_id]["left"]
    def width(self, s):  return content_r - self.left(s)
    def top(self, s):    return body_y_sub if self.meta[s.slide_id]["has_sub"] else body_y
    def bottom(self, s): return body_max_src if self.meta[s.slide_id]["has_source"] else body_max

    # ── section divider ──
    def divider(self, label, sidebar=None):
        s = self._new("divider")
        rect(s, 0, 0, slide_w, slide_h, fill=divider_bg)
        if (sidebar or self.sidebar) == "gradient": gradient_sidebar(s)
        else:
            p = self._a("tartan_strip")
            if p: s.shapes.add_picture(p, 0, 0, Inches(tartan_w), Inches(slide_h))
            else: native_tartan(s)
        text(s, 1.55, 3.05, 10.5, 0.85, label, 34, white, bold=True, italic=True)
        hline(s, 1.60, 4.06, 6.20, color=rose, lw=2.0)
        self._mark_furniture(s)
        return s

    # ── closing slide ──
    def closing(self, headline, columns, tagline=None):
        """Dark closing / contact slide. columns = [(header, [lines])], up to 4."""
        s = self._new("closing")
        rect(s, 0, 0, slide_w, slide_h, fill=closing_bg)
        self._mark_furniture(s)
        text(s, 0.85, 0.75, 11.0, 0.85, headline, 32, white, bold=True)
        if tagline: text(s, 0.85, 1.42, 11.0, 0.35, tagline, 13, closing_sub, italic=True)
        hline(s, 0.90, 1.92 if tagline else 1.62, 3.30, color=rose, lw=2.0)
        n = len(columns); cw = 11.6 / n
        for i, (head, lines) in enumerate(columns):
            cx = 0.85 + i * cw
            text(s, cx, 5.85, cw - 0.25, 0.28, head, 10.5, closing_hd, bold=True)
            text(s, cx, 6.16, cw - 0.25, 0.95, lines, 10, white, after=3)
        return s

    def save(self, path):
        self.prs.save(path)
        return path


# ───────────────────────────── components ─────────────────────────────
def bullets(slide, x, y, w, h, items, size=11, color=black, gap=7, marker="•", marker_color=None):
    """Bulleted list in one text box. items may use ** and [[ ]] markup."""
    _, tf = tbox(slide, x, y, w, h)
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        runs(p, marker + "   ", size, marker_color or color)
        runs(p, it, size, color)
    return tf


def rich_paragraphs(slide, x, y, w, h, paras, size=12.5, gap=11, line=1.25):
    """Executive-summary prose. Use [[...]] to highlight the key claim of each paragraph."""
    _, tf = tbox(slide, x, y, w, h)
    for i, ptxt in enumerate(paras):
        para(tf, ptxt, size, black, first=(i == 0), after=gap, line=line)
    return tf


def section_label(slide, x, y, w, label, color=iron, size=12):
    """Small bold grey label above a block ('Colour Hierarchy', 'What The Data Shows')."""
    return text(slide, x, y, w, 0.32, label, size, color, bold=True)


def kpi_banner(slide, x, y, w, h, items, gap=0.22, num_size=34):
    """Row of big-number boxes on tint. items = [(number, caption)], 2 to 4 items."""
    n = len(items); bw = (w - gap * (n - 1)) / n
    for i, (num, cap) in enumerate(items):
        bx = x + i * (bw + gap)
        rect(slide, bx, y, bw, h, fill=tint)
        nh = num_size * 1.22 / 72
        text(slide, bx + 0.22, y + 0.08, bw - 0.44, nh, num, num_size, red, bold=True)
        text(slide, bx + 0.22, y + 0.10 + nh, bw - 0.44, h - nh - 0.16, cap, 10.5, black)


def metric_grid(slide, x, y, w, h, cells, cols=2):
    """Label / big value / italic caption / source, each cell closed by a red rule.
    cells = [(label, value, caption, source)]"""
    n_r = (len(cells) + cols - 1) // cols
    cw = w / cols; ch = h / n_r
    for i, (label, value, cap, src) in enumerate(cells):
        r, c = divmod(i, cols)
        cx = x + c * cw; cy = y + r * ch
        text(slide, cx, cy + 0.04, cw - 0.18, 0.22, label, 11, black, bold=True)
        text(slide, cx, cy + 0.28, cw - 0.18, 0.42, value, 22, red, bold=True)
        text(slide, cx, cy + 0.72, cw - 0.18, 0.22, cap, 9.5, black, italic=True)
        text(slide, cx, cy + ch - 0.24, cw - 0.18, 0.20, src, 8.5, iron)
        hline(slide, cx, cy + ch - 0.03, cx + cw - 0.20, color=red, lw=1.2)


def bar_chart(slide, x, y, w, h, cats, vals, labels=None, colors=None, title=None,
              baseline=True, cat_size=10, lbl_size=10.5, lbl_inside=False, gap_ratio=0.42, headroom=1.02):
    """Native vertical bar chart. Returns the baseline y. headroom > 1 reserves space
    above the tallest bar (needed when line_overlay bubbles sit above the bars)."""
    if title:
        text(slide, x, y, w, 0.34, title, 11.5, black, bold=True, align="c")
        y += 0.40; h -= 0.40
    cat_h, lbl_h = 0.30, 0.30
    plot_h = h - cat_h - lbl_h
    n = len(vals); slot = w / n; bw = slot * (1 - gap_ratio)
    vmax = max(vals) * headroom
    base_y = y + lbl_h + plot_h
    for i, v in enumerate(vals):
        bh = plot_h * (v / vmax) if vmax else 0
        bx = x + i * slot + (slot - bw) / 2; by = base_y - bh
        rect(slide, bx, by, bw, max(bh, 0.01), fill=(colors[i] if colors else red))
        lab = labels[i] if labels else str(v)
        if lbl_inside and bh > 0.5:
            text(slide, bx, by + 0.06, bw, 0.42, lab, lbl_size, white, bold=True, align="c")
        else:
            text(slide, bx - 0.12, by - lbl_h, bw + 0.24, lbl_h, lab, lbl_size, black, bold=True,
                 align="c", anchor="b")
        text(slide, bx - 0.16, base_y + 0.05, bw + 0.32, cat_h, cats[i], cat_size, black, align="c")
    if baseline: hline(slide, x, base_y, x + w, color=iron, lw=1.0)
    return base_y


def stacked_bar_chart(slide, x, y, w, h, cats, series, series_names, colors, title=None,
                      totals=True, legend=True, seg_labels=True):
    """Native stacked vertical bars. series = [[v_cat0, v_cat1, ...] per series]."""
    if title:
        text(slide, x, y, w, 0.34, title, 11.5, black, bold=True, align="c")
        y += 0.40; h -= 0.40
    if legend:
        lx = x
        for nm, col in zip(series_names, colors):
            rect(slide, lx, y + 0.06, 0.16, 0.16, fill=col)
            adv = 0.32 + min(1.9, 0.075 * len(nm) + 0.2)
            text(slide, lx + 0.22, y + 0.02, adv - 0.26, 0.24, nm, 9, black)
            lx += adv
        y += 0.34; h -= 0.34
    cat_h, lbl_h = 0.30, 0.30
    plot_h = h - cat_h - lbl_h
    n = len(cats); slot = w / n; bw = slot * 0.58
    tot = [sum(s[i] for s in series) for i in range(n)]
    vmax = max(tot) * 1.02
    base_y = y + lbl_h + plot_h
    for i in range(n):
        bx = x + i * slot + (slot - bw) / 2; cur = base_y
        for s_vals, col in zip(series, colors):
            bh = plot_h * s_vals[i] / vmax
            rect(slide, bx, cur - bh, bw, bh, fill=col)
            if seg_labels and bh > 0.28:
                text(slide, bx, cur - bh, bw, bh, f"{s_vals[i]:g}", 9, white, bold=True, align="c", anchor="m")
            cur -= bh
        if totals:
            text(slide, bx - 0.12, cur - lbl_h, bw + 0.24, lbl_h, f"{tot[i]:g}", 10.5, black, bold=True,
                 align="c", anchor="b")
        text(slide, bx - 0.16, base_y + 0.05, bw + 0.32, cat_h, cats[i], 10, black, align="c")
    hline(slide, x, base_y, x + w, color=iron, lw=1.0)
    return base_y


def hbar_chart(slide, x, y, w, rows, bar_h=0.34, gap=0.16, label_w=2.4, colors=None, fmt=None, title=None):
    """Native horizontal bar chart / ranking. rows = [(label, value)]; longest bar = full width."""
    if title:
        text(slide, x, y, w, 0.34, title, 11.5, black, bold=True)
        y += 0.42
    vmax = max(v for _, v in rows)
    plot_w = w - label_w - 0.85
    for i, (lab, v) in enumerate(rows):
        ry = y + i * (bar_h + gap)
        text(slide, x, ry, label_w - 0.12, bar_h, lab, 10, black, align="r", anchor="m")
        bw = plot_w * v / vmax
        rect(slide, x + label_w, ry, max(bw, 0.02), bar_h, fill=(colors[i] if colors else red))
        text(slide, x + label_w + bw + 0.08, ry, 0.8, bar_h, fmt(v) if fmt else f"{v:g}", 10, black,
             bold=True, anchor="m")
    vline(slide, x + label_w, y - 0.05, y + len(rows) * (bar_h + gap) - gap + 0.05, color=iron, lw=1.0)


def line_overlay(slide, x, y_top, y_bot, w, n, fracs, labels, color=iron, bubble=True):
    """Rate line with bubbles in the band y_top..y_bot, aligned to n bar slots of width w.
    Reserve the band with bar_chart(headroom=1.8 to 2.0)."""
    slot = w / n; band = y_bot - y_top
    pts = [(x + i * slot + slot / 2, y_bot - band * f) for i, f in enumerate(fracs)]
    for a, b in zip(pts, pts[1:]):
        line(slide, a[0], a[1], b[0], b[1], color=color, lw=1.4)
    for (cx, cy), lab in zip(pts, labels):
        if bubble:
            d = 0.56
            o = oval(slide, cx - d / 2, cy - d / 2, d, d, fill=tint, line=red, lw=1.0)
            o.text_frame.word_wrap = False
            shape_text(o, lab, 9, red, bold=True)
            o.text_frame.margin_left = o.text_frame.margin_right = 0
        else:
            text(slide, cx - 0.30, cy - 0.22, 0.60, 0.24, lab, 9, color, bold=True, align="c")


def hbar_flow(slide, x, y, w, segs, h=0.62, label_size=11):
    """Proportional horizontal split bar. segs = [(fraction, big_label, small_label, colour)].
    Returns [(x_start, x_end)] per segment for drawing connectors to a second level."""
    cx = x; out = []
    for frac, big, small, col in segs:
        sw = w * frac
        rect(slide, cx, y, sw, h, fill=col)
        _, tf = tbox(slide, cx + 0.10, y + 0.06, sw - 0.20, h - 0.12, anchor="m")
        para(tf, big, label_size + 2, white, bold=True, first=True, after=0, align="r")
        para(tf, small, 8.5, white, after=0, align="r")
        out.append((cx, cx + sw)); cx += sw
    return out


def table(slide, x, y, w, h, data, col_w, row_h=None, fsize=10.5, header_fill=red,
          zebra=True, first_col_fill=None, align=None):
    """Native table. data[0] is the header row. col_w must sum to w.
    align: optional list per column of 'l'/'c'/'r' (default: first column centred, rest left)."""
    rows, cols = len(data), len(data[0])
    gt = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h)).table
    gt.first_row = True; gt.horz_banding = False
    for i, cw in enumerate(col_w): gt.columns[i].width = Inches(cw)
    if row_h:
        for i, rh in enumerate(row_h): gt.rows[i].height = Inches(rh)
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            c = gt.cell(ri, ci); c.text = ""
            c.margin_left = Inches(0.09); c.margin_right = Inches(0.07)
            c.margin_top = Inches(0.04); c.margin_bottom = Inches(0.04)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.text_frame.word_wrap = True
            p = c.text_frame.paragraphs[0]
            if ri == 0: p.alignment = PP_ALIGN.CENTER
            elif align: p.alignment = _al(align[ci])
            else: p.alignment = PP_ALIGN.CENTER if ci == 0 else PP_ALIGN.LEFT
            runs(p, val, fsize, white if ri == 0 else black, bold=(ri == 0))
            c.fill.solid()
            if ri == 0: c.fill.fore_color.rgb = header_fill
            elif ci == 0 and first_col_fill is not None: c.fill.fore_color.rgb = first_col_fill
            elif zebra and ri % 2 == 0: c.fill.fore_color.rgb = light
            else: c.fill.fore_color.rgb = white
    return gt


def grouped_table(slide, x, y, w, h, groups, col_w, headers=("Category", "Sub-Criteria", "Guiding Question", "Weight"),
                  fsize=10):
    """Category and weight cells merged across their sub-rows.
    groups = [(category, weight, [(sub_item, detail), ...])]"""
    total_rows = 1 + sum(len(g[2]) for g in groups)
    gt = slide.shapes.add_table(total_rows, 4, Inches(x), Inches(y), Inches(w), Inches(h)).table
    gt.first_row = True; gt.horz_banding = False
    for i, cw in enumerate(col_w): gt.columns[i].width = Inches(cw)
    for ci, ht in enumerate(headers):
        c = gt.cell(0, ci); c.text = ""; c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.margin_left = Inches(0.08); c.margin_top = Inches(0.03); c.margin_bottom = Inches(0.03)
        p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        runs(p, ht, fsize + 0.5, white, bold=True)
        c.fill.solid(); c.fill.fore_color.rgb = red
    ri = 1
    for gi, (cat, weight, subs) in enumerate(groups):
        first_r = ri
        for si, (sub, q) in enumerate(subs):
            for ci, val in enumerate([cat if si == 0 else "", sub, q, weight if si == 0 else ""]):
                c = gt.cell(ri, ci); c.text = ""
                c.margin_left = Inches(0.08); c.margin_right = Inches(0.06)
                c.margin_top = Inches(0.02); c.margin_bottom = Inches(0.02)
                c.vertical_anchor = MSO_ANCHOR.MIDDLE; c.text_frame.word_wrap = True
                p = c.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER if ci in (0, 3) else PP_ALIGN.LEFT
                runs(p, val, fsize, black, bold=ci in (0, 3))
                c.fill.solid(); c.fill.fore_color.rgb = tint if gi % 2 == 0 else white
            ri += 1
        if len(subs) > 1:
            gt.cell(first_r, 0).merge(gt.cell(ri - 1, 0))
            gt.cell(first_r, 3).merge(gt.cell(ri - 1, 3))
    return gt


def stat_card(slide, x, y, w, h, value, caption, rail=red):
    """Tinted panel, coloured left rail, big value, bold italic caption."""
    rect(slide, x, y, w, h, fill=tint)
    rect(slide, x, y, 0.055, h, fill=rail)
    text(slide, x + 0.26, y + 0.14, w - 0.46, h * 0.42, value, 30, red, bold=True)
    text(slide, x + 0.26, y + h * 0.58, w - 0.46, h * 0.36, caption, 10, black, bold=True, italic=True)


def header_card(slide, x, y, w, items, header, icon_name=None, fill=red, header_h=0.56, body_h=1.30):
    """Coloured header bar (optional white icon) over a bulleted body."""
    rect(slide, x, y, w, header_h, fill=fill)
    tx = x + 0.20
    if icon_name:
        icon(slide, icon_name, x + 0.16, y + (header_h - 0.30) / 2, 0.30, color="white"); tx = x + 0.60
    text(slide, tx, y + 0.06, w - (tx - x) - 0.15, header_h - 0.12, header, 13, white, bold=True, anchor="m")
    bullets(slide, x + 0.32, y + header_h + 0.16, w - 0.55, body_h, items, size=10.5)


def option_cards(slide, x, y, w, cards, body_h=1.95, gap=0.20, header_h=0.56, size=10.5):
    """Parallel comparison cards. cards = [(icon_name or None, header, fill, [bullets])], 2 to 4."""
    n = len(cards); cw = (w - gap * (n - 1)) / n
    for i, (ic, head, col, items) in enumerate(cards):
        cx = x + i * (cw + gap)
        rect(slide, cx, y, cw, header_h, fill=col)
        tx = cx + 0.20
        if ic:
            icon(slide, ic, cx + 0.16, y + (header_h - 0.30) / 2, 0.30, color="white"); tx = cx + 0.58
        text(slide, tx, y + 0.06, cw - (tx - cx) - 0.14, header_h - 0.12, head, 13, white, bold=True, anchor="m")
        rect(slide, cx, y + header_h, cw, body_h, fill=light)
        bullets(slide, cx + 0.22, y + header_h + 0.18, cw - 0.44, body_h - 0.30, items, size=size, gap=10)


def arrow_list(slide, x, y, w, items, size=11, gap=0.34, color=red):
    """Arrow-prefixed list, one text box per row (fixed rhythm)."""
    for i, it in enumerate(items):
        yy = y + i * gap
        text(slide, x, yy, 0.34, 0.26, "→", size + 1, color, bold=True)
        text(slide, x + 0.36, yy, w - 0.36, 0.28, it, size, black)


def callout(slide, x, y, w, h, lines, dashed=True, fill=None, size=10.5, italic=True):
    """Dashed red border insight box."""
    sh = rect(slide, x, y, w, h, fill=fill, line=red, lw=1.0,
              dash=MSO_LINE_DASH_STYLE.DASH if dashed else None)
    _, tf = tbox(slide, x + 0.18, y + 0.12, w - 0.36, h - 0.24, anchor="m")
    for i, ln in enumerate(lines):
        para(tf, ln, size, black, first=(i == 0), after=(6 if i < len(lines) - 1 else 0), italic=italic)
    return sh


def summary_bar(slide, x, y, w, content, h=0.52, fill=red, size=12):
    """Full-width conclusion bar that closes a slide. fill = red or iron (alternate)."""
    rect(slide, x, y, w, h, fill=fill)
    text(slide, x + 0.25, y + 0.05, w - 0.50, h - 0.10, content, size, white, bold=True, align="c", anchor="m")


def numbered_row(slide, x, y, w, idx, title, body, d=0.62, fill=red):
    """Numbered circle on a tinted row; stagger alternate rows by +0.35 in x."""
    rect(slide, x + d * 0.5, y, w - d * 0.5, d, fill=tint)
    o = oval(slide, x, y, d, d, fill=fill)
    shape_text(o, str(idx), 17, white, bold=True)
    _, tf = tbox(slide, x + d + 0.22, y + 0.04, w - d - 0.45, d - 0.08, anchor="m")
    para(tf, title, 11.5, red, bold=True, first=True, after=1)
    para(tf, body, 10, black, after=0)


def toc_row(slide, x, y, w, num, label, page, chip=0.30):
    """Agenda row: red number chip, label, page number at the right."""
    r = rect(slide, x, y, chip, chip, fill=red)
    shape_text(r, str(num), 11, white, bold=True)
    text(slide, x + chip + 0.22, y + 0.02, w - chip - 0.90, chip, label, 11.5, black)
    text(slide, x + w - 0.55, y + 0.02, 0.55, chip, str(page), 11.5, iron, align="r")


def agenda(slide, x, y, w, items, row_gap=0.62, split=5):
    """Two-column agenda. items = [(label, page)], numbered automatically; up to 2 x split rows."""
    col_w = (w - 0.70) / 2
    for i, (lab, pg) in enumerate(items):
        c, r = divmod(i, split)
        toc_row(slide, x + c * (col_w + 0.70), y + r * row_gap, col_w, i + 1, lab, pg)
    if len(items) > split:
        vline(slide, x + col_w + 0.35, y - 0.05, y + split * row_gap - 0.20, color=steel, lw=1.0)


def hub_diagram(slide, x, y, w, center_label, nodes, hub_d=2.35, node_w=3.35, h=4.0):
    """Centre circle with four labelled boxes (TL, TR, BL, BR) and spokes.
    nodes = [(header, body)] x 4. Occupies x..x+w, y..y+h."""
    cx = x + w / 2; cy = y + h / 2
    hub = oval(slide, cx - hub_d / 2, cy - hub_d / 2, hub_d, hub_d, fill=red)
    shape_text(hub, center_label, 13.5, white, bold=True)
    pos = [(x + 0.35, y), (x + w - node_w - 0.35, y),
           (x + 0.35, y + h - 1.08), (x + w - node_w - 0.35, y + h - 1.08)]
    for (nx, ny), (head, body) in zip(pos, nodes):
        left_side = nx < cx; top_side = ny < cy
        sx = cx + (-1 if left_side else 1) * hub_d / 2 * 0.72
        sy = cy + (-1 if top_side else 1) * hub_d / 2 * 0.72
        ex = nx + node_w if left_side else nx
        line(slide, sx, sy, ex, ny + 0.54, color=steel, lw=1.4)
        r = rect(slide, nx, ny, node_w, 0.46, fill=iron)
        shape_text(r, head, 11.5, white, bold=True, align="l")
        r.text_frame.margin_left = Inches(0.14)
        b = rect(slide, nx, ny + 0.46, node_w, 0.62, fill=white, line=red, lw=1.0, dash=MSO_LINE_DASH_STYLE.DASH)
        shape_text(b, body, 10, black, align="l")
        b.text_frame.margin_left = Inches(0.12)


def venn_diagram(slide, cx, cy, outer_d, inner_d, gap, left_label, right_label, center_label,
                 top_label=None, bottom_note=None, left_items=None, right_items=None):
    """Two overlapping circles inside a framing ring, with side annotations."""
    oval(slide, cx - outer_d / 2, cy - outer_d / 2, outer_d, outer_d, fill=light, line=red, lw=1.2)
    if top_label:
        text(slide, cx - outer_d / 2, cy - outer_d / 2 + 0.16, outer_d, 0.30, top_label, 12, red, bold=True, align="c")
    lx = cx - gap / 2 - inner_d / 2 + gap / 2 * 0.35
    rx = cx + gap / 2 - inner_d / 2 - gap / 2 * 0.35
    oval(slide, lx, cy - inner_d / 2, inner_d, inner_d, fill=tint, line=red, lw=1.0)
    oval(slide, rx, cy - inner_d / 2, inner_d, inner_d, fill=venn_right, line=skibo, lw=1.0)
    overlap_l, overlap_r = rx, lx + inner_d
    text(slide, lx + 0.08, cy - 0.32, overlap_l - lx - 0.12, 0.64, left_label, 10.5, red, bold=True,
         align="c", anchor="m")
    text(slide, overlap_r + 0.04, cy - 0.32, rx + inner_d - overlap_r - 0.12, 0.64, right_label, 10.5, skibo,
         bold=True, align="c", anchor="m")
    text(slide, overlap_l, cy - 0.24, overlap_r - overlap_l, 0.48, center_label, 9, black, bold=True,
         align="c", anchor="m")
    if left_items:
        for i, it in enumerate(left_items):
            text(slide, cx - outer_d / 2 - 3.15, cy - 0.62 + i * 0.42, 2.85, 0.38, it, 9.5, black,
                 italic=True, align="r")
        hline(slide, cx - outer_d / 2 - 0.25, cy, cx - outer_d / 2 - 0.02, color=red, lw=1.0)
    if right_items:
        for i, it in enumerate(right_items):
            text(slide, cx + outer_d / 2 + 0.30, cy - 0.62 + i * 0.42, 2.85, 0.38, it, 9.5, black, italic=True)
        hline(slide, cx + outer_d / 2 + 0.02, cy, cx + outer_d / 2 + 0.25, color=red, lw=1.0)
    if bottom_note:
        text(slide, cx - outer_d / 2 - 1.0, cy + outer_d / 2 + 0.08, outer_d + 2.0, 0.42, bottom_note, 9.5,
             black, italic=True, align="c")


def headshot_grid(slide, x, y, w, people, cols=4, d=1.15):
    """Team grid. people = [(photo_path_or_initials, name, role, [bullets])]."""
    cw = w / cols
    for i, (pic, name, role, items) in enumerate(people):
        c = i % cols; r = i // cols
        px = x + c * cw; py = y + r * (d + 2.3)
        cx = px + cw / 2
        if os.path.exists(str(pic)):
            circle_photo(slide, pic, cx - d / 2, py, d)
            oval(slide, cx - d / 2, py, d, d, line=red, lw=1.5)
        else:
            o = oval(slide, cx - d / 2, py, d, d, fill=light, line=red, lw=1.5)
            shape_text(o, pic, 20, red, bold=True)
        text(slide, px + 0.06, py + d + 0.12, cw - 0.12, 0.30, name, 12, black, bold=True)
        text(slide, px + 0.06, py + d + 0.42, cw - 0.12, 0.24, role, 9.5, red, italic=True)
        hline(slide, px + 0.06, py + d + 0.70, px + cw - 0.12, color=red, lw=1.0)
        bullets(slide, px + 0.06, py + d + 0.80, cw - 0.18, 1.3, items, size=8.7, gap=5)


def chevron_flow(slide, x, y, w, steps, side_notes=None, box_h=0.52, gap=0.30, note_w=4.55):
    """Vertical step boxes joined by chevrons, optional dashed side notes aligned to each step.
    steps = [(label, sub_caption or None, fill)]"""
    cy = y; tops = []
    for i, (label, sub, fill) in enumerate(steps):
        tops.append(cy)
        r = rect(slide, x, cy, w, box_h, fill=fill)
        shape_text(r, label, 11.5, white, bold=True)
        cy += box_h
        if sub:
            text(slide, x + 0.10, cy + 0.04, w - 0.20, 0.30, sub, 9.5, black, italic=True, align="c")
            cy += 0.40
        if i < len(steps) - 1:
            ch = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x + w / 2 - 0.20), Inches(cy),
                                        Inches(0.40), Inches(0.20))
            ch.rotation = 90; ch.fill.solid(); ch.fill.fore_color.rgb = steel
            ch.line.fill.background(); ch.shadow.inherit = False
            cy += gap
    if side_notes:
        nx = x + w + 0.55
        for t0, note in zip(tops, side_notes):
            ny = t0 + box_h / 2 - 0.31
            b = rect(slide, nx, ny, note_w, 0.62, line=red, lw=1.0, dash=MSO_LINE_DASH_STYLE.DASH)
            shape_text(b, note, 10, black, align="l")
            b.text_frame.margin_left = Inches(0.16)
            hline(slide, x + w + 0.05, t0 + box_h / 2, nx, color=iron, lw=1.0, dash=MSO_LINE_DASH_STYLE.SQUARE_DOT)
    return cy


def process_arrows(slide, x, y, w, steps, h=0.80, fills=None, desc=None, desc_h=1.2):
    """Horizontal chevron process (pentagon then chevrons) with optional description under each.
    steps = [label]; desc = [text] same length."""
    n = len(steps); step = w / n
    for i, lab in enumerate(steps):
        sx = x + i * step
        sw = step + (0.18 if i < n - 1 else 0.06)
        shp = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        col = (fills[i] if fills else (red if i % 2 == 0 else skibo))
        r = rect(slide, sx, y, sw - 0.06, h, fill=col, shape=shp)
        shape_text(r, lab, 11.5, white, bold=True)
        r.text_frame.margin_left = Inches(0.30 if i else 0.12); r.text_frame.margin_right = Inches(0.25)
        if desc:
            text(slide, sx + 0.12, y + h + 0.15, w / n - 0.30, desc_h, desc[i], 10, black)


def thesis_block(slide, x, y, w, word, tagline, body, panel_header, panel_items, panel_h=1.85):
    """Big single-word thesis on the left, tinted supporting panel on the right."""
    text(slide, x, y, w * 0.46, 0.75, word, 40, red, bold=True)
    text(slide, x, y + 0.74, w * 0.46, 0.30, tagline, 11, black, italic=True)
    text(slide, x, y + 1.12, w * 0.46, max(panel_h - 1.12, 0.6), body, 11, black, line=1.3)
    px = x + w * 0.50; pw = w * 0.50
    rect(slide, px, y, pw, panel_h, fill=tint)
    text(slide, px + 0.22, y + 0.14, pw - 0.40, 0.30, panel_header, 12, red, bold=True)
    hline(slide, px + 0.22, y + 0.48, px + pw - 0.22, color=red, lw=1.2)
    bullets(slide, px + 0.22, y + 0.58, pw - 0.42, panel_h - 0.70, panel_items, size=10, gap=6,
            marker="▶", marker_color=black)


def icon_trio(slide, x, y, w, items, box_h=1.0):
    """Icon + bold head + grey sub, 3 to 5 columns inside a dashed frame.
    items = [(icon_name, head, sub)]"""
    n = len(items); cw = w / n
    rect(slide, x, y, w, box_h, line=red, lw=1.0, dash=MSO_LINE_DASH_STYLE.DASH)
    for i, (ic, head, sub) in enumerate(items):
        cx = x + i * cw + cw / 2
        icon(slide, ic, cx - 0.20, y + 0.10, 0.40, color="red")
        text(slide, x + i * cw + 0.05, y + 0.55, cw - 0.10, 0.22, head, 10.5, black, bold=True, align="c")
        text(slide, x + i * cw + 0.05, y + 0.76, cw - 0.10, 0.20, sub, 8.5, iron, align="c")
        if i > 0: vline(slide, x + i * cw, y + 0.12, y + box_h - 0.12, color=steel, lw=0.8)


def logo_chip_table(slide, x, y, w, rows, chip_w=1.0, row_h=0.62, fsize=9.5, name_w=2.0):
    """Zebra list with a logo (image path) or a coloured initials chip per row.
    rows = [(logo_path_or_initials, chip_colour, name, description)]"""
    for i, (logo, col, name, desc) in enumerate(rows):
        ry = y + i * (row_h + 0.06)
        rect(slide, x, ry, w, row_h, fill=light if i % 2 == 0 else white)
        if os.path.exists(str(logo)):
            rect(slide, x + 0.12, ry + 0.08, chip_w, row_h - 0.16, fill=white)
            picture_fit(slide, logo, x + 0.16, ry + 0.11, chip_w - 0.08, row_h - 0.22)
        else:
            chip = rect(slide, x + 0.12, ry + (row_h - 0.40) / 2, chip_w, 0.40, fill=col)
            shape_text(chip, logo, 10, white, bold=True)
        text(slide, x + chip_w + 0.30, ry + 0.06, name_w, row_h - 0.12, name, fsize, red, bold=True, anchor="m")
        text(slide, x + chip_w + name_w + 0.40, ry + 0.06, w - chip_w - name_w - 0.55, row_h - 0.12, desc,
             fsize, black, anchor="m")


def logo_wall(slide, x, y, w, h, logos, cols=4, gap=0.20, captions=None):
    """Grid of real logos, each fitted inside a light cell. logos = [image paths]."""
    n = len(logos); rows_n = (n + cols - 1) // cols
    cw = (w - gap * (cols - 1)) / cols; chh = (h - gap * (rows_n - 1)) / rows_n
    for i, p in enumerate(logos):
        r, c = divmod(i, cols)
        cx = x + c * (cw + gap); cy = y + r * (chh + gap)
        rect(slide, cx, cy, cw, chh, fill=light)
        cap_h = 0.28 if captions else 0
        picture_fit(slide, p, cx + 0.20, cy + 0.15, cw - 0.40, chh - 0.30 - cap_h)
        if captions:
            text(slide, cx, cy + chh - 0.32, cw, 0.26, captions[i], 9, iron, align="c")


def resource_cards(slide, x, y, w, cards, card_h=1.75, gap=0.25):
    """Coloured thumbnail cards with title, plus italic date/subtitle below.
    cards = [(title, date, fill)]"""
    n = len(cards); cw = (w - gap * (n - 1)) / n
    for i, (ttl, date, col) in enumerate(cards):
        cx = x + i * (cw + gap)
        r = rect(slide, cx, y, cw, card_h * 0.62, fill=col)
        shape_text(r, ttl, 10.5, white, bold=True, align="l")
        r.text_frame.margin_left = Inches(0.14); r.text_frame.margin_right = Inches(0.14)
        text(slide, cx, y + card_h * 0.62 + 0.10, cw, 0.30, date, 9.5, iron, italic=True)


def timeline(slide, x, y, w, milestones, fill=red):
    """Horizontal timeline. milestones = [(date_label, title, detail)], labels alternate above/below."""
    n = len(milestones)
    hline(slide, x, y, x + w, color=iron, lw=2.0)
    step = w / n
    for i, (dt, ttl, det) in enumerate(milestones):
        cx = x + step * i + step / 2
        oval(slide, cx - 0.11, y - 0.11, 0.22, 0.22, fill=fill if i % 2 == 0 else skibo)
        above = i % 2 == 0
        ty = y - 1.45 if above else y + 0.28
        text(slide, cx - step / 2 + 0.05, ty, step - 0.10, 0.26, dt, 10, red, bold=True, align="c")
        text(slide, cx - step / 2 + 0.05, ty + 0.27, step - 0.10, 0.28, ttl, 11, black, bold=True, align="c")
        text(slide, cx - step / 2 + 0.05, ty + 0.56, step - 0.10, 0.60, det, 9, iron, align="c")
        vline(slide, cx, y - 0.30 if above else y + 0.11, y - 0.11 if above else y + 0.26, color=steel, lw=1.0)


def matrix_2x2(slide, x, y, w, h, quads, x_label, y_label, highlight=None):
    """2 x 2 prioritisation matrix. quads = [(title, body)] in order TL, TR, BL, BR.
    highlight = index of the quadrant to fill red (the recommendation)."""
    lw_ = 0.45
    gx, gy, gw, gh = x + lw_, y, w - lw_, h - 0.40
    qw, qh = (gw - 0.10) / 2, (gh - 0.10) / 2
    for i, (ttl, body) in enumerate(quads):
        r, c = divmod(i, 2)
        qx = gx + c * (qw + 0.10); qy = gy + r * (qh + 0.10)
        hot = (highlight == i)
        rect(slide, qx, qy, qw, qh, fill=red if hot else (tint if (r + c) % 2 == 0 else light))
        text(slide, qx + 0.20, qy + 0.16, qw - 0.40, 0.32, ttl, 12.5, white if hot else red, bold=True)
        text(slide, qx + 0.20, qy + 0.52, qw - 0.40, qh - 0.64, body, 10, white if hot else black)
    text(slide, gx, gy + gh + 0.08, gw, 0.28, x_label + "  →", 10, iron, bold=True, align="c")
    tb = text(slide, x - 1.0, y + gh / 2 - 0.15, gh, 0.30, y_label + "  →", 10, iron, bold=True, align="c")
    tb.rotation = 270
    tb.left = Inches(x + lw_ / 2 - gh / 2 - 0.08); tb.width = Inches(gh)


def quote_block(slide, x, y, w, quote, attribution, h=1.6):
    """Large pull quote with red rail and attribution."""
    rect(slide, x, y, 0.07, h, fill=red)
    text(slide, x + 0.35, y + 0.05, w - 0.40, h - 0.45, "“" + quote + "”", 18, black, italic=True,
         line=1.2)
    text(slide, x + 0.35, y + h - 0.34, w - 0.40, 0.30, attribution, 10.5, iron, bold=True)


def harvey_table(slide, x, y, w, h, row_labels, col_labels, scores, label_w=3.0, fsize=10.5):
    """Option-vs-criteria scoring grid with filled circles (0 to 4 quarters).
    scores[r][c] in 0..4. Uses pie shapes so it stays native."""
    nr, nc = len(row_labels), len(col_labels)
    cw = (w - label_w) / nc; rh = (h - 0.46) / nr
    rect(slide, x, y, w, 0.46, fill=red)
    for j, cl in enumerate(col_labels):
        text(slide, x + label_w + j * cw, y, cw, 0.46, cl, fsize, white, bold=True, align="c", anchor="m")
    for i, rl in enumerate(row_labels):
        ry = y + 0.46 + i * rh
        rect(slide, x, ry, w, rh, fill=light if i % 2 else white)
        text(slide, x + 0.12, ry, label_w - 0.2, rh, rl, fsize, black, bold=True, anchor="m")
        for j in range(nc):
            d = min(0.34, rh * 0.7)
            ccx = x + label_w + j * cw + cw / 2 - d / 2; ccy = ry + rh / 2 - d / 2
            oval(slide, ccx, ccy, d, d, fill=white, line=red, lw=1.2)
            sc = scores[i][j]
            if sc >= 4:
                oval(slide, ccx, ccy, d, d, fill=red)
            elif sc > 0:
                pie = slide.shapes.add_shape(MSO_SHAPE.PIE, Inches(ccx), Inches(ccy), Inches(d), Inches(d))
                # python-pptx scales angle adjustments by 1/100000 of 60000ths: value = degrees * 0.6
                pie.adjustments[0] = 270 * 0.6                       # start at 12 o'clock
                pie.adjustments[1] = ((270 + 90 * sc) % 360) * 0.6   # sweep clockwise sc quarters
                pie.fill.solid(); pie.fill.fore_color.rgb = red
                pie.line.fill.background(); pie.shadow.inherit = False


# ───────────────────────────── quality checks ─────────────────────────────
def _est_lines(txt, size_pt, width_in, bold=False):
    """Rough line count for Open Sans text wrapped at width_in."""
    cw = size_pt / 72.0 * (0.56 if bold else 0.52)
    if width_in <= 0: return 99
    per_line = max(1, int(width_in / cw))
    lines = 0
    for seg in txt.split("\n"):
        words = seg.split(" "); cur = 0; n = 1
        for wd in words:
            L = len(wd) + (1 if cur else 0)
            if cur + L > per_line and cur > 0: n += 1; cur = len(wd)
            else: cur += L
        lines += n
    return lines


def _tf_need(tf, width_in, wrap=True):
    """Estimated (height_in, widest_line_in) of a text frame."""
    hgt = 0.0; widest = 0.0
    for p in tf.paragraphs:
        txt = "".join(r.text for r in p.runs)
        sizes = [r.font.size.pt for r in p.runs if r.font.size] or [12]
        sz = max(sizes)
        bold = any(r.font.bold for r in p.runs)
        ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
        n = _est_lines(txt, sz, width_in, bold) if wrap else txt.count("\n") + 1
        hgt += n * sz * 1.22 * ls / 72.0
        if p.space_after is not None: hgt += p.space_after.pt / 72.0
        widest = max(widest, len(txt) * sz / 72.0 * (0.56 if bold else 0.52))
    if tf.paragraphs and tf.paragraphs[-1].space_after is not None:
        hgt -= tf.paragraphs[-1].space_after.pt / 72.0
    return hgt, widest


def _box(sh):
    return (Emu(sh.left).inches, Emu(sh.top).inches, Emu(sh.width).inches, Emu(sh.height).inches)


def _inter(a, b):
    ix = max(0, min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0]))
    iy = max(0, min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]))
    return ix * iy


def lint(deck, verbose=True):
    """Static layout checks on every slide. Returns a list of problem strings.
    Checks: off-canvas shapes, body content intruding into the title / logo /
    wordmark / sidebar / source zones, text that will overflow its box, tables
    whose rows will grow past the next element, and banned dash characters."""
    problems = []
    for idx, s in enumerate(deck.prs.slides, start=1):
        meta = deck.meta.get(s.slide_id, {})
        kind = meta.get("kind")
        nf = meta.get("n_furniture", 0)
        shapes = list(s.shapes)
        body = shapes[nf:] if kind == "content" else shapes
        for sh in shapes:
            if sh.left is None: continue
            l, t, w, h = _box(sh)
            if kind == "title" and sh.shape_type == 9: continue   # decorative lattice lines bleed by design
            if sh.rotation in (0, 0.0) and (l < -0.01 or t < -0.01 or l + w > slide_w + 0.02 or t + h > slide_h + 0.02):
                problems.append(f"slide {idx}: '{sh.name}' runs off the slide ({l:.2f},{t:.2f},{w:.2f},{h:.2f})")
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if "\u2014" in r.text or "\u2013" in r.text:
                            problems.append(f"slide {idx}: dash character in '{r.text[:50]}'")
            if sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        if "\u2014" in c.text or "\u2013" in c.text:
                            problems.append(f"slide {idx}: dash character in table cell '{c.text[:50]}'")
        if kind == "content":
            left = meta["left"]
            lim = body_max_src if meta.get("has_source") else body_max
            top_lim = body_y_sub - 0.08 if meta.get("has_sub") else rule_y + 0.12
            for sh in body:
                if sh.left is None: continue
                l, t, w, h = _box(sh)
                if sh.rotation not in (0, 0.0): continue
                tag = f"slide {idx}: '{sh.name}'"
                if l < left - 0.03: problems.append(f"{tag} starts left of the content margin ({l:.2f} < {left:.2f})")
                if l + w > content_r + 0.03: problems.append(f"{tag} passes the right content edge ({l + w:.2f} > {content_r})")
                if t < top_lim: problems.append(f"{tag} starts in the title zone (top {t:.2f} < {top_lim:.2f})")
                if t + h > lim + 0.02: problems.append(f"{tag} ends at {t + h:.2f}, below the body limit {lim:.2f}")
        # text overflow
        for sh in shapes:
            if not sh.has_text_frame or sh.left is None: continue
            tf = sh.text_frame
            if not tf.text.strip(): continue
            l, t, w, h = _box(sh)
            ml = Emu(tf.margin_left or 0).inches + Emu(tf.margin_right or 0).inches
            mt = Emu(tf.margin_top or 0).inches + Emu(tf.margin_bottom or 0).inches
            need_h, widest = _tf_need(tf, w - ml, tf.word_wrap is not False)
            if tf.word_wrap is False and widest > (slide_w - l) - 0.05:
                problems.append(f"slide {idx}: '{tf.text[:40]}' is one line and runs past the slide edge")
            elif need_h > (h - mt) * 1.18 + 0.06:
                problems.append(f"slide {idx}: text likely overflows its box ({need_h:.2f}in needed, "
                                f"{h - mt:.2f}in available): '{tf.text[:50]}'")
        # tables that will grow
        for sh in shapes:
            if not sh.has_table: continue
            l, t, w, h = _box(sh)
            est = 0.0
            for row in sh.table.rows:
                rh = Emu(row.height).inches; need = rh
                for ci, c in enumerate(row.cells):
                    if c.is_spanned: continue
                    cw = Emu(sh.table.columns[ci].width).inches - 0.16
                    nh, _ = _tf_need(c.text_frame, cw)
                    need = max(need, nh + 0.10)
                est += need
            bottom = t + est
            for other in shapes:
                if other is sh or other.left is None: continue
                ol, ot, ow, oh = _box(other)
                if ot > t + 0.05 and ot < bottom - 0.03 and _inter((l, t, w, est), (ol, ot, ow, oh)) > 0.01:
                    problems.append(f"slide {idx}: table will grow to {bottom:.2f}in and collide with "
                                    f"'{other.name}' at {ot:.2f}in; shrink rows, font or text")
                    break
            if kind == "content":
                lim = body_max_src if meta.get("has_source") else body_max
                if bottom > lim + 0.02:
                    problems.append(f"slide {idx}: table will grow to {bottom:.2f}in, past the body limit {lim:.2f}")
        # text boxes whose estimated text extents overlap (real collisions)
        ext = []
        for sh in body:
            if not (sh.has_text_frame and sh.text_frame.text.strip() and sh.shape_type == 17
                    and sh.rotation in (0, 0.0)): continue
            l, t, w, h = _box(sh); tf = sh.text_frame
            need, widest = _tf_need(tf, w)
            need = min(need, h * 1.6)
            anc = tf.vertical_anchor
            if anc == MSO_ANCHOR.MIDDLE: t = t + (h - need) / 2
            elif anc == MSO_ANCHOR.BOTTOM: t = t + h - need
            ww = min(w, widest) if tf.word_wrap is not False else widest
            al = tf.paragraphs[0].alignment
            if al == PP_ALIGN.CENTER: l = l + (w - ww) / 2
            elif al == PP_ALIGN.RIGHT: l = l + w - ww
            ext.append((sh, (l, t, ww, need)))
        for i in range(len(ext)):
            for j in range(i + 1, len(ext)):
                if _inter(ext[i][1], ext[j][1]) > 0.02:
                    problems.append(f"slide {idx}: text collides: '{ext[i][0].text_frame.text[:30]}' / "
                                    f"'{ext[j][0].text_frame.text[:30]}'")
    if verbose:
        print(f"lint: {len(problems)} problem(s)")
        for p in problems: print("  -", p)
    return problems


def render(pptx_path, out_dir=None, dpi=80):
    """Convert the deck to PDF with LibreOffice and to one PNG per slide.
    Returns (pdf_path, [png paths]); pngs is empty if no rasteriser is available."""
    out_dir = out_dir or os.path.join(os.path.dirname(os.path.abspath(pptx_path)), "render")
    os.makedirs(out_dir, exist_ok=True)
    office = shutil.which("soffice") or shutil.which("libreoffice")
    if not office:
        for cand in ["/Applications/LibreOffice.app/Contents/MacOS/soffice", "/opt/homebrew/bin/soffice"]:
            if os.path.exists(cand): office = cand
    if not office:
        print("render: LibreOffice not found; rely on lint() only"); return None, []
    subprocess.run([office, "--headless", "--convert-to", "pdf", "--outdir", out_dir, pptx_path],
                   check=True, capture_output=True, timeout=300)
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
    pngs = []
    for f in glob.glob(os.path.join(out_dir, "page-*.png")): os.remove(f)
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf, os.path.join(out_dir, "page")], check=True)
        pngs = sorted(glob.glob(os.path.join(out_dir, "page-*.png")))
    else:
        try:
            import fitz
            doc = fitz.open(pdf)
            for i, pg in enumerate(doc, start=1):
                p = os.path.join(out_dir, f"page-{i:02d}.png")
                pg.get_pixmap(dpi=dpi).save(p); pngs.append(p)
        except Exception:
            print("render: no pdftoppm or PyMuPDF; open the PDF pages directly")
    return pdf, pngs
