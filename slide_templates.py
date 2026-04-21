"""
Revenue Leak Audit — Slide Templates
=====================================
Every slide is a self-contained function.
Call generate(data) to produce a complete .pptx from a data dictionary.

Usage:
    from slide_templates import generate
    generate(data, "output.pptx")

Data dictionary keys are documented in DATA_SCHEMA at the bottom of this file.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn as _qn
from pptx.util import Pt
import copy
from lxml import etree

def qn(tag):
    return _qn(tag)

# ─────────────────────────────────────────────
# DESIGN CONSTANTS
# ─────────────────────────────────────────────

# Colors
C_DARK_BG      = RGBColor(0x1C, 0x1B, 0x2E)   # near-black navy
C_LIGHT_BG     = RGBColor(0xFF, 0xF8, 0xF0)   # warm cream
C_PURPLE       = RGBColor(0x7B, 0x5E, 0xA7)   # purple accent
C_GREEN        = RGBColor(0x00, 0xC3, 0x89)   # green highlight
C_RED          = RGBColor(0xE6, 0x39, 0x46)   # red accent
C_WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
C_DARK_TEXT    = RGBColor(0x1C, 0x1B, 0x2E)
C_MUTED        = RGBColor(0x99, 0x99, 0x99)
C_CARD_BG      = RGBColor(0xFF, 0xFF, 0xFF)
C_HIGHLIGHT    = RGBColor(0xEE, 0xF0, 0xFF)   # very light purple-blue
C_AMBER        = RGBColor(0xE0, 0x7B, 0x00)
C_DARK_CARD    = RGBColor(0x2A, 0x28, 0x45)   # slightly lighter than bg for dark cards
C_SCORE_RED    = RGBColor(0xCC, 0x22, 0x22)   # score circle red
C_GREEN_BADGE  = RGBColor(0x00, 0x99, 0x55)
C_AMBER_BADGE  = RGBColor(0xE0, 0x7B, 0x00)
C_RED_BADGE    = RGBColor(0xCC, 0x22, 0x22)
C_TEAL         = RGBColor(0x00, 0xC3, 0x89)   # same as green, used for guarantee borders
C_FUNNEL_BAR   = RGBColor(0x7B, 0x5E, 0xA7)   # purple funnel bars

# Slide dimensions
W = Inches(13.33)
H = Inches(7.5)

# Margins
LM = Inches(0.6)   # left margin
RM = Inches(0.6)   # right margin
TM = Inches(0.5)   # top margin (below accent bar)


# ─────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────

def new_slide(prs):
    """Add a blank slide and return it."""
    blank_layout = prs.slide_layouts[6]  # completely blank
    return prs.slides.add_slide(blank_layout)


def rgb_fill(shape, color):
    """Fill a shape with a solid RGB color."""
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = color


def no_fill(shape):
    shape.fill.background()


def add_rect(slide, left, top, width, height, fill_color=None, line_color=None, line_width=None):
    """Add a rectangle shape."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    if fill_color:
        rgb_fill(shape, fill_color)
    else:
        no_fill(shape)
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape


def add_rounded_rect(slide, left, top, width, height, fill_color, radius_pt=6, line_color=None):
    """Add a rounded rectangle."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    shape = slide.shapes.add_shape(
        5,  # ROUNDED_RECTANGLE
        left, top, width, height
    )
    # Set corner radius
    shape.adjustments[0] = radius_pt / 100.0
    if fill_color:
        rgb_fill(shape, fill_color)
    else:
        no_fill(shape)
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, left, top, width, height, text, font_name, font_size,
                bold=False, italic=False, color=C_DARK_TEXT, align=PP_ALIGN.LEFT,
                word_wrap=True, line_spacing=None):
    """Add a text box. \\n in text creates separate paragraphs."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap

    lines = text.split('\n')
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        if line_spacing:
            pPr = p._p.get_or_add_pPr()
            lnSpc = etree.SubElement(pPr, _qn('a:lnSpc'))
            spcPts = etree.SubElement(lnSpc, _qn('a:spcPts'))
            spcPts.set('val', str(int(line_spacing * 100)))

    # Remove internal padding
    txBody = txBox.text_frame._txBody
    bodyPr = txBody.find(_qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('lIns', '0')
        bodyPr.set('rIns', '0')
        bodyPr.set('tIns', '0')
        bodyPr.set('bIns', '0')
    return txBox


def add_vcenter_text(slide, left, top, width, height, text, font_name, font_size,
                     bold=False, italic=False, color=C_DARK_TEXT, align=PP_ALIGN.CENTER):
    """
    Vertically centres multi-line text within a region using a plain textbox.
    Calculates top padding so text sits in the middle of the available height.
    Uses pure black (overrides color param) for maximum legibility on light cards.
    """
    lines = text.split('\n')
    line_height_inches = font_size * 0.0175  # approx inches per pt at 1.25x leading
    total_text_h = len(lines) * line_height_inches
    top_pad = max(0, (height.inches if hasattr(height, 'inches') else height / 914400) - total_text_h) / 2
    adjusted_top = top + Inches(top_pad)

    txBox = slide.shapes.add_textbox(left, adjusted_top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  # always pure black

    # Zero internal padding
    txBody = txBox.text_frame._txBody
    bodyPr = txBody.find(_qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('lIns', '0')
        bodyPr.set('rIns', '0')
        bodyPr.set('tIns', '0')
        bodyPr.set('bIns', '0')
    return txBox


def add_accent_bar(slide, dark=True):
    """Add the full-width purple top accent bar."""
    height = Inches(0.12) if dark else Inches(0.08)
    bar = add_rect(slide, 0, 0, W, height, fill_color=C_PURPLE)
    return bar


def add_footer(slide, brand_name, slide_number):
    """Add standard footer: brand left, slide number right."""
    footer_y = H - Inches(0.35)
    footer_h = Inches(0.25)
    # Left text
    add_textbox(slide, LM, footer_y, Inches(6), footer_h,
                f"{brand_name}  ·  Revenue Leak Audit  ·  Confidential",
                "Calibri", 10, color=C_MUTED)
    # Right number
    add_textbox(slide, W - Inches(1.2), footer_y, Inches(0.8), footer_h,
                str(slide_number), "Calibri", 10, color=C_MUTED, align=PP_ALIGN.RIGHT)


def add_section_label(slide, text, left=None, top=Inches(0.45), dark=False):
    """Add a spaced-caps purple section label."""
    if left is None:
        left = LM
    color = C_PURPLE
    add_textbox(slide, left, top, W - left - RM, Inches(0.25),
                text, "Arial", 11, bold=True, color=color)


def add_title(slide, text, top=Inches(0.78), font_size=40, color=C_DARK_TEXT, width=None):
    """Add a slide title."""
    if width is None:
        width = W - LM - RM
    add_textbox(slide, LM, top, width, Inches(1.4),
                text, "Georgia", font_size, bold=True, color=color)


def set_shape_text(shape, text, font_name, font_size, bold=False, italic=False,
                   color=C_DARK_TEXT, align=PP_ALIGN.LEFT):
    """Set text on an existing shape's text frame."""
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_multi_para_textbox(slide, left, top, width, height, paragraphs,
                            word_wrap=True):
    """
    Add a textbox with multiple paragraphs.
    paragraphs = list of dicts:
      {text, font_name, font_size, bold, italic, color, align, space_before}
    """
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    # Remove padding
    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('lIns', '0'); bodyPr.set('rIns', '0')
        bodyPr.set('tIns', '0'); bodyPr.set('bIns', '0')

    for i, para in enumerate(paragraphs):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = para.get('align', PP_ALIGN.LEFT)
        if para.get('space_before'):
            pPr = p._p.get_or_add_pPr()
            spcBef = etree.SubElement(pPr, qn('a:spcBef'))
            spcPts = etree.SubElement(spcBef, qn('a:spcPts'))
            spcPts.set('val', str(int(para['space_before'] * 100)))
        run = p.add_run()
        run.text = para.get('text', '')
        run.font.name = para.get('font_name', 'Calibri')
        run.font.size = Pt(para.get('font_size', 14))
        run.font.bold = para.get('bold', False)
        run.font.italic = para.get('italic', False)
        run.font.color.rgb = para.get('color', C_DARK_TEXT)
    return txBox


def add_circle(slide, cx, cy, diameter, fill_color, text, text_color=C_WHITE,
               font_size=14, bold=True):
    """Add a filled circle with centered text (for numbered badges)."""
    r = diameter / 2
    shape = slide.shapes.add_shape(
        9,  # OVAL
        int(cx - r), int(cy - r), int(diameter), int(diameter)
    )
    rgb_fill(shape, fill_color)
    shape.line.fill.background()
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name = "Georgia"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = text_color
    # Vertical center
    from pptx.enum.text import MSO_ANCHOR
    tf.auto_size = None
    tf.word_wrap = False
    shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    return shape


# ─────────────────────────────────────────────
# SLIDE 1 — COVER
# ─────────────────────────────────────────────

def slide_01_cover(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_DARK_BG)
    add_accent_bar(slide, dark=True)

    # "REVENUE LEAK AUDIT" label — near top
    add_textbox(slide, LM, Inches(0.35), Inches(5), Inches(0.3),
                "REVENUE LEAK AUDIT", "Arial", 11, bold=True, color=C_PURPLE)

    # Client name — large, starts at roughly 1/4 down the slide
    add_textbox(slide, LM, Inches(1.3), W - LM - RM, Inches(2.1),
                d["client_name"], "Georgia", 72, bold=True, color=C_WHITE)

    # Subtitle — immediately below name
    add_textbox(slide, LM, Inches(3.45), Inches(9), Inches(0.38),
                "AI Readiness Assessment & Revenue Recovery Analysis",
                "Calibri", 18, color=C_MUTED)

    # Red divider line
    add_rect(slide, LM, Inches(3.98), Inches(1.2), Inches(0.05), fill_color=C_RED)

    # Prepared for / date block — clear gap below divider
    add_textbox(slide, LM, Inches(4.22), Inches(8), Inches(0.3),
                f"Prepared for {d['prepared_for']} — {d['client_name']}",
                "Calibri", 14, color=C_WHITE)
    add_textbox(slide, LM, Inches(4.56), Inches(4), Inches(0.3),
                d["date"], "Calibri", 14, color=C_WHITE)

    # CONFIDENTIAL — near bottom
    add_textbox(slide, LM, Inches(5.9), Inches(3), Inches(0.3),
                "CONFIDENTIAL", "Arial", 11, bold=True, color=C_RED)


# ─────────────────────────────────────────────
# SLIDE 2 — AGENDA
# ─────────────────────────────────────────────

def slide_02_agenda(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    add_section_label(slide, "TODAY'S AGENDA")
    add_title(slide, "What we'll walk through together", top=Inches(0.72), font_size=40)

    items = [
        ("01", "Your Industry",          d["agenda_01_desc"]),
        ("02", "Your Business Today",    "What we found when we audited your " + d["term_business"]),
        ("03", "Your AI Readiness Score","6 dimensions, scored with specific findings"),
        ("04", "The Revenue Leak",       "Exact dollars you're leaving on the table"),
        ("05", "Your AI Systems",        "The specific bots we recommend and how each works"),
        ("06", "The Guarantee",          "Our 90-day ROI guarantee — risk-free"),
        ("07", "Next Steps",             "Select your bots and we start building"),
    ]

    row_h = Inches(0.62)
    row_top = Inches(1.95)
    row_w = W - LM - RM

    for i, (num, title, desc) in enumerate(items):
        bg = C_HIGHLIGHT if i % 2 == 0 else C_WHITE
        add_rounded_rect(slide, LM, row_top + i * row_h, row_w, row_h - Inches(0.04),
                         fill_color=bg, radius_pt=4)
        # Number
        add_textbox(slide, LM + Inches(0.12), row_top + i * row_h + Inches(0.12),
                    Inches(0.5), Inches(0.38),
                    num, "Georgia", 22, bold=True, color=C_PURPLE)
        # Title
        add_textbox(slide, LM + Inches(0.65), row_top + i * row_h + Inches(0.14),
                    Inches(2.8), Inches(0.34),
                    title, "Calibri", 15, bold=True, color=C_DARK_TEXT)
        # Description
        add_textbox(slide, LM + Inches(3.6), row_top + i * row_h + Inches(0.16),
                    Inches(8.5), Inches(0.3),
                    desc, "Calibri", 14, color=C_MUTED)

    add_footer(slide, d["brand_name"], 2)


# ─────────────────────────────────────────────
# SLIDE 3 — PART ONE DIVIDER
# ─────────────────────────────────────────────

def slide_03_part_one(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_DARK_BG)
    add_accent_bar(slide, dark=True)

    # Left purple bar
    add_rect(slide, 0, 0, Inches(0.08), H, fill_color=C_PURPLE)

    add_textbox(slide, LM, Inches(1.8), Inches(4), Inches(0.3),
                "PART ONE", "Arial", 11, bold=True, color=C_PURPLE)

    add_textbox(slide, LM, Inches(2.25), W - LM - RM, Inches(2.0),
                d["part_one_title"], "Georgia", 52, bold=True, color=C_WHITE)

    add_textbox(slide, LM, Inches(4.45), Inches(9), Inches(0.4),
                d["part_one_subtitle"], "Calibri", 18, color=C_MUTED)


# ─────────────────────────────────────────────
# SLIDE 4 — CLIENT JOURNEY
# ─────────────────────────────────────────────

def slide_04_journey(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    add_section_label(slide, d["journey_label"])
    add_title(slide, d["journey_title"], top=Inches(0.72), font_size=36)

    stages = d["journey_stages"]
    n = len(stages)
    total_w = W - LM - RM
    card_w = (total_w - Inches(0.1) * (n - 1)) / n
    card_h = Inches(1.42)   # sized to content — no empty space inside
    card_top = Inches(2.05)

    stage_colors = [
        RGBColor(0x5B, 0x8D, 0xE8), RGBColor(0x7B, 0x5E, 0xA7),
        RGBColor(0xE0, 0x7B, 0x00), RGBColor(0x00, 0xC3, 0x89),
        RGBColor(0x2E, 0x8B, 0x6E), RGBColor(0x00, 0x99, 0xCC),
    ]

    for i, stage in enumerate(stages):
        cx = LM + i * (card_w + Inches(0.1))
        add_rounded_rect(slide, cx, card_top, card_w, card_h,
                         fill_color=C_CARD_BG, radius_pt=4,
                         line_color=RGBColor(0xDD, 0xDD, 0xDD))
        color = stage_colors[i % len(stage_colors)]
        add_rect(slide, cx, card_top, card_w, Inches(0.06), fill_color=color)
        add_textbox(slide, cx + Inches(0.08), card_top + Inches(0.1),
                    card_w - Inches(0.16), Inches(0.28),
                    stage["name"], "Arial", 9, bold=True, color=color,
                    align=PP_ALIGN.CENTER)
        # Description vertically centred in remaining card space
        add_vcenter_text(slide, cx + Inches(0.05), card_top + Inches(0.42),
                         card_w - Inches(0.1), card_h - Inches(0.42),
                         stage["desc"], "Calibri", 11,
                         color=RGBColor(0x55, 0x55, 0x55))

        # Drop-off callout — immediately below card
        if stage.get("pct_lost"):
            add_textbox(slide, cx, card_top + card_h + Inches(0.1),
                        card_w, Inches(0.28),
                        stage["pct_lost"] + " lost", "Calibri", 13, bold=True,
                        color=C_RED, align=PP_ALIGN.CENTER)
            add_textbox(slide, cx, card_top + card_h + Inches(0.38),
                        card_w, Inches(0.24),
                        stage.get("reason", ""), "Calibri", 10,
                        color=C_MUTED, align=PP_ALIGN.CENTER)

    # Closing line — anchored just below callouts
    add_textbox(slide, LM, card_top + card_h + Inches(0.78), W - LM - RM, Inches(0.35),
                d["journey_closing"], "Calibri", 13, italic=True, color=C_RED)

    add_footer(slide, d["brand_name"], 4)


# ─────────────────────────────────────────────
# SLIDE 5 — INDUSTRY STATS
# ─────────────────────────────────────────────

def slide_05_stats(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    add_section_label(slide, d["stat_section_label"])
    add_title(slide, "What the data says about businesses like yours",
              top=Inches(0.72), font_size=36)

    stats = d["industry_stats"]
    card_w = Inches(5.8)
    card_h = Inches(1.9)   # taller so number doesn't crowd label
    gap = Inches(0.25)
    top_row_y = Inches(2.1)
    bot_row_y = top_row_y + card_h + Inches(0.2)
    positions = [
        (LM, top_row_y),
        (LM + card_w + gap, top_row_y),
        (LM, bot_row_y),
        (LM + card_w + gap, bot_row_y),
    ]

    for i, (stat, (cx, cy)) in enumerate(zip(stats, positions)):
        add_rounded_rect(slide, cx, cy, card_w, card_h,
                         fill_color=C_CARD_BG, radius_pt=4,
                         line_color=RGBColor(0xE0, 0xE0, 0xE0))
        # Big number — more vertical room
        add_textbox(slide, cx + Inches(0.18), cy + Inches(0.12),
                    card_w - Inches(0.36), Inches(0.82),
                    stat["number"], "Georgia", 48, bold=True, color=C_DARK_TEXT)
        # Label — pushed down so it doesn't overlap number
        add_textbox(slide, cx + Inches(0.18), cy + Inches(0.9),
                    card_w - Inches(0.36), Inches(0.68),
                    stat["label"], "Calibri", 12, color=C_DARK_TEXT)
        # Source
        add_textbox(slide, cx + Inches(0.18), cy + Inches(1.58),
                    card_w - Inches(0.36), Inches(0.26),
                    stat["source"], "Calibri", 10, italic=True, color=C_MUTED)

    add_footer(slide, d["brand_name"], 5)


# ─────────────────────────────────────────────
# SLIDE 6 — LEAD LIFECYCLE
# ─────────────────────────────────────────────

def slide_06_lifecycle(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    add_section_label(slide, "THE LEAD LIFECYCLE")
    add_title(slide, d["lifecycle_title"], top=Inches(0.72), font_size=32)

    # Intro paragraph — more gap below title
    add_textbox(slide, LM, Inches(1.88), W - LM - RM, Inches(0.5),
                d["lifecycle_intro"], "Calibri", 13, color=C_DARK_TEXT)

    stages = d["lifecycle_stages"]
    n = 6
    total_w = W - LM - RM
    card_w = (total_w - Inches(0.08) * (n - 1)) / n
    card_h = Inches(1.95)   # height matches content: label + desc + badge
    card_top = Inches(2.55)

    for i, stage in enumerate(stages):
        cx = LM + i * (card_w + Inches(0.08))
        add_rounded_rect(slide, cx, card_top, card_w, card_h,
                         fill_color=C_CARD_BG, radius_pt=4,
                         line_color=RGBColor(0xDD, 0xDD, 0xDD))
        # Stage name
        add_textbox(slide, cx + Inches(0.06), card_top + Inches(0.1),
                    card_w - Inches(0.12), Inches(0.32),
                    stage["name"], "Arial", 9, bold=True, color=C_PURPLE,
                    align=PP_ALIGN.CENTER)
        # Description vertically centred between name and badge
        add_vcenter_text(slide, cx + Inches(0.05), card_top + Inches(0.44),
                         card_w - Inches(0.1), card_h - Inches(0.44) - Inches(0.44),
                         stage["desc"], "Calibri", 10,
                         color=RGBColor(0x55, 0x55, 0x55))

        # Status badge pinned to bottom
        status = stage["status"].upper()
        badge_color = C_GREEN_BADGE if status == "WORKING" else (
            C_AMBER_BADGE if status == "MANUAL" else C_RED_BADGE)
        badge_h = Inches(0.32)
        badge_top = card_top + card_h - badge_h - Inches(0.1)
        badge = add_rounded_rect(slide, cx + Inches(0.1), badge_top,
                                  card_w - Inches(0.2), badge_h,
                                  fill_color=badge_color, radius_pt=4)
        set_shape_text(badge, status, "Arial", 9, bold=True,
                       color=C_WHITE, align=PP_ALIGN.CENTER)
        badge.text_frame.vertical_anchor = __import__(
            'pptx.enum.text', fromlist=['MSO_ANCHOR']).MSO_ANCHOR.MIDDLE

    # Summary line — anchored just below cards
    add_textbox(slide, LM, card_top + card_h + Inches(0.18), W - LM - RM, Inches(0.45),
                d["lifecycle_summary"], "Calibri", 13, italic=True, color=C_RED)

    add_footer(slide, d["brand_name"], 6)


# ─────────────────────────────────────────────
# SLIDE 7 — REVENUE FUNNEL
# ─────────────────────────────────────────────

def slide_07_funnel(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    add_section_label(slide, "YOUR REVENUE FUNNEL")
    add_title(slide, "Where your monthly leads actually end up",
              top=Inches(0.72), font_size=36)

    funnel = d["funnel_rows"]  # list of 5: {label, value, lost}   last row has no lost
    bar_top = Inches(1.95)
    bar_h = Inches(0.56)
    bar_gap = Inches(0.06)
    max_w = W - LM - RM - Inches(1.8)

    widths = [1.0, 0.82, 0.62, 0.42, 0.28]

    for i, (row, w_frac) in enumerate(zip(funnel, widths)):
        bar_w = max_w * w_frac
        y = bar_top + i * (bar_h + bar_gap)
        # Funnel bar
        bar = add_rounded_rect(slide, LM, y, bar_w, bar_h,
                                fill_color=C_FUNNEL_BAR, radius_pt=3)
        # Label inside bar
        add_textbox(slide, LM + Inches(0.15), y + Inches(0.1),
                    bar_w - Inches(0.3), Inches(0.36),
                    row["label"], "Calibri", 13, color=C_WHITE)
        # Value at right end of bar
        add_textbox(slide, LM + bar_w - Inches(1.2), y + Inches(0.05),
                    Inches(1.1), Inches(0.46),
                    str(row["value"]), "Georgia", 26, bold=True,
                    color=C_WHITE, align=PP_ALIGN.RIGHT)
        # Lost text to right of bar
        if row.get("lost"):
            add_textbox(slide, LM + bar_w + Inches(0.15), y + Inches(0.1),
                        Inches(1.3), Inches(0.36),
                        row["lost"], "Calibri", 13, bold=True, color=C_RED)

    # Insight paragraph
    add_textbox(slide, LM, Inches(5.3), W - LM - RM, Inches(0.7),
                d["funnel_insight"], "Calibri", 13, italic=True, color=C_DARK_TEXT)

    add_footer(slide, d["brand_name"], 7)


# ─────────────────────────────────────────────
# SLIDE 8 — PART TWO DIVIDER
# ─────────────────────────────────────────────

def slide_08_part_two(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_DARK_BG)
    add_accent_bar(slide, dark=True)
    add_rect(slide, 0, 0, Inches(0.08), H, fill_color=C_PURPLE)

    add_textbox(slide, LM, Inches(1.8), Inches(4), Inches(0.3),
                "PART TWO", "Arial", 11, bold=True, color=C_PURPLE)
    add_textbox(slide, LM, Inches(2.25), W - LM - RM, Inches(1.2),
                "Your Business Today", "Georgia", 52, bold=True, color=C_WHITE)
    add_textbox(slide, LM, Inches(3.6), Inches(9), Inches(0.4),
                f"What we found when we audited {d['client_name']}",
                "Calibri", 18, color=C_MUTED)


# ─────────────────────────────────────────────
# SLIDE 9 — YOUR NUMBERS
# ─────────────────────────────────────────────

def slide_09_numbers(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    add_section_label(slide, "YOUR NUMBERS")
    add_title(slide, f"{d['client_name']} — current state",
              top=Inches(0.72), font_size=36)

    cards = d["number_cards"]
    card_w = Inches(3.8)
    card_h = Inches(2.1)   # taller — fills more slide height
    gap_x = Inches(0.28)
    gap_y = Inches(0.28)
    top_y = Inches(1.88)

    positions = [
        (LM,                        top_y),
        (LM + card_w + gap_x,       top_y),
        (LM + 2*(card_w + gap_x),   top_y),
        (LM,                        top_y + card_h + gap_y),
        (LM + card_w + gap_x,       top_y + card_h + gap_y),
        (LM + 2*(card_w + gap_x),   top_y + card_h + gap_y),
    ]

    for card, (cx, cy) in zip(cards, positions):
        add_rounded_rect(slide, cx, cy, card_w, card_h,
                         fill_color=C_CARD_BG, radius_pt=4,
                         line_color=RGBColor(0xE0, 0xE0, 0xE0))
        # Big value
        add_textbox(slide, cx + Inches(0.18), cy + Inches(0.14),
                    card_w - Inches(0.36), Inches(0.9),
                    str(card["value"]), "Georgia", 44, bold=True, color=C_DARK_TEXT)
        # Label
        add_textbox(slide, cx + Inches(0.18), cy + Inches(1.0),
                    card_w - Inches(0.36), Inches(0.55),
                    card["label"], "Calibri", 14, color=C_DARK_TEXT)
        # Sub-note
        add_textbox(slide, cx + Inches(0.18), cy + Inches(1.58),
                    card_w - Inches(0.36), Inches(0.38),
                    card["subnote"], "Calibri", 11, italic=True, color=C_MUTED)

    add_footer(slide, d["brand_name"], 9)


# ─────────────────────────────────────────────
# SLIDE 10 — AUDIT METHODOLOGY
# ─────────────────────────────────────────────

def slide_10_methodology(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    add_section_label(slide, "AUDIT METHODOLOGY")
    add_title(slide, "What we actually tested", top=Inches(0.72), font_size=36)

    # Intro italic line
    add_textbox(slide, LM, Inches(1.72), W - LM - RM, Inches(0.35),
                f"We didn't just ask questions — we tested your {d['term_business']} "
                f"as a real {d['term_client']} would.",
                "Calibri", 13, italic=True, color=C_DARK_TEXT)

    tests = d["audit_tests"]  # list of 6: {name, finding}
    # 2-column, 3-row grid
    col_w = Inches(5.9)
    row_h = Inches(1.3)
    gap_x = Inches(0.26)
    gap_y = Inches(0.1)
    start_y = Inches(2.18)

    for i, test in enumerate(tests):
        col = i % 2
        row = i // 2
        cx = LM + col * (col_w + gap_x)
        cy = start_y + row * (row_h + gap_y)

        add_rounded_rect(slide, cx, cy, col_w, row_h,
                         fill_color=C_CARD_BG, radius_pt=4,
                         line_color=RGBColor(0xE0, 0xE0, 0xE0))

        # Number circle
        add_circle(slide,
                   cx + Inches(0.35), cy + row_h / 2,
                   Inches(0.42), C_PURPLE,
                   str(i + 1), font_size=16)

        # Test name
        add_textbox(slide, cx + Inches(0.7), cy + Inches(0.1),
                    col_w - Inches(0.85), Inches(0.38),
                    test["name"], "Calibri", 14, bold=True, color=C_DARK_TEXT)
        # Finding
        add_textbox(slide, cx + Inches(0.7), cy + Inches(0.5),
                    col_w - Inches(0.85), Inches(0.65),
                    test["finding"], "Calibri", 12, color=C_MUTED)

    add_footer(slide, d["brand_name"], 10)


# ─────────────────────────────────────────────
# SLIDE 11 — AI READINESS SCORE
# ─────────────────────────────────────────────

def slide_11_score(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_DARK_BG)
    add_accent_bar(slide, dark=True)

    add_textbox(slide, LM, Inches(0.3), Inches(8), Inches(0.3),
                "YOUR AI READINESS SCORE", "Arial", 11, bold=True, color=C_PURPLE)

    dims = d["dimensions"]  # list of 6: {name, weight, score}
    bar_top = Inches(0.78)
    bar_h = Inches(0.52)
    bar_gap = Inches(0.1)
    bar_track_w = Inches(4.5)
    label_w = Inches(2.4)

    for i, dim in enumerate(dims):
        y = bar_top + i * (bar_h + bar_gap)
        score_frac = dim["score"] / 100.0

        # Dimension name
        add_textbox(slide, LM, y + Inches(0.08), label_w, Inches(0.36),
                    dim["name"], "Calibri", 14, color=C_WHITE)
        # Weight
        add_textbox(slide, LM + label_w, y + Inches(0.1), Inches(0.6), Inches(0.3),
                    dim["weight"], "Calibri", 11, color=C_MUTED)

        # Track (background bar)
        track_x = LM + label_w + Inches(0.7)
        add_rect(slide, track_x, y + Inches(0.14),
                 bar_track_w, Inches(0.24),
                 fill_color=RGBColor(0x3A, 0x38, 0x5A))

        # Fill bar
        score = dim["score"]
        fill_color = C_RED if score < 30 else (C_AMBER if score < 60 else C_GREEN)
        fill_w = max(Inches(0.15), bar_track_w * score_frac)
        add_rect(slide, track_x, y + Inches(0.14),
                 fill_w, Inches(0.24), fill_color=fill_color)

        # Score number
        add_textbox(slide, track_x + bar_track_w + Inches(0.1), y + Inches(0.06),
                    Inches(0.5), Inches(0.36),
                    str(score), "Calibri", 16, bold=True, color=fill_color)

    # KEY FINDINGS label — extra gap from bars
    kf_y = bar_top + 6 * (bar_h + bar_gap) + Inches(0.18)
    add_textbox(slide, LM, kf_y, Inches(5), Inches(0.28),
                "KEY FINDINGS", "Arial", 10, bold=True, color=C_PURPLE)

    findings = d["key_findings"]  # list of 5 strings
    for i, finding in enumerate(findings):
        add_textbox(slide, LM, kf_y + Inches(0.34) + i * Inches(0.42),
                    Inches(7.8), Inches(0.36),
                    "→  " + finding, "Calibri", 12, color=C_WHITE)

    # Score circle (right side)
    circle_cx = Inches(10.8)
    circle_cy = Inches(2.8)
    circle_d = Inches(2.8)
    add_circle(slide, circle_cx, circle_cy, circle_d,
               C_SCORE_RED, str(d["overall_score"]),
               font_size=72)

    # /100 label
    add_textbox(slide, circle_cx - Inches(1.4), circle_cy + Inches(0.85),
                Inches(2.8), Inches(0.4),
                "/100", "Calibri", 22, color=C_WHITE, align=PP_ALIGN.CENTER)

    # Status badge
    add_textbox(slide, circle_cx - Inches(1.6), circle_cy + Inches(1.35),
                Inches(3.2), Inches(0.5),
                d["score_status_label"], "Arial", 13, bold=True,
                color=C_RED, align=PP_ALIGN.CENTER)

    add_footer(slide, d["brand_name"], 11)


# ─────────────────────────────────────────────
# SLIDES 12–14 — DIMENSION DEEP DIVE (reusable)
# ─────────────────────────────────────────────

def slide_deep_dive(prs, d, dive, slide_number):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    # Section label
    add_section_label(slide, f"DIMENSION DEEP DIVE — {dive['score']}/100")

    # Score badge (top right)
    badge = add_rounded_rect(slide, W - RM - Inches(2.6), Inches(0.3),
                              Inches(2.6), Inches(1.1),
                              fill_color=C_SCORE_RED, radius_pt=4)
    add_textbox(slide, W - RM - Inches(2.55), Inches(0.32),
                Inches(2.5), Inches(0.62),
                f"{dive['score']}/100", "Georgia", 36, bold=True,
                color=C_WHITE, align=PP_ALIGN.CENTER)
    add_textbox(slide, W - RM - Inches(2.55), Inches(0.92),
                Inches(2.5), Inches(0.3),
                dive["subtitle"], "Calibri", 11,
                color=C_WHITE, align=PP_ALIGN.CENTER)

    # Dimension title
    add_title(slide, dive["name"], top=Inches(0.72), font_size=36,
              width=W - LM - RM - Inches(3.0))

    # 4 sub-score cards in 2×2 grid
    subs = dive["sub_scores"]  # list of 4: {score, label, note}
    sub_card_w = Inches(5.7)
    sub_card_h = Inches(0.72)
    sub_gap_x = Inches(0.28)
    sub_gap_y = Inches(0.1)
    sub_top = Inches(1.75)

    for i, sub in enumerate(subs):
        col = i % 2
        row = i // 2
        cx = LM + col * (sub_card_w + sub_gap_x)
        cy = sub_top + row * (sub_card_h + sub_gap_y)

        add_rounded_rect(slide, cx, cy, sub_card_w, sub_card_h,
                         fill_color=C_CARD_BG, radius_pt=4,
                         line_color=RGBColor(0xE0, 0xE0, 0xE0))

        # Score number
        add_textbox(slide, cx + Inches(0.12), cy + Inches(0.05),
                    Inches(0.6), Inches(0.62),
                    str(sub["score"]), "Georgia", 28, bold=True, color=C_PURPLE)
        # Label
        add_textbox(slide, cx + Inches(0.75), cy + Inches(0.06),
                    sub_card_w - Inches(0.9), Inches(0.3),
                    sub["label"], "Calibri", 13, bold=True, color=C_DARK_TEXT)
        # Note
        add_textbox(slide, cx + Inches(0.75), cy + Inches(0.36),
                    sub_card_w - Inches(0.9), Inches(0.28),
                    sub["note"], "Calibri", 11, color=C_MUTED)

    # AUDIT FINDING block
    af_top = Inches(3.42)
    af_h = Inches(1.3)
    add_rect(slide, LM, af_top, W - LM - RM, af_h,
             fill_color=RGBColor(0xFF, 0xF0, 0xF0))
    # Red left border
    add_rect(slide, LM, af_top, Inches(0.05), af_h, fill_color=C_RED)
    add_textbox(slide, LM + Inches(0.15), af_top + Inches(0.08),
                Inches(3), Inches(0.28),
                "AUDIT FINDING", "Arial", 10, bold=True, color=C_RED)
    add_textbox(slide, LM + Inches(0.15), af_top + Inches(0.36),
                W - LM - RM - Inches(0.3), Inches(0.85),
                dive["audit_finding"], "Calibri", 12, color=C_DARK_TEXT)

    # REVENUE IMPACT block
    ri_top = af_top + af_h + Inches(0.12)
    add_rect(slide, LM, ri_top, W - LM - RM, Inches(1.0),
             fill_color=RGBColor(0xF0, 0xFF, 0xF6))
    add_rect(slide, LM, ri_top, Inches(0.05), Inches(1.0), fill_color=C_GREEN)
    add_textbox(slide, LM + Inches(0.15), ri_top + Inches(0.06),
                Inches(3), Inches(0.28),
                "REVENUE IMPACT", "Arial", 10, bold=True, color=C_GREEN)
    add_textbox(slide, LM + Inches(0.15), ri_top + Inches(0.34),
                W - LM - RM - Inches(0.3), Inches(0.58),
                dive["revenue_impact"], "Calibri", 12, color=C_DARK_TEXT)

    add_footer(slide, d["brand_name"], slide_number)


# ─────────────────────────────────────────────
# SLIDE 15 — BOTTOM LINE
# ─────────────────────────────────────────────

def slide_15_bottom_line(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_DARK_BG)
    add_accent_bar(slide, dark=True)

    add_textbox(slide, LM, Inches(0.35), Inches(5), Inches(0.28),
                "THE BOTTOM LINE", "Arial", 11, bold=True, color=C_PURPLE)

    add_textbox(slide, LM, Inches(0.85), W - LM - RM, Inches(0.55),
                f"Your {d['bottom_line_biz_word']} is leaving",
                "Calibri", 30, color=C_WHITE)

    add_textbox(slide, LM, Inches(1.38), W - LM - RM, Inches(1.85),
                d["annual_leak"], "Georgia", 96, bold=True, color=C_GREEN)

    add_textbox(slide, LM, Inches(3.28), W - LM - RM, Inches(0.55),
                "per year on the table.", "Calibri", 30, color=C_WHITE)

    add_textbox(slide, LM, Inches(4.05), W - LM - RM, Inches(0.75),
                d["bottom_line_context"], "Calibri", 15, color=C_MUTED)

    add_textbox(slide, LM, Inches(5.05), W - LM - RM, Inches(0.75),
                d["closing_line"], "Georgia", 19, italic=True, color=C_WHITE)


# ─────────────────────────────────────────────
# SLIDE 16 — PART THREE DIVIDER
# ─────────────────────────────────────────────

def slide_16_part_three(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_DARK_BG)
    add_accent_bar(slide, dark=True)
    add_rect(slide, 0, 0, Inches(0.08), H, fill_color=C_PURPLE)

    add_textbox(slide, LM, Inches(1.8), Inches(4), Inches(0.3),
                "PART THREE", "Arial", 11, bold=True, color=C_PURPLE)
    add_textbox(slide, LM, Inches(2.25), W - LM - RM, Inches(1.5),
                "Your AI Revenue\nRecovery Systems", "Georgia", 52, bold=True,
                color=C_WHITE)
    add_textbox(slide, LM, Inches(4.05), Inches(9), Inches(0.4),
                f"Recommended systems custom-built for {d['client_name']}",
                "Calibri", 18, color=C_MUTED)


# ─────────────────────────────────────────────
# SLIDE 17 — HOW IT WORKS
# ─────────────────────────────────────────────

def slide_17_how_it_works(prs, d):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    add_section_label(slide, "HOW IT WORKS")
    add_title(slide, f"What changes in your {d['term_client']} journey",
              top=Inches(0.72), font_size=36)

    before_steps = [
        "Lead calls in",
        "Voicemail / wait",
        "One email reply",
        "No follow-up",
        d["before_dropoff_pct"],
        "Never return",
    ]

    col_w = (W - LM - RM) / 6 - Inches(0.07)
    step_h = Inches(1.3)   # significantly taller

    # BEFORE label
    add_textbox(slide, LM, Inches(1.88), Inches(6), Inches(0.3),
                "BEFORE (current)", "Arial", 11, bold=True, color=C_RED)
    step_top_before = Inches(2.25)

    for i, step in enumerate(before_steps):
        cx = LM + i * (col_w + Inches(0.07))
        add_rounded_rect(slide, cx, step_top_before, col_w, step_h,
                         fill_color=RGBColor(0xF0, 0xEE, 0xEE), radius_pt=3)
        add_vcenter_text(slide, cx + Inches(0.05), step_top_before,
                         col_w - Inches(0.1), step_h,
                         step, "Calibri", 12, color=C_DARK_TEXT)
        if i < 5:
            add_textbox(slide, cx + col_w, step_top_before + Inches(0.48),
                        Inches(0.07), Inches(0.3),
                        "—", "Calibri", 11, color=C_MUTED, align=PP_ALIGN.CENTER)

    # AFTER label
    add_textbox(slide, LM, Inches(3.75), Inches(6), Inches(0.3),
                "AFTER (with AI systems)", "Arial", 11, bold=True, color=C_GREEN)
    step_top_after = Inches(4.12)

    after_steps = [
        "Lead calls in",
        "AI answers in\n<3 rings, 24/7",
        "7-touch nurture\nfor non-bookers",
        "Smart reminders\nslash no-shows",
        d["after_step_5"],
        "Reviews, referrals,\nreactivation",
    ]

    for i, step in enumerate(after_steps):
        cx = LM + i * (col_w + Inches(0.07))
        add_rounded_rect(slide, cx, step_top_after, col_w, step_h,
                         fill_color=RGBColor(0xE8, 0xF8, 0xF2), radius_pt=3)
        add_vcenter_text(slide, cx + Inches(0.05), step_top_after,
                         col_w - Inches(0.1), step_h,
                         step, "Calibri", 11, color=C_DARK_TEXT)
        if i < 5:
            add_textbox(slide, cx + col_w, step_top_after + Inches(0.48),
                        Inches(0.07), Inches(0.3),
                        "—", "Calibri", 11, color=C_MUTED, align=PP_ALIGN.CENTER)

    # Closing line
    add_textbox(slide, LM, Inches(5.65), W - LM - RM, Inches(0.45),
                d["how_it_works_closing"], "Calibri", 13, italic=True, color=C_GREEN)

    add_footer(slide, d["brand_name"], 17)


# ─────────────────────────────────────────────
# SYSTEM SLIDES (18+, one per system)
# ─────────────────────────────────────────────

def slide_system(prs, d, sys_data, idx, total, slide_number):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    # Section label
    add_textbox(slide, LM, Inches(0.18), Inches(5), Inches(0.3),
                f"SYSTEM {idx} OF {total}", "Arial", 11, bold=True, color=C_PURPLE)

    # System name — font size auto-scales with name length
    name_len = len(sys_data["name"])
    title_size = 34 if name_len <= 20 else (28 if name_len <= 30 else 22)
    add_title(slide, sys_data["name"], top=Inches(0.55), font_size=title_size,
              width=Inches(7.6))

    # Phase badge
    badge = add_rounded_rect(slide, LM, Inches(1.28), Inches(2.6), Inches(0.36),
                              fill_color=C_DARK_TEXT, radius_pt=4)
    set_shape_text(badge, sys_data["label"], "Arial", 10, bold=True,
                   color=C_WHITE, align=PP_ALIGN.CENTER)
    badge.text_frame.vertical_anchor = __import__(
        'pptx.enum.text', fromlist=['MSO_ANCHOR']).MSO_ANCHOR.MIDDLE

    # Description
    add_textbox(slide, LM, Inches(1.78), Inches(7.5), Inches(0.95),
                sys_data["description"], "Calibri", 13, color=C_DARK_TEXT)

    # HOW IT WORKS label
    add_textbox(slide, LM, Inches(2.82), Inches(4), Inches(0.25),
                "HOW IT WORKS", "Arial", 10, bold=True, color=C_PURPLE)

    # 5 step cards
    steps = sys_data["steps"]  # list of 5 strings
    step_w = (Inches(7.5) - Inches(0.08) * 4) / 5
    step_h = Inches(1.1)
    step_top = Inches(3.12)

    for i, step in enumerate(steps):
        cx = LM + i * (step_w + Inches(0.08))
        add_rounded_rect(slide, cx, step_top, step_w, step_h,
                         fill_color=C_HIGHLIGHT, radius_pt=4)
        # Number circle
        add_circle(slide, cx + step_w / 2, step_top + Inches(0.28),
                   Inches(0.38), C_PURPLE, str(i + 1), font_size=13)
        # Step text vertically centred in bottom half of card
        add_vcenter_text(slide, cx + Inches(0.06), step_top + Inches(0.52),
                         step_w - Inches(0.12), Inches(0.52),
                         step, "Calibri", 10, color=C_DARK_TEXT)

    # AUDIT FINDING block
    af_top = Inches(4.38)
    af_h = Inches(1.08)
    add_rect(slide, LM, af_top, Inches(7.5), af_h,
             fill_color=RGBColor(0xFF, 0xF0, 0xF0))
    add_rect(slide, LM, af_top, Inches(0.05), af_h, fill_color=C_RED)
    add_textbox(slide, LM + Inches(0.15), af_top + Inches(0.06),
                Inches(3), Inches(0.25),
                "AUDIT FINDING", "Arial", 10, bold=True, color=C_RED)
    add_textbox(slide, LM + Inches(0.15), af_top + Inches(0.32),
                Inches(7.2), Inches(0.7),
                sys_data["audit_finding"], "Calibri", 12, color=C_DARK_TEXT)

    # ROI callout — sits immediately below audit finding
    roi_top = af_top + af_h + Inches(0.1)
    add_rounded_rect(slide, LM, roi_top, Inches(7.5), Inches(0.68),
                     fill_color=C_DARK_TEXT, radius_pt=4)
    add_textbox(slide, LM + Inches(0.2), roi_top + Inches(0.1),
                Inches(4.5), Inches(0.28),
                "PROJECTED ROI", "Arial", 10, bold=True, color=C_MUTED)
    add_textbox(slide, LM + Inches(4.0), roi_top + Inches(0.06),
                Inches(3.3), Inches(0.48),
                f"{sys_data['roi']} return on investment",
                "Calibri", 18, bold=True, color=C_GREEN, align=PP_ALIGN.RIGHT)

    # RIGHT COLUMN — 5 equal cards, top to bottom, uniform spacing
    rc_x = Inches(8.35)
    rc_w = W - rc_x - RM
    rc_top = Inches(0.18)
    rc_bottom = Inches(6.88)
    n_cards = 5
    gap = Inches(0.1)
    card_h = (rc_bottom - rc_top - gap * (n_cards - 1)) / n_cards

    cards_rc = [
        (sys_data["metric_label"],  str(sys_data["metric_value"]), C_DARK_TEXT, 40, C_CARD_BG),
        ("Monthly Revenue",          sys_data["monthly_revenue"],   C_GREEN,     28, C_CARD_BG),
        ("Monthly Cost",             sys_data["monthly_cost"],       C_AMBER,     28, C_CARD_BG),
        ("Setup Fee",                sys_data["setup_fee"],          C_DARK_TEXT, 28, C_CARD_BG),
        ("ROI",                      sys_data["roi"],                C_GREEN,     28, C_DARK_TEXT),
    ]

    for i, (label, value, val_color, val_size, bg) in enumerate(cards_rc):
        cy = rc_top + i * (card_h + gap)
        add_rounded_rect(slide, rc_x, cy, rc_w, card_h,
                         fill_color=bg, radius_pt=4,
                         line_color=None if bg == C_DARK_TEXT else RGBColor(0xE0, 0xE0, 0xE0))
        label_color = C_MUTED if bg == C_CARD_BG else RGBColor(0x99, 0x99, 0xBB)
        add_textbox(slide, rc_x + Inches(0.12), cy + Inches(0.1),
                    rc_w - Inches(0.24), Inches(0.28),
                    label, "Calibri", 10, color=label_color)
        add_textbox(slide, rc_x + Inches(0.12), cy + Inches(0.38),
                    rc_w - Inches(0.24), card_h - Inches(0.48),
                    value, "Georgia", val_size, bold=True, color=val_color)

    add_footer(slide, d["brand_name"], slide_number)


# ─────────────────────────────────────────────
# REVENUE SUMMARY SLIDE
# ─────────────────────────────────────────────

def slide_revenue_summary(prs, d, slide_number):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_DARK_BG)
    add_accent_bar(slide, dark=True)

    add_textbox(slide, LM, Inches(0.3), Inches(8), Inches(0.28),
                "REVENUE RECOVERY SUMMARY", "Arial", 11, bold=True, color=C_PURPLE)

    # Big number
    add_textbox(slide, LM, Inches(0.65), Inches(8), Inches(1.5),
                d["total_monthly_revenue"], "Georgia", 80, bold=True, color=C_GREEN)
    add_textbox(slide, LM, Inches(2.1), Inches(8), Inches(0.4),
                f"additional revenue per month  ·  {d['annual_recovery']} per year",
                "Calibri", 16, color=C_MUTED)

    # 4 KPI boxes
    kpis = [
        ("Monthly Investment", d["total_monthly_cost"],     C_WHITE),
        ("Net Monthly Gain",   d["net_monthly_gain"],       C_GREEN),
        ("Annual Recovery",    d["annual_recovery"],        C_GREEN),
        ("Overall ROI",        d["overall_roi"] + " return", C_PURPLE),
    ]
    kpi_w = Inches(3.0)
    kpi_h = Inches(1.0)
    kpi_gap = Inches(0.14)
    kpi_top = Inches(2.62)

    for i, (label, value, color) in enumerate(kpis):
        cx = LM + i * (kpi_w + kpi_gap)
        add_rounded_rect(slide, cx, kpi_top, kpi_w, kpi_h,
                         fill_color=C_DARK_CARD, radius_pt=4)
        add_textbox(slide, cx + Inches(0.12), kpi_top + Inches(0.06),
                    kpi_w - Inches(0.24), Inches(0.28),
                    label, "Calibri", 10, color=C_MUTED)
        add_textbox(slide, cx + Inches(0.12), kpi_top + Inches(0.35),
                    kpi_w - Inches(0.24), Inches(0.55),
                    value, "Georgia", 26, bold=True, color=color)

    # Table
    systems = d["systems"]
    table_top = Inches(3.78)
    col_widths = [Inches(4.0), Inches(2.0), Inches(1.6), Inches(1.6), Inches(1.6)]
    headers = ["AI System", "Revenue/mo", "Cost/mo", "Setup", "ROI"]
    row_h = Inches(0.46)

    # Header row
    x = LM
    for j, (header, cw) in enumerate(zip(headers, col_widths)):
        cell = add_rect(slide, x, table_top, cw, row_h,
                        fill_color=C_PURPLE)
        add_textbox(slide, x + Inches(0.08), table_top + Inches(0.1),
                    cw - Inches(0.16), Inches(0.26),
                    header, "Calibri", 12, bold=True, color=C_WHITE)
        x += cw

    # Data rows
    for r, sys in enumerate(systems):
        row_y = table_top + (r + 1) * row_h
        bg = C_DARK_CARD if r % 2 == 0 else RGBColor(0x24, 0x22, 0x3A)
        row_data = [
            (sys["name"],           C_WHITE,  False),
            (sys["monthly_revenue"], C_GREEN, False),
            (sys["monthly_cost"],    C_WHITE, False),
            (sys["setup_fee"],       C_WHITE, False),
            (sys["roi"],             C_PURPLE, False),
        ]
        x = LM
        for j, ((text, color, bold), cw) in enumerate(zip(row_data, col_widths)):
            add_rect(slide, x, row_y, cw, row_h, fill_color=bg)
            add_textbox(slide, x + Inches(0.08), row_y + Inches(0.1),
                        cw - Inches(0.16), Inches(0.26),
                        text, "Calibri", 12, bold=bold, color=color)
            x += cw

    # Total row
    total_y = table_top + (len(systems) + 1) * row_h
    total_data = [
        ("TOTAL",                   C_WHITE,  True),
        (d["total_monthly_revenue"] + "/mo", C_GREEN, True),
        (d["total_monthly_cost"] + "/mo",    C_WHITE, True),
        (d["total_setup_fees"],     C_WHITE,  True),
        (d["overall_roi"],          C_PURPLE, True),
    ]
    x = LM
    for (text, color, bold), cw in zip(total_data, col_widths):
        add_rect(slide, x, total_y, cw, row_h, fill_color=C_DARK_CARD)
        add_textbox(slide, x + Inches(0.08), total_y + Inches(0.1),
                    cw - Inches(0.16), Inches(0.26),
                    text, "Calibri", 12, bold=bold, color=color)
        x += cw

    add_footer(slide, d["brand_name"], slide_number)


# ─────────────────────────────────────────────
# BEFORE / AFTER TRANSFORMATION
# ─────────────────────────────────────────────

def slide_transformation(prs, d, slide_number):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    add_section_label(slide, "THE TRANSFORMATION")
    add_title(slide, "Before vs. After AI Systems", top=Inches(0.72), font_size=36)

    rows = d["transformation_rows"]  # list of 7: {label, current, with_ai}
    table_top = Inches(1.85)
    row_h = Inches(0.5)
    label_w = Inches(4.5)
    val_w = Inches(3.8)
    header_h = Inches(0.45)

    # Header
    add_rect(slide, LM, table_top, label_w + val_w * 2 + Inches(0.04),
             header_h, fill_color=C_DARK_TEXT)
    add_textbox(slide, LM + label_w + Inches(0.1), table_top + Inches(0.08),
                val_w - Inches(0.2), Inches(0.28),
                "CURRENT STATE", "Arial", 10, bold=True, color=C_RED,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, LM + label_w + val_w + Inches(0.14), table_top + Inches(0.08),
                val_w - Inches(0.2), Inches(0.28),
                "WITH AI SYSTEMS", "Arial", 10, bold=True, color=C_GREEN,
                align=PP_ALIGN.CENTER)

    for i, row in enumerate(rows):
        ry = table_top + header_h + i * row_h
        bg = C_HIGHLIGHT if i % 2 == 0 else C_WHITE
        add_rect(slide, LM, ry, label_w + val_w * 2 + Inches(0.04), row_h,
                 fill_color=bg)
        add_textbox(slide, LM + Inches(0.12), ry + Inches(0.12),
                    label_w - Inches(0.24), Inches(0.26),
                    row["label"], "Calibri", 14, color=C_DARK_TEXT)
        add_textbox(slide, LM + label_w + Inches(0.1), ry + Inches(0.1),
                    val_w - Inches(0.2), Inches(0.3),
                    row["current"], "Calibri", 14, bold=True, color=C_RED,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, LM + label_w + val_w + Inches(0.14), ry + Inches(0.1),
                    val_w - Inches(0.2), Inches(0.3),
                    row["with_ai"], "Calibri", 14, bold=True, color=C_GREEN,
                    align=PP_ALIGN.CENTER)

    add_footer(slide, d["brand_name"], slide_number)


# ─────────────────────────────────────────────
# GUARANTEE
# ─────────────────────────────────────────────

def slide_guarantee(prs, d, slide_number):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_DARK_BG)

    # Teal border lines top and bottom
    add_rect(slide, Inches(0.5), Inches(0.25), W - Inches(1.0), Inches(0.06),
             fill_color=C_TEAL)
    add_rect(slide, Inches(0.5), H - Inches(0.31), W - Inches(1.0), Inches(0.06),
             fill_color=C_TEAL)

    # Inner card
    add_rounded_rect(slide, Inches(0.5), Inches(0.35),
                     W - Inches(1.0), H - Inches(0.7),
                     fill_color=C_DARK_CARD, radius_pt=6)

    add_textbox(slide, 0, Inches(0.52), W, Inches(0.3),
                "OUR GUARANTEE", "Arial", 11, bold=True, color=C_PURPLE,
                align=PP_ALIGN.CENTER)

    # Headline
    add_textbox(slide, Inches(1.0), Inches(0.88), W - Inches(2.0), Inches(1.65),
                d["guarantee_headline"],
                "Georgia", 44, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # Italic subline
    add_textbox(slide, Inches(1.0), Inches(2.6), W - Inches(2.0), Inches(0.48),
                "Or we work for free until you do.",
                "Georgia", 24, italic=True, color=C_GREEN, align=PP_ALIGN.CENTER)

    # Math box — 3 equally-spaced rows inside, fills bottom ~40% of card
    math_y = Inches(3.22)
    math_h = H - Inches(0.7) - math_y + Inches(0.35) - Inches(0.1)
    add_rounded_rect(slide, Inches(1.5), math_y, W - Inches(3.0), math_h,
                     fill_color=RGBColor(0x28, 0x26, 0x42), radius_pt=4)

    # THE MATH label centred at top of box
    add_textbox(slide, Inches(1.7), math_y + Inches(0.14), W - Inches(3.4), Inches(0.28),
                "THE MATH", "Arial", 11, bold=True, color=C_MUTED,
                align=PP_ALIGN.CENTER)

    # Row spacing: 3 content rows evenly across remaining height
    row_h = (math_h - Inches(0.52)) / 3
    for idx, (label, value) in enumerate([
        (f"Your 12-month investment:", d["guarantee_investment"]),
        (f"Projected revenue in first 90 days:", d["guarantee_90day"]),
        (f"That's {d['guarantee_surplus']} surplus — before you've even used 3 months.", ""),
    ]):
        ry = math_y + Inches(0.48) + idx * row_h
        if idx < 2:
            # Label + green value on same line
            add_textbox(slide, Inches(1.7), ry, W - Inches(3.4), row_h,
                        f"{label}  {value}", "Calibri", 17,
                        color=C_WHITE if idx < 2 else C_GREEN)
        else:
            # Surplus line in green
            add_textbox(slide, Inches(1.7), ry, W - Inches(3.4), row_h,
                        label, "Calibri", 17, color=C_GREEN)


# ─────────────────────────────────────────────
# IMPLEMENTATION TIMELINE
# ─────────────────────────────────────────────

def slide_implementation(prs, d, slide_number):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_LIGHT_BG)
    add_accent_bar(slide, dark=False)

    add_section_label(slide, "IMPLEMENTATION")
    add_title(slide, "Live in 14 days. Results in 30.", top=Inches(0.72), font_size=36)

    phases = d["implementation_phases"]  # list of 3: {period, title, items[]}
    col_w = Inches(3.85)
    col_h = Inches(4.2)
    col_top = Inches(1.7)
    gap = Inches(0.22)

    phase_colors = [C_PURPLE, C_GREEN, RGBColor(0x00, 0x88, 0xCC)]

    for i, phase in enumerate(phases):
        cx = LM + i * (col_w + gap)
        add_rounded_rect(slide, cx, col_top, col_w, col_h,
                         fill_color=C_HIGHLIGHT, radius_pt=6)
        # Period label
        add_textbox(slide, cx + Inches(0.15), col_top + Inches(0.12),
                    col_w - Inches(0.3), Inches(0.25),
                    phase["period"], "Arial", 10, bold=True,
                    color=phase_colors[i])
        # Title
        add_textbox(slide, cx + Inches(0.15), col_top + Inches(0.4),
                    col_w - Inches(0.3), Inches(0.38),
                    phase["title"], "Georgia", 17, bold=True, color=C_DARK_TEXT)
        # Items
        for j, item in enumerate(phase["items"]):
            add_textbox(slide, cx + Inches(0.15), col_top + Inches(0.9) + j * Inches(0.42),
                        col_w - Inches(0.3), Inches(0.38),
                        "→  " + item, "Calibri", 12, color=C_DARK_TEXT)

    # Closing line
    add_textbox(slide, LM, Inches(6.1), W - LM - RM, Inches(0.35),
                d["implementation_closing"], "Calibri", 13, italic=True,
                color=C_MUTED, align=PP_ALIGN.CENTER)

    add_footer(slide, d["brand_name"], slide_number)


# ─────────────────────────────────────────────
# NEXT STEPS
# ─────────────────────────────────────────────

def slide_next_steps(prs, d, slide_number):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_DARK_BG)
    add_accent_bar(slide, dark=True)

    add_textbox(slide, LM, Inches(0.32), Inches(5), Inches(0.28),
                "NEXT STEPS", "Arial", 11, bold=True, color=C_PURPLE)
    add_textbox(slide, LM, Inches(0.65), W - LM - RM, Inches(0.58),
                "Select your bots. We'll start building.",
                "Georgia", 32, bold=True, color=C_WHITE)

    steps = d["next_steps"]
    card_w = Inches(2.9)
    card_h = Inches(4.2)
    gap = Inches(0.24)
    top_y = Inches(1.45)

    for i, step in enumerate(steps):
        cx = LM + i * (card_w + gap)
        add_rounded_rect(slide, cx, top_y, card_w, card_h,
                         fill_color=C_DARK_CARD, radius_pt=6)
        # Big number
        add_textbox(slide, cx + Inches(0.18), top_y + Inches(0.15),
                    card_w - Inches(0.36), Inches(1.0),
                    str(i + 1), "Georgia", 60, bold=True, color=C_PURPLE)
        # Title
        add_textbox(slide, cx + Inches(0.18), top_y + Inches(1.18),
                    card_w - Inches(0.36), Inches(0.46),
                    step["title"], "Calibri", 14, bold=True, color=C_WHITE)
        # Description
        add_textbox(slide, cx + Inches(0.18), top_y + Inches(1.68),
                    card_w - Inches(0.36), Inches(1.8),
                    step["description"], "Calibri", 12, color=C_MUTED)

    # CTA pinned near bottom
    add_textbox(slide, LM, Inches(5.85), W - LM - RM, Inches(0.32),
                f"{d['cta_url']}  ·  {d['cta_label']}",
                "Calibri", 13, color=C_MUTED, align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────
# CLOSING SLIDE
# ─────────────────────────────────────────────

def slide_closing(prs, d, slide_number):
    slide = new_slide(prs)
    add_rect(slide, 0, 0, W, H, fill_color=C_DARK_BG)
    add_accent_bar(slide, dark=True)

    add_textbox(slide, LM, Inches(1.8), W - LM - RM, Inches(1.4),
                d["closing_urgency_line"],
                "Georgia", 32, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    add_textbox(slide, LM, Inches(3.35), W - LM - RM, Inches(0.8),
                d["closing_cta"],
                "Georgia", 44, bold=True, color=C_GREEN, align=PP_ALIGN.CENTER)

    add_textbox(slide, LM, H - Inches(0.75), W - LM - RM, Inches(0.3),
                d["brand_line"],
                "Calibri", 14, color=C_MUTED, align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────
# MAIN GENERATE FUNCTION
# ─────────────────────────────────────────────

def generate(data, output_path="output.pptx"):
    """
    Build the complete Revenue Leak Audit deck from a data dictionary.
    See DATA_SCHEMA below for all required keys.
    """
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    # Slides 1–11 (fixed)
    slide_01_cover(prs, data)
    slide_02_agenda(prs, data)
    slide_03_part_one(prs, data)
    slide_04_journey(prs, data)
    slide_05_stats(prs, data)
    slide_06_lifecycle(prs, data)
    slide_07_funnel(prs, data)
    slide_08_part_two(prs, data)
    slide_09_numbers(prs, data)
    slide_11_score(prs, data)

    # Deep dive slides (variable count)
    current_slide = 11
    for dive in data["deep_dives"]:
        slide_deep_dive(prs, data, dive, current_slide)
        current_slide += 1

    # Slide 15 equivalent (bottom line)
    slide_15_bottom_line(prs, data)
    current_slide += 1

    # Part three divider
    slide_16_part_three(prs, data)
    current_slide += 1

    # How it works
    slide_17_how_it_works(prs, data)
    current_slide += 1

    # System slides (variable count)
    systems = data["systems"]
    total_systems = len(systems)
    for i, sys in enumerate(systems):
        slide_system(prs, data, sys, i + 1, total_systems, current_slide)
        current_slide += 1

    # Summary, transformation, guarantee, implementation, next steps, closing
    slide_revenue_summary(prs, data, current_slide); current_slide += 1
    slide_transformation(prs, data, current_slide);  current_slide += 1
    slide_guarantee(prs, data, current_slide);        current_slide += 1
    slide_implementation(prs, data, current_slide);   current_slide += 1
    slide_next_steps(prs, data, current_slide);       current_slide += 1
    slide_closing(prs, data, current_slide)

    prs.save(output_path)
    print(f"✓ Saved: {output_path}  ({current_slide} slides)")
    return output_path


# ─────────────────────────────────────────────
# DATA_SCHEMA — full example (Law Biz / Legal)
# ─────────────────────────────────────────────

EXAMPLE_DATA = {
    # ── Identity ──────────────────────────────
    "brand_name":    "Booked Solid AI",
    "client_name":   "Law Biz",
    "prepared_for":  "the Directors",
    "date":          "15 April 2026",
    "cta_url":       "bookedsolidai.com.au/your-plan",
    "cta_label":     "Your personalised bot selector is ready",
    "brand_line":    "Booked Solid AI  ·  bookedsolidai.com.au",

    # ── Industry language ─────────────────────
    "term_client":       "prospect",
    "term_client_pl":    "prospects",
    "term_engagement":   "consultation",
    "term_ongoing":      "case",
    "term_completing":   "closing the case",
    "term_practitioner": "attorney",
    "term_business":     "firm",
    "bottom_line_biz_word": "firm",

    # ── Agenda ────────────────────────────────
    "agenda_01_desc": "The legal prospect journey and where cases are lost",

    # ── Part One content ──────────────────────
    "part_one_title":    "Your Industry &\nThe Client Journey",
    "part_one_subtitle": "Understanding where revenue leaks in a law firm",

    # ── Journey (Slide 4) ─────────────────────
    "journey_label":   "THE PROSPECTIVE CLIENT JOURNEY",
    "journey_title":   "How a legal prospect finds you — and where they drop off",
    "journey_stages": [
        {"name": "AWARENESS",     "desc": "Google, referral,\nword of mouth",       "pct_lost": None,  "reason": ""},
        {"name": "ENQUIRY",       "desc": "Calls, web form,\nlive chat",            "pct_lost": "42%", "reason": "No response / too slow"},
        {"name": "RESPONSE",      "desc": "Firm responds",                          "pct_lost": "25%", "reason": "Never followed up"},
        {"name": "CONSULTATION",  "desc": "Free case\nevaluation",                  "pct_lost": None,  "reason": ""},
        {"name": "RETAINER",      "desc": "Signs and pays\nretainer",              "pct_lost": "45%", "reason": "Went with faster competitor"},
        {"name": "LIFETIME",      "desc": "Referrals, reviews,\nrepeat matters",   "pct_lost": "85%", "reason": "No referral or review ask"},
    ],
    "journey_closing": "Every red zone is a case fee you generated interest for — but never collected.",

    # ── Industry Stats (Slide 5) ──────────────
    "stat_section_label": "LEGAL INDUSTRY DATA",
    "industry_stats": [
        {"number": "26%",    "label": "of law firms don't respond to online leads at all — and 39% take more than 2 hours, handing cases to faster competitors", "source": "Hennessey Digital Lead Form Response Study, 2025"},
        {"number": "42%",    "label": "of potential client enquiries arrive outside business hours — evenings, weekends, holidays — and go completely unanswered", "source": "Ruby Receptionists, 2025"},
        {"number": "14%",    "label": "average lead-to-retainer conversion across law firms — top-performing firms achieve 40–50% by fixing speed and follow-up", "source": "Clio Legal Trends Report, 2025"},
        {"number": "$200k+", "label": "lost annually by the average multi-attorney firm to unanswered calls and slow response, before any marketing spend is wasted", "source": "Lead Docket Case Studies, 2025"},
    ],

    # ── Lifecycle (Slide 6) ───────────────────
    "lifecycle_title":   "A prospect isn't a one-time transaction — it's a lifecycle",
    "lifecycle_intro":   "Most firms only focus on generating new enquiries. But sustainable revenue comes from managing the full client lifecycle — and most practices have no systems for the stages marked below.",
    "lifecycle_stages": [
        {"name": "ATTRACT",  "desc": "Google Ads, SEO,\nreferrals bring enquiries",   "status": "WORKING"},
        {"name": "CAPTURE",  "desc": "Answer calls,\nrespond to web forms fast",      "status": "WORKING"},
        {"name": "NURTURE",  "desc": "Follow up prospects\nwho didn't retain",        "status": "NO SYSTEM"},
        {"name": "CONVERT",  "desc": "Book consultation,\ncollect intake info",       "status": "MANUAL"},
        {"name": "RETAIN",   "desc": "Keep clients updated\nthrough case lifecycle",  "status": "NO SYSTEM"},
        {"name": "GROW",     "desc": "Reviews, referrals\nfrom closed cases",         "status": "NO SYSTEM"},
    ],
    "lifecycle_summary": "4 of 6 lifecycle stages have no system. You're paying to attract prospects, then losing them to faster-responding competitors.",

    # ── Funnel (Slide 7) ─────────────────────
    "funnel_rows": [
        {"label": "Monthly Enquiries (calls + web + referrals)", "value": 15,  "lost": "-5 lost"},
        {"label": "Answered / Responded To (within 1hr)",        "value": 10,  "lost": "-7 lost"},
        {"label": "Actually Followed Up (if didn't retain)",      "value": 3,   "lost": "-2 lost"},
        {"label": "Booked Free Consultation",                     "value": 1,   "lost": ""},
        {"label": "Signed Retainer",                              "value": 1,   "lost": None},
    ],
    "funnel_insight": "14 of 15 enquiries never result in a signed retainer. That's 93% of marketing spend generating zero return — not because of bad leads, but because of missing systems.",

    # ── Your Numbers (Slide 9) ────────────────
    "number_cards": [
        {"value": "15",    "label": "Monthly Enquiries",                    "subnote": "Calls, web forms, CRM, referrals"},
        {"value": "~8",    "label": "Missed / Unanswered\nCalls per Month", "subnote": "No after-hours coverage"},
        {"value": "6%",    "label": "Close Rate\n(enquiry → retainer)",     "subnote": "1 client from 15 enquiries"},
        {"value": "$100",  "label": "Average First\nConsultation Fee",      "subnote": "Initial case evaluation"},
        {"value": "$12,000","label": "Avg Case Value",                      "subnote": "Avg deal / contract value"},
        {"value": "44%",   "label": "Retainer\nConversion Rate",            "subnote": "56% of consultations don't convert"},
    ],

    # ── Audit Methodology (Slide 10) ──────────
    "audit_tests": [
        {"name": "Phone Call Test (business hours)",  "finding": "Called at 11am Wednesday. Rang 7 times.\nVoicemail — callback took 2+ hours"},
        {"name": "Phone Call Test (after hours)",     "finding": "Called at 6:30pm Thursday.\nVoicemail — no callback until next morning"},
        {"name": "Web Enquiry Test",                  "finding": "Submitted web form at 10am Tuesday.\nEmail reply 4 hours later, no phone call"},
        {"name": "CRM Lead Flow Test",                "finding": "Tested CRM intake flow end-to-end.\nWorks but no confirmation SMS sent"},
        {"name": "Follow-Up Persistence",             "finding": "After web form, we didn't respond.\nZero follow-up attempts received"},
        {"name": "Google Review Audit",               "finding": "Compared profile vs top 5 local competitors.\nFewer reviews than top local law firms"},
    ],

    # ── AI Readiness Score (Slide 11) ─────────
    "overall_score": 24,
    "score_status_label": "CRITICAL — IMMEDIATE\nACTION NEEDED",
    "dimensions": [
        {"name": "Speed to Lead",        "weight": "25%", "score": 16},
        {"name": "Follow-Up Systems",    "weight": "20%", "score": 8},
        {"name": "Pipeline Visibility",  "weight": "15%", "score": 45},
        {"name": "Client Communication", "weight": "15%", "score": 35},
        {"name": "Post-Case Nurture",    "weight": "15%", "score": 20},
        {"name": "Automation Maturity",  "weight": "10%", "score": 35},
    ],
    "key_findings": [
        "~8 calls/month going unanswered with no after-hours coverage whatsoever",
        "Only 1–2 follow-up touches after initial contact — leads fall through the cracks",
        "No automated follow-up after unconverted web enquiries — leads simply abandoned",
        "Fewer Google reviews than top local competing law firms",
        "No system for re-engaging past clients or generating referrals post-case",
    ],

    # ── Deep Dives (one per dimension < 40) ───
    "deep_dives": [
        {
            "score": 16,
            "name": "Speed to Lead",
            "subtitle": "How fast you respond",
            "sub_scores": [
                {"score": 15, "label": "Phone answer rate (business hrs)", "note": "~60% answered, long ring times"},
                {"score": 0,  "label": "After-hours handling",             "note": "None — voicemail only"},
                {"score": 15, "label": "Web enquiry response time",        "note": "2–4 hours, email only"},
                {"score": 20, "label": "CRM follow-up automation",         "note": "Manual only, no sequences"},
            ],
            "audit_finding": "We called Law Biz during business hours and reached voicemail. Callback took 2+ hours. After-hours test: voicemail, no callback until the following morning. Legal prospects contact multiple firms — the first to respond wins 79% of the time.",
            "revenue_impact": "With ~8 missed calls per month, recovering 50% at a 20% retainer rate = 1 additional case per month = $12,000/month in case value. Even one recovered case covers the entire AI system investment.",
        },
        {
            "score": 8,
            "name": "Follow-Up Systems",
            "subtitle": "Follow-up persistence",
            "sub_scores": [
                {"score": 0,  "label": "Follow-up attempts",          "note": "1–2 touches maximum"},
                {"score": 0,  "label": "Unconverted lead nurturing",   "note": "None — leads abandoned"},
                {"score": 20, "label": "Channels used for follow-up",  "note": "Email only, no SMS"},
                {"score": 15, "label": "Cancellation recovery",        "note": "Manual, inconsistent"},
            ],
            "audit_finding": "After our web enquiry test, we received one email, then silence. No SMS follow-up, no multi-touch sequence, no phone callback. With approximately 12 unconverted enquiries per month, you're abandoning thousands in potential retainer revenue without a single follow-up.",
            "revenue_impact": "A 7-touch automated nurture sequence typically converts 20–30% of unconverted legal leads. At 12 unconverted enquiries, that's 2–4 extra retained cases per month = $24,000–$48,000/month from leads already paid for.",
        },
        {
            "score": 20,
            "name": "Post-Case Nurture",
            "subtitle": "After the case closes",
            "sub_scores": [
                {"score": 5,  "label": "Client re-engagement system", "note": "None — clients self-manage"},
                {"score": 10, "label": "Review collection",            "note": "Rarely asked, no system"},
                {"score": 0,  "label": "Referral program",             "note": "None"},
                {"score": 20, "label": "Lapsed client reactivation",   "note": "Occasional email, no system"},
            ],
            "audit_finding": "Law Biz has no automated review requests, no referral incentive program, no check-in after case closure, and no reactivation campaign for past clients. Legal referrals are the highest-quality, lowest-cost lead source — and there's no system to generate them.",
            "revenue_impact": "Reactivating 5% of past clients = 2–3 return matters per month = $24,000–$36,000/month. Reviews driving organic search = 3–5 new enquiries/month. This dimension alone could recover $5,000+/month.",
        },
    ],

    # ── Bottom Line (Slide 15) ────────────────
    "annual_leak":          "$60,480",
    "bottom_line_context":  "That's $5,040 every month in recoverable revenue from missed calls, abandoned leads, prospects going to faster competitors, a dormant database of past clients, and a reputation falling behind local law firms.",
    "closing_line":         "The question isn't whether you can afford AI systems.\nIt's whether you can afford not to have them.",

    # ── How It Works (Slide 17) ───────────────
    "before_dropoff_pct":     "56% drop off",
    "after_step_5":           "Retention bot\nkeeps cases moving",
    "how_it_works_closing":   "Every stage is automated. Your team focuses on winning cases — AI handles the rest.",

    # ── Systems ───────────────────────────────
    "systems": [
        {
            "name":            "AI Receptionist",
            "label":           "PHASE 1 — CORE",
            "metric_label":    "recovered calls →\nretainers/mo",
            "metric_value":    "4",
            "monthly_revenue": "$4,500",
            "monthly_cost":    "$1,000",
            "setup_fee":       "$2,000",
            "roi":             "4.5x",
            "description":     "Your AI-powered intake desk that never misses a call. Answers within 3 rings 24/7, qualifies the legal matter using natural conversation, checks CRM availability, and books consultations — all without human intervention. After hours, captures every enquiry and triggers instant SMS + email confirmation.",
            "steps":           ["Prospect calls\n(any hour)", "AI qualifies\nmatter type", "Books consultation,\nsends SMS", "Intake form\nsent automatically", "Attorney alerted\nwith full summary"],
            "audit_finding":   "We called Law Biz during business hours and reached voicemail with a 2-hour callback. After hours: voicemail until next morning. ~8 calls/month going unanswered — each one a potential $12,000 case.",
        },
        {
            "name":            "Lead Nurture Bot",
            "label":           "PHASE 2 — NURTURE",
            "metric_label":    "extra retainers from\nunconverted leads/mo",
            "metric_value":    "3",
            "monthly_revenue": "$3,600",
            "monthly_cost":    "$600",
            "setup_fee":       "$1,500",
            "roi":             "6.0x",
            "description":     "A 7-touch automated follow-up sequence for every enquiry that doesn't retain on first contact. Uses SMS + email over 21 days, handles objections around cost and timing, includes social proof from past client outcomes, and creates urgency with limited consultation availability.",
            "steps":           ["Lead enquires but\ndoesn't retain", "Instant: SMS +\nemail with booking link", "Days 1–3: Value\ncontent + FAQs", "Days 7–14:\nAddress objections", "Day 21: Final\nurgency message"],
            "audit_finding":   "After our web enquiry test, we received one email then silence. Zero SMS follow-up. With ~12 unconverted enquiries/month, a 25% nurture conversion rate = 3 extra retained cases = $36,000/month in case value.",
        },
        {
            "name":            "Review Bot + Client Reactivation",
            "label":           "ADD-ON — GROWTH",
            "metric_label":    "new reviews/mo +\npast clients reactivated",
            "metric_value":    "10+",
            "monthly_revenue": "$2,400",
            "monthly_cost":    "$800",
            "setup_fee":       "$2,000",
            "roi":             "3.0x",
            "description":     "Two systems working together: (1) automatically requests Google reviews from satisfied clients after case closure with one-tap links, (2) runs quarterly reactivation campaigns targeting past clients in the CRM with personalised win-back messages based on their case history.",
            "steps":           ["Case closes\nsuccessfully", "Review bot sends\none-tap Google link", "Past clients get\nquarterly reactivation", "Personalised by\ncase type", "Dashboard tracks\nreviews + reactivations"],
            "audit_finding":   "Fewer Google reviews than top local competing law firms. Years of past clients sitting untouched in CRM. Reactivating 5% of past clients = 2–3 new matters/month. Reviews driving SEO = 3–5 new organic enquiries/month.",
        },
    ],

    # ── Revenue Summary ───────────────────────
    "total_monthly_revenue": "$10,500",
    "total_monthly_cost":    "$2,400",
    "total_setup_fees":      "$5,500",
    "net_monthly_gain":      "$8,100",
    "annual_recovery":       "$97,200",
    "overall_roi":           "4.4x",

    # ── Guarantee ─────────────────────────────
    "guarantee_headline":  "Make your full 12-month investment back in 90 days.",
    "guarantee_investment": "$28,800",
    "guarantee_90day":      "$31,500",
    "guarantee_surplus":    "$2,700",

    # ── Before / After Transformation ─────────
    "transformation_rows": [
        {"label": "Missed calls/mo",       "current": "~8/month",        "with_ai": "< 2"},
        {"label": "Response time",         "current": "2–4 hours",       "with_ai": "< 60 seconds"},
        {"label": "After-hours coverage",  "current": "None",            "with_ai": "24/7 AI"},
        {"label": "Follow-up touches",     "current": "1–2 emails",      "with_ai": "7-touch multi-channel"},
        {"label": "Lead drop-off rate",    "current": "94%",             "with_ai": "< 60%"},
        {"label": "Google reviews",        "current": "Low",             "with_ai": "200+ in 6 months"},
        {"label": "Past client outreach",  "current": "None",            "with_ai": "Quarterly automated"},
    ],

    # ── Implementation ────────────────────────
    "implementation_phases": [
        {
            "period": "WEEK 1–2",
            "title":  "Build & Deploy",
            "items":  [
                "AI Receptionist goes live (24/7)",
                "CRM integration complete",
                "Smart consultation reminders on",
                "Intake automation active",
            ],
        },
        {
            "period": "WEEK 3–4",
            "title":  "Nurture & Convert",
            "items":  [
                "Lead nurture sequences live",
                "Unconverted lead follow-up running",
                "Cancellation/no-show recovery on",
                "Team trained on dashboards",
            ],
        },
        {
            "period": "MONTH 2–3",
            "title":  "Grow & Optimise",
            "items":  [
                "Client reactivation campaigns running",
                "Review generation bot launched",
                "A/B test all messaging",
                "Full ROI report delivered",
            ],
        },
    ],
    "implementation_closing": "Zero disruption to your team. We handle the entire build, CRM integration, and testing.",

    # ── Next Steps ────────────────────────────
    "next_steps": [
        {"title": "Select your systems",        "description": "Use the link we'll send you to choose which bots you want and customise your numbers."},
        {"title": "We build everything in 14 days", "description": "Complete setup, CRM integration, and testing — zero work required from your team."},
        {"title": "Revenue starts recovering",  "description": "Most clients see measurable results within the first 30 days."},
        {"title": "90-day guarantee kicks in",  "description": "If you haven't made your 12-month investment back in 90 days, we work for free until you do."},
    ],

    # ── Closing ───────────────────────────────
    "closing_urgency_line": "Every month without AI systems\n= $8,100 in lost revenue for Law Biz.",
    "closing_cta":          "Let's fix it.",
}


if __name__ == "__main__":
    generate(EXAMPLE_DATA, "Law_Biz_Revenue_Audit_v1.pptx")
