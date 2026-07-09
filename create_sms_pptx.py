#!/usr/bin/env python3
"""Ultra-modern SMS Management Presentation Generator."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn, nsmap
from lxml import etree
import copy

# ── Design System ──────────────────────────────────────────────
NAVY = RGBColor(0x0B, 0x1F, 0x3A)
NAVY_MID = RGBColor(0x14, 0x32, 0x5C)
NAVY_SOFT = RGBColor(0x1E, 0x3A, 0x5F)
GOLD = RGBColor(0xE8, 0xB9, 0x23)
GOLD_SOFT = RGBColor(0xF5, 0xD0, 0x6A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xF4, 0xF7, 0xFB)
SLATE = RGBColor(0x5A, 0x6B, 0x7D)
DARK = RGBColor(0x1A, 0x23, 0x32)
TEAL = RGBColor(0x1A, 0xB3, 0xA6)
BLUE = RGBColor(0x2E, 0x86, 0xDE)
ORANGE = RGBColor(0xE6, 0x7E, 0x22)
RED = RGBColor(0xE7, 0x4C, 0x3C)
GREEN = RGBColor(0x27, 0xAE, 0x60)
AMBER = RGBColor(0xF3, 0x9C, 0x12)
LIGHT_BLUE = RGBColor(0xD6, 0xE8, 0xF8)
LIGHT_TEAL = RGBColor(0xD4, 0xF1, 0xEE)
LIGHT_GOLD = RGBColor(0xFD, 0xF0, 0xD5)
GRAY = RGBColor(0x95, 0xA5, 0xA6)
SOFT_GRAY = RGBColor(0xEC, 0xF0, 0xF3)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# OOXML namespaces for morph
NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


def set_run(run, size=18, color=DARK, bold=False, font="Segoe UI"):
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.name = font


def add_text(shape, text, size=18, color=DARK, bold=False, align=PP_ALIGN.LEFT, font="Segoe UI"):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size, color, bold, font)
    return tf


def add_para(tf, text, size=14, color=DARK, bold=False, align=PP_ALIGN.LEFT, space_before=6, space_after=0, font="Segoe UI"):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = text
    set_run(run, size, color, bold, font)
    return p


def fill_shape(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def fill_line(shape, color, width=Pt(1.5)):
    shape.line.color.rgb = color
    shape.line.width = width


def rect(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    fill_shape(s, color)
    return s


def round_rect(slide, l, t, w, h, color, radius=None):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    fill_shape(s, color)
    return s


def oval(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, w, h)
    fill_shape(s, color)
    return s


def triangle(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, l, t, w, h)
    fill_shape(s, color)
    return s


def set_shape_name(shape, name):
    """Name shapes so Morph can match them across slides."""
    sp = shape._element
    nvSpPr = sp.find(qn("p:nvSpPr"))
    if nvSpPr is not None:
        cNvPr = nvSpPr.find(qn("p:cNvPr"))
        if cNvPr is not None:
            cNvPr.set("name", name)


def add_morph_transition(slide):
    """Add Morph transition (PowerPoint 2019+/Microsoft 365)."""
    sld = slide._element
    for child in list(sld):
        if "transition" in child.tag:
            sld.remove(child)

    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
    MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

    # Declare mc & p14 on the AlternateContent element itself so the
    # Requires="p14" prefix reference is in scope when PowerPoint parses it.
    alt = etree.SubElement(
        sld,
        f"{{{MC_NS}}}AlternateContent",
        nsmap={"mc": MC_NS, "p14": P14_NS},
    )
    choice = etree.SubElement(alt, f"{{{MC_NS}}}Choice")
    choice.set("Requires", "p14")
    transition = etree.SubElement(choice, f"{{{P_NS}}}transition")
    transition.set("spd", "med")
    morph = etree.SubElement(transition, f"{{{P14_NS}}}morph")
    morph.set("option", "byObject")

    fallback = etree.SubElement(alt, f"{{{MC_NS}}}Fallback")
    fb_trans = etree.SubElement(fallback, f"{{{P_NS}}}transition")
    fb_trans.set("spd", "med")
    etree.SubElement(fb_trans, f"{{{P_NS}}}fade")


def add_fade_transition(slide):
    sld = slide._element
    for child in list(sld):
        if "transition" in child.tag:
            sld.remove(child)
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    transition = etree.SubElement(sld, f"{{{P_NS}}}transition")
    transition.set("spd", "med")
    etree.SubElement(transition, f"{{{P_NS}}}fade")


def add_entrance_animations(slide, skip=1, stagger_ms=150, dur_ms=500):
    """Auto-playing staggered fade entrance for every shape on the slide.

    Builds the p:timing tree the same way PowerPoint does for a
    'Fade / Start: With Previous' effect on each shape, so the whole
    slide assembles itself when it appears.
    """
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

    shapes = list(slide.shapes)[skip:]
    if not shapes:
        return

    # Keep the full build under ~4 seconds on shape-heavy slides.
    total_budget = 4000
    stagger = min(stagger_ms, max(40, total_budget // max(1, len(shapes))))

    sld = slide._element
    for child in list(sld):
        if child.tag == f"{{{P_NS}}}timing":
            sld.remove(child)

    def el(parent, tag, **attrs):
        e = etree.SubElement(parent, f"{{{P_NS}}}{tag}")
        for k, v in attrs.items():
            e.set(k, str(v))
        return e

    next_id = [1]

    def nid():
        i = next_id[0]
        next_id[0] += 1
        return i

    timing = el(sld, "timing")
    tnLst = el(timing, "tnLst")
    root_par = el(tnLst, "par")
    root_ctn = el(root_par, "cTn", id=nid(), dur="indefinite", restart="never", nodeType="tmRoot")
    root_children = el(root_ctn, "childTnLst")

    seq = el(root_children, "seq", concurrent="1", nextAc="seek")
    main_ctn = el(seq, "cTn", id=nid(), dur="indefinite", nodeType="mainSeq")
    main_seq_id = main_ctn.get("id")
    main_children = el(main_ctn, "childTnLst")

    # Single group that fires automatically when the slide appears.
    grp_par = el(main_children, "par")
    grp_ctn = el(grp_par, "cTn", id=nid(), fill="hold")
    grp_st = el(grp_ctn, "stCondLst")
    el(grp_st, "cond", delay="indefinite")
    on_begin = el(grp_st, "cond", evt="onBegin", delay="0")
    tn_ref = el(on_begin, "tn")
    tn_ref.set("val", main_seq_id)
    grp_children = el(grp_ctn, "childTnLst")

    inner_par = el(grp_children, "par")
    inner_ctn = el(inner_par, "cTn", id=nid(), fill="hold")
    inner_st = el(inner_ctn, "stCondLst")
    el(inner_st, "cond", delay="0")
    inner_children = el(inner_ctn, "childTnLst")

    for i, shape in enumerate(shapes):
        spid = shape.shape_id
        delay = i * stagger

        eff_par = el(inner_children, "par")
        eff_ctn = el(
            eff_par, "cTn", id=nid(), presetID="10", presetClass="entr",
            presetSubtype="0", fill="hold", grpId="0", nodeType="withEffect",
        )
        eff_st = el(eff_ctn, "stCondLst")
        el(eff_st, "cond", delay=delay)
        eff_children = el(eff_ctn, "childTnLst")

        # visibility set
        set_el = el(eff_children, "set")
        set_bhvr = el(set_el, "cBhvr")
        set_ctn = el(set_bhvr, "cTn", id=nid(), dur="1", fill="hold")
        set_st = el(set_ctn, "stCondLst")
        el(set_st, "cond", delay="0")
        set_tgt = el(set_bhvr, "tgtEl")
        sp_tgt = el(set_tgt, "spTgt")
        sp_tgt.set("spid", str(spid))
        attr_lst = el(set_bhvr, "attrNameLst")
        attr = el(attr_lst, "attrName")
        attr.text = "style.visibility"
        to = el(set_el, "to")
        str_val = el(to, "strVal")
        str_val.set("val", "visible")

        # fade in
        anim = el(eff_children, "animEffect")
        anim.set("transition", "in")
        anim.set("filter", "fade")
        anim_bhvr = el(anim, "cBhvr")
        el(anim_bhvr, "cTn", id=nid(), dur=dur_ms)
        anim_tgt = el(anim_bhvr, "tgtEl")
        sp_tgt2 = el(anim_tgt, "spTgt")
        sp_tgt2.set("spid", str(spid))

    prev_cond_lst = el(seq, "prevCondLst")
    prev_cond = el(prev_cond_lst, "cond", evt="onPrev", delay="0")
    prev_tgt = el(prev_cond, "tgtEl")
    el(prev_tgt, "sldTgt")
    next_cond_lst = el(seq, "nextCondLst")
    next_cond = el(next_cond_lst, "cond", evt="onNext", delay="0")
    next_tgt = el(next_cond, "tgtEl")
    el(next_tgt, "sldTgt")


def accent_bar(slide, color=GOLD):
    return rect(slide, Inches(0), Inches(0), Inches(0.12), SLIDE_H, color)


def top_rule(slide, color=GOLD):
    return rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.08), color)


def section_label(slide, text, l=Inches(0.7), t=Inches(0.35)):
    box = slide.shapes.add_textbox(l, t, Inches(6), Inches(0.35))
    add_text(box, text.upper(), size=11, color=GOLD, bold=True, font="Segoe UI")
    set_shape_name(box, "sectionLabel")
    return box


def slide_title(slide, text, l=Inches(0.7), t=Inches(0.65), w=Inches(11.5)):
    box = slide.shapes.add_textbox(l, t, w, Inches(0.7))
    add_text(box, text, size=32, color=NAVY, bold=True, font="Segoe UI Semibold")
    set_shape_name(box, "slideTitle")
    return box


def icon_circle(slide, l, t, size, bg, symbol, symbol_color=WHITE, symbol_size=18):
    c = oval(slide, l, t, size, size, bg)
    tb = slide.shapes.add_textbox(l, t + size * 0.22, size, size * 0.6)
    add_text(tb, symbol, size=symbol_size, color=symbol_color, bold=True, align=PP_ALIGN.CENTER, font="Segoe UI Symbol")
    return c, tb


# ── Slides ─────────────────────────────────────────────────────

def slide_01_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    # Diagonal accent
    accent = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(9.5), Inches(0), Inches(3.9), SLIDE_H)
    fill_shape(accent, NAVY_MID)
    # Gold line
    rect(slide, Inches(0.9), Inches(3.15), Inches(1.8), Inches(0.06), GOLD)

    brand = slide.shapes.add_textbox(Inches(0.9), Inches(1.8), Inches(10), Inches(0.4))
    add_text(brand, "SAFETY MANAGEMENT SYSTEM", size=14, color=GOLD, bold=True, font="Segoe UI")
    set_shape_name(brand, "brand")

    title = slide.shapes.add_textbox(Inches(0.9), Inches(3.4), Inches(10), Inches(1.2))
    add_text(title, "Building a Culture of\nContinuous Safety", size=40, color=WHITE, bold=True, font="Segoe UI Semibold")
    set_shape_name(title, "mainTitle")

    sub = slide.shapes.add_textbox(Inches(0.9), Inches(5.5), Inches(9), Inches(0.5))
    add_text(sub, "Executive Briefing  ·  Policy · Risk · Assurance · Promotion", size=16, color=RGBColor(0xA8, 0xB8, 0xC8), font="Segoe UI")
    set_shape_name(sub, "subtitle")

    # Decorative dots
    for i, c in enumerate([GOLD, TEAL, BLUE]):
        o = oval(slide, Inches(0.9 + i * 0.35), Inches(6.5), Inches(0.18), Inches(0.18), c)
        set_shape_name(o, f"dot{i}")

    add_fade_transition(slide)
    return slide


def slide_02_agenda(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Roadmap")
    slide_title(slide, "Today's Focus")

    items = [
        ("01", "Safety Policy", "Values, objectives & ownership"),
        ("02", "Safety Risk Management", "Identify · Assess · Mitigate"),
        ("03", "Safety Assurance", "Measure what matters"),
        ("04", "Safety Promotion", "Culture, training & improvement"),
    ]
    colors = [GOLD, TEAL, BLUE, ORANGE]
    for i, ((num, title, desc), col) in enumerate(zip(items, colors)):
        x = Inches(0.7 + (i % 4) * 3.1)
        y = Inches(2.2)
        card = round_rect(slide, x, y, Inches(2.9), Inches(3.8), WHITE)
        set_shape_name(card, f"agendaCard{i}")
        bar = rect(slide, x, y, Inches(2.9), Inches(0.1), col)
        set_shape_name(bar, f"agendaBar{i}")
        n = slide.shapes.add_textbox(x + Inches(0.25), y + Inches(0.5), Inches(2.4), Inches(0.7))
        add_text(n, num, size=36, color=col, bold=True)
        set_shape_name(n, f"agendaNum{i}")
        t = slide.shapes.add_textbox(x + Inches(0.25), y + Inches(1.5), Inches(2.4), Inches(1))
        add_text(t, title, size=18, color=NAVY, bold=True)
        set_shape_name(t, f"agendaTitle{i}")
        d = slide.shapes.add_textbox(x + Inches(0.25), y + Inches(2.5), Inches(2.4), Inches(0.9))
        add_text(d, desc, size=13, color=SLATE)
        set_shape_name(d, f"agendaDesc{i}")

    add_morph_transition(slide)
    return slide


def slide_03_policy(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Component 1")
    slide_title(slide, "Safety Policy")

    points = [
        ("◆", "Define Values", "Brainstorm your organization's values and write them down."),
        ("◆", "Anchor Safety", "The opening paragraph of your safety policy should reflect where safety fits into your values."),
        ("◆", "Dream", "What would you like the SMS to do for your organization?"),
        ("◆", "Set Objectives", "Set those as the SMS objectives in your policy."),
        ("◆", "Assign Ownership", "Outline high level responsibilities for all employees of your organization."),
    ]
    for i, (icon, title, desc) in enumerate(points):
        y = Inches(1.7 + i * 1.0)
        pill = round_rect(slide, Inches(0.7), y, Inches(11.8), Inches(0.85), WHITE)
        set_shape_name(pill, f"policyRow{i}")
        num_bg = oval(slide, Inches(0.95), y + Inches(0.15), Inches(0.55), Inches(0.55), NAVY if i % 2 == 0 else TEAL)
        set_shape_name(num_bg, f"policyIcon{i}")
        nb = slide.shapes.add_textbox(Inches(0.95), y + Inches(0.25), Inches(0.55), Inches(0.4))
        add_text(nb, str(i + 1), size=16, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(nb, f"policyNum{i}")
        tb = slide.shapes.add_textbox(Inches(1.8), y + Inches(0.12), Inches(10), Inches(0.35))
        add_text(tb, title, size=16, color=NAVY, bold=True)
        set_shape_name(tb, f"policyTitle{i}")
        db = slide.shapes.add_textbox(Inches(1.8), y + Inches(0.45), Inches(10), Inches(0.3))
        add_text(db, desc, size=13, color=SLATE)
        set_shape_name(db, f"policyDesc{i}")

    add_morph_transition(slide)
    return slide


def slide_04_objective_intro(prs):
    """Morph step 1: objective statement only."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    label = slide.shapes.add_textbox(Inches(0.9), Inches(2.4), Inches(11), Inches(0.4))
    add_text(label, "OUR SAFETY OBJECTIVE IS", size=16, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
    set_shape_name(label, "objLabel")
    title = slide.shapes.add_textbox(Inches(0.9), Inches(3.1), Inches(11.5), Inches(1))
    add_text(title, '"A Continuous Reduction in Mishaps"', size=36, color=WHITE, bold=True, align=PP_ALIGN.CENTER, font="Segoe UI Semibold")
    set_shape_name(title, "objTitle")
    add_morph_transition(slide)
    return slide


def slide_05_pyramid(prs):
    """Morph step 2: safety pyramid + bar chart."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)

    label = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(8), Inches(0.35))
    add_text(label, "OUR SAFETY OBJECTIVE IS", size=12, color=GOLD, bold=True)
    set_shape_name(label, "objLabel")
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.65), Inches(8), Inches(0.5))
    add_text(title, '"A Continuous Reduction in Mishaps"', size=22, color=WHITE, bold=True, font="Segoe UI Semibold")
    set_shape_name(title, "objTitle")

    # Pyramid tiers (bottom to top visually drawn bottom-up)
    # Bottom: Close Calls 10000
    t1 = triangle(slide, Inches(1.2), Inches(1.8), Inches(5.2), Inches(4.6), RGBColor(0xF1, 0xC4, 0x0F))
    set_shape_name(t1, "pyramidBase")
    # Middle overlay
    t2 = triangle(slide, Inches(1.95), Inches(1.8), Inches(3.7), Inches(3.2), ORANGE)
    set_shape_name(t2, "pyramidMid")
    # Top
    t3 = triangle(slide, Inches(2.7), Inches(1.8), Inches(2.2), Inches(1.8), RED)
    set_shape_name(t3, "pyramidTop")

    # Labels left of pyramid
    labels = [
        (Inches(0.4), Inches(2.3), "ACCIDENT", "1", RED),
        (Inches(0.4), Inches(3.6), "INCIDENTS", "1,000", ORANGE),
        (Inches(0.4), Inches(5.0), "CLOSE CALLS", "10,000", GOLD),
    ]
    for i, (lx, ly, name, num, col) in enumerate(labels):
        lb = slide.shapes.add_textbox(lx, ly, Inches(1.8), Inches(0.35))
        add_text(lb, name, size=12, color=col, bold=True)
        set_shape_name(lb, f"pyrLabel{i}")

    # Numbers on pyramid
    nums = [
        (Inches(3.2), Inches(2.3), "1"),
        (Inches(3.1), Inches(3.7), "1,000"),
        (Inches(2.9), Inches(5.2), "10,000"),
    ]
    for i, (nx, ny, n) in enumerate(nums):
        nb = slide.shapes.add_textbox(nx, ny, Inches(1.5), Inches(0.4))
        add_text(nb, n, size=16, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(nb, f"pyrNum{i}")

    # Magnifying glass accent on top
    mag = oval(slide, Inches(4.6), Inches(1.5), Inches(0.55), Inches(0.55), WHITE)
    set_shape_name(mag, "magnifier")
    mag_inner = oval(slide, Inches(4.7), Inches(1.6), Inches(0.35), Inches(0.35), RED)
    set_shape_name(mag_inner, "magnifierInner")

    # Bar chart on right
    chart_data = CategoryChartData()
    chart_data.categories = ["Close Calls", "Incidents", "Accidents"]
    chart_data.add_series("Events", (10000, 1000, 1))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(7.0), Inches(1.8), Inches(5.6), Inches(4.2),
        chart_data,
    ).chart
    chart.has_legend = False
    plot = chart.plots[0]
    plot.gap_width = 80
    series = chart.series[0]
    point_colors = [RGBColor(0xF1, 0xC4, 0x0F), ORANGE, RED]
    for idx, col in enumerate(point_colors):
        point = series.points[idx]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = col

    try:
        chart.has_title = False
    except Exception:
        pass

    footer = slide.shapes.add_textbox(Inches(0.5), Inches(6.7), Inches(12), Inches(0.45))
    add_text(
        footer,
        "We must be vigilant to FIND the next POTENTIAL accident and prevent it from HAPPENING",
        size=14, color=GOLD_SOFT, bold=True, align=PP_ALIGN.CENTER,
    )
    set_shape_name(footer, "pyrFooter")

    add_morph_transition(slide)
    return slide


def slide_06_commitment(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Accountability")
    slide_title(slide, "Management Commitment & Safety Accountabilities")

    cards = [
        (NAVY, "01", "One Responsible Person", "One person must have the responsibility to oversee SMS development, implementation and operation."),
        (TEAL, "02", "The \u201cChampion\u201d", "This person must be the champion for the SMS program — but does not bear the principal responsibility for safety management."),
        (BLUE, "03", "Line Managers", "Managers of the \u201cline\u201d operational functions, from middle management to front-line supervisors, manage the operations in which risk is incurred."),
        (ORANGE, "04", "Owners of the SMS", "These managers and supervisors are the \u201cowners\u201d of the SMS."),
    ]
    for i, (col, num, title, desc) in enumerate(cards):
        x = Inches(0.6 + (i % 2) * 6.3)
        y = Inches(1.8 + (i // 2) * 2.5)
        card = round_rect(slide, x, y, Inches(6.0), Inches(2.2), WHITE)
        set_shape_name(card, f"commitCard{i}")
        stripe = rect(slide, x, y, Inches(0.12), Inches(2.2), col)
        set_shape_name(stripe, f"commitStripe{i}")
        nb = slide.shapes.add_textbox(x + Inches(0.4), y + Inches(0.35), Inches(1), Inches(0.5))
        add_text(nb, num, size=24, color=col, bold=True)
        set_shape_name(nb, f"commitNum{i}")
        tb = slide.shapes.add_textbox(x + Inches(0.4), y + Inches(0.9), Inches(5.2), Inches(0.4))
        add_text(tb, title, size=20, color=NAVY, bold=True)
        set_shape_name(tb, f"commitTitle{i}")
        db = slide.shapes.add_textbox(x + Inches(0.4), y + Inches(1.4), Inches(5.2), Inches(0.55))
        add_text(db, desc, size=13, color=SLATE)
        set_shape_name(db, f"commitDesc{i}")

    add_morph_transition(slide)
    return slide


def slide_07_swiss_cheese(prs):
    """Swiss cheese / barriers model — visual."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Management & Supervision")
    slide_title(slide, "Barriers or Controls")

    work = slide.shapes.add_textbox(Inches(9.0), Inches(2.2), Inches(3.5), Inches(0.5))
    add_text(work, "WORK", size=20, color=GREEN, bold=True, align=PP_ALIGN.CENTER)
    set_shape_name(work, "workLabel")

    # Hazard
    haz = round_rect(slide, Inches(0.5), Inches(3.2), Inches(1.8), Inches(1.0), GOLD)
    set_shape_name(haz, "hazardBox")
    ht = slide.shapes.add_textbox(Inches(0.5), Inches(3.4), Inches(1.8), Inches(0.6))
    add_text(ht, "HAZARD\n/ RISK", size=12, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    set_shape_name(ht, "hazardText")

    # Arrow path
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(2.4), Inches(3.45), Inches(0.7), Inches(0.45))
    fill_shape(arrow, RED)
    set_shape_name(arrow, "riskArrow")

    # Four barrier slices
    barrier_colors = [ORANGE, BLUE, GREEN, GRAY]
    barrier_names = ["Policy", "Process", "Training", "Supervision"]
    for i, (col, name) in enumerate(zip(barrier_colors, barrier_names)):
        x = Inches(3.3 + i * 1.35)
        bar = round_rect(slide, x, Inches(2.0), Inches(1.1), Inches(3.5), col)
        set_shape_name(bar, f"barrier{i}")
        # holes
        for j, hy in enumerate([Inches(2.5), Inches(3.5), Inches(4.4)]):
            hole = oval(slide, x + Inches(0.3), hy, Inches(0.5), Inches(0.5), OFF_WHITE)
            set_shape_name(hole, f"hole{i}_{j}")
        # label under
        lb = slide.shapes.add_textbox(x - Inches(0.1), Inches(5.6), Inches(1.3), Inches(0.4))
        add_text(lb, name, size=11, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(lb, f"barrierLabel{i}")

    # Outcome blocked
    out = round_rect(slide, Inches(9.0), Inches(3.1), Inches(3.5), Inches(1.2), RGBColor(0xFD, 0xE8, 0xE6))
    set_shape_name(out, "outcomeBox")
    ot = slide.shapes.add_textbox(Inches(9.0), Inches(3.35), Inches(3.5), Inches(0.7))
    add_text(ot, "UNDESIRABLE\nOUTCOME  ✕", size=14, color=RED, bold=True, align=PP_ALIGN.CENTER)
    set_shape_name(ot, "outcomeText")

    # Supervisors callout
    sup = round_rect(slide, Inches(3.3), Inches(6.2), Inches(5.4), Inches(0.8), NAVY)
    set_shape_name(sup, "supervisorBar")
    st = slide.shapes.add_textbox(Inches(3.3), Inches(6.35), Inches(5.4), Inches(0.5))
    add_text(st, "SUPERVISORS maintain every barrier layer", size=14, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    set_shape_name(st, "supervisorText")

    note = slide.shapes.add_textbox(Inches(9.0), Inches(4.6), Inches(3.5), Inches(1.2))
    add_text(note, "Aligned holes = pathway to failure.\nStrong supervision closes the gaps.", size=12, color=SLATE, align=PP_ALIGN.CENTER)
    set_shape_name(note, "swissNote")

    add_morph_transition(slide)
    return slide


def slide_08_key_personnel(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Structure")
    slide_title(slide, "Key Safety Personnel")

    # Top management banner
    top = round_rect(slide, Inches(0.6), Inches(1.6), Inches(12.1), Inches(1.1), NAVY)
    set_shape_name(top, "topMgmt")
    tt = slide.shapes.add_textbox(Inches(0.9), Inches(1.75), Inches(11.5), Inches(0.8))
    tf = add_text(tt, "TOP MANAGEMENT", size=16, color=GOLD, bold=True)
    add_para(tf, "Has the ultimate responsibility for the SMS, provides the resources essential to implement and maintain it, and appoints a member of management — the Safety Manager.", size=12, color=WHITE, space_before=4)
    set_shape_name(tt, "topMgmtText")

    # Safety manager card
    sm = round_rect(slide, Inches(0.6), Inches(3.0), Inches(4.0), Inches(3.8), WHITE)
    set_shape_name(sm, "smCard")
    rect(slide, Inches(0.6), Inches(3.0), Inches(4.0), Inches(0.1), TEAL)
    smt = slide.shapes.add_textbox(Inches(0.9), Inches(3.3), Inches(3.5), Inches(0.5))
    add_text(smt, "Safety Manager", size=18, color=NAVY, bold=True)
    set_shape_name(smt, "smTitle")
    duties = [
        "Ensuring the processes needed for the SMS are established, implemented and maintained",
        "Reporting to top management on the performance of the SMS and the need for improvement",
        "Ensuring the promotion of awareness of safety requirements throughout the organization",
        "Ensuring safety-related positions, responsibilities and authorities are defined, documented and communicated",
    ]
    for i, d in enumerate(duties):
        db = slide.shapes.add_textbox(Inches(0.9), Inches(3.85 + i * 0.75), Inches(3.5), Inches(0.72))
        add_text(db, f"→  {d}", size=10.5, color=SLATE)
        set_shape_name(db, f"smDuty{i}")

    # Icon metric cards
    metrics = [
        (GOLD, "100%", "Resource\nCommitment"),
        (TEAL, "1", "Appointed\nSafety Lead"),
        (BLUE, "All", "Roles\nDocumented"),
    ]
    for i, (col, val, label) in enumerate(metrics):
        x = Inches(5.0 + i * 2.7)
        card = round_rect(slide, x, Inches(3.0), Inches(2.5), Inches(3.8), WHITE)
        set_shape_name(card, f"metricCard{i}")
        circ = oval(slide, x + Inches(0.65), Inches(3.5), Inches(1.2), Inches(1.2), col)
        set_shape_name(circ, f"metricCirc{i}")
        vb = slide.shapes.add_textbox(x + Inches(0.65), Inches(3.8), Inches(1.2), Inches(0.6))
        add_text(vb, val, size=18, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(vb, f"metricVal{i}")
        lb = slide.shapes.add_textbox(x + Inches(0.2), Inches(5.1), Inches(2.1), Inches(1))
        add_text(lb, label, size=14, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(lb, f"metricLabel{i}")

    add_morph_transition(slide)
    return slide


def slide_09_erp(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Readiness")
    slide_title(slide, "Emergency Response Plan")

    # Definition
    defn = round_rect(slide, Inches(0.6), Inches(1.6), Inches(12.1), Inches(1.2), NAVY)
    set_shape_name(defn, "erpDef")
    dt = slide.shapes.add_textbox(Inches(0.9), Inches(1.8), Inches(11.5), Inches(0.9))
    add_text(dt, "An ERP outlines in writing what is done when an emergency occurs, what to do after an accident happens — and who is responsible for each action.", size=15, color=WHITE, align=PP_ALIGN.CENTER)
    set_shape_name(dt, "erpDefText")

    # Before / During / After
    phases = [
        (TEAL, "PREPARE", "Be Ready", "The better prepared, the better the chances of minimizing injury & damage"),
        (GOLD, "RESPOND", "Act Fast", "Clear roles — who is responsible for each action"),
        (ORANGE, "ACCESS", "Keep It Close", "Readily available at the work stations of first responders"),
    ]
    for i, (col, phase, title, desc) in enumerate(phases):
        x = Inches(0.6 + i * 4.2)
        card = round_rect(slide, x, Inches(3.2), Inches(3.9), Inches(2.4), WHITE)
        set_shape_name(card, f"erpPhase{i}")
        bar = rect(slide, x, Inches(3.2), Inches(3.9), Inches(0.12), col)
        set_shape_name(bar, f"erpBar{i}")
        pb = slide.shapes.add_textbox(x + Inches(0.3), Inches(3.5), Inches(3.3), Inches(0.35))
        add_text(pb, phase, size=12, color=col, bold=True)
        set_shape_name(pb, f"erpPhaseLabel{i}")
        tb = slide.shapes.add_textbox(x + Inches(0.3), Inches(3.95), Inches(3.3), Inches(0.45))
        add_text(tb, title, size=22, color=NAVY, bold=True)
        set_shape_name(tb, f"erpTitle{i}")
        db = slide.shapes.add_textbox(x + Inches(0.3), Inches(4.6), Inches(3.3), Inches(0.6))
        add_text(db, desc, size=13, color=SLATE)
        set_shape_name(db, f"erpDesc{i}")

    # Impact bar chart concept
    footer = slide.shapes.add_textbox(Inches(0.6), Inches(6.0), Inches(12.1), Inches(0.9))
    tf = add_text(footer, "GOAL", size=12, color=GOLD, bold=True)
    add_para(tf, "Minimize injuries to personnel and damage to equipment, property or the environment.", size=14, color=NAVY, space_before=4)
    set_shape_name(footer, "erpFooter")

    add_morph_transition(slide)
    return slide


def slide_10_documentation(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Governance")
    slide_title(slide, "SMS Documentation")

    sub = slide.shapes.add_textbox(Inches(0.7), Inches(1.45), Inches(11), Inches(0.4))
    add_text(sub, "Clearly defined and documented safety policies, objectives and procedures — in paper or electronic format.", size=14, color=SLATE)
    set_shape_name(sub, "docSub")

    standards = [
        ("01", "Orderly", "Maintained in a structured system"),
        ("02", "Identifiable", "Easy to find and recognize"),
        ("03", "Retrievable", "Available when needed"),
        ("04", "Legible", "Clear and usable"),
        ("05", "Dated", "Revision date always visible"),
    ]
    for i, (num, title, desc) in enumerate(standards):
        x = Inches(0.5 + i * 2.5)
        card = round_rect(slide, x, Inches(2.2), Inches(2.35), Inches(3.5), WHITE)
        set_shape_name(card, f"docCard{i}")
        circ = oval(slide, x + Inches(0.6), Inches(2.6), Inches(1.15), Inches(1.15), NAVY if i % 2 == 0 else TEAL)
        set_shape_name(circ, f"docCirc{i}")
        nb = slide.shapes.add_textbox(x + Inches(0.6), Inches(2.9), Inches(1.15), Inches(0.55))
        add_text(nb, num, size=18, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(nb, f"docNum{i}")
        tb = slide.shapes.add_textbox(x + Inches(0.15), Inches(4.1), Inches(2.05), Inches(0.45))
        add_text(tb, title, size=16, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(tb, f"docTitle{i}")
        db = slide.shapes.add_textbox(x + Inches(0.15), Inches(4.6), Inches(2.05), Inches(0.8))
        add_text(db, desc, size=12, color=SLATE, align=PP_ALIGN.CENTER)
        set_shape_name(db, f"docDesc{i}")

    foot = slide.shapes.add_textbox(Inches(0.7), Inches(6.2), Inches(12), Inches(0.5))
    add_text(foot, "The organization should determine how long records should be retained.", size=13, color=SLATE, align=PP_ALIGN.CENTER)
    set_shape_name(foot, "docFoot")

    add_morph_transition(slide)
    return slide


def slide_11_risk_strategies(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Component 2")
    slide_title(slide, "Safety Risk Management")

    strategies = [
        (ORANGE, "REACTIVE", "(Past)", "Responds to events that have already happened."),
        (GOLD, "PROACTIVE", "(Present)", "Actively seeks the identification of hazardous conditions through the analysis of the organization's processes."),
        (TEAL, "PREDICTIVE", "(Future)", "Analyzes system processes and environment to identify potential future problems."),
    ]
    for i, (col, name, time, desc) in enumerate(strategies):
        x = Inches(0.6 + i * 4.2)
        card = round_rect(slide, x, Inches(1.8), Inches(3.95), Inches(4.6), WHITE)
        set_shape_name(card, f"riskCard{i}")
        header = rect(slide, x, Inches(1.8), Inches(3.95), Inches(1.4), col)
        set_shape_name(header, f"riskHeader{i}")
        tb = slide.shapes.add_textbox(x + Inches(0.25), Inches(2.05), Inches(3.45), Inches(0.45))
        add_text(tb, name, size=18, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(tb, f"riskName{i}")
        timeb = slide.shapes.add_textbox(x + Inches(0.25), Inches(2.55), Inches(3.45), Inches(0.4))
        add_text(timeb, time, size=14, color=WHITE, align=PP_ALIGN.CENTER)
        set_shape_name(timeb, f"riskTime{i}")
        db = slide.shapes.add_textbox(x + Inches(0.35), Inches(3.7), Inches(3.25), Inches(1.8))
        add_text(db, desc, size=15, color=NAVY, align=PP_ALIGN.CENTER)
        set_shape_name(db, f"riskDesc{i}")

        # Timeline dots
        for j in range(3):
            c = oval(slide, x + Inches(1.2 + j * 0.55), Inches(5.6), Inches(0.28), Inches(0.28), col if j == i else SOFT_GRAY)
            set_shape_name(c, f"riskDot{i}_{j}")

    add_morph_transition(slide)
    return slide


def slide_12_hazard_reporting(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Detection")
    slide_title(slide, "Hazard Identification & Reporting")

    # Left: hazard ID
    left = round_rect(slide, Inches(0.6), Inches(1.7), Inches(5.8), Inches(5.0), WHITE)
    set_shape_name(left, "hazCard")
    rect(slide, Inches(0.6), Inches(1.7), Inches(5.8), Inches(0.1), GOLD)
    lt = slide.shapes.add_textbox(Inches(0.9), Inches(2.0), Inches(5.2), Inches(0.5))
    add_text(lt, "HAZARD IDENTIFICATION", size=16, color=NAVY, bold=True)
    set_shape_name(lt, "hazTitle")
    points = [
        "The SMS identifies hazards and develops processes to identify and manage risks",
        "Proactive identification of existing and potential hazards",
        "Includes hazards from organizational change — rapid growth",
        "New services, new equipment or new personnel",
    ]
    for i, p in enumerate(points):
        pb = slide.shapes.add_textbox(Inches(0.9), Inches(2.8 + i * 0.75), Inches(5.2), Inches(0.6))
        add_text(pb, f"▸  {p}", size=14, color=SLATE)
        set_shape_name(pb, f"hazPoint{i}")

    # Right: reporting
    right = round_rect(slide, Inches(6.8), Inches(1.7), Inches(5.9), Inches(5.0), NAVY)
    set_shape_name(right, "repCard")
    rt = slide.shapes.add_textbox(Inches(7.1), Inches(2.0), Inches(5.3), Inches(0.5))
    add_text(rt, "REPORTING SYSTEMS", size=16, color=GOLD, bold=True)
    set_shape_name(rt, "repTitle")
    q = slide.shapes.add_textbox(Inches(7.1), Inches(2.55), Inches(5.3), Inches(0.4))
    add_text(q, "AND HOW WILL YOU REPORT..?", size=13, color=RGBColor(0xA8, 0xB8, 0xC8))
    set_shape_name(q, "repQ")
    reps = [
        "Keep it simple and accessible",
        "Re-active and pro-active processes can overlap",
        "Ensure people submitting reports get feedback",
        "Find a simple way to file and track reports",
    ]
    for i, r in enumerate(reps):
        rb = slide.shapes.add_textbox(Inches(7.1), Inches(3.2 + i * 0.7), Inches(5.3), Inches(0.55))
        add_text(rb, f"●  {r}", size=14, color=WHITE)
        set_shape_name(rb, f"repPoint{i}")

    add_morph_transition(slide)
    return slide


def slide_13_risk_matrix(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Decision Tool")
    slide_title(slide, "Safety Risk Matrix")

    sub = slide.shapes.add_textbox(Inches(0.7), Inches(1.4), Inches(12), Inches(0.35))
    add_text(sub, "Likelihood × Severity  →  Risk Index", size=14, color=SLATE)
    set_shape_name(sub, "matrixSub")

    # Matrix: rows = probability 5..1, cols = severity A..E
    # Color map
    # Red: 5A,5B,5C,4A,4B,3A
    # Yellow: 5D,5E,4C,4D,4E,3B,3C,3D,2A,2B,2C,1A
    # Green: 3E,2D,2E,1B,1C,1D,1E
    red_cells = {"5A", "5B", "5C", "4A", "4B", "3A"}
    yellow_cells = {"5D", "5E", "4C", "4D", "4E", "3B", "3C", "3D", "2A", "2B", "2C", "1A"}

    probs = [("5", "Frequent"), ("4", "Occasional"), ("3", "Remote"), ("2", "Improbable"), ("1", "Ext. Improb.")]
    sevs = [("A", "Catastrophic"), ("B", "Hazardous"), ("C", "Major"), ("D", "Minor"), ("E", "Negligible")]

    cell_w, cell_h = Inches(1.35), Inches(0.72)
    origin_x, origin_y = Inches(3.2), Inches(2.0)

    # Column headers
    for j, (code, name) in enumerate(sevs):
        x = origin_x + j * cell_w
        hb = slide.shapes.add_textbox(x, Inches(1.75), cell_w, Inches(0.3))
        add_text(hb, code, size=11, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(hb, f"sevHead{j}")

    # Row headers + cells
    for i, (pcode, pname) in enumerate(probs):
        y = origin_y + i * cell_h
        rh = slide.shapes.add_textbox(Inches(0.5), y + Inches(0.15), Inches(2.5), Inches(0.45))
        add_text(rh, f"{pcode}  {pname}", size=11, color=NAVY, bold=True)
        set_shape_name(rh, f"probHead{i}")
        for j, (scode, _) in enumerate(sevs):
            key = f"{pcode}{scode}"
            if key in red_cells:
                col = RED
            elif key in yellow_cells:
                col = AMBER
            else:
                col = GREEN
            x = origin_x + j * cell_w
            cell = round_rect(slide, x + Inches(0.04), y + Inches(0.04), cell_w - Inches(0.08), cell_h - Inches(0.08), col)
            set_shape_name(cell, f"cell_{key}")
            ct = slide.shapes.add_textbox(x, y + Inches(0.18), cell_w, Inches(0.4))
            add_text(ct, key, size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
            set_shape_name(ct, f"cellText_{key}")

    # Legend
    legend = [
        (RED, "INTOLERABLE"),
        (AMBER, "TOLERABLE"),
        (GREEN, "ACCEPTABLE"),
    ]
    for i, (col, name) in enumerate(legend):
        x = Inches(10.3)
        y = Inches(2.2 + i * 1.1)
        box = round_rect(slide, x, y, Inches(2.5), Inches(0.9), col)
        set_shape_name(box, f"legend{i}")
        tb = slide.shapes.add_textbox(x, y + Inches(0.25), Inches(2.5), Inches(0.4))
        add_text(tb, name, size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(tb, f"legendText{i}")

    note = slide.shapes.add_textbox(Inches(0.7), Inches(6.15), Inches(12), Inches(0.7))
    add_text(note, "Note — In determining the safety risk tolerability, the quality and reliability of the data used for the hazard identification and safety risk probability should be taken into consideration.", size=11, color=SLATE, align=PP_ALIGN.CENTER)
    set_shape_name(note, "matrixNote")

    add_morph_transition(slide)
    return slide


def slide_14_risk_actions(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Action Framework")
    slide_title(slide, "Risk Response Levels")

    rows = [
        (RED, "INTOLERABLE", "5A, 5B, 5C, 4A, 4B, 3A",
         "Take immediate action to mitigate the risk or stop the activity. Perform priority safety risk mitigation to ensure additional or enhanced preventative controls are in place to bring the safety risk index down to tolerable."),
        (AMBER, "TOLERABLE", "5D, 5E, 4C, 4D, 4E, 3B, 3C, 3D, 2A, 2B, 2C, 1A",
         "Can be tolerated based on the safety risk mitigation. It may require management decision to accept the risk."),
        (GREEN, "ACCEPTABLE", "3E, 2D, 2E, 1B, 1C, 1D, 1E",
         "Acceptable as is. No further safety risk mitigation required."),
    ]
    for i, (col, title, rng, desc) in enumerate(rows):
        y = Inches(1.75 + i * 1.65)
        card = round_rect(slide, Inches(0.7), y, Inches(12), Inches(1.5), WHITE)
        set_shape_name(card, f"actionCard{i}")
        badge = round_rect(slide, Inches(0.95), y + Inches(0.2), Inches(3.2), Inches(0.7), col)
        set_shape_name(badge, f"actionBadge{i}")
        bt = slide.shapes.add_textbox(Inches(0.95), y + Inches(0.35), Inches(3.2), Inches(0.45))
        add_text(bt, title, size=15, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(bt, f"actionTitle{i}")
        rb = slide.shapes.add_textbox(Inches(0.95), y + Inches(0.95), Inches(3.2), Inches(0.5))
        add_text(rb, rng, size=10, color=SLATE, align=PP_ALIGN.CENTER)
        set_shape_name(rb, f"actionRange{i}")
        db = slide.shapes.add_textbox(Inches(4.5), y + Inches(0.2), Inches(7.9), Inches(1.2))
        add_text(db, desc, size=12.5, color=NAVY)
        set_shape_name(db, f"actionDesc{i}")

    note = slide.shapes.add_textbox(Inches(0.7), Inches(6.75), Inches(12), Inches(0.4))
    add_text(note, "Based on hazard identification analysis — likelihood of occurrence × severity of resulting consequences.", size=12, color=SLATE, align=PP_ALIGN.CENTER)
    set_shape_name(note, "actionNote")

    add_morph_transition(slide)
    return slide


def slide_15_assurance(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Component 3")
    slide_title(slide, "Safety Assurance")

    sub = slide.shapes.add_textbox(Inches(0.7), Inches(1.4), Inches(12), Inches(0.35))
    add_text(sub, "Safety performance measures link directly to operational performance.", size=14, color=SLATE)
    set_shape_name(sub, "assSub")

    # 2x2 strategy map
    cells = [
        (NAVY, "Organization Objective", "Reduce Costs"),
        (TEAL, "Organization Performance Measure", "Reduction in Insurance Rates"),
        (GOLD, "Safety Objective", "Decrease Number and Severity\nof Hangar Incidents"),
        (BLUE, "Safety Performance Measure", "Reduction in total number of events:\nDamage-only · Near-miss · Lessons Learned · Corrective Action Plans"),
    ]
    for i, (col, header, body) in enumerate(cells):
        x = Inches(0.6 + (i % 2) * 6.3)
        y = Inches(2.0 + (i // 2) * 2.4)
        card = round_rect(slide, x, y, Inches(6.0), Inches(2.15), WHITE)
        set_shape_name(card, f"assCard{i}")
        bar = rect(slide, x, y, Inches(6.0), Inches(0.1), col)
        set_shape_name(bar, f"assBar{i}")
        hb = slide.shapes.add_textbox(x + Inches(0.35), y + Inches(0.35), Inches(5.3), Inches(0.4))
        add_text(hb, header, size=13, color=col, bold=True)
        set_shape_name(hb, f"assHead{i}")
        bb = slide.shapes.add_textbox(x + Inches(0.35), y + Inches(0.85), Inches(5.3), Inches(1.1))
        add_text(bb, body, size=15, color=NAVY, bold=True)
        set_shape_name(bb, f"assBody{i}")

    add_morph_transition(slide)
    return slide


def slide_16_spi_chart(prs):
    """Bar chart of example SPIs."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Monitoring & Measurement")
    slide_title(slide, "Safety Performance Monitoring")

    sub = slide.shapes.add_textbox(Inches(0.7), Inches(1.35), Inches(12), Inches(0.35))
    add_text(sub, "Safety performance is proactively and reactively monitored so key safety goals continue to be achieved. Example SPI format:", size=13, color=SLATE)
    set_shape_name(sub, "spiSub")

    chart_data = CategoryChartData()
    chart_data.categories = ["Damage-only\nEvents", "Near-miss\nAccidents", "Lessons\nLearned", "CAPs\nClosed"]
    chart_data.add_series("Previous", (18, 42, 12, 20))
    chart_data.add_series("Current", (11, 28, 24, 35))

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.8), Inches(1.7), Inches(8.2), Inches(5.0),
        chart_data,
    )
    chart = chart_shape.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False

    # Color series
    for s_idx, col in enumerate([SLATE, TEAL]):
        ser = chart.series[s_idx]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = col

    # Insight cards
    insights = [
        (GREEN, "↓ 39%", "Damage events"),
        (GREEN, "↓ 33%", "Near-misses"),
        (TEAL, "↑ 100%", "Lessons captured"),
        (TEAL, "↑ 75%", "CAPs closed"),
    ]
    for i, (col, val, label) in enumerate(insights):
        y = Inches(1.8 + i * 1.2)
        card = round_rect(slide, Inches(9.4), y, Inches(3.3), Inches(1.05), WHITE)
        set_shape_name(card, f"spiCard{i}")
        vb = slide.shapes.add_textbox(Inches(9.6), y + Inches(0.15), Inches(2.9), Inches(0.4))
        add_text(vb, val, size=20, color=col, bold=True)
        set_shape_name(vb, f"spiVal{i}")
        lb = slide.shapes.add_textbox(Inches(9.6), y + Inches(0.55), Inches(2.9), Inches(0.35))
        add_text(lb, label, size=12, color=SLATE)
        set_shape_name(lb, f"spiLabel{i}")

    add_morph_transition(slide)
    return slide


def slide_17_moc(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Change Control")
    slide_title(slide, "Management of Change")

    intro = slide.shapes.add_textbox(Inches(0.7), Inches(1.4), Inches(12), Inches(0.4))
    add_text(intro, "The MOC process has four basic phases. Both the effect of change and the effect of implementing change are considered.", size=13, color=SLATE)
    set_shape_name(intro, "mocIntro")

    phases = [
        (NAVY, "1", "Screening"),
        (TEAL, "2", "Review"),
        (GOLD, "3", "Approval"),
        (ORANGE, "4", "Implementation"),
    ]
    for i, (col, num, name) in enumerate(phases):
        x = Inches(0.7 + i * 3.15)
        circ = oval(slide, x + Inches(0.95), Inches(1.95), Inches(0.95), Inches(0.95), col)
        set_shape_name(circ, f"mocCirc{i}")
        nb = slide.shapes.add_textbox(x + Inches(0.95), Inches(2.17), Inches(0.95), Inches(0.5))
        add_text(nb, num, size=20, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(nb, f"mocNum{i}")
        tb = slide.shapes.add_textbox(x + Inches(0.15), Inches(2.95), Inches(2.65), Inches(0.4))
        add_text(tb, name, size=14, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(tb, f"mocName{i}")
        if i < 3:
            arr = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                x + Inches(2.55), Inches(2.3),
                Inches(0.35), Inches(0.25),
            )
            fill_shape(arr, SOFT_GRAY)
            set_shape_name(arr, f"mocArrow{i}")

    # Procedures for managing change (verbatim)
    procs = [
        "Risk assessment",
        "Identifying the goals, objectives and nature of the proposed change",
        "Identifying operational procedures",
        "Analyzing changes in location, equipment or operating conditions",
        "Posting current changes in maintenance and operator manuals",
        "All personnel being made aware of and understanding changes",
        "Identifying the level of management with authority to approve changes",
        "Reviewing, evaluating and recording potential safety hazards from the change or its implementation",
        "Approval of the agreed change and the implementation procedure(s)",
    ]
    for i, p in enumerate(procs):
        x = Inches(0.7 + (i % 2) * 6.15)
        y = Inches(3.55 + (i // 2) * 0.72)
        chip = round_rect(slide, x, y, Inches(5.95), Inches(0.62), WHITE)
        set_shape_name(chip, f"mocProc{i}")
        tb = slide.shapes.add_textbox(x + Inches(0.2), y + Inches(0.07), Inches(5.6), Inches(0.5))
        add_text(tb, f"{i+1:02d}   {p}", size=10.5, color=NAVY, bold=True)
        set_shape_name(tb, f"mocProcText{i}")

    add_morph_transition(slide)
    return slide


def slide_18_improvement(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Component 4")
    slide_title(slide, "Continuous Improvement")

    # Cycle diagram - 4 ovals
    cycle = [
        (Inches(5.5), Inches(1.7), TEAL, "Safety Risk\nManagement"),
        (Inches(8.8), Inches(3.3), GOLD, "Safety\nAssurance"),
        (Inches(5.5), Inches(4.9), ORANGE, "Lessons\nLearned"),
        (Inches(2.2), Inches(3.3), BLUE, "Communicate\nto All"),
    ]
    for i, (x, y, col, label) in enumerate(cycle):
        c = oval(slide, x, y, Inches(2.6), Inches(1.6), col)
        set_shape_name(c, f"cycle{i}")
        tb = slide.shapes.add_textbox(x + Inches(0.15), y + Inches(0.4), Inches(2.3), Inches(0.9))
        add_text(tb, label, size=14, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(tb, f"cycleText{i}")

    # Center
    center = oval(slide, Inches(5.7), Inches(3.4), Inches(2.2), Inches(1.4), NAVY)
    set_shape_name(center, "cycleCenter")
    ct = slide.shapes.add_textbox(Inches(5.7), Inches(3.75), Inches(2.2), Inches(0.7))
    add_text(ct, "SMS\nIMPROVE", size=14, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
    set_shape_name(ct, "cycleCenterText")

    note = slide.shapes.add_textbox(Inches(0.7), Inches(6.55), Inches(12), Inches(0.6))
    add_text(note, "Continual improvement through recurring application of Safety Risk Management, Safety Assurance — using safety lessons learned and communicating them to all personnel.", size=12, color=SLATE, align=PP_ALIGN.CENTER)
    set_shape_name(note, "improveNote")

    add_morph_transition(slide)
    return slide


def slide_19_promotion(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "People")
    slide_title(slide, "Safety Promotion & Culture")

    pillars = [
        (NAVY, "TRAINING & EDUCATION", "Minimum safety training to all the employees"),
        (TEAL, "SMS TRAINING", "All personnel should be given introductory and recurrent SMS training"),
        (GOLD, "SAFETY CULTURE", "Shared values that make safe choices the normal way of working"),
        (ORANGE, "PROMOTION", "Visible leadership, open reporting and recognition"),
    ]
    for i, (col, title, desc) in enumerate(pillars):
        x = Inches(0.55 + i * 3.2)
        card = round_rect(slide, x, Inches(1.9), Inches(3.0), Inches(4.4), WHITE)
        set_shape_name(card, f"promoCard{i}")
        top = rect(slide, x, Inches(1.9), Inches(3.0), Inches(1.6), col)
        set_shape_name(top, f"promoTop{i}")
        # Icon circle
        circ = oval(slide, x + Inches(0.9), Inches(2.2), Inches(1.2), Inches(1.2), WHITE)
        set_shape_name(circ, f"promoCirc{i}")
        icons = ["✎", "◎", "★", "▲"]
        ib = slide.shapes.add_textbox(x + Inches(0.9), Inches(2.5), Inches(1.2), Inches(0.6))
        add_text(ib, icons[i], size=22, color=col, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(ib, f"promoIcon{i}")
        tb = slide.shapes.add_textbox(x + Inches(0.2), Inches(3.8), Inches(2.6), Inches(0.7))
        add_text(tb, title, size=14, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        set_shape_name(tb, f"promoTitle{i}")
        db = slide.shapes.add_textbox(x + Inches(0.25), Inches(4.7), Inches(2.5), Inches(1.2))
        add_text(db, desc, size=13, color=SLATE, align=PP_ALIGN.CENTER)
        set_shape_name(db, f"promoDesc{i}")

    add_morph_transition(slide)
    return slide


def slide_20_training_chart(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Capability")
    slide_title(slide, "Training Coverage Target")

    chart_data = CategoryChartData()
    chart_data.categories = ["All Employees\n(Minimum Safety)", "Introductory\nSMS", "Recurrent\nSMS", "Leadership\nSafety Brief"]
    chart_data.add_series("Target %", (100, 100, 100, 100))
    chart_data.add_series("Current %", (92, 78, 65, 88))

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.7), Inches(1.6), Inches(8.5), Inches(5.2),
        chart_data,
    )
    chart = chart_shape.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM

    for s_idx, col in enumerate([GOLD, NAVY]):
        ser = chart.series[s_idx]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = col

    # Callout
    call = round_rect(slide, Inches(9.5), Inches(2.5), Inches(3.3), Inches(3.2), NAVY)
    set_shape_name(call, "trainCall")
    ct = slide.shapes.add_textbox(Inches(9.7), Inches(2.9), Inches(2.9), Inches(2.5))
    tf = add_text(ct, "GOAL", size=12, color=GOLD, bold=True, align=PP_ALIGN.CENTER)
    add_para(tf, "100%", size=40, color=WHITE, bold=True, align=PP_ALIGN.CENTER, space_before=12)
    add_para(tf, "of personnel trained\n& current on SMS", size=13, color=RGBColor(0xA8, 0xB8, 0xC8), align=PP_ALIGN.CENTER, space_before=10)
    set_shape_name(ct, "trainCallText")

    add_morph_transition(slide)
    return slide


def slide_21_summary(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    accent_bar(slide)
    section_label(slide, "Executive Summary")
    slide_title(slide, "What We Need From Leadership")

    asks = [
        (GOLD, "01", "Endorse", "Approve the Safety Policy & SMS objectives"),
        (TEAL, "02", "Resource", "Fund the Safety Manager role & tools"),
        (BLUE, "03", "Own", "Hold line managers accountable as SMS owners"),
        (ORANGE, "04", "Champion", "Visible leadership for reporting & culture"),
    ]
    for i, (col, num, title, desc) in enumerate(asks):
        x = Inches(0.6 + i * 3.15)
        card = round_rect(slide, x, Inches(1.9), Inches(3.0), Inches(4.2), WHITE)
        set_shape_name(card, f"askCard{i}")
        top = rect(slide, x, Inches(1.9), Inches(3.0), Inches(0.12), col)
        set_shape_name(top, f"askTop{i}")
        nb = slide.shapes.add_textbox(x + Inches(0.25), Inches(2.3), Inches(2.5), Inches(0.5))
        add_text(nb, num, size=28, color=col, bold=True)
        set_shape_name(nb, f"askNum{i}")
        tb = slide.shapes.add_textbox(x + Inches(0.25), Inches(3.1), Inches(2.5), Inches(0.5))
        add_text(tb, title, size=22, color=NAVY, bold=True)
        set_shape_name(tb, f"askTitle{i}")
        db = slide.shapes.add_textbox(x + Inches(0.25), Inches(3.9), Inches(2.5), Inches(1.5))
        add_text(db, desc, size=14, color=SLATE)
        set_shape_name(db, f"askDesc{i}")

    add_morph_transition(slide)
    return slide


def slide_22_close(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    rect(slide, Inches(0), Inches(0), Inches(0.12), SLIDE_H, GOLD)

    label = slide.shapes.add_textbox(Inches(0.9), Inches(2.2), Inches(11), Inches(0.4))
    add_text(label, "NEXT STEP", size=14, color=GOLD, bold=True)
    set_shape_name(label, "closeLabel")

    title = slide.shapes.add_textbox(Inches(0.9), Inches(2.8), Inches(11.5), Inches(1.2))
    add_text(title, "Approve. Resource. Lead.", size=40, color=WHITE, bold=True, font="Segoe UI Semibold")
    set_shape_name(title, "closeTitle")

    sub = slide.shapes.add_textbox(Inches(0.9), Inches(4.2), Inches(10), Inches(0.8))
    add_text(sub, "A continuous reduction in mishaps starts with management commitment today.", size=16, color=RGBColor(0xA8, 0xB8, 0xC8))
    set_shape_name(sub, "closeSub")

    for i, c in enumerate([GOLD, TEAL, BLUE]):
        o = oval(slide, Inches(0.9 + i * 0.35), Inches(5.5), Inches(0.18), Inches(0.18), c)
        set_shape_name(o, f"closeDot{i}")

    add_fade_transition(slide)
    return slide


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    builders = [
        slide_01_title,
        slide_02_agenda,
        slide_03_policy,
        slide_04_objective_intro,
        slide_05_pyramid,
        slide_06_commitment,
        slide_07_swiss_cheese,
        slide_08_key_personnel,
        slide_09_erp,
        slide_10_documentation,
        slide_11_risk_strategies,
        slide_12_hazard_reporting,
        slide_13_risk_matrix,
        slide_14_risk_actions,
        slide_15_assurance,
        slide_16_spi_chart,
        slide_17_moc,
        slide_18_improvement,
        slide_19_promotion,
        slide_20_training_chart,
        slide_21_summary,
        slide_22_close,
    ]
    for build in builders:
        slide = build(prs)
        # Staggered fade-in build for everything except the background fill.
        add_entrance_animations(slide, skip=1)

    out = "/workspace/SMS_Management_Briefing.pptx"
    prs.save(out)
    print(f"Saved: {out}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
