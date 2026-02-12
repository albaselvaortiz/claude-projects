#!/usr/bin/env python3
"""
Restyle ESG Ready Módulo 2 presentation with ISI branding.

Changes v2:
  - Typography: Syne Bold (no ExtraBold), SemiBold, Medium, Regular
  - Corner rounding: increased radius (8%)
  - More blue #2A3EF4: footer bars, callout boxes, accent bars on light slides
  - Vector watermark: ISI logo at ~15% opacity on dark cover/section slides
  - Fix slide 24 overlapping elements
  - Google Slides compatible (no transitions/animations)
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
import copy
import os

# ── ISI Color Palette ──────────────────────────────────────────
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLUE = RGBColor(0x2A, 0x3E, 0xF4)
DARK_NAVY = RGBColor(0x15, 0x00, 0x47)
MINT = RGBColor(0x95, 0xEB, 0xDA)

# ── Color Mapping ──────────────────────────────────────────────
COLOR_MAP = {
    (0x15, 0x10, 0x50): DARK_NAVY,
    (0x1E, 0x29, 0x3B): DARK_NAVY,
    (0x33, 0x41, 0x55): DARK_NAVY,
    (0x5C, 0xE0, 0xD2): MINT,
    (0x00, 0xB4, 0xA0): MINT,
    (0x94, 0xA3, 0xB8): MINT,
    (0xFF, 0xFF, 0xFF): WHITE,
    (0xF5, 0xF7, 0xFA): WHITE,
}

# Fill mapping for shapes on LIGHT background slides → use more BLUE
FILL_COLOR_MAP_LIGHT = {
    (0x15, 0x10, 0x50): BLUE,       # dark fills → BLUE (footer, callouts, highlights)
    (0x5C, 0xE0, 0xD2): BLUE,       # teal accent bars → BLUE
    (0x00, 0xB4, 0xA0): BLUE,       # green accents → BLUE
    (0xF5, 0xF7, 0xFA): WHITE,
    (0x33, 0x41, 0x55): BLUE,
    (0x1E, 0x29, 0x3B): BLUE,
    (0x94, 0xA3, 0xB8): BLUE,
}

# Fill mapping for shapes on DARK background slides → keep navy/mint
FILL_COLOR_MAP_DARK = {
    (0x15, 0x10, 0x50): DARK_NAVY,
    (0x5C, 0xE0, 0xD2): MINT,
    (0x00, 0xB4, 0xA0): MINT,
    (0xF5, 0xF7, 0xFA): WHITE,
    (0x33, 0x41, 0x55): DARK_NAVY,
    (0x1E, 0x29, 0x3B): DARK_NAVY,
    (0x94, 0xA3, 0xB8): BLUE,
}

# Corner rounding radius — increased to 8%
CORNER_RADIUS = 0.08

# Logo dimensions
LOGO_WIDTH = Emu(1097280)
LOGO_HEIGHT = Emu(274320)
LOGO_LEFT = Emu(457200)
LOGO_BOTTOM_MARGIN = Emu(274320)

# Watermark dimensions (large, bottom-right)
WATERMARK_SIZE = Emu(2743200)   # ~3 inches
WATERMARK_OPACITY = 15          # 15% opacity

# Paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, "ESG_Ready_Modulo2_Presentacion.pptx")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "claude_ESG_Ready_Modulo2_ISI.pptx")
LOGO_BLACK = os.path.join(SCRIPT_DIR, "LOGO ISI BLACK.png")
LOGO_CLARO = os.path.join(SCRIPT_DIR, "LOGO ISI CLARO.png")
LOGO_WHITE = os.path.join(SCRIPT_DIR, "logo_isi_WHITE.png")


def rgb_tuple(color):
    if color is None:
        return None
    return (color[0], color[1], color[2])


def map_text_color(old_rgb, is_dark_bg):
    if old_rgb is None:
        return None
    t = rgb_tuple(old_rgb)
    if t == (0x94, 0xA3, 0xB8):
        return MINT if is_dark_bg else DARK_NAVY
    if t in COLOR_MAP:
        return COLOR_MAP[t]
    brightness = (t[0] * 299 + t[1] * 587 + t[2] * 114) / 1000
    if is_dark_bg:
        return WHITE if brightness > 128 else MINT
    else:
        return DARK_NAVY if brightness < 128 else WHITE


def map_fill_color(old_rgb, is_dark_bg):
    if old_rgb is None:
        return None
    t = rgb_tuple(old_rgb)
    fill_map = FILL_COLOR_MAP_DARK if is_dark_bg else FILL_COLOR_MAP_LIGHT
    if t in fill_map:
        return fill_map[t]
    brightness = (t[0] * 299 + t[1] * 587 + t[2] * 114) / 1000
    if brightness > 200:
        return WHITE
    elif brightness > 100:
        return BLUE
    else:
        return DARK_NAVY


def get_slide_bg_color(slide):
    try:
        bg = slide.background
        fill = bg.fill
        if fill.type is not None and fill.type == 1:
            return rgb_tuple(fill.fore_color.rgb)
    except Exception:
        pass
    return (0xFF, 0xFF, 0xFF)


def is_dark_background(bg_tuple):
    brightness = (bg_tuple[0] * 299 + bg_tuple[1] * 587 + bg_tuple[2] * 114) / 1000
    return brightness < 128


def get_font_name_for_role(original_font, original_bold, original_size):
    """Syne Bold (no ExtraBold) for titles, SemiBold for subtitles, etc."""
    size_pt = original_size / 12700 if original_size else 14

    # Arial Black → was ExtraBold, now just Bold
    if original_font and "arial black" in original_font.lower():
        return "Syne Bold", False

    # Large bold text (titles)
    if original_bold and size_pt >= 20:
        return "Syne Bold", False

    # Medium bold text (subtitles, labels)
    if original_bold and size_pt >= 14:
        return "Syne SemiBold", False

    # Small bold text
    if original_bold:
        return "Syne Medium", False

    # Regular text
    return "Syne", False


def restyle_run(run, is_dark_bg):
    font = run.font
    original_font = font.name
    original_bold = font.bold
    original_size = font.size

    font_name, bold = get_font_name_for_role(original_font, original_bold, original_size)
    font.name = font_name
    font.bold = bold

    rPr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is not None:
            el.set('typeface', font_name)
        else:
            el = rPr.makeelement(qn(tag), {'typeface': font_name})
            rPr.append(el)

    try:
        if font.color and font.color.rgb:
            new_color = map_text_color(font.color.rgb, is_dark_bg)
            if new_color:
                font.color.rgb = new_color
    except Exception:
        pass


def restyle_text_frame(text_frame, is_dark_bg):
    for para in text_frame.paragraphs:
        for run in para.runs:
            restyle_run(run, is_dark_bg)


def round_freeform_corners(shape):
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    custGeom = spPr.find(qn('a:custGeom'))
    if custGeom is not None:
        prstGeom = spPr.makeelement(qn('a:prstGeom'), {'prst': 'roundRect'})
        avLst = prstGeom.makeelement(qn('a:avLst'), {})
        gd = avLst.makeelement(qn('a:gd'), {
            'name': 'adj', 'fmla': f'val {int(CORNER_RADIUS * 100000)}'
        })
        avLst.append(gd)
        prstGeom.append(avLst)
        spPr.replace(custGeom, prstGeom)


def round_autoshape_corners(shape):
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    prstGeom = spPr.find(qn('a:prstGeom'))
    if prstGeom is None:
        return
    prst = prstGeom.get('prst')
    if prst in ('rect', 'roundRect'):
        prstGeom.set('prst', 'roundRect')
        avLst = prstGeom.find(qn('a:avLst'))
        if avLst is None:
            avLst = prstGeom.makeelement(qn('a:avLst'), {})
            prstGeom.append(avLst)
        for gd in avLst.findall(qn('a:gd')):
            avLst.remove(gd)
        gd = avLst.makeelement(qn('a:gd'), {
            'name': 'adj', 'fmla': f'val {int(CORNER_RADIUS * 100000)}'
        })
        avLst.append(gd)


def round_picture_corners(shape):
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    prstGeom = spPr.find(qn('a:prstGeom'))
    if prstGeom is None:
        prstGeom = spPr.makeelement(qn('a:prstGeom'), {'prst': 'roundRect'})
        avLst = prstGeom.makeelement(qn('a:avLst'), {})
        gd = avLst.makeelement(qn('a:gd'), {
            'name': 'adj', 'fmla': f'val {int(CORNER_RADIUS * 100000)}'
        })
        avLst.append(gd)
        prstGeom.append(avLst)
        xfrm = spPr.find(qn('a:xfrm'))
        if xfrm is not None:
            xfrm.addnext(prstGeom)
        else:
            spPr.insert(0, prstGeom)
    else:
        prstGeom.set('prst', 'roundRect')
        avLst = prstGeom.find(qn('a:avLst'))
        if avLst is None:
            avLst = prstGeom.makeelement(qn('a:avLst'), {})
            prstGeom.append(avLst)
        for gd in avLst.findall(qn('a:gd')):
            avLst.remove(gd)
        gd = avLst.makeelement(qn('a:gd'), {
            'name': 'adj', 'fmla': f'val {int(CORNER_RADIUS * 100000)}'
        })
        avLst.append(gd)


def recolor_shape_fill(shape, is_dark_bg):
    try:
        fill = shape.fill
        if fill.type == 1:  # SOLID
            old_rgb = fill.fore_color.rgb
            new_color = map_fill_color(old_rgb, is_dark_bg)
            if new_color:
                fill.fore_color.rgb = new_color
    except Exception:
        pass


def recolor_freeform_fill(shape, is_dark_bg):
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    solidFill = spPr.find(qn('a:solidFill'))
    if solidFill is not None:
        srgb = solidFill.find(qn('a:srgbClr'))
        if srgb is not None:
            val = srgb.get('val')
            if val:
                r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
                new_color = map_fill_color(RGBColor(r, g, b), is_dark_bg)
                if new_color:
                    srgb.set('val', f'{new_color[0]:02X}{new_color[1]:02X}{new_color[2]:02X}')


def recolor_line(shape):
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    ln = spPr.find(qn('a:ln'))
    if ln is not None:
        solidFill = ln.find(qn('a:solidFill'))
        if solidFill is not None:
            srgb = solidFill.find(qn('a:srgbClr'))
            if srgb is not None:
                val = srgb.get('val')
                if val:
                    r, g, b = int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)
                    new_color = map_fill_color(RGBColor(r, g, b), False)
                    if new_color:
                        srgb.set('val', f'{new_color[0]:02X}{new_color[1]:02X}{new_color[2]:02X}')


def is_decorative_freeform(shape, slide_width, slide_height):
    left = shape.left
    top = shape.top
    w = shape.width
    h = shape.height

    if left < 0 or top < 0:
        return True
    if left + w > slide_width + Emu(50000):
        return True
    if top + h > slide_height + Emu(50000):
        return True
    if w > slide_width * 0.4 and h > slide_height * 0.4:
        return True
    if w > 0 and h > 0:
        aspect = max(w / h, h / w)
        if aspect < 1.3 and min(w, h) > Emu(900000):
            return True
    if w > slide_width * 0.9 and h < Emu(100000):
        return True
    return False


def should_round_shape(shape, slide_width=None, slide_height=None):
    w = shape.width
    h = shape.height
    if w < Emu(200000) or h < Emu(200000):
        return False
    if w > 0 and h > 0:
        ratio = max(w / h, h / w)
        if ratio > 8:
            return False
    if shape.shape_type == MSO_SHAPE_TYPE.FREEFORM and slide_width and slide_height:
        if is_decorative_freeform(shape, slide_width, slide_height):
            return False
    return True


def process_shape(shape, is_dark_bg, slide_width, slide_height):
    st = shape.shape_type

    if shape.has_text_frame:
        restyle_text_frame(shape.text_frame, is_dark_bg)

    if st == MSO_SHAPE_TYPE.AUTO_SHAPE and should_round_shape(shape, slide_width, slide_height):
        round_autoshape_corners(shape)
    if st == MSO_SHAPE_TYPE.PICTURE and should_round_shape(shape, slide_width, slide_height):
        round_picture_corners(shape)
    if st == MSO_SHAPE_TYPE.FREEFORM and should_round_shape(shape, slide_width, slide_height):
        round_freeform_corners(shape)

    if st == MSO_SHAPE_TYPE.FREEFORM:
        recolor_freeform_fill(shape, is_dark_bg)
    elif st in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.TEXT_BOX):
        recolor_shape_fill(shape, is_dark_bg)

    recolor_line(shape)

    if st == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            process_shape(child, is_dark_bg, slide_width, slide_height)


def set_slide_background(slide, bg_color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = bg_color


def has_footer_bar(slide, slide_height):
    for shape in slide.shapes:
        if shape.top + shape.height >= slide_height - Emu(50000):
            if shape.width > Emu(8000000) and shape.height < Emu(600000):
                return True
    return False


def add_logo_to_slide(slide, prs, is_dark_bg, slide_index=0):
    if is_dark_bg:
        logo_path = LOGO_CLARO
    else:
        logo_path = LOGO_BLACK

    if not os.path.exists(logo_path):
        return

    footer = has_footer_bar(slide, prs.slide_height)

    # Slide 24 (index 23): place logo right of Doc card to avoid test bar overlap
    if slide_index == 23:
        logo_left = prs.slide_width - LOGO_WIDTH - Emu(457200)
        logo_top = prs.slide_height - Emu(434340) - LOGO_HEIGHT - Emu(91440)
        slide.shapes.add_picture(logo_path, logo_left, logo_top, LOGO_WIDTH, LOGO_HEIGHT)
        return

    if footer and not is_dark_bg:
        logo_top = prs.slide_height - Emu(434340) - LOGO_HEIGHT - Emu(91440)
    else:
        logo_top = prs.slide_height - LOGO_HEIGHT - LOGO_BOTTOM_MARGIN

    slide.shapes.add_picture(logo_path, LOGO_LEFT, logo_top, LOGO_WIDTH, LOGO_HEIGHT)


def add_watermark(slide, prs):
    """Add ISI logo as a large, low-opacity watermark on dark slides."""
    logo_path = LOGO_WHITE
    if not os.path.exists(logo_path):
        return

    wm_left = prs.slide_width - WATERMARK_SIZE - Emu(274320)
    wm_top = prs.slide_height - WATERMARK_SIZE - Emu(274320)

    pic = slide.shapes.add_picture(logo_path, wm_left, wm_top, WATERMARK_SIZE, WATERMARK_SIZE)

    # Set opacity via XML: add alphaModFix to the blip
    sp = pic._element
    blipFill = sp.find(qn('p:blipFill'))
    if blipFill is not None:
        blip = blipFill.find(qn('a:blip'))
        if blip is not None:
            # alphaModFix: amt is in 1/1000 of percent (15% = 15000)
            alpha = blip.makeelement(qn('a:alphaModFix'), {
                'amt': str(WATERMARK_OPACITY * 1000)
            })
            blip.append(alpha)


def fix_slide_24(slide):
    """Fix overlapping elements on slide 24 (checklist slide)."""
    # Row 3 (Documentación) overlaps with the Test bar in the original.
    # Fix: compress row 2 and row 3, then reposition test bar clearly below.

    SHIFT_R2 = Emu(80000)    # shift row 2 up
    SHIFT_R3 = Emu(200000)   # shift row 3 up
    SHRINK_R3 = Emu(250000)  # shrink row 3 card height
    TEST_TOP = Emu(4160000)  # absolute position for test bar
    TEST_H = Emu(400000)     # test bar height

    for shape in slide.shapes:
        t = shape.top

        # Row 2 shapes (top ~2057400-2560000): shift up
        if Emu(2000000) <= t < Emu(2700000):
            shape.top = t - SHIFT_R2

        # Row 3 shapes (top ~3200000-3800000): shift up + shrink card bgs
        elif Emu(3200000) <= t < Emu(3800000):
            shape.top = t - SHIFT_R3
            if shape.height > Emu(800000):
                shape.height = shape.height - SHRINK_R3

        # Test bar shapes (top ~3931920): reposition
        elif Emu(3900000) <= t < Emu(4200000):
            shape.top = TEST_TOP
            shape.height = TEST_H


# ── Section divider slide indices (0-based) ──
# These are the dark slides with "01", "02", etc.
SECTION_DIVIDERS = {2, 5, 8, 11, 13, 15, 18, 20, 22}
COVER_SLIDES = {0, 24}  # Title and closing slides


def process_presentation():
    print(f"Loading: {INPUT_FILE}")
    prs = Presentation(INPUT_FILE)
    print(f"Slides: {len(prs.slides)}")

    for i, slide in enumerate(prs.slides):
        bg_tuple = get_slide_bg_color(slide)
        is_dark_bg = is_dark_background(bg_tuple)

        print(f"Slide {i+1}: bg={'dark' if is_dark_bg else 'light'}")

        # Remap slide background
        if is_dark_bg:
            set_slide_background(slide, DARK_NAVY)
        else:
            set_slide_background(slide, WHITE)

        # Process all shapes
        for shape in slide.shapes:
            process_shape(shape, is_dark_bg, prs.slide_width, prs.slide_height)

        # Fix slide 24 overlapping
        if i == 23:
            fix_slide_24(slide)

        # Add watermark on dark cover/section slides
        if i in COVER_SLIDES or i in SECTION_DIVIDERS:
            add_watermark(slide, prs)

        # Add logo
        add_logo_to_slide(slide, prs, is_dark_bg, slide_index=i)

    print(f"\nSaving: {OUTPUT_FILE}")
    prs.save(OUTPUT_FILE)
    print("Done!")


if __name__ == "__main__":
    process_presentation()
