#!/usr/bin/env python3
"""
Restyle ESG Ready Módulo 2 presentation with ISI (Institute for Sustainable Innovation) branding.

Transformations:
  - Typography: All text → Syne family (ExtraBold, Bold, SemiBold, Regular)
  - Corner rounding: All rectangular shapes and images get rounded corners
  - Color palette: Only #FFFFFF, #2A3EF4, #150047, #95EBDA, #FF2990
  - Logo: ISI logo on every slide (color variant based on background)
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
ERROR_PINK = RGBColor(0xFF, 0x29, 0x90)

# ── Color Mapping (old → new) ─────────────────────────────────
# Map existing colors to ISI palette
COLOR_MAP = {
    # Dark backgrounds/text
    (0x15, 0x10, 0x50): DARK_NAVY,      # #151050 → #150047
    (0x1E, 0x29, 0x3B): DARK_NAVY,      # #1E293B → #150047
    (0x33, 0x41, 0x55): DARK_NAVY,      # #334155 → #150047

    # Accent teal/mint colors
    (0x5C, 0xE0, 0xD2): MINT,           # #5CE0D2 → #95EBDA
    (0x00, 0xB4, 0xA0): MINT,           # #00B4A0 → #95EBDA

    # Grays → context-dependent (subtitle gray → lighter treatment)
    (0x94, 0xA3, 0xB8): MINT,           # #94A3B8 subtitle gray → mint on dark, dark on light

    # White stays white
    (0xFF, 0xFF, 0xFF): WHITE,

    # Light background
    (0xF5, 0xF7, 0xFA): WHITE,          # #F5F7FA → #FFFFFF
}

# For shapes with solid fills (decorative elements)
FILL_COLOR_MAP = {
    (0x15, 0x10, 0x50): DARK_NAVY,
    (0x5C, 0xE0, 0xD2): MINT,
    (0x00, 0xB4, 0xA0): MINT,
    (0xF5, 0xF7, 0xFA): WHITE,
    (0x33, 0x41, 0x55): DARK_NAVY,
    (0x1E, 0x29, 0x3B): DARK_NAVY,
    (0x94, 0xA3, 0xB8): BLUE,           # gray decorative → blue
}

# Corner rounding radius (fraction of shorter side)
CORNER_RADIUS = 0.04

# Logo dimensions
LOGO_WIDTH = Emu(1097280)   # ~1.2 inches
LOGO_HEIGHT = Emu(274320)   # ~0.3 inches (aspect ratio of the logo)

# Logo position: bottom-left, consistent with ISI dossier reference
LOGO_LEFT = Emu(457200)     # ~0.5 inches from left
LOGO_BOTTOM_MARGIN = Emu(274320)  # ~0.3 inches from bottom

# Paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, "ESG_Ready_Modulo2_Presentacion.pptx")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "ESG_Ready_Modulo2_ISI.pptx")
LOGO_BLACK = os.path.join(SCRIPT_DIR, "LOGO ISI BLACK.png")
LOGO_CLARO = os.path.join(SCRIPT_DIR, "LOGO ISI CLARO.png")
LOGO_WHITE = os.path.join(SCRIPT_DIR, "logo_isi_WHITE.png")


def rgb_tuple(color):
    """Extract (r, g, b) tuple from an RGBColor."""
    if color is None:
        return None
    return (color[0], color[1], color[2])


def map_text_color(old_rgb, is_dark_bg):
    """Map a text color to the ISI palette, considering background."""
    if old_rgb is None:
        return None
    t = rgb_tuple(old_rgb)
    # Subtitle gray gets special treatment based on background
    if t == (0x94, 0xA3, 0xB8):
        return MINT if is_dark_bg else DARK_NAVY
    if t in COLOR_MAP:
        return COLOR_MAP[t]
    # Fallback: if it's a light color on dark bg, keep white; if dark color on light bg, use navy
    brightness = (t[0] * 299 + t[1] * 587 + t[2] * 114) / 1000
    if is_dark_bg:
        return WHITE if brightness > 128 else MINT
    else:
        return DARK_NAVY if brightness < 128 else WHITE


def map_fill_color(old_rgb):
    """Map a fill color to the ISI palette."""
    if old_rgb is None:
        return None
    t = rgb_tuple(old_rgb)
    if t in FILL_COLOR_MAP:
        return FILL_COLOR_MAP[t]
    # Fallback by brightness
    brightness = (t[0] * 299 + t[1] * 587 + t[2] * 114) / 1000
    if brightness > 200:
        return WHITE
    elif brightness > 100:
        return BLUE
    else:
        return DARK_NAVY


def get_slide_bg_color(slide):
    """Determine the background color of a slide."""
    try:
        bg = slide.background
        fill = bg.fill
        if fill.type is not None and fill.type == 1:  # SOLID
            return rgb_tuple(fill.fore_color.rgb)
    except Exception:
        pass
    return (0xFF, 0xFF, 0xFF)  # default white


def is_dark_background(bg_tuple):
    """Check if a background color is dark."""
    brightness = (bg_tuple[0] * 299 + bg_tuple[1] * 587 + bg_tuple[2] * 114) / 1000
    return brightness < 128


def get_font_name_for_role(original_font, original_bold, original_size):
    """Determine the Syne variant based on the original font role."""
    size_pt = original_size / 12700 if original_size else 14

    # Arial Black → big titles, numbers → Syne ExtraBold
    if original_font and "arial black" in original_font.lower():
        return "Syne ExtraBold", False  # (font_name, bold_flag)

    # Large bold text (titles) → Syne Bold
    if original_bold and size_pt >= 20:
        return "Syne Bold", False

    # Medium bold text (subtitles, labels) → Syne SemiBold
    if original_bold and size_pt >= 14:
        return "Syne SemiBold", False

    # Small bold text → Syne Medium
    if original_bold:
        return "Syne Medium", False

    # Regular text
    return "Syne", False


def restyle_run(run, is_dark_bg):
    """Apply ISI font and color to a single text run."""
    font = run.font
    original_font = font.name
    original_bold = font.bold
    original_size = font.size

    # Determine Syne variant
    font_name, bold = get_font_name_for_role(original_font, original_bold, original_size)
    font.name = font_name
    font.bold = bold

    # Also set East Asian and Complex Script fonts to Syne
    rPr = run._r.get_or_add_rPr()
    # Set a:ea (East Asian) font
    ea = rPr.find(qn('a:ea'))
    if ea is not None:
        ea.set('typeface', font_name)
    else:
        ea = rPr.makeelement(qn('a:ea'), {'typeface': font_name})
        rPr.append(ea)
    # Set a:cs (Complex Script) font
    cs = rPr.find(qn('a:cs'))
    if cs is not None:
        cs.set('typeface', font_name)
    else:
        cs = rPr.makeelement(qn('a:cs'), {'typeface': font_name})
        rPr.append(cs)

    # Map color
    try:
        if font.color and font.color.rgb:
            new_color = map_text_color(font.color.rgb, is_dark_bg)
            if new_color:
                font.color.rgb = new_color
    except Exception:
        pass


def restyle_text_frame(text_frame, is_dark_bg):
    """Apply ISI styling to all text in a text frame."""
    for para in text_frame.paragraphs:
        for run in para.runs:
            restyle_run(run, is_dark_bg)


def round_freeform_corners(shape):
    """
    For freeform shapes that act as rectangles/cards,
    try to convert them to rounded rectangles via XML manipulation.
    Only apply if shape is roughly rectangular.
    """
    sp = shape._element
    # Check if it's a large enough shape to benefit from rounding
    w = shape.width
    h = shape.height
    if w < Emu(200000) or h < Emu(200000):
        return  # Too small (decorative lines/bars)

    # For thin bars (aspect ratio > 10:1 or < 1:10), skip rounding
    if w > 0 and h > 0:
        ratio = max(w / h, h / w)
        if ratio > 8:
            return  # It's a bar/line, not a card

    # For freeform shapes, we can't easily convert to roundRect
    # since they have custom geometry. Instead, we'll leave them as-is
    # if they're genuinely freeform. But many "freeform" shapes from
    # Google Slides are actually just rectangles with custom paths.
    # We'll handle this at the XML level.
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return

    custGeom = spPr.find(qn('a:custGeom'))
    if custGeom is not None:
        # Replace custom geometry with preset rounded rectangle
        prstGeom = spPr.makeelement(qn('a:prstGeom'), {'prst': 'roundRect'})
        avLst = prstGeom.makeelement(qn('a:avLst'), {})
        gd = avLst.makeelement(qn('a:gd'), {
            'name': 'adj',
            'fmla': f'val {int(CORNER_RADIUS * 100000)}'
        })
        avLst.append(gd)
        prstGeom.append(avLst)
        spPr.replace(custGeom, prstGeom)


def round_autoshape_corners(shape):
    """Round corners of an AutoShape."""
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
        # Set adjustment value
        avLst = prstGeom.find(qn('a:avLst'))
        if avLst is None:
            avLst = prstGeom.makeelement(qn('a:avLst'), {})
            prstGeom.append(avLst)
        # Clear existing adjustments
        for gd in avLst.findall(qn('a:gd')):
            avLst.remove(gd)
        gd = avLst.makeelement(qn('a:gd'), {
            'name': 'adj',
            'fmla': f'val {int(CORNER_RADIUS * 100000)}'
        })
        avLst.append(gd)


def round_picture_corners(shape):
    """Round corners of a picture by changing its geometry mask."""
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return

    prstGeom = spPr.find(qn('a:prstGeom'))
    if prstGeom is None:
        # Create prstGeom for the picture
        prstGeom = spPr.makeelement(qn('a:prstGeom'), {'prst': 'roundRect'})
        avLst = prstGeom.makeelement(qn('a:avLst'), {})
        gd = avLst.makeelement(qn('a:gd'), {
            'name': 'adj',
            'fmla': f'val {int(CORNER_RADIUS * 100000)}'
        })
        avLst.append(gd)
        prstGeom.append(avLst)
        # Insert after xfrm if present, or as first child
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
            'name': 'adj',
            'fmla': f'val {int(CORNER_RADIUS * 100000)}'
        })
        avLst.append(gd)


def recolor_solid_fill(element, new_color):
    """Set a solid fill color on an XML element's solidFill."""
    solidFill = element.find(qn('a:solidFill'))
    if solidFill is not None:
        # Remove existing color children
        for child in list(solidFill):
            solidFill.remove(child)
        srgb = solidFill.makeelement(qn('a:srgbClr'), {
            'val': f'{new_color[0]:02X}{new_color[1]:02X}{new_color[2]:02X}'
        })
        solidFill.append(srgb)
        return True
    return False


def recolor_shape_fill(shape, is_dark_bg):
    """Recolor the fill of a shape to ISI palette."""
    try:
        fill = shape.fill
        if fill.type == 1:  # SOLID
            old_rgb = fill.fore_color.rgb
            new_color = map_fill_color(old_rgb)
            if new_color:
                fill.fore_color.rgb = new_color
    except Exception:
        pass


def recolor_freeform_fill(shape, is_dark_bg):
    """Recolor freeform shapes via XML."""
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
                new_color = map_fill_color(RGBColor(r, g, b))
                if new_color:
                    srgb.set('val', f'{new_color[0]:02X}{new_color[1]:02X}{new_color[2]:02X}')


def recolor_line(shape):
    """Recolor shape border/line if present."""
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
                    new_color = map_fill_color(RGBColor(r, g, b))
                    if new_color:
                        srgb.set('val', f'{new_color[0]:02X}{new_color[1]:02X}{new_color[2]:02X}')


def is_decorative_freeform(shape, slide_width, slide_height):
    """
    Detect if a freeform is a decorative/brand element (circles, abstract shapes)
    vs a content card/container that should get rounded corners.

    Decorative indicators:
    - Partially off-screen (negative position or extends beyond slide)
    - Very large relative to slide (>40% of slide in both dimensions)
    - Nearly circular (aspect ratio ~1:1 and large)
    - On the top edge spanning full width (top bar)
    """
    left = shape.left
    top = shape.top
    w = shape.width
    h = shape.height

    # Off-screen: negative position or extends beyond slide
    if left < 0 or top < 0:
        return True
    if left + w > slide_width + Emu(50000):  # small tolerance
        return True
    if top + h > slide_height + Emu(50000):
        return True

    # Very large shapes (>40% of slide in both dims) are usually decorative
    if w > slide_width * 0.4 and h > slide_height * 0.4:
        return True

    # Nearly circular and large → decorative (e.g., the leaf, circles)
    if w > 0 and h > 0:
        aspect = max(w / h, h / w)
        if aspect < 1.3 and min(w, h) > Emu(900000):  # ~1 inch, nearly square
            return True

    # Full-width top bar
    if w > slide_width * 0.9 and h < Emu(100000):
        return True

    return False


def should_round_shape(shape, slide_width=None, slide_height=None):
    """Determine if a shape should get rounded corners."""
    w = shape.width
    h = shape.height
    # Skip very small shapes
    if w < Emu(200000) or h < Emu(200000):
        return False
    # Skip thin bars (aspect ratio > 8:1)
    if w > 0 and h > 0:
        ratio = max(w / h, h / w)
        if ratio > 8:
            return False
    # For freeforms, check if it's a decorative element
    if shape.shape_type == MSO_SHAPE_TYPE.FREEFORM and slide_width and slide_height:
        if is_decorative_freeform(shape, slide_width, slide_height):
            return False
    return True


def process_shape(shape, is_dark_bg, slide_width, slide_height):
    """Process a single shape: restyle text, round corners, recolor."""
    st = shape.shape_type

    # ── Text restyling ──
    if shape.has_text_frame:
        restyle_text_frame(shape.text_frame, is_dark_bg)

    # ── Corner rounding ──
    if st == MSO_SHAPE_TYPE.AUTO_SHAPE and should_round_shape(shape, slide_width, slide_height):
        round_autoshape_corners(shape)

    if st == MSO_SHAPE_TYPE.PICTURE and should_round_shape(shape, slide_width, slide_height):
        round_picture_corners(shape)

    if st == MSO_SHAPE_TYPE.FREEFORM and should_round_shape(shape, slide_width, slide_height):
        round_freeform_corners(shape)

    # ── Fill recoloring ──
    if st == MSO_SHAPE_TYPE.FREEFORM:
        recolor_freeform_fill(shape, is_dark_bg)
    elif st in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.TEXT_BOX):
        recolor_shape_fill(shape, is_dark_bg)

    # ── Line/border recoloring ──
    recolor_line(shape)

    # ── Recurse into groups ──
    if st == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            process_shape(child, is_dark_bg, slide_width, slide_height)


def set_slide_background(slide, bg_color):
    """Set slide background to a solid color."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = bg_color


def has_footer_bar(slide, slide_height):
    """Check if slide has a dark footer bar at the bottom."""
    for shape in slide.shapes:
        # Footer bars are typically full-width shapes near the bottom
        if shape.top + shape.height >= slide_height - Emu(50000):
            if shape.width > Emu(8000000):  # nearly full width
                if shape.height < Emu(600000):  # and relatively thin
                    return True
    return False


def add_logo_to_slide(slide, prs, is_dark_bg):
    """Add the appropriate ISI logo to the slide."""
    # Choose logo variant based on background
    if is_dark_bg:
        logo_path = LOGO_CLARO  # mint/turquoise on dark backgrounds
    else:
        logo_path = LOGO_BLACK  # black on light/white backgrounds

    if not os.path.exists(logo_path):
        print(f"  Warning: Logo not found at {logo_path}")
        return

    # If the slide has a footer bar, place logo above it
    footer = has_footer_bar(slide, prs.slide_height)
    if footer and not is_dark_bg:
        # Place above the footer bar (footer is ~434340 EMU high at the bottom)
        logo_top = prs.slide_height - Emu(434340) - LOGO_HEIGHT - Emu(91440)
    else:
        logo_top = prs.slide_height - LOGO_HEIGHT - LOGO_BOTTOM_MARGIN

    slide.shapes.add_picture(logo_path, LOGO_LEFT, logo_top, LOGO_WIDTH, LOGO_HEIGHT)


def process_presentation():
    """Main processing function."""
    print(f"Loading: {INPUT_FILE}")
    prs = Presentation(INPUT_FILE)
    print(f"Slides: {len(prs.slides)}")
    print(f"Slide size: {prs.slide_width/914400:.2f}\" x {prs.slide_height/914400:.2f}\"")

    for i, slide in enumerate(prs.slides):
        bg_tuple = get_slide_bg_color(slide)
        is_dark_bg = is_dark_background(bg_tuple)

        print(f"\nSlide {i+1}: bg={'dark' if is_dark_bg else 'light'} (#{bg_tuple[0]:02X}{bg_tuple[1]:02X}{bg_tuple[2]:02X})")

        # ── Remap slide background ──
        if is_dark_bg:
            set_slide_background(slide, DARK_NAVY)
        else:
            set_slide_background(slide, WHITE)

        # ── Process all shapes ──
        for shape in slide.shapes:
            process_shape(shape, is_dark_bg, prs.slide_width, prs.slide_height)

        # ── Add logo ──
        add_logo_to_slide(slide, prs, is_dark_bg)

    # ── Save output ──
    print(f"\nSaving: {OUTPUT_FILE}")
    prs.save(OUTPUT_FILE)
    print("Done!")


if __name__ == "__main__":
    process_presentation()
