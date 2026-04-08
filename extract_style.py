import json
from pptx import Presentation
from pptx.util import Emu
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR

PPTX_PATH = "MRI early rectal - OncoNovo+.pptx"

prs = Presentation(PPTX_PATH)

# --- Slide dimensions ---
width_emu = prs.slide_width
height_emu = prs.slide_height
dimensions = {
    "width_emu": width_emu,
    "height_emu": height_emu,
    "width_cm": round(width_emu / 914400 * 2.54, 2),
    "height_cm": round(height_emu / 914400 * 2.54, 2),
}

# --- Theme colors ---
def get_theme_colors(prs):
    theme_colors = {}
    try:
        theme_element = prs.slide_master.element.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}theme')
        if theme_element is None:
            # Try finding via slide_master XML
            from lxml import etree
            xml = prs.slide_master.element.xml
            # Parse color map from theme
            ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
            master_el = prs.slide_master.element
            clrMap = master_el.find('.//{%s}clrMap' % ns)
            if clrMap is not None:
                theme_colors['clrMap'] = dict(clrMap.attrib)
    except Exception as e:
        theme_colors['error'] = str(e)
    return theme_colors

# --- Collect fonts and colors from shapes ---
fonts_seen = {}
colors_seen = set()

def rgb_to_hex(rgb):
    if rgb is None:
        return None
    return f"#{rgb.red:02X}{rgb.green:02X}{rgb.blue:02X}"

def process_text_frame(tf, context=""):
    for para in tf.paragraphs:
        for run in para.runs:
            font = run.font
            name = font.name
            size = font.size
            bold = font.bold
            italic = font.italic
            color = None
            try:
                if font.color and font.color.type is not None:
                    color = rgb_to_hex(font.color.rgb)
                    if color:
                        colors_seen.add(color)
            except Exception:
                pass
            key = f"{name}_{size}_{bold}_{italic}"
            if key not in fonts_seen:
                fonts_seen[key] = {
                    "name": name,
                    "size_pt": round(size / 12700, 1) if size else None,
                    "bold": bold,
                    "italic": italic,
                    "color": color,
                    "context": context,
                }

def process_shape(shape, context=""):
    # Fill color
    try:
        if shape.fill and shape.fill.type is not None:
            fg = shape.fill.fore_color
            if fg and fg.type is not None:
                c = rgb_to_hex(fg.rgb)
                if c:
                    colors_seen.add(c)
    except Exception:
        pass
    # Line color
    try:
        if shape.line and shape.line.color and shape.line.color.type is not None:
            c = rgb_to_hex(shape.line.color.rgb)
            if c:
                colors_seen.add(c)
    except Exception:
        pass
    # Text
    try:
        if shape.has_text_frame:
            process_text_frame(shape.text_frame, context)
    except Exception:
        pass
    # Table cells
    try:
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    process_text_frame(cell.text_frame, context + "/table")
    except Exception:
        pass

# Process slide master
for shape in prs.slide_master.shapes:
    process_shape(shape, "master")

# Process all slides
slide_backgrounds = []
for i, slide in enumerate(prs.slides):
    # Background
    bg_color = None
    try:
        bg = slide.background
        fill = bg.fill
        fill.fore_color  # trigger
        bg_color = rgb_to_hex(fill.fore_color.rgb)
    except Exception:
        pass
    slide_backgrounds.append({"slide": i + 1, "background_color": bg_color})

    for shape in slide.shapes:
        process_shape(shape, f"slide_{i+1}")

# Process slide layouts
layouts = []
for layout in prs.slide_layouts:
    ph_info = []
    for ph in layout.placeholders:
        ph_info.append({"idx": ph.placeholder_format.idx, "type": str(ph.placeholder_format.type), "name": ph.name})
    layouts.append({"name": layout.name, "placeholders": ph_info})

# --- Assemble result ---
result = {
    "source_file": PPTX_PATH,
    "slide_count": len(prs.slides),
    "dimensions": dimensions,
    "theme_colors": get_theme_colors(prs),
    "explicit_colors_hex": sorted(colors_seen),
    "fonts": list(fonts_seen.values()),
    "slide_backgrounds": slide_backgrounds,
    "slide_layouts": layouts,
}

with open("radiology_style_guide.json", "w") as f:
    json.dump(result, f, indent=2)

print("Done. Colors found:", sorted(colors_seen))
print("Fonts found:", len(fonts_seen))
print("Slides:", len(prs.slides))
print("Layouts:", [l["name"] for l in layouts])
