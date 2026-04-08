"""
Generate radiology_style_guide.json (enriched), radiology_style_guide.md,
and radiology_template.pptx from MRI early rectal - OncoNovo+.pptx
"""
import json, re, zipfile
from pptx import Presentation
from pptx.util import Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from lxml import etree

PPTX_PATH = "MRI early rectal - OncoNovo+.pptx"
prs = Presentation(PPTX_PATH)

# ── 1. Dimensions ─────────────────────────────────────────────────────────────
dims = {
    "width_emu":  prs.slide_width,
    "height_emu": prs.slide_height,
    "width_cm":   round(prs.slide_width  / 914400 * 2.54, 2),
    "height_cm":  round(prs.slide_height / 914400 * 2.54, 2),
}

# ── 2. Theme colors from zip ──────────────────────────────────────────────────
theme_colors = {}
with zipfile.ZipFile(PPTX_PATH) as z:
    xml = z.read("ppt/theme/theme1.xml").decode("utf-8", errors="replace")

sys_colors   = re.findall(r'<a:(\w+)>\s*<a:sysClr\s+val="\w+"\s+lastClr="([0-9A-Fa-f]{6})"', xml)
rgb_colors   = re.findall(r'<a:(\w+)>\s*<a:srgbClr\s+val="([0-9A-Fa-f]{6})"', xml)
for tag, val in sys_colors + rgb_colors:
    theme_colors[tag] = "#" + val.upper()

major_m = re.search(r'<a:majorFont>.*?<a:latin typeface="([^"]+)"', xml, re.DOTALL)
minor_m = re.search(r'<a:minorFont>.*?<a:latin typeface="([^"]+)"', xml, re.DOTALL)
heading_font = major_m.group(1) if major_m else "Calibri Light"
body_font    = minor_m.group(1) if minor_m else "Calibri"

# ── 3. Background colors ──────────────────────────────────────────────────────
bg_colors = set()
for slide in prs.slides:
    xml_bg = etree.tostring(slide.background.element, pretty_print=True).decode()
    for c in re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', xml_bg):
        bg_colors.add("#" + c.upper())

dominant_bg  = "#0F1117"   # darkest most-used background
primary_bg   = "#0F1117"
secondary_bg = "#1A1C25"

# ── 4. Font usage ─────────────────────────────────────────────────────────────
fonts_seen = {}
def process_tf(tf, ctx):
    for para in tf.paragraphs:
        for run in para.runs:
            f = run.font
            key = f"{f.name}_{f.size}_{f.bold}"
            if key not in fonts_seen:
                fonts_seen[key] = {
                    "name":    f.name,
                    "size_pt": round(f.size / 12700, 1) if f.size else None,
                    "bold":    f.bold,
                    "italic":  f.italic,
                    "context": ctx,
                }

for shape in prs.slide_master.shapes:
    if shape.has_text_frame:
        process_tf(shape.text_frame, "master")

for i, slide in enumerate(prs.slides):
    for shape in slide.shapes:
        if shape.has_text_frame:
            process_tf(shape.text_frame, f"slide_{i+1}")

# Summarise unique sizes used
sizes_used = sorted(set(v["size_pt"] for v in fonts_seen.values() if v["size_pt"]))

# ── 5. Layouts ────────────────────────────────────────────────────────────────
layouts = []
for layout in prs.slide_layouts:
    phs = [{"idx": ph.placeholder_format.idx,
             "type": str(ph.placeholder_format.type),
             "name": ph.name}
            for ph in layout.placeholders]
    layouts.append({"name": layout.name, "placeholders": phs})

# ── 6. Build enriched JSON ────────────────────────────────────────────────────
style = {
    "source_file":    PPTX_PATH,
    "slide_count":    len(prs.slides),
    "dimensions":     dims,
    "theme": {
        "heading_font": heading_font,
        "body_font":    body_font,
        "colors":       theme_colors,
    },
    "background_colors_found": sorted(bg_colors),
    "dominant_backgrounds": {
        "primary":   primary_bg,
        "secondary": secondary_bg,
    },
    "font_sizes_used_pt": sizes_used,
    "fonts_detail":       list(fonts_seen.values()),
    "slide_layouts":      layouts,
    "radiology_notes": {
        "background_strategy": "Dark backgrounds (#0F1117–#2E3140) used throughout for DICOM-image contrast",
        "min_title_size_pt":   32,
        "min_body_size_pt":    20,
        "accent_colors":       ["#4F98A3","#6CC0CC","#6DAA45","#E8A033"],
    }
}

with open("radiology_style_guide.json", "w") as f:
    json.dump(style, f, indent=2)
print("✓ radiology_style_guide.json")

# ── 7. Markdown style guide ───────────────────────────────────────────────────
md = f"""# Radiology Presentation Style Guide
*Extracted from: {PPTX_PATH} ({len(prs.slides)} slides)*

---

## Slide Dimensions
| Property | Value |
|---|---|
| Width | {dims['width_cm']} cm ({dims['width_emu']} EMU) |
| Height | {dims['height_cm']} cm ({dims['height_emu']} EMU) |
| Aspect ratio | 16:9 |

---

## Typography
| Role | Font | Notes |
|---|---|---|
| Headings / Titles | **{heading_font}** | Major font (theme) |
| Body / Content | **{body_font}** | Minor font (theme) |

### Font sizes found in presentation
`{', '.join(str(s) + 'pt' for s in sizes_used if s)}`

### Recommended minimums (lecture hall)
- Title: **≥ 32 pt**
- Body: **≥ 20 pt**
- Captions: **≥ 14 pt**

---

## Color Palette

### Theme Colors
| Role | Hex | Swatch |
|---|---|---|
| Dark 1 (text) | `{theme_colors.get('dk1','#000000')}` | ■ |
| Light 1 (bg) | `{theme_colors.get('lt1','#FFFFFF')}` | ■ |
| Dark 2 | `{theme_colors.get('dk2','#44546A')}` | ■ |
| Light 2 | `{theme_colors.get('lt2','#E7E6E6')}` | ■ |
| Accent 1 | `{theme_colors.get('accent1','#4472C4')}` | ■ |
| Accent 2 | `{theme_colors.get('accent2','#ED7D31')}` | ■ |
| Accent 3 | `{theme_colors.get('accent3','#A5A5A5')}` | ■ |
| Accent 4 | `{theme_colors.get('accent4','#FFC000')}` | ■ |
| Accent 5 | `{theme_colors.get('accent5','#5B9BD5')}` | ■ |
| Accent 6 | `{theme_colors.get('accent6','#70AD47')}` | ■ |
| Hyperlink | `{theme_colors.get('hlink','#0563C1')}` | ■ |

### Background Colors (used across slides)
| Purpose | Hex |
|---|---|
| Primary background | `#0F1117` (near-black navy) |
| Secondary background | `#1A1C25` (dark blue-gray) |
| Section divider | `#2E3140` (dark slate) |
| Teal accent bg | `#4F98A3` / `#6CC0CC` |
| Amber accent bg | `#E8A033` |
| Green accent bg | `#6DAA45` |
| White (contrast) | `#FFFFFF` |

---

## Slide Layouts
| Layout | Placeholders |
|---|---|
{''.join(f"| {l['name']} | {', '.join(p['name'] for p in l['placeholders']) or '—'} |" + chr(10) for l in layouts)}

---

## Radiology-Specific Usage Notes

### Background strategy
- Use **dark backgrounds** (`#0F1117`–`#2E3140`) as the default — they provide maximum contrast for DICOM MRI/CT images and reduce eye strain in darkened lecture rooms.
- Reserve light backgrounds (`#FFFFFF`, `#E0DFD8`) for text-heavy summary slides only.

### Image slides
- Place medical images on a **pure black** (`#000000`) or near-black background.
- Do **not** add drop shadows or borders around DICOM images — they distort perceived window/level.
- Leave ≥ 10% margin around images to avoid edge cropping during projection.

### Color usage with DICOM grayscale
- Avoid pure red (`#FF0000`) and yellow (`#FFFF00`) as annotation colors — use **teal** (`#6CC0CC`) or **amber** (`#E8A033`) instead for better visibility without implying pathology urgency.
- Arrows and callout lines: use `#6CC0CC` (teal) at 2–3 pt thickness.

### Font & contrast
- Minimum font sizes for lecture hall projection:
  - **Title**: 32 pt bold
  - **Body**: 20 pt
  - **Captions / footnotes**: 14 pt
- Text color on dark backgrounds: **white** (`#FFFFFF`) or light gray (`#E0DFD8`).
- Avoid colored text on colored backgrounds — use white text on all accent backgrounds.

### Slide count guidance
- This presentation used **{len(prs.slides)} slides** — aim for ≤ 1 slide per minute of talk time.
"""

with open("radiology_style_guide.md", "w") as f:
    f.write(md)
print("✓ radiology_style_guide.md")

# ── 8. Create template PPTX ───────────────────────────────────────────────────
tpl = Presentation()
tpl.slide_width  = prs.slide_width
tpl.slide_height = prs.slide_height

def hex_to_rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def add_slide(tpl, layout_idx, title_text, body_text=None, bg_hex="#0F1117"):
    layout = tpl.slide_layouts[min(layout_idx, len(tpl.slide_layouts)-1)]
    slide  = tpl.slides.add_slide(layout)
    # Background
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = hex_to_rgb(bg_hex)
    # Title
    if slide.shapes.title:
        tf = slide.shapes.title.text_frame
        tf.text = title_text
        p = tf.paragraphs[0]
        run = p.runs[0] if p.runs else p.add_run()
        run.font.name  = heading_font
        run.font.size  = Pt(36)
        run.font.bold  = True
        run.font.color.rgb = hex_to_rgb("#FFFFFF")
    # Body
    if body_text:
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 1:
                tf = ph.text_frame
                tf.text = body_text
                p = tf.paragraphs[0]
                run = p.runs[0] if p.runs else p.add_run()
                run.font.name  = body_font
                run.font.size  = Pt(24)
                run.font.color.rgb = hex_to_rgb("#E0DFD8")
                break
    return slide

# Slide 1 — Title slide
add_slide(tpl, 0, "Presentation Title", "Subtitle / Author · Institution · Date", bg_hex="#0F1117")
# Slide 2 — Content slide (dark)
add_slide(tpl, 1, "Section Heading", "• Key finding\n• Supporting point\n• Conclusion", bg_hex="#1A1C25")
# Slide 3 — Image slide (near-black)
add_slide(tpl, 5, "Image Slide", bg_hex="#000000")
# Slide 4 — Summary slide (dark slate)
add_slide(tpl, 1, "Summary", "• Point 1\n• Point 2\n• Point 3", bg_hex="#2E3140")
# Slide 5 — Thank you / questions
add_slide(tpl, 0, "Thank You", "Questions?", bg_hex="#0F1117")

tpl.save("radiology_template.pptx")
print("✓ radiology_template.pptx")
print("\nDone! Style extracted from", len(prs.slides), "slides.")
print("Heading font:", heading_font, "/ Body font:", body_font)
print("Primary bg:", primary_bg)
