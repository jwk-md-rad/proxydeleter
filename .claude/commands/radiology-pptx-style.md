# /radiology-pptx-style

Extract the visual style from an existing PowerPoint presentation and produce:
1. A documented **style guide** (JSON + Markdown summary)
2. A reusable **PPTX template** with those styles applied

This skill is tailored for radiology-related presentations.

---

## Steps

### 1. Locate the source presentation

If the user provided a file path as an argument to this command, use that.
Otherwise, ask:

> "Please provide the path to your existing .pptx presentation."

### 2. Install dependencies if needed

```bash
pip install python-pptx 2>/dev/null || pip3 install python-pptx
```

### 3. Extract the style

Write and run a Python script (`extract_style.py`) that uses `python-pptx` to extract:

- **Slide dimensions** (width × height in cm and EMU)
- **Theme colors** (accent 1–6, background 1–2, text 1–2, hyperlink, followed hyperlink)
- **Explicit RGB colors** used on shapes (fill, line, text) across all slides
- **Fonts**: name, size, bold/italic for each placeholder type (title, body, caption)
- **Background**: solid fill color or image path per slide layout
- **Slide layouts**: names and which placeholders they contain
- **Slide master**: background color, font defaults

Save the result as `radiology_style_guide.json`.

### 4. Generate the style guide Markdown

From the extracted data, produce `radiology_style_guide.md` with sections:

```
# Radiology Presentation Style Guide

## Slide Dimensions
## Color Palette
## Typography
## Slide Layouts
## Usage Notes (radiology-specific tips)
```

In the Usage Notes, include best practices for radiology talks:
- High-contrast backgrounds for medical images (prefer dark backgrounds)
- Font size minimums for lecture halls (title ≥ 36pt, body ≥ 24pt)
- Recommended color pairs that don't clash with DICOM grayscale images

### 5. Create the template PPTX

Write and run a second Python script (`create_template.py`) that:

1. Creates a new `Presentation()` with the extracted slide dimensions
2. Applies the theme colors to the slide master (use `prs.slide_master`)
3. Adds slide layouts matching the originals (at minimum: Title Slide, Title + Content, Blank)
4. Sets default fonts and text colors on each layout's placeholders
5. Saves as `radiology_template.pptx`

### 6. Report to the user

After all files are saved, summarize:

- Path to `radiology_style_guide.json`
- Path to `radiology_style_guide.md`
- Path to `radiology_template.pptx`
- Key style facts found: primary colors, font family, slide size

---

## Arguments

`$ARGUMENTS` — optional path to the source .pptx file.

Example: `/radiology-pptx-style ./my_talk_2025.pptx`

---

## Notes

- If `python-pptx` cannot fully replicate the theme (e.g. embedded images, complex gradients), note the limitation and document those elements manually in the style guide instead.
- Do not modify the original source file.
- Generated files are saved in the current working directory unless the user specifies otherwise.
