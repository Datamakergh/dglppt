---
name: dglppt
description: Restyle any PowerPoint deck so it matches a "truth" .pptx pixel for pixel — layouts, margins, spacing, fonts, colors, headers, footers and logos — without touching the content. Use when the user starts a truth-file session, uploads a truth deck, or asks to make a deck match their truth/master/reference file.
---

# dglppt — Truth File Restyle

The user supplies a **truth file**: a finished, correctly designed .pptx. Every deck they send afterwards is a **target**. Each target is restyled to match the truth file exactly. Only design changes. The words, numbers, images, charts, tables, order and notes stay as they are.

## The two unbreakable rules

1. **Content is locked.** Never edit, reword, delete, add, reorder or "fix" any content without the user's explicit permission in chat. Content means all text (including typos), numbers, table cells, chart data, images, slide order, speaker notes, hyperlinks and the number of slides. If restyling seems to need a content change (splitting a slide, dropping an old footer line, removing a duplicate logo), **stop and ask**. One approval covers one change only.
2. **The truth file wins on every design decision.** Positions, sizes, margins, gaps, fonts, sizes, colors, line spacing, bullets, insets, fills, borders, shadows, header, footer and logo all come from the truth file. Nothing is left as a "close enough" guess. When the truth file doesn't define something the target needs, ask rather than invent.

## Session flow

```
1. Get the truth file        → nothing else happens before this
2. Extract the truth spec    → truth_spec.json + summary for the user
3. Take in a target          → content fingerprint + slide-to-layout map
4. Restyle the target        → in place, content untouched
5. Verify                    → content check, geometry check, visual check
6. Deliver                   → restyled deck + before/after + report
```
Steps 3–6 repeat for every new target. The truth spec stays in force for the whole session until the user uploads a new truth file.

---

## Step 1 — Get the truth file

When the skill starts, ask for the truth file first and wait for it:

> "Please upload your truth file, the finished .pptx that every other deck should match. I'll read its design rules first and then restyle anything you send after it."

- If the user sends a target before a truth file, don't restyle it. Ask for the truth file.
- If more than one file arrives at once, ask which one is the truth file.
- A `.potx` truth file is fine. A legacy `.ppt` must be converted to `.pptx` first (see the pptx skill).

## Step 2 — Extract the truth spec

Work on a copy, never on the upload. Run the extractor below on the truth file to produce `truth_spec.json`. All measurements are in **EMU** (914,400 EMU = 1 inch, 12,700 EMU = 1 pt) so nothing is lost to rounding.

The spec has to capture:

| Area | What is recorded |
|---|---|
| Canvas | Slide width and height |
| Theme | All 12 theme colors, heading and body fonts, the raw theme XML |
| Masters and layouts | Every layout's name, placeholders (type, idx, x, y, w, h), text styles per level, background, and raw XML |
| Fixed elements | Header, footer, slide number, date, logo(s): exact position, size, image hash, crop, and whether they live on the master, layout or slide |
| Borders / safe area | Left, right, top and bottom space left empty on each layout, plus the content safe area inside the header and footer zones |
| Gaps | Spacing between title and body, between stacked blocks, between columns, as most-used values per layout |
| Text styles | For each level and role: font, size, bold/italic, color, alignment, line spacing, space before/after, bullet character/color/size, indent, hanging indent, text box insets, autofit and wrap |
| Shapes | Fill, outline width/color/dash, corner radius, shadow, effects (raw `spPr` XML) |
| Tables | Table style, header row/banding, cell fill, borders, cell margins, font sizes |
| Charts | Palette order, fonts and sizes of axis/data labels/legend, gridlines, label positions |
| Images | Border, shadow, corner radius, crop style |

### Extractor

```python
# extract_truth.py  —  python extract_truth.py truth.pptx truth_spec.json
import sys, json, hashlib
from collections import Counter
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.dml import MSO_COLOR_TYPE
from lxml import etree

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

def xml(el):
    return etree.tostring(el, encoding='unicode') if el is not None else None

def color_of(font):
    try:
        if font.color is None or font.color.type is None: return None
        if font.color.type == MSO_COLOR_TYPE.RGB: return str(font.color.rgb)
        if font.color.type == MSO_COLOR_TYPE.SCHEME: return str(font.color.theme_color)
    except Exception: return None

def text_spec(tf):
    paras = []
    for p in tf.paragraphs:
        runs = [dict(font=r.font.name, size_emu=r.font.size, bold=r.font.bold,
                     italic=r.font.italic, color=color_of(r.font)) for r in p.runs]
        paras.append(dict(level=p.level, align=str(p.alignment),
                          line_spacing=str(p.line_spacing), space_before=p.space_before,
                          space_after=p.space_after, pPr=xml(p._p.find(A+'pPr')), runs=runs))
    body = tf._txBody.find(A+'bodyPr')
    return dict(insets=[tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom],
                wrap=tf.word_wrap, autosize=str(tf.auto_size), bodyPr=xml(body),
                lstStyle=xml(tf._txBody.find(A+'lstStyle')), paragraphs=paras)

def shape_spec(sh):
    d = dict(name=sh.name, kind=str(sh.shape_type), x=sh.left, y=sh.top, w=sh.width,
             h=sh.height, rot=getattr(sh, 'rotation', 0),
             spPr=xml(sh._element.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}spPr')))
    if sh.is_placeholder:
        d['ph_type'] = str(sh.placeholder_format.type); d['ph_idx'] = sh.placeholder_format.idx
    if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
        d['image_sha1'] = hashlib.sha1(sh.image.blob).hexdigest()
        d['crop'] = [sh.crop_left, sh.crop_top, sh.crop_right, sh.crop_bottom]
    if sh.has_text_frame: d['text'] = text_spec(sh.text_frame)
    if getattr(sh, 'has_table', False) and sh.has_table:
        d['table_xml'] = xml(sh._element.find('.//'+A+'tbl'))
    if getattr(sh, 'has_chart', False) and sh.has_chart:
        d['chart_xml'] = etree.tostring(sh.chart._chartSpace, encoding='unicode')
    return d

FIXED_PH = {PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE}

def main(src, out):
    prs = Presentation(src)
    W, H = prs.slide_width, prs.slide_height
    spec = dict(canvas=dict(w=W, h=H), masters=[], slides=[])

    for m in prs.slide_masters:
        mspec = dict(name=m.name, shapes=[shape_spec(s) for s in m.shapes], layouts=[])
        for lay in m.slide_layouts:
            mspec['layouts'].append(dict(name=lay.name,
                                         shapes=[shape_spec(s) for s in lay.shapes],
                                         xml=xml(lay._element)))
        spec['masters'].append(mspec)

    pic_seen = Counter()
    for i, s in enumerate(prs.slides, 1):
        shapes = [shape_spec(sh) for sh in s.shapes]
        spec['slides'].append(dict(n=i, layout=s.slide_layout.name, shapes=shapes))
        for sh in shapes:
            if 'image_sha1' in sh: pic_seen[(sh['image_sha1'], sh['x'], sh['y'], sh['w'], sh['h'])] += 1

    # Fixed elements: header/footer/number/date placeholders, master/layout art,
    # and any picture repeated at the same spot (logos). Single-slide truth files:
    # pictures in the top or bottom 15% are treated as logos too.
    n_slides = len(spec['slides'])
    for sl in spec['slides']:
        for sh in sl['shapes']:
            in_band = sh['y'] < 0.15*H or sh['y'] + sh['h'] > 0.85*H
            repeated = 'image_sha1' in sh and (pic_seen[(sh['image_sha1'], sh['x'], sh['y'], sh['w'], sh['h'])] > 1
                                               or (n_slides == 1 and in_band))
            sh['fixed'] = repeated or sh.get('ph_type') in {str(p) for p in FIXED_PH}

    # Borders / safe area and gaps per layout, from content shapes only
    # (not fixed, not full-bleed backgrounds).
    per_layout = {}
    for sl in spec['slides']:
        content = [s for s in sl['shapes'] if not s['fixed']
                   and not (s['w'] > 0.9*W and s['h'] > 0.9*H)]
        if not content: continue
        L = per_layout.setdefault(sl['layout'], dict(left=[], top=[], right=[], bottom=[], vgaps=[], hgaps=[]))
        L['left'].append(min(s['x'] for s in content))
        L['top'].append(min(s['y'] for s in content))
        L['right'].append(W - max(s['x']+s['w'] for s in content))
        L['bottom'].append(H - max(s['y']+s['h'] for s in content))
        by_y = sorted(content, key=lambda s: s['y'])
        for a, b in zip(by_y, by_y[1:]):
            g = b['y'] - (a['y']+a['h'])
            if g > 0: L['vgaps'].append(g)
        by_x = sorted(content, key=lambda s: s['x'])
        for a, b in zip(by_x, by_x[1:]):
            g = b['x'] - (a['x']+a['w'])
            if g > 0: L['hgaps'].append(g)
    mode = lambda v: Counter(v).most_common(1)[0][0] if v else None
    spec['borders'] = {k: dict(left=mode(v['left']), top=mode(v['top']), right=mode(v['right']),
                               bottom=mode(v['bottom']), vgap=mode(v['vgaps']), hgap=mode(v['hgaps']))
                       for k, v in per_layout.items()}

    json.dump(spec, open(out, 'w'), indent=1, default=str)

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
```

Also keep the truth file's raw parts for exact copying later:
`ppt/theme/`, `ppt/slideMasters/`, `ppt/slideLayouts/`, their `_rels/`, and `ppt/media/` items referenced by them.

### Check the spec visually

Render the truth file to images (pptx skill → *Converting to Images*) and look at every slide next to `truth_spec.json`. Confirm that:
- every logo, header and footer in the render was detected as `fixed`,
- the border values match the empty space you can see,
- every text role (title, subtitle, body levels 1–3, captions, table text, chart labels, footer) has a recorded style.

If the extractor missed something, correct the spec by hand and note it.

### Confirm with the user

Send a short summary: slide size, fonts, main colors, the layouts found, the border/margin values in inches and cm, and where the header, footer and logo sit. Ask the user to confirm or correct it. **Only the confirmed spec is used.** Then ask for the first target.

---

## Step 3 — Take in a target

1. **Copy the upload.** Work on the copy only. Keep the untouched original for the before/after.
2. **Fingerprint the content** with the script below and save it as `before_fingerprint.json`. This is what the final check compares against.
3. **Render the original** to images (`before-*.jpg`).
4. **Map every target slide to a truth layout.** Match by structure: title only, title + body, two columns, section divider, table, chart, image + text. Show the user a short table:

   | Target slide | Looks like | Truth layout | Confidence |
   |---|---|---|---|

   Ask the user only about low-confidence or unmatched slides. If no truth layout fits a slide, ask: use the nearest layout, or keep the slide's arrangement and only apply truth margins, fonts, colors and fixed elements.
5. **List leftovers from the old design.** These are things in the target that look like its own header, footer, logo, page number or decorative bars. They will be replaced by the truth file's fixed elements. **Removing the old ones needs permission** whenever they contain text (a footer saying "Confidential" is content). Ask once, as a list.
6. **Different slide size.** If the target's slide size differs from the truth's, tell the user. The truth size is used, and content is scaled into the new canvas. Ask first if the aspect ratio differs (for example 4:3 to 16:9).

### Content fingerprint

```python
# fingerprint.py  —  python fingerprint.py deck.pptx out.json
import sys, json, hashlib
from pptx import Presentation

def shape_items(sh, out):
    if sh.shape_type == 6:  # group
        for s in sh.shapes: shape_items(s, out)
        return
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            t = ''.join(r.text for r in p.runs)
            if t.strip(): out['text'].append(t)
            for r in p.runs:
                if r.hyperlink and r.hyperlink.address: out['links'].append(r.hyperlink.address)
    if getattr(sh, 'has_table', False) and sh.has_table:
        out['tables'].append([[c.text for c in row.cells] for row in sh.table.rows])
    if getattr(sh, 'has_chart', False) and sh.has_chart:
        ch = sh.chart
        out['charts'].append(dict(cats=[str(c) for c in ch.plots[0].categories] if ch.plots else [],
                                  series=[[s.name, list(s.values)] for p in ch.plots for s in p.series]))
    if sh.shape_type == 13:  # picture
        out['images'].append(hashlib.sha1(sh.image.blob).hexdigest())

prs = Presentation(sys.argv[1])
slides = []
for s in prs.slides:
    out = dict(text=[], links=[], tables=[], charts=[], images=[])
    for sh in s.shapes: shape_items(sh, out)
    out['notes'] = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ''
    for k in ('text', 'links', 'images'): out[k] = sorted(out[k])
    slides.append(out)
json.dump(dict(slide_count=len(slides), slides=slides), open(sys.argv[2], 'w'), indent=1, default=str)
```

---

## Step 4 — Restyle the target

Restyle **inside the target's own file**, so its text, images, charts and tables never leave it. Use the pptx skill's unpack, edit and zip workflow and its XML rules (parse with `defusedxml`/`lxml`, never `ElementTree`).

### 4a. Bring in the truth design system
1. Replace the target's theme, slide master(s) and slide layouts with the truth file's, including their rels and the media they reference (backgrounds, logos). Update `[Content_Types].xml`, `ppt/presentation.xml` (`sldMasterIdLst`, `sldSz`) and `ppt/_rels/presentation.xml.rels`.
2. Point each target slide's layout relationship at its mapped truth layout.
3. Remove the old masters, layouts and theme only after nothing points at them (pptx skill `clean.py`).

### 4b. Place the fixed elements — exact, never adjusted
- Header, footer, slide number, date and logos come from the truth file **at the exact EMU position, size and crop**. If they live on the truth master or layout, they arrive with 4a. If the truth file has them on its slides, copy those shapes onto each target slide.
- Fixed elements are never shrunk, moved or restyled to make room for content. Content moves around them.
- Remove the target's own old header, footer and logo **only where the user approved it** in Step 3.

### 4c. Place the content inside the borders
- Placeholders: clear the slide-level `<a:xfrm>` so each one inherits the truth layout's geometry. Otherwise set it to the layout's exact x, y, w, h.
- Non-placeholder content (images, charts, tables, free text boxes): fit it inside the truth safe area for that layout. Use the truth borders (left/top/right/bottom) and the truth gaps between blocks and columns. Keep image aspect ratios and never crop more than the original did. Align edges to the same lines the truth file uses.
- Nothing may sit in the border space or overlap a fixed element.

### 4d. Apply the text styles
For every paragraph that is copied or kept:
- **Strip local design overrides.** On `<a:rPr>` remove size, font (`latin`/`ea`/`cs`), color fills and character spacing. On `<a:pPr>` remove alignment, spacing, indents and bullet overrides. Then the truth layout's styles apply through inheritance.
- **Keep meaning-carrying emphasis by default:** bold, italic, underline, strikethrough and hyperlinks on individual words inside a sentence. Keep the paragraph `lvl` (indent level), and keep numbered lists numbered (use the truth file's numbered style if one exists).
- Text boxes get the truth insets, wrap and autofit settings for their role.
- Tables get the truth table style, borders, cell margins and fonts. Charts get the truth palette order, label fonts, sizes and gridlines. Only the styling XML changes. Chart data and table values are never touched.

### 4e. When content doesn't fit — shrink, but gently
Work in this order and stop at the first step that works:
1. Use the full truth safe area for that layout (no extra margin beyond the truth borders).
2. Shrink the text in that shape in steps of 1 pt, **down to 85% of the truth size for that role** (for example 20 pt → 17 pt). Never below 10 pt for body text or 8 pt for footnotes and table text. Shrink line spacing by at most one step (for example 1.0 → 0.9) after the size reaches its floor.
3. Shrink tables and charts proportionally inside the safe area, using the same floors for their text.
4. **Still too much? Stop and ask.** Offer: split the slide in two (keeps every word, adds a slide); allow a smaller size this once; or the user trims the text. Never cut, reword or move content to another slide without a yes.

Every shrink is listed in the report with the slide, the shape and the before/after size.

### 4f. Rebuild and validate the file
Zip the deck and run the pptx skill's validator with the target as the original:
`python <pptx skill>/scripts/office/validate.py restyled.pptx --original target.pptx`
Fix everything it reports before moving on.

---

## Step 5 — Verify

All three checks must pass before anything is delivered. If one fails, fix and run all three again.

### 5a. Content check (must be identical)
Run `fingerprint.py` on the restyled deck and compare it with `before_fingerprint.json`: slide count, every text paragraph, hyperlinks, table cells, chart categories and values, image hashes and speaker notes. **Any difference is a failure**, unless it is a change the user approved, which is listed in the report. Old footer/logo text the user approved for removal is the only expected difference.

### 5b. Geometry check (exact numbers)
For each restyled slide, compare against the truth spec in EMU:
- Fixed elements (header, footer, logo, slide number): **exact match**, 0 EMU difference in x, y, w, h and crop.
- Placeholders: exact match with the truth layout.
- Borders: no content shape closer to any slide edge than the truth border for that layout.
- Gaps: block-to-block and column gaps equal the truth gap values, unless shrinking (4e) required a change, which goes in the report.
- Text: every run's effective font, size, color and paragraph spacing equals the truth role style, apart from listed shrinks.

### 5c. Visual check
Render the restyled deck to images (`after-*.jpg`) at 150 dpi and look at every slide fresh. Compare it with the truth slide of the same layout:
- Header, footer and logo are in the same spot on every slide. Toggling between the truth render and the after render shows no shift.
- No text overflow, clipping, overlap or words broken across lines.
- Same empty borders, same alignment lines, same rhythm of gaps.

LibreOffice substitutes fonts it doesn't have. If a truth font is missing from the render environment, say so, and don't trust text-fit in the preview for that font. The geometry check (5b) still applies.

---

## Step 6 — Deliver

Give the user three things:

1. **The restyled deck** (`<original name> - restyled.pptx`).
2. **Before / after comparison.** Build one PDF (or HTML page) with one row per slide: original on the left, restyled on the right, slide number and mapped truth layout above. Build it from the `before-*.jpg` and `after-*.jpg` renders, for example with Pillow pasting the pairs side by side and saving them as a multi-page PDF.
3. **A short report:**
   - content check: passed (or the approved changes, listed),
   - slide → truth layout used,
   - every shrink applied (slide, shape, from → to),
   - anything that couldn't match the truth exactly and why,
   - anything waiting on the user's decision.

Then ask for the next target. The truth spec stays loaded.

---

## Running inside PowerPoint (Claude in PowerPoint / Microsoft 365 add-in)

When the skill runs inside PowerPoint instead of on uploaded files, the same rules and the same order apply. Only the tools change:

- **Truth file:** the user opens or attaches the truth deck first. Read its slide size, layouts, placeholder positions, fonts, colors, fixed elements, borders and gaps through the add-in's PowerPoint actions, and build the same spec. Summarize it and get confirmation exactly as in Step 2.
- **Target:** the open deck is the target. Save a duplicate first (or ask the user to) so the original survives for the before/after.
- **Restyle** with the add-in's actions to set each shape's position, size, font, size, color, spacing and fill to the spec values, and to apply the truth layouts/theme where the add-in supports it. If the add-in can't do something (for example importing a slide master), say so and offer the file-based route for that part.
- **Verify:** re-read every shape's properties after changes and compare them with the spec (5a and 5b). For the before/after, export or screenshot slides from both copies where the add-in allows it. Otherwise list the changes per slide.
- **The content lock is the same.** Never type into, delete or add slides or text boxes without permission.

---

## Things that always need the user's permission

- Changing, deleting or adding any text, number, image, chart data, table cell, note or link
- Adding or deleting a slide, including splitting an overfull slide
- Reordering slides
- Removing old header/footer/logo items that contain text
- Shrinking below the 85% floor or the minimum sizes
- Using a layout the truth file doesn't have, or inventing a style the truth file doesn't define
- Scaling content to a different aspect ratio
