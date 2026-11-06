---
name: dglppt
description: Restyle any PowerPoint deck to match a "truth" .pptx/.potx exactly — layouts, borders, spacing, fonts, colors, header, footer, slide numbers and logos — without changing any content, including GIFs, video, audio and animations. Also makes visuals that belong together (icon rows, repeated screenshots) the same size and alignment. Use when the user starts a truth-file session, uploads a truth/master/reference/brand deck or template, or asks to "apply our template", "rebrand this deck", "make it look like this deck" or "match the truth file".
---

# dglppt — Truth File Restyle

The user supplies a **truth file**: a finished, correctly designed .pptx (or a .potx template). Every deck they send afterwards is a **target**. Each target is restyled to match the truth file. Only design changes. The words, numbers, images, GIFs, videos, audio, charts, tables, animations, slide order and notes stay as they are.

All the mechanical work is done by the scripts in `scripts/` (python-pptx, lxml, Pillow). In this file, `$S` means `<this skill's directory>/scripts`. Each script prints what it did and writes a `.meta.json` next to its output that the next script reads. Run them in the order below, or run `pipeline.py` for all of them at once.

## The two unbreakable rules

1. **Content is locked.** Never edit, reword, delete, add, reorder or "fix" any content without the user's explicit permission in chat. Content means:
   - all text (including typos), numbers, table cells, chart data, speaker notes, hyperlinks, alt text, section names, hidden-slide flags and the number and order of slides;
   - every image, **GIF (all its frames and timings), video and audio file**, its poster frame, trim, bookmarks and playback settings;
   - **animations and transitions**, and the shapes they point at.

   Media and animated content may change in **placement only**: position and size, with the aspect ratio kept. They also take the truth file's frame styling (border, shadow, corners) when the truth defines one. If restyling seems to need a content change (splitting a slide, dropping an old footer line, removing a duplicate logo), **stop and ask**. One approval covers one change only. Record every approval in `approved.json` (format below).
2. **The truth file wins on every design decision.** Positions, sizes, margins, gaps, fonts, sizes, colors, line spacing, bullets, insets, fills, borders, header, footer, slide number and logo all come from the truth file. When the truth file doesn't define something the target needs (no table in the truth, for example), ask rather than invent.

## Session flow

```
0. Preflight                → tools and fonts available?
1. Get the truth file       → nothing else happens before this
2. Extract the truth spec   → truth_spec.json + summary for the user to confirm
3. Take in a target         → fingerprint, layout map, leftovers, visual groups (one list of questions)
4. Restyle                  → transplant → remap → strip → place → group → fit
5. Verify                   → content, media, geometry, pixel and visual checks
6. Deliver                  → restyled deck + before/after PDF + report
```
Steps 3–6 repeat for every new target. The truth spec stays in force until the user uploads a new truth file.

**Auto mode.** If the user says something like "just do it" or "no questions unless needed", run `pipeline.py` (Step 4 below) with the defaults. It still stops (exit code 2) for anything on the permission list. All other decisions go into the report.

---

## Step 0 — Preflight

`python $S/preflight.py [truth.pptx]` checks python-pptx/lxml/Pillow (required), LibreOffice + pdftoppm (renders), fc-match (font metrics) and the pptx skill's `validate.py`, and lists truth fonts that aren't installed.
- **No LibreOffice:** skip renders and the pixel check, say so, and rely on the content, geometry and media checks.
- **No pptx skill:** skip `validate.py`. Opening the output once in LibreOffice (render) is then the only file-integrity check. Say so.
- **Truth font missing:** renders substitute it, so text fit in the preview is unreliable. Ask whether the user can provide the font files. The geometry checks are unaffected.

## Step 1 — Get the truth file

Ask for it first and wait:

> "Please upload your truth file, the finished .pptx (or .potx template) that every other deck should match. I'll read its design rules first and then restyle anything you send after it."

- If a target arrives first, don't restyle it. Ask for the truth file.
- If several files arrive at once, ask which one is the truth file.
- `.pptx`, `.potx` and `.ppsx` all work. Convert a legacy `.ppt` to `.pptx` first (pptx skill: `soffice.py --headless --convert-to pptx`).
- **Reuse:** if a `truth_spec.json` from an earlier session exists and its `source.sha1` matches the uploaded truth file, offer to reuse it. Otherwise extract again.

## Step 2 — Extract the truth spec

Work on copies, never on uploads.

```
python $S/extract_truth.py truth.pptx truth_spec.json --summary truth_summary.md
```
The output records everything in EMU (914,400 per inch, 12,700 per pt):
- **Canvas and theme:** slide size, the 12 theme colors, heading and body fonts, and the master text styles (title/body/other, per level).
- **Layouts:** every layout, keyed `m<master>/<layout name>`, with its placeholders (type, idx, box) and the *resolved* text style of each placeholder per level (inheritance included).
- **Borders and safe area:** taken from the layout placeholders first, so a `.potx` without slides works. When truth slides exist, they cross-check the values, which are clustered to 0.01 in. Also recorded: the content area and the vertical and horizontal gaps.
- **Fixed elements:** master and layout art, the footer, slide-number and date placeholders, and anything repeated at the same spot on truth slides (logos, badges, header text, bars). Also which footer elements the truth slides actually *show* per layout, and their text.
- **Element styles:** table style (id, flags, cell fills and margins, text size), chart style (palette order, font sizes, legend, gridlines), picture frame style, box style, and the free text-box roles (body text box, footnote).

**Check it visually.** Run `python $S/render.py truth.pptx renders/truth` and look at the slides next to the summary:
- every logo, header and footer you can see appears under "Fixed elements";
- the borders match the empty space you can see;
- every text role (title, subtitle, body levels, captions, table text, chart labels, footer) has a style.

Fix the spec JSON by hand if something was missed, and note what you changed.

**Confirm with the user.** Send the summary: slide size, fonts, main colors, layouts with borders in inches and cm, and where the header, footer, slide number and logo sit. Include anything under "Needs confirmation" (for example "the truth has no table: how should tables look?"). **Only the confirmed spec is used.** Then ask for the first target.

---

## Step 3 — Take in a target (ask everything in ONE message)

1. **Copy the upload** and work on the copy. Keep the original for the before/after.
2. **Fingerprint the content:** `python $S/fingerprint.py target.pptx before_fingerprint.json`. It captures text in order (including fields and line breaks), links, alt text, tables, chart data, SmartArt text, image/GIF/video/audio hashes, GIF frames, media settings, animations, transitions, notes, sections and hidden slides.
3. **Render the original:** `python $S/render.py target.pptx renders/before`.
4. **Map slides to truth layouts:** `python $S/map_layouts.py truth_spec.json target.pptx layout_map.json`. Show the table and ask only about **low/none** rows. For a slide no layout fits, offer two choices: the nearest layout, or keep the slide's arrangement and apply only the truth margins, fonts, colors and fixed elements (map it to the truth "Title Only" or "Blank" layout).
5. **Leftovers from the old design:** the target's own footer text, logo, page number or decorative bars. They will be replaced by the truth's fixed elements, but **removing them needs permission** whenever they hold text or an image ("Confidential" is content). List them by slide and shape id, ask once, and record the yes answers in `approved.json`.
6. **Different slide size:** say so. The truth size is used, and content is scaled uniformly into the new canvas (never stretched). Ask first if the aspect ratio changes (4:3 → 16:9).
7. **Visual groups** (visuals that should look identical). Run this after placement in Step 4, but ask about the groups here. Proposals come from identical images, perceptually similar images, rows or columns of similar visuals, and the same visual repeated across slides. Look at each group's contact sheet, drop wrong groups (an icon and a GIF that merely sit above each other), and **add** groups that depict the same idea without matching pixels (three different feature icons). Ask the user about medium-confidence groups.
8. **Free text-box roles (optional):** text boxes get the truth "textbox" style by default. If one is really a title, caption or footnote, record that in `roles.json`: `{"<slide>": {"<shape id>": "caption"}}`.

`approved.json`, which every script reads:
```json
{"remove": [{"slide": 2, "shape_id": 4, "why": "old footer 'Confidential' — user said yes"}],
 "change": [{"slide": 5, "shape_id": 3, "field": "text", "why": "user asked to fix the typo"}],
 "add": [], "media_scale": [{"slide": 4, "shape_id": 6, "why": "user allowed 80%"}]}
```

## Step 4 — Restyle (inside the target's own file)

The restyle happens inside the target's package, so its slides, media, charts and notes never leave it. Every step edits shapes **in place**, so shape ids, and the animations that point at them, survive.

```
python $S/transplant_design.py truth.pptx target.pptx layout_map.json s1.pptx --spec truth_spec.json
python $S/remap_placeholders.py s1.pptx truth.pptx truth_spec.json s2.pptx --approved approved.json
python $S/strip_overrides.py   s2.pptx truth_spec.json s3.pptx [--roles roles.json]
python $S/place.py             s3.pptx truth_spec.json s4.pptx [--plan plan.json]
python $S/visual_groups.py detect s4.pptx groups.json --sheets group_sheets   # review, edit, confirm
python $S/visual_groups.py apply  s4.pptx truth_spec.json groups.json s5.pptx
python $S/fit_text.py          s5.pptx "<name> - restyled.pptx"
```
Or run it all at once: `python $S/pipeline.py truth.pptx target.pptx --work WORK --approved approved.json [--groups groups.json] [--map layout_map.json] [--roles roles.json]`. Exit code 0 = done, 2 = questions for the user (listed in `report.md`), 1 = a check failed.

What each step guarantees:
- **4a. transplant_design:**
  - first freezes every placeholder's current geometry;
  - then swaps the target's theme, masters and layouts for the truth's, renaming parts so nothing collides and fixing relationships, content types, master and layout ids, slide size, default text style and table styles;
  - garbage-collects only parts no slide uses anymore. Videos, GIFs, charts and workbooks stay.
- **4b. remap_placeholders:**
  - matches each slide placeholder to its truth layout placeholder (same idx, then the same kind) and rewrites its type and idx. A body with `idx=1` becomes the truth's `idx=13` when needed;
  - clears its local geometry, so it inherits the truth position exactly;
  - **adds slide-level footer, slide-number and date placeholders** wherever the truth slides show them (a layout placeholder alone never appears on a slide), with the truth footer text;
  - copies truth slide-level fixed elements (a badge on every truth slide) at their exact box;
  - deletes only approved leftovers and never reuses a shape id.
- **4c. place:** moves all free content of a slide (pictures, GIFs, video, charts, tables, text boxes, groups) as **one block**, using one uniform scale plus one move into the layout's content area, aligned to the truth's left content line. Arrangement and aspect ratios survive, and nothing is cropped. It reports anything that touches a fixed element. To place a slide by hand, use `plan.json`: `{"<slide>": {"area": [x,y,w,h], "boxes": {"<id>": [x,y,w,h]}}}`.
- **4c-bis. visual_groups apply:** members of a confirmed group get the **same box**. Each visual is fitted inside the box and centered, never stretched or cropped beyond the original.
  - Rows share one top edge and use equal gaps, spread over the row's span. Add `"gap": "truth"` to a group to use the truth gap.
  - Captions keep one shared line and stay centered under their own visual.
  - A visual repeated across slides gets the first occurrence's box, and its spot too unless that would overlap other content there (then the size matches, and it's reported).
- **4d. strip_overrides:**
  - **Removed:** local fonts, sizes, colors, spacing, alignment, indents, bullet styling, shape list styles and stale autofit scales. Free text boxes get the truth role style written into their own list style, and keep their alignment.
  - **Kept:** bold, italic, underline, strike, super/subscript, highlight, links, list type (bullet / number / none) and paragraph level. **Semantic color** (a red "12%" inside black text) is kept and reported. Drop it only with `--drop-semantic-color` and the user's yes.
  - **Tables** get the truth table style, flags, cell margins and text size. **Charts** get the truth palette order, font sizes, legend and gridlines. Data, categories and workbooks are never touched.
  - **Pictures and media frames** get the truth frame style. The image and the crop never change.
- **4e. fit_text:** measures text with real font metrics. If text doesn't fit, it shrinks 1 pt at a time down to **85% of the truth size** (never below 10 pt body or 8 pt small text), then tightens line spacing one step (100% → 90%). If it still doesn't fit, it **stops and asks**: split the slide (keeps every word, adds a slide), allow a smaller size this once (`--allow SLIDE:SHAPE`), or the user trims the text. Every shrink is listed in the report.
- **Media floor:** a GIF, video or audio frame scaled below 85% of its size needs approval (`approved.json` → `media_scale`).
- **4f. Validate the file:** `python <pptx skill>/scripts/office/validate.py "<name> - restyled.pptx" --original target.pptx`. Fix everything it reports. `pipeline.py` runs this automatically when the pptx skill is present.

## Step 5 — Verify (all must pass before delivery)

```
python $S/fingerprint.py restyled.pptx after_fingerprint.json
python $S/compare_fingerprint.py before_fingerprint.json after_fingerprint.json --approved approved.json
python $S/check_media.py target.pptx restyled.pptx --approved approved.json
python $S/check_geometry.py restyled.pptx truth.pptx truth_spec.json
python $S/render.py restyled.pptx renders/after
python $S/band_diff.py truth_spec.json restyled.pptx renders/truth renders/after --out band_diffs
```
- **5a. Content lock (`compare_fingerprint`):** any difference is a failure, except approved changes and the truth furniture that was added. It catches a changed word, swapped bullets, a changed chart value or edited notes.
- **5a-bis. Media and animation integrity (`check_media`):** same GIF, video and audio bytes, poster frames, trim and playback settings, GIF frames; same animation timeline and transitions; every animation target still exists and is the same shape; aspect ratios kept; no unapproved shrink below 85%.
- **5b. Geometry and style (`check_geometry`, exact, in EMU):**
  - layout XML identical to the truth;
  - fixed elements at 0 EMU difference, and every footer, slide number and date the truth shows is present and inherited;
  - matched placeholders inherit the truth geometry;
  - no free content in the borders or over a fixed element;
  - every run's *effective* font, size and color, and every paragraph's spacing and alignment, equal what the truth gives it, apart from emphasis, kept semantic colors and listed shrinks;
  - visual groups have identical slots, a shared edge and equal gaps.
- **5c. Visual:**
  - `band_diff` compares the fixed-element regions pixel by pixel against a truth slide of the same layout, with slide numbers and dates masked. A failure usually means text running into the footer, or furniture that moved.
  - Then look yourself, with fresh eyes. For long decks, start with the contact sheets (`before_after.py … --sheets`) and open full renders only for flagged slides. Look for overflow, clipping, overlap, words broken across lines, and the same empty borders, alignment lines and rhythm of gaps as the truth.
  - Renders show only a video's poster frame and a GIF's first frame, so the visual check for media covers placement only. `check_media` covers the rest.

If a check fails: fix, then run **all** checks again.

## Step 6 — Deliver

1. **The restyled deck:** `<original name> - restyled.pptx`.
2. **Before / after PDF:** `python $S/before_after.py renders/before renders/after "<name> - before-after.pdf" --labels "<restyled>.meta.json"`. One page per slide: original left, restyled right, with the slide number and truth layout.
3. **A short report** (`pipeline.py` writes `report.md`):
   - checks passed;
   - slide → truth layout;
   - every shrink (slide, shape, from → to);
   - semantic colors kept;
   - what was removed with approval;
   - visual groups made identical;
   - media placements (content verified unchanged);
   - anything that couldn't match the truth exactly, and why;
   - missing fonts;
   - anything waiting on the user.

Then ask for the next target. The truth spec stays loaded, and `truth_spec.json` is saved for later sessions.

---

## Running inside PowerPoint (Claude in PowerPoint / Microsoft 365 add-in)

The same rules and order apply. Only the tools change. **Prefer the file-based route above for a full restyle**: only it can run the scripts and their checks. Use the add-in for touch-ups, decisions on flagged slides, and a final review with the real fonts. When the user is in the add-in and asks for a full restyle, say this once and offer the file-based route.
- **Truth:** read the slide size, layouts, placeholder positions, fonts, colors, fixed elements, borders and gaps through the add-in's actions. Build the same spec, then summarize it and confirm as in Step 2.
- **Target:** the open deck. Save a duplicate first (or ask the user to), so the original survives for the before/after.
- **Restyle:** set each shape's position, size, font, size, color, spacing and fill to the spec values, and apply the truth layouts and theme where the add-in supports it. **Move and resize** GIFs, videos and animated shapes only; never replace or re-insert them, because that would drop their animations. If the add-in can't do something (for example importing a slide master), say so and offer the file-based route for that part.
- **Verify:** re-read every shape's properties and compare them with the spec (5a, 5b). Export or screenshot slides from both copies for the before/after where possible; otherwise list the changes per slide.
- **The content lock is the same.** Never type into, delete or add slides or text boxes without permission.

## Things that always need the user's permission

- Changing, deleting or adding any text, number, image, GIF, video, audio, chart data, table cell, note, link or alt text
- Changing, removing or re-creating animations or transitions, or replacing a media file or its poster frame
- Adding or deleting a slide, including splitting an overfull slide
- Reordering slides
- Removing old header, footer or logo items that contain text or images
- Shrinking text below the 85% floor or the minimum sizes, or a media frame below 85% of its size
- Removing semantic text color, or a picture background
- Using a layout the truth file doesn't have, or inventing a style the truth file doesn't define
- Scaling content to a different aspect ratio

## Known limits (say so when they apply)

- **Charts pasted as pictures** can't be restyled; they are placed like images.
- **Diagrams built from plain shapes** (arrows, boxes) keep their own fills unless the user asks for `--restyle-boxes` (truth box style).
- **SmartArt** text is locked and checked, but SmartArt colors follow the theme only.
- **Embedded fonts** in the target aren't carried over. The truth fonts apply.
- **Font substitution:** renders use substitutes for fonts that aren't installed, so text fit for those fonts is an estimate. The geometry checks are exact regardless.
- **Fit estimates** assume PowerPoint-like line breaking. On a tight box, look at the render before you trust it.

## Testing the skill

`python tests/run_tests.py` builds fixture decks and runs the full flow. Covered:
- a 16:9 truth with custom placeholder idx, a logo, a footer and a badge;
- a 4:3 target with overrides, semantic red, a hyperlink, notes, leftovers, an icon row, a GIF, animations, a transition, an MP4 with a poster frame, a table, a chart, a group, a repeated screenshot and overflow;
- mutation tests that must each fail the right check.

Cases are described in `tests/evals.md`.
