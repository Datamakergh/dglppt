# dglppt

A Claude skill that restyles PowerPoint decks to match a **truth file**: a finished, correctly designed `.pptx` (or a `.potx` template).

Upload the truth file first. Every deck you send after it is restyled to the truth file's layouts, borders, spacing, fonts, colors, header, footer, slide numbers and logo, **without changing any content**. That includes GIFs, videos, audio, animations and transitions, which are only repositioned. Visuals that belong together (an icon row, the same screenshot on several slides) get identical size, alignment and placement.

Each run ends with the restyled deck, a before/after PDF and a report.

## Install

Add this folder (with `SKILL.md`, `scripts/` and `tests/`) as a skill in Claude, then start it with `/dglppt`.

### Requirements
| Need | For | Without it |
|---|---|---|
| Python 3.9+, `python-pptx`, `lxml`, `Pillow` | everything | required: `pip install python-pptx lxml Pillow` |
| LibreOffice (`soffice`) + `pdftoppm` (poppler) | renders, pixel check, before/after PDF | content, media and geometry checks still run |
| `fontconfig` (`fc-match`) | real font metrics for text fit | rough estimates |
| the pptx skill (`scripts/office/validate.py`) | OOXML schema validation | LibreOffice round-trip only |
| `ffmpeg` | building the test fixtures only | — |

Run `python scripts/preflight.py truth.pptx` to see what's available and which truth fonts are missing.

### Where to run it
**Run the restyle in Claude with code execution on** (Claude Code, or claude.ai with code execution enabled). Every guarantee comes from the scripts: the master and layout transplant, the placeholder remap, the content lock, the 0-EMU geometry check, the media and animation check, and the renders and pixel check. Without code execution none of these can run.

**Then open the result in PowerPoint** for:
- a final look with the real fonts and renderer (text fit, video playback, animations), since LibreOffice substitutes missing fonts;
- small fixes on the slides the report flags.

The Claude in PowerPoint add-in can follow the same rules, but it can't import a slide master or verify media bytes, so it checks its own work less strictly. For the best previews, upload the truth deck's font files.

## Example session

```
You:    /dglppt
Claude: Please upload your truth file…
You:    [brand-2026.potx]
Claude: Truth spec: 13.33 × 7.5 in, Caladea / Carlito, navy #1E2761…
        Borders 0.6 in left/right, footer "ACME Corp — Internal" + slide number on content slides,
        logo top right. Correct?
You:    yes
You:    [q3-review.pptx]
Claude: 12 slides mapped (2 need you: …). Old footer "Confidential" on 9 slides — remove? 
        Icon row on slide 4 and the screenshot on slides 6 & 9 — make them identical?
You:    yes to all
Claude: Done. q3-review - restyled.pptx, before/after PDF, report:
        content/media/geometry checks passed, slide 7 text 20→18 pt,
        slide 11 doesn't fit even at 17 pt — split it, allow smaller, or trim?
```

Auto mode, one command:

```bash
python scripts/pipeline.py truth.pptx target.pptx --work out/ --approved approved.json
```

## What's inside

| Script | Step |
|---|---|
| `preflight.py` | check the tools and fonts |
| `extract_truth.py` | truth spec: theme, layouts, borders, gaps, fixed elements, table, chart and picture styles |
| `fingerprint.py`, `compare_fingerprint.py` | content lock: text, links, tables, charts, SmartArt, images, GIF frames, video, audio, animations, notes |
| `map_layouts.py` | target slide → truth layout, with confidence |
| `transplant_design.py` | move the truth theme, masters and layouts into the target |
| `remap_placeholders.py` | placeholder idx remap, footer / slide number / date, slide-level fixed elements, approved removals |
| `strip_overrides.py` | remove local styling, keep emphasis and semantic color, restyle tables, charts and pictures |
| `place.py` | fit free content into the truth safe area as one block |
| `visual_groups.py` | detect and normalize groups of visuals that belong together |
| `fit_text.py` | gentle shrink (85% floor), or stop and ask |
| `check_geometry.py`, `check_media.py`, `band_diff.py` | exact geometry and style; media and animation integrity; pixel check of fixed elements |
| `render.py`, `before_after.py` | renders, before/after PDF, contact sheets |
| `pipeline.py` | everything in order, plus `report.md` |

## Tests

```bash
python tests/run_tests.py
```

The suite builds fixture decks, runs the full flow and runs mutation tests. Each deliberate content, media or geometry change must fail its check. See `tests/evals.md`.

## Limits

Charts pasted as pictures can't be restyled. SmartArt colors follow the theme only. Fonts that aren't installed are substituted in renders, so text fit for them is an estimate. Details are in `SKILL.md` → *Known limits*.
