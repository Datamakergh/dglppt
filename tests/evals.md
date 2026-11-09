# dglppt eval cases

`run_tests.py` automates cases 1–9 and 15. Cases 10–14 are prompts to try with the skill loaded, judged by a person.

| # | Input | Expected outcome |
|---|---|---|
| 1 | Truth `.pptx` with slides | Spec has theme fonts and colors, custom placeholder idx, master logo and bar, slide-level badge, footer shown on content slides but not on the title slide |
| 2 | Truth `.potx` with **no slides** | Extracts without crashing. Borders come from layout and master placeholders. It warns that footer visibility must be confirmed |
| 3 | Target 4:3 → truth 16:9 | Truth slide size; content scaled uniformly and never stretched; asks first when the aspect ratio changes |
| 4 | Target body placeholder `idx=1`, truth body `idx=13` | Placeholder remapped to idx 13 and inherits the truth geometry (no local xfrm) |
| 5 | Truth shows footer and slide number | Every content slide gets slide-level footer and slide-number placeholders with the truth text; the title slide doesn't |
| 6 | Target with local font, size, alignment, a red "12%" and a hyperlink | Local styling gone, red kept and reported, link kept, bold kept |
| 7 | Old footer "Confidential" and old logo | Listed as leftovers; removed only when `approved.json` says so; ids never reused |
| 8 | GIF, MP4 with poster frame, entrance animations, a transition | Bytes, frames, poster, trim and settings unchanged; animations still target the same shapes; placement and size only, aspect kept |
| 9 | Icon row of three sizes; the same screenshot on two slides at different sizes | Identical icon boxes, one top edge, equal gaps, captions on one line and centered; screenshot the same size on both slides |
| 10 | 13 bullets that slightly overflow | Shrink 20 → 18 pt (≥ 85%), reported |
| 11 | 31 bullets | Stops: split / allow smaller / trim; nothing cut |
| 12 | Target sent before the truth file | Doesn't restyle; asks for the truth file |
| 13 | Truth has no table, target has one | Asks how tables should look instead of inventing a style |
| 14 | "Fix the typos while you're at it" | Lists the typos and asks; changes only what the user confirms, recorded in `approved.json` → `change` |
| 15 | Truth slide with hand-made design: moved placeholder, fill, bottom anchor, right alignment, 32 pt red text, a background, a decorative line and a text callout | Targets on that layout get position, fill, anchor, alignment, size, color, background and the line; the callout is asked about and copied only when approved; other layouts untouched; best-matching truth slide donates when several share a layout; overflow from the copied design asks first and is never shrunk |

Mutation tests (must FAIL their check): edit one word, swap two bullets, change a chart value, change speaker notes (content); change a video trim, delete an animation effect, flatten a GIF to one frame, stretch a video (media); move the badge 1 pt, a local title size, a picture in the border, an icon out of its row (geometry).
