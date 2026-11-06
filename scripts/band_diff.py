#!/usr/bin/env python3
"""Step 5c (automated part) — pixel-compare the fixed elements against the truth render.

    python band_diff.py truth_spec.json restyled.pptx TRUTH_PREFIX AFTER_PREFIX [--threshold 8] [--out DIR]

TRUTH_PREFIX / AFTER_PREFIX are render.py prefixes (TRUTH_PREFIX-1.jpg …).
For each restyled slide, finds a truth slide that uses the same layout and
compares only the regions of the fixed elements (accent bars, logos, header,
footer, badges), so the slide's own content never counts. Slide numbers and
dates are masked (they differ per slide by design). Score = mean absolute
grayscale difference (0–255) per region; above --threshold = FAIL, and a
side-by-side crop is written to DIR for a close look.
Slides whose layout no truth slide uses are skipped (the exact geometry check
still covers them).
"""
import argparse
import json
import os
import sys

from PIL import Image, ImageChops, ImageStat

from dglcore import load_json

MASK_PH = {'sldNum', 'dt'}


def regions_for(spec, key, W, H, kept_footer):
    L = spec['layouts'][key]
    out = []
    for fe in spec['fixed_elements']:
        b = fe.get('box')
        if not b or (b[2] >= 0.98 * W and b[3] >= 0.98 * H):
            continue
        ph = (fe.get('ph') or [None])[0]
        if ph in MASK_PH or (ph == 'ftr' and kept_footer):
            continue
        if ph and ph not in L.get('show_fixed_placeholders', []):
            continue
        src = fe['source']
        if src == 'master' and (not L.get('show_master_shapes', True) or ph):
            continue
        if src == 'layout' and fe.get('layout') != key:
            continue
        if src == 'slide' and key not in (fe.get('layouts') or []):
            continue
        out.append((fe.get('name') or ph, b))
    return out


def crop(im, box, W, H, pad=2):
    sx, sy = im.width / W, im.height / H
    x0, y0 = max(0, int(box[0] * sx) - pad), max(0, int(box[1] * sy) - pad)
    x1, y1 = min(im.width, int((box[0] + box[2]) * sx) + pad), min(im.height, int((box[1] + box[3]) * sy) + pad)
    return im.crop((x0, y0, x1, y1))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('spec')
    ap.add_argument('deck')
    ap.add_argument('truth_prefix')
    ap.add_argument('after_prefix')
    ap.add_argument('--threshold', type=float, default=8.0)
    ap.add_argument('--out')
    ap.add_argument('--json')
    a = ap.parse_args()
    spec = load_json(a.spec)
    meta = load_json(a.deck + '.meta.json')
    W, H = spec['canvas']['w'], spec['canvas']['h']
    first_truth = {}
    for s in spec['slides']:
        first_truth.setdefault(s['layout'], s['n'])
    kept = {k['slide'] for k in meta.get('remap', {}).get('kept_footer_text', [])}
    fails, skipped, rows = [], [], []
    if a.out:
        os.makedirs(a.out, exist_ok=True)
    for n_str, key in sorted(meta['layout_for_slide'].items(), key=lambda kv: int(kv[0])):
        n = int(n_str)
        t = first_truth.get(key)
        if t is None:
            skipped.append(n)
            continue
        ti = Image.open(f'{a.truth_prefix}-{t}.jpg').convert('L')
        ai = Image.open(f'{a.after_prefix}-{n}.jpg').convert('L')
        if ai.size != ti.size:
            ai = ai.resize(ti.size)
        for name, box in regions_for(spec, key, W, H, n in kept):
            ct, ca = crop(ti, box, W, H), crop(ai, box, W, H)
            score = ImageStat.Stat(ImageChops.difference(ct, ca)).mean[0]
            rows.append(dict(slide=n, truth_slide=t, element=name, score=round(score, 2)))
            if score > a.threshold:
                fails.append(rows[-1])
                if a.out:
                    pair = Image.new('L', (ct.width * 2 + 6, ct.height), 255)
                    pair.paste(ct, (0, 0)); pair.paste(ca, (ct.width + 6, 0))
                    pair.save(os.path.join(a.out, f'slide{n}-{name}.png'.replace(' ', '_')))
    print('FIXED-ELEMENT PIXEL CHECK:', 'PASS' if not fails else 'FAIL', f'({len(rows)} regions)')
    for f in fails:
        print(f"  FAIL slide {f['slide']} '{f['element']}': differs from truth slide {f['truth_slide']} "
              f"(score {f['score']} > {a.threshold})")
    if skipped:
        print(f'  skipped slides {skipped}: no truth slide uses their layout (geometry check covers them)')
    if a.json:
        json.dump(dict(passed=not fails, rows=rows, fails=fails, skipped=skipped), open(a.json, 'w'), indent=1)
    sys.exit(0 if not fails else 1)


if __name__ == '__main__':
    main()
