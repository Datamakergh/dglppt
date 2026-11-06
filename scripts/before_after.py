#!/usr/bin/env python3
"""Before/after comparison PDF and contact sheets.

    python before_after.py BEFORE_PREFIX AFTER_PREFIX out.pdf [--labels labels.json]
                           [--sheets SHEET_PREFIX] [--per-sheet 12]

One page per slide: original on the left, restyled on the right, slide number
and truth layout above (labels.json: {"1": "m1/Title Slide", ...} — the
"layout_for_slide" map from the .meta.json works as is).
--sheets also writes grid images of the restyled deck (12 slides each) for a
fast visual pass over long decks: look at the sheets first, then open the
full-size render only for slides a check flagged or that look off.
"""
import argparse
import glob
import json

from PIL import Image, ImageDraw, ImageFont


def font(size):
    try:
        return ImageFont.load_default(size=size)
    except TypeError:  # Pillow < 10.1
        return ImageFont.load_default()


def numbered(prefix):
    files = glob.glob(prefix + '-*.jpg')
    return {int(f.rsplit('-', 1)[1].split('.')[0]): f for f in files}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('before')
    ap.add_argument('after')
    ap.add_argument('out')
    ap.add_argument('--labels')
    ap.add_argument('--sheets')
    ap.add_argument('--per-sheet', type=int, default=12)
    a = ap.parse_args()
    labels = {}
    if a.labels:
        d = json.load(open(a.labels))
        labels = d.get('layout_for_slide', d)
    before, after = numbered(a.before), numbered(a.after)
    pages = []
    for n in sorted(after):
        ai = Image.open(after[n]).convert('RGB')
        bi = Image.open(before[n]).convert('RGB') if n in before else Image.new('RGB', ai.size, 'white')
        h = ai.height
        bi = bi.resize((int(bi.width * h / bi.height), h))
        gap, top = 30, 90
        big, small = font(max(18, h // 30)), font(max(14, h // 40))
        page = Image.new('RGB', (bi.width + ai.width + gap * 3, h + top + gap), 'white')
        d = ImageDraw.Draw(page)
        d.text((gap, 15), f'Slide {n}  |  truth layout: {labels.get(str(n), "?")}', fill='black', font=big)
        d.text((gap, top - small.size - 8), 'Before', fill='gray', font=small)
        d.text((bi.width + gap * 2, top - small.size - 8), 'After', fill='gray', font=small)
        page.paste(bi, (gap, top))
        page.paste(ai, (bi.width + gap * 2, top))
        d.rectangle([gap - 1, top - 1, gap + bi.width, top + h], outline='lightgray')
        d.rectangle([bi.width + gap * 2 - 1, top - 1, bi.width + gap * 2 + ai.width, top + h], outline='lightgray')
        pages.append(page)
    if not pages:
        raise SystemExit('no renders found for ' + a.after)
    pages[0].save(a.out, save_all=True, append_images=pages[1:], resolution=150)
    print('before/after:', a.out, f'({len(pages)} pages)')
    if a.sheets:
        nums = sorted(after)
        cols = 4
        for s, i in enumerate(range(0, len(nums), a.per_sheet), 1):
            chunk = nums[i:i + a.per_sheet]
            ims = [Image.open(after[n]).convert('RGB') for n in chunk]
            w, h = 480, int(480 * ims[0].height / ims[0].width)
            rows = (len(ims) + cols - 1) // cols
            sheet = Image.new('RGB', (cols * (w + 10) + 10, rows * (h + 30) + 10), (235, 235, 235))
            d = ImageDraw.Draw(sheet)
            for k, (n, im) in enumerate(zip(chunk, ims)):
                x, y = 10 + (k % cols) * (w + 10), 10 + (k // cols) * (h + 30)
                sheet.paste(im.resize((w, h)), (x, y + 20))
                d.text((x, y + 2), f'{n}', fill='black', font=font(16))
            path = f'{a.sheets}-{s}.jpg'
            sheet.save(path, quality=85)
            print('contact sheet:', path)


if __name__ == '__main__':
    main()
