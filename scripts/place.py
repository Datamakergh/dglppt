#!/usr/bin/env python3
"""Step 4c — fit free content inside the truth safe area of each slide.

    python place.py styled.pptx truth_spec.json out.pptx [--plan plan.json]
                    [--halign left|center] [--max-scale 1.333] [--report place.json]

"Free content" = everything on a slide that is not a matched placeholder and not
a truth fixed element: pictures, GIFs, video/audio frames, charts, tables, text
boxes, shapes, groups and unmatched placeholders. Matched placeholders already
sit exactly on the truth layout geometry (remap step).

Per slide the free content is moved as ONE block, so its internal arrangement
survives: one uniform scale (aspect ratios never change, nothing is cropped)
and one translation into the target area, which is
  * the layout's content area (body placeholders / master body zone) when the
    slide doesn't use a body placeholder, or
  * the safe area below the last used placeholder when it does,
aligned to the truth left content line (or centered with --halign center).
Scale is capped at --max-scale (default: how much the canvas grew, at least
1.0) so pictures aren't blown up. Groups scale through their own xfrm; tables scale
their column widths and row heights too.

Media rule: a video/audio/GIF frame shrunk below 85% of its size is listed as
"needs approval" (stop and ask), like overflowing text.

plan.json overrides per slide (EMU): {"3": {"area": [x, y, w, h], "halign": "center",
"boxes": {"7": [x, y, w, h]}, "skip": false}}
"""
import argparse

from dglcore import (A, FIXED_PH_TYPES, TITLE_PH_TYPES, Package, dump_json, image_info,
                     inherited_box, iter_shapes, load_json, ph_of, read_xfrm, rects_overlap,
                     shape_id, shape_kind, shape_name, sp_tree, write_xfrm, blip_part)

MEDIA_FLOOR = 0.85


def is_media(pkg, part, el):
    k = shape_kind(el)
    if k in ('video', 'audio'):
        return True
    if k == 'picture':
        img = blip_part(pkg, part, el)
        info = image_info(pkg.blob(img)) if img else None
        return bool(info and info['frames'] > 1)
    return False


def scale_table(el, s):
    tbl = el.find('.//' + A + 'tbl')
    if tbl is None:
        return
    for gc in tbl.iter(A + 'gridCol'):
        gc.set('w', str(int(round(int(gc.get('w')) * s))))
    for tr in tbl.findall(A + 'tr'):
        tr.set('h', str(int(round(int(tr.get('h')) * s))))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('deck')
    ap.add_argument('spec')
    ap.add_argument('out')
    ap.add_argument('--plan')
    ap.add_argument('--halign', default='left', choices=['left', 'center'])
    ap.add_argument('--max-scale', type=float)
    ap.add_argument('--report')
    a = ap.parse_args()

    pkg = Package(open(a.deck, 'rb').read())
    spec = load_json(a.spec)
    meta = load_json(a.deck + '.meta.json')
    plan = load_json(a.plan) if a.plan else {}
    W, H = pkg.slide_size()
    ow, oh = meta.get('original_size', [W, H])
    max_scale = a.max_scale or max(1.0, W / ow, H / oh)
    fixed_ids = {(x['slide'], x['shape_id']) for x in meta.get('remap', {}).get('added_fixed', [])}
    positioned = {(x['slide'], x['shape_id']) for x in meta.get('positions', {}).get('copied', [])}
    report = dict(slides=[], needs_approval=[], overlaps=[])

    fixed_boxes = [fe['box'] for fe in spec['fixed_elements'] if fe.get('box') and fe['source'] != 'slide'
                   and not (fe['box'][2] >= 0.98 * W and fe['box'][3] >= 0.98 * H)]

    for n, part in enumerate(pkg.slide_parts(), 1):
        root = pkg.xml(part)
        key = meta['layout_for_slide'][str(n)]
        L = spec['layouts'][key]
        sp = plan.get(str(n), {})
        if sp.get('skip'):
            continue
        free, used_ph = [], []
        for el, depth in iter_shapes(sp_tree(root), recurse=False):
            sid = shape_id(el)
            if (n, sid) in fixed_ids:
                continue
            ph = ph_of(el)
            if ph and (n, sid) in positioned:  # sits where the truth slide put it
                if ph[0] not in FIXED_PH_TYPES:
                    used_ph.append((ph, read_xfrm(el)))
                continue
            if ph and read_xfrm(el) is None:  # matched: inherits truth geometry
                if ph[0] not in FIXED_PH_TYPES:
                    used_ph.append((ph, inherited_box(pkg, part, el)))
                continue
            if ph and ph[0] in FIXED_PH_TYPES:
                continue
            box = read_xfrm(el)
            if box:
                free.append((el, box))
        if not free:
            continue

        safe = L.get('safe_area') or [0, 0, W, H]
        uses_body = any(p[0] not in TITLE_PH_TYPES for p, _ in used_ph)
        if sp.get('area'):
            area = sp['area']
        elif not uses_body and L.get('content_area'):
            area = L['content_area']
        else:
            bottom = max([b[1] + b[3] for _, b in used_ph if b] + [safe[1]])
            gap = (L.get('gaps') or {}).get('vgap') or 0
            top = bottom + gap if used_ph else safe[1]
            area = [safe[0], top, safe[2], safe[1] + safe[3] - top]
        bx0 = min(b[0] for _, b in free); by0 = min(b[1] for _, b in free)
        bx1 = max(b[0] + b[2] for _, b in free); by1 = max(b[1] + b[3] for _, b in free)
        bw, bh = max(bx1 - bx0, 1), max(by1 - by0, 1)
        s = min(area[2] / bw, area[3] / bh, max_scale)
        halign = sp.get('halign', a.halign)
        dx = area[0] + ((area[2] - bw * s) / 2 if halign == 'center' else 0)
        dy = area[1]
        moved = []
        for el, box in free:
            sid = str(shape_id(el))
            if sid in sp.get('boxes', {}):
                nb = sp['boxes'][sid]
            else:
                nb = [dx + (box[0] - bx0) * s, dy + (box[1] - by0) * s, box[2] * s, box[3] * s]
            write_xfrm(el, nb)
            if shape_kind(el) == 'table':
                scale_table(el, nb[2] / box[2] if box[2] else 1)
            ratio = nb[2] / box[2] if box[2] else 1
            if is_media(pkg, part, el) and ratio < MEDIA_FLOOR:
                report['needs_approval'].append(dict(slide=n, shape_id=int(sid), name=shape_name(el),
                                                     scale=round(ratio, 3),
                                                     why='media frame below 85% of its size'))
            moved.append(dict(shape_id=int(sid), name=shape_name(el), kind=shape_kind(el),
                              before=box, after=[int(round(v)) for v in nb]))
            for fb in fixed_boxes + [b for _, b in used_ph if b]:
                if rects_overlap([int(v) for v in nb], fb):
                    report['overlaps'].append(dict(slide=n, shape_id=int(sid), name=shape_name(el), with_box=fb))
        report['slides'].append(dict(slide=n, layout=key, area=[int(v) for v in area], scale=round(s, 4),
                                     moved=moved))

    pkg.save(a.out)
    meta['place'] = report
    dump_json(meta, a.out + '.meta.json')
    if a.report:
        dump_json(report, a.report)
    print(f"placed free content on {len(report['slides'])} slides")
    for s in report['slides']:
        print(f"  slide {s['slide']}: {len(s['moved'])} shapes, scale {s['scale']}")
    for x in report['needs_approval']:
        print(f"  NEEDS APPROVAL slide {x['slide']} '{x['name']}': {x['why']} ({x['scale']})")
    for o in report['overlaps']:
        print(f"  overlap: slide {o['slide']} '{o['name']}' touches a placeholder or fixed element")


if __name__ == '__main__':
    main()
