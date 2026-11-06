#!/usr/bin/env python3
"""Step 4c-bis — give visuals that belong together the same size, alignment and placement.

Find candidate groups (writes proposals + one contact sheet per group):
    python visual_groups.py detect deck.pptx proposals.json --sheets SHEET_DIR

Apply the groups the model/user confirmed (edit proposals.json first: drop
wrong groups, add groups that show the same idea but don't match by pixels):
    python visual_groups.py apply deck.pptx truth_spec.json groups.json out.pptx

Visuals = pictures, GIFs (first frame), video/audio frames (poster frame).
Detected groups:
  identical  same image bytes (sha1)                         -> "across" when on several slides
  similar    perceptual hash (dHash) distance <= 10 and aspect ratio within 5%
  row/column 2+ visuals on one slide sharing a row (or column), aspect within
             25% and sizes within 2x (icon rows, logo strips, photo pairs)
Confidence: high = identical/similar or row with sizes within 15%; medium otherwise.

Apply rules (aspect ratios are never changed, images never stretched or cropped):
  row     every member gets the same box (median member size, shrunk if the row
          would leave the safe area); members are "contain"-fitted and centered
          in it; one shared top edge; equal gaps spread over the group's current
          span ("gap": "truth" on a group uses the truth gap instead); order
          kept. Captions (text boxes directly below/above a member) move with it:
          one shared caption line, each centered on its own visual.
  column  the same with the axes swapped (shared left edge, truth vgap).
  across  the same visual on several slides: every occurrence gets the box of
          the first occurrence; its position too unless that would overlap other
          content on that slide (then it keeps its own top-left; reported).
groups.json format: {"groups": [{"kind": "row"|"column"|"across",
  "members": [{"slide": 3, "shape_id": 3}, ...]}, ...]}
"""
import argparse
import os
import statistics

from dglcore import (Package, blip_part, dhash, dump_json, fit_contain, hamming, image_info,
                     iter_shapes, load_json, ph_of, read_xfrm, rects_overlap, sha1, shape_id,
                     shape_kind, shape_name, sp_tree, write_xfrm)

VISUAL = ('picture', 'video', 'audio')


def visuals(pkg):
    out = []
    for n, part in enumerate(pkg.slide_parts(), 1):
        for el, depth in iter_shapes(sp_tree(pkg.xml(part)), recurse=False):
            if shape_kind(el) not in VISUAL or ph_of(el):
                continue
            name = shape_name(el)
            if name.startswith('dglppt-fixed'):
                continue
            box = read_xfrm(el)
            img = blip_part(pkg, part, el)
            if not box or not img:
                continue
            blob = pkg.blob(img)
            info = image_info(blob) or {}
            out.append(dict(slide=n, shape_id=shape_id(el), name=name, kind=shape_kind(el), box=box,
                            sha1=sha1(blob), dhash=dhash(blob), gif=info.get('frames', 1) > 1,
                            aspect=box[2] / box[3] if box[3] else 1, part=part, img=img))
    return out


def y_overlap(a, b):
    lo, hi = max(a[1], b[1]), min(a[1] + a[3], b[1] + b[3])
    return max(0, hi - lo) / min(a[3], b[3])


def x_overlap(a, b):
    lo, hi = max(a[0], b[0]), min(a[0] + a[2], b[0] + b[2])
    return max(0, hi - lo) / min(a[2], b[2])


def similar_aspect(a, b, tol):
    return abs(a['aspect'] - b['aspect']) / max(a['aspect'], b['aspect']) <= tol


def detect(pkg):
    vs = visuals(pkg)
    groups = []
    # identical / similar across slides
    used = set()
    for i, a in enumerate(vs):
        if i in used:
            continue
        members = [a]
        for j in range(i + 1, len(vs)):
            b = vs[j]
            if j in used:
                continue
            same = a['sha1'] == b['sha1']
            near = hamming(a['dhash'], b['dhash']) <= 10 and similar_aspect(a, b, 0.05)
            if same or near:
                members.append(b)
                used.add(j)
        if len(members) > 1 and len({m['slide'] for m in members}) > 1:
            used.add(i)
            groups.append(dict(kind='across', confidence='high',
                               reason='identical image' if len({m['sha1'] for m in members}) == 1
                               else 'perceptually similar image',
                               members=members))
    # rows / columns on one slide
    by_slide = {}
    for v in vs:
        by_slide.setdefault(v['slide'], []).append(v)
    for n, items in by_slide.items():
        for kind, overlap, axis in (('row', y_overlap, 0), ('column', x_overlap, 1)):
            pool = sorted(items, key=lambda v: v['box'][axis])
            seen = set()
            for i, a in enumerate(pool):
                if i in seen:
                    continue
                members = [a]
                for j in range(i + 1, len(pool)):
                    b = pool[j]
                    size_ratio = max(a['box'][2] * a['box'][3], b['box'][2] * b['box'][3]) / max(
                        1, min(a['box'][2] * a['box'][3], b['box'][2] * b['box'][3]))
                    if overlap(a['box'], b['box']) >= 0.3 and similar_aspect(a, b, 0.25) and size_ratio <= 4:
                        members.append(b)
                        seen.add(j)
                if len(members) >= 2:
                    sizes = [m['box'][2] for m in members]
                    conf = 'high' if max(sizes) / min(sizes) <= 1.15 else 'medium'
                    groups.append(dict(kind=kind, confidence=conf,
                                       reason=f'{len(members)} visuals in one {kind}, similar shape', members=members))
    # a pair can be both a row and a column only if they overlap — keep the stronger
    return groups


def contact_sheet(pkg, group, path):
    from PIL import Image, ImageDraw
    import io
    thumbs = []
    for m in group['members']:
        try:
            im = Image.open(io.BytesIO(pkg.blob(m['img'])))
            im.seek(0)
            im = im.convert('RGB')
            im.thumbnail((220, 220))
        except Exception:
            im = Image.new('RGB', (220, 220), 'lightgray')
        thumbs.append((im, f"s{m['slide']} #{m['shape_id']}"))
    W = sum(t.width for t, _ in thumbs) + 20 * (len(thumbs) + 1)
    sheet = Image.new('RGB', (W, 270), 'white')
    d = ImageDraw.Draw(sheet)
    x = 20
    for im, label in thumbs:
        sheet.paste(im, (x, 20))
        d.text((x, 245), label, fill='black')
        x += im.width + 20
    sheet.save(path)


def strip_internal(m):
    return {k: m[k] for k in ('slide', 'shape_id', 'name', 'kind', 'box', 'gif')}


def cmd_detect(a):
    pkg = Package(open(a.deck, 'rb').read())
    groups = detect(pkg)
    if a.sheets:
        os.makedirs(a.sheets, exist_ok=True)
    out = []
    for i, g in enumerate(groups, 1):
        sheet = None
        if a.sheets:
            sheet = os.path.join(a.sheets, f'group-{i}.png')
            contact_sheet(pkg, g, sheet)
        out.append(dict(id=i, kind=g['kind'], confidence=g['confidence'], reason=g['reason'], sheet=sheet,
                        members=[strip_internal(m) for m in g['members']]))
    dump_json(dict(groups=out), a.out)
    for g in out:
        mem = ', '.join(f"s{m['slide']}#{m['shape_id']} {m['name']}" for m in g['members'])
        print(f"group {g['id']} [{g['kind']}, {g['confidence']}] {g['reason']}: {mem}")
    print(f'{len(out)} proposals -> {a.out}. Look at the sheets, drop wrong groups, add same-idea groups, '
          'then run apply.')


# ------------------------------------------------------------------ apply
def companions(root, member_el, box, taken):
    """Text boxes that caption this visual: centered under/over it, close by."""
    out = []
    cx = box[0] + box[2] / 2
    for el, depth in iter_shapes(sp_tree(root), recurse=False):
        if el is member_el or shape_kind(el) not in ('textbox', 'autoshape') or ph_of(el):
            continue
        if shape_name(el).startswith('dglppt-fixed') or id(el) in taken:
            continue
        b = read_xfrm(el)
        if not b:
            continue
        tcx = b[0] + b[2] / 2
        below = 0 <= b[1] - (box[1] + box[3]) <= 914400
        above = 0 <= box[1] - (b[1] + b[3]) <= 914400
        if (below or above) and abs(tcx - cx) <= box[2] / 2:
            out.append((el, b, 'below' if below else 'above'))
            taken.add(id(el))
    return out


def find_el(pkg, slide, sid):
    part = pkg.slide_parts()[slide - 1]
    root = pkg.xml(part)
    for el, _ in iter_shapes(sp_tree(root)):
        if shape_id(el) == sid:
            return root, el
    return root, None


def apply_line(pkg, spec, meta, g, axis, report):
    """axis 0 = row (left->right), 1 = column (top->bottom)."""
    items = []
    for m in g['members']:
        root, el = find_el(pkg, m['slide'], m['shape_id'])
        if el is None:
            report['skipped'].append(dict(m, why='shape not found'))
            continue
        items.append((root, el, read_xfrm(el)))
    if len(items) < 2:
        return
    items.sort(key=lambda t: t[2][axis])
    slide = g['members'][0]['slide']
    key = meta['layout_for_slide'][str(slide)]
    L = spec['layouts'][key]
    W, H = pkg.slide_size()
    safe = L.get('safe_area') or [0, 0, W, H]
    bw = int(statistics.median(b[2] for _, _, b in items))
    bh = int(statistics.median(b[3] for _, _, b in items))
    start = min(b[axis] for _, _, b in items)
    end = max(b[axis] + b[axis + 2] for _, _, b in items)
    unit = bw if axis == 0 else bh
    truth_gap = (L.get('gaps') or {}).get('hgap' if axis == 0 else 'vgap')
    n = len(items)
    if g.get('gap') == 'truth' and truth_gap:
        gap = int(truth_gap)
    else:  # default: keep the group's span, spread the members evenly
        gap = int((end - start - n * unit) / (n - 1))
        if gap < 0:
            gap = int(truth_gap or 0)
    span_limit = (safe[0] + safe[2] - start) if axis == 0 else (safe[1] + safe[3] - start)
    need = n * unit + (n - 1) * gap
    if need > span_limit:
        f = (span_limit - (n - 1) * gap) / (n * unit)
        bw, bh = int(bw * f), int(bh * f)
        unit = bw if axis == 0 else bh
        report['notes'].append(f"group {g.get('id')}: common box shrunk to {f:.2f} to stay in the safe area")
    cross = min(b[1 - axis] for _, _, b in items)
    taken = set()
    placed = []
    pos = start
    for root, el, b in items:
        slot = [pos, cross, bw, bh] if axis == 0 else [cross, pos, bw, bh]
        nb = fit_contain(b[2], b[3], slot)
        placed.append((root, el, b, slot, nb, companions(root, el, b, taken)))
        pos += unit + gap
    # captions: one shared line per side, each centered on its own visual (rows);
    # for columns they keep their offset to their visual
    below_gap = [cb[1] - (b[1] + b[3]) for _, _, b, _, _, cs in placed for _, cb, w in cs if w == 'below']
    above_gap = [b[1] - (cb[1] + cb[3]) for _, _, b, _, _, cs in placed for _, cb, w in cs if w == 'above']
    line_below = max(p[3][1] + p[3][3] for p in placed) + (int(statistics.median(below_gap)) if below_gap else 0)
    final = []
    for root, el, b, slot, nb, comps in placed:
        write_xfrm(el, nb)
        for cel, cb, where in comps:
            cx = nb[0] + nb[2] / 2 - cb[2] / 2
            if axis == 0:
                ny = line_below if where == 'below' else slot[1] - (int(statistics.median(above_gap)) if above_gap else 0) - cb[3]
            else:
                ny = (nb[1] + nb[3] + (cb[1] - (b[1] + b[3]))) if where == 'below' else nb[1] - (b[1] - cb[1])
            write_xfrm(cel, [cx, ny, cb[2], cb[3]])
        final.append(dict(slide=slide, shape_id=shape_id(el), slot=slot, box=nb,
                          companions=[shape_id(c[0]) for c in comps]))
    report['applied'].append(dict(id=g.get('id'), kind=g['kind'], box_size=[bw, bh], gap=gap, members=final))


def apply_across(pkg, g, report):
    first = g['members'][0]
    root0, el0 = find_el(pkg, first['slide'], first['shape_id'])
    if el0 is None:
        return
    ref = read_xfrm(el0)
    final = [dict(slide=first['slide'], shape_id=first['shape_id'], slot=ref, box=ref)]
    for m in g['members'][1:]:
        root, el = find_el(pkg, m['slide'], m['shape_id'])
        if el is None:
            continue
        b = read_xfrm(el)
        slot = list(ref)
        nb = fit_contain(b[2], b[3], slot)
        others = [read_xfrm(o) for o, _ in iter_shapes(sp_tree(root), recurse=False)
                  if o is not el and read_xfrm(o) and not ph_of(o)
                  and not shape_name(o).startswith('dglppt-fixed')]
        if any(rects_overlap(nb, o) for o in others):
            slot = [b[0], b[1], ref[2], ref[3]]
            nb = fit_contain(b[2], b[3], slot)
            report['notes'].append(f"group {g.get('id')}: slide {m['slide']} keeps its own position "
                                   "(the first occurrence's spot is taken there); size matched")
        write_xfrm(el, nb)
        final.append(dict(slide=m['slide'], shape_id=m['shape_id'], slot=slot, box=nb))
    report['applied'].append(dict(id=g.get('id'), kind='across', box_size=ref[2:], members=final))


def cmd_apply(a):
    pkg = Package(open(a.deck, 'rb').read())
    spec = load_json(a.spec)
    meta = load_json(a.deck + '.meta.json')
    groups = load_json(a.groups)['groups']
    report = dict(applied=[], skipped=[], notes=[])
    # rows/columns first (they decide sizes on a slide), then cross-slide matches
    for g in sorted(groups, key=lambda g: g['kind'] == 'across'):
        if g['kind'] == 'row':
            apply_line(pkg, spec, meta, g, 0, report)
        elif g['kind'] == 'column':
            apply_line(pkg, spec, meta, g, 1, report)
        elif g['kind'] == 'across':
            apply_across(pkg, g, report)
    pkg.save(a.out)
    meta['visual_groups'] = report
    dump_json(meta, a.out + '.meta.json')
    for r in report['applied']:
        print(f"group {r.get('id')} [{r['kind']}]: {len(r['members'])} visuals -> box "
              f"{r['box_size'][0] / 914400:.2f} x {r['box_size'][1] / 914400:.2f} in")
    for n in report['notes']:
        print('  note:', n)


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)
    d = sub.add_parser('detect'); d.add_argument('deck'); d.add_argument('out'); d.add_argument('--sheets')
    p = sub.add_parser('apply'); p.add_argument('deck'); p.add_argument('spec'); p.add_argument('groups')
    p.add_argument('out')
    a = ap.parse_args()
    cmd_detect(a) if a.cmd == 'detect' else cmd_apply(a)


if __name__ == '__main__':
    main()
