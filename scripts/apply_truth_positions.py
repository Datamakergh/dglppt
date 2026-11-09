#!/usr/bin/env python3
"""Step 4b-bis — copy the truth SLIDES' placeholder positions onto matching target slides.

    python apply_truth_positions.py in.pptx truth.pptx truth_spec.json out.pptx [--skip-layouts KEY ...]

remap_placeholders.py clears every placeholder's local geometry so it inherits the
truth LAYOUT. But a truth deck is often designed on its slides: the author moved the
title or body on a slide, and that placement is part of the design. This step puts it
back: for each target slide mapped to truth layout L, the first truth slide that uses L
donates the box (x, y, w, h) of every placeholder that carries its own geometry, matched
by placeholder type + idx (remap has already rewritten the target's type/idx to the
truth's). Placeholders the truth slide leaves to the layout stay inherited.

Not touched: footer / date / slide-number placeholders (they must inherit), and
placeholders with no truth match. Text, runs and shape ids are never touched.

Writes the copied boxes to meta['positions'] so check_geometry can verify them exactly.
"""
import argparse
import copy

from dglcore import (A, FIXED_PH_TYPES, P, Package, as_pptx_bytes, dump_json, iter_shapes,
                     load_json, ph_of, read_xfrm, shape_id, shape_name, sp_tree)


def ph_key(el):
    ph = ph_of(el)
    return None if ph is None else (ph[0], str(ph[1]))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('deck')
    ap.add_argument('truth')
    ap.add_argument('spec')
    ap.add_argument('out')
    ap.add_argument('--skip-layouts', nargs='*', default=[],
                    help='layout keys that should keep inheriting the layout position')
    a = ap.parse_args()

    tgt = Package(open(a.deck, 'rb').read())
    truth = Package(as_pptx_bytes(open(a.truth, 'rb').read()))
    spec = load_json(a.spec)
    meta = load_json(a.deck + '.meta.json')

    # first truth slide per layout key
    donor = {}
    for s, part in zip(spec['slides'], truth.slide_parts()):
        donor.setdefault(s['layout'], part)

    report = dict(copied=[], no_donor=[], inherited=[], unmatched=[])
    for n, part in enumerate(tgt.slide_parts(), 1):
        key = meta['layout_for_slide'][str(n)]
        if key in a.skip_layouts:
            continue
        dpart = donor.get(key)
        if dpart is None:
            report['no_donor'].append(dict(slide=n, layout=key))
            continue
        boxes = {}
        for el, depth in iter_shapes(sp_tree(truth.xml(dpart))):
            k = ph_key(el) if depth == 0 else None
            if k and k[0] not in FIXED_PH_TYPES and read_xfrm(el) is not None:
                boxes[k] = copy.deepcopy(el.find(P + 'spPr').find(A + 'xfrm'))
        root = tgt.xml(part)
        for el, depth in iter_shapes(sp_tree(root)):
            if depth != 0:
                continue
            k = ph_key(el)
            if not k or k[0] in FIXED_PH_TYPES:
                continue
            xf = boxes.get(k)
            if xf is None:
                report['inherited'].append(dict(slide=n, shape_id=shape_id(el), name=shape_name(el)))
                continue
            spPr = el.find(P + 'spPr')
            old = spPr.find(A + 'xfrm')
            if old is not None:
                spPr.remove(old)
            spPr.insert(0, xf)
            report['copied'].append(dict(slide=n, shape_id=shape_id(el), name=shape_name(el),
                                         box=read_xfrm(el)))

    tgt.save(a.out)
    meta['positions'] = report
    dump_json(meta, a.out + '.meta.json')
    print(f"truth-slide positions copied: {len(report['copied'])}, "
          f"left to the layout: {len(report['inherited'])}, layouts without a truth slide: "
          f"{len(report['no_donor'])}")
    for c in report['copied']:
        print(f"  slide {c['slide']} '{c['name']}' -> {c['box']}")


if __name__ == '__main__':
    main()
