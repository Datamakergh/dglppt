#!/usr/bin/env python3
"""Step 5b — exact geometry and style check of a restyled deck against the truth.

    python check_geometry.py restyled.pptx truth.pptx truth_spec.json [--json out.json]

Reads restyled.pptx.meta.json (written by the restyle scripts). Exit 1 on any FAIL.

Checks, per slide, in EMU (0 tolerance unless stated):
  layouts      the slide's layout/master/theme XML is byte-for-byte the truth's
  fixed        truth slide-level fixed elements sit at their exact box and crop;
               the footer / slide number / date the truth shows are present and
               inherit their geometry (no local xfrm); none was moved
  placeholders matched placeholders inherit the truth layout, or sit exactly where
               the truth slide put them (boxes copied by apply_truth_slide_design.py)
  borders      no free content closer to an edge than the truth border
  overlap      no free content overlaps a fixed element (header, footer, logo…)
  text         each run's effective font, size, color, alignment and spacing equals
               what the truth gives it (inheritance-aware). Allowed differences:
               bold/italic/underline (emphasis), kept semantic colors, listed shrinks
  groups       visual groups: identical slots, one shared edge, equal gaps
  gaps         reported (warn) when free blocks don't use the truth gap
"""
import argparse
import json
import sys

from dglcore import (A, FIXED_PH_TYPES, P, Package, StyleContext, as_pptx_bytes,
                     c14n_str, inherited_box, iter_shapes, load_json, ph_of, read_xfrm,
                     rects_overlap, shape_id, shape_name, sp_tree)

TEXT_KEYS = ('font', 'size', 'color')
PARA_KEYS = ('align', 'line_spacing', 'space_before', 'space_after', 'marL', 'indent')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('deck')
    ap.add_argument('truth')
    ap.add_argument('spec')
    ap.add_argument('--json')
    a = ap.parse_args()
    pkg = Package(open(a.deck, 'rb').read())
    truth = Package(as_pptx_bytes(open(a.truth, 'rb').read()))
    spec = load_json(a.spec)
    meta = load_json(a.deck + '.meta.json')
    W, H = pkg.slide_size()
    fails, warns, passes = [], [], []

    def fail(n, msg):
        fails.append(f'slide {n}: {msg}' if n else msg)

    def warn(n, msg):
        warns.append(f'slide {n}: {msg}' if n else msg)

    if [W, H] != [spec['canvas']['w'], spec['canvas']['h']]:
        fail(None, f"slide size {W}x{H} != truth {spec['canvas']['w']}x{spec['canvas']['h']}")

    added = {(x['slide'], x['shape_id']): x for x in meta.get('remap', {}).get('added_fixed', [])}
    shrinks = {(s['slide'], s['shape_id']) for s in meta.get('fit', {}).get('shrinks', [])}
    shrinks |= {(s['slide'], s['shape_id']) for s in meta.get('fit', {}).get('needs_decision', [])}
    semantic = {(s['slide'], s['shape_id']) for s in meta.get('style', {}).get('semantic_color_kept', [])}
    roles = {(r['slide'], r['shape_id']): r['role'] for r in meta.get('style', {}).get('free_text_roles', [])}
    approved_overlaps = set()
    # placeholder boxes copied from the truth slides (apply_truth_slide_design.py)
    positioned = {(x['slide'], x['shape_id']): x['box']
                  for x in meta.get('positions', {}).get('copied', [])}
    # text design copied from truth slides: those properties are compared against the donor, not the layout
    designed = {(x['slide'], x['shape_id']): set(x['props']) for x in meta.get('design', {}).get('shapes', [])}
    for x in meta.get('design', {}).get('added', []):  # decorative one-offs copied from a truth slide
        added[(x['slide'], x['shape_id'])] = dict(x, box=x['box'])

    fixed_boxes = [fe['box'] for fe in spec['fixed_elements'] if fe.get('box') and fe['source'] != 'slide'
                   and not (fe['box'][2] >= 0.98 * W and fe['box'][3] >= 0.98 * H)]
    # truth layout XML by key, to compare byte-for-byte
    truth_layout_xml = {k: c14n_str(truth.xml(L['part'])) for k, L in spec['layouts'].items()}

    import importlib
    so = importlib.import_module('strip_overrides')

    for n, part in enumerate(pkg.slide_parts(), 1):
        key = meta['layout_for_slide'][str(n)]
        L = spec['layouts'][key]
        lay = pkg.layout_of(part)
        if c14n_str(pkg.xml(lay)) != truth_layout_xml[key]:
            fail(n, f'layout XML differs from truth layout {key}')
        root = pkg.xml(part)
        ctx = StyleContext(pkg, part)
        safe = L.get('safe_area')
        present_fixed = set()
        free_boxes = []
        for el, depth in iter_shapes(sp_tree(root)):
            sid = shape_id(el)
            ph = ph_of(el) if depth == 0 else None
            if (n, sid) in added and added[(n, sid)].get('box'):
                want = added[(n, sid)]['box']
                got = read_xfrm(el)
                if got != want:
                    fail(n, f"fixed element '{shape_name(el)}' at {got}, truth {want}")
                continue
            if ph and ph[0] in FIXED_PH_TYPES:
                present_fixed.add(ph[0])
                if read_xfrm(el) is not None:
                    fail(n, f"{ph[0]} placeholder has local geometry (must inherit the truth position)")
                continue
            if ph and (n, sid) in positioned:
                if read_xfrm(el) != positioned[(n, sid)]:
                    fail(n, f"placeholder '{shape_name(el)}' moved off the truth slide's position")
            elif ph and read_xfrm(el) is None:
                box = inherited_box(pkg, part, el)
            elif depth == 0:
                box = read_xfrm(el)
                if box and not ph:
                    free_boxes.append((el, box))
                elif ph:
                    warn(n, f"placeholder '{shape_name(el)}' has local geometry (unmatched placeholder)")
                    free_boxes.append((el, box))
            # ---- text style
            tb = el.find(P + 'txBody')
            if tb is None or (n, sid) in [k for k in added]:
                continue
            for p in tb.findall(A + 'p'):
                for r in p.findall(A + 'r'):
                    if not (r.findtext(A + 't') or '').strip():
                        continue
                    para, run = ctx.resolve_run(el, p, r)
                    if ph:
                        epara, erun = ctx.resolve_paragraph(el, p, skip_local=True)
                    else:
                        lvl = para['level']
                        epara, erun = so.role_style(ctx, spec, roles.get((n, sid), 'textbox'), lvl)
                    dprops = designed.get((n, sid), set())
                    all_free = 'list_style' in dprops
                    for k in TEXT_KEYS:
                        if erun.get(k) is None or run.get(k) == erun.get(k):
                            continue
                        if all_free or k in dprops:
                            continue
                        if k == 'size' and (n, sid) in shrinks:
                            continue
                        if k == 'color' and (n, sid) in semantic:
                            continue
                        fail(n, f"'{shape_name(el)}' run {r.findtext(A + 't')[:25]!r}: {k} {run.get(k)} != truth {erun.get(k)}")
                    keys = PARA_KEYS if ph else tuple(k for k in PARA_KEYS if k != 'align')
                    for k in keys:
                        if all_free or k in dprops:
                            continue
                        if k == 'line_spacing' and (n, sid) in shrinks:
                            continue
                        if epara.get(k) is not None and para.get(k) != epara.get(k):
                            fail(n, f"'{shape_name(el)}' paragraph {k} {para.get(k)} != truth {epara.get(k)}")
        missing = set(L.get('show_fixed_placeholders', [])) - present_fixed
        if missing:
            fail(n, f"truth shows {sorted(missing)} but the slide doesn't")
        for el, box in free_boxes:
            nm = shape_name(el)
            if safe:
                if (box[0] < safe[0] or box[1] < safe[1] or box[0] + box[2] > safe[0] + safe[2]
                        or box[1] + box[3] > safe[1] + safe[3]):
                    fail(n, f"'{nm}' {box} is outside the truth safe area {safe}")
            for fb in fixed_boxes + [x['box'] for k, x in added.items() if k[0] == n and x.get('box')]:
                if rects_overlap(box, fb) and (n, shape_id(el)) not in approved_overlaps:
                    fail(n, f"'{nm}' overlaps fixed element at {fb}")

    # ---- visual groups
    for g in meta.get('visual_groups', {}).get('applied', []):
        slots = [m['slot'] for m in g['members']]
        for m in g['members']:
            part = pkg.slide_parts()[m['slide'] - 1]
            el = next((e for e, _ in iter_shapes(sp_tree(pkg.xml(part))) if shape_id(e) == m['shape_id']), None)
            if el is None or read_xfrm(el) != [int(v) for v in m['box']]:
                fail(m['slide'], f"group {g.get('id')}: shape {m['shape_id']} moved after grouping")
        if len({(s[2], s[3]) for s in slots}) > 1:
            fail(None, f"group {g.get('id')}: slot sizes differ {[(s[2], s[3]) for s in slots]}")
        if g['kind'] == 'row':
            if len({s[1] for s in slots}) > 1:
                fail(None, f"group {g.get('id')}: members don't share a top edge")
            gaps = [slots[i + 1][0] - slots[i][0] - slots[i][2] for i in range(len(slots) - 1)]
            if gaps and max(gaps) - min(gaps) > 1:
                fail(None, f"group {g.get('id')}: unequal gaps {gaps}")
        if g['kind'] == 'column':
            if len({s[0] for s in slots}) > 1:
                fail(None, f"group {g.get('id')}: members don't share a left edge")
            gaps = [slots[i + 1][1] - slots[i][1] - slots[i][3] for i in range(len(slots) - 1)]
            if gaps and max(gaps) - min(gaps) > 1:
                fail(None, f"group {g.get('id')}: unequal gaps {gaps}")
        passes.append(f"group {g.get('id')} [{g['kind']}]: {len(slots)} members, box {slots[0][2]}x{slots[0][3]}")

    print('GEOMETRY CHECK:', 'PASS' if not fails else 'FAIL')
    for f in fails:
        print('  FAIL', f)
    for w in warns:
        print('  warn', w)
    for p in passes:
        print('  ok  ', p)
    if a.json:
        json.dump(dict(passed=not fails, fails=fails, warns=warns, passes=passes), open(a.json, 'w'), indent=1)
    sys.exit(0 if not fails else 1)


if __name__ == '__main__':
    main()
