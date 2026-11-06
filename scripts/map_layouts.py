#!/usr/bin/env python3
"""Propose a target-slide -> truth-layout map.

    python map_layouts.py truth_spec.json target.pptx layout_map.json

Matches by structure (which placeholders the slide uses and what free content
it holds) plus layout-name similarity. Prints a review table; only rows marked
"low" (or "none") need the user's decision. Edit layout_map.json to override:
{"slides": {"3": {"layout": "m1/Title Only", ...}}}
"""
import re
import sys
from collections import Counter

from dglcore import (FIXED_PH_TYPES, TITLE_PH_TYPES, dump_json, iter_shapes, load_json,
                     open_package, ph_of, shape_kind, sp_tree)

BODYISH = {'body', 'obj'}


def norm_ph(t):
    if t in TITLE_PH_TYPES:
        return 'title' if t == 'title' else 'ctrTitle'
    if t in BODYISH:
        return 'body'
    return t


def tokens(name):
    return set(re.findall(r'[a-z]+', (name or '').lower())) - {'and', 'with', 'the', 'm'}


def slide_features(pkg, part):
    root = pkg.xml(part)
    phs, free = Counter(), Counter()
    for el, depth in iter_shapes(sp_tree(root), recurse=False):
        ph = ph_of(el)
        if ph:
            if ph[0] not in FIXED_PH_TYPES:
                phs[norm_ph(ph[0])] += 1
        else:
            free[shape_kind(el)] += 1
    lay = pkg.layout_of(part)
    name = pkg.xml(lay).find('{http://schemas.openxmlformats.org/presentationml/2006/main}cSld').get('name') if lay else ''
    return phs, free, name


def layout_features(L):
    return Counter(norm_ph(p['type']) for p in L['placeholders'] if p['type'] not in FIXED_PH_TYPES)


def score(phs, free, name, L):
    lf = layout_features(L)
    inter = sum((phs & lf).values())
    union = sum((phs | lf).values()) or 1
    s = inter / union
    # slides with free content but only a title want a title-only layout, not a body layout
    if free and phs.get('body', 0) == 0 and lf.get('body', 0) > 0:
        s -= 0.25
    tn, ln = tokens(name), tokens(L['name'])
    if tn and tn == ln:
        s += 0.35
    elif tn & ln:
        s += 0.1 * len(tn & ln) / max(len(tn | ln), 1)
    # vertical / caption / comparison layouts only when the names say so
    if ln & {'vertical'} and not tn & {'vertical'}:
        s -= 0.3
    return round(s, 3)


def main():
    if len(sys.argv) != 4:
        sys.exit(__doc__)
    spec = load_json(sys.argv[1])
    pkg = open_package(sys.argv[2])
    out = dict(truth=spec['source'], slides={})
    rows = []
    for n, part in enumerate(pkg.slide_parts(), 1):
        phs, free, name = slide_features(pkg, part)
        ranked = sorted(((score(phs, free, name, L), k) for k, L in spec['layouts'].items()), reverse=True)
        best, key = ranked[0] if ranked else (0, None)
        second = ranked[1][0] if len(ranked) > 1 else 0
        if best >= 0.9 and best - second >= 0.1:
            conf = 'high'
        elif best >= 0.55:
            conf = 'medium'
        elif best > 0.2:
            conf = 'low'
        else:
            conf, key = 'none', None
        looks = ', '.join(f'{k}×{v}' for k, v in sorted(phs.items())) or 'no placeholders'
        if free:
            looks += ' + ' + ', '.join(f'{k}×{v}' for k, v in sorted(free.items()))
        out['slides'][str(n)] = dict(layout=key, confidence=conf, target_layout=name,
                                     looks_like=looks,
                                     alternatives=[k for _, k in ranked[1:3]])
        rows.append(f'| {n} | {name} — {looks} | {key} | {conf} |')
    dump_json(out, sys.argv[3])
    print('| Target slide | Looks like | Truth layout | Confidence |')
    print('|---|---|---|---|')
    print('\n'.join(rows))
    print(f'\nmap written: {sys.argv[3]} (ask the user only about low / none rows)')


if __name__ == '__main__':
    main()
