#!/usr/bin/env python3
"""Step 4a — move the truth design system into the target deck.

    python transplant_design.py truth.pptx target.pptx layout_map.json out.pptx [--spec truth_spec.json]

What it does, inside the target's own package (its slides, media, charts,
notes and animations never leave it):
  1. Freezes every slide placeholder's current geometry (writes an explicit
     xfrm where it inherited one), so nothing jumps before remapping decides.
  2. Detaches the target's old masters, layouts and theme and garbage-collects
     every part that only they used (never anything a slide still references:
     videos, GIFs, charts and their workbooks stay).
  3. Copies the truth theme(s), master(s), layouts and the media they use,
     renaming parts so nothing collides, rewriting relationships, content types
     and the presentation's master list (truth master/layout ids are kept;
     they are unique and >= 2147483648 by construction).
  4. Copies the truth slide size, default text style and table styles (target
     table styles the truth lacks are kept so no table loses its style).
  5. Points each slide at its mapped truth layout.
Writes out.pptx and out.pptx.meta.json (original size, frozen shapes,
original placeholder ids, part renames) for the next steps.
"""
import argparse
import copy
import posixpath
import re
import sys

from lxml import etree

from dglcore import (CT, P, PR, R, RT_LAYOUT, RT_MASTER, RT_NOTES, RT_SLIDE, RT_THEME,
                     RT_TABLESTYLES, Package, as_pptx_bytes, dump_json, inherited_box,
                     iter_shapes, load_json, ph_of, read_xfrm, sha1, shape_id, sp_tree,
                     unique_name, write_xfrm)

_RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
SKIP_TYPES = {RT_SLIDE, RT_NOTES, _RT + 'notesMaster', _RT + 'handoutMaster'}


def design_closure(truth: Package):
    """All parts reachable from the truth masters, minus slides and notes."""
    seen, stack = [], list(truth.master_parts())
    while stack:
        part = stack.pop(0)
        if part in seen:
            continue
        seen.append(part)
        for r in truth.rels(part):
            if r['mode'] == 'External' or r['type'] in SKIP_TYPES or not r['abs']:
                continue
            if truth.has(r['abs']) and r['abs'] not in seen:
                stack.append(r['abs'])
    return seen


def new_name_for(tgt: Package, old: str, reserved: set):
    d, f = posixpath.split(old)
    m = re.match(r'^(.*?)(\d*)\.([^.]+)$', f)
    stem, ext = (m.group(1), m.group(3)) if m else (f, 'bin')
    if d == 'ppt/media':
        stem = 'truth_' + (stem or 'media')
    name = unique_name(tgt, d, stem, ext, reserved)
    reserved.add(name)
    return name


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('truth')
    ap.add_argument('target')
    ap.add_argument('layout_map')
    ap.add_argument('out')
    ap.add_argument('--spec', help='truth_spec.json (to check the truth file is the one the spec came from)')
    a = ap.parse_args()

    truth_raw = open(a.truth, 'rb').read()
    truth = Package(as_pptx_bytes(truth_raw))
    tgt = Package(open(a.target, 'rb').read())
    lmap = load_json(a.layout_map)
    spec = load_json(a.spec) if a.spec else None
    if spec and spec['source']['sha1'] != sha1(truth_raw):
        sys.exit('truth file does not match truth_spec.json (different sha1) — re-run extract_truth.py')

    meta = dict(original_size=list(tgt.slide_size()), frozen=[], original_ph={}, renames={},
                layout_for_slide={}, warnings=[])

    # ---- 1. freeze placeholder geometry ------------------------------------
    slides = tgt.slide_parts()
    for n, sp in enumerate(slides, 1):
        root = tgt.xml(sp)
        for el, depth in iter_shapes(sp_tree(root), recurse=False):
            ph = ph_of(el)
            if not ph:
                continue
            meta['original_ph'].setdefault(str(n), {})[str(shape_id(el))] = list(ph)
            if read_xfrm(el) is None:
                box = inherited_box(tgt, sp, el)
                if box:
                    write_xfrm(el, box)
                    meta['frozen'].append(dict(slide=n, shape_id=shape_id(el), box=box))

    # ---- 2. detach old design ---------------------------------------------
    pres = tgt.main_part
    pres_root = tgt.xml(pres)
    prels = tgt.rels_root(pres)
    for rel in list(prels):
        if rel.get('Type') in (RT_MASTER, RT_THEME):
            prels.remove(rel)
    mlist = pres_root.find(P + 'sldMasterIdLst')
    if mlist is not None:
        for c in list(mlist):
            mlist.remove(c)
    slide_layout_rid = {}
    for sp in slides:
        root = tgt.rels_root(sp)
        for rel in list(root):
            if rel.get('Type') == RT_LAYOUT:
                slide_layout_rid[sp] = rel.get('Id')
                root.remove(rel)
    deleted = tgt.gc()
    meta['deleted_parts'] = deleted

    # ---- 3. copy truth design parts ----------------------------------------
    parts = design_closure(truth)
    reserved, rename = set(), {}
    for old in parts:
        rename[old] = new_name_for(tgt, old, reserved)
    meta['renames'] = rename
    for old, new in rename.items():
        tgt.set_blob(new, truth.blob(old))
        ctype = None
        tct = truth.xml('[Content_Types].xml')
        for o in tct.findall(CT + 'Override'):
            if o.get('PartName').lstrip('/') == old:
                ctype = o.get('ContentType')
        if ctype:
            tgt.set_override(new, ctype)
        else:
            ext = new.rsplit('.', 1)[-1]
            if not tgt.has_default(ext):
                dct = truth.content_type(old)
                if dct:
                    tgt.ensure_default(ext, dct)
        # relationships of the copied part
        src_rels = truth.rels(old)
        if src_rels:
            root = etree.Element(PR + 'Relationships', nsmap={None: 'http://schemas.openxmlformats.org/package/2006/relationships'})
            for r in src_rels:
                el = etree.SubElement(root, PR + 'Relationship')
                el.set('Id', r['id'])
                el.set('Type', r['type'])
                if r['mode'] == 'External':
                    el.set('Target', r['target'])
                    el.set('TargetMode', 'External')
                elif r['abs'] in rename:
                    el.set('Target', Package.relative(new, rename[r['abs']]))
                else:
                    root.remove(el)
                    meta['warnings'].append(f'{old}: dropped relationship to {r["abs"]} (outside the design)')
            tgt.set_xml(Package.rels_name(new), root)

    # master list + rels, in truth order, keeping truth ids
    truth_pres = truth.xml(truth.main_part)
    tml = truth_pres.find(P + 'sldMasterIdLst')
    if mlist is None:
        mlist = etree.Element(P + 'sldMasterIdLst')
        pres_root.insert(0, mlist)
    for mid in tml.findall(P + 'sldMasterId'):
        old_master = truth.rel_target(truth.main_part, mid.get(R + 'id'))['abs']
        rid = tgt.add_rel(pres, RT_MASTER, rename[old_master])
        el = etree.SubElement(mlist, P + 'sldMasterId')
        el.set('id', mid.get('id'))
        el.set(R + 'id', rid)
    first_theme = truth.theme_of(truth.master_parts()[0])
    tgt.add_rel(pres, RT_THEME, rename[first_theme])

    # ---- 4. size, default text style, table styles ---------------------------
    tsz, osz = truth_pres.find(P + 'sldSz'), pres_root.find(P + 'sldSz')
    for k in ('cx', 'cy', 'type'):
        if tsz.get(k) is not None:
            osz.set(k, tsz.get(k))
        elif k in osz.attrib:
            del osz.attrib[k]
    meta['new_size'] = [int(tsz.get('cx')), int(tsz.get('cy'))]
    tdts, odts = truth_pres.find(P + 'defaultTextStyle'), pres_root.find(P + 'defaultTextStyle')
    if tdts is not None:
        if odts is not None:
            odts.addprevious(copy.deepcopy(tdts))
            pres_root.remove(odts)
        else:
            pres_root.find(P + 'notesSz').addnext(copy.deepcopy(tdts))
    t_ts = truth.rels_of_type(truth.main_part, RT_TABLESTYLES)
    o_ts = tgt.rels_of_type(pres, RT_TABLESTYLES)
    if t_ts and t_ts[0]['abs'] and truth.has(t_ts[0]['abs']):
        troot = etree.fromstring(truth.blob(t_ts[0]['abs']))
        if o_ts and tgt.has(o_ts[0]['abs']):
            oroot = tgt.xml(o_ts[0]['abs'])
            have = {s.get('styleId') for s in troot}
            for s in oroot:
                if s.get('styleId') not in have:
                    troot.append(copy.deepcopy(s))
            tgt.set_xml(o_ts[0]['abs'], troot)
        else:
            tgt.set_xml('ppt/tableStyles.xml', troot)
            tgt.set_override('ppt/tableStyles.xml',
                             'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml')
            tgt.add_rel(pres, RT_TABLESTYLES, 'ppt/tableStyles.xml')

    # ---- 5. point slides at their truth layouts -----------------------------
    by_key = {}
    if spec:
        by_key = {k: L['part'] for k, L in spec['layouts'].items()}
    else:  # build keys the same way extract_truth does
        for mi, mp in enumerate(truth.master_parts(), 1):
            seen = {}
            for lp in truth.layout_parts(mp):
                nm = truth.xml(lp).find(P + 'cSld').get('name')
                seen[nm] = seen.get(nm, 0) + 1
                by_key[f'm{mi}/{nm}' + (f'#{seen[nm]}' if seen[nm] > 1 else '')] = lp
    for n, sp in enumerate(slides, 1):
        entry = lmap['slides'].get(str(n))
        key = entry and entry.get('layout')
        if not key or key not in by_key:
            sys.exit(f'slide {n}: no truth layout in the map (got {key!r}) — decide it with the user first')
        new_layout = rename[by_key[key]]
        root = tgt.rels_root(sp)
        el = etree.SubElement(root, PR + 'Relationship')
        el.set('Id', slide_layout_rid.get(sp) or 'rIdLayout')
        el.set('Type', RT_LAYOUT)
        el.set('Target', Package.relative(sp, new_layout))
        meta['layout_for_slide'][str(n)] = key

    tgt.save(a.out)
    dump_json(meta, a.out + '.meta.json')
    print(f'design transplanted: {a.out}')
    print(f"  slide size {meta['original_size']} -> {meta['new_size']}")
    print(f"  parts copied from truth: {len(rename)}, old parts removed: {len(deleted)}")
    print(f"  placeholders frozen: {len(meta['frozen'])}")
    for w in meta['warnings']:
        print('  warning:', w)


if __name__ == '__main__':
    main()
