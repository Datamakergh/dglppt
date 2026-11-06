#!/usr/bin/env python3
"""Step 4b — remap placeholders and place the truth's fixed elements.

    python remap_placeholders.py transplanted.pptx truth.pptx truth_spec.json out.pptx
                                 [--approved approved.json] [--report remap.json]

* Each slide placeholder is matched to a placeholder of its new truth layout
  (same idx first, then the same kind: title, body/content, subtitle, picture,
  footer…). Its (type, idx) is rewritten to the truth layout's and its frozen
  xfrm is cleared so it inherits the truth geometry exactly. Placeholders with
  no counterpart keep their frozen geometry and are reported (place.py fits
  them into the safe area).
* Footer / slide-number / date placeholders the truth slides show are added to
  every slide as slide-level placeholders (a layout placeholder alone is never
  shown on a slide). Text comes from the truth slides. A target's own footer
  text is kept (content) and reported.
* Slide-level fixed elements of the truth deck (a badge or logo placed on each
  truth slide rather than the master) are copied at their exact box.
* Shapes the user approved for removal (approved.json "remove": old footer,
  old logo…) are deleted. Nothing else is deleted. Shapes are edited in place,
  so shape ids — and the animations that point at them — survive.
"""
import argparse
import copy

from lxml import etree

from dglcore import (A, FIXED_PH_TYPES, P, R, RT_IMAGE, Package, as_pptx_bytes,
                     clear_xfrm, dump_json, iter_shapes, load_json, next_shape_id, paragraph_text,
                     ph_el, ph_of, read_xfrm, sha1, shape_id, shape_name, sp_tree, unique_name)

KIND = {'title': 'title', 'ctrTitle': 'title', 'subTitle': 'subtitle', 'body': 'body', 'obj': 'body',
        'pic': 'pic', 'chart': 'body', 'tbl': 'body', 'dgm': 'body', 'media': 'body', 'clipArt': 'pic',
        'ftr': 'ftr', 'sldNum': 'sldNum', 'dt': 'dt', 'hdr': 'hdr'}


def layout_phs(layout_root):
    out = []
    for el, _ in iter_shapes(sp_tree(layout_root), recurse=False):
        ph = ph_of(el)
        if ph:
            out.append(dict(el=el, type=ph[0], idx=ph[1], explicit_type=ph_el(el).get('type')))
    return out


def find_match(ph, cands, used):
    t, idx = ph
    for c in cands:
        if id(c['el']) in used:
            continue
        if c['idx'] == idx and KIND.get(c['type']) == KIND.get(t):
            return c
    for c in cands:
        if id(c['el']) in used:
            continue
        if KIND.get(c['type']) == KIND.get(t):
            return c
    # a subtitle can live in a body placeholder and vice versa
    alt = {'subtitle': 'body', 'body': 'subtitle', 'pic': 'body'}.get(KIND.get(t))
    for c in cands:
        if id(c['el']) not in used and KIND.get(c['type']) == alt:
            return c
    return None


def set_ph(el, cand):
    phe = ph_el(el)
    if cand['explicit_type']:
        phe.set('type', cand['explicit_type'])
    elif 'type' in phe.attrib:
        del phe.attrib['type']
    if cand['idx']:
        phe.set('idx', str(cand['idx']))
    elif 'idx' in phe.attrib:
        del phe.attrib['idx']


def copy_picture_part(tgt: Package, truth: Package, truth_part, rid, cache):
    rel = truth.rel_target(truth_part, rid)
    blob = truth.blob(rel['abs'])
    key = sha1(blob)
    if key not in cache:
        ext = rel['abs'].rsplit('.', 1)[-1]
        name = unique_name(tgt, 'ppt/media', 'truth_fixed', ext)
        tgt.set_blob(name, blob)
        if not tgt.has_default(ext):
            ct = truth.content_type(rel['abs'])
            if ct:
                tgt.ensure_default(ext, ct)
        cache[key] = name
    return cache[key]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('transplanted')
    ap.add_argument('truth')
    ap.add_argument('spec')
    ap.add_argument('out')
    ap.add_argument('--approved')
    ap.add_argument('--report')
    a = ap.parse_args()

    tgt = Package(open(a.transplanted, 'rb').read())
    truth = Package(as_pptx_bytes(open(a.truth, 'rb').read()))
    spec = load_json(a.spec)
    meta = load_json(a.transplanted + '.meta.json')
    appr = load_json(a.approved) if a.approved else {}
    removals = {(x['slide'], x['shape_id']) for x in appr.get('remove', [])}
    report = dict(remapped=[], unmatched=[], added_fixed=[], kept_footer_text=[], removed=[])
    media_cache = {}

    for n, sp in enumerate(tgt.slide_parts(), 1):
        root = tgt.xml(sp)
        key = meta['layout_for_slide'][str(n)]
        L = spec['layouts'][key]
        cands = layout_phs(tgt.xml(tgt.layout_of(sp)))
        used = set()

        # never reuse an id, not even one of a removed shape: animations and the
        # content check identify shapes by id
        id_counter = [next_shape_id(root)]

        def new_id():
            id_counter[0] += 1
            return id_counter[0] - 1

        # approved removals first (old footer/logo leftovers)
        for el, _ in list(iter_shapes(sp_tree(root))):
            if (n, shape_id(el)) in removals:
                report['removed'].append(dict(slide=n, shape_id=shape_id(el), name=shape_name(el)))
                el.getparent().remove(el)

        present_fixed = set()
        for el, _ in iter_shapes(sp_tree(root), recurse=False):
            ph = ph_of(el)
            if not ph:
                continue
            cand = find_match(ph, cands, used)
            if cand is None:
                report['unmatched'].append(dict(slide=n, shape_id=shape_id(el), name=shape_name(el), ph=list(ph),
                                                box=read_xfrm(el)))
                continue
            used.add(id(cand['el']))
            set_ph(el, cand)
            clear_xfrm(el)
            report['remapped'].append(dict(slide=n, shape_id=shape_id(el), frm=list(ph),
                                           to=[cand['type'], cand['idx']]))
            if cand['type'] in FIXED_PH_TYPES:
                present_fixed.add(cand['type'])
                txt = ' '.join(paragraph_text(p) for p in el.iter(A + 'p')).strip()
                truth_tx = spec.get('fixed_placeholder_txbody', {}).get(cand['type'])
                if cand['type'] == 'ftr' and txt and truth_tx and txt not in truth_tx:
                    report['kept_footer_text'].append(dict(slide=n, shape_id=shape_id(el), text=txt))

        # footer / slide number / date the truth shows
        for typ in L.get('show_fixed_placeholders', []):
            if typ in present_fixed:
                continue
            cand = next((c for c in cands if c['type'] == typ and id(c['el']) not in used), None)
            if cand is None:
                continue
            new = copy.deepcopy(cand['el'])
            cnv = new.find('.//' + P + 'cNvPr')
            cnv.set('id', str(new_id()))
            for k in ('descr', 'title'):
                cnv.attrib.pop(k, None)
            clear_xfrm(new)
            tx = spec.get('fixed_placeholder_txbody', {}).get(typ)
            if tx:
                old = new.find(P + 'txBody')
                newtx = etree.fromstring(tx)
                if old is not None:
                    old.addprevious(newtx)
                    new.remove(old)
                else:
                    new.append(newtx)
            # layouts mark these userDrawn; slide copies must not be
            nvpr = new.find('.//' + P + 'nvPr')
            nvpr.attrib.pop('userDrawn', None)
            sp_tree(root).append(new)
            report['added_fixed'].append(dict(slide=n, type=typ, shape_id=int(cnv.get('id'))))

        # slide-level fixed elements of the truth deck
        for fe in spec['fixed_elements']:
            if fe.get('source') != 'slide':
                continue
            truth_layouts = {s['layout'] for s in spec['slides']}
            on_layout = key in (fe.get('layouts') or [])
            # layouts the truth slides never use: follow the content-slide pattern,
            # but never put furniture on a title slide unless the truth does
            unseen_content = key not in truth_layouts and not any(
                p['type'] == 'ctrTitle' for p in L['placeholders'])
            if not (on_layout or unseen_content):
                continue
            troot = truth.xml(fe['part'])
            src = next((el for el, _ in iter_shapes(sp_tree(troot)) if shape_id(el) == fe['id']), None)
            if src is None:
                continue
            new = copy.deepcopy(src)
            cnv = new.find('.//' + P + 'cNvPr')
            cnv.set('id', str(new_id()))
            cnv.set('name', 'dglppt-fixed ' + (fe.get('name') or 'element'))
            for blip in new.iter(A + 'blip'):
                rid = blip.get(R + 'embed')
                if rid:
                    part = copy_picture_part(tgt, truth, fe['part'], rid, media_cache)
                    blip.set(R + 'embed', tgt.add_rel(sp, RT_IMAGE, part))
            for h in list(new.iter(A + 'hlinkClick')):
                h.getparent().remove(h)
            sp_tree(root).append(new)
            report['added_fixed'].append(dict(slide=n, name=fe.get('name'), shape_id=int(cnv.get('id')),
                                              box=fe.get('box')))

    tgt.save(a.out)
    meta['remap'] = report
    dump_json(meta, a.out + '.meta.json')
    if a.report:
        dump_json(report, a.report)
    print(f"placeholders remapped: {len(report['remapped'])}, unmatched: {len(report['unmatched'])}, "
          f"fixed elements added: {len(report['added_fixed'])}, removed (approved): {len(report['removed'])}")
    for u in report['unmatched']:
        print(f"  unmatched placeholder: slide {u['slide']} '{u['name']}' {u['ph']} — kept, place.py will fit it")
    for k in report['kept_footer_text']:
        print(f"  slide {k['slide']}: target footer text kept: '{k['text']}' (ask before replacing)")


if __name__ == '__main__':
    main()
