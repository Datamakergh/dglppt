#!/usr/bin/env python3
"""Step 4d-bis — copy what the author did BY HAND on the truth slides onto matching targets.

    python apply_truth_slide_design.py in.pptx truth.pptx truth_spec.json out.pptx
        [--approved approved.json] [--skip-layouts KEY ...]

Run after strip_overrides (which removed the target's own local styling). A truth deck is
often designed on its slides: the author moved or resized a title, changed a font size or
color, anchored text at the bottom, filled a box, set a background, drew a line. All of that
is design. For every target slide mapped to truth layout L, the BEST-MATCHING truth slide on
L (see score below) donates:

  placeholders (matched by type + idx; footer/date/slide number always inherit)
    position and size, fill, outline, effects, shape geometry   (never a picture fill)
    text-box settings: anchor, insets, wrap, autofit, columns   (stale autofit scale dropped)
    the shape's list style (hand-set per-level text style)
    paragraph: alignment, margins/indent, line and paragraph spacing, bullet appearance
    run: size, font, color, letter spacing, caps  (only where the target run sets none:
         kept semantic colors and the target's bold/italic/underline always win)
  slide background  (solid / gradient / theme; an image background is content: asked, not copied)
  one-off shapes    decorative ones (no text, no picture, no media) are copied; ones with text
                    or a picture are content: listed as questions, copied only when approved.json
                    -> add has {"layout": KEY, "shape_name": NAME}

Which truth slide donates: the one whose placeholder set matches the target slide best, then
whose text load per placeholder is closest, then the first. The choice is reported.

Never touched: any text, run split, shape id, picture, table, chart, media, animation, notes.
Writes meta['positions'] (boxes copied, read by place.py / check_geometry.py) and
meta['design'] (what was copied per shape, read by fit_text.py / check_geometry.py).
"""
import argparse
import copy
import math

from lxml import etree

from dglcore import (A, FIXED_PH_TYPES, P, Package, as_pptx_bytes, dump_json, iter_shapes,
                     load_json, next_shape_id, ph_of, read_xfrm, shape_id, shape_kind,
                     shape_name, sp_tree)

PARA_ATTRS = ('algn', 'marL', 'indent')
PARA_CHILDREN = ('lnSpc', 'spcBef', 'spcAft')
BULLET_LOOK = ('buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts', 'buFontTx', 'buFont')
BULLET_KIND = ('buNone', 'buAutoNum', 'buChar')
RUN_ATTRS = ('sz', 'spc', 'cap')
RUN_CHILDREN = ('solidFill', 'latin', 'ea', 'cs')
PARA_ORDER = ['lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts',
              'buFontTx', 'buFont', 'buNone', 'buAutoNum', 'buChar', 'buBlip', 'tabLst', 'defRPr', 'extLst']
RUN_ORDER = ['ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'effectLst',
             'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym',
             'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst']


def ph_key(el):
    ph = ph_of(el)
    return None if ph is None else (ph[0], str(ph[1]))


def local(el):
    return etree.QName(el).localname


def text_of(el):
    return ''.join(t.text or '' for t in el.iter(A + 't'))


def insert_ordered(parent, child, order):
    """insert child into parent respecting the schema order list."""
    rank = order.index(local(child)) if local(child) in order else len(order)
    for i, c in enumerate(parent):
        r = order.index(local(c)) if local(c) in order else len(order)
        if r > rank:
            parent.insert(i, child)
            return
    parent.append(child)


def paras(el):
    tb = el.find(P + 'txBody')
    return [] if tb is None else tb.findall(A + 'p')


def lvl_of(p):
    ppr = p.find(A + 'pPr')
    return int(ppr.get('lvl', '0')) if ppr is not None else 0


def pick_donor_para(donors, lvl, i):
    same = [d for d in donors if lvl_of(d) == lvl]
    pool = same or donors
    return pool[min(i, len(pool) - 1)]


def score(donor_phs, target_phs):
    """higher = better match: same placeholders first, then similar amount of text in each."""
    s = 0.0
    for k in target_phs:
        if k in donor_phs:
            s += 10 - min(5.0, abs(math.log1p(len(target_phs[k])) - math.log1p(len(donor_phs[k]))))
        else:
            s -= 5
    s -= 5 * len([k for k in donor_phs if k not in target_phs])
    return s


def copy_spPr(el, d_el, props):
    sp, dsp = el.find(P + 'spPr'), d_el.find(P + 'spPr')
    if sp is None or dsp is None or not len(dsp):
        return
    new = [copy.deepcopy(c) for c in dsp if local(c) != 'blipFill']
    for c in list(sp):
        sp.remove(c)
    for c in new:
        sp.append(c)
    for c in dsp:
        n = local(c)
        props.add({'xfrm': 'xfrm', 'prstGeom': 'geometry', 'custGeom': 'geometry'}.get(n, 'shape_style'))


def copy_bodyPr(el, d_el, props):
    b, db = el.find(P + 'txBody/' + A + 'bodyPr'), d_el.find(P + 'txBody/' + A + 'bodyPr')
    if b is None or db is None or (not db.attrib and not len(db)):
        return
    for k in list(b.attrib):
        del b.attrib[k]
    for c in list(b):
        b.remove(c)
    for k, v in db.attrib.items():
        b.set(k, v)
    for c in db:
        nc = copy.deepcopy(c)
        if local(nc) == 'normAutofit':  # a stale scale belongs to the donor's text, not ours
            nc.attrib.pop('fontScale', None)
            nc.attrib.pop('lnSpcReduction', None)
        b.append(nc)
    props.add('text_box')


def copy_lstStyle(el, d_el, props):
    ls, dls = el.find(P + 'txBody/' + A + 'lstStyle'), d_el.find(P + 'txBody/' + A + 'lstStyle')
    if ls is None or dls is None or not len(dls):
        return
    new = copy.deepcopy(dls)
    ls.getparent().replace(ls, new)
    props.add('list_style')


def copy_text(el, d_el, props):
    tp, dps = paras(el), [p for p in paras(d_el) if text_of(p).strip()] or paras(d_el)
    if not tp or not dps:
        return
    seen = {}
    for p in tp:
        lvl = lvl_of(p)
        i = seen.get(lvl, 0)
        seen[lvl] = i + 1
        dp = pick_donor_para(dps, lvl, i)
        dppr = dp.find(A + 'pPr')
        if dppr is not None:
            ppr = p.find(A + 'pPr')
            for k in PARA_ATTRS:
                if dppr.get(k) is not None:
                    if ppr is None:
                        ppr = etree.Element(A + 'pPr')
                        p.insert(0, ppr)
                    ppr.set(k, dppr.get(k))
                    props.add({'algn': 'align'}.get(k, k))
            for tag in PARA_CHILDREN + BULLET_LOOK:
                src = dppr.find(A + tag)
                if src is None:
                    continue
                if ppr is None:
                    ppr = etree.Element(A + 'pPr')
                    p.insert(0, ppr)
                old = ppr.find(A + tag)
                if old is not None:
                    ppr.remove(old)
                insert_ordered(ppr, copy.deepcopy(src), PARA_ORDER)
                props.add({'lnSpc': 'line_spacing', 'spcBef': 'space_before',
                           'spcAft': 'space_after'}.get(tag, 'bullet_look'))
            # bullet character / numbering style only where the target already has that kind
            for tag in ('buChar', 'buAutoNum'):
                src = dppr.find(A + tag)
                if src is not None and ppr is not None and ppr.find(A + tag) is not None:
                    ppr.replace(ppr.find(A + tag), copy.deepcopy(src))
                    props.add('bullet_look')
        drun = next((r for r in dp.findall(A + 'r') if (r.findtext(A + 't') or '').strip()), None)
        drpr = drun.find(A + 'rPr') if drun is not None else None
        if drpr is None:
            continue
        for r in p.findall(A + 'r'):
            rpr = r.find(A + 'rPr')
            if rpr is None:
                rpr = etree.Element(A + 'rPr')
                r.insert(0, rpr)
            for k in RUN_ATTRS:
                if drpr.get(k) is not None and rpr.get(k) is None:
                    rpr.set(k, drpr.get(k))
                    props.add({'sz': 'size'}.get(k, k))
            for tag in RUN_CHILDREN:
                src = drpr.find(A + tag)
                if src is None or rpr.find(A + tag) is not None:
                    continue
                if tag == 'solidFill' and any(rpr.find(A + t) is not None for t in
                                              ('gradFill', 'noFill', 'blipFill', 'pattFill')):
                    continue
                insert_ordered(rpr, copy.deepcopy(src), RUN_ORDER)
                props.add({'solidFill': 'color', 'latin': 'font', 'ea': 'font', 'cs': 'font'}[tag])


def decorative(el):
    kind = shape_kind(el)
    if local(el) in ('graphicFrame', 'pic', 'grpSp') or kind in ('picture', 'video', 'audio', 'table', 'chart'):
        return False
    return not text_of(el).strip() and not any(True for _ in el.iter(A + 'blip'))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('deck')
    ap.add_argument('truth')
    ap.add_argument('spec')
    ap.add_argument('out')
    ap.add_argument('--approved')
    ap.add_argument('--skip-layouts', nargs='*', default=[],
                    help='layout keys that should keep following the bare layout')
    a = ap.parse_args()

    tgt = Package(open(a.deck, 'rb').read())
    truth = Package(as_pptx_bytes(open(a.truth, 'rb').read()))
    spec = load_json(a.spec)
    meta = load_json(a.deck + '.meta.json')
    appr = load_json(a.approved) if a.approved else {}
    approved_oneoff = {(x.get('layout'), x.get('shape_name')) for x in appr.get('add', []) if x.get('layout')}

    donors = {}  # layout key -> [(truth slide number, part)]
    for i, (s, part) in enumerate(zip(spec['slides'], truth.slide_parts()), 1):
        donors.setdefault(s['layout'], []).append((i, part))
    fixed_truth_ids = {(fe['part'], fe['id']) for fe in spec['fixed_elements'] if fe.get('source') == 'slide'}

    pos = dict(copied=[], inherited=[], no_donor=[])
    des = dict(shapes=[], donors={}, backgrounds=[], added=[], questions=[])

    for n, part in enumerate(tgt.slide_parts(), 1):
        key = meta['layout_for_slide'][str(n)]
        if key in a.skip_layouts:
            continue
        root = tgt.xml(part)
        tph = {ph_key(el): text_of(el) for el, d in iter_shapes(sp_tree(root))
               if d == 0 and ph_key(el) and ph_key(el)[0] not in FIXED_PH_TYPES}
        cands = donors.get(key, [])
        if not cands:
            pos['no_donor'].append(dict(slide=n, layout=key))
            continue
        best = None
        for num, dpart in cands:
            dph = {ph_key(el): text_of(el) for el, d in iter_shapes(sp_tree(truth.xml(dpart)))
                   if d == 0 and ph_key(el) and ph_key(el)[0] not in FIXED_PH_TYPES}
            sc = score(dph, tph)
            if best is None or sc > best[0]:
                best = (sc, num, dpart)
        _, dnum, dpart = best
        droot = truth.xml(dpart)
        des['donors'][str(n)] = dict(truth_slide=dnum, layout=key, candidates=[c[0] for c in cands])
        dmap = {ph_key(el): el for el, d in iter_shapes(sp_tree(droot))
                if d == 0 and ph_key(el) and ph_key(el)[0] not in FIXED_PH_TYPES}

        for el, depth in iter_shapes(sp_tree(root)):
            if depth != 0:
                continue
            k = ph_key(el)
            if not k or k[0] in FIXED_PH_TYPES:
                continue
            d_el = dmap.get(k)
            if d_el is None:
                pos['inherited'].append(dict(slide=n, shape_id=shape_id(el), name=shape_name(el)))
                continue
            props = set()
            copy_spPr(el, d_el, props)
            if local(el) == 'sp':
                copy_bodyPr(el, d_el, props)
                copy_lstStyle(el, d_el, props)
                copy_text(el, d_el, props)
            if not props:
                pos['inherited'].append(dict(slide=n, shape_id=shape_id(el), name=shape_name(el)))
                continue
            if 'xfrm' in props and read_xfrm(el) is not None:
                pos['copied'].append(dict(slide=n, shape_id=shape_id(el), name=shape_name(el),
                                          box=read_xfrm(el)))
            des['shapes'].append(dict(slide=n, shape_id=shape_id(el), name=shape_name(el),
                                      truth_slide=dnum, props=sorted(props)))

        # background
        dbg, bg = droot.find(P + 'cSld/' + P + 'bg'), root.find(P + 'cSld/' + P + 'bg')
        if dbg is not None:
            has_img = any(True for _ in dbg.iter(A + 'blip'))
            tgt_img = bg is not None and any(True for _ in bg.iter(A + 'blip'))
            if has_img:
                des['questions'].append(f"Slide {n}: truth slide {dnum} has a picture background; "
                                        "it is content, so it was not copied. Add it to this slide?")
            elif not tgt_img:
                new = copy.deepcopy(dbg)
                if bg is not None:
                    bg.getparent().replace(bg, new)
                else:
                    root.find(P + 'cSld').insert(0, new)
                des['backgrounds'].append(dict(slide=n, truth_slide=dnum))

        # one-off shapes the author drew on that truth slide
        counter = next_shape_id(root)
        d_children = [c for c in sp_tree(droot) if local(c) in ('sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame')]
        first_ph = next((i for i, c in enumerate(d_children) if ph_of(c)), len(d_children))
        for i, c in enumerate(d_children):
            if ph_of(c) or (dpart, shape_id(c)) in fixed_truth_ids:
                continue
            nm = shape_name(c)
            if not (decorative(c) or (key, nm) in approved_oneoff):
                des['questions'].append(
                    f"Slide {n}: truth slide {dnum} has a one-off '{nm}' with text or a picture. It is content, "
                    "so it was not copied. Add it to this slide? "
                    f"(approved.json -> add: {{\"layout\": \"{key}\", \"shape_name\": \"{nm}\"}})")
                continue
            box = read_xfrm(c)
            if any(read_xfrm(e) == box and shape_name(e).endswith(nm) for e, dd in iter_shapes(sp_tree(root))
                   if dd == 0 and (e.find('.//' + P + 'cNvPr') is not None)):
                continue  # already there (a repeated fixed element)
            new = copy.deepcopy(c)
            cnv = new.find('.//' + P + 'cNvPr')
            cnv.set('id', str(counter))
            cnv.set('name', 'dglppt-design ' + nm)
            for kk in ('descr', 'title'):
                cnv.attrib.pop(kk, None)
            for h in list(new.iter(A + 'hlinkClick')):
                h.getparent().remove(h)
            tree = sp_tree(root)
            if i < first_ph:  # drawn behind the placeholders on the truth slide
                tree.insert(2, new)
            else:
                tree.append(new)
            des['added'].append(dict(slide=n, shape_id=counter, name=nm, box=box, truth_slide=dnum))
            counter += 1

    tgt.save(a.out)
    meta['positions'] = pos
    meta['design'] = des
    dump_json(meta, a.out + '.meta.json')
    print(f"truth-slide design copied onto {len({s['slide'] for s in des['shapes']})} slide(s): "
          f"{len(des['shapes'])} placeholder(s), {len(des['backgrounds'])} background(s), "
          f"{len(des['added'])} one-off shape(s); left to the layout: {len(pos['inherited'])}; "
          f"layouts without a truth slide: {len(pos['no_donor'])}")
    for n, d in des['donors'].items():
        print(f"  slide {n} <- truth slide {d['truth_slide']} (of {d['candidates']} on {d['layout']})")
    for s in des['shapes']:
        print(f"  slide {s['slide']} '{s['name']}': {', '.join(s['props'])}")
    for q in des['questions']:
        print('  ASK:', q)


if __name__ == '__main__':
    main()
