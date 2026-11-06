#!/usr/bin/env python3
"""Extract the design rules ("truth spec") of a truth deck.

    python extract_truth.py truth.pptx|truth.potx truth_spec.json [--summary summary.md]

Records, all in EMU (914400 per inch, 12700 per pt):
  canvas, theme (12 colors, heading/body fonts) per master, default text style,
  every master and layout (keyed "m<i>/<layout name>") with placeholders,
  resolved text styles per placeholder and level (inheritance-aware), borders,
  safe area, content area and gaps, fixed elements (master/layout art, footer /
  slide-number / date placeholders, logos and anything repeated on the truth
  slides), which footer placeholders the truth slides show, table, chart,
  picture and box styles, and the free-text-box roles.

Borders come from the layout placeholders first (works for .potx files with
no slides) and are cross-checked against the truth slides when there are any.
Values are clustered to a 0.01 in grid so a few EMU of drift don't count as
different values.
"""
import argparse
import os
from collections import Counter, defaultdict

from lxml import etree

from dglcore import (A, EMU_PER_IN, FIXED_PH_TYPES, NS, P, R, TITLE_PH_TYPES,
                     StyleContext, Theme, absolute_box, blip_part, c14n_str, clr_map_of,
                     cluster_value, dump_json, emu_to_cm, emu_to_in, inherited_box, iter_shapes,
                     open_package, paragraph_text, ph_of, read_xfrm, sha1,
                     shape_id, shape_kind, shape_name, sp_tree)

LEVELS = 5


def dummy_p(lvl):
    p = etree.Element(A + 'p')
    ppr = etree.SubElement(p, A + 'pPr')
    ppr.set('lvl', str(lvl))
    return p


def ph_styles(ctx, el, levels=LEVELS):
    out = []
    for lvl in range(levels):
        para, run = ctx.resolve_paragraph(el, dummy_p(lvl))
        out.append(dict(para=para, run=run))
    return out


def shape_info(pkg, part, el, inherited=None):
    box = absolute_box(el, inherited)
    kind = shape_kind(el)
    d = dict(id=shape_id(el), name=shape_name(el), kind=kind,
             box=box, top_level=el.getparent().tag == P + 'spTree')
    ph = ph_of(el)
    if ph:
        d['ph'] = list(ph)
    if kind in ('picture', 'video', 'audio'):
        img = blip_part(pkg, part, el)
        if img:
            d['image_sha1'] = sha1(pkg.blob(img))
        src = el.find('.//' + A + 'srcRect')
        d['crop'] = dict(src.attrib) if src is not None else {}
    tb = el.find(P + 'txBody')
    if tb is not None:
        txt = '\n'.join(paragraph_text(p) for p in tb.findall(A + 'p')).strip()
        if txt:
            d['text'] = txt
    sppr = el.find(P + 'spPr')
    if sppr is not None:
        d['style_sha1'] = sha1(c14n_str(sppr).encode())[:12]
    return d


def is_background_like(box, W, H):
    """Full-bleed art or edge-to-edge bands: design, never content."""
    if not box:
        return True
    x, y, w, h = box
    spans_w = x <= 0.01 * W and x + w >= 0.99 * W
    spans_h = y <= 0.01 * H and y + h >= 0.99 * H
    return spans_w or spans_h


def borders_from_boxes(boxes, W, H):
    if not boxes:
        return None
    return dict(left=min(b[0] for b in boxes), top=min(b[1] for b in boxes),
                right=W - max(b[0] + b[2] for b in boxes),
                bottom=H - max(b[1] + b[3] for b in boxes))


def gaps(boxes):
    v, h = [], []
    for i, a in enumerate(boxes):
        for b in boxes[i + 1:]:
            # vertical stack: x ranges overlap
            if min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0]) > 0:
                top, bot = (a, b) if a[1] <= b[1] else (b, a)
                g = bot[1] - (top[1] + top[3])
                if g > 0:
                    v.append(g)
            # side by side: y ranges overlap
            if min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]) > 0:
                lft, rgt = (a, b) if a[0] <= b[0] else (b, a)
                g = rgt[0] - (lft[0] + lft[2])
                if g > 0:
                    h.append(g)
    return v, h


def nearest_neighbor_gaps(boxes):
    """Only gaps between neighbours (no shape in between) — avoids long-range noise."""
    v, h = gaps(boxes)
    return v, h


def signature(d):
    b = d.get('box') or [0, 0, 0, 0]
    rb = tuple(int(round(x / (EMU_PER_IN / 100))) for x in b)
    return (d['kind'], rb, d.get('image_sha1') or d.get('text') or d.get('style_sha1'))


def resolved_run_styles(ctx, el, in_table=False):
    out = []
    tb = el.find(P + 'txBody')
    if tb is None:
        return out
    for p in tb.findall(A + 'p'):
        for r in p.findall(A + 'r'):
            if (r.findtext(A + 't') or '').strip():
                para, run = ctx.resolve_run(el, p, r, in_table)
                out.append((para, run))
    return out


def style_key(run):
    return (run.get('font'), run.get('size'), run.get('color'), run.get('bold'))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('truth')
    ap.add_argument('out')
    ap.add_argument('--summary')
    a = ap.parse_args()

    raw = open(a.truth, 'rb').read()
    pkg = open_package(a.truth)
    W, H = pkg.slide_size()
    pres = pkg.xml(pkg.main_part)
    spec = dict(source=dict(file=os.path.basename(a.truth), sha1=sha1(raw)),
                canvas=dict(w=W, h=H), masters=[], layouts={}, fixed_elements=[],
                slides=[], warnings=[])
    dts = pres.find(P + 'defaultTextStyle')
    spec['default_text_style_sha1'] = sha1(c14n_str(dts).encode())[:12] if dts is not None else None

    layout_key_of = {}
    for mi, mpart in enumerate(pkg.master_parts(), 1):
        mroot = pkg.xml(mpart)
        tpart = pkg.theme_of(mpart)
        theme = Theme(pkg.xml(tpart)) if tpart else None
        mctx = StyleContext(pkg, mpart)
        tx = mroot.find(P + 'txStyles')
        txs = {}
        if tx is not None:
            for nm in ('titleStyle', 'bodyStyle', 'otherStyle'):
                lst = tx.find(P + nm)
                levels = []
                for lvl in range(LEVELS):
                    para, run = {}, {}
                    from dglcore import _read_ppr, _read_rpr, _lvl_ppr
                    ppr = _lvl_ppr(lst, lvl)
                    _read_ppr(ppr, para)
                    if ppr is not None:
                        _read_rpr(ppr.find(A + 'defRPr'), theme, mctx.clr_map, run)
                    levels.append(dict(para=para, run=run))
                txs[nm] = levels
        hf = mroot.find(P + 'hf')
        mshapes = []
        for el, depth in iter_shapes(sp_tree(mroot)):
            if depth:
                continue
            mshapes.append(shape_info(pkg, mpart, el))
        mkey = f'm{mi}'
        spec['masters'].append(dict(
            key=mkey, part=mpart, theme=dict(name=theme.name if theme else None,
                                             colors=theme.colors if theme else {},
                                             heading_font=theme.major if theme else None,
                                             body_font=theme.minor if theme else None),
            clr_map=clr_map_of(mroot), txStyles=txs,
            hf=dict(hf.attrib) if hf is not None else None,
            has_background=mroot.find(P + 'cSld/' + P + 'bg') is not None,
            shapes=mshapes))
        # fixed elements on the master
        for d in mshapes:
            if d.get('ph'):
                if d['ph'][0] in FIXED_PH_TYPES:
                    spec['fixed_elements'].append(dict(source='master', part=mpart, master=mkey, **d))
            elif not (d['box'] and d['box'][2] >= 0.98 * W and d['box'][3] >= 0.98 * H):
                spec['fixed_elements'].append(dict(source='master', part=mpart, master=mkey, **d))

        seen_names = Counter()
        for lpart in pkg.layout_parts(mpart):
            lroot = pkg.xml(lpart)
            cs = lroot.find(P + 'cSld')
            lname = cs.get('name') or os.path.basename(lpart)
            seen_names[lname] += 1
            key = f'{mkey}/{lname}' + (f'#{seen_names[lname]}' if seen_names[lname] > 1 else '')
            layout_key_of[lpart] = key
            ctx = StyleContext(pkg, lpart)
            phs, fixed, content_boxes = [], [], []
            for el, depth in iter_shapes(sp_tree(lroot)):
                if depth:
                    continue
                ph = ph_of(el)
                if ph:
                    box = read_xfrm(el)
                    if box is None:
                        from dglcore import match_master_ph
                        mel = match_master_ph(mroot, ph)
                        box = read_xfrm(mel) if mel is not None else None
                    info = dict(type=ph[0], idx=ph[1], name=shape_name(el), box=box,
                                styles=ph_styles(ctx, el) if el.find(P + 'txBody') is not None
                                or ph[0] not in ('pic', 'chart', 'tbl', 'media', 'clipArt', 'dgm')
                                else None)
                    phs.append(info)
                    if ph[0] in FIXED_PH_TYPES:
                        fixed.append(dict(source='layout', part=lpart, layout=key,
                                          **shape_info(pkg, lpart, el, box)))
                    elif box:
                        content_boxes.append(box)
                else:
                    d = shape_info(pkg, lpart, el)
                    if not (d['box'] and d['box'][2] >= 0.98 * W and d['box'][3] >= 0.98 * H):
                        fixed.append(dict(source='layout', part=lpart, layout=key, **d))
            spec['fixed_elements'].extend(fixed)
            if not any(p['type'] not in TITLE_PH_TYPES | FIXED_PH_TYPES for p in phs):
                # title-only / blank layouts: the master body defines the content zone
                for mel_box in master_body_boxes(mroot):
                    content_boxes.append(mel_box)
                if not any(p['type'] in TITLE_PH_TYPES for p in phs):
                    content_boxes.extend(master_title_boxes(mroot))
            b = borders_from_boxes(content_boxes, W, H)
            title = next((p for p in phs if p['type'] in TITLE_PH_TYPES and p['box']), None)
            bodies = [p for p in phs if p['type'] not in TITLE_PH_TYPES | FIXED_PH_TYPES and p['box']
                      and p['type'] != 'subTitle']
            vg, hg = gaps(content_boxes)
            content_area = None
            if bodies:
                bb = [p['box'] for p in bodies]
                x0, y0 = min(x[0] for x in bb), min(x[1] for x in bb)
                x1, y1 = max(x[0] + x[2] for x in bb), max(x[1] + x[3] for x in bb)
                content_area = [x0, y0, x1 - x0, y1 - y0]
            elif not bodies and master_body_boxes(mroot):
                content_area = master_body_boxes(mroot)[0]
            elif b and title:
                g = cluster_value(vg) or int(0.2 * EMU_PER_IN)
                top = title['box'][1] + title['box'][3] + g
                content_area = [b['left'], top, W - b['left'] - b['right'], H - b['bottom'] - top]
            elif b:
                content_area = [b['left'], b['top'], W - b['left'] - b['right'], H - b['top'] - b['bottom']]
            spec['layouts'][key] = dict(
                key=key, master=mkey, part=lpart, name=lname,
                show_master_shapes=lroot.get('showMasterSp', '1') != '0',
                placeholders=phs,
                borders=dict(b or {}, source='layout'),
                gaps=dict(vgap=cluster_value(vg), hgap=cluster_value(hg), source='layout'),
                content_area=content_area,
                title_box=title['box'] if title else None,
                has_background=lroot.find(P + 'cSld/' + P + 'bg') is not None)

    # ------------------------------------------------------------------ slides
    slide_parts = pkg.slide_parts()
    shown = defaultdict(set)
    per_layout_boxes = defaultdict(lambda: dict(b=[], v=[], h=[]))
    sig_count = Counter()
    sig_layouts = defaultdict(set)
    slide_infos = []
    for n, spart in enumerate(slide_parts, 1):
        root = pkg.xml(spart)
        lkey = layout_key_of.get(pkg.layout_of(spart))
        shapes = []
        for el, depth in iter_shapes(sp_tree(root)):
            if depth:
                continue
            d = shape_info(pkg, spart, el, inherited_box(pkg, spart, el))
            shapes.append(d)
            if d.get('ph') and d['ph'][0] in FIXED_PH_TYPES:
                shown[lkey].add(d['ph'][0])
        slide_infos.append(dict(n=n, part=spart, layout=lkey, shapes=shapes))
        for d in shapes:
            if not d.get('ph'):
                s = signature(d)
                sig_count[s] += 1
                sig_layouts[s].add(lkey)

    n_slides = len(slide_infos)
    for sl in slide_infos:
        for d in sl['shapes']:
            fixed = False
            if d.get('ph'):
                fixed = d['ph'][0] in FIXED_PH_TYPES
            else:
                s = signature(d)
                box = d.get('box')
                in_band = box is not None and (box[1] + box[3] <= 0.15 * H or box[1] >= 0.85 * H)
                if sig_count[s] > 1 and n_slides > 1:
                    fixed = True
                elif n_slides == 1 and in_band and d['kind'] in ('picture', 'textbox', 'autoshape'):
                    fixed = True
                    spec['warnings'].append(
                        f"single-slide truth: '{d['name']}' in the header/footer band treated as fixed; confirm")
            d['fixed'] = fixed
            if fixed and not d.get('ph'):
                s = signature(d)
                if not any(fe.get('signature') == list(map(str, s)) for fe in spec['fixed_elements']):
                    spec['fixed_elements'].append(dict(
                        source='slide', part=sl['part'], slide=sl['n'], layouts=sorted(
                            k for k in sig_layouts[s] if k), signature=list(map(str, s)), **d))
        content = [d['box'] for d in sl['shapes']
                   if not d['fixed'] and d.get('box') and not is_background_like(d['box'], W, H)]
        if content and sl['layout']:
            pl = per_layout_boxes[sl['layout']]
            pl['b'].append(borders_from_boxes(content, W, H))
            v, h = gaps(content)
            pl['v'].extend(v)
            pl['h'].extend(h)
        spec['slides'].append(dict(n=sl['n'], layout=sl['layout'],
                                   shapes=[{k: v for k, v in d.items() if k != 'style_sha1'}
                                           for d in sl['shapes']]))

    # cross-check borders with slides: content must never sit closer to an edge
    # than the truth slides themselves put it
    for key, L in spec['layouts'].items():
        pl = per_layout_boxes.get(key)
        if not pl or not pl['b']:
            if not L['borders'].get('left') and L['borders'].get('left') != 0:
                spec['warnings'].append(f'layout {key}: no placeholders and no slides — borders unknown')
            continue
        obs = {side: cluster_value([b[side] for b in pl['b']]) for side in ('left', 'top', 'right', 'bottom')}
        lb = L['borders']
        merged = {}
        for side, v in obs.items():
            if lb.get(side) is None:
                merged[side] = v
            else:
                merged[side] = min(lb[side], v)
        L['borders'] = dict(merged, source='layout+slides', layout_only={k: lb.get(k) for k in obs},
                            slides_observed=obs)
        if pl['v']:
            L['gaps']['vgap'] = cluster_value(pl['v'])
            L['gaps']['source'] = 'slides'
        if pl['h']:
            L['gaps']['hgap'] = cluster_value(pl['h'])
            L['gaps']['source'] = 'slides'
    for key, L in spec['layouts'].items():
        b = L['borders']
        if all(b.get(s) is not None for s in ('left', 'top', 'right', 'bottom')):
            L['safe_area'] = [b['left'], b['top'], W - b['left'] - b['right'], H - b['top'] - b['bottom']]
        else:
            L['safe_area'] = None

    # footer / date / slide-number text as the truth slides show it (copied onto targets)
    fixed_tx = defaultdict(Counter)
    for spart in slide_parts:
        for el, depth in iter_shapes(sp_tree(pkg.xml(spart)), recurse=False):
            ph = ph_of(el)
            if ph and ph[0] in FIXED_PH_TYPES and el.find(P + 'txBody') is not None:
                fixed_tx[ph[0]][etree.tostring(el.find(P + 'txBody'), encoding='unicode')] += 1
    spec['fixed_placeholder_txbody'] = {t: c.most_common(1)[0][0] for t, c in fixed_tx.items()}

    # which footer placeholders the truth deck shows
    if n_slides:
        union = set().union(*shown.values()) if shown else set()
        for key, L in spec['layouts'].items():
            types_on_layout = {p['type'] for p in L['placeholders'] if p['type'] in FIXED_PH_TYPES}
            used_by_truth = any(sl['layout'] == key for sl in slide_infos)
            seen = shown.get(key, set()) if used_by_truth else union
            L['show_fixed_placeholders'] = sorted(seen & types_on_layout)
    else:
        for key, L in spec['layouts'].items():
            L['show_fixed_placeholders'] = sorted(p['type'] for p in L['placeholders']
                                                  if p['type'] in FIXED_PH_TYPES)
        spec['warnings'].append('truth file has no slides: footer/slide number/date assumed shown '
                                'wherever the layout defines them; confirm with the user')

    # ---------------------------------------------------- tables / charts / pics
    tables, charts, pics, boxes, textbox_runs = [], [], [], [], []
    for sl, spart in zip(slide_infos, slide_parts):
        root = pkg.xml(spart)
        ctx = StyleContext(pkg, spart)
        for el, depth in iter_shapes(sp_tree(root)):
            kind = shape_kind(el)
            if kind == 'table':
                tbl = el.find('.//' + A + 'tbl')
                tblpr = tbl.find(A + 'tblPr')
                sid = tblpr.findtext(A + 'tableStyleId') if tblpr is not None else None
                rows = tbl.findall(A + 'tr')
                sizes, margins = [], []
                for tr in rows:
                    for tc in tr.findall(A + 'tc'):
                        tcpr = tc.find(A + 'tcPr')
                        if tcpr is not None:
                            margins.append(tuple(tcpr.get(k) for k in ('marL', 'marR', 'marT', 'marB')))
                        for r in tc.iter(A + 'r'):
                            rp = r.find(A + 'rPr')
                            if rp is not None and rp.get('sz'):
                                sizes.append(int(rp.get('sz')))
                def tcpr_of(tr):
                    tc = tr.find(A + 'tc') if tr is not None else None
                    t = tc.find(A + 'tcPr') if tc is not None else None
                    return etree.tostring(t, encoding='unicode') if t is not None else None
                tables.append(dict(slide=sl['n'], style_id=sid,
                                   flags={k: v for k, v in (tblpr.attrib.items() if tblpr is not None else [])},
                                   header_tcPr=tcpr_of(rows[0] if rows else None),
                                   body_tcPr=[tcpr_of(rows[1]) if len(rows) > 1 else None,
                                              tcpr_of(rows[2]) if len(rows) > 2 else None],
                                   font_size=Counter(sizes).most_common(1)[0][0] / 100 if sizes else None,
                                   margins=Counter(margins).most_common(1)[0][0] if margins else None))
            elif kind == 'chart':
                cref = el.find('.//{%s}chart' % NS['c'])
                rel = pkg.rel_target(spart, cref.get(R + 'id')) if cref is not None else None
                if rel and pkg.has(rel['abs']):
                    charts.append(chart_style(pkg.xml(rel['abs'])))
            elif kind == 'picture' and not any(fe.get('part') == spart and fe.get('id') == shape_id(el)
                                               for fe in spec['fixed_elements']):
                sppr = el.find(P + 'spPr')
                pics.append(dict(
                    ln=etree.tostring(sppr.find(A + 'ln'), encoding='unicode') if sppr is not None and sppr.find(A + 'ln') is not None else None,
                    effect=etree.tostring(sppr.find(A + 'effectLst'), encoding='unicode') if sppr is not None and sppr.find(A + 'effectLst') is not None else None,
                    geom=etree.tostring(sppr.find(A + 'prstGeom'), encoding='unicode') if sppr is not None and sppr.find(A + 'prstGeom') is not None else None))
            elif kind == 'autoshape' and el.find(P + 'txBody') is not None and paragraph_text_any(el):
                sppr = el.find(P + 'spPr')
                fill = next((c for c in sppr if etree.QName(c).localname in
                             ('solidFill', 'gradFill', 'noFill', 'pattFill')), None) if sppr is not None else None
                ln = sppr.find(A + 'ln') if sppr is not None else None
                boxes.append((etree.tostring(fill, encoding='unicode') if fill is not None else None,
                              etree.tostring(ln, encoding='unicode') if ln is not None else None))
            elif kind == 'textbox' and depth == 0:
                is_fixed = any(fe.get('part') == spart and fe.get('id') == shape_id(el)
                               for fe in spec['fixed_elements'])
                if not is_fixed:
                    box = absolute_box(el)
                    for para, run in resolved_run_styles(ctx, el):
                        textbox_runs.append(dict(run=run, para=para, bottom_band=bool(box and box[1] > 0.8 * H)))

    def most_common(items):
        c = Counter(items)
        return c.most_common(1)[0][0] if c else None

    spec['tables'] = dict(count=len(tables))
    if tables:
        sid = most_common([t['style_id'] for t in tables])
        ref = next(t for t in tables if t['style_id'] == sid)
        spec['tables'].update(style_id=sid, flags=ref['flags'], header_tcPr=ref['header_tcPr'],
                              body_tcPr=ref['body_tcPr'],
                              font_size=most_common([t['font_size'] for t in tables if t['font_size']]),
                              margins=most_common([t['margins'] for t in tables if t['margins']]))
    else:
        spec['warnings'].append('truth slides contain no table: table styling must be confirmed with the user')
    spec['charts'] = dict(count=len(charts))
    if charts:
        spec['charts'].update(charts[0])
    else:
        spec['charts']['palette'] = [f'<a:schemeClr xmlns:a="{NS["a"]}" val="accent{i}"/>' for i in range(1, 7)]
        spec['charts']['palette_source'] = 'theme accents (truth has no chart)'
    spec['pictures'] = dict(count=len(pics))
    if pics:
        spec['pictures'].update(ln=most_common([p['ln'] for p in pics]),
                                effect=most_common([p['effect'] for p in pics]),
                                geom=most_common([p['geom'] for p in pics]))
    spec['boxes'] = dict(count=len(boxes))
    if boxes:
        f, l = most_common(boxes)
        spec['boxes'].update(fill=f, ln=l)
    roles = {}
    if textbox_runs:
        body = [t for t in textbox_runs if not t['bottom_band']]
        if body:
            k = most_common([style_key(t['run']) for t in body])
            roles['textbox'] = next(t for t in body if style_key(t['run']) == k)
        foot = [t for t in textbox_runs if t['bottom_band']]
        if foot:
            k = most_common([style_key(t['run']) for t in foot])
            roles['footnote'] = next(t for t in foot if style_key(t['run']) == k)
    spec['free_text_roles'] = roles

    dump_json(spec, a.out)
    print('truth spec written:', a.out)
    if spec['warnings']:
        print('warnings:')
        for w in spec['warnings']:
            print('  -', w)
    if a.summary:
        write_summary(spec, a.summary)
        print('summary written:', a.summary)


def master_body_boxes(mroot):
    return [read_xfrm(el) for el, _ in iter_shapes(sp_tree(mroot), recurse=False)
            if ph_of(el) and ph_of(el)[0] == 'body' and read_xfrm(el)]


def master_title_boxes(mroot):
    return [read_xfrm(el) for el, _ in iter_shapes(sp_tree(mroot), recurse=False)
            if ph_of(el) and ph_of(el)[0] in TITLE_PH_TYPES and read_xfrm(el)]


def paragraph_text_any(el):
    tb = el.find(P + 'txBody')
    return any(paragraph_text(p).strip() for p in tb.findall(A + 'p')) if tb is not None else False


def chart_style(root):
    Cn = '{%s}' % NS['c']
    palette = []
    for ser in root.iter(Cn + 'ser'):
        sppr = ser.find(Cn + 'spPr')
        clr = None
        if sppr is not None:
            sf = sppr.find(A + 'solidFill')
            if sf is None and sppr.find(A + 'ln/' + A + 'solidFill') is not None:
                sf = sppr.find(A + 'ln/' + A + 'solidFill')
            if sf is not None and len(sf):
                clr = etree.tostring(sf[0], encoding='unicode')
        if clr:
            palette.append(clr)
        for dpt in ser.findall(Cn + 'dPt'):
            sf = dpt.find(Cn + 'spPr/' + A + 'solidFill')
            if sf is not None and len(sf):
                palette.append(etree.tostring(sf[0], encoding='unicode'))

    def size_of(path):
        el = root.find(path)
        if el is None:
            return None
        d = el.find('.//' + A + 'defRPr')
        return int(d.get('sz')) / 100 if d is not None and d.get('sz') else None

    out = dict(palette=palette or [f'<a:schemeClr xmlns:a="{NS["a"]}" val="accent{i}"/>' for i in range(1, 7)],
               axis_font_size=size_of('.//' + Cn + 'catAx/' + Cn + 'txPr') or size_of('.//' + Cn + 'valAx/' + Cn + 'txPr'),
               legend_font_size=size_of('.//' + Cn + 'legend/' + Cn + 'txPr'),
               data_label_font_size=size_of('.//' + Cn + 'dLbls/' + Cn + 'txPr'),
               chart_font_size=size_of(Cn + 'txPr'),
               legend_pos=(root.find('.//' + Cn + 'legend/' + Cn + 'legendPos').get('val')
                           if root.find('.//' + Cn + 'legend/' + Cn + 'legendPos') is not None else None),
               major_gridlines=root.find('.//' + Cn + 'valAx/' + Cn + 'majorGridlines') is not None)
    lat = root.find('.//' + Cn + 'txPr//' + A + 'latin')
    out['font'] = lat.get('typeface') if lat is not None else None
    return out


def fmt_len(v):
    return '—' if v is None else f'{emu_to_in(v)} in / {emu_to_cm(v)} cm'


def write_summary(spec, path):
    W, H = spec['canvas']['w'], spec['canvas']['h']
    L = ['# Truth spec summary', '',
         f"**File:** {spec['source']['file']}  ",
         f"**Slide size:** {emu_to_in(W)} × {emu_to_in(H)} in ({emu_to_cm(W)} × {emu_to_cm(H)} cm)", '']
    for m in spec['masters']:
        t = m['theme']
        L.append(f"## Master {m['key']} — theme '{t['name']}'")
        L.append(f"- Heading font: **{t['heading_font']}**, body font: **{t['body_font']}**")
        L.append('- Colors: ' + ', '.join(f"{k} `#{v}`" for k, v in t['colors'].items()))
        for nm, levels in m['txStyles'].items():
            r = levels[0]['run']
            if r:
                L.append(f"- {nm} level 1: {r.get('font', 'theme')} {r.get('size', '?')} pt, color #{r.get('color', '?')}")
        L.append('')
    L.append('## Layouts')
    L.append('| Layout | Placeholders | Borders L / T / R / B (in) | Gap v / h (in) | Shows |')
    L.append('|---|---|---|---|---|')
    for key, lay in spec['layouts'].items():
        b = lay['borders']
        g = lay['gaps']
        phs = ', '.join(f"{p['type']}#{p['idx']}" for p in lay['placeholders'])
        L.append(f"| {key} | {phs} | {emu_to_in(b.get('left'))} / {emu_to_in(b.get('top'))} / "
                 f"{emu_to_in(b.get('right'))} / {emu_to_in(b.get('bottom'))} | "
                 f"{emu_to_in(g.get('vgap'))} / {emu_to_in(g.get('hgap'))} | "
                 f"{', '.join(lay.get('show_fixed_placeholders', []))} |")
    L.append('')
    L.append('## Fixed elements (header, footer, logo, slide number)')
    grouped = {}
    for fe in spec['fixed_elements']:
        k = (fe['kind'], (fe.get('ph') or [None])[0], tuple(fe.get('box') or []), fe.get('text'))
        where = fe.get('layout') or fe.get('master') or f"slides {fe.get('layouts')}"
        grouped.setdefault(k, dict(fe=fe, where=[]))['where'].append(f"{fe['source']} {where}")
    for (kind, ph, box, text), g in grouped.items():
        b = list(box) or [None] * 4
        name = ph or g['fe']['name']
        where = g['where'][0] + (f" (+{len(g['where']) - 1} more)" if len(g['where']) > 1 else '')
        L.append(f"- {kind} '{name}' — {where}: x {fmt_len(b[0])}, y {fmt_len(b[1])}, "
                 f"w {fmt_len(b[2])}, h {fmt_len(b[3])}" + (f" — text “{text}”" if text else ''))
    L.append('')
    t = spec['tables']
    L.append(f"## Tables: {t['count']} in truth" + (f" — style `{t.get('style_id')}`, text {t.get('font_size')} pt" if t['count'] else ' — **not defined**'))
    c = spec['charts']
    L.append(f"## Charts: {c['count']} in truth — palette from {c.get('palette_source', 'truth charts')}")
    if spec['free_text_roles']:
        for role, s in spec['free_text_roles'].items():
            L.append(f"- free text role '{role}': {s['run'].get('font')} {s['run'].get('size')} pt #{s['run'].get('color')}")
    if spec['warnings']:
        L += ['', '## Needs confirmation'] + [f'- {w}' for w in spec['warnings']]
    open(path, 'w', encoding='utf-8').write('\n'.join(L) + '\n')


if __name__ == '__main__':
    main()
