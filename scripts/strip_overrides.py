#!/usr/bin/env python3
"""Step 4d — remove the target's local design overrides and apply truth styles.

    python strip_overrides.py remapped.pptx truth_spec.json out.pptx
        [--roles roles.json] [--drop-semantic-color] [--restyle-boxes] [--report style.json]

Text (placeholders, text boxes, shapes, groups, table cells)
  * removed: font faces, sizes, character spacing, caps, text fills/outlines and
    effects on runs; alignment, indents, line/paragraph spacing and bullet
    char/font/color/size on paragraphs; shape-level list styles; stale autofit
    scales (fontScale / lnSpcReduction); placeholder insets and anchoring
  * kept (meaning): bold, italic, underline, strike, baseline (super/subscript),
    highlight, hyperlinks, language, paragraph level, numbered vs plain vs
    bulleted list type, tab stops
  * semantic color: a run whose color differs from the shape's dominant text
    color (red "12%") is emphasis and is kept and reported, unless
    --drop-semantic-color (only with the user's yes)
  * free text boxes don't inherit from layouts, so they get the truth role style
    written into their own list style: role "textbox" by default, or what
    roles.json says: {"3": {"4": "caption", "6": "body"}} (slide -> shape id ->
    role; roles: title, subtitle, body, textbox, caption, footnote). Their
    alignment is kept (it is arrangement, e.g. a caption centered under an icon)
Slides: solid/gradient slide backgrounds are removed (the truth layout's
  applies); picture backgrounds are content and are kept and reported.
Tables: truth table style id and flags, cell margins, font size, header/body
  cell fills and borders when the truth defines them. Cell text is untouched.
Charts: truth palette order, font sizes, legend position, gridlines. Only
  chart formatting XML changes; c:cat / c:val / the workbook never do.
Pictures and media frames: truth border, effects and corner geometry when the
  truth pictures share one style; srcRect (crop) and the image never change.
Boxes (--restyle-boxes): text-bearing autoshapes get the truth box fill/line.
"""
import argparse
import copy
from collections import Counter

from lxml import etree

from dglcore import (A, C, NS, P, R, StyleContext, TITLE_PH_TYPES, Package, c14n_str, dump_json,
                     iter_shapes, load_json, ph_of, shape_id, shape_kind, shape_name, sp_tree)

RPR_DROP_ATTRS = ('sz', 'spc', 'kern', 'cap')
RPR_DROP_CHILDREN = ('latin', 'ea', 'cs', 'sym', 'ln', 'effectLst', 'effectDag', 'uLnTx', 'uFillTx')
FILLS = ('solidFill', 'gradFill', 'noFill', 'pattFill', 'blipFill', 'grpFill')
PPR_DROP_ATTRS = ('algn', 'marL', 'marR', 'indent', 'defTabSz', 'fontAlgn')
PPR_DROP_CHILDREN = ('lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts',
                     'buFontTx', 'buFont', 'buChar', 'buBlip', 'defRPr')
BODYPR_DROP_ATTRS = ('lIns', 'tIns', 'rIns', 'bIns', 'anchor', 'anchorCtr')
# pPr child order (CT_TextParagraphProperties)
PPR_ORDER = ['lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts', 'buFontTx',
             'buFont', 'buNone', 'buAutoNum', 'buChar', 'buBlip', 'tabLst', 'defRPr', 'extLst']
RPR_ORDER = ['ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'effectLst',
             'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym',
             'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst']


def local(el):
    return etree.QName(el).localname


def insert_ordered(parent, child, order):
    pos = order.index(local(child))
    for i, c in enumerate(parent):
        if local(c) in order and order.index(local(c)) > pos:
            parent.insert(i, child)
            return
    parent.append(child)


def run_fill(rpr):
    if rpr is None:
        return None
    for c in rpr:
        if local(c) in FILLS:
            return c14n_str(c)
    return None


def strip_rpr(rpr, keep_fill):
    if rpr is None:
        return
    for k in RPR_DROP_ATTRS:
        rpr.attrib.pop(k, None)
    for c in list(rpr):
        n = local(c)
        if n in RPR_DROP_CHILDREN or (n in FILLS and not keep_fill):
            rpr.remove(c)


def strip_ppr(ppr, keep_align=False):
    if ppr is None:
        return
    for k in PPR_DROP_ATTRS:
        if k == 'algn' and keep_align:
            continue
        ppr.attrib.pop(k, None)
    for c in list(ppr):
        if local(c) in PPR_DROP_CHILDREN:
            ppr.remove(c)


def strip_txbody(txb, *, placeholder, keep_align, drop_semantic, report, where):
    """Returns list of kept semantic colors."""
    if txb is None:
        return
    lst = txb.find(A + 'lstStyle')
    if lst is not None:
        for c in list(lst):
            lst.remove(c)
    body = txb.find(A + 'bodyPr')
    if body is not None:
        for c in list(body):
            if local(c) == 'normAutofit':
                c.attrib.pop('fontScale', None)
                c.attrib.pop('lnSpcReduction', None)
        if placeholder:
            for k in BODYPR_DROP_ATTRS:
                body.attrib.pop(k, None)
            for c in list(body):
                if local(c) in ('normAutofit', 'spAutoFit', 'noAutofit'):
                    body.remove(c)
    # dominant text color by characters
    weights = Counter()
    for r in txb.iter(A + 'r'):
        weights[run_fill(r.find(A + 'rPr'))] += len(r.findtext(A + 't') or '')
    dominant = weights.most_common(1)[0][0] if weights else None
    for p in txb.findall(A + 'p'):
        strip_ppr(p.find(A + 'pPr'), keep_align)
        for r in list(p.iter(A + 'r')) + list(p.iter(A + 'fld')):
            rpr = r.find(A + 'rPr')
            f = run_fill(rpr)
            semantic = f is not None and f != dominant and not drop_semantic
            if semantic:
                report['semantic_color_kept'].append(dict(where, text=r.findtext(A + 't')))
            strip_rpr(rpr, keep_fill=semantic)
        strip_rpr(p.find(A + 'endParaRPr'), keep_fill=False)


# ---------------------------------------------------------------- role styles
def to_color_el(hexv, theme_colors, clr_map):
    """schemeClr when the hex is a theme color, else srgbClr."""
    inv = {v: k for k, v in clr_map.items()}
    for name, val in theme_colors.items():
        if val.upper() == (hexv or '').upper():
            el = etree.Element(A + 'schemeClr')
            el.set('val', inv.get(name, name))
            return el
    el = etree.Element(A + 'srgbClr')
    el.set('val', (hexv or '000000').upper())
    return el


def role_style(ctx: StyleContext, spec, role, level):
    roles = spec.get('free_text_roles', {})
    if role == 'caption':
        role = 'footnote' if 'footnote' in roles else 'textbox'
    if role in ('textbox', 'footnote') and role in roles:
        return roles[role]['para'], roles[role]['run']
    want = {'title': TITLE_PH_TYPES, 'subtitle': {'subTitle'}}.get(role, {'body', 'obj'})
    for root, cctx in ((ctx.layout, StyleContext(ctx.pkg, ctx.layout_part)),
                       (ctx.master, StyleContext(ctx.pkg, ctx.master_part))):
        for el, _ in iter_shapes(sp_tree(root), recurse=False):
            ph = ph_of(el)
            if ph and ph[0] in want:
                p = etree.Element(A + 'p')
                etree.SubElement(p, A + 'pPr').set('lvl', str(level))
                para, run = cctx.resolve_paragraph(el, p)
                return para, run
    return {}, {}


def apply_role(txb, ctx, spec, role):
    lst = txb.find(A + 'lstStyle')
    if lst is None:
        lst = etree.Element(A + 'lstStyle')
        txb.find(A + 'bodyPr').addnext(lst)
    levels = sorted({int(p.find(A + 'pPr').get('lvl', '0')) if p.find(A + 'pPr') is not None else 0
                     for p in txb.findall(A + 'p')})
    for lvl in levels:
        para, run = role_style(ctx, spec, role, lvl)
        ppr = etree.SubElement(lst, A + f'lvl{lvl + 1}pPr')
        for key, tag in (('line_spacing', 'lnSpc'), ('space_before', 'spcBef'), ('space_after', 'spcAft')):
            v = para.get(key)
            if v:
                kind, num = v.split(':')
                sp = etree.SubElement(ppr, A + tag)
                if kind == 'pct':
                    etree.SubElement(sp, A + 'spcPct').set('val', str(int(float(num) * 1000)))
                else:
                    etree.SubElement(sp, A + 'spcPts').set('val', str(int(float(num) * 100)))
        d = etree.SubElement(ppr, A + 'defRPr')
        if run.get('size'):
            d.set('sz', str(int(round(run['size'] * 100))))
        if run.get('bold') is not None:
            d.set('b', '1' if run['bold'] else '0')
        if run.get('color'):
            sf = etree.SubElement(d, A + 'solidFill')
            sf.append(to_color_el(run['color'], ctx.theme.colors if ctx.theme else {}, ctx.clr_map))
        if run.get('font'):
            lat = etree.SubElement(d, A + 'latin')
            th = ctx.theme
            lat.set('typeface', '+mn-lt' if th and run['font'] == th.minor else
                    '+mj-lt' if th and run['font'] == th.major else run['font'])


# ---------------------------------------------------------------- tables
def restyle_table(frame, spec, report, where):
    T = spec.get('tables') or {}
    tbl = frame.find('.//' + A + 'tbl')
    if tbl is None:
        return
    tblpr = tbl.find(A + 'tblPr')
    if tblpr is None:
        tblpr = etree.Element(A + 'tblPr')
        tbl.insert(0, tblpr)
    if T.get('style_id'):
        sid = tblpr.find(A + 'tableStyleId')
        if sid is None:
            sid = etree.SubElement(tblpr, A + 'tableStyleId')
        sid.text = T['style_id']
        for k in list(tblpr.attrib):
            del tblpr.attrib[k]
        for k, v in (T.get('flags') or {}).items():
            tblpr.set(k, v)
        for c in list(tblpr):
            if local(c) in FILLS or local(c) in ('effectLst', 'effectDag'):
                tblpr.remove(c)
    rows = tbl.findall(A + 'tr')
    for ri, tr in enumerate(rows):
        for tc in tr.findall(A + 'tc'):
            txb = tc.find(A + 'txBody')
            strip_txbody(txb, placeholder=False, keep_align=True, drop_semantic=False, report=report,
                         where=dict(where, cell=ri))
            if T.get('font_size') and txb is not None:
                lst = txb.find(A + 'lstStyle')
                if lst is None:
                    lst = etree.Element(A + 'lstStyle')
                    txb.find(A + 'bodyPr').addnext(lst)
                l1 = lst.find(A + 'lvl1pPr')
                if l1 is None:
                    l1 = etree.SubElement(lst, A + 'lvl1pPr')
                d = l1.find(A + 'defRPr')
                if d is None:
                    d = etree.SubElement(l1, A + 'defRPr')
                d.set('sz', str(int(T['font_size'] * 100)))
            tcpr = tc.find(A + 'tcPr')
            if T.get('style_id') and tcpr is not None:
                for c in list(tcpr):
                    if local(c) in FILLS or local(c).startswith('ln'):
                        tcpr.remove(c)
            ref = T.get('header_tcPr') if ri == 0 and T.get('flags', {}).get('firstRow') == '1' else (
                (T.get('body_tcPr') or [None, None])[(ri - 1) % 2] if ri else None)
            if ref:
                src = etree.fromstring(ref)
                if tcpr is None:
                    tcpr = etree.SubElement(tc, A + 'tcPr')
                for c in src:
                    if local(c) in FILLS or local(c).startswith('ln'):
                        tcpr.append(copy.deepcopy(c))
            if T.get('margins') and tcpr is not None:
                for k, v in zip(('marL', 'marR', 'marT', 'marB'), T['margins']):
                    if v is not None:
                        tcpr.set(k, v)
    report['tables'].append(where)


# ---------------------------------------------------------------- charts
def chart_txpr(size, typeface='+mn-lt'):
    x = (f'<c:txPr xmlns:c="{NS["c"]}" xmlns:a="{NS["a"]}"><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
         f'<a:defRPr sz="{int(size * 100)}"><a:latin typeface="{typeface}"/></a:defRPr></a:pPr>'
         f'<a:endParaRPr lang="en-US"/></a:p></c:txPr>')
    return etree.fromstring(x)


def set_txpr(parent, size, before_tags):
    if parent is None or not size:
        return
    old = parent.find(C + 'txPr')
    new = chart_txpr(size)
    if old is not None:
        old.addprevious(new)
        parent.remove(old)
        return
    for c in parent:
        if local(c) in before_tags:
            c.addprevious(new)
            return
    parent.append(new)


def restyle_chart(pkg, part, spec, report, where):
    ch = spec.get('charts') or {}
    root = pkg.xml(part)
    palette = [etree.fromstring(x) for x in ch.get('palette', [])]
    if palette:
        i = 0
        for plot in root.iter():
            if not local(plot).endswith('Chart') or plot.getparent() is None or local(plot.getparent()) != 'plotArea':
                continue
            linelike = local(plot) in ('lineChart', 'scatterChart', 'radarChart', 'line3DChart')
            vary = plot.find(C + 'varyColors')
            for ser in plot.findall(C + 'ser'):
                color = palette[i % len(palette)]
                i += 1
                sppr = ser.find(C + 'spPr')
                if sppr is None:
                    sppr = etree.Element(C + 'spPr')
                    anchor = ser.find(C + 'tx') if ser.find(C + 'tx') is not None else ser.find(C + 'order')
                    anchor.addnext(sppr)
                target = sppr
                if linelike:
                    target = sppr.find(A + 'ln')
                    if target is None:
                        target = etree.SubElement(sppr, A + 'ln')
                for c in list(target):
                    if local(c) in FILLS:
                        target.remove(c)
                sf = etree.Element(A + 'solidFill')
                sf.append(copy.deepcopy(color))
                target.insert(0, sf)
                # pies / varyColors: color each data point in palette order
                if vary is not None and vary.get('val') == '1':
                    for k, dpt in enumerate(ser.findall(C + 'dPt')):
                        dsp = dpt.find(C + 'spPr')
                        if dsp is None:
                            continue
                        for c in list(dsp):
                            if local(c) in FILLS:
                                dsp.remove(c)
                        sf = etree.Element(A + 'solidFill')
                        sf.append(copy.deepcopy(palette[k % len(palette)]))
                        dsp.insert(0, sf)
    if ch.get('chart_font_size') or ch.get('axis_font_size'):
        set_txpr(root, ch.get('chart_font_size') or ch.get('axis_font_size'),
                 ('externalData', 'printSettings', 'userShapes', 'extLst'))
    for ax in list(root.iter(C + 'catAx')) + list(root.iter(C + 'valAx')) + list(root.iter(C + 'dateAx')):
        set_txpr(ax, ch.get('axis_font_size'), ('crossAx',))
        if local(ax) == 'valAx' and 'major_gridlines' in ch:
            g = ax.find(C + 'majorGridlines')
            if ch['major_gridlines'] and g is None:
                ax.find(C + 'axPos').addnext(etree.Element(C + 'majorGridlines'))
            elif not ch['major_gridlines'] and g is not None:
                ax.remove(g)
    leg = root.find('.//' + C + 'legend')
    if leg is not None:
        set_txpr(leg, ch.get('legend_font_size'), ('extLst',))
        if ch.get('legend_pos') and leg.find(C + 'legendPos') is not None:
            leg.find(C + 'legendPos').set('val', ch['legend_pos'])
    for dl in root.iter(C + 'dLbls'):
        set_txpr(dl, ch.get('data_label_font_size'),
                 ('dLblPos', 'showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent',
                  'showBubbleSize', 'separator', 'showLeaderLines', 'leaderLines', 'extLst'))
    report['charts'].append(dict(where, part=part))


# ---------------------------------------------------------------- pictures / boxes
def restyle_spPr(el, ln_xml, effect_xml, geom_xml):
    sppr = el.find(P + 'spPr')
    if sppr is None:
        return
    order = ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill',
             'grpFill', 'ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst']
    for xml, names in ((geom_xml, ('prstGeom', 'custGeom')), (ln_xml, ('ln',)), (effect_xml, ('effectLst',))):
        if not xml:
            continue
        for c in list(sppr):
            if local(c) in names:
                sppr.remove(c)
        insert_ordered(sppr, etree.fromstring(xml), order)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('deck')
    ap.add_argument('spec')
    ap.add_argument('out')
    ap.add_argument('--roles')
    ap.add_argument('--drop-semantic-color', action='store_true')
    ap.add_argument('--restyle-boxes', action='store_true')
    ap.add_argument('--report')
    a = ap.parse_args()
    pkg = Package(open(a.deck, 'rb').read())
    spec = load_json(a.spec)
    roles = load_json(a.roles) if a.roles else {}
    meta_path = a.deck + '.meta.json'
    try:
        meta = load_json(meta_path)
    except FileNotFoundError:
        meta = {}
    fixed_ids = {(x['slide'], x['shape_id']) for x in meta.get('remap', {}).get('added_fixed', [])}
    report = dict(semantic_color_kept=[], free_text_roles=[], backgrounds_removed=[],
                  backgrounds_kept=[], tables=[], charts=[], pictures=[], boxes=[])
    pic = spec.get('pictures') or {}
    boxes = spec.get('boxes') or {}

    for n, part in enumerate(pkg.slide_parts(), 1):
        root = pkg.xml(part)
        ctx = StyleContext(pkg, part)
        bg = root.find(P + 'cSld/' + P + 'bg')
        if bg is not None:
            if bg.find('.//' + A + 'blip') is not None:
                report['backgrounds_kept'].append(dict(slide=n, why='picture background is content'))
            else:
                bg.getparent().remove(bg)
                report['backgrounds_removed'].append(n)
        for el, depth in iter_shapes(sp_tree(root)):
            sid = shape_id(el)
            if (n, sid) in fixed_ids:
                continue
            where = dict(slide=n, shape_id=sid, name=shape_name(el))
            kind = shape_kind(el)
            ph = ph_of(el) if depth == 0 else None
            txb = el.find(P + 'txBody')
            if txb is not None:
                free = ph is None
                strip_txbody(txb, placeholder=not free, keep_align=free,
                             drop_semantic=a.drop_semantic_color, report=report, where=where)
                if free and any((t.text or '').strip() for t in txb.iter(A + 't')):
                    role = roles.get(str(n), {}).get(str(sid), 'textbox')
                    apply_role(txb, ctx, spec, role)
                    report['free_text_roles'].append(dict(where, role=role))
                if kind == 'autoshape' and a.restyle_boxes and boxes.get('count'):
                    sppr = el.find(P + 'spPr')
                    for c in list(sppr):
                        if local(c) in FILLS or local(c) == 'ln':
                            sppr.remove(c)
                    order = ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill',
                             'pattFill', 'grpFill', 'ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst']
                    for xml in (boxes.get('fill'), boxes.get('ln')):
                        if xml:
                            insert_ordered(sppr, etree.fromstring(xml), order)
                    report['boxes'].append(where)
            if kind == 'table':
                restyle_table(el, spec, report, where)
            elif kind == 'chart':
                cref = el.find('.//{%s}chart' % NS['c'])
                rel = pkg.rel_target(part, cref.get(R + 'id')) if cref is not None else None
                if rel and pkg.has(rel['abs']):
                    restyle_chart(pkg, rel['abs'], spec, report, where)
            elif kind in ('picture', 'video') and pic.get('count'):
                restyle_spPr(el, pic.get('ln'), pic.get('effect'), pic.get('geom'))
                report['pictures'].append(where)

    pkg.save(a.out)
    meta['style'] = report
    dump_json(meta, a.out + '.meta.json')
    if a.report:
        dump_json(report, a.report)
    print(f"styles applied: {len(report['free_text_roles'])} free text boxes, {len(report['tables'])} tables, "
          f"{len(report['charts'])} charts, {len(report['pictures'])} pictures/media frames")
    for s in report['semantic_color_kept']:
        print(f"  semantic color kept: slide {s['slide']} '{s['name']}': {s.get('text')!r}")
    for b in report['backgrounds_kept']:
        print(f"  slide {b['slide']}: picture background kept (content) — ask the user")


if __name__ == '__main__':
    main()
