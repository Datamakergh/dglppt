#!/usr/bin/env python3
"""Step 4e — detect text that doesn't fit and shrink it gently.

    python fit_text.py deck.pptx out.pptx [--floor 0.85] [--min-body 10] [--min-small 8]
                       [--allow SLIDE:SHAPE ...] [--report fit.json]

Measures every text shape with the real font metrics (Pillow + fc-match):
words are wrapped to the box width minus insets and indents, line height is
size x 1.2 x line spacing, plus space before/after. Then, in order, stopping
at the first step that fits:
  1. shrink that shape's text 1 pt at a time down to --floor (85%) of its
     truth size, never below --min-body (10 pt; --min-small 8 pt for text
     whose truth size is under 12 pt: footnotes, captions, table text)
  2. tighten line spacing one step (100% -> 90%)
  3. still too big -> "needs decision": split the slide, allow a smaller size
     this once, or the user trims the text. Nothing is cut or moved.
Every shrink is written explicitly on the runs and listed in the report.
If a font file isn't installed the estimate uses fc-match's substitute and
says so: re-check those slides by eye.
"""
import argparse
import functools
import subprocess

from lxml import etree

from dglcore import (A, EMU_PER_PT, P, Package, StyleContext, dump_json, inherited_box, iter_shapes,
                     load_json, ph_of, shape_id, shape_name, sp_tree)

DEFAULT_INS = dict(lIns=91440, rIns=91440, tIns=45720, bIns=45720)


@functools.lru_cache(maxsize=None)
def font_file(family, bold):
    pattern = f'{family}:bold' if bold else family
    try:
        out = subprocess.run(['fc-match', '-f', '%{file}|%{family}', pattern], capture_output=True,
                             text=True, timeout=10).stdout
        path, fam = out.split('|', 1)
        return path, family.lower() in fam.lower()
    except Exception:
        return None, False


@functools.lru_cache(maxsize=None)
def pil_font(family, bold, size_pt):
    from PIL import ImageFont
    path, exact = font_file(family or 'Liberation Sans', bold)
    try:
        # 1 pt = 1 unit; measure at 10x for precision
        return ImageFont.truetype(path, int(round(size_pt * 10))), exact
    except Exception:
        return ImageFont.load_default(), False


def text_width_pt(text, family, bold, size):
    f, exact = pil_font(family, bold, size)
    try:
        return f.getlength(text) / 10, exact
    except Exception:
        return len(text) * size * 0.5, False


def bodypr_insets(pkg, part, el):
    out = dict(DEFAULT_INS)
    chain = [el.find(P + 'txBody/' + A + 'bodyPr')]
    ph = ph_of(el)
    if ph:
        from dglcore import match_layout_ph, match_master_ph
        lay = pkg.layout_of(part)
        lel = match_layout_ph(pkg.xml(lay), ph) if lay else None
        if lel is not None:
            chain.append(lel.find(P + 'txBody/' + A + 'bodyPr'))
        mel = match_master_ph(pkg.xml(pkg.master_of(lay)), ph) if lay else None
        if mel is not None:
            chain.append(mel.find(P + 'txBody/' + A + 'bodyPr'))
    found = {}
    for b in chain:
        if b is None:
            continue
        for k in DEFAULT_INS:
            if k not in found and b.get(k) is not None:
                found[k] = int(b.get(k))
    out.update(found)
    return out


def spacing_pt(v, size):
    if not v:
        return 0.0
    kind, num = v.split(':')
    return float(num) / 100 * size * 1.2 if kind == 'pct' else float(num)


def line_factor(v):
    if not v:
        return 1.0
    kind, num = v.split(':')
    return float(num) / 100 if kind == 'pct' else None


def measure(ctx, el, box, insets, scale=1.0, lnspc_override=None):
    """Height in pt the text needs, available height in pt, words that can't fit a line."""
    avail_w = (box[2] - insets['lIns'] - insets['rIns']) / EMU_PER_PT
    avail_h = (box[3] - insets['tIns'] - insets['bIns']) / EMU_PER_PT
    total = 0.0
    broken, approx = [], False
    for i, p in enumerate(el.find(P + 'txBody').findall(A + 'p')):
        para, base = ctx.resolve_paragraph(el, p)
        runs = p.findall(A + 'r')
        size = base['size']
        if runs:
            _, r0 = ctx.resolve_run(el, p, runs[0])
            size = r0['size']
        size *= scale
        indent = ((para.get('marL') or 0) + max(0, para.get('indent') or 0)) / EMU_PER_PT
        width = max(10.0, avail_w - indent)
        lines, cur = 1, 0.0
        text = ''.join((r.findtext(A + 't') or '') for r in p if etree.QName(r).localname in ('r', 'fld'))
        family, bold = base.get('font'), base.get('bold')
        if runs:
            _, r0 = ctx.resolve_run(el, p, runs[0])
            family, bold = r0.get('font'), r0.get('bold')
        space_w, exact = text_width_pt(' ', family, bold, size)
        approx = approx or not exact
        for w in text.split():
            ww, _ = text_width_pt(w, family, bold, size)
            if ww > width:
                broken.append(w)
            if cur == 0:
                cur = ww
            elif cur + space_w + ww <= width:
                cur += space_w + ww
            else:
                lines += 1
                cur = ww
        lf = lnspc_override if lnspc_override else line_factor(para.get('line_spacing'))
        lh = size * 1.2 * lf if lf else float(para['line_spacing'].split(':')[1])
        total += lines * lh
        if i:
            total += spacing_pt(para.get('space_before'), size)
        total += spacing_pt(para.get('space_after'), size)
    return total, avail_h, broken, approx


def set_size(el, size_pt):
    for r in el.find(P + 'txBody').iter(A + 'r', A + 'fld'):
        rpr = r.find(A + 'rPr')
        if rpr is None:
            rpr = etree.Element(A + 'rPr')
            r.insert(0, rpr)
        rpr.set('sz', str(int(round(size_pt * 100))))


def set_lnspc(el, pct):
    for p in el.find(P + 'txBody').findall(A + 'p'):
        ppr = p.find(A + 'pPr')
        if ppr is None:
            ppr = etree.Element(A + 'pPr')
            p.insert(0, ppr)
        old = ppr.find(A + 'lnSpc')
        if old is not None:
            ppr.remove(old)
        ln = etree.Element(A + 'lnSpc')
        etree.SubElement(ln, A + 'spcPct').set('val', str(int(pct * 1000)))
        ppr.insert(0, ln)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('deck')
    ap.add_argument('out')
    ap.add_argument('--floor', type=float, default=0.85)
    ap.add_argument('--min-body', type=float, default=10)
    ap.add_argument('--min-small', type=float, default=8)
    ap.add_argument('--allow', nargs='*', default=[], help='SLIDE:SHAPE pairs the user allowed below the floor')
    ap.add_argument('--report')
    a = ap.parse_args()
    pkg = Package(open(a.deck, 'rb').read())
    try:
        meta = load_json(a.deck + '.meta.json')
    except FileNotFoundError:
        meta = {}
    fixed_ids = {(x['slide'], x['shape_id']) for x in meta.get('remap', {}).get('added_fixed', [])}
    allowed = set(a.allow)
    report = dict(shrinks=[], needs_decision=[], broken_words=[], approximate_fonts=set(), checked=0)

    for n, part in enumerate(pkg.slide_parts(), 1):
        ctx = StyleContext(pkg, part)
        for el, depth in iter_shapes(sp_tree(pkg.xml(part))):
            if el.find(P + 'txBody') is None or (n, shape_id(el)) in fixed_ids:
                continue
            ph = ph_of(el) if depth == 0 else None
            if ph and ph[0] in ('ftr', 'sldNum', 'dt'):
                continue
            if not ''.join(t.text or '' for t in el.iter(A + 't')).strip():
                continue
            box = inherited_box(pkg, part, el) if depth == 0 else None
            if depth:
                from dglcore import absolute_box
                box = absolute_box(el)
            if not box:
                continue
            body = el.find(P + 'txBody/' + A + 'bodyPr')
            if body is not None and body.find(A + 'spAutoFit') is not None:
                continue  # box grows with its text: nothing can overflow
            report['checked'] += 1
            ins = bodypr_insets(pkg, part, el)
            need, have, broken, approx = measure(ctx, el, box, ins)
            if approx:
                for p in el.find(P + 'txBody').findall(A + 'p'):
                    _, r = ctx.resolve_paragraph(el, p)
                    report['approximate_fonts'].add(r.get('font'))
            where = dict(slide=n, shape_id=shape_id(el), name=shape_name(el))
            if broken:
                report['broken_words'].append(dict(where, words=broken[:5]))
            if need <= have:
                continue
            # truth size of this shape's first paragraph
            p0 = el.find(P + 'txBody/' + A + 'p')
            _, truth_run = ctx.resolve_paragraph(el, p0)
            runs = p0.findall(A + 'r')
            current = ctx.resolve_run(el, p0, runs[0])[1]['size'] if runs else truth_run['size']
            truth_size = truth_run['size']
            minimum = a.min_small if truth_size < 12 else a.min_body
            floor = max(truth_size * a.floor, minimum)
            if f'{n}:{shape_id(el)}' in allowed:
                floor = minimum
            size = current
            fitted = False
            while size - 1 >= floor - 1e-6:
                size -= 1
                if measure(ctx, el, box, ins, scale=size / current)[0] <= have:
                    fitted = True
                    break
            lnspc = None
            if not fitted and measure(ctx, el, box, ins, scale=max(size, floor) / current, lnspc_override=0.9)[0] <= have:
                fitted, lnspc = True, 90
            size = max(size, floor)
            if fitted:
                set_size(el, size)
                if lnspc:
                    set_lnspc(el, lnspc)
                report['shrinks'].append(dict(where, from_pt=current, to_pt=size, line_spacing=lnspc,
                                              truth_pt=truth_size))
            else:
                report['needs_decision'].append(dict(
                    where, need_pt=round(need, 1), have_pt=round(have, 1), truth_pt=truth_size,
                    floor_pt=floor, options=['split the slide in two (keeps every word)',
                                             'allow a smaller size this once (--allow)',
                                             'the user trims the text']))

    report['approximate_fonts'] = sorted(f for f in report['approximate_fonts'] if f)
    pkg.save(a.out)
    meta['fit'] = report
    dump_json(meta, a.out + '.meta.json')
    if a.report:
        dump_json(report, a.report)
    print(f"text shapes checked: {report['checked']}, shrunk: {len(report['shrinks'])}, "
          f"needs decision: {len(report['needs_decision'])}")
    for s in report['shrinks']:
        print(f"  shrink slide {s['slide']} '{s['name']}': {s['from_pt']} -> {s['to_pt']} pt"
              + (f", line spacing {s['line_spacing']}%" if s['line_spacing'] else ''))
    for s in report['needs_decision']:
        print(f"  STOP slide {s['slide']} '{s['name']}': needs {s['need_pt']} pt of height, has {s['have_pt']} "
              f"even at the {s['floor_pt']} pt floor — ask the user")
    for b in report['broken_words']:
        print(f"  word too long for its box: slide {b['slide']} '{b['name']}': {b['words']}")
    if report['approximate_fonts']:
        print('  fonts not installed (estimates use substitutes): ' + ', '.join(report['approximate_fonts']))


if __name__ == '__main__':
    main()
