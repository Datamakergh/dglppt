#!/usr/bin/env python3
"""End-to-end tests for the dglppt scripts.

    python tests/run_tests.py [WORKDIR]

Builds the fixtures, runs the pipeline the way a session would (auto mode, then
again with the user's decisions), checks the restyled deck, and runs mutation
tests: each deliberate content, media or geometry change must make its check fail.
Needs python-pptx, lxml, Pillow, ffmpeg (fixture video); LibreOffice for the
render steps (skipped without it).
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(HERE, '..', 'scripts')
sys.path.insert(0, SCRIPTS)

from lxml import etree  # noqa: E402

from dglcore import (A, P, P14, Package, blip_part, iter_shapes, load_json, ph_of, read_xfrm,  # noqa: E402
                     shape_name, sp_tree, write_xfrm)

RESULTS = []


def check(name, cond, detail=''):
    RESULTS.append((name, bool(cond)))
    print(f"  {'PASS' if cond else 'FAIL'}  {name}" + (f' — {detail}' if detail and not cond else ''))


def script(name, *args):
    p = subprocess.run([sys.executable, os.path.join(SCRIPTS, name)] + [str(a) for a in args],
                       capture_output=True, text=True)
    return p.returncode, p.stdout + p.stderr


def shapes(pkg, n):
    part = pkg.slide_parts()[n - 1]
    return [el for el, _ in iter_shapes(sp_tree(pkg.xml(part)))]


def by_name(pkg, n, name):
    return next(el for el in shapes(pkg, n) if shape_name(el) == name)


def mutate(src, dst, fn):
    pkg = Package(open(src, 'rb').read())
    fn(pkg)
    pkg.save(dst)
    if os.path.exists(src + '.meta.json'):
        shutil.copyfile(src + '.meta.json', dst + '.meta.json')


def main():
    work = sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp(prefix='dglppt-test-')
    fx = os.path.join(work, 'fixtures')
    os.makedirs(fx, exist_ok=True)
    print('workdir:', work)
    subprocess.run([sys.executable, os.path.join(HERE, 'make_fixtures.py'), fx], check=True, capture_output=True)
    truth, potx, target = (os.path.join(fx, f) for f in ('truth.pptx', 'truth.potx', 'target.pptx'))
    approved = os.path.join(fx, 'approved.json')
    json.dump({'remove': [{'slide': 2, 'shape_id': 4, 'why': 'old footer'},
                          {'slide': 2, 'shape_id': 5, 'why': 'old logo'}]}, open(approved, 'w'))

    print('\n# truth extraction')
    rc, out = script('extract_truth.py', truth, os.path.join(work, 'spec.json'))
    spec = load_json(os.path.join(work, 'spec.json'))
    check('extract .pptx', rc == 0, out)
    check('theme fonts read', spec['masters'][0]['theme']['heading_font'] == 'Caladea')
    check('custom placeholder idx recorded', any(p['idx'] == 13 for p in spec['layouts']['m1/Title and Content']['placeholders']))
    check('title-only borders come from the master body', spec['layouts']['m1/Title Only']['borders']['bottom'] == 914400)
    check('logo + bar fixed on master', {'Logo', 'Accent Bar'} <= {f['name'] for f in spec['fixed_elements']})
    check('slide-level badge detected', any(f['source'] == 'slide' and f['name'] == 'Badge' for f in spec['fixed_elements']))
    check('title slide shows no footer', spec['layouts']['m1/Title Slide']['show_fixed_placeholders'] == [])
    rc, out = script('extract_truth.py', potx, os.path.join(work, 'spec_potx.json'))
    check('extract .potx without slides', rc == 0 and load_json(os.path.join(work, 'spec_potx.json'))['layouts'], out)

    print('\n# pipeline, auto mode')
    w1 = os.path.join(work, 'auto')
    rc, out = subprocess.run([sys.executable, os.path.join(SCRIPTS, 'pipeline.py'), truth, target, '--work', w1,
                              '--approved', approved], capture_output=True, text=True).returncode, ''
    check('auto mode stops for decisions (exit 2)', rc == 2)
    prop = load_json(os.path.join(w1, 'target', 'groups_proposed.json'))
    kinds = sorted(g['kind'] for g in prop['groups'])
    check('groups proposed: across + row (+ a column to reject)', kinds == ['across', 'column', 'row'], kinds)

    print('\n# pipeline with the user\'s decisions')
    confirmed = os.path.join(work, 'groups_confirmed.json')
    json.dump(dict(groups=[g for g in prop['groups'] if g['kind'] != 'column']), open(confirmed, 'w'))
    w2 = os.path.join(work, 'decided')
    p = subprocess.run([sys.executable, os.path.join(SCRIPTS, 'pipeline.py'), truth, target, '--work', w2,
                        '--approved', approved, '--groups', confirmed, '--spec', os.path.join(w1, 'truth_spec.json')],
                       capture_output=True, text=True)
    D = os.path.join(w2, 'target')
    final = os.path.join(D, 'target - restyled.pptx')
    check('only the overflow question remains (exit 2)', p.returncode == 2, p.stdout[-400:])
    for k in ('content', 'geometry', 'media'):
        d = load_json(os.path.join(D, f'{k}_check.json'))
        check(f'{k} check passes', d['passed'], d.get('fails'))
    if os.path.exists(os.path.join(D, 'pixel_check.json')):
        px = load_json(os.path.join(D, 'pixel_check.json'))
        check('pixel check fails only on the overflow slide 7', {f['slide'] for f in px['fails']} <= {7}, px['fails'])

    pkg = Package(open(final, 'rb').read())
    meta = load_json(final + '.meta.json')
    check('slide size is the truth size', list(pkg.slide_size()) == [spec['canvas']['w'], spec['canvas']['h']])
    for n in range(2, 8):
        types = {ph_of(el)[0] for el in shapes(pkg, n) if ph_of(el)}
        check(f'slide {n} shows footer + slide number', {'ftr', 'sldNum'} <= types)
        check(f'slide {n} has the truth badge', any(shape_name(el).startswith('dglppt-fixed') for el in shapes(pkg, n)))
    check('title slide has no footer', not any(ph_of(el) and ph_of(el)[0] == 'ftr' for el in shapes(pkg, 1)))
    icons = [read_xfrm(by_name(pkg, 3, f'Icon {i}')) for i in (1, 2, 3)]
    check('icon row: identical size', len({(b[2], b[3]) for b in icons}) == 1, icons)
    check('icon row: shared top edge', len({b[1] for b in icons}) == 1, icons)
    gaps = [icons[i + 1][0] - icons[i][0] - icons[i][2] for i in range(2)]
    check('icon row: equal gaps', abs(gaps[0] - gaps[1]) <= 1, gaps)
    caps = [read_xfrm(el) for el in shapes(pkg, 3) if shape_name(el).startswith('TextBox')]
    check('captions on one line', len({c[1] for c in caps}) == 1, caps)
    check('captions centered on their icons',
          all(abs((c[0] + c[2] / 2) - (i[0] + i[2] / 2)) <= 1 for c, i in zip(sorted(caps), sorted(icons))))
    s4, s5 = read_xfrm(by_name(pkg, 4, 'Screenshot')), read_xfrm(by_name(pkg, 5, 'Screenshot'))
    check('same screenshot, same size on both slides', s4[2:] == s5[2:], (s4, s5))
    check('slide 6 shrunk 20 -> 18 pt', any(s['slide'] == 6 and s['to_pt'] == 18 for s in meta['fit']['shrinks']))
    check('slide 7 stopped for a decision', any(s['slide'] == 7 for s in meta['fit']['needs_decision']))
    check('semantic red kept', any(s.get('text') == '12%' for s in meta['style']['semantic_color_kept']))
    gif = by_name(pkg, 3, 'Spinner GIF')
    b = read_xfrm(gif)
    check('GIF keeps its aspect ratio', b[2] == b[3])
    before = os.path.join(D, 'original.pptx')

    print('\n# mutation tests: content lock')

    def edit_word(pkg):
        for el in shapes(pkg, 6):
            for t in el.iter(A + 't'):
                if t.text and 'Point 3' in t.text:
                    t.text = t.text.replace('detail', 'details')
                    return

    def swap_bullets(pkg):
        el = next(e for e in shapes(pkg, 6) if ph_of(e) and ph_of(e)[0] != 'title' and e.find(P + 'txBody') is not None
                  and len(e.find(P + 'txBody').findall(A + 'p')) > 3)
        ps = el.find(P + 'txBody').findall(A + 'p')
        ps[1].addprevious(ps[2])

    def chart_value(pkg):
        name = next(n for n in pkg.names() if n.startswith('ppt/charts/chart') and n.endswith('.xml'))
        v = next(pkg.xml(name).iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}v'))
        v.text = '99'

    def drop_note(pkg):
        name = next(n for n in pkg.names() if n.startswith('ppt/notesSlides/') and n.endswith('.xml'))
        t = next(pkg.xml(name).iter(A + 't'))
        t.text = 'changed'

    for label, fn in (('edit one word', edit_word), ('swap two bullets', swap_bullets),
                      ('change a chart value', chart_value), ('change speaker notes', drop_note)):
        m = os.path.join(work, 'mut.pptx')
        mutate(final, m, fn)
        script('fingerprint.py', m, os.path.join(work, 'mut.json'))
        rc, out = script('compare_fingerprint.py', os.path.join(D, 'before_fingerprint.json'),
                         os.path.join(work, 'mut.json'), '--approved', approved)
        check(f'content check catches: {label}', rc == 1, out[-300:])

    print('\n# mutation tests: media')

    def trim_video(pkg):
        for el in shapes(pkg, 4):
            for med in el.iter(P14 + 'media'):
                etree.SubElement(med, P14 + 'trim').set('st', '500')

    def drop_effect(pkg):
        root = pkg.xml(pkg.slide_parts()[2])
        lst = root.find('.//' + P + 'seq/' + P + 'cTn/' + P + 'childTnLst')
        lst.remove(lst[-1])

    def flatten_gif(pkg):
        import io
        from PIL import Image
        part = pkg.slide_parts()[2]
        img = blip_part(pkg, part, by_name(pkg, 3, 'Spinner GIF'))
        buf = io.BytesIO()
        Image.open(io.BytesIO(pkg.blob(img))).convert('RGB').save(buf, 'GIF')
        pkg.set_blob(img, buf.getvalue())

    def stretch_video(pkg):
        el = by_name(pkg, 4, 'Demo video')
        b = read_xfrm(el)
        write_xfrm(el, [b[0], b[1], b[2], int(b[3] * 1.2)])

    for label, fn in (('video trim changed', trim_video), ('animation effect deleted', drop_effect),
                      ('GIF re-encoded to one frame', flatten_gif), ('video stretched', stretch_video)):
        m = os.path.join(work, 'mut.pptx')
        mutate(final, m, fn)
        rc, out = script('check_media.py', before, m)
        check(f'media check catches: {label}', rc == 1, out[-300:])

    print('\n# mutation tests: geometry')

    def nudge_badge(pkg):
        el = next(e for e in shapes(pkg, 3) if shape_name(e).startswith('dglppt-fixed'))
        b = read_xfrm(el)
        write_xfrm(el, [b[0] + 12700, b[1], b[2], b[3]])

    def local_title_size(pkg):
        el = next(e for e in shapes(pkg, 2) if ph_of(e) and ph_of(e)[0] == 'title')
        r = el.find('.//' + A + 'r')
        rpr = r.find(A + 'rPr')
        if rpr is None:
            rpr = etree.Element(A + 'rPr')
            r.insert(0, rpr)
        rpr.set('sz', '4000')

    def into_border(pkg):
        el = by_name(pkg, 4, 'Screenshot')
        b = read_xfrm(el)
        write_xfrm(el, [0, b[1], b[2], b[3]])

    def break_row(pkg):
        el = by_name(pkg, 3, 'Icon 2')
        b = read_xfrm(el)
        write_xfrm(el, [b[0], b[1] + 50000, b[2], b[3]])

    for label, fn in (('logo/badge moved 1 pt', nudge_badge), ('local title size', local_title_size),
                      ('picture in the border', into_border), ('icon out of its row', break_row)):
        m = os.path.join(work, 'mut.pptx')
        mutate(final, m, fn)
        rc, out = script('check_geometry.py', m, truth, os.path.join(w1, 'truth_spec.json'))
        check(f'geometry check catches: {label}', rc == 1, out[-300:])

    print('\n# truth-slide design copied onto targets')
    moved = [123456, 234567, 3456789, 456789]
    tp = os.path.join(work, 'truth_designed.pptx')
    ts0 = spec['slides'][0]
    info = {}

    def design_truth(tpkg):
        part = tpkg.slide_parts()[0]
        root = tpkg.xml(part)
        for el in shapes(tpkg, 1):
            ph = ph_of(el)
            if not ph or ph[0] in ('ftr', 'sldNum', 'dt', 'hdr') or el.find(P + 'txBody') is None:
                continue
            info['key'] = (ph[0], ph[1])
            write_xfrm(el, moved)
            sp = el.find(P + 'spPr')
            sp.append(etree.fromstring('<a:solidFill xmlns:a="%s"><a:srgbClr val="FFEEAA"/></a:solidFill>' % A[1:-1]))
            el.find(P + 'txBody/' + A + 'bodyPr').set('anchor', 'b')
            p = el.find(P + 'txBody/' + A + 'p')
            p.insert(0, etree.fromstring('<a:pPr xmlns:a="%s" algn="r"/>' % A[1:-1]))
            for r in p.findall(A + 'r'):
                rpr = r.find(A + 'rPr')
                if rpr is None:
                    rpr = etree.Element(A + 'rPr')
                    r.insert(0, rpr)
                rpr.set('sz', '3200')
                rpr.insert(0, etree.fromstring('<a:solidFill xmlns:a="%s"><a:srgbClr val="C00000"/></a:solidFill>' % A[1:-1]))
            break
        cs = root.find(P + 'cSld')
        cs.insert(0, etree.fromstring('<p:bg xmlns:p="%s" xmlns:a="%s"><p:bgPr><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill>'
                                      '<a:effectLst/></p:bgPr></p:bg>' % (P[1:-1], A[1:-1])))
        tree = sp_tree(root)
        tree.append(etree.fromstring(
            '<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="900" name="Rule"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="500000" y="900000"/><a:ext cx="2000000" cy="40000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></p:spPr></p:sp>'
            % (P[1:-1], A[1:-1])))
        tree.append(etree.fromstring(
            '<p:sp xmlns:p="%s" xmlns:a="%s"><p:nvSpPr><p:cNvPr id="901" name="Callout"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="600000" y="1500000"/><a:ext cx="2000000" cy="400000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>NEW!</a:t></a:r></a:p></p:txBody></p:sp>'
            % (P[1:-1], A[1:-1])))

    mutate(truth, tp, design_truth)
    check('truth fixture: hand-made design on a truth slide', 'key' in info, info)
    step3 = os.path.join(D, 'step3.pptx')
    out3 = os.path.join(work, 'design.pptx')
    rc, out = script('apply_truth_slide_design.py', step3, tp, os.path.join(w1, 'truth_spec.json'), out3)
    check('apply_truth_slide_design runs', rc == 0, out)
    dm = load_json(out3 + '.meta.json')
    pk = Package(open(out3, 'rb').read())
    hits = [int(n) for n, k in load_json(step3 + '.meta.json')['layout_for_slide'].items() if k == ts0['layout']]
    check('some target uses that layout', bool(hits), hits)
    allfine, other_ok = True, True
    for n in hits:
        el = next(e for e in shapes(pk, n) if ph_of(e) and (ph_of(e)[0], ph_of(e)[1]) == info['key'])
        allfine = allfine and read_xfrm(el) == moved
        allfine = allfine and el.find(P + 'spPr/' + A + 'solidFill/' + A + 'srgbClr').get('val') == 'FFEEAA'
        allfine = allfine and el.find(P + 'txBody/' + A + 'bodyPr').get('anchor') == 'b'
        p0 = el.find(P + 'txBody/' + A + 'p')
        allfine = allfine and p0.find(A + 'pPr').get('algn') == 'r'
        r0 = p0.find(A + 'r/' + A + 'rPr')
        allfine = allfine and r0.get('sz') == '3200'
        allfine = allfine and r0.find(A + 'solidFill/' + A + 'srgbClr').get('val') == 'C00000'
        root = pk.xml(pk.slide_parts()[n - 1])
        allfine = allfine and root.find(P + 'cSld/' + P + 'bg') is not None
        names = [shape_name(e) for e in shapes(pk, n)]
        allfine = allfine and 'dglppt-design Rule' in names and 'dglppt-design Callout' not in names
    check('position, fill, anchor, alignment, size, color, background and decorative shape copied', allfine, hits)
    check('a one-off shape with text is asked about, not copied',
          any('Callout' in q for q in dm['design']['questions']), dm['design']['questions'])
    for n, k in load_json(step3 + '.meta.json')['layout_for_slide'].items():
        if k != ts0['layout']:
            other_ok = other_ok and not any(x['slide'] == int(n) and 'Rule' in x.get('name', '') for x in dm['design']['added'])
    check('slides on other layouts are untouched', other_ok)
    apd = os.path.join(work, 'approved_callout.json')
    json.dump({'add': [{'layout': ts0['layout'], 'shape_name': 'Callout'}]}, open(apd, 'w'))
    out3b = os.path.join(work, 'design_b.pptx')
    rc, out = script('apply_truth_slide_design.py', step3, tp, os.path.join(w1, 'truth_spec.json'), out3b, '--approved', apd)
    pkb = Package(open(out3b, 'rb').read())
    check('an approved text one-off is copied', rc == 0 and any(
        shape_name(e) == 'dglppt-design Callout' for n in hits for e in shapes(pkb, n)), out)

    # best-matching truth slide wins
    import apply_truth_slide_design as atd
    one = {('title', '0'): 'Quarterly results'}
    two = {('title', '0'): 'Quarterly results', ('body', '1'): 'x' * 400}
    check('donor score prefers the same placeholders', atd.score(two, two) > atd.score(one, two))
    check('donor score prefers a similar amount of text',
          atd.score({('body', '1'): 'x' * 380}, {('body', '1'): 'x' * 400}) >
          atd.score({('body', '1'): 'x'}, {('body', '1'): 'x' * 400}))

    # a copied change that pushes text out of its box is never shrunk silently: it asks
    def huge(p):
        for n in hits:
            el = next(e for e in shapes(p, n) if ph_of(e) and (ph_of(e)[0], ph_of(e)[1]) == info['key'])
            for r in el.iter(A + 'r'):
                r.find(A + 'rPr').set('sz', '9000')
    mutate(out3, os.path.join(work, 'design_huge.pptx'), huge)
    rc, out = script('fit_text.py', os.path.join(work, 'design_huge.pptx'), os.path.join(work, 'fit_out.pptx'),
                     '--report', os.path.join(work, 'fit_huge.json'))
    fr = load_json(os.path.join(work, 'fit_huge.json'))
    check('overflow from a copied design asks first, never shrinks', any(
        x.get('reason') == 'truth-slide design' for x in fr['needs_decision']) and not any(
        s['slide'] in hits and s['name'] == next(d['name'] for d in dm['design']['shapes'] if d['slide'] == s['slide'])
        for s in fr['shrinks']), fr['needs_decision'])

    failed = [n for n, ok in RESULTS if not ok]
    print(f'\n{len(RESULTS) - len(failed)}/{len(RESULTS)} passed')
    sys.exit(1 if failed else 0)


if __name__ == '__main__':
    main()
