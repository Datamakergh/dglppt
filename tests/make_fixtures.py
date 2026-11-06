#!/usr/bin/env python3
"""Build the test decks used by tests/run_tests.py.

    python tests/make_fixtures.py OUT_DIR

Creates:
  truth.pptx          16:9 truth deck: custom theme/fonts, master accent bar + logo,
                      custom placeholder idx on "Title and Content", footer + slide
                      number shown, a slide-level badge, table, chart, captioned picture
  truth.potx          the same design as a template with no slides
  target.pptx         4:3 deck in the default Office design with local overrides,
                      semantic red, hyperlink, notes, old footer/logo leftovers,
                      an icon row (3 sizes), GIF, entrance animations, a transition,
                      embedded MP4 with poster frame, table, chart, a group,
                      the same screenshot on two slides at different sizes, a slide that
                      needs a gentle shrink and one that must stop and ask
"""
import copy
import io
import os
import subprocess
import sys

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))
from dglcore import A, P, NS, Package, iter_shapes, ph_of, sp_tree  # noqa: E402

IN = 914400


def png(w, h, color, text=None, shape='rect'):
    im = Image.new('RGBA', (w, h), (255, 255, 255, 0))
    d = ImageDraw.Draw(im)
    if shape == 'circle':
        d.ellipse([2, 2, w - 3, h - 3], fill=color)
    else:
        d.rectangle([0, 0, w - 1, h - 1], fill=color)
    if text:
        d.text((8, h // 2 - 6), text, fill='white')
    buf = io.BytesIO()
    im.save(buf, 'PNG')
    buf.seek(0)
    return buf


def screenshot():
    im = Image.new('RGB', (400, 300), 'white')
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, 399, 40], fill=(40, 60, 120))
    for i in range(5):
        d.rectangle([20, 60 + i * 45, 380, 90 + i * 45], fill=(200, 210, 230))
    buf = io.BytesIO(); im.save(buf, 'PNG'); buf.seek(0)
    return buf


def gif():
    frames = []
    for c in ((255, 0, 0), (0, 160, 0), (0, 0, 255)):
        im = Image.new('RGB', (120, 120), c)
        frames.append(im)
    buf = io.BytesIO()
    frames[0].save(buf, 'GIF', save_all=True, append_images=frames[1:], duration=300, loop=0)
    buf.seek(0)
    return buf


def mp4(path):
    subprocess.run(['ffmpeg', '-loglevel', 'error', '-y', '-f', 'lavfi', '-i',
                    'testsrc=duration=1:size=160x90:rate=10', '-pix_fmt', 'yuv420p', path], check=True)


def set_xfrm(el, x, y, w, h):
    sppr = el.find(P + 'spPr')
    xf = sppr.find(A + 'xfrm')
    if xf is None:
        xf = etree.SubElement(sppr, A + 'xfrm')
        sppr.insert(0, xf)
    for c in list(xf):
        xf.remove(c)
    off = etree.SubElement(xf, A + 'off'); off.set('x', str(int(x * IN))); off.set('y', str(int(y * IN)))
    ext = etree.SubElement(xf, A + 'ext'); ext.set('cx', str(int(w * IN))); ext.set('cy', str(int(h * IN)))


def drop_xfrm(el):
    sppr = el.find(P + 'spPr')
    xf = sppr.find(A + 'xfrm') if sppr is not None else None
    if xf is not None:
        sppr.remove(xf)


def add_master_picture(master, image, x, y, w, h, name):
    _, rid = master.part.get_or_add_image_part(image)
    pic = etree.fromstring(f'''<p:pic xmlns:p="{NS['p']}" xmlns:a="{NS['a']}" xmlns:r="{NS['r']}">
  <p:nvPicPr><p:cNvPr id="90" name="{name}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="{int(x*IN)}" y="{int(y*IN)}"/><a:ext cx="{int(w*IN)}" cy="{int(h*IN)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>''')
    sp_tree(master._element).append(pic)


def add_master_bar(master, W):
    sp = etree.fromstring(f'''<p:sp xmlns:p="{NS['p']}" xmlns:a="{NS['a']}">
  <p:nvSpPr><p:cNvPr id="91" name="Accent Bar"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{int(0.12*IN)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>
</p:sp>''')
    sp_tree(master._element).insert(2, sp)


def theme_part(prs):
    master = prs.slide_masters[0]
    for rel in master.part.rels.values():
        if rel.reltype.endswith('/theme'):
            return rel.target_part
    raise RuntimeError('no theme')


def restyle_theme(prs):
    tp = theme_part(prs)
    root = etree.fromstring(tp.blob)
    colors = dict(dk1='1A1A1A', lt1='FFFFFF', dk2='1E2761', lt2='F2F2F2', accent1='1E2761',
                  accent2='F96167', accent3='2F3C7E', accent4='97BC62', accent5='CADCFC', accent6='F9E795',
                  hlink='2F3C7E', folHlink='6D2E46')
    cs = root.find('.//' + A + 'clrScheme')
    cs.set('name', 'Truth Colors')
    for child in cs:
        name = etree.QName(child).localname
        for c in list(child):
            child.remove(c)
        e = etree.SubElement(child, A + 'srgbClr'); e.set('val', colors[name])
    fs = root.find('.//' + A + 'fontScheme')
    fs.set('name', 'Truth Fonts')
    fs.find(A + 'majorFont/' + A + 'latin').set('typeface', 'Caladea')
    fs.find(A + 'minorFont/' + A + 'latin').set('typeface', 'Carlito')
    root.set('name', 'Truth Theme')
    tp._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def set_txstyles(master):
    tx = master._element.find(P + 'txStyles')
    t1 = tx.find(P + 'titleStyle/' + A + 'lvl1pPr')
    d = t1.find(A + 'defRPr'); d.set('sz', '3200'); d.set('b', '1')
    for c in list(d):
        if etree.QName(c).localname == 'solidFill':
            d.remove(c)
    sf = etree.Element(A + 'solidFill'); sc = etree.SubElement(sf, A + 'schemeClr'); sc.set('val', 'accent1')
    d.insert(0, sf)
    t1.set('algn', 'l')
    body = tx.find(P + 'bodyStyle')
    for lvl, sz in ((1, 2000), (2, 1800), (3, 1600), (4, 1400), (5, 1400)):
        l = body.find(A + f'lvl{lvl}pPr')
        l.find(A + 'defRPr').set('sz', str(sz))


def build_truth(out_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    W = prs.slide_width
    restyle_theme(prs)
    m = prs.slide_masters[0]
    set_txstyles(m)
    for el, _ in iter_shapes(sp_tree(m._element)):
        ph = ph_of(el)
        if not ph:
            continue
        t = ph[0]
        if t == 'title': set_xfrm(el, 0.6, 0.45, 11.2, 0.9)
        elif t == 'body': set_xfrm(el, 0.6, 1.6, 12.133, 4.9)
        elif t == 'dt': set_xfrm(el, 0.6, 6.85, 2.5, 0.35)
        elif t == 'ftr': set_xfrm(el, 4.4, 6.85, 4.5, 0.35)
        elif t == 'sldNum': set_xfrm(el, 10.233, 6.85, 2.5, 0.35)
    add_master_bar(m, W)
    add_master_picture(m, png(180, 70, (30, 39, 97), 'TRUTH'), 12.033, 0.25, 0.9, 0.35, 'Logo')
    for lay in m.slide_layouts:
        for el, _ in iter_shapes(sp_tree(lay._element)):
            ph = ph_of(el)
            if not ph:
                continue
            if lay.name == 'Title Slide':
                if ph[0] == 'ctrTitle': set_xfrm(el, 0.6, 2.4, 12.133, 1.5)
                elif ph[0] == 'subTitle': set_xfrm(el, 0.6, 4.0, 12.133, 1.0)
                else: drop_xfrm(el)
            elif lay.name == 'Two Content':
                if ph[1] == 1: set_xfrm(el, 0.6, 1.6, 5.9, 4.9)
                elif ph[1] == 2: set_xfrm(el, 6.833, 1.6, 5.9, 4.9)
                else: drop_xfrm(el)
            else:
                # scale 4:3 layouts horizontally, or inherit from the master
                if lay.name in ('Title and Content', 'Title Only', 'Blank', 'Section Header'):
                    drop_xfrm(el)
                else:
                    xf = el.find(P + 'spPr/' + A + 'xfrm')
                    if xf is not None:
                        off, ext = xf.find(A + 'off'), xf.find(A + 'ext')
                        off.set('x', str(int(int(off.get('x')) * 4 / 3)))
                        ext.set('cx', str(int(int(ext.get('cx')) * 4 / 3)))
        if lay.name == 'Title and Content':
            # custom placeholder indexes, like many corporate templates
            for el, _ in iter_shapes(sp_tree(lay._element)):
                phe = el.find('.//' + P + 'ph')
                if phe is None:
                    continue
                remap = {'1': '13', '10': '20', '11': '21', '12': '22'}
                if phe.get('idx') in remap:
                    phe.set('idx', remap[phe.get('idx')])

    badge = png(64, 64, (249, 97, 103), shape='circle')

    def add_furniture(slide, footer_text='ACME Corp — Internal'):
        lay = slide.slide_layout
        for el, _ in iter_shapes(sp_tree(lay._element)):
            ph = ph_of(el)
            if ph and ph[0] in ('ftr', 'sldNum'):
                new = copy.deepcopy(el)
                drop_xfrm(new)
                if ph[0] == 'ftr':
                    for t in new.iter(A + 't'):
                        t.text = ''
                    ts = list(new.iter(A + 't'))
                    if ts:
                        ts[0].text = footer_text
                    else:
                        p = new.find(P + 'txBody/' + A + 'p')
                        r = etree.SubElement(p, A + 'r'); etree.SubElement(r, A + 'rPr').set('lang', 'en-US')
                        etree.SubElement(r, A + 't').text = footer_text
                sp_tree(slide._element).append(new)
        badge.seek(0)
        slide.shapes.add_picture(badge, Inches(0.15), Inches(6.85), Inches(0.35), Inches(0.35)).name = 'Badge'

    s = prs.slides.add_slide(m.slide_layouts.get_by_name('Title Slide'))
    s.shapes.title.text = 'Quarterly Review'
    s.placeholders[1].text = 'Finance team · Q3'

    s = prs.slides.add_slide(m.slide_layouts.get_by_name('Title and Content'))
    s.shapes.title.text = 'What changed'
    body = [p for p in s.placeholders if p.placeholder_format.idx == 13][0]
    tf = body.text_frame
    tf.text = 'Costs fell in every region'
    p = tf.add_paragraph(); p.text = 'Freight renegotiated'; p.level = 1
    p = tf.add_paragraph(); p.text = 'Hiring paused until Q1'
    add_furniture(s)

    s = prs.slides.add_slide(m.slide_layouts.get_by_name('Title Only'))
    s.shapes.title.text = 'By region'
    gf = s.shapes.add_table(3, 3, Inches(0.6), Inches(1.6), Inches(12.133), Inches(1.5))
    for r in range(3):
        for c in range(3):
            cell = gf.table.cell(r, c)
            cell.text = f'R{r}C{c}'
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(14)
    add_furniture(s)

    s = prs.slides.add_slide(m.slide_layouts.get_by_name('Title Only'))
    s.shapes.title.text = 'Trend'
    cd = CategoryChartData(); cd.categories = ['Q1', 'Q2', 'Q3']
    cd.add_series('Plan', (1, 2, 3)); cd.add_series('Actual', (1.5, 2.5, 2.8))
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.6), Inches(12.133),
                            Inches(4.9), cd).chart
    ch.plots[0].series[0].format.fill.solid(); ch.plots[0].series[0].format.fill.fore_color.rgb = RGBColor(0x1E, 0x27, 0x61)
    ch.plots[0].series[1].format.fill.solid(); ch.plots[0].series[1].format.fill.fore_color.rgb = RGBColor(0xF9, 0x61, 0x67)
    ch.has_legend = True
    ch.font.size = Pt(12)
    add_furniture(s)

    s = prs.slides.add_slide(m.slide_layouts.get_by_name('Title Only'))
    s.shapes.title.text = 'Our office'
    pic = s.shapes.add_picture(screenshot(), Inches(0.6), Inches(1.6), Inches(6.0), Inches(4.5))
    pic.line.color.rgb = RGBColor(0xCA, 0xDC, 0xFC); pic.line.width = Pt(1.5)
    tb = s.shapes.add_textbox(Inches(6.9), Inches(1.6), Inches(5.833), Inches(1.0))
    tb.text_frame.text = 'Opened in May'
    r = tb.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(16); r.font.color.rgb = RGBColor(0x2F, 0x3C, 0x7E)
    add_furniture(s)

    path = os.path.join(out_dir, 'truth.pptx')
    prs.save(path)
    return path


def build_potx(truth_path, out_dir):
    pkg = Package(open(truth_path, 'rb').read())
    pres = pkg.xml(pkg.main_part)
    lst = pres.find(P + 'sldIdLst')
    root = pkg.rels_root(pkg.main_part)
    for sid in list(lst):
        rid = sid.get('{%s}id' % NS['r'])
        lst.remove(sid)
        for rel in list(root):
            if rel.get('Id') == rid:
                root.remove(rel)
    pkg.gc()
    ct = pkg.xml('[Content_Types].xml')
    for o in ct:
        if o.get('ContentType') == 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml':
            o.set('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml')
    path = os.path.join(out_dir, 'truth.potx')
    pkg.save(path)
    return path


TIMING = '''<p:timing xmlns:p="{p}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{pars}</p:childTnLst></p:cTn>
<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>
</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'''

PAR = '''<p:par><p:cTn id="{a}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="{b}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="{c}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">
<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:set><p:cBhvr><p:cTn id="{d}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>
<p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'''


def add_animations(slide, spids):
    pars = ''.join(PAR.format(a=3 + 4 * i, b=4 + 4 * i, c=5 + 4 * i, d=6 + 4 * i, spid=s) for i, s in enumerate(spids))
    timing = etree.fromstring(TIMING.format(p=NS['p'], pars=pars))
    trans = etree.fromstring(f'<p:transition xmlns:p="{NS["p"]}" spd="med"><p:fade/></p:transition>')
    root = slide._element
    clr = root.find(P + 'clrMapOvr')
    (clr if clr is not None else root.find(P + 'cSld')).addnext(trans)
    trans.addnext(timing)


def build_target(out_dir):
    prs = Presentation()  # 4:3, default Office theme
    L = {l.name: l for l in prs.slide_layouts}

    s = prs.slides.add_slide(L['Title Slide'])
    s.shapes.title.text = 'Project Phoenix'
    s.placeholders[1].text = 'Status update'

    s = prs.slides.add_slide(L['Title and Content'])
    s.shapes.title.text = 'Highlights'
    tf = s.placeholders[1].text_frame
    tf.text = 'Revenue grew '
    p0 = tf.paragraphs[0]
    r = p0.add_run(); r.text = '12%'; r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0, 0)
    r = p0.add_run(); r.text = ' year on year'
    for run in p0.runs:
        run.font.name = 'Comic Sans MS'; run.font.size = Pt(28)
    p0.alignment = 2  # center
    p = tf.add_paragraph(); p.text = 'Driven by the new pricing page'; p.level = 1
    p = tf.add_paragraph()
    r = p.add_run(); r.text = 'Details on the wiki'
    r.hyperlink.address = 'https://example.com/wiki'
    s.notes_slide.notes_text_frame.text = 'Mention the pricing test.'
    tb = s.shapes.add_textbox(Inches(0.3), Inches(7.0), Inches(3), Inches(0.4)); tb.name = 'Old footer'
    tb.text_frame.text = 'Confidential'
    s.shapes.add_picture(png(120, 50, (0, 112, 192), 'OLD'), Inches(8.6), Inches(0.2), Inches(1.2), Inches(0.5)).name = 'Old logo'

    s = prs.slides.add_slide(L['Title Only'])
    s.shapes.title.text = 'Our pillars'
    ids = []
    for i, (x, size, col) in enumerate(((1.0, 1.0, (30, 39, 97)), (4.0, 1.3, (249, 97, 103)), (7.2, 0.8, (151, 188, 98)))):
        pic = s.shapes.add_picture(png(200, 200, col, shape='circle'), Inches(x), Inches(2.0), Inches(size), Inches(size))
        pic.name = f'Icon {i + 1}'
        ids.append(pic.shape_id)
        cap = s.shapes.add_textbox(Inches(x - 0.3), Inches(3.6), Inches(1.8), Inches(0.5))
        cap.text_frame.text = ['Speed', 'Trust', 'Growth'][i]
    g = s.shapes.add_picture(gif(), Inches(0.5), Inches(4.6), Inches(1.5), Inches(1.5))
    g.name = 'Spinner GIF'
    add_animations(s, ids)

    tmp = os.path.join(out_dir, '_clip.mp4'); mp4(tmp)
    s = prs.slides.add_slide(L['Title Only'])
    s.shapes.title.text = 'Product'
    s.shapes.add_picture(screenshot(), Inches(0.5), Inches(1.5), Inches(4.0), Inches(3.0)).name = 'Screenshot'
    s.shapes.add_movie(tmp, Inches(5.0), Inches(1.5), Inches(4.0), Inches(2.25),
                       poster_frame_image=png(160, 90, (20, 20, 20), 'PLAY'), mime_type='video/mp4').name = 'Demo video'
    gf = s.shapes.add_table(3, 3, Inches(0.5), Inches(5.0), Inches(9.0), Inches(1.2))
    for rr in range(3):
        for cc in range(3):
            gf.table.cell(rr, cc).text = f'{rr}-{cc}'
    os.remove(tmp)

    s = prs.slides.add_slide(L['Title Only'])
    s.shapes.title.text = 'Product again'
    s.shapes.add_picture(screenshot(), Inches(5.5), Inches(1.6), Inches(3.6), Inches(2.7)).name = 'Screenshot'
    cd = CategoryChartData(); cd.categories = ['A', 'B', 'C']
    cd.add_series('S1', (3, 1, 2)); cd.add_series('S2', (2, 2, 1))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1.6), Inches(4.8), Inches(3.6), cd)
    grp = s.shapes.add_group_shape()
    a = grp.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(4), Inches(0.6)); a.text_frame.text = 'Left note'
    b = grp.shapes.add_textbox(Inches(5.0), Inches(5.5), Inches(4), Inches(0.6)); b.text_frame.text = 'Right note'
    grp.name = 'Notes group'

    s = prs.slides.add_slide(L['Title and Content'])
    s.shapes.title.text = 'Details'
    tf = s.placeholders[1].text_frame
    tf.text = 'This slide has a lot of text that will not fit at the truth body size without some help.'
    for i in range(12):
        p = tf.add_paragraph(); p.text = f'Point {i + 2}: another fairly long line of supporting detail for the reader.'

    s = prs.slides.add_slide(L['Title and Content'])
    s.shapes.title.text = 'Appendix'
    tf = s.placeholders[1].text_frame
    tf.text = 'Far too much text for one slide at any sensible size.'
    for i in range(30):
        p = tf.add_paragraph(); p.text = f'Line {i + 2}: the restyle must stop and ask instead of cutting this.'
    path = os.path.join(out_dir, 'target.pptx')
    prs.save(path)
    return path


if __name__ == '__main__':
    out = sys.argv[1] if len(sys.argv) > 1 else 'fixtures'
    os.makedirs(out, exist_ok=True)
    t = build_truth(out)
    print(t)
    print(build_potx(t, out))
    print(build_target(out))
