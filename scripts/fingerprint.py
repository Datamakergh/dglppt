#!/usr/bin/env python3
"""Content fingerprint of a deck — everything the content lock protects.

    python fingerprint.py deck.pptx out.json

Per slide, keyed by shape id (restyling edits shapes in place, so ids survive):
  * text of every paragraph, in order, including fields and line breaks
  * hyperlinks and click actions, alt text (descr/title)
  * table cells, chart categories/series values, SmartArt text
  * image hashes (pictures, picture placeholders, fills), GIF frames/durations
  * video/audio blobs or links, poster frames, trim/bookmarks/playback settings
  * OLE objects, the animation timeline (p:timing) and the transition
Deck level: slide count and order, hidden flags, section names, speaker notes.
"""
import sys

from dglcore import (A, P, P14, R, NS, RT_PACKAGE,
                     blip_part, c14n_str, dump_json, image_info, iter_shapes,
                     media_parts, nv_pr, nv_props, open_package, paragraph_text,
                     ph_of, sha1, shape_id, shape_kind, shape_name, sp_tree)


def hyperlinks(pkg, part, el):
    out = []
    for tag in ('hlinkClick', 'hlinkHover'):
        for h in el.iter(A + tag):
            rid = h.get(R + 'id')
            tgt = None
            if rid:
                rel = pkg.rel_target(part, rid)
                if rel:
                    tgt = rel['target'] if rel['mode'] == 'External' else rel['abs']
            out.append(dict(kind=tag, target=tgt, action=h.get('action')))
    return out


def chart_data(pkg, part):
    """Categories and values straight from the chart XML caches (c:cat/c:val/c:tx)."""
    root = pkg.xml(part)
    C = '{%s}' % NS['c']
    series = []
    for ser in root.iter(C + 'ser'):
        def pts(tag):
            node = ser.find(C + tag)
            if node is None:
                return []
            vals = []
            for pt in node.iter(C + 'pt'):
                v = pt.find(C + 'v')
                vals.append((int(pt.get('idx')), v.text if v is not None else None))
            return [v for _, v in sorted(vals)]
        name = ser.find(C + 'tx')
        series.append(dict(name=[v.text for v in name.iter(C + 'v')] if name is not None else [],
                           cat=pts('cat'), val=pts('val'), x=pts('xVal'), y=pts('yVal'),
                           bubble=pts('bubbleSize')))
    title = root.find('.//' + C + 'title')
    title_text = ''.join(t.text or '' for t in title.iter(A + 't')) if title is not None else ''
    ext = None
    for r in pkg.rels(part):
        if r['type'] == RT_PACKAGE and r['abs'] and pkg.has(r['abs']):
            ext = sha1(pkg.blob(r['abs']))
    return dict(title=title_text, series=series, workbook_sha1=ext)


def smartart_text(pkg, part, frame):
    rel_ids = frame.find('.//{%s}relIds' % NS['dgm'])
    if rel_ids is None:
        return []
    rel = pkg.rel_target(part, rel_ids.get(R + 'dm'))
    if not rel or not pkg.has(rel['abs']):
        return []
    root = pkg.xml(rel['abs'])
    out = []
    for pt in root.iter('{%s}pt' % NS['dgm']):
        t = pt.find('{%s}t' % NS['dgm'])
        if t is None:
            continue
        paras = [paragraph_text(p) for p in t.findall(A + 'p')]
        if any(x.strip() for x in paras):
            out.append(dict(model_id=pt.get('modelId'), text=paras))
    return out


def picture_record(pkg, part, el, rec):
    img = blip_part(pkg, part, el)
    if img:
        blob = pkg.blob(img)
        rec['image_sha1'] = sha1(blob)
        info = image_info(blob)
        if info and info['frames'] > 1:
            rec['gif'] = dict(frames=info['frames'], durations=info['durations'])
    else:
        blip = el.find('.//' + A + 'blip')
        if blip is not None and blip.get(R + 'link'):
            rel = pkg.rel_target(part, blip.get(R + 'link'))
            rec['image_link'] = rel['target'] if rel else None
    src = el.find('.//' + A + 'srcRect')
    if src is not None:
        rec['crop'] = dict(src.attrib)
    media = media_parts(pkg, part, el)
    if media:
        rec['media'] = []
        for kind, tgt, external in media:
            if external:
                rec['media'].append(dict(kind=kind, link=tgt))
            else:
                rec['media'].append(dict(kind=kind, sha1=sha1(pkg.blob(tgt)) if pkg.has(tgt) else None,
                                         ext=tgt.rsplit('.', 1)[-1].lower()))
        nvpr = nv_pr(el)
        m = nvpr.find('.//' + P14 + 'media') if nvpr is not None else None
        if m is not None:
            rec['media_settings'] = c14n_str(m)  # trim, fade, bookmarks


def shape_record(pkg, part, el):
    kind = shape_kind(el)
    c = nv_props(el)
    rec = dict(id=shape_id(el), name=shape_name(el), kind=kind)
    if c is not None and (c.get('descr') or c.get('title')):
        rec['alt'] = dict(descr=c.get('descr'), title=c.get('title'))
    ph = ph_of(el)
    if ph:
        rec['ph'] = list(ph)
    tb = el.find(P + 'txBody')
    if tb is not None:
        paras = [paragraph_text(p) for p in tb.findall(A + 'p')]
        # trailing empty paragraphs carry no content
        while paras and not paras[-1].strip():
            paras.pop()
        if paras:
            rec['text'] = paras
    links = hyperlinks(pkg, part, el)
    if links:
        rec['links'] = links
    if kind in ('picture', 'video', 'audio'):
        picture_record(pkg, part, el, rec)
    # picture/gradient fills inside shapes are content when they hold an image
    if kind in ('autoshape', 'textbox'):
        bf = el.find(P + 'spPr/' + A + 'blipFill')
        if bf is not None:
            img = blip_part(pkg, part, el)
            if img:
                rec['fill_image_sha1'] = sha1(pkg.blob(img))
    if kind == 'table':
        rows = []
        for tr in el.iter(A + 'tr'):
            rows.append([[paragraph_text(p) for p in tc.iter(A + 'p')] for tc in tr.findall(A + 'tc')])
        rec['table'] = rows
        rec['table_merge'] = [dict(tc.attrib) for tr in el.iter(A + 'tr') for tc in tr.findall(A + 'tc')
                              if set(tc.attrib) & {'gridSpan', 'rowSpan', 'hMerge', 'vMerge'}]
    if kind == 'chart':
        cref = el.find('.//{%s}chart' % NS['c'])
        rel = pkg.rel_target(part, cref.get(R + 'id')) if cref is not None else None
        if rel and pkg.has(rel['abs']):
            rec['chart'] = chart_data(pkg, rel['abs'])
    if kind == 'smartart':
        rec['smartart'] = smartart_text(pkg, part, el)
    if kind == 'ole':
        ole = el.find('.//' + P + 'oleObj')
        if ole is not None:
            rec['ole'] = dict(progId=ole.get('progId'))
            rid = ole.get(R + 'id')
            rel = pkg.rel_target(part, rid) if rid else None
            if rel and rel['abs'] and pkg.has(rel['abs']):
                rec['ole']['sha1'] = sha1(pkg.blob(rel['abs']))
    return rec


def timing_record(root):
    timing = root.find(P + 'timing')
    if timing is None:
        return None
    effects = []
    for ctn in timing.iter(P + 'cTn'):
        if ctn.get('presetClass'):
            tgt = ctn.find('.//' + P + 'spTgt')
            effects.append(dict(cls=ctn.get('presetClass'), preset=ctn.get('presetID'),
                                sub=ctn.get('presetSubtype'), node=ctn.get('nodeType'),
                                spid=tgt.get('spid') if tgt is not None else None,
                                delay=ctn.find(P + 'stCondLst/' + P + 'cond').get('delay')
                                if ctn.find(P + 'stCondLst/' + P + 'cond') is not None else None,
                                dur=ctn.get('dur')))
    media = [dict(spid=m.find('.//' + P + 'spTgt').get('spid') if m.find('.//' + P + 'spTgt') is not None else None,
                  vol=m.get('vol'), mute=m.get('mute'), showWhenStopped=m.get('showWhenStopped'))
             for m in timing.iter(P + 'video', P + 'audio')]
    return dict(xml_sha1=sha1(c14n_str(timing).encode()), effects=effects, media=media)


def transition_record(root):
    found = []
    for t in root.iter(P + 'transition'):
        found.append(c14n_str(t))
    return sha1(''.join(found).encode()) if found else None


def sections(pkg):
    pres = pkg.xml(pkg.main_part)
    out = []
    for sec in pres.iter(P14 + 'section'):
        out.append(dict(name=sec.get('name'),
                        slide_ids=[s.get('id') for s in sec.iter(P14 + 'sldId')]))
    return out


def notes_text(pkg, slide_part):
    from dglcore import RT_NOTES
    rs = pkg.rels_of_type(slide_part, RT_NOTES)
    if not rs or not pkg.has(rs[0]['abs']):
        return []
    root = pkg.xml(rs[0]['abs'])
    for el, _ in iter_shapes(sp_tree(root)):
        ph = ph_of(el)
        if ph and ph[0] == 'body':
            tb = el.find(P + 'txBody')
            return [paragraph_text(p) for p in tb.findall(A + 'p')] if tb is not None else []
    return []


def fingerprint(path):
    pkg = open_package(path)
    pres = pkg.xml(pkg.main_part)
    sld_ids = [s.get('id') for s in pres.iter(P + 'sldId')]
    slides = []
    for n, part in enumerate(pkg.slide_parts(), 1):
        root = pkg.xml(part)
        shapes = [shape_record(pkg, part, el) for el, _ in iter_shapes(sp_tree(root))]
        bg = root.find(P + 'cSld/' + P + 'bg')
        bg_img = None
        if bg is not None:
            blip = bg.find('.//' + A + 'blip')
            if blip is not None and blip.get(R + 'embed'):
                rel = pkg.rel_target(part, blip.get(R + 'embed'))
                if rel and pkg.has(rel['abs']):
                    bg_img = sha1(pkg.blob(rel['abs']))
        slides.append(dict(n=n, sld_id=sld_ids[n - 1] if n - 1 < len(sld_ids) else None,
                           hidden=root.get('show') == '0', shapes=shapes,
                           background_image_sha1=bg_img,
                           timing=timing_record(root), transition=transition_record(root),
                           notes=notes_text(pkg, part)))
    return dict(file=str(path), slide_count=len(slides), sections=sections(pkg), slides=slides)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    dump_json(fingerprint(sys.argv[1]), sys.argv[2])
    print('fingerprint written:', sys.argv[2])
