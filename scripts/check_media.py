#!/usr/bin/env python3
"""Media and animation integrity check (GIF, video, audio, animations, transitions).

    python check_media.py original.pptx restyled.pptx [--approved approved.json] [--json out.json]

Exit 1 on any FAIL. For every slide (shapes matched by id — the restyle edits
shapes in place, so ids survive):
  * every video/audio/GIF in the original is still there with the same bytes,
    poster frame, link/embed, trim/fade/bookmarks (p14:media) and playback
    settings; GIFs keep their frame count and frame durations
  * a media frame keeps its aspect ratio (±0.5%); shrinking it below 85% of its
    original size needs the user's approval ("media_scale" in approved.json)
  * the animation timeline (p:timing) and the transition are unchanged, and
    every animation target (spTgt spid) still exists and is the same shape
  * every media relationship resolves to a part that exists, and every media
    extension has a content type
"""
import argparse
import json
import sys

from dglcore import (P, P14, Package, as_pptx_bytes, blip_part, c14n_str, image_info, iter_shapes,
                     load_json, media_parts, nv_pr, read_xfrm, sha1, shape_id, shape_kind, shape_name,
                     sp_tree)


def media_shapes(pkg):
    out = {}
    for n, part in enumerate(pkg.slide_parts(), 1):
        root = pkg.xml(part)
        names = {}
        for el, _ in iter_shapes(sp_tree(root)):
            sid = shape_id(el)
            names[sid] = shape_name(el)
            kind = shape_kind(el)
            if kind not in ('video', 'audio', 'picture'):
                continue
            img = blip_part(pkg, part, el)
            info = image_info(pkg.blob(img)) if img else None
            is_gif = bool(info and info['frames'] > 1)
            if kind == 'picture' and not is_gif:
                continue
            rec = dict(kind='gif' if is_gif else kind, name=shape_name(el), box=read_xfrm(el),
                       poster=sha1(pkg.blob(img)) if img else None,
                       frames=info['frames'] if info else None,
                       durations=info['durations'] if info else None, media=[])
            for k, tgt, external in media_parts(pkg, part, el):
                if external:
                    rec['media'].append(dict(kind=k, link=tgt))
                else:
                    rec['media'].append(dict(kind=k, part=tgt, exists=pkg.has(tgt),
                                             sha1=sha1(pkg.blob(tgt)) if pkg.has(tgt) else None,
                                             ctype=pkg.content_type(tgt) if pkg.has(tgt) else None))
            nvpr = nv_pr(el)
            m = nvpr.find('.//' + P14 + 'media') if nvpr is not None else None
            rec['settings'] = c14n_str(m) if m is not None else None
            out[(n, sid)] = rec
        timing = root.find(P + 'timing')
        out[(n, 'timing')] = dict(xml=c14n_str(timing) if timing is not None else None,
                                  spids=sorted({int(t.get('spid')) for t in root.iter(P + 'spTgt')
                                                if t.get('spid', '').isdigit()}),
                                  names=names,
                                  transition=sha1(''.join(c14n_str(t) for t in root.iter(P + 'transition')).encode()))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('original')
    ap.add_argument('restyled')
    ap.add_argument('--approved')
    ap.add_argument('--json')
    a = ap.parse_args()
    before = media_shapes(Package(as_pptx_bytes(open(a.original, 'rb').read())))
    after = media_shapes(Package(as_pptx_bytes(open(a.restyled, 'rb').read())))
    appr = load_json(a.approved) if a.approved else {}
    ok_scale = {(x['slide'], x['shape_id']) for x in appr.get('media_scale', [])}
    fails, infos = [], []

    for key, b in before.items():
        n, sid = key
        if sid == 'timing':
            t = after.get(key)
            if t is None:
                fails.append(f'slide {n}: slide missing')
                continue
            if b['xml'] != t['xml']:
                fails.append(f'slide {n}: animation timeline changed')
            if b['transition'] != t['transition']:
                fails.append(f'slide {n}: transition changed')
            for spid in t['spids']:
                if spid not in t['names']:
                    fails.append(f'slide {n}: animation targets shape {spid}, which no longer exists')
                elif b['names'].get(spid) != t['names'][spid]:
                    fails.append(f"slide {n}: animation target {spid} is now '{t['names'][spid]}', "
                                 f"was '{b['names'].get(spid)}'")
            if t['spids']:
                infos.append(f"slide {n}: {len(t['spids'])} animation targets intact")
            continue
        aft = after.get(key)
        where = f"slide {n} '{b['name']}' ({b['kind']})"
        if aft is None:
            still_there = sid in (after.get((n, 'timing')) or {}).get('names', {})
            fails.append(f'{where}: ' + ('is no longer animated / no longer media (file re-encoded or replaced)'
                                         if still_there else 'missing after restyle'))
            continue
        for k in ('kind', 'poster', 'frames', 'durations', 'settings'):
            if b[k] != aft[k]:
                fails.append(f'{where}: {k} changed')
        bm = [(m['kind'], m.get('sha1') or m.get('link')) for m in b['media']]
        am = [(m['kind'], m.get('sha1') or m.get('link')) for m in aft['media']]
        if bm != am:
            fails.append(f'{where}: media file or link changed')
        for m in aft['media']:
            if 'part' in m and not m['exists']:
                fails.append(f"{where}: media part {m['part']} is missing")
            elif 'part' in m and not m['ctype']:
                fails.append(f"{where}: no content type for {m['part']}")
        if b['box'] and aft['box']:
            ra, rb = b['box'][2] / b['box'][3], aft['box'][2] / aft['box'][3]
            if abs(ra - rb) / ra > 0.005:
                fails.append(f'{where}: aspect ratio changed {ra:.4f} -> {rb:.4f}')
            scale = aft['box'][2] / b['box'][2]
            if scale < 0.85 and key not in ok_scale:
                fails.append(f'{where}: shrunk to {scale:.0%} of its size without approval')
            infos.append(f'{where}: placement {b["box"]} -> {aft["box"]} (scale {scale:.2f}); content unchanged')

    print('MEDIA CHECK:', 'PASS' if not fails else 'FAIL')
    for f in fails:
        print('  FAIL', f)
    for i in infos:
        print('  info', i)
    if a.json:
        json.dump(dict(passed=not fails, fails=fails, infos=infos), open(a.json, 'w'), indent=1)
    sys.exit(0 if not fails else 1)


if __name__ == '__main__':
    main()
