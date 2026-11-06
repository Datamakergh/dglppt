#!/usr/bin/env python3
"""Compare two content fingerprints and print a readable diff.

    python compare_fingerprint.py before.json after.json [--approved approved.json] [--json out.json]

Exit code 0 = content identical (apart from approved changes), 1 = content changed.

approved.json (optional) lists what the user explicitly allowed in chat:
{
  "remove": [{"slide": 3, "shape_id": 7, "why": "old footer 'Confidential'"}],
  "change": [{"slide": 5, "shape_id": 4, "field": "text", "why": "user fixed typo"}],
  "add":    [{"slide": 2, "name": "Picture 9", "why": "split slide"}]
}
Shapes added by the restyle as truth furniture (footer / slide number / date
placeholders, or shapes named "dglppt-fixed ...") are always allowed and listed.
"""
import argparse
import json
import sys

IGNORED_SHAPE_KEYS = {'ph', 'name'}  # remapping placeholders may change these; not content


def key_desc(slide, sh):
    return f"slide {slide}, shape {sh.get('id')} ({sh.get('kind')} '{sh.get('name', '')}')"


def short(v, n=120):
    s = json.dumps(v, ensure_ascii=False)
    return s if len(s) <= n else s[:n] + '…'


def diff_value(path, a, b, out):
    if a == b:
        return
    if isinstance(a, dict) and isinstance(b, dict):
        for k in sorted(set(a) | set(b)):
            diff_value(f'{path}.{k}', a.get(k), b.get(k), out)
        return
    if isinstance(a, list) and isinstance(b, list) and len(a) == len(b):
        for i, (x, y) in enumerate(zip(a, b)):
            diff_value(f'{path}[{i}]', x, y, out)
        return
    out.append(f'{path}: {short(a)} -> {short(b)}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('before')
    ap.add_argument('after')
    ap.add_argument('--approved')
    ap.add_argument('--json')
    a = ap.parse_args()
    before = json.load(open(a.before, encoding='utf-8'))
    after = json.load(open(a.after, encoding='utf-8'))
    appr = json.load(open(a.approved, encoding='utf-8')) if a.approved else {}
    removals = {(x['slide'], x['shape_id']) for x in appr.get('remove', [])}
    changes = {(x['slide'], x['shape_id'], x.get('field')) for x in appr.get('change', [])}
    adds = {(x['slide'], x.get('name')) for x in appr.get('add', [])}

    fails, infos = [], []
    if before['slide_count'] != after['slide_count']:
        fails.append(f"slide count {before['slide_count']} -> {after['slide_count']}")
    if [s['sld_id'] for s in before['slides']] != [s['sld_id'] for s in after['slides']]:
        fails.append('slide order or slide identity changed')
    if before.get('sections') != after.get('sections'):
        fails.append(f"sections changed: {short(before.get('sections'))} -> {short(after.get('sections'))}")

    for sb, sa in zip(before['slides'], after['slides']):
        n = sb['n']
        for key in ('hidden', 'notes', 'background_image_sha1'):
            if sb.get(key) != sa.get(key):
                fails.append(f'slide {n}: {key} changed: {short(sb.get(key))} -> {short(sa.get(key))}')
        if sb.get('timing') != sa.get('timing'):
            out = []
            diff_value('timing', sb.get('timing'), sa.get('timing'), out)
            fails.extend(f'slide {n}: animation {d}' for d in out)
        if sb.get('transition') != sa.get('transition'):
            fails.append(f'slide {n}: transition changed')

        bmap = {s['id']: s for s in sb['shapes']}
        amap = {s['id']: s for s in sa['shapes']}
        for sid, shb in bmap.items():
            if sid not in amap:
                if (n, sid) in removals:
                    infos.append(f'approved removal: {key_desc(n, shb)} text={short(shb.get("text"))}')
                else:
                    fails.append(f'MISSING {key_desc(n, shb)} text={short(shb.get("text"))}')
                continue
            sha = amap[sid]
            for k in sorted((set(shb) | set(sha)) - IGNORED_SHAPE_KEYS):
                if shb.get(k) == sha.get(k):
                    continue
                if (n, sid, k) in changes or (n, sid, None) in changes:
                    infos.append(f'approved change: {key_desc(n, shb)} field {k}')
                    continue
                out = []
                diff_value(k, shb.get(k), sha.get(k), out)
                fails.extend(f'CHANGED {key_desc(n, shb)} {d}' for d in out)
        for sid, sha in amap.items():
            if sid in bmap:
                continue
            ph = sha.get('ph') or [None]
            if ph[0] in ('ftr', 'sldNum', 'dt', 'hdr') or sha.get('name', '').startswith('dglppt-fixed'):
                infos.append(f'truth furniture added: {key_desc(n, sha)}')
            elif (n, sha.get('name')) in adds:
                infos.append(f'approved addition: {key_desc(n, sha)}')
            else:
                fails.append(f'ADDED {key_desc(n, sha)} text={short(sha.get("text"))}')

    print('CONTENT CHECK:', 'PASS' if not fails else 'FAIL')
    for f in fails:
        print('  FAIL ', f)
    for i in infos:
        print('  info ', i)
    if a.json:
        json.dump(dict(passed=not fails, fails=fails, infos=infos), open(a.json, 'w'), indent=1,
                  ensure_ascii=False)
    sys.exit(0 if not fails else 1)


if __name__ == '__main__':
    main()
