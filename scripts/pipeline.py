#!/usr/bin/env python3
"""Run the whole restyle for one target (auto mode, or after the user's decisions).

    python pipeline.py truth.pptx target.pptx --work WORKDIR
        [--spec truth_spec.json] [--map layout_map.json] [--approved approved.json]
        [--roles roles.json] [--groups groups.json] [--plan plan.json]
        [--allow-low-confidence] [--no-render]

Steps (each one is also a script you can run alone):
  extract_truth (if no --spec) -> fingerprint (before) -> map_layouts (if no --map)
  -> transplant_design -> remap_placeholders -> apply_truth_positions
  -> strip_overrides -> place
  -> visual_groups (confirmed --groups, else only high-confidence proposals)
  -> fit_text -> validate (pptx skill) -> fingerprint + compare (content lock)
  -> check_geometry -> check_media -> render + band_diff -> before_after -> report.md

Exit codes: 0 = delivered, every check passed; 2 = stopped for a decision (low
confidence layout, text that won't fit, media below its floor) — the report
lists the questions; 1 = a check failed (fix, then run again).
"""
import argparse
import json
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def run(step, args, log, allow_fail=False):
    cmd = [sys.executable, os.path.join(HERE, step)] + [str(x) for x in args]
    p = subprocess.run(cmd, capture_output=True, text=True)
    out = (p.stdout + p.stderr).strip()
    log.append(f'## {step}\n```\n{out}\n```\n')
    print(f'[{step}] exit {p.returncode}')
    if p.returncode and not allow_fail:
        print(out)
        sys.exit(f'{step} failed — see above')
    return p.returncode, out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('truth')
    ap.add_argument('target')
    ap.add_argument('--work', required=True)
    ap.add_argument('--spec')
    ap.add_argument('--map')
    ap.add_argument('--approved')
    ap.add_argument('--roles')
    ap.add_argument('--groups')
    ap.add_argument('--plan')
    ap.add_argument('--allow-low-confidence', action='store_true')
    ap.add_argument('--no-render', action='store_true')
    a = ap.parse_args()

    os.makedirs(a.work, exist_ok=True)
    name = os.path.splitext(os.path.basename(a.target))[0]
    W = os.path.join(a.work, name)
    os.makedirs(W, exist_ok=True)
    log = []
    j = lambda f: os.path.join(W, f)  # noqa: E731

    original = j('original.pptx')
    shutil.copyfile(a.target, original)  # never touch the upload
    spec = a.spec or os.path.join(a.work, 'truth_spec.json')
    if not a.spec:
        run('extract_truth.py', [a.truth, spec, '--summary', os.path.join(a.work, 'truth_summary.md')], log)
    run('fingerprint.py', [original, j('before_fingerprint.json')], log)

    lmap = a.map or j('layout_map.json')
    if not a.map:
        run('map_layouts.py', [spec, original, lmap], log)
    m = json.load(open(lmap))
    questions = []
    low = [(n, s) for n, s in m['slides'].items() if s['confidence'] in ('low', 'none')]
    if low and not a.allow_low_confidence:
        for n, s in low:
            questions.append(f"Slide {n} ({s['looks_like']}): which truth layout? Best guess {s['layout']}, "
                             f"alternatives {s['alternatives']} — or keep its arrangement and apply only truth "
                             "margins, fonts, colors and fixed elements.")
        write_report(W, name, log, questions, stopped=True)
        print('\n'.join(questions))
        sys.exit(2)

    s1, s2, s3, s4, s5, s6 = (j(f'step{i}.pptx') for i in range(1, 7))
    run('transplant_design.py', [a.truth, original, lmap, s1, '--spec', spec], log)
    remap = [s1, a.truth, spec, s2] + (['--approved', a.approved] if a.approved else [])
    run('remap_placeholders.py', remap, log)
    s2b = j('step2b.pptx')
    run('apply_truth_positions.py', [s2, a.truth, spec, s2b], log)
    run('strip_overrides.py', [s2b, spec, s3] + (['--roles', a.roles] if a.roles else []), log)
    run('place.py', [s3, spec, s4] + (['--plan', a.plan] if a.plan else []), log)
    groups = a.groups
    if not groups:
        run('visual_groups.py', ['detect', s4, j('groups_proposed.json'), '--sheets', j('group_sheets')], log)
        prop = json.load(open(j('groups_proposed.json')))
        auto = dict(groups=[g for g in prop['groups'] if g['confidence'] == 'high'])
        json.dump(auto, open(j('groups_applied_auto.json'), 'w'), indent=1)
        groups = j('groups_applied_auto.json')
        for g in prop['groups']:
            if g['confidence'] != 'high':
                questions.append(f"Visual group {g['id']} ({g['kind']}: " + ', '.join(
                    f"slide {x['slide']} '{x['name']}'" for x in g['members']) +
                    f") — make them the same size and aligned? Sheet: {g['sheet']}")
    run('visual_groups.py', ['apply', s4, spec, groups, s5], log)
    run('fit_text.py', [s5, s6, '--report', j('fit.json')], log)
    final = os.path.join(W, f'{name} - restyled.pptx')
    shutil.copyfile(s6, final)
    shutil.copyfile(s6 + '.meta.json', final + '.meta.json')

    meta = json.load(open(final + '.meta.json'))
    for s in meta.get('fit', {}).get('needs_decision', []):
        questions.append(f"Slide {s['slide']} '{s['name']}' doesn't fit even at {s['floor_pt']} pt "
                         "(85% of the truth size): split the slide in two, allow a smaller size this once, "
                         "or trim the text yourself?")
    for s in meta.get('place', {}).get('needs_approval', []):
        questions.append(f"Slide {s['slide']} '{s['name']}' (media) would shrink to {s['scale']:.0%}: allow?")
    for s in meta.get('remap', {}).get('kept_footer_text', []):
        questions.append(f"Slide {s['slide']} keeps its own footer text '{s['text']}': replace it with the "
                         "truth footer text?")
    for s in meta.get('style', {}).get('backgrounds_kept', []):
        questions.append(f"Slide {s['slide']} has a picture background (content): keep it?")

    results = {}
    from render import pptx_skill_dir
    skill = pptx_skill_dir()
    if skill:
        results['validate'] = run_val = subprocess.run(
            [sys.executable, os.path.join(skill, 'scripts', 'office', 'validate.py'), final, '--original', original],
            capture_output=True, text=True)
        log.append(f"## validate.py\n```\n{run_val.stdout}{run_val.stderr}\n```\n")
        results['validate'] = run_val.returncode
    results['content'] = run('fingerprint.py', [final, j('after_fingerprint.json')], log)[0] or run(
        'compare_fingerprint.py', [j('before_fingerprint.json'), j('after_fingerprint.json'), '--json',
                                   j('content_check.json')] + (['--approved', a.approved] if a.approved else []),
        log, allow_fail=True)[0]
    results['geometry'] = run('check_geometry.py', [final, a.truth, spec, '--json', j('geometry_check.json')],
                              log, allow_fail=True)[0]
    results['media'] = run('check_media.py', [original, final, '--json', j('media_check.json')] +
                           (['--approved', a.approved] if a.approved else []), log, allow_fail=True)[0]
    if not a.no_render and shutil.which('soffice'):
        R = j('renders')
        os.makedirs(R, exist_ok=True)
        run('render.py', [a.truth, os.path.join(R, 'truth')], log)
        run('render.py', [original, os.path.join(R, 'before')], log)
        run('render.py', [final, os.path.join(R, 'after')], log)
        results['pixels'] = run('band_diff.py', [spec, final, os.path.join(R, 'truth'), os.path.join(R, 'after'),
                                                 '--out', j('band_diffs'), '--json', j('pixel_check.json')],
                                log, allow_fail=True)[0]
        run('before_after.py', [os.path.join(R, 'before'), os.path.join(R, 'after'),
                                os.path.join(W, f'{name} - before-after.pdf'), '--labels', final + '.meta.json',
                                '--sheets', os.path.join(R, 'sheet')], log)
    write_report(W, name, log, questions, stopped=False, results=results, meta=meta)
    failed = [k for k, v in results.items() if v]
    # overflow that is waiting on the user shows up as a pixel/geometry failure too
    print(f"\nreport: {os.path.join(W, 'report.md')}")
    if failed and not questions:
        print('FAILED checks:', ', '.join(failed))
        sys.exit(1)
    if questions:
        print(f'{len(questions)} question(s) for the user — see the report')
        sys.exit(2)
    print('all checks passed')


def write_report(W, name, log, questions, stopped, results=None, meta=None):
    L = [f'# Restyle report — {name}', '']
    if stopped:
        L += ['**Stopped before restyling: the layout map needs the user.**', '']
    if results:
        L.append('## Checks')
        files = dict(content='content_check.json', geometry='geometry_check.json', media='media_check.json',
                     pixels='pixel_check.json')
        for k, v in results.items():
            L.append(f"- {k}: {'PASS' if not v else 'FAIL'}")
            path = os.path.join(W, files.get(k, '-'))
            if v and os.path.exists(path):
                d = json.load(open(path))
                for f in d.get('fails', [])[:15]:
                    if isinstance(f, dict):
                        f = f"slide {f['slide']} '{f['element']}' differs from the truth render (score {f['score']})"
                    L.append(f'  - {f}')
        L.append('')
    if meta:
        L.append('## Slide → truth layout')
        for n, k in sorted(meta.get('layout_for_slide', {}).items(), key=lambda kv: int(kv[0])):
            L.append(f'- slide {n}: {k}')
        fit = meta.get('fit', {})
        L += ['', '## Text shrinks']
        L += [f"- slide {s['slide']} '{s['name']}': {s['from_pt']} → {s['to_pt']} pt"
              + (f", line spacing {s['line_spacing']}%" if s.get('line_spacing') else '')
              for s in fit.get('shrinks', [])] or ['- none']
        st = meta.get('style', {})
        L += ['', '## Kept as content']
        L += [f"- slide {s['slide']} '{s['name']}': colored text {s.get('text')!r}"
              for s in st.get('semantic_color_kept', [])] or ['- nothing special']
        rm = meta.get('remap', {})
        if rm.get('removed'):
            L += ['', '## Removed with approval'] + [f"- slide {r['slide']} '{r['name']}'" for r in rm['removed']]
        vg = meta.get('visual_groups', {})
        if vg.get('applied'):
            L += ['', '## Visual groups made identical']
            for g in vg['applied']:
                sl = sorted({m['slide'] for m in g['members']})
                L.append(f"- {g['kind']} of {len(g['members'])} on slide(s) {sl}: "
                         f"{g['box_size'][0] / 914400:.2f} × {g['box_size'][1] / 914400:.2f} in")
            L += [f'- note: {n}' for n in vg.get('notes', [])]
        if fit.get('approximate_fonts'):
            L += ['', '## Fonts not installed here',
                  '- ' + ', '.join(fit['approximate_fonts']) + ': renders use substitutes; text fit estimated']
        L += ['', '## Media (GIF, video, audio) and animations',
              '- content, animations and transitions verified unchanged; placement only (see media check)',
              '- renders show only the poster frame / first GIF frame']
    if questions:
        L += ['', '## Waiting on the user'] + [f'- {q}' for q in questions]
    L += ['', '---', '# Log', ''] + log
    open(os.path.join(W, 'report.md'), 'w').write('\n'.join(L) + '\n')


if __name__ == '__main__':
    main()
