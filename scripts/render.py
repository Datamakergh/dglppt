#!/usr/bin/env python3
"""Render a deck to one JPG per slide (LibreOffice -> PDF -> pdftoppm).

    python render.py deck.pptx OUT_PREFIX [--dpi 150]

Writes OUT_PREFIX-1.jpg, OUT_PREFIX-2.jpg, … and prints the list.
Uses the pptx skill's soffice wrapper when it can find it (it works around
sandboxes that block AF_UNIX sockets); otherwise runs soffice directly with a
throwaway profile. Set DGLPPT_PPTX_SKILL to the pptx skill folder if needed.
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

CANDIDATES = [os.environ.get('DGLPPT_PPTX_SKILL', ''), '/mnt/skills/public/pptx',
              os.path.expanduser('~/.claude/skills/pptx'), '/mnt/user-data/skills/pptx']


def pptx_skill_dir():
    for c in CANDIDATES:
        if c and os.path.isfile(os.path.join(c, 'scripts', 'office', 'soffice.py')):
            return c
    return None


def run_soffice(args, **kw):
    skill = pptx_skill_dir()
    if skill:
        sys.path.insert(0, os.path.join(skill, 'scripts'))
        try:
            from office.soffice import run_soffice as rs  # type: ignore
            return rs(args, **kw)
        except Exception:
            pass
    exe = shutil.which('soffice') or shutil.which('libreoffice')
    if not exe:
        sys.exit('LibreOffice (soffice) is not installed: rendering is unavailable. '
                 'Skip the visual check and say so in the report.')
    with tempfile.TemporaryDirectory(prefix='lo_profile_') as prof:
        env = dict(os.environ, SAL_USE_VCLPLUGIN='svp')
        return subprocess.run([exe, f'-env:UserInstallation=file://{prof}'] + list(args), env=env, **kw)


def to_pdf(deck, outdir):
    src = deck
    if deck.lower().endswith(('.potx', '.ppsx')):
        src = os.path.join(outdir, os.path.splitext(os.path.basename(deck))[0] + '.pptx')
        sys.path.insert(0, os.path.dirname(__file__))
        from dglcore import as_pptx_bytes
        open(src, 'wb').write(as_pptx_bytes(open(deck, 'rb').read()))
    run_soffice(['--headless', '--convert-to', 'pdf', '--outdir', outdir, src],
                check=True, capture_output=True, timeout=300)
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + '.pdf')
    if not os.path.exists(pdf):
        sys.exit('LibreOffice produced no PDF')
    return pdf


def render(deck, prefix, dpi=150):
    with tempfile.TemporaryDirectory() as tmp:
        pdf = to_pdf(os.path.abspath(deck), tmp)
        for old in glob.glob(prefix + '-*.jpg'):
            os.remove(old)
        subprocess.run(['pdftoppm', '-jpeg', '-r', str(dpi), pdf, prefix], check=True)
    files = sorted(glob.glob(prefix + '-*.jpg'), key=lambda f: int(f.rsplit('-', 1)[1].split('.')[0]))
    # pdftoppm zero-pads (-01) for decks with 10+ pages; normalise to -1, -2, …
    out = []
    for i, f in enumerate(files, 1):
        want = f'{prefix}-{i}.jpg'
        if f != want:
            os.replace(f, want)
        out.append(want)
    return out


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('deck')
    ap.add_argument('prefix')
    ap.add_argument('--dpi', type=int, default=150)
    a = ap.parse_args()
    for f in render(a.deck, a.prefix, a.dpi):
        print(f)
