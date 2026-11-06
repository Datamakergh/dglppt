#!/usr/bin/env python3
"""Check the environment before a dglppt session.

    python preflight.py [truth.pptx]

Required: python-pptx, lxml, Pillow. Optional (with fallbacks):
  LibreOffice + pdftoppm   renders for the visual check and before/after
                           (without them: skip 5c, say so, deliver the report only)
  fc-match / fc-list       real font metrics for text fit (without: estimates)
  pptx skill validate.py   OOXML schema validation (without: the LibreOffice
                           round-trip in step 4f is the only open check)
With a truth file it also lists the truth fonts that aren't installed: renders
will substitute them, so text fit in the preview can't be trusted for those.
"""
import importlib
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def ok(flag):
    return 'ok     ' if flag else 'MISSING'


def main():
    missing_required = False
    print('Python modules')
    for mod, req in (('pptx', True), ('lxml', True), ('PIL', True), ('defusedxml', False)):
        try:
            importlib.import_module(mod)
            print(f'  {ok(True)} {mod}')
        except ImportError:
            print(f"  {ok(False)} {mod} {'(required: pip install python-pptx lxml Pillow)' if req else '(optional)'}")
            missing_required |= req
    print('Tools')
    for exe, why in (('soffice', 'renders (visual check, before/after)'), ('pdftoppm', 'PDF -> images'),
                     ('fc-match', 'font metrics for text fit'), ('fc-list', 'installed font list')):
        print(f'  {ok(shutil.which(exe))} {exe:9} {why}')
    from render import pptx_skill_dir
    skill = pptx_skill_dir()
    v = os.path.join(skill, 'scripts', 'office', 'validate.py') if skill else None
    print(f"  {ok(bool(v))} pptx skill validate.py {v or '(set DGLPPT_PPTX_SKILL; fallback: LibreOffice round-trip)'}")
    if len(sys.argv) > 1:
        from dglcore import Theme, open_package, A
        pkg = open_package(sys.argv[1])
        fonts = set()
        for m in pkg.master_parts():
            t = pkg.theme_of(m)
            if t:
                th = Theme(pkg.xml(t))
                fonts |= {th.major, th.minor}
        for name in pkg.names():
            if name.startswith('ppt/') and name.endswith('.xml') and ('slide' in name or 'theme' in name):
                for lat in pkg.xml(name).iter(A + 'latin'):
                    tf = lat.get('typeface', '')
                    if tf and not tf.startswith('+'):
                        fonts.add(tf)
        installed = set()
        if shutil.which('fc-list'):
            out = subprocess.run(['fc-list', ':', 'family'], capture_output=True, text=True).stdout
            for line in out.splitlines():
                installed |= {f.strip().lower() for f in line.split(',')}
        print('Truth fonts')
        for f in sorted(x for x in fonts if x):
            print(f"  {ok(f.lower() in installed)} {f}" + ('' if f.lower() in installed else
                  '  -> renders substitute it; ask the user for the font files, or trust geometry checks only'))
    sys.exit(1 if missing_required else 0)


if __name__ == '__main__':
    main()
