"""dglcore — shared helpers for the dglppt scripts.

Package-level (OPC) access with lxml, .potx-safe python-pptx loading, an
inheritance-aware text-style resolver, geometry helpers and small image tools.
Every other script imports this module; keep it dependency-light
(python-pptx, lxml, Pillow).
"""
from __future__ import annotations

import colorsys
import copy
import hashlib
import io
import json
import posixpath
import re
import zipfile
from dataclasses import dataclass

from lxml import etree

# --------------------------------------------------------------------------
# Constants and namespaces
# --------------------------------------------------------------------------
EMU_PER_IN = 914400
EMU_PER_PT = 12700
EMU_PER_CM = 360000

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
    'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
A = '{%s}' % NS['a']
P = '{%s}' % NS['p']
R = '{%s}' % NS['r']
C = '{%s}' % NS['c']
P14 = '{%s}' % NS['p14']
PR = '{%s}' % NS['pr']
CT = '{%s}' % NS['ct']

RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
RT_SLIDE = RT + 'slide'
RT_LAYOUT = RT + 'slideLayout'
RT_MASTER = RT + 'slideMaster'
RT_THEME = RT + 'theme'
RT_IMAGE = RT + 'image'
RT_NOTES = RT + 'notesSlide'
RT_CHART = RT + 'chart'
RT_DGM_DATA = RT + 'diagramData'
RT_TABLESTYLES = RT + 'tableStyles'
RT_VIDEO = RT + 'video'
RT_AUDIO = RT + 'audio'
RT_MEDIA = 'http://schemas.microsoft.com/office/2007/relationships/media'
RT_OLE = RT + 'oleObject'
RT_PACKAGE = RT + 'package'

CT_PRES_MAIN = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
CT_TEMPLATE_MAIN = 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'
CT_SLIDESHOW_MAIN = 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml'

PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True,
                         remove_blank_text=False)

# Placeholder types whose text is the deck's fixed furniture (shown per slide).
FIXED_PH_TYPES = {'ftr', 'sldNum', 'dt', 'hdr'}
TITLE_PH_TYPES = {'title', 'ctrTitle'}


def q(tag: str) -> str:
    """'a:rPr' -> Clark notation."""
    pfx, local = tag.split(':')
    return '{%s}%s' % (NS[pfx], local)


def parse_xml(data: bytes):
    return etree.fromstring(data, PARSER)


def to_bytes(el) -> bytes:
    return etree.tostring(el, xml_declaration=True, encoding='UTF-8', standalone=True)


def c14n_str(el) -> str:
    """Stable string for comparing XML subtrees."""
    if el is None:
        return ''
    return etree.tostring(el, method='c14n').decode('utf-8')


def sha1(data: bytes) -> str:
    return hashlib.sha1(data).hexdigest()


def emu_to_in(v):
    return None if v is None else round(v / EMU_PER_IN, 3)


def emu_to_cm(v):
    return None if v is None else round(v / EMU_PER_CM, 2)


def dump_json(obj, path):
    with open(path, 'w', encoding='utf-8') as fh:
        json.dump(obj, fh, indent=1, ensure_ascii=False, default=str)


def load_json(path):
    with open(path, encoding='utf-8') as fh:
        return json.load(fh)


# --------------------------------------------------------------------------
# OPC package
# --------------------------------------------------------------------------
class Package:
    """A .pptx/.potx opened at the part level.

    Parts are kept as bytes; XML parts parsed through .xml() are cached and
    written back on save(), so edits to the returned tree persist.
    """

    def __init__(self, path_or_bytes):
        if isinstance(path_or_bytes, (bytes, bytearray)):
            zf = zipfile.ZipFile(io.BytesIO(path_or_bytes))
        else:
            zf = zipfile.ZipFile(path_or_bytes)
        self.parts: dict[str, bytes] = {}
        for info in zf.infolist():
            if info.is_dir():
                continue
            self.parts[info.filename] = zf.read(info.filename)
        zf.close()
        self._trees: dict[str, object] = {}

    # ---- raw access -------------------------------------------------------
    def has(self, name):
        return name in self.parts or name in self._trees

    def xml(self, name):
        if name not in self._trees:
            self._trees[name] = parse_xml(self.parts[name])
        return self._trees[name]

    def set_xml(self, name, root):
        self._trees[name] = root
        self.parts[name] = b''  # placeholder; real bytes produced on save

    def blob(self, name) -> bytes:
        if name in self._trees:
            return to_bytes(self._trees[name])
        return self.parts[name]

    def set_blob(self, name, data: bytes):
        self._trees.pop(name, None)
        self.parts[name] = data

    def delete(self, name):
        self.parts.pop(name, None)
        self._trees.pop(name, None)

    def names(self):
        return list(self.parts.keys())

    # ---- relationships ----------------------------------------------------
    @staticmethod
    def rels_name(part: str) -> str:
        d, f = posixpath.split(part)
        return posixpath.join(d, '_rels', f + '.rels') if part else '_rels/.rels'

    def rels(self, part: str):
        """List of dicts: id, type, target (raw), mode, abs (resolved part name or None)."""
        rn = self.rels_name(part)
        if not self.has(rn):
            return []
        out = []
        for rel in self.xml(rn).findall(PR + 'Relationship'):
            mode = rel.get('TargetMode', 'Internal')
            tgt = rel.get('Target')
            absn = None if mode == 'External' else self.resolve(part, tgt)
            out.append(dict(id=rel.get('Id'), type=rel.get('Type'), target=tgt, mode=mode, abs=absn))
        return out

    def rels_root(self, part: str):
        rn = self.rels_name(part)
        if not self.has(rn):
            root = etree.Element(PR + 'Relationships', nsmap={None: NS['pr']})
            self.set_xml(rn, root)
        return self.xml(rn)

    def rel_target(self, part: str, rid: str):
        for r in self.rels(part):
            if r['id'] == rid:
                return r
        return None

    def rels_of_type(self, part, rtype):
        return [r for r in self.rels(part) if r['type'] == rtype]

    def add_rel(self, part: str, rtype: str, target_part: str, external=False) -> str:
        root = self.rels_root(part)
        used = {r.get('Id') for r in root}
        n = 1
        while f'rId{n}' in used:
            n += 1
        rid = f'rId{n}'
        el = etree.SubElement(root, PR + 'Relationship')
        el.set('Id', rid)
        el.set('Type', rtype)
        if external:
            el.set('Target', target_part)
            el.set('TargetMode', 'External')
        else:
            el.set('Target', self.relative(part, target_part))
        return rid

    @staticmethod
    def resolve(source_part: str, target: str) -> str:
        if target.startswith('/'):
            return target.lstrip('/')
        base = posixpath.dirname(source_part)
        return posixpath.normpath(posixpath.join(base, target))

    @staticmethod
    def relative(source_part: str, target_part: str) -> str:
        base = posixpath.dirname(source_part) or '.'
        return posixpath.relpath(target_part, base)

    # ---- content types ----------------------------------------------------
    def content_type(self, name: str):
        ct = self.xml('[Content_Types].xml')
        for o in ct.findall(CT + 'Override'):
            if o.get('PartName').lstrip('/') == name:
                return o.get('ContentType')
        ext = name.rsplit('.', 1)[-1].lower()
        for d in ct.findall(CT + 'Default'):
            if d.get('Extension').lower() == ext:
                return d.get('ContentType')
        return None

    def set_override(self, name: str, ctype: str):
        ct = self.xml('[Content_Types].xml')
        for o in ct.findall(CT + 'Override'):
            if o.get('PartName').lstrip('/') == name:
                o.set('ContentType', ctype)
                return
        el = etree.SubElement(ct, CT + 'Override')
        el.set('PartName', '/' + name)
        el.set('ContentType', ctype)

    def remove_override(self, name: str):
        ct = self.xml('[Content_Types].xml')
        for o in list(ct.findall(CT + 'Override')):
            if o.get('PartName').lstrip('/') == name:
                ct.remove(o)

    def ensure_default(self, ext: str, ctype: str):
        ct = self.xml('[Content_Types].xml')
        for d in ct.findall(CT + 'Default'):
            if d.get('Extension').lower() == ext.lower():
                return
        el = etree.Element(CT + 'Default')
        el.set('Extension', ext)
        el.set('ContentType', ctype)
        # Defaults must precede Overrides
        first_override = ct.find(CT + 'Override')
        if first_override is not None:
            first_override.addprevious(el)
        else:
            ct.append(el)

    def has_default(self, ext: str) -> bool:
        ct = self.xml('[Content_Types].xml')
        return any(d.get('Extension').lower() == ext.lower() for d in ct.findall(CT + 'Default'))

    # ---- presentation structure -------------------------------------------
    @property
    def main_part(self) -> str:
        for r in self.rels(''):
            if r['type'] == RT + 'officeDocument':
                return r['abs']
        return 'ppt/presentation.xml'

    def slide_parts(self):
        pres = self.main_part
        root = self.xml(pres)
        lst = root.find(P + 'sldIdLst')
        out = []
        if lst is None:
            return out
        for sid in lst.findall(P + 'sldId'):
            rel = self.rel_target(pres, sid.get(R + 'id'))
            out.append(rel['abs'])
        return out

    def master_parts(self):
        pres = self.main_part
        root = self.xml(pres)
        lst = root.find(P + 'sldMasterIdLst')
        out = []
        if lst is None:
            return out
        for mid in lst.findall(P + 'sldMasterId'):
            out.append(self.rel_target(pres, mid.get(R + 'id'))['abs'])
        return out

    def layout_parts(self, master):
        return [r['abs'] for r in self.rels_of_type(master, RT_LAYOUT)]

    def layout_of(self, slide):
        rs = self.rels_of_type(slide, RT_LAYOUT)
        return rs[0]['abs'] if rs else None

    def master_of(self, layout):
        rs = self.rels_of_type(layout, RT_MASTER)
        return rs[0]['abs'] if rs else None

    def theme_of(self, master):
        rs = self.rels_of_type(master, RT_THEME)
        return rs[0]['abs'] if rs else None

    def notes_of(self, slide):
        rs = self.rels_of_type(slide, RT_NOTES)
        return rs[0]['abs'] if rs else None

    def slide_size(self):
        sz = self.xml(self.main_part).find(P + 'sldSz')
        return int(sz.get('cx')), int(sz.get('cy'))

    # ---- garbage collection -----------------------------------------------
    def reachable(self):
        seen = set()
        stack = ['']
        while stack:
            part = stack.pop()
            for r in self.rels(part):
                t = r['abs']
                if t and t not in seen and self.has(t):
                    seen.add(t)
                    stack.append(t)
        return seen

    def gc(self):
        """Delete parts no relationship chain reaches. Returns deleted names."""
        keep = self.reachable()
        deleted = []
        for name in list(self.parts):
            if name == '[Content_Types].xml' or name.endswith('.rels'):
                continue
            if name not in keep:
                self.delete(name)
                self.remove_override(name)
                deleted.append(name)
        # rels whose source part is gone
        for name in list(self.parts):
            if name.endswith('.rels') and name != '_rels/.rels':
                d, f = posixpath.split(name)
                src = posixpath.join(posixpath.dirname(d), f[:-5])
                if not self.has(src):
                    self.delete(name)
                    deleted.append(name)
        return deleted

    # ---- save -------------------------------------------------------------
    def save(self, path):
        names = list(self.parts)
        order = ['[Content_Types].xml', '_rels/.rels'] + sorted(
            n for n in names if n not in ('[Content_Types].xml', '_rels/.rels'))
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for n in order:
                if n in self.parts or n in self._trees:
                    zf.writestr(n, self.blob(n))

    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        self.save(buf)
        return buf.getvalue()


def unique_name(pkg: Package, directory: str, stem: str, ext: str, reserved=()):
    """First free 'directory/stemN.ext'."""
    n = 1
    while True:
        cand = f'{directory}/{stem}{n}.{ext}'
        if not pkg.has(cand) and cand not in reserved:
            return cand
        n += 1


# --------------------------------------------------------------------------
# python-pptx loading (handles .potx / .ppsx)
# --------------------------------------------------------------------------
def open_presentation(path):
    """python-pptx Presentation for .pptx, .potx or .ppsx files."""
    from pptx import Presentation
    with open(path, 'rb') as fh:
        data = fh.read()
    ct = zipfile.ZipFile(io.BytesIO(data)).read('[Content_Types].xml')
    if CT_TEMPLATE_MAIN.encode() in ct or CT_SLIDESHOW_MAIN.encode() in ct:
        data = as_pptx_bytes(data)
    return Presentation(io.BytesIO(data))


def as_pptx_bytes(data: bytes) -> bytes:
    """Rewrite the main content type of a .potx/.ppsx so it reads as .pptx."""
    src = zipfile.ZipFile(io.BytesIO(data))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            blob = src.read(info.filename)
            if info.filename == '[Content_Types].xml':
                blob = (blob.replace(CT_TEMPLATE_MAIN.encode(), CT_PRES_MAIN.encode())
                            .replace(CT_SLIDESHOW_MAIN.encode(), CT_PRES_MAIN.encode()))
            out.writestr(info, blob)
    return buf.getvalue()


def open_package(path) -> Package:
    with open(path, 'rb') as fh:
        data = fh.read()
    return Package(as_pptx_bytes(data))


# --------------------------------------------------------------------------
# Shape tree helpers (XML level)
# --------------------------------------------------------------------------
SHAPE_TAGS = {P + 'sp', P + 'pic', P + 'graphicFrame', P + 'grpSp', P + 'cxnSp', P + 'contentPart'}


def sp_tree(root):
    return root.find(P + 'cSld/' + P + 'spTree')


def iter_shapes(container, recurse=True, _depth=0):
    """Yield (shape_el, depth) in document (z) order. Descends into groups and
    mc:AlternateContent (Choice branch first, falls back to Fallback)."""
    for el in container:
        tag = el.tag
        if tag == '{%s}AlternateContent' % NS['mc']:
            branch = el.find('{%s}Choice' % NS['mc'])
            if branch is None:
                branch = el.find('{%s}Fallback' % NS['mc'])
            if branch is not None:
                yield from iter_shapes(branch, recurse, _depth)
            continue
        if tag in SHAPE_TAGS:
            yield el, _depth
            if tag == P + 'grpSp' and recurse:
                yield from iter_shapes(el, recurse, _depth + 1)


def nv_props(shape_el):
    """Return the cNvPr element of a shape."""
    for child in shape_el:
        if child.tag.startswith(P + 'nv'):
            c = child.find(P + 'cNvPr')
            if c is not None:
                return c
    return None


def nv_pr(shape_el):
    for child in shape_el:
        if child.tag.startswith(P + 'nv'):
            n = child.find(P + 'nvPr')
            if n is not None:
                return n
    return None


def shape_id(shape_el):
    c = nv_props(shape_el)
    return int(c.get('id')) if c is not None and c.get('id', '').isdigit() else None


def shape_name(shape_el):
    c = nv_props(shape_el)
    return c.get('name', '') if c is not None else ''


def ph_of(shape_el):
    """(type, idx) of a placeholder shape or None. type defaults to 'obj', idx to 0."""
    n = nv_pr(shape_el)
    if n is None:
        return None
    ph = n.find(P + 'ph')
    if ph is None:
        return None
    return ph.get('type', 'obj'), int(ph.get('idx', '0'))


def ph_el(shape_el):
    n = nv_pr(shape_el)
    return None if n is None else n.find(P + 'ph')


def shape_kind(shape_el):
    tag = etree.QName(shape_el).localname
    if tag == 'graphicFrame':
        gd = shape_el.find('.//' + A + 'graphicData')
        uri = gd.get('uri', '') if gd is not None else ''
        if uri.endswith('/table'):
            return 'table'
        if uri.endswith('/chart'):
            return 'chart'
        if uri.endswith('/diagram'):
            return 'smartart'
        if uri.endswith('/ole'):
            return 'ole'
        return 'graphicFrame'
    if tag == 'pic':
        nvpr = nv_pr(shape_el)
        if nvpr is not None and (nvpr.find(A + 'videoFile') is not None):
            return 'video'
        if nvpr is not None and (nvpr.find(A + 'audioFile') is not None):
            return 'audio'
        if nvpr is not None and nvpr.find('.//' + P14 + 'media') is not None:
            return 'video'
        return 'picture'
    if tag == 'sp':
        if shape_el.find(P + 'txBody') is not None:
            nvsp = shape_el.find(P + 'nvSpPr/' + P + 'cNvSpPr')
            if nvsp is not None and nvsp.get('txBox') == '1':
                return 'textbox'
            return 'autoshape'
        return 'autoshape'
    if tag == 'grpSp':
        return 'group'
    if tag == 'cxnSp':
        return 'connector'
    return tag


def xfrm_el(shape_el):
    """The element holding off/ext for this shape (a:xfrm or p:xfrm)."""
    tag = etree.QName(shape_el).localname
    if tag == 'graphicFrame':
        return shape_el.find(P + 'xfrm')
    if tag == 'grpSp':
        g = shape_el.find(P + 'grpSpPr')
        return None if g is None else g.find(A + 'xfrm')
    sppr = shape_el.find(P + 'spPr')
    return None if sppr is None else sppr.find(A + 'xfrm')


def read_xfrm(shape_el):
    x = xfrm_el(shape_el)
    if x is None:
        return None
    off, ext = x.find(A + 'off'), x.find(A + 'ext')
    if off is None or ext is None:
        return None
    return [int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy'))]


def write_xfrm(shape_el, box):
    """Set off/ext. Creates the xfrm if missing. box = [x, y, w, h] (ints)."""
    x, y, w, h = [int(round(v)) for v in box]
    tag = etree.QName(shape_el).localname
    xf = xfrm_el(shape_el)
    if xf is None:
        if tag == 'graphicFrame':
            xf = etree.Element(P + 'xfrm')
            nv = shape_el.find(P + 'nvGraphicFramePr')
            nv.addnext(xf)
        else:
            parent_tag = P + ('grpSpPr' if tag == 'grpSp' else 'spPr')
            sppr = shape_el.find(parent_tag)
            if sppr is None:
                sppr = etree.Element(parent_tag)
                shape_el.find(P + ('nvGrpSpPr' if tag == 'grpSp' else
                                   'nvPicPr' if tag == 'pic' else
                                   'nvCxnSpPr' if tag == 'cxnSp' else 'nvSpPr')).addnext(sppr)
            xf = etree.Element(A + 'xfrm')
            sppr.insert(0, xf)
    off, ext = xf.find(A + 'off'), xf.find(A + 'ext')
    if off is None:
        off = etree.SubElement(xf, A + 'off')
    if ext is None:
        ext = etree.SubElement(xf, A + 'ext')
        if off.getnext() is not None and off.getnext() is not ext:
            off.addnext(ext)
    off.set('x', str(x)); off.set('y', str(y))
    ext.set('cx', str(w)); ext.set('cy', str(h))
    if tag == 'grpSp' and xf.find(A + 'chOff') is None:
        # A group without child extents: mirror the box so children don't jump.
        cho = etree.SubElement(xf, A + 'chOff'); cho.set('x', str(x)); cho.set('y', str(y))
        che = etree.SubElement(xf, A + 'chExt'); che.set('cx', str(w)); che.set('cy', str(h))


def clear_xfrm(shape_el):
    xf = xfrm_el(shape_el)
    if xf is not None:
        xf.getparent().remove(xf)


def is_top_level(shape_el):
    return shape_el.getparent().tag == P + 'spTree'


def group_transform(group_el):
    """Map child coordinates of a group to the parent space: returns f(box)->box."""
    xf = xfrm_el(group_el)
    if xf is None:
        return lambda b: b
    off, ext = xf.find(A + 'off'), xf.find(A + 'ext')
    cho, che = xf.find(A + 'chOff'), xf.find(A + 'chExt')
    if None in (off, ext, cho, che):
        return lambda b: b
    ox, oy, ew, eh = int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy'))
    cx, cy, cw, chh = int(cho.get('x')), int(cho.get('y')), int(che.get('cx')), int(che.get('cy'))
    sx = ew / cw if cw else 1.0
    sy = eh / chh if chh else 1.0
    return lambda b: [ox + (b[0] - cx) * sx, oy + (b[1] - cy) * sy, b[2] * sx, b[3] * sy]


def absolute_box(shape_el, inherited=None):
    """Slide-space box of a shape, applying the transforms of enclosing groups.
    `inherited` supplies the box for placeholders without their own xfrm."""
    box = read_xfrm(shape_el) or inherited
    if box is None:
        return None
    parent = shape_el.getparent()
    while parent is not None and parent.tag != P + 'spTree':
        if parent.tag == P + 'grpSp':
            box = group_transform(parent)(box)
        parent = parent.getparent()
    return [int(round(v)) for v in box]


def next_shape_id(slide_root):
    ids = [int(c.get('id')) for c in slide_root.iter(P + 'cNvPr') if c.get('id', '').isdigit()]
    return (max(ids) if ids else 1) + 1


def rects_overlap(a, b, tol=0):
    return not (a[0] + a[2] <= b[0] + tol or b[0] + b[2] <= a[0] + tol or
                a[1] + a[3] <= b[1] + tol or b[1] + b[3] <= a[1] + tol)


def fit_contain(src_w, src_h, box):
    """Largest src-aspect rect centered inside box [x,y,w,h]."""
    bx, by, bw, bh = box
    if src_w <= 0 or src_h <= 0:
        return list(box)
    s = min(bw / src_w, bh / src_h)
    w, h = src_w * s, src_h * s
    return [int(round(bx + (bw - w) / 2)), int(round(by + (bh - h) / 2)), int(round(w)), int(round(h))]


def round_emu(v, step=EMU_PER_IN // 100):
    """Round to a grid (default 0.01 in) so tiny drift clusters together."""
    return None if v is None else int(round(v / step) * step)


def median(vals):
    vals = sorted(v for v in vals if v is not None)
    if not vals:
        return None
    m = len(vals) // 2
    return vals[m] if len(vals) % 2 else (vals[m - 1] + vals[m]) // 2


def cluster_value(vals, step=EMU_PER_IN // 100):
    """Most common value after rounding to `step`; ties -> the median of the tie."""
    from collections import Counter
    vals = [round_emu(v, step) for v in vals if v is not None]
    if not vals:
        return None
    cnt = Counter(vals).most_common()
    top = cnt[0][1]
    tied = [v for v, n in cnt if n == top]
    return median(tied)


# --------------------------------------------------------------------------
# Placeholder inheritance (slide -> layout -> master)
# --------------------------------------------------------------------------
def _ph_candidates(root):
    out = []
    for el, _ in iter_shapes(sp_tree(root), recurse=False):
        ph = ph_of(el)
        if ph:
            out.append((el, ph))
    return out


MASTER_TYPE = {'ctrTitle': 'title', 'title': 'title', 'dt': 'dt', 'ftr': 'ftr',
               'sldNum': 'sldNum', 'hdr': 'hdr'}


def match_layout_ph(layout_root, ph):
    """Find the layout placeholder a slide placeholder (type, idx) inherits from."""
    typ, idx = ph
    cands = _ph_candidates(layout_root)
    # PowerPoint matches on idx first, then type
    for el, (t, i) in cands:
        if i == idx and (idx != 0 or t == typ or (t in TITLE_PH_TYPES and typ in TITLE_PH_TYPES)):
            return el
    for el, (t, i) in cands:
        if t == typ or (t in TITLE_PH_TYPES and typ in TITLE_PH_TYPES):
            return el
    return None


def match_master_ph(master_root, ph):
    typ, _ = ph
    want = MASTER_TYPE.get(typ, 'body')
    for el, (t, _i) in _ph_candidates(master_root):
        if MASTER_TYPE.get(t, 'body') == want and (t == want or want == 'body'):
            return el
    return None


def inherited_box(pkg, slide_part, shape_el):
    """Box for a slide placeholder following slide -> layout -> master."""
    box = read_xfrm(shape_el)
    if box:
        return box
    ph = ph_of(shape_el)
    if not ph:
        return None
    layout = pkg.layout_of(slide_part)
    if layout is None:
        return None
    lel = match_layout_ph(pkg.xml(layout), ph)
    if lel is not None and read_xfrm(lel):
        return read_xfrm(lel)
    master = pkg.master_of(layout)
    mel = match_master_ph(pkg.xml(master), ph if lel is None else ph_of(lel))
    if mel is not None:
        return read_xfrm(mel)
    return None


# --------------------------------------------------------------------------
# Theme + colors
# --------------------------------------------------------------------------
CLR_TAGS = ('srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr')
PRESET_COLORS = {'black': '000000', 'white': 'FFFFFF', 'red': 'FF0000', 'green': '008000',
                 'blue': '0000FF', 'yellow': 'FFFF00', 'gray': '808080'}


class Theme:
    def __init__(self, root):
        self.root = root
        self.colors = {}
        cs = root.find('.//' + A + 'clrScheme')
        if cs is not None:
            for child in cs:
                name = etree.QName(child).localname
                v = child[0] if len(child) else None
                if v is None:
                    continue
                if etree.QName(v).localname == 'srgbClr':
                    self.colors[name] = v.get('val').upper()
                elif etree.QName(v).localname == 'sysClr':
                    self.colors[name] = (v.get('lastClr') or '000000').upper()
        fs = root.find('.//' + A + 'fontScheme')
        self.major = self.minor = None
        if fs is not None:
            mj = fs.find(A + 'majorFont/' + A + 'latin')
            mn = fs.find(A + 'minorFont/' + A + 'latin')
            self.major = mj.get('typeface') if mj is not None else None
            self.minor = mn.get('typeface') if mn is not None else None
        self.name = root.get('name')

    def font(self, typeface):
        if typeface in ('+mj-lt', '+mj-ea', '+mj-cs'):
            return self.major
        if typeface in ('+mn-lt', '+mn-ea', '+mn-cs'):
            return self.minor
        return typeface


def _apply_mods(hexv, clr_el):
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    for m in clr_el:
        n, v = etree.QName(m).localname, int(m.get('val', '0')) / 100000
        if n == 'lumMod':
            l *= v
        elif n == 'lumOff':
            l += v
        elif n == 'tint':
            l = l * v + (1 - v)
        elif n == 'shade':
            l *= v
        elif n == 'satMod':
            s *= v
    l = min(max(l, 0), 1); s = min(max(s, 0), 1)
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return '%02X%02X%02X' % (round(r * 255), round(g * 255), round(b * 255))


def resolve_color(fill_parent, theme: Theme, clr_map: dict):
    """Hex for the first color child of a fill element (e.g. a:solidFill)."""
    if fill_parent is None:
        return None
    for el in fill_parent:
        n = etree.QName(el).localname
        if n not in CLR_TAGS:
            continue
        if n == 'srgbClr':
            base = el.get('val').upper()
        elif n == 'sysClr':
            base = (el.get('lastClr') or '000000').upper()
        elif n == 'prstClr':
            base = PRESET_COLORS.get(el.get('val'), '000000')
        elif n == 'schemeClr':
            key = el.get('val')
            key = clr_map.get(key, key)
            base = theme.colors.get(key, '000000') if theme else '000000'
        else:
            return None
        return _apply_mods(base, el) if len(el) else base
    return None


def clr_map_of(master_root):
    cm = master_root.find(P + 'clrMap')
    return dict(cm.attrib) if cm is not None else {
        'bg1': 'lt1', 'tx1': 'dk1', 'bg2': 'lt2', 'tx2': 'dk2'}


# --------------------------------------------------------------------------
# Effective text style resolver
# --------------------------------------------------------------------------
RUN_KEYS = ('font', 'size', 'bold', 'italic', 'underline', 'color')
PARA_KEYS = ('align', 'line_spacing', 'space_before', 'space_after', 'marL', 'indent', 'bullet')


def _lvl_ppr(lst, lvl):
    if lst is None:
        return None
    return lst.find(A + 'lvl%dpPr' % (lvl + 1))


def _read_rpr(rpr, theme, clr_map, out):
    if rpr is None:
        return
    if 'size' not in out and rpr.get('sz'):
        out['size'] = int(rpr.get('sz')) / 100
    if 'bold' not in out and rpr.get('b') is not None:
        out['bold'] = rpr.get('b') in ('1', 'true')
    if 'italic' not in out and rpr.get('i') is not None:
        out['italic'] = rpr.get('i') in ('1', 'true')
    if 'underline' not in out and rpr.get('u') is not None:
        out['underline'] = rpr.get('u') != 'none'
    if 'font' not in out:
        lat = rpr.find(A + 'latin')
        if lat is not None and lat.get('typeface'):
            out['font'] = theme.font(lat.get('typeface')) if theme else lat.get('typeface')
    if 'color' not in out:
        sf = rpr.find(A + 'solidFill')
        if sf is not None:
            out['color'] = resolve_color(sf, theme, clr_map)


def _spc(el):
    if el is None:
        return None
    p = el.find(A + 'spcPct')
    if p is not None:
        return 'pct:%g' % (int(p.get('val')) / 1000)
    t = el.find(A + 'spcPts')
    if t is not None:
        return 'pt:%g' % (int(t.get('val')) / 100)
    return None


def _read_ppr(ppr, out):
    if ppr is None:
        return
    if 'align' not in out and ppr.get('algn'):
        out['align'] = ppr.get('algn')
    if 'marL' not in out and ppr.get('marL') is not None:
        out['marL'] = int(ppr.get('marL'))
    if 'indent' not in out and ppr.get('indent') is not None:
        out['indent'] = int(ppr.get('indent'))
    for key, tag in (('line_spacing', 'lnSpc'), ('space_before', 'spcBef'), ('space_after', 'spcAft')):
        if key not in out:
            v = _spc(ppr.find(A + tag))
            if v is not None:
                out[key] = v
    if 'bullet' not in out:
        if ppr.find(A + 'buNone') is not None:
            out['bullet'] = 'none'
        elif ppr.find(A + 'buAutoNum') is not None:
            out['bullet'] = 'num:' + ppr.find(A + 'buAutoNum').get('type', '')
        elif ppr.find(A + 'buChar') is not None:
            out['bullet'] = 'char:' + ppr.find(A + 'buChar').get('char', '')
        elif ppr.find(A + 'buBlip') is not None:
            out['bullet'] = 'blip'


class StyleContext:
    """Everything needed to resolve text on one slide (or layout/master)."""

    def __init__(self, pkg: Package, part: str):
        self.pkg = pkg
        self.part = part
        root = pkg.xml(part)
        self.kind = etree.QName(root).localname  # sld, sldLayout, sldMaster
        self.slide = root if self.kind == 'sld' else None
        if self.kind == 'sld':
            self.layout_part = pkg.layout_of(part)
        elif self.kind == 'sldLayout':
            self.layout_part = part
        else:
            self.layout_part = None
        self.master_part = part if self.kind == 'sldMaster' else pkg.master_of(self.layout_part)
        self.layout = pkg.xml(self.layout_part) if self.layout_part else None
        self.master = pkg.xml(self.master_part)
        tp = pkg.theme_of(self.master_part)
        self.theme = Theme(pkg.xml(tp)) if tp else None
        self.clr_map = clr_map_of(self.master)
        pres = pkg.xml(pkg.main_part)
        self.default_text = pres.find(P + 'defaultTextStyle')

    def chain(self, shape_el, in_table=False):
        """List of lstStyle-like elements from most to least specific (excluding the shape's own)."""
        out = []
        ph = ph_of(shape_el) if not in_table else None
        tx = self.master.find(P + 'txStyles')
        if ph:
            lph = None
            if self.kind == 'sld' and self.layout is not None:
                lph = match_layout_ph(self.layout, ph)
                if lph is not None:
                    out.append(lph.find(P + 'txBody/' + A + 'lstStyle'))
            mph = match_master_ph(self.master, ph_of(lph) if lph is not None else ph)
            if mph is not None and self.kind != 'sldMaster':
                out.append(mph.find(P + 'txBody/' + A + 'lstStyle'))
            t = ph[0]
            if t in TITLE_PH_TYPES:
                out.append(tx.find(P + 'titleStyle') if tx is not None else None)
            elif t in FIXED_PH_TYPES:
                out.append(tx.find(P + 'otherStyle') if tx is not None else None)
            else:
                out.append(tx.find(P + 'bodyStyle') if tx is not None else None)
        else:
            out.append(tx.find(P + 'otherStyle') if tx is not None else None)
            out.append(self.default_text)
        return [e for e in out if e is not None]

    def resolve_paragraph(self, shape_el, p_el, in_table=False, skip_local=False):
        """Effective paragraph + 'paragraph default run' props. skip_local=True
        ignores everything set on the slide shape itself (what inheritance alone gives)."""
        lvl = int((p_el.find(A + 'pPr').get('lvl', '0')) if p_el.find(A + 'pPr') is not None else 0)
        para, run = {}, {}
        sources = []
        if not skip_local:
            sources.append(p_el.find(A + 'pPr'))
            tb = shape_el.find(P + 'txBody') if not in_table else None
            if tb is not None:
                sources.append(_lvl_ppr(tb.find(A + 'lstStyle'), lvl))
        for lst in self.chain(shape_el, in_table):
            sources.append(_lvl_ppr(lst, lvl))
        for ppr in sources:
            if ppr is None:
                continue
            _read_ppr(ppr, para)
            _read_rpr(ppr.find(A + 'defRPr'), self.theme, self.clr_map, run)
        if 'font' not in run and self.theme:
            ph = ph_of(shape_el)
            run['font'] = self.theme.major if ph and ph[0] in TITLE_PH_TYPES else self.theme.minor
        run.setdefault('size', 18.0)
        run.setdefault('bold', False)
        run.setdefault('italic', False)
        run.setdefault('underline', False)
        run.setdefault('color', self.theme.colors.get(self.clr_map.get('tx1', 'dk1'), '000000')
                       if self.theme else '000000')
        para.setdefault('align', 'l')
        para.setdefault('bullet', 'none')
        para['level'] = lvl
        return para, run

    def resolve_run(self, shape_el, p_el, r_el, in_table=False):
        para, base = self.resolve_paragraph(shape_el, p_el, in_table)
        run = {}
        _read_rpr(r_el.find(A + 'rPr'), self.theme, self.clr_map, run)
        for k, v in base.items():
            run.setdefault(k, v)
        return para, run


def paragraph_text(p_el):
    """Text of an a:p including fields and line breaks (as '\\n')."""
    out = []
    for child in p_el:
        n = etree.QName(child).localname
        if n in ('r', 'fld'):
            t = child.find(A + 't')
            out.append(t.text or '' if t is not None else '')
        elif n == 'br':
            out.append('\n')
    return ''.join(out)


def text_paragraphs(tx_body):
    if tx_body is None:
        return []
    return [paragraph_text(p) for p in tx_body.findall(A + 'p')]


# --------------------------------------------------------------------------
# Images
# --------------------------------------------------------------------------
def image_info(blob: bytes):
    """dict(format, w, h, frames, durations) or None when Pillow can't read it."""
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(blob))
        frames = getattr(im, 'n_frames', 1)
        durations = []
        if frames > 1:
            for i in range(frames):
                im.seek(i)
                durations.append(im.info.get('duration'))
            im.seek(0)
        return dict(format=im.format, w=im.width, h=im.height, frames=frames, durations=durations)
    except Exception:
        return None


def dhash(blob: bytes, size=8):
    """64-bit difference hash as int, or None."""
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(blob))
        im.seek(0)
        im = im.convert('RGBA')
        bg = Image.new('RGBA', im.size, (255, 255, 255, 255))
        bg.alpha_composite(im)
        g = bg.convert('L').resize((size + 1, size), Image.LANCZOS)
        px = list(g.getdata())
        bits = 0
        for row in range(size):
            for col in range(size):
                left = px[row * (size + 1) + col]
                right = px[row * (size + 1) + col + 1]
                bits = (bits << 1) | (1 if left > right else 0)
        return bits
    except Exception:
        return None


def hamming(a, b):
    if a is None or b is None:
        return 64
    return bin(a ^ b).count('1')


def blip_part(pkg: Package, slide_part: str, pic_el):
    """Part name of the picture's (poster) image, or None for linked/missing."""
    blip = pic_el.find('.//' + A + 'blip')
    if blip is None:
        return None
    rid = blip.get(R + 'embed')
    if not rid:
        return None
    rel = pkg.rel_target(slide_part, rid)
    return rel['abs'] if rel and rel['abs'] and pkg.has(rel['abs']) else None


def media_parts(pkg: Package, slide_part: str, pic_el):
    """Video/audio parts referenced by a media picture: list of (kind, part_or_url, external)."""
    out = []
    nvpr = nv_pr(pic_el)
    if nvpr is None:
        return out
    refs = []
    for tag in ('videoFile', 'audioFile', 'quickTimeFile', 'wavAudioFile'):
        el = nvpr.find(A + tag)
        if el is not None:
            refs.append((tag, el.get(R + 'link') or el.get(R + 'embed')))
    for m in nvpr.iter(P14 + 'media'):
        refs.append(('p14:media', m.get(R + 'embed') or m.get(R + 'link')))
    for kind, rid in refs:
        rel = pkg.rel_target(slide_part, rid) if rid else None
        if rel is None:
            continue
        if rel['mode'] == 'External':
            out.append((kind, rel['target'], True))
        else:
            out.append((kind, rel['abs'], False))
    return out


# --------------------------------------------------------------------------
# Misc
# --------------------------------------------------------------------------
@dataclass
class Finding:
    level: str  # 'fail' | 'warn' | 'info'
    slide: int | None
    what: str

    def as_dict(self):
        return dict(level=self.level, slide=self.slide, what=self.what)


def natural_key(s):
    return [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', s)]


def deepcopy(el):
    return copy.deepcopy(el)
