#!/usr/bin/env python3
"""Repair schema-invalid OOXML in the 1851 chapter decks so Windows PowerPoint
opens them without the 'repair' prompt.

Fixes, applied to every XML part of each .pptx:
1. <a:p>  child order  -> pPr?, (r|br|fld)*, endParaRPr?
2. <a:ln> child order  -> fill, prstDash, custDash, join, headEnd, tailEnd, extLst
3. negative <a:ext> cx/cy -> absolute value + flipH/flipV on parent <a:xfrm>
   (and abs() on <a:chExt>)
4. xml:space attribute removed from <a:t> (not allowed by CT_TextString)
5. <a:buSzPct val="100000"> -> val="100%" (schema-valid form, same meaning)
"""
import glob, os, shutil, zipfile
from lxml import etree

BASE = "/Users/iahmad/Creator/Courses_and_conferences/LT/courses/1851/1851-FTP-upload"
BAKDIR = os.path.join(BASE, "prefix_backup")
os.makedirs(BAKDIR, exist_ok=True)

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

P_ORDER = {"pPr": 0, "r": 1, "br": 1, "fld": 1, "endParaRPr": 2}
LN_ORDER = {
    "noFill": 0, "solidFill": 0, "gradFill": 0, "blipFill": 0, "pattFill": 0, "grpFill": 0,
    "prstDash": 1, "custDash": 2,
    "round": 3, "bevel": 3, "miter": 3,
    "headEnd": 4, "tailEnd": 5, "extLst": 6,
}

def local(el):
    return etree.QName(el).localname

def reorder(el, order):
    kids = list(el)
    keyed = [(order.get(local(c), 99), i, c) for i, c in enumerate(kids)]
    keyed.sort(key=lambda t: (t[0], t[1]))
    if [k[2] for k in keyed] != kids:
        for c in kids:
            el.remove(c)
        for _, _, c in keyed:
            el.append(c)
        return True
    return False

def fix_xml(data):
    """Return (new_bytes, changed, stats)."""
    try:
        root = etree.fromstring(data)
    except Exception:
        return data, False, {}
    changed = False
    stats = {}

    # 1. paragraph child order
    for el in root.iter(f"{A}p"):
        if reorder(el, P_ORDER):
            changed = True; stats['p_order'] = stats.get('p_order', 0) + 1

    # 2. line-properties child order
    for el in root.iter(f"{A}ln"):
        if reorder(el, LN_ORDER):
            changed = True; stats['ln_order'] = stats.get('ln_order', 0) + 1

    # 3. negative extents -> flips
    for el in root.iter(f"{A}xfrm"):
        ext = el.find(f"{A}ext")
        if ext is not None:
            cx, cy = ext.get('cx'), ext.get('cy')
            if cx is not None and int(cx) < 0:
                ext.set('cx', str(-int(cx))); el.set('flipH', '1')
                changed = True; stats['neg_cx'] = stats.get('neg_cx', 0) + 1
            if cy is not None and int(cy) < 0:
                ext.set('cy', str(-int(cy))); el.set('flipV', '1')
                changed = True; stats['neg_cy'] = stats.get('neg_cy', 0) + 1
        chext = el.find(f"{A}chExt")
        if chext is not None:
            for attr in ('cx', 'cy'):
                v = chext.get(attr)
                if v is not None and int(v) < 0:
                    chext.set(attr, str(-int(v)))
                    changed = True; stats['neg_chext'] = stats.get('neg_chext', 0) + 1

    # 4. xml:space on <a:t>
    for el in root.iter(f"{A}t"):
        if el.get(XML_SPACE) is not None:
            del el.attrib[XML_SPACE]
            changed = True; stats['xml_space'] = stats.get('xml_space', 0) + 1

    # 5. buSzPct numeric -> percent string
    for el in root.iter(f"{A}buSzPct"):
        v = el.get('val')
        if v is not None and v.isdigit():
            el.set('val', f"{int(v) // 1000}%")
            changed = True; stats['buszpct'] = stats.get('buszpct', 0) + 1

    if not changed:
        return data, False, stats
    out = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    return out, True, stats

def fix_file(path):
    with open(path, 'rb') as fh:
        zin = zipfile.ZipFile(fh)
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    total = {}
    new_parts = []
    any_change = False
    for info, data in items:
        if info.filename.endswith('.xml'):
            data, ch, stats = fix_xml(data)
            if ch:
                any_change = True
                for k, v in stats.items():
                    total[k] = total.get(k, 0) + v
        new_parts.append((info, data))
    if not any_change:
        return None
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info, data in new_parts:
            zout.writestr(info, data)
    os.replace(tmp, path)
    return total

for f in sorted(glob.glob(os.path.join(BASE, "1851-Ch*.pptx"))):
    if 'backup' in f:
        continue
    name = os.path.basename(f)
    bak = os.path.join(BAKDIR, name)
    if not os.path.exists(bak):
        shutil.copy2(f, bak)
    stats = fix_file(f)
    print(f"{name}: {'no changes' if stats is None else stats}")
print("backups in", BAKDIR)
