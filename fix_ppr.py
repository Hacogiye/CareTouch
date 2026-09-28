# -*- coding: utf-8 -*-
"""Post-process pptx: keep only the first <a:pPr> in each <a:p> (schema fix),
then repack. Also verifies the file reopens."""
import sys, zipfile, shutil, os, re
import xml.etree.ElementTree as ET

NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
ET.register_namespace("a", NS)

def fix_ppr(xml_bytes):
    root = ET.fromstring(xml_bytes)
    changed = 0
    for p in root.iter(f"{{{NS}}}p"):
        pprs = [c for c in list(p) if c.tag == f"{{{NS}}}pPr"]
        if len(pprs) > 1:
            # keep first occurrence, remove the rest
            first = pprs[0]
            for extra in pprs[1:]:
                p.remove(extra)
                changed += 1
        elif len(pprs) == 1:
            first = pprs[0]
            idx = list(p).index(first)
            if idx != 0:  # pPr must be first child
                p.remove(first)
                p.insert(0, first)
                changed += 1
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8"), changed

def main(src):
    tmp = src + ".tmp"
    total = 0
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        # content types first
        for n in ["[Content_Types].xml"] + [x for x in names if x != "[Content_Types].xml"]:
            data = zin.read(n)
            if re.match(r"ppt/slides/slide\d+\.xml$", n):
                data, c = fix_ppr(data)
                total += c
            zout.writestr(n, data)
    shutil.move(tmp, src)
    print(f"fixed {total} pPr issues in {src}")
    # verify reopen with python-pptx
    from pptx import Presentation
    prs = Presentation(src)
    print(f"reopens OK — {len(prs.slides)} slides")

if __name__ == "__main__":
    main(sys.argv[1])
