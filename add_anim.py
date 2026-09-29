# -*- coding: utf-8 -*-
"""Chèn hiệu ứng trình chiếu vào CareTouch-deck.pptx:
- Entrance "fade (+nhích lên)" theo nhóm: mỗi lần bấm chuột hiện 1 khối
- S1 bìa / S14 kết: tự chạy khi mở slide
- Chuyển slide: fade 600ms
Cách dùng:
  python add_anim.py                       -> chèn cả timing + transition vào CareTouch-deck.pptx
  python -c "import add_anim; add_anim.make('A.pptx','B.pptx', timing=True, transition=False, rise=True)"
"""
import zipfile, shutil, re

SRC = "CareTouch-deck.pptx"

# Plan: slide -> ("auto" | "click"), groups = [ [(spid, delay_ms), ...], ... ]
# Shape tĩnh (kicker, tiêu đề, motif, số trang, dòng nguồn) KHÔNG nằm trong plan -> luôn hiển thị.
PLAN = {
    1: ("auto", [[(2,0),(3,200),(4,450),(5,650),(6,850),(7,1050)]]),
    2: ("click", [[(7,0),(8,100),(9,200)], [(10,0),(11,100),(12,200),(13,300)],
                  [(14,0),(15,100),(16,200),(17,300)],
                  [(18,0),(19,150),(20,300),(21,380),(22,460),(23,540),(24,620),(25,700)]]),
    3: ("click", [[(7,0),(8,100),(9,200)], [(10,0),(11,100),(12,200)],
                  [(13,0),(14,100),(15,200)], [(16,0),(17,100),(18,200)]]),
    4: ("click", [[(7,0),(8,100),(9,200),(10,300),(11,400)], [(12,0),(13,100),(14,200),(15,300),(16,400)],
                  [(17,0),(18,100),(19,200),(20,300),(21,400)], [(22,0),(23,100)]]),
    5: ("click", [[(7,0),(8,100),(9,200),(10,300),(11,400)],
                  [(12,0),(13,100),(14,200),(15,300),(16,400),(17,500)],
                  [(18,0),(19,100),(20,250)],
                  [(21,0),(22,100),(23,200),(24,300),(25,400),(26,500)],
                  [(27,0),(28,100),(29,200),(30,300),(31,400),(32,500)]]),
    6: ("click", [[(7,0)], [(8,0),(9,100),(10,200)], [(11,0),(12,100)]]),
    7: ("click", [[(7,0),(8,100),(9,200),(10,300)], [(11,0),(12,100),(13,200),(14,300)],
                  [(15,0),(16,100),(17,200),(18,300)], [(19,0),(20,100)]]),
    8: ("click", [[(7,0),(8,100),(9,200),(10,300),(11,400)], [(12,0),(13,100),(14,200),(15,300),(16,400)],
                  [(17,0),(18,100),(19,200),(20,300),(21,400)], [(22,0),(23,100),(24,200),(25,300)],
                  [(26,0),(27,100)]]),
    9: ("click", [[(10,0)]]),
    10: ("click", [[(7,0),(8,100),(9,200),(26,300)], [(22,0)],
                   [(10,0),(11,80),(12,160),(13,240),(14,320),(15,400),(16,480),(17,560),(18,640),(19,720),(20,800),(21,880)]]),
    11: ("click", [[(7,0),(8,100),(9,200),(10,300),(11,400),(12,500)], [(13,0),(14,100),(15,200),(16,300),(17,400)],
                   [(18,0),(19,100),(20,200),(21,300),(22,400)], [(23,0),(24,100),(25,200),(26,300),(27,400)],
                   [(28,0),(29,100)]]),
    12: ("click", [[(7,0),(8,100),(9,200),(10,300),(11,400)], [(12,0),(13,100),(14,200),(15,300),(16,400)],
                   [(17,0),(18,100),(19,200),(20,300),(21,400)], [(22,0),(23,100),(24,200),(25,300),(26,400)]]),
    13: ("click", [[(7,0),(8,100),(9,200),(10,300),(11,400)], [(12,0),(13,100),(14,200),(15,300),(16,400)],
                   [(17,0),(18,100),(19,200),(20,300),(21,400)], [(22,0),(23,100),(24,200),(25,300),(26,400)],
                   [(27,0),(28,100),(29,200),(30,300),(31,400)]]),
    14: ("auto", [[(2,0),(3,400),(4,700),(5,850),(6,1050),(7,1130),(8,1210),(9,1290),(10,1400),(11,1500),(12,1650)]]),
}

TRANSITION = (
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    '<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">'
    '<p:transition spd="med" p14:dur="600"><p:fade/></p:transition></mc:Choice>'
    '<mc:Fallback><p:transition spd="med"><p:fade/></p:transition></mc:Fallback>'
    '</mc:AlternateContent>'
)

_id = [1]
def nid():
    _id[0] += 1
    return _id[0]

def effect_par(spid, delay, ntype, rise=True):
    a = nid()
    parts = [
        f'<p:par><p:cTn id="{a}" presetID="42" presetClass="entr" presetSubtype="0" '
        f'fill="hold" grpId="0" nodeType="{ntype}">'
        f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>',
        f'<p:set><p:cBhvr><p:cTn id="{nid()}" dur="1" fill="hold">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
        f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
        f'<p:to><p:strVal val="visible"/></p:to></p:set>',
        f'<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{nid()}" dur="500"/>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>',
    ]
    if rise:
        parts.append(
            f'<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base">'
            f'<p:cTn id="{nid()}" dur="500" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
            f'<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_y+0.018"/></p:val></p:tav>'
            f'<p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav></p:tavLst></p:anim>'
        )
    parts.append('</p:childTnLst></p:cTn></p:par>')
    return "".join(parts)

def click_group(group, mode, rise=True):
    g1, g2 = nid(), nid()
    start = '<p:cond delay="0"/>' if mode == "auto" else '<p:cond delay="indefinite"/>'
    effects = []
    for j, (spid, delay) in enumerate(group):
        if mode == "auto":
            nt = "afterEffect" if j == 0 else "withEffect"
        else:
            nt = "clickEffect" if j == 0 else "withEffect"
        effects.append(effect_par(spid, delay, nt, rise))
    return (
        f'<p:par><p:cTn id="{g1}" fill="hold"><p:stCondLst>{start}</p:stCondLst>'
        f'<p:childTnLst><p:par><p:cTn id="{g2}" fill="hold">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
        + "".join(effects) +
        f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
    )

def timing_xml(slide_no, tables_charts, rise=True, with_bld=True):
    _id[0] = 1
    mode, groups = PLAN[slide_no]
    click_blocks = "".join(click_group(g, mode, rise) for g in groups)
    anim_spids = sorted({spid for g in groups for spid, _ in g})
    bld_part = ""
    if with_bld:
        bld = "".join(
            f'<p:bldGraphic spid="{s}" grpId="0"><p:bldAsOne/></p:bldGraphic>' if s in tables_charts
            else f'<p:bldP spid="{s}" grpId="0"/>'
            for s in anim_spids
        )
        bld_part = f'<p:bldLst>{bld}</p:bldLst>'
    return (
        '<p:timing><p:tnLst><p:par>'
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        '<p:seq concurrent="1" nextAc="seek">'
        f'<p:cTn id="{nid()}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        + click_blocks +
        '</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
        + bld_part + '</p:timing>'
    )

def make(src, dst, timing=True, transition=True, rise=True, with_bld=True):
    """Đọc src, chèn timing/transition, ghi ra dst (dst == src thì ghi tạm rồi thay)."""
    from pptx import Presentation
    prs = Presentation(src)
    gf_ids = {}
    for i, slide in enumerate(prs.slides, 1):
        gf_ids[i] = {sh.shape_id for sh in slide.shapes
                     if getattr(sh, "has_table", False) or getattr(sh, "has_chart", False)}
    tmp = dst + ".building"
    zin = zipfile.ZipFile(src, "r")
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    pat = re.compile(r"ppt/slides/slide(\d+)\.xml$")
    injected = 0
    for item in zin.infolist():
        data = zin.read(item.filename)
        m = pat.match(item.filename)
        if m:
            n = int(m.group(1))
            if n in PLAN:
                xml = data.decode("utf-8")
                root = re.search(r"<(\w+):sld\b", xml)
                assert root, "không tìm thấy root <sld>"
                pfx = root.group(1)
                inject = (TRANSITION if transition else "") + (timing_xml(n, gf_ids[n], rise, with_bld) if timing else "")
                if inject:
                    inject = inject.replace("p:", pfx + ":")
                    xml = re.sub(rf"<{pfx}:timing>.*?</{pfx}:timing>", "", xml, flags=re.S)
                    assert xml.count(f"</{pfx}:sld>") == 1
                    xml = xml.replace(f"</{pfx}:sld>", inject + f"</{pfx}:sld>")
                    injected += 1
                data = xml.encode("utf-8")
        zout.writestr(item, data)
    zout.close(); zin.close()
    if dst == src:
        shutil.move(tmp, src)
    else:
        shutil.move(tmp, dst)
    return injected

if __name__ == "__main__":
    n = make(SRC, SRC)
    print(f"Đã chèn timing + transition cho {n} slide.")
