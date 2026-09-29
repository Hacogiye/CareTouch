# -*- coding: utf-8 -*-
"""Chèn hiệu ứng trình chiếu vào CareTouch-deck.pptx.

- Entrance "fade (+ nhích lên)" theo NHÓM: mỗi lần bấm chuột hiện 1 khối.
- Slide bìa & slide kết: tự chạy khi mở (một nhóm, cascade).
- Chuyển slide: fade 600ms.

PLAN được TỰ SINH từ shape thật của file PPTX (không hard-code id), nên
build lại deck bao nhiêu lần cũng không sinh spid "treo" -> không còn lỗi
PowerPoint repair.

Cách dùng:
  python add_anim.py                       # chèn vào CareTouch-deck.pptx
  python add_anim.py --dry                 # chỉ in PLAN, không ghi file
"""
import zipfile, shutil, re, sys

SRC = "CareTouch-deck.pptx"

TRANSITION = (
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    '<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">'
    '<p:transition spd="med" p14:dur="600"><p:fade/></p:transition></mc:Choice>'
    '<mc:Fallback><p:transition spd="med"><p:fade/></p:transition></mc:Fallback>'
    '</mc:AlternateContent>'
)

EMU = 914400.0


# ---------------------------------------------------------------- static parts
def _is_static(sh, slide_w, slide_h):
    """kicker, tiêu đề, motif, số trang, dòng nguồn -> luôn hiển thị."""
    x, y = sh.left / EMU, sh.top / EMU
    w, h = sh.width / EMU, sh.height / EMU
    xc, yc = x + w / 2, y + h / 2
    if yc < 1.72:                                   # dải trên: kicker + tiêu đề
        return True
    if xc > 12.2 and yc < 1.0:                      # motif 2 vòng tròn
        return True
    if xc > 12.4 and yc > 6.9:                      # số trang
        return True
    if yc > 6.6:                                    # dòng nguồn / thanh tổng kết
        return True
    return False


def _cluster(vals, tol):
    """Gom giá trị đã sort thành cụm: cắt khi khe hở >= tol."""
    out, cur = [], [vals[0]]
    for v in vals[1:]:
        if v - cur[-1] >= tol:
            out.append(cur); cur = []
        cur.append(v)
    out.append(cur)
    return out


def _bbox(sh):
    return (sh.left, sh.top, sh.left + sh.width,
            sh.top + max(sh.height, 91440))          # dòng kẻ cao 0 vẫn có hộp


def _overlap(a, b, pad):
    al, at, ar, ab = a
    bl, bt, br, bb = b
    return not (ar < bl - pad or br < al - pad or ab < bt - pad or bb < at - pad)


def _components(shapes, pad=0.2):
    """Gom shape thành CỤM theo hộp bao chồng nhau (union-find).

    Mỗi cụm ≈ một "thẻ" nội dung (khung + số + chữ bên trong) nên một lần
    bấm chuột hiện trọn thẻ — giống bản PLAN chỉnh tay trước đây nhưng tự
    sinh, không phụ thuộc id cứng."""
    shapes = list(shapes)
    n = len(shapes)
    par = list(range(n))

    def find(x):
        while par[x] != x:
            par[x] = par[par[x]]
            x = par[x]
        return x

    boxes = [_bbox(s) for s in shapes]
    for i in range(n):
        for j in range(i + 1, n):
            if _overlap(boxes[i], boxes[j], pad):
                par[find(i)] = find(j)
    groups = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(shapes[i])
    return list(groups.values())


def build_plan(prs):
    """Sinh PLAN: {slide_no: (mode, [[(spid, delay), ...], ...])}.

    Mọi slide đều ở chế độ "auto": nội dung tự hiện khi mở slide, không cần
    bấm chuột. Thứ tự hiện theo hàng-dải đọc (trên xuống, trái sang), mỗi
    khối so le một nhịp; nhịp được co lại khi slide nhiều khối để tổng thời
    gian hiện không quá ~3 giây."""
    plan = {}
    for idx, slide in enumerate(prs.slides, 1):
        W = prs.slide_width / EMU
        H = prs.slide_height / EMU
        anim = [sh for sh in slide.shapes if not _is_static(sh, W, H)]
        if not anim:
            continue
        # Dải đọc: gom theo hàng (band 0,6in), trong hàng đi trái→phải;
        # shape lớn (nền thẻ) hiện trước nội dung của nó nhờ -area.
        anim.sort(key=lambda s: (
            round((s.top + s.height / 2) / 0.6),
            s.left,
            -(s.width * s.height),
        ))
        stagger = int(min(180, 2800 / len(anim)))
        plan[idx] = ("auto", [[(s.shape_id, i * stagger) for i, s in enumerate(anim)]])
    return plan


# ---------------------------------------------------------------- timing XML
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


def timing_xml(mode, groups, tables_charts, rise=True, with_bld=True):
    _id[0] = 1
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


# ---------------------------------------------------------------- main
def _normalize_ids(src, dst):
    """pptxgenjs cấp id bảng và id ảnh từ hai bộ đếm riêng nên có thể trùng
    (vd slide 11: bảng id 12 và ảnh id 12). Trùng cNvPr id làm PowerPoint
    báo repair và làm hiệu ứng trỏ sai shape -> đánh lại id cho duy nhất."""
    import re as _re
    zin = zipfile.ZipFile(src, "r")
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    pat = _re.compile(r"ppt/slides/slide\d+\.xml$")
    fixed = 0
    for item in zin.infolist():
        data = zin.read(item.filename)
        if pat.match(item.filename):
            xml = data.decode("utf-8")
            ids = [int(i) for i in _re.findall(r'<p:cNvPr id="(\d+)"', xml)]
            if len(ids) != len(set(ids)):
                used = set(ids)
                nxt = max(ids) + 1
                seen = set()
                def _sub(m):
                    nonlocal nxt, fixed
                    i = int(m.group(1))
                    if i in seen:
                        while nxt in used:
                            nxt += 1
                        used.add(nxt); seen.add(nxt); fixed += 1
                        return f'<p:cNvPr id="{nxt}"'
                    seen.add(i)
                    return m.group(0)
                xml = _re.sub(r'<p:cNvPr id="(\d+)"', _sub, xml)
                data = xml.encode("utf-8")
        zout.writestr(item, data)
    zout.close(); zin.close()
    return fixed


def make(src, dst, timing=True, transition=True, rise=True, with_bld=True, dry=False):
    from pptx import Presentation
    norm = src + ".norm"
    nfix = _normalize_ids(src, norm)
    if nfix:
        print(f"Đã sửa {nfix} shape id trùng.")
    prs = Presentation(norm)
    plan = build_plan(prs)
    gf = {}
    for i, slide in enumerate(prs.slides, 1):
        gf[i] = {sh.shape_id for sh in slide.shapes
                 if getattr(sh, "has_table", False) or getattr(sh, "has_chart", False)}

    if dry:
        for n in sorted(plan):
            mode, groups = plan[n]
            print(f"S{n:02d} [{mode}] " + " | ".join(
                ",".join(str(s) for s, _ in g) for g in groups))
        import os
        os.remove(norm)
        return plan

    tmp = dst + ".building"
    zin = zipfile.ZipFile(norm, "r")
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    pat = re.compile(r"ppt/slides/slide(\d+)\.xml$")
    injected = 0
    for item in zin.infolist():
        data = zin.read(item.filename)
        m = pat.match(item.filename)
        if m:
            n = int(m.group(1))
            if n in plan:
                xml = data.decode("utf-8")
                root = re.search(r"<(\w+):sld\b", xml)
                assert root, "không tìm thấy root <sld>"
                pfx = root.group(1)
                mode, groups = plan[n]
                inject = (TRANSITION if transition else "") + (
                    timing_xml(mode, groups, gf[n], rise, with_bld) if timing else "")
                if inject:
                    inject = inject.replace("p:", pfx + ":")
                    xml = re.sub(rf"<{pfx}:timing>.*?</{pfx}:timing>", "", xml, flags=re.S)
                    assert xml.count(f"</{pfx}:sld>") == 1
                    xml = xml.replace(f"</{pfx}:sld>", inject + f"</{pfx}:sld>")
                    injected += 1
                data = xml.encode("utf-8")
        zout.writestr(item, data)
    zout.close(); zin.close()
    import os
    os.remove(norm)
    shutil.move(tmp, src if dst == src else dst)
    return injected


if __name__ == "__main__":
    if "--dry" in sys.argv:
        make(SRC, SRC, dry=True)
    else:
        n = make(SRC, SRC)
        print(f"Đã chèn timing + transition cho {n} slide.")
