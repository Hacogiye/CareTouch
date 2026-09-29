# -*- coding: utf-8 -*-
"""CareTouch — báo cáo dự án khởi nghiệp (DOCX, tiếng Việt, nguồn siêu liên kết)."""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__))
A = lambda p: os.path.join(HERE, "assets", p)

# brand colors
DARK = RGBColor(0x0B, 0x3B, 0x34)
PRIMARY = RGBColor(0x0E, 0x7C, 0x6B)
ACCENT = RGBColor(0xE8, 0x59, 0x0C)
MUTED = RGBColor(0x5C, 0x73, 0x70)
TEXT = RGBColor(0x16, 0x30, 0x2B)
HEX_DARK = "0B3B34"; HEX_PRIMARY = "0E7C6B"; HEX_TINT = "E8F5F1"; HEX_TINT2 = "F4FAF8"; HEX_ACCENT = "E8590C"; HEX_ACCENTSOFT = "FDEBDD"

doc = Document()

# ---------- page & base styles ----------
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.4)
sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(2.0)

def set_font(style, name="Segoe UI", size=11, color=TEXT, bold=False, italic=False):
    style.font.name = name; style.font.size = Pt(size)
    style.font.color.rgb = color; style.font.bold = bold; style.font.italic = italic
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(attr), name)

normal = doc.styles["Normal"]
set_font(normal, size=11)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.28

for hname, hsize, hcolor, hbefore in [("Heading 1", 16, DARK, 14), ("Heading 2", 13, PRIMARY, 10), ("Heading 3", 11.5, TEXT, 8)]:
    st = doc.styles[hname]
    set_font(st, size=hsize, color=hcolor, bold=True)
    st.paragraph_format.space_before = Pt(hbefore)
    st.paragraph_format.space_after = Pt(5)
    st.paragraph_format.keep_with_next = True

# ---------- helpers ----------
def para(text="", style=None, size=None, color=None, bold=False, italic=False,
         align=None, space_after=None, space_before=None):
    p = doc.add_paragraph(style=style)
    if text:
        r = p.add_run(text)
        if size: r.font.size = Pt(size)
        r.font.color.rgb = color if color else TEXT
        r.font.bold = bold; r.font.italic = italic
    if align is not None: p.alignment = align
    if space_after is not None: p.paragraph_format.space_after = Pt(space_after)
    if space_before is not None: p.paragraph_format.space_before = Pt(space_before)
    return p

def rich(p, parts):
    """parts: list of (text, dict(color=,bold=,italic=,link=))"""
    for t, o in parts:
        if o.get("link"):
            add_hyperlink(p, t, o.get("link"), color=o.get("color", PRIMARY), bold=o.get("bold", False))
        else:
            r = p.add_run(t)
            r.font.color.rgb = o.get("color", TEXT)
            r.font.bold = o.get("bold", False)
            r.font.italic = o.get("italic", False)
            if o.get("size"): r.font.size = Pt(o["size"])
    return p

def add_hyperlink(p, text, url, color=PRIMARY, bold=False, size=None):
    part = p.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hl = OxmlElement("w:hyperlink"); hl.set(qn("r:id"), r_id)
    r = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs"): rf.set(qn(attr), "Segoe UI")
    rPr.append(rf)
    col = OxmlElement("w:color"); col.set(qn("w:val"), "0E7C6B" if color == PRIMARY else ("E8590C" if color == ACCENT else "16302B")); rPr.append(col)
    if bold:
        b = OxmlElement("w:b"); rPr.append(b)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    if size:
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size * 2))); rPr.append(sz)
    r.append(rPr)
    t = OxmlElement("w:t"); t.text = text; t.set(qn("xml:space"), "preserve")
    r.append(t); hl.append(r); p._p.append(hl)
    return hl

def shade_cell(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)

def style_table(table, col_widths=None, header_fill=HEX_PRIMARY, zebra=True, font_size=10.5, header=True):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = doc.styles["Table Grid"]
    # borders -> light
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "DCEBE6")
        borders.append(el)
    tblPr.append(borders)
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.space_before = Pt(2)
                for r in p.runs:
                    r.font.size = Pt(font_size)
                    r.font.name = "Segoe UI"
                    if ri == 0 and header:
                        r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    else:
                        r.font.color.rgb = TEXT
            if ri == 0 and header:
                shade_cell(cell, header_fill)
            elif zebra and ri % 2 == 0:
                shade_cell(cell, HEX_TINT2)
            if col_widths:
                cell.width = Cm(col_widths[ci])
    if col_widths:
        for ri in range(len(table.rows)):
            for ci, w in enumerate(col_widths):
                table.rows[ri].cells[ci].width = Cm(w)

def caption(text):
    p = para(align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=10)
    rich(p, [(text, dict(color=MUTED, italic=True, size=9.5))])
    return p

def bullet(text_parts, level=0):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    if isinstance(text_parts, str):
        text_parts = [(text_parts, {})]
    rich(p, text_parts)
    p.paragraph_format.space_after = Pt(3)
    return p

def chart_img(fname, width_cm=14.6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_picture(A(fname), width=Cm(width_cm))

def page_break():
    doc.add_paragraph().add_run().add_break()
    doc.paragraphs[-1].runs[0].add_break() if False else None

# ================= TITLE PAGE =================
p = para(align=WD_ALIGN_PARAGRAPH.CENTER, space_before=30, space_after=4)
p.add_run().add_picture(A("logo.png"), width=Cm(7.2))
para("BÁO CÁO DỰ ÁN KHỞI NGHIỆP", size=13, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
rich(p, [("Care", dict(color=DARK, bold=True, size=30)), ("Touch", dict(color=ACCENT, bold=True, size=30)),
         ("  —  Healing Touch, Healthy Life", dict(color=MUTED, italic=True, size=15))])
para("Dịch vụ massage – xoa bóp, bấm huyệt, gội đầu dưỡng sinh và chườm thảo dược theo phương pháp y học cổ truyền — hỗ trợ thư giãn cơ khớp; tại cơ sở và tận nhà",
     size=12.5, color=TEXT, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=8, space_after=18)
p = para(align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)
rich(p, [("Đổi thương hiệu từ “An Khang Đường” sang ", dict(color=MUTED, size=11)), ("CareTouch", dict(color=ACCENT, bold=True, size=11))])

info = doc.add_table(rows=5, cols=2)
info_data = [
    ("Nhóm thực hiện", "Nhóm 3 · Tổ 14 · Lớp A1K79"),
    ("Thành viên", "Trịnh Hương Thảo (2401634) · Bạch Thùy Trâm (2401679) · Bùi Thị Anh Trâm (2401681) · Phạm Quốc Việt (2401742) · Viengkeo Shengchan (2401761)"),
    ("Môn học", "Thực tập Nghiên cứu khoa học — Đổi mới sáng tạo Khởi nghiệp · Bài 3: Lựa chọn ý tưởng kinh doanh"),
    ("Buổi thực tập", "Ca chiều 1 · Thứ 7, ngày 26 tháng 9 năm 2026"),
    ("Phiên bản tài liệu", "Báo cáo đầy đủ cho giảng viên & nhà đầu tư — tháng 9/2026"),
]
for i, (k, v) in enumerate(info_data):
    c0, c1 = info.rows[i].cells
    c0.width, c1.width = Cm(4.2), Cm(11.6)
    r = c0.paragraphs[0].add_run(k); r.font.bold = True; r.font.size = Pt(10); r.font.color.rgb = PRIMARY
    r = c1.paragraphs[0].add_run(v); r.font.size = Pt(10); r.font.color.rgb = TEXT
    shade_cell(c0, HEX_TINT)
style_table(info, col_widths=[4.2, 11.6], header_fill=HEX_TINT, zebra=True, font_size=10, header=False)

para("", space_after=0)
para("Tài liệu phục vụ mục đích học tập và thuyết trình gọi vốn minh họa. Toàn bộ số liệu thị trường có dẫn nguồn chi tiết ở phần “Nguồn tham khảo”; số tài chính là giả định dự phóng của nhóm.",
     size=9.5, color=MUTED, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER)

doc.add_page_break()

# ================= MỤC LỤC =================
para("MỤC LỤC", style=None, size=16, color=DARK, bold=True, space_after=8)
toc_items = [
    "1. Tóm tắt điều hành",
    "2. Vấn đề và nhu cầu thị trường",
    "3. Khách hàng mục tiêu",
    "4. Giải pháp: dịch vụ, giá và quy trình",
    "5. Gói dịch vụ và mô hình doanh thu lặp lại",
    "6. Giá trị khác biệt và bối cảnh cạnh tranh",
    "7. Nguồn lực, pháp lý và điều kiện vận hành",
    "8. Marketing và kênh tiếp cận khách hàng",
    "9. Tài chính dự phóng",
    "10. Rủi ro và giải pháp ứng phó",
    "11. Lộ trình triển khai",
    "12. Kết luận và khuyến nghị",
    "Nguồn tham khảo",
    "Phụ lục A. Ghi chú quá trình lựa chọn ý tưởng",
    "Phụ lục B. Ghi chú đổi thương hiệu: từ An Khang Đường đến CareTouch",
]
for t in toc_items:
    p = para(t, size=11, space_after=4)
    p.paragraph_format.left_indent = Cm(0.4)
para("Ghi chú: số trang tự động không chèn được từ code — bạn có thể dùng References → Table of Contents trong Word để tạo mục lục có số trang sau khi mở tài liệu.",
     size=9.5, color=MUTED, italic=True, space_before=8)

doc.add_page_break()

# ================= 1. TÓM TẮT =================
doc.add_heading("1. Tóm tắt điều hành", level=1)
para("CareTouch là dịch vụ chăm sóc sức khỏe bằng phương pháp chạm trị liệu của y học cổ truyền Việt Nam — gội đầu dưỡng sinh thảo dược, massage trị liệu, day ấn bấm huyệt và chườm thảo dược — hỗ trợ thư giãn, giảm mỏi cơ khớp; phục vụ khách hàng tại cơ sở và tận nhà. Dự án được phát triển từ ý tưởng “An Khang Đường” của Nhóm 3, Lớp A1K79, và được đổi thương hiệu thành CareTouch với định vị quốc tế hơn: Healing Touch, Healthy Life (lý do và so sánh chi tiết tại Phụ lục B).")
para("Dự án nhắm vào một khoảng trống rất cụ thể của thị trường: người trung niên và cao tuổi — nhóm dân số tăng nhanh nhất và sẵn chi trả nhất cho việc chăm sóc sức khỏe — chưa được các dịch vụ chạm trị liệu phục vụ một cách bài bản, an toàn và thuận tiện. Spa cao cấp hướng tới khách trẻ với giá cao; các quán massage nhỏ thì thiếu quy trình và niềm tin. CareTouch đặt mình vào chính khoảng giữa đó: trị liệu cổ truyền có kiểm chứng về an toàn, giá minh bạch 120.000–320.000 đồng mỗi phiên, và phục vụ tận giường, tận nhà.")
para("Ba con số nói lên tiềm năng của thị trường: Việt Nam có 16,1 triệu người trên 60 tuổi (hơn 16% dân số) năm 2025 — nhóm nước già hóa nhanh nhất châu Á; tổng chi tiêu y tế đạt 27,5 tỷ USD năm 2025 và dự báo 34,1 tỷ USD vào 2028; riêng ngành spa Việt Nam đã đạt 1,4 tỷ USD doanh thu năm 2023, thuộc top 20 thế giới. Global Wellness Institute xếp Việt Nam là thị trường wellness tăng trưởng nhanh nhất châu Á với tốc độ 15,6% mỗi năm.")
para("Về mặt vận hành, tổng nhu cầu vốn khởi đầu khoảng 165 triệu đồng, gồm 95 triệu đồng đầu tư ban đầu và 70 triệu đồng vốn lưu động bù lỗ giai đoạn đầu (phép tính chi tiết tại chương 9). Doanh thu vượt chi phí cố định từ tháng thứ năm; theo kịch bản cơ sở, toàn bộ 165 triệu đồng vốn được thu hồi trong khoảng 12 tháng, kịch bản thận trọng với tỷ lệ khách quay lại 55% hoàn vốn trong 12–14 tháng. Bên dưới là bốn trụ cột tạo nên khác biệt của CareTouch:")
for t in [
    [("Tận giường – tận nhà: ", dict(bold=True)), ("mang theo đệm, dầu xoa, khăn sạch; khách cao tuổi không cần ai đưa đón.", dict())],
    [("An toàn được kiểm soát ngay từ đầu: ", dict(bold=True)), ("mọi buổi đều bắt đầu bằng khảo sát thể trạng (hỏi bệnh nền, đo huyết áp); không an toàn thì không thực hiện.", dict())],
    [("Giá niêm yết minh bạch: ", dict(bold=True)), ("gói combo tiết kiệm 15–20%, không phát sinh chi phí ngoài bảng giá.", dict())],
    [("Sổ sức khỏe cá nhân: ", dict(bold=True)), ("ghi nhận phản ứng sau mỗi buổi để liệu trình liền mạch, có căn cứ điều chỉnh.", dict())],
]:
    bullet(t)

# (Nội dung đổi thương hiệu đã chuyển xuống Phụ lục B)

# ================= 2. VẤN ĐỀ & NHU CẦU =================
doc.add_heading("2. Vấn đề và nhu cầu thị trường", level=1)
doc.add_heading("2.1. Ba nhóm người, một vấn đề chung", level=2)
para("Chức năng vận động, xương khớp và giấc ngủ suy giảm do lao động nặng, thói quen ngồi lâu hoặc do tuổi tác — nhưng phần lớn người dân chưa được phục vụ một cách bài bản:")
bullet([("Người lao động nặng (xây dựng, nông nghiệp, vận chuyển): ", dict(bold=True)), ("đau lưng, thoái hóa khớp sau nhiều năm; thường tự xử lý bằng dầu nóng, miếng dán vì đi viện tốn thời gian.", dict())])
bullet([("Dân văn phòng: ", dict(bold=True)), ("đau cổ vai gáy, tê bì tay, mất ngủ chức năng do ngồi 8–10 tiếng mỗi ngày; có khả năng chi trả nhưng bận rộn, cần dịch vụ nhanh và tiện.", dict())])
bullet([("Người cao tuổi: ", dict(bold=True)), ("nhiều bệnh mạn tính cùng lúc (bình quân 3,5–4 bệnh mỗi lần nhập viện); quen và tin dùng y học cổ truyền nhưng đi lại khó khăn.", dict())])
para("Hệ thống y tế công lập tập trung cho khám chữa bệnh; nhu cầu chăm sóc, phục hồi và thư giãn định kỳ nằm rải rác ở các spa (hướng tới khách trẻ, giá cao) và các quán nhỏ (giá rẻ nhưng thiếu chuẩn quy trình và niềm tin). Đó chính là khoảng trống CareTouch hướng tới.")

doc.add_heading("2.2. Dân số già hóa — nhu cầu có sẵn và tăng theo từng năm", level=2)
para("Năm 2025, Việt Nam có khoảng 16,1 triệu người trên 60 tuổi, chiếm hơn 16% dân số; tốc độ già hóa được đánh giá là nhanh nhất châu Á. Dự báo đến năm 2030, số người trên 60 tuổi đạt khoảng 18 triệu. Đây là nhóm khách hàng lớn nhất lịch sử của dịch vụ chăm sóc sức khỏe tại Việt Nam — và cũng là nhóm trung thành nhất với phương pháp xoa bóp, bấm huyệt.")
chart_img("chart_aging.png")
caption("Hình 1. Dân số Việt Nam từ 60 tuổi trở lên. Nguồn: Tuổi Trẻ, 31/5/2025; dự báo 2030 theo báo cáo của nhóm.")

doc.add_heading("2.3. Chi tiêu sức khỏe tăng đều", level=2)
para("Chi y tế bình quân đầu người đạt khoảng 270 USD (khoảng 7,3 triệu đồng) mỗi năm; tổng chi tiêu y tế cả nước tăng từ 17,4 tỷ USD năm 2019 lên 27,5 tỷ USD năm 2025, dự báo đạt 34,1 tỷ USD vào năm 2028. Trong 6 tháng đầu năm 2026, lượt khám chữa bệnh có bảo hiểm y tế đạt 94,4 triệu lượt — hệ thống khám chữa bệnh vận hành hết công suất, trong khi nhu cầu chăm sóc – phục hồi chưa được đáp ứng tương xứng.")
chart_img("chart_spend.png")
caption("Hình 2. Chi tiêu y tế Việt Nam. Nguồn: VnExpress, 24/12/2025; US Commercial Service (trade.gov).")

doc.add_heading("2.4. Kinh tế wellness: Việt Nam là “ngôi sao tăng trưởng” của châu Á", level=2)
para("Theo Global Wellness Institute (GWI), kinh tế wellness toàn cầu đạt 6.300 tỷ USD năm 2023 và tăng nhanh hơn GDP thế giới. Việt Nam được GWI xếp vào nhóm thị trường tăng trưởng nhanh nhất châu Á với tốc độ 15,6% mỗi năm — cao hơn Ấn Độ (8,1%), Trung Quốc (7,6%) và Indonesia (7,5%). Riêng doanh thu spa Việt Nam đạt 1,4 tỷ USD năm 2023, từ 850 triệu USD năm 2019, đưa Việt Nam vào top 20 thị trường spa toàn cầu.")
para("Về chính sách, Thủ tướng Chính phủ đã ban hành Quyết định số 1289/QĐ-TTg ngày 28/10/2024 phê duyệt Đề án phát triển du lịch chăm sóc sức khỏe dựa trên nền tảng y học cổ truyền giai đoạn 2024–2030, định hướng đến 2035 — tín hiệu thể chế rõ ràng cho đúng phân khúc CareTouch lựa chọn.")
chart_img("chart_wellness.png", 13.8)
caption("Hình 3. Tốc độ tăng trưởng kinh tế wellness. Nguồn: Global Wellness Institute, dẫn theo Nikkei Asia, 1/2025.")

# ================= 3. KHÁCH HÀNG =================
doc.add_heading("3. Khách hàng mục tiêu", level=1)
t = doc.add_table(rows=4, cols=3)
rows = [
    ("Nhóm khách", "Đặc điểm", "Hành vi mua"),
    ("Chính: Trung niên 40–60 tuổi", "Thu nhập ổn định nhất; bắt đầu đau nhức cơ – xương khớp; thường là người quyết định dịch vụ cho cả cha mẹ già", "Chuộng gói định kỳ, doanh thu lặp lại; tiếp cận qua Zalo, Facebook, giới thiệu"),
    ("Phụ: Nhân viên văn phòng", "Đau cổ vai gáy do nghề; bận rộn", "Mua lẻ ban đầu, chuyển dần sang gói combo"),
    ("Phụ: Người cao tuổi", "Nhiều bệnh mạn tính; khó đi lại; tin tưởng y học cổ truyền", "Con cháu đặt và thanh toán hộ; trung thành cao nếu buổi đầu an toàn"),
]
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        t.rows[i].cells[j].paragraphs[0].add_run(v)
style_table(t, col_widths=[4.4, 6.0, 5.4])
caption("Bảng 1. Phân khúc khách hàng mục tiêu")
para("Về khả năng chi trả: thu nhập bình quân đầu người năm 2025 đạt 5,9 triệu đồng/tháng, tăng 9,3% so với năm 2024 (Tổng cục Thống kê — ")
p = doc.paragraphs[-1]
add_hyperlink(p, "gso.gov.vn", "https://www.gso.gov.vn/")
rich(p, [("). Với mức giá 120.000–320.000 đồng một phiên — tương đương một bữa cơm gia đình — dịch vụ nằm trong khả năng chi trả định kỳ của nhóm mục tiêu. Khảo sát thị trường cho thấy khách hàng đô thị sẵn sàng chi 250.000–500.000 đồng mỗi buổi cho liệu trình chất lượng, hoặc mua trọn gói 2,5–5 triệu đồng.", dict())])

# ================= 4. GIẢI PHÁP =================
doc.add_heading("4. Giải pháp: dịch vụ, giá và quy trình", level=1)
doc.add_heading("4.1. Danh mục dịch vụ và giá đề xuất", level=2)
para("Giá đề xuất căn theo mặt bằng thị trường đã khảo sát: dịch vụ massage tận nhà cho người lớn tuổi tại TP.HCM phổ biến 350.000–620.000 đồng mỗi phiên; spa tầm trung 350.000–400.000 đồng. CareTouch định vị ngay dưới phân khúc spa — vừa túi tiền nhóm trung niên, vẫn đảm bảo chất lượng chuyên nghiệp.")
t = doc.add_table(rows=6, cols=4)
rows = [
    ("Dịch vụ", "Thời lượng", "Tại cơ sở (đ)", "Tại nhà (đ)"),
    ("Gội đầu dưỡng sinh thảo dược", "45 phút", "120.000", "200.000"),
    ("Massage cổ – vai – gáy trị liệu", "60 phút", "180.000", "270.000"),
    ("Massage toàn thân thư giãn", "70 phút", "220.000", "320.000"),
    ("Bấm huyệt – đả thông kinh lạc", "60 phút", "200.000", "290.000"),
    ("Chườm thảo dược – xông hơi đông y", "60 phút", "180.000", "—"),
]
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        t.rows[i].cells[j].paragraphs[0].add_run(v)
style_table(t, col_widths=[6.4, 2.8, 2.9, 2.9])
caption("Bảng 2. Danh mục dịch vụ và giá đề xuất (VNĐ/phiên). Nguồn tham chiếu giá thị trường: Massagenha, 7/2025.")
para("Phục vụ tận nhà miễn phí trong bán kính 5 km; xa hơn phụ thu 20.000 đồng mỗi 5 km, và khách được khuyến nghị đặt tối thiểu 2 buổi mỗi lượt. Mức chênh lệch 80.000–100.000 đồng giữa dịch vụ tận nhà và tại cơ sở được tính từ thời gian thực tế của một ca tại nhà: 45–60 phút trị liệu cộng 30–45 phút di chuyển và chuẩn bị, tức khoảng 1,5–2 tiếng mỗi ca — nhờ đó thu nhập trên mỗi giờ làm việc của kỹ thuật viên không bị suy giảm so với phục vụ tại cơ sở. Các con số sẽ được hiệu chỉnh theo địa bàn triển khai thực tế.")

doc.add_heading("4.2. Quy trình một buổi trị liệu chuẩn", level=2)
para("Mỗi buổi trải qua bốn bước:")
for i, s in enumerate([
    [("Khảo sát thể trạng an toàn (5 phút đầu): ", dict(bold=True)), ("hỏi bệnh nền, đo huyết áp, kiểm tra chống chỉ định — không an toàn thì không thực hiện.", dict())],
    [("Đánh giá: ", dict(bold=True)), ("xác định vùng đau, tư thế sai lệch để trị liệu nhắm đúng chỗ.", dict())],
    [("Trị liệu: ", dict(bold=True)), ("massage – bấm huyệt – chườm thảo dược theo bài chuẩn 45–70 phút.", dict())],
    [("Theo dõi: ", dict(bold=True)), ("dặn dò bài tập nhẹ và ghi nhận phản ứng vào sổ sức khỏe riêng của từng khách để buổi sau điều chỉnh tốt hơn.", dict())],
], start=1):
    bullet(s)
para("Nguyên tắc xuyên suốt: an toàn là tính năng của dịch vụ, không phải khẩu hiệu. Khách có chống chỉ định sẽ được từ chối hoặc khuyên đi khám; cơ sở mua bảo hiểm trách nhiệm dịch vụ.")

# ================= 5. GÓI =================
doc.add_heading("5. Gói dịch vụ và mô hình doanh thu lặp lại", level=1)
para("Liệu trình y học cổ truyền vốn cần 5–10 buổi liên tục để có tác dụng rõ rệt — đặc điểm y khoa này trở thành nền tảng của thiết kế doanh thu: khách mua theo liệu trình, quay lại định kỳ, và dòng tiền trở nên dự đoán được.")
t = doc.add_table(rows=4, cols=3)
rows = [
    ("Cấp độ", "Giá", "Quyền lợi"),
    ("Mua lẻ", "120.000–290.000 đ/phiên", "Không ràng buộc; tích điểm thành viên từ phiên đầu"),
    ("Combo 5 buổi", "890.000 đ (tiết kiệm ~15%)", "Tự chọn kết hợp 3 dịch vụ; đặt lịch linh hoạt trong 3 tháng; 1 buổi tổng kết liệu trình"),
    ("CarePlus 10 buổi", "1.590.000 đ (tiết kiệm ~20%)", "Giá cố định 12 tháng; tích điểm 5%; ưu đãi tháng sinh nhật; cho người thân dùng chung"),
]
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        t.rows[i].cells[j].paragraphs[0].add_run(v)
style_table(t, col_widths=[3.6, 4.6, 7.6])
caption("Bảng 3. Cấp độ gói dịch vụ")
para("Theo dự phóng, khoảng 60% doanh thu tháng thứ 12 đến từ gói combo và thẻ thành viên. Tỷ lệ khách quay lại là chỉ số quan trọng nhất nhóm theo dõi từ tháng đầu tiên — mục tiêu ≥55%.")

# ================= 6. KHÁC BIỆT =================
doc.add_heading("6. Giá trị khác biệt và bối cảnh cạnh tranh", level=1)
t = doc.add_table(rows=7, cols=5)
rows = [
    ("Tiêu chí", "CareTouch", "Spa cao cấp", "Quán massage nhỏ", "Dịch vụ tại nhà tự phát"),
    ("Giá niêm yết minh bạch", "✓", "✓", "✕", "~"),
    ("Trị liệu cổ truyền theo bài chuẩn", "✓", "✕", "~", "~"),
    ("Khảo sát thể trạng an toàn trước mỗi buổi", "✓", "✕", "✕", "~"),
    ("Phục vụ tận giường – tận nhà", "✓", "✕", "~", "✓"),
    ("Sổ sức khỏe cá nhân", "✓", "✕", "✕", "✕"),
    ("Giá phù hợp chi trả định kỳ", "✓", "✕", "✓", "~"),
]
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        cell = t.rows[i].cells[j]
        cell.paragraphs[0].add_run(v)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
style_table(t, col_widths=[5.4, 2.7, 2.7, 2.7, 2.7])
caption("Bảng 4. So sánh với các nhóm dịch vụ hiện có (do nhóm thực hiện trên cơ sở khảo sát dịch vụ công khai tại TP.HCM, 2025–2026)")
para("Bốn giá trị khác biệt làm nên CareTouch: (1) tiện lợi đúng chỗ — phục vụ tận giường, tận nhà; (2) cổ truyền có kiểm chứng — khảo sát thể trạng an toàn là quy trình bắt buộc; (3) giá minh bạch — niêm yết công khai, không phát sinh; (4) sổ sức khỏe cá nhân — liệu trình được theo dõi liên tục thay vì từng buổi rời rạc.")

# ================= 7. NGUỒN LỰC & PHÁP LÝ =================
doc.add_heading("7. Nguồn lực, pháp lý và điều kiện vận hành", level=1)
para("Nhóm xác định rõ ranh giới pháp lý ngay từ đầu: CareTouch định vị là dịch vụ chăm sóc – thư giãn – hỗ trợ vận động, không phải cơ sở khám chữa bệnh. Theo Luật Khám bệnh, chữa bệnh năm 2023 và Nghị định số 96/2023/NĐ-CP hướng dẫn, các hoạt động thuộc phạm vi khám chữa bệnh (kể cả phục hồi chức năng theo nghĩa y khoa) chỉ được thực hiện tại cơ sở đã đăng ký hoạt động. Vì vậy, CareTouch chủ động giới hạn phạm vi dịch vụ để vận hành đúng khung pháp lý, cụ thể:")
bullet([("Không khám bệnh, không kê đơn, không quảng cáo chữa bệnh; không thực hiện các kỹ thuật thuộc phạm vi khám chữa bệnh. ", dict(bold=True)), ("Mọi thông tin chuyên môn do bác sĩ y học cổ truyền cộng tác đảm nhận.", dict())])
bullet([("Kỹ thuật viên thi chứng chỉ nghề ", dict(bold=True)), ("theo quy định hiện hành về massage – xoa bóp trước khi phục vụ độc lập.", dict())])
bullet([("Bảo hiểm trách nhiệm dịch vụ ", dict(bold=True)), ("cho toàn bộ buổi trị liệu tại cơ sở và tận nhà.", dict())])
bullet([("Khảo sát thể trạng bắt buộc: ", dict(bold=True)), ("khách có chống chỉ định sẽ bị từ chối hoặc khuyên đi khám — đây vừa là biện pháp an toàn, vừa là ranh giới chuyên môn.", dict())])
para("Về nhân sự, giai đoạn đầu mô hình có 6 người tham gia vận hành thường xuyên: 4 kỹ thuật viên có chứng chỉ nghề (2 người phục vụ tại cơ sở hai giường, 2 người đảm nhận các ca tại nhà) và 5 thành viên sáng lập trực tiếp luân phiên đảm nhận công việc điều phối lịch, marketing và chăm sóc khách hàng. Trong giai đoạn kiểm chứng, nhóm sáng lập không rút quỹ lương điều phối mà được ưu tiên chia từ lợi nhuận sau khi hoàn vốn; kỹ thuật viên nhận lương cứng bình quân 8 triệu đồng mỗi tháng, kèm phụ cấp theo từng ca phục vụ tại nhà trích từ phần chênh lệch giá dịch vụ tận nhà. Sang năm thứ hai, khi mở cơ sở thứ hai, dự kiến tuyển thêm một điều phối viên toàn thời gian. Chi phí đầu tư chủ yếu là cải tạo mặt bằng, dụng cụ trị liệu và quảng bá khai trương — mô hình không cần thiết bị đắt tiền vì giá trị cốt lõi nằm ở tay nghề và quy trình.")

# ================= 8. MARKETING =================
doc.add_heading("8. Marketing và kênh tiếp cận khách hàng", level=1)
para("Nhóm khách trung niên – cao tuổi không mua dịch vụ vì quảng cáo ồn ào, mà vì niềm tin; chiến lược tiếp thị vì vậy tập trung vào ba kênh chi phí thấp, chuyển đổi cao:")
bullet([("Giới thiệu truyền miệng có cơ chế: ", dict(bold=True)), ("khách giới thiệu khách nhận 1 buổi gội đầu dưỡng sinh miễn phí; đây là kênh số một trong nhóm khách cao tuổi.", dict())])
bullet([("Zalo & Facebook địa phương: ", dict(bold=True)), ("nội dung giáo dục sức khỏe ngắn (bài tập cổ vai gáy, gội đầu dưỡng sinh tại nhà) kèm đặt lịch trực tiếp; Zalo OA là kênh nhắc liệu trình và chăm sóc sau buổi.", dict())])
bullet([("Liên kết cộng đồng: ", dict(bold=True)), ("hợp tác với câu lạc bộ dưỡng sinh, câu lạc bộ người cao tuổi, và các dược sĩ quận/phường — nơi nhóm khách chính sinh hoạt hằng ngày.", dict())])
para("Nguyên tắc đo lường: mỗi kênh được gắn mã ưu đãi riêng để tính chi phí thu hút một khách hàng (CAC) và tỷ lệ khách quay lại theo nguồn — chỉ mở rộng ngân sách cho kênh nào giữ khách tốt.")

# ================= 9. TÀI CHÍNH =================
doc.add_heading("9. Tài chính dự phóng", level=1)
para("Toàn bộ số liệu trong chương này là giả định minh họa phục vụ mục đích học tập, dựng trên mô hình vận hành: một cơ sở 2 giường trị liệu với 4 kỹ thuật viên (2 người tại cơ sở, 2 người đảm nhận ca tại nhà); giá bán theo Bảng 2. Các con số sẽ được thay bằng khảo sát địa bàn thực tế khi triển khai.")
t = doc.add_table(rows=6, cols=3)
rows = [
    ("Hạng mục", "Giá trị", "Ghi chú"),
    ("Đầu tư ban đầu", "~95 triệu đồng", "Cải tạo mặt bằng, dụng cụ trị liệu, đồng phục, quảng bá khai trương"),
    ("Vốn lưu động bù lỗ (tháng 1–4)", "~70 triệu đồng", "Lỗ lũy kế 4 tháng đầu: 26 + 18 + 10 + 4 = 58 triệu đồng; dự trù thêm 12 triệu đồng làm đệm an toàn"),
    ("Tổng nhu cầu vốn khởi đầu", "~165 triệu đồng", "95 triệu đồng đầu tư + 70 triệu đồng vốn lưu động"),
    ("Chi phí cố định", "~48 triệu đồng/tháng", "Thuê mặt bằng 12 triệu; lương 4 kỹ thuật viên ~32 triệu (lương cứng ~8 triệu/người, kèm phụ cấp ca tại nhà); điện nước, vật tư, khấu hao còn lại"),
    ("Doanh thu tháng thứ 12", "~88 triệu đồng/tháng", "Tương đương ~350 khách/tháng, trong đó ~60% từ gói combo và thẻ thành viên"),
]
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        t.rows[i].cells[j].paragraphs[0].add_run(v)
style_table(t, col_widths=[4.2, 4.2, 7.4])
caption("Bảng 5. Các giả định tài chính chính")
chart_img("chart_revenue.png", 15.2)
caption("Hình 4. Doanh thu dự phóng 12 tháng đầu vận hành (giả định minh họa của nhóm)")
para("Phép tính vốn lưu động: trong bốn tháng đầu, doanh thu còn thấp hơn chi phí cố định nên dự án lỗ 26 triệu đồng vào tháng 1, 18 triệu đồng vào tháng 2, 10 triệu đồng vào tháng 3 và 4 triệu đồng vào tháng 4; cộng lại, lỗ lũy kế bốn tháng đầu là 58 triệu đồng. Nếu chỉ có 95 triệu đồng đầu tư ban đầu, dòng tiền của dự án sẽ âm ngay từ tháng thứ hai. Vì vậy, nhu cầu vốn khởi đầu gồm 95 triệu đồng đầu tư cộng tối thiểu 58 triệu đồng vốn lưu động, được nhóm dự trù lên 70 triệu đồng để có đệm an toàn; tổng nhu cầu vốn khởi đầu là 165 triệu đồng.")
para("Phép tính hoàn vốn: kể từ tháng thứ năm, lãi ròng hàng tháng lần lượt là 2, 8, 14, 20, 25, 30, 35 và 40 triệu đồng (doanh thu trừ chi phí cố định 48 triệu đồng); lũy kế từ tháng 5 đến tháng 12 đạt 174 triệu đồng, lớn hơn tổng vốn 165 triệu đồng — tức toàn bộ vốn được thu hồi trong khoảng 12 tháng. Riêng phần đầu tư ban đầu 95 triệu đồng được thu hồi vào khoảng tháng thứ 10, khi lũy kế lãi ròng đạt 99 triệu đồng. Nếu tỷ lệ khách quay lại chỉ đạt ngưỡng 55%, hoàn vốn sẽ trôi về khoảng 12–14 tháng theo kịch bản thận trọng.")
para("Điểm nghẽn đầu tiên của mô hình là công suất nhân sự. Một kỹ thuật viên phục vụ được 5–6 khách tại cơ sở hoặc tối đa 4 ca tại nhà mỗi ngày (mỗi ca tại nhà chiếm 1,5–2 tiếng); bốn kỹ thuật viên cho công suất tối đa khoảng 20 lượt mỗi ngày, tương đương 450–500 lượt mỗi tháng với 26 ngày hoạt động. Doanh thu tháng thứ 12, tương đương khoảng 350 khách, đã sử dụng 70–80% công suất này; vì vậy bước tăng trưởng tiếp theo phải đến từ tuyển thêm nhân sự hoặc mở cơ sở thứ hai — đây chính là lý do của lộ trình mở rộng ở chương 11.")

# ================= 10. RỦI RO =================
doc.add_heading("10. Rủi ro và giải pháp ứng phó", level=1)
t = doc.add_table(rows=5, cols=3)
rows = [
    ("Rủi ro", "Mức độ", "Giải pháp ứng phó"),
    ("Ranh giới pháp lý với “khám chữa bệnh”", "Cao", "Định vị dịch vụ chăm sóc – thư giãn – hỗ trợ vận động; không khám bệnh, không kê thuốc, không thực hiện kỹ thuật thuộc phạm vi khám chữa bệnh (Luật Khám bệnh, chữa bệnh 2023; Nghị định 96/2023/NĐ-CP); bác sĩ y học cổ truyền cộng tác đảm nhận chuyên môn; kỹ thuật viên thi chứng chỉ nghề"),
    ("Sự cố với khách có bệnh nền", "Cao", "Khảo sát bắt buộc buổi đầu (hỏi bệnh nền, đo huyết áp); từ chối hoặc khuyên đi khám khi có chống chỉ định; mua bảo hiểm trách nhiệm dịch vụ"),
    ("Nhân lực và tay nghề", "Trung bình", "Đào tạo nội bộ theo giáo trình chuẩn, kèm cặp thực tế; chất lượng đo bằng đánh giá sau mỗi buổi; người dạy nghề giữ cổ phần nhỏ"),
    ("Cạnh tranh về giá", "Trung bình", "Không đua giá rẻ nhất; cạnh tranh bằng quy trình an toàn và niềm tin; gói combo và CarePlus giữ khách dài hạn"),
]
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        t.rows[i].cells[j].paragraphs[0].add_run(v)
style_table(t, col_widths=[4.4, 2.2, 9.2])
caption("Bảng 6. Danh mục rủi ro và giải pháp ứng phó")

# ================= 11. LỘ TRÌNH =================
doc.add_heading("11. Lộ trình triển khai", level=1)
for i, s in enumerate([
    [("Giai đoạn 1 (tháng 1–6) — Kiểm chứng: ", dict(bold=True)), ("vận hành cơ sở đầu tiên, chuẩn hóa quy trình khảo sát thể trạng và đào tạo, đo tỷ lệ khách quay lại từ tháng đầu.", dict())],
    [("Giai đoạn 2 (tháng 7–12) — Tăng trưởng: ", dict(bold=True)), ("đạt 88 triệu đồng/tháng doanh thu, ≥55% khách quay lại, hoàn tất hồ sơ pháp lý và chứng chỉ nghề.", dict())],
    [("Giai đoạn 3 (năm 2) — Mở rộng: ", dict(bold=True)), ("mở cơ sở thứ hai, xây dựng Zalo Mini App đặt lịch và nhắc liệu trình, huy động vốn mở rộng.", dict())],
    [("Giai đoạn 4 (năm 3) — Chuỗi hóa: ", dict(bold=True)), ("nhượng quyền 3–5 cơ sở vệ tinh; xây dựng chuẩn đào tạo CareTouch Academy cho kỹ thuật viên.", dict())],
], start=1):
    bullet(s)
para("Nguyên tắc xuyên suốt: chỉ mở cơ sở mới khi cơ sở hiện tại đạt ≥55% khách quay lại và quy trình an toàn đã được kiểm chứng — tăng trưởng đi theo chất lượng, không đi trước chất lượng.")

# ================= 12. KẾT LUẬN =================
doc.add_heading("12. Kết luận và khuyến nghị", level=1)
para("Ý tưởng CareTouch đứng vững trên ba nền tảng rất cụ thể. Thứ nhất, thị trường người cao tuổi Việt Nam lớn dần theo từng năm và được các định chế quốc tế (GWI) đánh giá là tăng trưởng nhanh nhất châu Á. Thứ hai, xoa bóp – bấm huyệt là phương pháp người dân đã tin dùng từ lâu; doanh nghiệp không phải tạo thói quen tiêu dùng mới. Thứ ba, dịch vụ chuyên nghiệp phục vụ tận nhà cho nhóm trung niên – cao tuổi với giá hợp lý hiện vẫn là khoảng trống cạnh tranh thực sự.")
para("Khuyến nghị của nhóm: bắt đầu với quy mô nhỏ để kiểm chứng quy trình — mở buổi trị liệu đầu tiên, phục vụ thật tốt nhóm khách đầu tiên, đo tỷ lệ quay lại từ tháng đầu; hoàn thiện hồ sơ pháp lý và quy trình khảo sát thể trạng an toàn trước khi mở rộng. Mục tiêu năm đầu tiên là tỷ lệ khách quay lại từ 55% trở lên và doanh thu 88 triệu đồng/tháng vào tháng thứ 12 theo dự phóng. Tổng nhu cầu vốn khởi đầu 165 triệu đồng được thiết kế để dự án an toàn về dòng tiền trong suốt giai đoạn kiểm chứng. Với khoản vốn gọi đề xuất minh họa 500 triệu đồng, nhóm dự kiến mở rộng lên 2 cơ sở kèm nền tảng đặt lịch Zalo Mini App trong năm thứ hai.")

# ================= NGUỒN =================
doc.add_heading("Nguồn tham khảo", level=1)
para("Các nguồn dưới đây được trích dẫn trực tiếp trong báo cáo; đường link nhấn được (Ctrl+Click để mở).", size=10.5, color=MUTED)
sources = [
    ("Tuổi Trẻ (31/5/2025). “Năm 2025, Việt Nam có 16% dân số là người cao tuổi, 1 người gánh 3–4 bệnh mạn tính”.",
     "https://tuoitre.vn/nam-2025-viet-nam-co-16-dan-so-la-nguoi-cao-tuoi-1-nguoi-ganh-3-4-benh-man-tinh-20250531122410678.htm"),
    ("Tuổi Trẻ (31/5/2025). “Việt Nam đang có tốc độ già hóa nhanh nhất châu Á”.",
     "https://tuoitre.vn/nld/viet-nam-dang-co-toc-do-gia-hoa-nhanh-nhat-chau-a-196250531143638561.htm"),
    ("VnExpress (24/12/2025). “Người Việt trung bình chi hơn 7 triệu đồng một năm cho y tế”.",
     "https://vnexpress.net/chi-phi-cho-y-te-cua-nguoi-viet-nam-tang-thang-dung-4997605.html"),
    ("Global Wellness Institute (2024). Global Wellness Economy Monitor 2024 (PDF).",
     "https://globalwellnessinstitute.org/wp-content/uploads/2024/11/WellnessEconMonitor2024PDF.pdf"),
    ("Nikkei Asia (28/1/2025). “Vietnam fires up wellness tourism to rival Thailand, Indonesia”.",
     "https://asia.nikkei.com/Business/Health-Care/Vietnam-fires-up-wellness-tourism-to-rival-Thailand-Indonesia"),
    ("US Commercial Service — Trade.gov. Vietnam Healthcare (Country Commercial Guide).",
     "https://www.trade.gov/country-commercial-guides/vietnam-healthcare"),
    ("Massagenha.com (7/2025). “Bảng giá massage cho người lớn tuổi tận nhà tại TPHCM”.",
     "https://massagenha.com/dich-vu/massage-nguoi-lon-tuoi/"),
    ("Tổng cục Thống kê. Thu nhập bình quân đầu người năm 2025 (tăng 9,3%).",
     "https://www.gso.gov.vn/"),
    ("Thủ tướng Chính phủ. Quyết định 1289/QĐ-TTg ngày 28/10/2024 phê duyệt Đề án phát triển du lịch chăm sóc sức khỏe dựa trên nền tảng y học cổ truyền (tra cứu tại Cổng thông tin điện tử Chính phủ).",
     "https://vanban.chinhphu.vn/"),
    ("Luật Khám bệnh, chữa bệnh năm 2023 và Nghị định 96/2023/NĐ-CP hướng dẫn Luật Khám bệnh, chữa bệnh (tra cứu tại Cổng thông tin điện tử Chính phủ).",
     "https://vanban.chinhphu.vn/"),
]
for i, (txt, url) in enumerate(sources, 1):
    p = doc.add_paragraph()
    rich(p, [(f"[{i}] ", dict(bold=True, color=MUTED)), (txt + " ", dict(size=10.5))])
    add_hyperlink(p, url, url, size=10)
    p.paragraph_format.space_after = Pt(5)
para("Số liệu lượt khám chữa bệnh có BHYT 6 tháng đầu năm 2026 (93,7 → 94,4 triệu lượt) do nhóm thu thập trong quá trình làm bài; khảo sát giá spa công khai (KKday, Jackfruit Adventure, 2025–2026) dùng cho mục 7.",
     size=9.5, color=MUTED, italic=True)

# ================= PHỤ LỤC =================
doc.add_heading("Phụ lục A. Ghi chú quá trình lựa chọn ý tưởng", level=1)
para("Nhóm đã sinh 5 ý tưởng ban đầu quanh vấn đề “chăm sóc sức khỏe chủ động, tiện lợi cho người già”: hỗ trợ đi khám bệnh; khám sức khỏe định kỳ tại nhà; thiết bị nhắc uống thuốc; massage – xoa bóp, bấm huyệt và đả thông kinh lạc theo cổ truyền; và dịch vụ nấu ăn chế độ cho người già. Ba ý tưởng khả thi nhất được đưa vào bảng đánh giá 5 tiêu chí, thang điểm 5:")
t = doc.add_table(rows=7, cols=4)
rows = [
    ("Tiêu chí", "Ý tưởng 1: Khám định kỳ tại nhà", "Ý tưởng 2: Thiết bị nhắc uống thuốc", "Ý tưởng 3: Massage – trị liệu cổ truyền"),
    ("Mức độ đáp ứng nhu cầu khách hàng", "4", "3,5", "4,5"),
    ("Giá trị và sự khác biệt", "4", "4", "4"),
    ("Tiềm năng thị trường", "3,5", "3", "4,5"),
    ("Tính khả thi", "3", "3,5", "4,5"),
    ("Phù hợp với năng lực nhóm", "3", "4", "5"),
    ("Điểm trung bình", "3,5", "3,6", "4,5"),
]
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        cell = t.rows[i].cells[j]
        cell.paragraphs[0].add_run(v)
        if j > 0: cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
style_table(t, col_widths=[5.2, 3.7, 3.7, 3.6])
caption("Bảng 7. Bảng đánh giá và so sánh 3 ý tưởng (thang điểm 5)")
para("Ý tưởng 3 đạt điểm cao nhất trên tổng thể và được chọn, với ba lý do chính: xu hướng chăm sóc sức khỏe chủ động giúp y học cổ truyền ngày càng được ưa chuộng nhờ an toàn, không dùng thuốc; mô hình dịch vụ có biên lợi nhuận tốt và doanh thu ổn định nhờ khách hàng dùng theo liệu trình 5–10 buổi; và rào cản chứng chỉ hành nghề cùng yếu tố chuyên môn chuẩn y khoa tạo lợi thế khác biệt bền vững so với các spa thư giãn thông thường.")

# ================= PHỤ LỤC B =================
doc.add_heading("Phụ lục B. Ghi chú đổi thương hiệu: từ An Khang Đường đến CareTouch", level=1)
para("Tên ban đầu “An Khang Đường” mang ý nghĩa tốt đẹp — “An Khang” là mong ước sức khỏe bình an, “Đường” theo cách đặt tên các nhà thuốc, cơ sở y học cổ truyền xưa. Tuy nhiên, khi dự án bước vào giai đoạn chuẩn bị gọi vốn và mở rộng, nhóm nhận thấy tên gọi này bộc lộ ba hạn chế: khó phát âm với đối tác quốc tế và khách hàng trẻ; dài, khó nhớ trên các kênh số như Zalo, Facebook hay ứng dụng đặt lịch; và hình ảnh “cửa hàng thuốc cổ” dễ khiến khách hàng nhầm lẫn về phạm vi dịch vụ.")
p = para("Sau khi đánh giá nhiều phương án, nhóm chọn tên ")
rich(doc.paragraphs[-1], [("CareTouch", dict(bold=True, color=ACCENT)),
     (" — ghép hai chữ “Care” (sự chăm sóc chu đáo) và “Touch” (cái chạm trị liệu). Tên gọi nói đúng bản chất dịch vụ: hỗ trợ thư giãn, giảm mỏi cơ khớp bằng đôi tay, đồng thời gợi cảm giác tận tâm. Slogan kèm theo: ", dict()),
     ("“CareTouch – Healing Touch, Healthy Life”", dict(italic=True, bold=True)),
     (" (Chạm trị liệu — Sống khỏe mỗi ngày).", dict())])
para("Bảng dưới đây tóm tắt lý do đổi tên:")
t = doc.add_table(rows=5, cols=3)
rows = [
    ("Tiêu chí", "An Khang Đường", "CareTouch"),
    ("Khả năng ghi nhớ", "Dài, dễ nhầm với nhà thuốc", "Hai âm tiết, đọc đúng ngay lần đầu"),
    ("Khách hàng quốc tế", "Khó phát âm, khó viết", "Tiếng Anh đơn giản, phù hợp khách du lịch chăm sóc sức khỏe"),
    ("Hình ảnh số", "Kém phù hợp làm app/website", "Thân thiện logo, ứng dụng, mạng xã hội"),
    ("Thông điệp", "Mong ước chung chung", "Nói đúng phương pháp: chăm sóc bằng đôi tay"),
]
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        t.rows[i].cells[j].paragraphs[0].add_run(v)
style_table(t, col_widths=[3.6, 5.6, 6.6])
caption("Bảng 8. So sánh tên gọi cũ và mới")

# footer
footer = sec.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = fp.add_run("CareTouch — Healing Touch, Healthy Life  ·  Báo cáo dự án khởi nghiệp · Nhóm 3 – Lớp A1K79")
r.font.size = Pt(9); r.font.color.rgb = MUTED; r.font.name = "Segoe UI"

doc.save(os.path.join(HERE, "CareTouch-bao-cao.docx"))
print("DOCX saved OK")
