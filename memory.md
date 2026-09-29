# MEMORY — Dự án CareTouch (Nhóm 3 · Tổ 14 · Lớp A1K79)

> Ghi chú trước khi compact. Thư mục làm việc: `C:\Users\Tu_DZ\Downloads\bai tap cua Tram`

## 1. Dự án
- Khởi nghiệp sinh viên — môn **Thực tập NCKH Đổi mới sáng tạo Khởi nghiệp, Bài 3: Lựa chọn ý tưởng kinh doanh**. Buổi thực tập: ca chiều 1, thứ 7, 26/9/2026.
- Dịch vụ: massage – xoa bóp, bấm huyệt, gội đầu dưỡng sinh, chườm thảo dược (YHCT) — tại cơ sở + tận nhà, khách chính trung niên 40–60.
- **Đổi thương hiệu: An Khang Đường → CareTouch** — "Healing Touch, Healthy Life" (lý do ở Phụ lục B báo cáo).
- Thành viên: Trịnh Hương Thảo 2401634, Bạch Thùy Trâm 2401679, Bùi Thị Anh Trâm 2401681, Phạm Quốc Việt 2401742, Viengkeo Shengchan 2401761.

## 2. Sản phẩm bàn giao (trong thư mục làm việc)
| File | Nội dung |
|---|---|
| `CareTouch-deck.pptx` | 14 slide pitch nhà đầu tư, brand riêng, có speaker notes |
| `CareTouch-bao-cao.docx` | 12 chương + Nguồn tham khảo (9 nguồn hyperlink) + Phụ lục A (lựa chọn ý tưởng) + Phụ lục B (đổi thương hiệu) |
| `CareTouch-slides.html` | HTML deck tự chứa (Chart.js inline 245KB+), dựa trên bản ở `C:\Users\Tu_DZ\Downloads\CareTouch-slides.html` (bản user ưng nhất) |
| `save.php` + `content.json` | đi kèm HTML khi up cPanel — lưu chỉnh sửa chung cho nhóm |
| `assets/` | bg_cover, bg_close, logo, 4 chart PNG, chart.umd.min.js |
| Script build | `make_assets.py` (assets+charts), `build_pptx.js` (pptxgenjs, cần NODE_PATH=C:\Users\Tu_DZ\AppData\Roaming\npm\node_modules), `fix_ppr.py` (sửa schema pPr sau build — làm XML đổi prefix p: → ns0:), `add_anim.py` (chèn hiệu ứng: CHẠY SAU CÙNG), `build_docx.py` (python-docx) |
| `deck_template.html` | nguồn của HTML hiện tại (phiên bản mới có stepper — **KHÔNG dùng, user đã chê**; file chạy thật là CareTouch-slides.html vá từ base) |

## 3. Brand tokens
- Màu: DARK `#0B3B34`, PRIMARY teal `#0E7C6B`, TEAL_MID `#2AA88F`, TEAL_SOFT `#9ED9CC`, TINT `#E8F5F1`, ACCENT cam `#E8590C`, ACCENT_SOFT `#FDEBDD`, TEXT `#16302B`, MUTED `#5C7370`.
- Motif: 2 vòng tròn chồng nhau (Care + Touch). Font: Segoe UI (toàn bộ tiếng Việt), Georgia italic chỉ cho tagline tiếng Anh. PPTX canvas 13.33×7.5".

## 4. Số liệu & nguồn (link đã verify 200)
- 16,1 triệu người 60+ (16% dân số, 2025) + già hóa nhanh nhất châu Á — Tuổi Trẻ 31/5/2025 (2 link trong tài liệu cũ).
- Chi y tế 17,4 (2019) → 27,5 (2025) → 34,1 tỷ USD (2028F) — VnExpress 24/12/2025.
- Wellness toàn cầu 6.300 tỷ USD (2023); **Việt Nam +15,6%/năm cao nhất châu Á** (Ấn Độ 8,1%, TQ 7,6%, Indo 7,5%) — GWI Wellness Economy Monitor 2024 PDF.
- Spa VN 1,4 tỷ USD 2023, top 20 TG — Nikkei Asia 28/1/2025 (URL constructed: `https://asia.nikkei.com/Business/Health-Care/Vietnam-fires-up-wellness-tourism-to-rival-Thailand-Indonesia`).
- Thu nhập BQ 5,9 tr/đ/tháng (+9,3%) — Tổng cục Thống kê (link gso.gov.vn).
- Quyết định 1289/QĐ-TTg 28/10/2024 (du lịch YHCT) + Luật KBBC 2023 & NĐ 96/2023/NĐ-CP — link vanban.chinhphu.vn.
- Massagenha 7/2025: giá thị trường tận nhà TP.HCM 350–620k/phiên.

## 5. Nội dung đã đồng bộ v1.1 (theo phản biện AI + yêu cầu user)
- Thuật ngữ: "massage" thống nhất (bỏ "matxa"); "Tự chọn kết hợp 3 dịch vụ"; "Tại nhà tự phát"; "dịch vụ chăm sóc chuyên nghiệp"; bỏ ký hiệu "K" → số đầy đủ `120.000–290.000 đ` / `890.000 đ` / `1.590.000 đ (~20%)`.
- "sàng lọc y khoa" → **"khảo sát thể trạng an toàn"**; pháp lý nêu Luật KBBC 2023 + NĐ 96/2023/NĐ-CP; định vị "chăm sóc – thư giãn – hỗ trợ vận động".
- **Tài chính (phép tính hiển thị)**: lỗ lũy kế T1–T4 = 26+18+10+4 = 58 tr → vốn lưu động 70 tr → **tổng vốn khởi đầu 165 tr** (=95 đầu tư + 70 VC); lãi ròng T5–T12 lũy kế 174 tr ≥ 165 → hoàn vốn đủ ~12 tháng (thận trọng 55%: 12–14 tháng). Doanh thu T12 = 88 tr (~350 khách), chi phí cố định 48 tr, hòa vốn T5 (T4: 44 < 48 < T5: 50).
- **Giá tận nhà nâng**: 200/270/320/290k (chênh lệch 80–100k bù 1,5–2 tiếng/ca: 45–60' trị liệu + 30–45' di chuyển); tối thiểu 2 buổi/lượt.
- **Nhân sự**: 4 KTV lương cứng ~8tr + phụ cấp ca tại nhà; 5 sáng lập trực tiếp điều phối (chưa rút lương điều phối); năm 2 tuyển thêm điều phối viên. Trần công suất ~20 lượt/ngày ≈ 450–500 lượt/tháng; T12 dùng 70–80%.
- "Hóa giải" → **"Giải pháp ứng phó"** (giữ chữ "Rủi ro").
- DOCX: chương đổi tên chuyển thành Phụ lục B; đánh số lại 1–12; Bảng 1–8; phụ lục A = Bảng 7, phụ lục B = Bảng 8.

## 6. HTML (CareTouch-slides.html) — tính năng & cách test
- Engine: v1.0 đơn giản (showSlide toggle active, auto-reveal bằng CSS `.reveal d1-d6`), chart lazy-init qua CHART_BUILDERS + animateIn.
- **Ẩn/hiện điều khiển**: tự ẩn khi bấm phím hoặc chuột đứng yên 2,6s; hiện khi di chuột; phím **H** toggle, **F** fullscreen; **E** mở panel sửa.
- **Sửa chữ**: nút "✎ Sửa nội dung" góc trái dưới → contenteditable → Lưu với mã **1234321**. Có save.php: POST JSON {code, edits} → ghi content.json (đã sanitize script/iframe/on*). Không có PHP → fallback localStorage (`caretouch_edits_v1`). Eid = `s<slideIdx>-<n>` do assignIds() gán.
- Đã test OK: điều hướng tới/lùi, chrome-hide, H, lưu+persist sau reload, chart label slide 3 (plugin valueLabels tự viết: `afterDatasetsDraw` + `_showLabels:true` + `layout.padding.top`).
- ⚠️ Bài học đã gặp: khi vá JS phải node --check cú pháp (đã từng mất 1 dấu `}` làm cả script điều hướng chết). Cache trình duyệt khi test: thêm `?v=N` vào URL.

## 6b. Hiệu ứng trình chiếu PPTX (v1.2, add_anim.py)
- Entrance "Float In" (fade 500ms + nhích lên 0.018×chiều cao) theo NHÓM: mỗi lần bấm chuột hiện 1 khối, shape con lệch 80–150ms (cascade).
- Tiêu đề/kicker/motif/số trang/dòng nguồn: luôn hiển thị (không animate). S1 bìa + S14 kết: TỰ chạy khi mở slide (delay 0), không cần bấm.
- Chuyển slide: fade 600ms (mc:AlternateContent p14:dur + fallback).
- LƯU Ý: fix_ppr.py đổi namespace prefix → add_anim.py tự dò prefix (`<ns0:sld>`) trước khi chèn `<p:timing>` + `<p:transition>` trước `</ns0:sld>`.
- Chuỗi build lại PPTX: `node build_pptx.js` → `python fix_ppr.py CareTouch-deck.pptx` → `python add_anim.py`.
- Plan nhóm shape theo slide id nằm trong PLAN dict của add_anim.py (shape id từ python-pptx).

## 7. Git (trong thư mục làm việc, branch main) — v1.2 = 64998dc (đồng bộ) + commit hiệu ứng
- `161816e` — **Phiên bản 1.0** (trước phản biện)
- `37519ca` — **Phiên bản 1.1** (hiện tại)
- `.gitignore`: `~$*`, `render/`, `render_sync/`, `.playwright-mcp/`
- Quay lại: `git checkout 161816e -- .` hoặc `git checkout 161816e` (detached).

## 8. Sở thích/user preference đã chốt
- Chỉ lấy NỘI DUNG tài liệu cũ, thiết kế làm mới hoàn toàn (không ảnh cũ).
- HTML: KHÔNG cần hiệu ứng hiện từng phần (stepper) — user chê, đã bỏ; KHÔNG cần tối ưu mật độ/khoảng trắng nữa — user nói "thôi, không cần làm gì thêm" về chuyện này.
- Bản HTML chuẩn = bản ở `C:\Users\Tu_DZ\Downloads\CareTouch-slides.html` + 3 thay đổi (sync nội dung, sửa chữ, ẩn/hiện) — đã làm xong.
- "Rủi ro" giữ nguyên, "hóa giải" → "giải pháp ứng phó".
- Báo cáo Word = nộp lớp/đọc trước; thuyết trình dùng PPTX hoặc HTML (không cần lập thêm).
- PPTX: user ok, không sửa thêm layout.

## 9. Môi trường
- Windows, Git Bash. Python 3.11 + python-docx, python-pptx, matplotlib, PIL, pymupdf.
- LibreOffice: dùng full path `"C:\Program Files\LibreOffice\program\soffice.exe"`.
- Render PNG: soffice → PDF → pymupdf (không có poppler/pdftoppm).
- Node v24 + pptxgenjs global (NODE_PATH cần set). Test HTML qua `python -m http.server` (file:// bị chặn).
- Visual-judge subagent KHÔNG khả dụng (provider error) → tự inspect ảnh.

## 10. Việc còn treo / lưu ý
- v1.2 đã rà soát đồng bộ HTML↔PPTX theo CẢ ẢNH (14 cặp, render LibreOffice vs screenshot Playwright): pass.
- HTML slide 10: legend Chart.js tự vẽ bằng HTML (span line + dashed) vì legend gốc render ô đen.
- Counter slide 2 HTML: đặt sẵn text đích (16,1 / 94,4), JS vẫn đếm từ 0 khi mở slide.
- Không có việc dở nào. Nếu user yêu cầu sửa tiếp: sửa cả 3 nguồn (build_pptx.js / build_docx.py / CareTouch-slides.html trực tiếp) để giữ đồng bộ.
- File gốc nhóm (`bao-cao-an-khang-duong.docx`, `slide-an-khang-duong.pptx`, `Tài liệu không có tiêu đề.docx`) là read-only, KHÔNG được sửa.
- PPTX slides chú ý: slide 2 panel tối (3 stats), slide 10 annotation "T4: 44 < 48 < T5: 50", speaker notes đã cập nhật số 165tr/4 KTV.
- Mã bảo mật sửa chữ mặc định: **1234321** (user chọn); khuyên đổi trong save.php khi up host thật.
