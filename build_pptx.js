/* CareTouch — investor deck builder (pptxgenjs)
   Brand: Healing Teal + Warm Touch. Motif: paired circles (Care + Touch). */
const pptxgen = require("pptxgenjs");
const path = require("path");

const A = p => path.join(__dirname, "assets", p);

// ---------- palette ----------
const DARK = "0B3B34", DARK2 = "0E4A41";
const PRIMARY = "0E7C6B", TEALMID = "2AA88F", TEALSOFT = "9ED9CC";
const TINT = "E8F5F1", TINT2 = "F4FAF8";
const ACCENT = "E8590C", ACCENTSOFT = "FDEBDD";
const TEXT = "16302B", MUTED = "5C7370", BORDER = "DCEBE6";
const W = 13.33, H = 7.5, M = 0.55;
const FONT = "Segoe UI";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Nhóm 3 - Tổ 14 - Lớp A1K79";
pres.title = "CareTouch — Healing Touch, Healthy Life";

// ---------- helpers (fresh objects every call) ----------
const shadow = () => ({ type: "outer", color: "0B3B34", blur: 8, offset: 2, angle: 90, opacity: 0.13 });
const softShadow = () => ({ type: "outer", color: "0B3B34", blur: 5, offset: 1, angle: 90, opacity: 0.09 });

function motif(s, x = 12.35, y = 0.42, r = 0.17) {
  s.addShape("ellipse", { x: x, y: y, w: r * 2, h: r * 2, fill: { color: TEALSOFT, transparency: 25 }, line: { type: "none" } });
  s.addShape("ellipse", { x: x + r * 0.85, y: y + r * 0.25, w: r * 1.7, h: r * 1.7, fill: { color: ACCENT, transparency: 12 }, line: { type: "none" } });
}
function kicker(s, txt, opts = {}) {
  s.addText(txt.toUpperCase(), Object.assign({
    x: M, y: 0.42, w: 9.5, h: 0.32, fontSize: 12, bold: true, charSpacing: 3,
    color: TEALMID, fontFace: FONT, margin: 0,
  }, opts));
}
function title(s, txt, opts = {}) {
  s.addText(txt, Object.assign({
    x: M, y: 0.74, w: 11.6, h: 0.85, fontSize: 27, bold: true,
    color: TEXT, fontFace: FONT, margin: 0, valign: "top",
  }, opts));
}
function src(s, runs, y = 7.08) {
  // runs: [{t, url?}] — url parts become clickable links
  const items = runs.map(r => ({
    text: r.t,
    options: Object.assign(
      { color: MUTED, fontFace: FONT, fontSize: 10.5, italic: true },
      r.url ? { hyperlink: { url: r.url }, color: PRIMARY, italic: false } : {}
    ),
  }));
  s.addText(items, { x: M, y: y, w: 12.2, h: 0.3, margin: 0, valign: "middle" });
}
function pageNo(s, n) {
  s.addText(String(n).padStart(2, "0"), { x: 12.62, y: 7.06, w: 0.5, h: 0.3, fontSize: 10.5, color: MUTED, fontFace: FONT, align: "right", margin: 0 });
}
function card(s, x, y, w, h, fill = "FFFFFF", useShadow = true, radius = 0.09) {
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: radius, fill: { color: fill },
    line: { color: BORDER, width: 0.75 }, shadow: useShadow ? shadow() : undefined,
  });
}
function band(s, txt, y = 6.28, h = 0.62, fill = TINT, textColor = TEXT) {
  s.addShape("rect", { x: 0, y, w: W, h, fill: { color: fill }, line: { type: "none" } });
  s.addText(txt, { x: M, y, w: W - 2 * M, h, fontSize: 13.5, color: textColor, fontFace: FONT, valign: "middle", margin: 0 });
}
const bu = () => ({ code: "25B8", indent: 12 });

/* ============================== S1 · COVER ============================== */
{
  const s = pres.addSlide();
  s.background = { path: A("bg_cover.png") };
  s.addText("HỒ SƠ DỰ ÁN KHỞI NGHIỆP  ·  THÁNG 9/2026", {
    x: 0.9, y: 1.15, w: 8, h: 0.35, fontSize: 13, bold: true, charSpacing: 4,
    color: TEALSOFT, fontFace: FONT, margin: 0,
  });
  s.addText([
    { text: "Care", options: { color: "FFFFFF" } },
    { text: "Touch", options: { color: "F7965A" } },
  ], { x: 0.86, y: 1.62, w: 9, h: 1.5, fontSize: 76, bold: true, fontFace: FONT, margin: 0, charSpacing: 1 });
  s.addText("Healing Touch, Healthy Life", {
    x: 0.9, y: 3.18, w: 8, h: 0.55, fontSize: 24, italic: true, color: TEALSOFT,
    fontFace: "Georgia", margin: 0,
  });
  s.addText(
    "Chạm trị liệu y học cổ truyền — gội đầu dưỡng sinh, massage trị liệu, bấm huyệt, chườm thảo dược. Phục vụ tại cơ sở và tận nhà cho người trung niên và cao tuổi.",
    { x: 0.9, y: 3.95, w: 7.2, h: 1.1, fontSize: 15, color: "CFE6DF", fontFace: FONT, margin: 0, lineSpacingMultiple: 1.25 }
  );
  s.addShape("line", { x: 0.9, y: 5.9, w: 3.2, h: 0, line: { color: TEALMID, width: 1.5 } });
  s.addText([
    { text: "Nhóm 3 · Tổ 14 · Lớp A1K79", options: { bold: true, color: "FFFFFF", breakLine: true } },
    { text: "Thực tập Nghiên cứu khoa học — Đổi mới sáng tạo Khởi nghiệp", options: { color: "9EC9BF", breakLine: true } },
    { text: "Tiền thân: An Khang Đường", options: { color: "9EC9BF", italic: true } },
  ], { x: 0.9, y: 6.05, w: 8, h: 1.1, fontSize: 13, fontFace: FONT, margin: 0, paraSpaceAfter: 4 });
  s.addNotes("Chào thầy/cô và các nhà đầu tư. CareTouch — tên mới của dự án An Khang Đường — dịch vụ chạm trị liệu y học cổ truyền tại cơ sở và tận nhà.");
}

/* ============================== S2 · VẤN ĐỀ ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Vấn đề thị trường");
  title(s, "Già hóa nhanh nhất châu Á — nhu cầu chăm sóc chưa được phục vụ bài bản");
  motif(s); pageNo(s, 2);

  const rows = [
    ["Người lao động nặng", "Đau lưng, thoái hóa khớp sau nhiều năm. Thói quen tự xử lý bằng dầu nóng, miếng dán vì đi viện tốn thời gian."],
    ["Dân văn phòng", "Đau cổ vai gáy, tê bì tay, mất ngủ chức năng do ngồi 8–10 tiếng/ngày. Có khả năng chi trả nhưng bận rộn, cần dịch vụ nhanh, tiện."],
    ["Người cao tuổi", "Bình quân 3,5–4 bệnh mạn tính cùng lúc. Tin dùng y học cổ truyền nhưng đi lại khó khăn, cần người đến tận nơi."],
  ];
  let y = 1.95;
  rows.forEach((r, i) => {
    s.addShape("ellipse", { x: M, y: y + 0.06, w: 0.52, h: 0.52, fill: { color: i === 2 ? ACCENTSOFT : TINT }, line: { type: "none" } });
    s.addText(String(i + 1), { x: M, y: y + 0.06, w: 0.52, h: 0.52, fontSize: 18, bold: true, color: i === 2 ? ACCENT : PRIMARY, align: "center", valign: "middle", fontFace: FONT, margin: 0 });
    s.addText([
      { text: r[0], options: { bold: true, fontSize: 15.5, color: TEXT, breakLine: true } },
      { text: r[1], options: { fontSize: 12.5, color: MUTED } },
    ], { x: M + 0.75, y: y - 0.05, w: 6.55, h: 1.5, fontFace: FONT, margin: 0, valign: "top", paraSpaceAfter: 5, lineSpacingMultiple: 1.12 });
    if (i < 2) s.addShape("line", { x: M + 0.75, y: y + 1.36, w: 6.4, h: 0, line: { color: BORDER, width: 0.75 } });
    y += 1.62;
  });

  // right stat panel (dark)
  card(s, 8.15, 1.82, 4.6, 4.98, DARK, true, 0.12);
  s.addText("VÌ SAO LÀ BÂY GIỜ?", { x: 8.5, y: 2.06, w: 3.9, h: 0.3, fontSize: 11.5, bold: true, charSpacing: 2.5, color: TEALSOFT, fontFace: FONT, margin: 0 });
  const stats = [
    ["16,1 triệu", "người Việt trên 60 tuổi (2025) — hơn 16% dân số"],
    ["94,4 triệu", "lượt khám BHYT 6 tháng đầu 2026 — hệ thống khám chữa bệnh chạy hết công suất"],
    ["3,5–4 bệnh", "mạn tính bình quân mỗi lần nhập viện của người cao tuổi"],
  ];
  let sy = 2.44;
  stats.forEach((st, i) => {
    s.addText(st[0], { x: 8.5, y: sy, w: 3.9, h: 0.6, fontSize: i === 0 ? 36 : 32, bold: true, color: i === 0 ? "F7965A" : "FFFFFF", fontFace: FONT, margin: 0, valign: "middle" });
    s.addText(st[1], { x: 8.5, y: sy + 0.62, w: 3.9, h: 0.58, fontSize: 11, color: "B9D8D0", fontFace: FONT, margin: 0, lineSpacingMultiple: 1.12 });
    sy += i === 0 ? 1.5 : 1.44;
  });
  src(s, [
    { t: "Nguồn: " },
    { t: "Tuổi Trẻ, 31/5/2025", url: "https://tuoitre.vn/nam-2025-viet-nam-co-16-dan-so-la-nguoi-cao-tuoi-1-nguoi-ganh-3-4-benh-man-tinh-20250531122410678.htm" },
    { t: "  ·  " },
    { t: "Tuổi Trẻ — già hóa nhanh nhất châu Á", url: "https://tuoitre.vn/nld/viet-nam-dang-co-toc-do-gia-hoa-nhanh-nhat-chau-a-196250531143638561.htm" },
    { t: "  ·  Lượt khám BHYT: nhóm tổng hợp, 2026" },
  ]);
  s.addNotes("Ba nhóm khách cùng một vấn đề: vận động, xương khớp, giấc ngủ suy giảm nhưng chưa được phục vụ bài bản. Con số 16,1 triệu người 60+ và 94,4 triệu lượt khám BHYT cho thấy nhu cầu tràn ra ngoài hệ thống.");
}

/* ============================== S3 · THỊ TRƯỜNG ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Cơ hội thị trường");
  title(s, "Chi tiêu sức khỏe tăng đều — Việt Nam là “ngôi sao tăng trưởng” wellness của châu Á");
  motif(s); pageNo(s, 3);

  // left: native bar chart
  card(s, M, 1.9, 6.5, 4.85, "FFFFFF", true);
  s.addText("Chi tiêu y tế Việt Nam (tỷ USD)", { x: M + 0.35, y: 2.12, w: 5.8, h: 0.35, fontSize: 14, bold: true, color: TEXT, fontFace: FONT, margin: 0 });
  s.addChart(pres.charts.BAR, [{
    name: "Chi tiêu y tế (tỷ USD)",
    labels: ["2019", "2025", "2028 (dự báo)"],
    values: [17.4, 27.5, 34.1],
  }], {
    x: M + 0.3, y: 2.55, w: 5.9, h: 3.85, barDir: "col", barGapWidthPct: 60,
    chartColors: [TEALSOFT, PRIMARY, ACCENT], varyColors: true,
    chartArea: { fill: { color: "FFFFFF" } },
    catAxisLabelColor: MUTED, catAxisLabelFontSize: 12, catAxisLabelFontFace: FONT,
    valAxisHidden: true, valAxisMaxVal: 40, valAxisMinVal: 0,
    valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TEXT,
    dataLabelFontSize: 14, dataLabelFontFace: FONT, dataLabelFormatCode: "0.0",
    showLegend: false, showTitle: false,
  });

  // right: 3 stat cards
  const rs = [
    ["6.300 tỷ USD", "quy mô kinh tế wellness toàn cầu (2023) — tăng nhanh hơn GDP thế giới", "https://globalwellnessinstitute.org/wp-content/uploads/2024/11/WellnessEconMonitor2024PDF.pdf", "Global Wellness Institute, 2024"],
    ["+15,6%/năm", "tăng trưởng wellness của Việt Nam — cao nhất châu Á, trước Ấn Độ (8,1%) và Trung Quốc (7,6%)", "https://globalwellnessinstitute.org/wp-content/uploads/2024/11/WellnessEconMonitor2024PDF.pdf", "GWI Wellness Economy Monitor"],
    ["1,4 tỷ USD", "doanh thu spa Việt Nam 2023 — top 20 thị trường toàn cầu, từ 850 triệu USD (2019)", "https://asia.nikkei.com/Business/Health-Care/Vietnam-fires-up-wellness-tourism-to-rival-Thailand-Indonesia", "Nikkei Asia, 1/2025"],
  ];
  let ry = 1.9;
  rs.forEach((r) => {
    card(s, 7.45, ry, 5.33, 1.5, TINT2, true);
    s.addText([
      { text: r[0] + "   ", options: { fontSize: 25, bold: true, color: ACCENT } },
      { text: r[1], options: { fontSize: 11.5, color: MUTED } },
    ], { x: 7.75, y: ry + 0.14, w: 4.8, h: 1.0, fontFace: FONT, margin: 0, valign: "top", lineSpacingMultiple: 1.1 });
    s.addText([{ text: "Nguồn: ", options: { italic: true } }, { text: r[3], options: { hyperlink: { url: r[2] } } }],
      { x: 7.75, y: ry + 1.13, w: 4.8, h: 0.28, fontSize: 10, color: PRIMARY, fontFace: FONT, margin: 0 });
    ry += 1.68;
  });
  src(s, [
    { t: "Nguồn: " },
    { t: "VnExpress, 24/12/2025", url: "https://vnexpress.net/chi-phi-cho-y-te-cua-nguoi-viet-nam-tang-thang-dung-4997605.html" },
    { t: " (chi tiêu y tế)  ·  GWI 2024  ·  Nikkei Asia" },
  ], 6.98);
  s.addNotes("Chi y tế tăng từ 17,4 lên 27,5 tỷ USD, dự báo 34,1 tỷ USD năm 2028. GWI xếp Việt Nam là thị trường wellness tăng trưởng nhanh nhất châu Á với 15,6%/năm.");
}

/* ============================== S4 · KHÁCH HÀNG ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Khách hàng mục tiêu");
  title(s, "Ba nhóm khách, một nhu cầu chung: chạm trị liệu định kỳ, thuận tiện và đáng tin");
  motif(s); pageNo(s, 4);

  const seg = [
    ["TRUNG NIÊN 40–60 TUỔI", "Nhóm chính", PRIMARY, TINT,
      ["Thu nhập ổn định nhất, bắt đầu đau nhức cơ – xương khớp", "Người quyết định dịch vụ cho cả cha mẹ già", "Chuộng gói định kỳ — tiếp cận qua Zalo, Facebook, giới thiệu"]],
    ["NHÂN VIÊN VĂN PHÒNG", "Nhóm phụ", TEALMID, TINT2,
      ["Đau cổ vai gáy nghề nghiệp, bận rộn", "Mua lẻ ban đầu, chuyển dần sang gói combo", "Cần khung giờ sớm – tối và phục vụ nhanh"]],
    ["NGƯỜI CAO TUỔI", "Nhóm phụ", ACCENT, ACCENTSOFT,
      ["Nhiều bệnh mạn tính, khó di chuyển", "Tin tưởng y học cổ truyền từ lâu", "Con cháu đặt và thanh toán hộ — trung thành cao nếu buổi đầu an toàn"]],
  ];
  let x = M;
  seg.forEach((g) => {
    card(s, x, 1.95, 3.93, 3.3, "FFFFFF", true);
    s.addShape("roundRect", { x: x + 0.28, y: 2.2, w: 1.18, h: 0.34, rectRadius: 0.17, fill: { color: g[3] }, line: { type: "none" } });
    s.addText(g[1], { x: x + 0.28, y: 2.2, w: 1.18, h: 0.34, fontSize: 10, bold: true, color: g[2], align: "center", valign: "middle", fontFace: FONT, margin: 0 });
    s.addText(g[0], { x: x + 0.28, y: 2.66, w: 3.4, h: 0.4, fontSize: 15, bold: true, color: TEXT, fontFace: FONT, margin: 0 });
    s.addText(g[4].map((t, i) => ({ text: t, options: { bullet: bu(), breakLine: i < g[4].length - 1 } })),
      { x: x + 0.28, y: 3.14, w: 3.45, h: 2.0, fontSize: 11.5, color: MUTED, fontFace: FONT, margin: 0, paraSpaceAfter: 7, valign: "top", lineSpacingMultiple: 1.12 });
    x += 4.13;
  });
  band(s, [
    { text: "Khả năng chi trả:  ", options: { bold: true, color: TEXT } },
    { text: "thu nhập bình quân đầu người 2025 đạt 5,9 triệu đồng/tháng (+9,3% so với 2024) — mức giá 120.000–320.000 đ/phiên nằm trong tầm chi trả định kỳ của nhóm mục tiêu.  ", options: { color: TEXT } },
    { text: "Nguồn: Tổng cục Thống kê (gso.gov.vn)", options: { color: PRIMARY, hyperlink: { url: "https://www.gso.gov.vn/" }, fontSize: 11 } },
  ], 5.85, 0.75);
  s.addNotes("Nhóm chính là trung niên 40–60: thu nhập tốt nhất, mua cho cả cha mẹ. Giá một phiên tương đương một bữa cơm gia đình — dễ quyết định định kỳ.");
}

/* ============================== S5 · GIẢI PHÁP & ĐỊNH VỊ ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Giải pháp CareTouch");
  title(s, "Trị liệu cổ truyền có kiểm chứng — giá phải chăng, phục vụ tận giường, tận nhà");
  motif(s); pageNo(s, 5);

  // left: positioning map
  card(s, M, 1.95, 6.3, 4.75, "FFFFFF", true);
  const px = M + 0.55, py = 2.25, pw = 5.3, ph = 3.6;
  s.addShape("line", { x: px, y: py, w: 0, h: ph, line: { color: BORDER, width: 1.5 } });
  s.addShape("line", { x: px, y: py + ph, w: pw, h: 0, line: { color: BORDER, width: 1.5 } });
  s.addText("Giá cao ↑", { x: px - 0.45, y: py - 0.28, w: 1.4, h: 0.26, fontSize: 10.5, color: MUTED, fontFace: FONT, margin: 0 });
  s.addText("Chuyên môn trị liệu →", { x: px + pw - 1.9, y: py + ph + 0.08, w: 1.9, h: 0.26, fontSize: 10.5, color: MUTED, fontFace: FONT, margin: 0, align: "right" });
  // competitors
  const dot = (dx, dy, w, h, color, label, sub, dark = false) => {
    s.addShape("ellipse", { x: dx - w / 2, y: dy - h / 2, w, h, fill: { color }, line: { type: "none" }, shadow: softShadow() });
    s.addText(label, { x: dx - 1.05, y: dy + h / 2 + 0.04, w: 2.1, h: 0.3, fontSize: 10.5, bold: true, color: dark ? TEXT : MUTED, align: "center", fontFace: FONT, margin: 0 });
  };
  dot(px + 1.15, py + 2.75, 1.1, 0.62, "C9D8D3", "Quán massage nhỏ");
  dot(px + 1.5, py + 0.62, 1.25, 0.68, TEALSOFT, "Spa cao cấp (thư giãn)");
  dot(px + 4.45, py + 2.62, 1.15, 0.6, "C9D8D3", "Tại nhà tự phát");
  // CareTouch star
  s.addShape("ellipse", { x: px + 3.35, y: py + 0.68, w: 1.7, h: 0.95, fill: { color: ACCENT }, line: { color: "FFFFFF", width: 2 }, shadow: shadow() });
  s.addText("CareTouch", { x: px + 3.35, y: py + 0.68, w: 1.7, h: 0.95, fontSize: 12.5, bold: true, color: "FFFFFF", align: "center", valign: "middle", fontFace: FONT, margin: 0 });
  s.addText("trị liệu cổ truyền · giá phải chăng", { x: px + 2.5, y: py + 1.74, w: 3.4, h: 0.3, fontSize: 10.5, bold: true, color: ACCENT, align: "center", fontFace: FONT, margin: 0 });

  // right: 4 pillars
  const pil = [
    ["Tận giường – tận nhà", "Mang theo đệm, dầu xoa, khăn sạch; khách cao tuổi không cần ai đưa đón."],
    ["Khảo sát thể trạng trước buổi trị liệu", "Mỗi buổi bắt đầu bằng hỏi bệnh nền, đo huyết áp — không an toàn thì không thực hiện."],
    ["Giá niêm yết minh bạch", "Bảng giá công khai, không phát sinh chi phí; combo tiết kiệm 15–20%."],
    ["Sổ sức khỏe cá nhân", "Ghi nhận phản ứng sau mỗi buổi — liệu trình liên tục, không rời rạc."],
  ];
  let pyy = 1.95;
  pil.forEach((p, i) => {
    s.addShape("roundRect", { x: 7.25, y: pyy, w: 0.5, h: 0.5, rectRadius: 0.12, fill: { color: i === 1 ? ACCENTSOFT : TINT }, line: { type: "none" } });
    s.addText(String(i + 1), { x: 7.25, y: pyy, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: i === 1 ? ACCENT : PRIMARY, align: "center", valign: "middle", fontFace: FONT, margin: 0 });
    s.addText([
      { text: p[0], options: { bold: true, fontSize: 14.5, color: TEXT, breakLine: true } },
      { text: p[1], options: { fontSize: 11.5, color: MUTED } },
    ], { x: 7.95, y: pyy - 0.04, w: 4.85, h: 1.2, fontFace: FONT, margin: 0, valign: "top", paraSpaceAfter: 3, lineSpacingMultiple: 1.1 });
    pyy += 1.22;
  });
  src(s, [{ t: "Định vị dựa trên khảo sát giá công khai 2025: massage tại nhà cho người lớn tuổi 350.000–620.000 đ/phiên (" },
    { t: "Massagenha, 7/2025", url: "https://massagenha.com/dich-vu/massage-nguoi-lon-tuoi/" },
    { t: "); spa tầm trung 350.000–400.000 đ." }]);
  s.addNotes("Bản đồ định vị: CareTouch chiếm góc 'chuyên môn cao - giá phải chăng' — nơi không ai đang phục vụ nhóm trung niên-cao tuổi một cách bài bản.");
}

/* ============================== S6 · DỊCH VỤ & GIÁ ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Sản phẩm dịch vụ");
  title(s, "Danh mục dịch vụ & giá đề xuất (VNĐ/phiên)");
  motif(s); pageNo(s, 6);

  const head = ["Dịch vụ", "Thời lượng", "Tại cơ sở", "Tại nhà"].map(t => ({
    text: t, options: { fill: { color: PRIMARY }, color: "FFFFFF", bold: true, fontSize: 13.5, align: t === "Dịch vụ" ? "left" : "center", valign: "middle" },
  }));
  const rowsData = [
    ["Gội đầu dưỡng sinh thảo dược", "45 phút", "120.000", "200.000"],
    ["Massage cổ – vai – gáy trị liệu", "60 phút", "180.000", "270.000"],
    ["Massage toàn thân thư giãn", "70 phút", "220.000", "320.000"],
    ["Bấm huyệt – đả thông kinh lạc", "60 phút", "200.000", "290.000"],
    ["Chườm thảo dược – xông hơi đông y", "60 phút", "180.000", "—"],
  ];
  const body = rowsData.map((r, i) => r.map((c, j) => ({
    text: c,
    options: {
      fill: { color: i % 2 ? TINT2 : "FFFFFF" }, fontSize: 13,
      color: j === 2 ? PRIMARY : (j === 3 ? ACCENT : TEXT), bold: j >= 2 && c !== "—",
      align: j === 0 ? "left" : "center", valign: "middle", italic: c === "—",
    },
  })));
  s.addTable([head, ...body], {
    x: M, y: 2.0, w: 8.1, colW: [3.6, 1.3, 1.6, 1.6], rowH: [0.5, 0.62, 0.62, 0.62, 0.62, 0.62],
    border: { pt: 0.75, color: BORDER }, fontFace: FONT, margin: 0.08,
  });

  // right column
  card(s, 9.0, 2.0, 3.78, 2.2, DARK, true, 0.1);
  s.addText("PHỤC VỤ TẬN NHÀ", { x: 9.3, y: 2.24, w: 3.2, h: 0.3, fontSize: 11, bold: true, charSpacing: 2, color: TEALSOFT, fontFace: FONT, margin: 0 });
  s.addText([
    { text: "Miễn phí trong bán kính 5 km", options: { bold: true, fontSize: 15, color: "FFFFFF", breakLine: true } },
    { text: "Xa hơn: phụ thu 20.000 đ mỗi 5 km; khách đặt tối thiểu 2 buổi mỗi lượt. Kỹ thuật viên mang theo đầy đủ đệm, dầu xoa, khăn sạch.", options: { fontSize: 11.5, color: "B9D8D0" } },
  ], { x: 9.3, y: 2.58, w: 3.2, h: 1.5, fontFace: FONT, margin: 0, paraSpaceAfter: 6, lineSpacingMultiple: 1.15 });
  card(s, 9.0, 4.45, 3.78, 2.1, ACCENTSOFT, true, 0.1);
  s.addText([
    { text: "Định vị giá", options: { bold: true, fontSize: 14, color: ACCENT, breakLine: true } },
    { text: "Thị trường tại nhà cho người lớn tuổi: 350.000–620.000 đ/phiên (TP.HCM). CareTouch ở khoảng 200.000–320.000 đ — chênh lệch 80–100 nghìn đ bù thời gian di chuyển, vẫn dưới phân khúc spa.", options: { fontSize: 11.5, color: TEXT } },
  ], { x: 9.3, y: 4.65, w: 3.2, h: 1.7, fontFace: FONT, margin: 0, paraSpaceAfter: 6, lineSpacingMultiple: 1.15 });

  src(s, [{ t: "Giá đề xuất của nhóm, căn theo khảo sát thị trường: " },
    { t: "Massagenha.com, 7/2025", url: "https://massagenha.com/dich-vu/massage-nguoi-lon-tuoi/" },
    { t: " — con số sẽ hiệu chỉnh theo địa bàn triển khai thực tế." }]);
  s.addNotes("Năm dịch vụ cốt lõi. Giá tại nhà cao hơn tại cơ sở 60–70 nghìn để bù chi phí đi lại. Chọn phân khúc giá ngay dưới spa, trên quán nhỏ.");
}

/* ============================== S7 · GÓI & DOANH THU LẶP LẠI ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Mô hình doanh thu");
  title(s, "Thiết kế doanh thu lặp lại — liệu trình 5–10 buổi, khách quay lại là mặc định");
  motif(s); pageNo(s, 7);

  const packs = [
    ["MUA LẺ", "120.000–290.000 đ", "", ["Không ràng buộc", "Giá niêm yết, không phát sinh", "Tích điểm thành viên từ phiên đầu"], "FFFFFF", TEXT, false, 21],
    ["COMBO 5 BUỔI", "890.000 đ", "tiết kiệm ~15%", ["Tự chọn kết hợp 3 dịch vụ", "Đặt lịch linh hoạt trong 3 tháng", "1 buổi tổng kết liệu trình"], "FFFFFF", TEXT, false, 34],
    ["CAREPLUS 10 BUỔI", "1.590.000 đ", "~20%", ["Giá cố định 12 tháng", "Tích điểm 5% · ưu đãi tháng sinh nhật", "Cho người thân dùng chung"], DARK, "FFFFFF", true, 31],
  ];
  let x = M;
  packs.forEach((p) => {
    const dark = p[6];
    card(s, x, 2.0, 3.93, 3.6, p[4], true, 0.1);
    s.addText(p[0], { x: x + 0.3, y: 2.28, w: 3.3, h: 0.32, fontSize: 12, bold: true, charSpacing: 2, color: dark ? TEALSOFT : TEALMID, fontFace: FONT, margin: 0 });
    const pSize = p[7] || 34;
    s.addText([
      { text: p[1], options: { fontSize: pSize, bold: true, color: dark ? "F7965A" : (p[0] === "COMBO 5 BUỔI" ? PRIMARY : TEXT) } },
      { text: p[2] ? "  " + p[2] : "", options: { fontSize: 12, color: dark ? "B9D8D0" : MUTED } },
    ], { x: x + 0.3, y: 2.62, w: 3.45, h: 0.75, fontFace: FONT, margin: 0, valign: "middle" });
    s.addText(p[3].map((t, i) => ({ text: t, options: { bullet: bu(), breakLine: i < p[3].length - 1 } })),
      { x: x + 0.3, y: 3.55, w: 3.4, h: 1.9, fontSize: 12, color: dark ? "D5EAE4" : MUTED, fontFace: FONT, margin: 0, paraSpaceAfter: 8, lineSpacingMultiple: 1.12 });
    x += 4.13;
  });
  band(s, [
    { text: "60% ", options: { bold: true, fontSize: 16, color: ACCENT } },
    { text: "doanh thu tháng thứ 12 đến từ gói combo và thẻ thành viên — dòng tiền dự đoán được, chi phí giữ chân khách thấp. ", options: { color: TEXT } },
    { text: "(dự phóng minh họa)", options: { italic: true, color: MUTED, fontSize: 11.5 } },
  ], 6.05, 0.72);
  s.addNotes("Liệu trình YHCT vốn cần 5-10 buổi — mô hình gói biến nhu cầu y khoa thành doanh thu định kỳ. CarePlus 10 là gói chủ lực.");
}

/* ============================== S8 · QUY TRÌNH ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Quy trình vận hành");
  title(s, "Một buổi trị liệu chuẩn — an toàn là tính năng của dịch vụ, không phải khẩu hiệu");
  motif(s); pageNo(s, 8);

  const steps = [
    ["01", "Khảo sát", "5 phút đầu: hỏi bệnh nền, đo huyết áp, kiểm tra chống chỉ định."],
    ["02", "Đánh giá", "Xác định vùng đau, tư thế sai lệch để trị liệu nhắm đúng chỗ."],
    ["03", "Trị liệu", "Matxa – bấm huyệt – chườm thảo dược theo bài chuẩn 45–70 phút."],
    ["04", "Theo dõi", "Dặn bài tập nhẹ, ghi sổ sức khỏe — buổi sau điều chỉnh tốt hơn."],
  ];
  let x = M;
  steps.forEach((st, i) => {
    card(s, x, 2.15, 2.85, 3.15, i === 0 ? TINT : "FFFFFF", true, 0.1);
    s.addText(st[0], { x: x + 0.25, y: 2.35, w: 1.6, h: 0.85, fontSize: 46, bold: true, color: i === 0 ? "BFE0D6" : TEALSOFT, fontFace: FONT, margin: 0 });
    s.addText(st[1], { x: x + 0.25, y: 3.3, w: 2.35, h: 0.4, fontSize: 17, bold: true, color: TEXT, fontFace: FONT, margin: 0 });
    s.addText(st[2], { x: x + 0.25, y: 3.75, w: 2.4, h: 1.4, fontSize: 11.5, color: MUTED, fontFace: FONT, margin: 0, lineSpacingMultiple: 1.2 });
    if (i < 3) s.addShape("rightArrow", { x: x + 2.87, y: 3.55, w: 0.36, h: 0.3, fill: { color: ACCENT }, line: { type: "none" } });
    x += 3.08;
  });
  band(s, [
    { text: "Nguyên tắc bất di bất dịch:  ", options: { bold: true } },
    { text: "không an toàn → không thực hiện. Khách có chống chỉ định sẽ được từ chối hoặc khuyên đi khám; cơ sở mua bảo hiểm trách nhiệm dịch vụ.", options: {} },
  ], 6.0, 0.72, ACCENTSOFT);
  s.addNotes("Điểm khác biệt vận hành quan trọng nhất: khảo sát thể trạng bắt buộc trước mỗi buổi. Đây là thứ spa thư giãn và quán nhỏ không làm.");
}

/* ============================== S9 · KHÁC BIỆT CẠNH TRANH ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Giá trị khác biệt");
  title(s, "Khoảng trống thị trường: dịch vụ chăm sóc chuyên nghiệp cho trung niên – cao tuổi còn bỏ ngỏ");
  motif(s); pageNo(s, 9);

  const C = "✓", X = "✕", P = "~";
  const head = ["Tiêu chí", "CareTouch", "Spa cao cấp", "Quán massage nhỏ", "Tại nhà tự phát"].map((t, j) => ({
    text: t, options: {
      fill: { color: j === 1 ? ACCENT : PRIMARY }, color: "FFFFFF", bold: true, fontSize: 12.5,
      align: j === 0 ? "left" : "center", valign: "middle",
    },
  }));
  const rows = [
    ["Giá niêm yết minh bạch, không phát sinh", C, C, X, P],
    ["Trị liệu cổ truyền theo bài chuẩn", C, X, P, P],
    ["Khảo sát thể trạng an toàn trước mỗi buổi", C, X, X, P],
    ["Phục vụ tận giường – tận nhà", C, X, P, C],
    ["Sổ sức khỏe cá nhân theo dõi liệu trình", C, X, X, X],
    ["Giá phù hợp chi trả định kỳ (120–320 nghìn đồng)", C, X, C, P],
  ];
  const body = rows.map((r, i) => r.map((c, j) => ({
    text: c, options: {
      fill: { color: j === 1 ? ACCENTSOFT : (i % 2 ? TINT2 : "FFFFFF") },
      color: j === 1 ? (c === C ? ACCENT : MUTED) : (c === C ? PRIMARY : (c === X ? "B9C6C2" : MUTED)),
      bold: j === 1 && c === C, fontSize: j === 0 ? 12 : 15, align: j === 0 ? "left" : "center", valign: "middle",
    },
  })));
  s.addTable([head, ...body], {
    x: M, y: 1.98, w: 12.23, colW: [4.63, 1.9, 1.9, 1.9, 1.9], rowH: [0.55, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7],
    border: { pt: 0.75, color: BORDER }, fontFace: FONT, margin: 0.07,
  });
  src(s, [{ t: "So sánh do nhóm thực hiện trên cơ sở khảo sát dịch vụ công khai tại TP.HCM, 2025–2026 (spa: KKday, Jackfruit Adventure; home care: Massagenha)." }]);
  s.addNotes("Bảng so sánh: chỉ CareTouch đạt đồng thời cả 6 tiêu chí. Spa cao cấp bỏ lại nhóm trung niên vì giá; quán nhỏ thiếu an toàn; home care rời rạc thiếu quy trình.");
}

/* ============================== S10 · TÀI CHÍNH ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Tài chính dự phóng");
  title(s, "Hòa vốn vận hành từ tháng thứ 5 — hoàn vốn đầy đủ trong khoảng 12 tháng");
  motif(s); pageNo(s, 10);

  card(s, M, 1.95, 7.35, 4.75, "FFFFFF", true);
  s.addText("Doanh thu dự phóng 12 tháng đầu (triệu đồng/tháng)", { x: M + 0.32, y: 2.14, w: 6.6, h: 0.32, fontSize: 13.5, bold: true, color: TEXT, fontFace: FONT, margin: 0 });
  const months = ["T1","T2","T3","T4","T5","T6","T7","T8","T9","T10","T11","T12"];
  const rev = [22, 30, 38, 44, 50, 56, 62, 68, 73, 78, 83, 88];
  const cols = ["9ED9CC","9ED9CC","9ED9CC","9ED9CC","E8590C","0E7C6B","0E7C6B","0E7C6B","0E7C6B","0E7C6B","0E7C6B","0E7C6B"];
  s.addChart([
    {
      type: pres.charts.BAR,
      data: [{ name: "Doanh thu", labels: months, values: rev }],
      options: { chartColors: cols, varyColors: true, barDir: "col", barGapWidthPct: 45 },
    },
    {
      type: pres.charts.LINE,
      data: [{ name: "Chi phí cố định 48 tr./tháng", labels: months, values: months.map(() => 48) }],
      options: { chartColors: [TEXT], lineSize: 1.75, lineDash: "dash", lineDataSymbol: "none" },
    },
  ], {
    x: M + 0.25, y: 2.52, w: 6.85, h: 3.6,
    chartArea: { fill: { color: "FFFFFF" } },
    catAxisLabelColor: MUTED, catAxisLabelFontSize: 10.5, catAxisLabelFontFace: FONT,
    valAxisLabelColor: MUTED, valAxisLabelFontSize: 10.5, valAxisLabelFontFace: FONT,
    valAxisMaxVal: 100, valAxisMinVal: 0,
    valGridLine: { color: "EAF3F0", size: 0.5 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "b", legendFontSize: 10.5, legendFontFace: FONT, legendColor: MUTED,
    showValue: false,
  });

  const fin = [
    ["165 tr.", "tổng vốn khởi đầu = 95 tr. đầu tư ban đầu + 70 tr. vốn lưu động bù lỗ 4 tháng đầu", TINT, PRIMARY],
    ["48 tr./th.", "chi phí cố định: thuê 12tr + lương 4 kỹ thuật viên 32tr + điện nước, vật tư, khấu hao", TINT, PRIMARY],
    ["88 tr./th.", "doanh thu tháng 12 — tương đương ~350 khách/tháng", TINT, PRIMARY],
    ["≥ 55%", "tỷ lệ khách quay lại để hoàn vốn đúng hạn — đo từ tháng đầu tiên", ACCENTSOFT, ACCENT],
  ];
  let fy = 1.95;
  fin.forEach((f) => {
    card(s, 8.3, fy, 4.48, 1.08, f[2], true, 0.08);
    s.addText(f[0], { x: 8.58, y: fy + 0.12, w: 1.62, h: 0.85, fontSize: f[0] === "165 tr." ? 19 : 21, bold: true, color: f[3], fontFace: FONT, margin: 0, valign: "middle" });
    s.addText(f[1], { x: 10.24, y: fy + 0.1, w: 2.4, h: 0.9, fontSize: 10.5, color: MUTED, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.08 });
    fy += 1.24;
  });
  s.addText([
    { text: "T4: 44 < 48 < T5: 50  →  ", options: { color: MUTED } },
    { text: "hòa vốn ngay tháng thứ 5", options: { color: ACCENT, bold: true } },
  ], { x: M + 0.32, y: 6.2, w: 6.6, h: 0.3, fontSize: 10.5, italic: true, fontFace: FONT, margin: 0 });
  src(s, [
    { t: "Phép tính: lỗ lũy kế T1–T4 = 26 + 18 + 10 + 4 = 58 triệu → dự trù vốn lưu động 70 triệu; lãi ròng lũy kế T5–T12 = 174 triệu ≥ 165 triệu → hoàn vốn đủ trong ~12 tháng." },
  ], 6.74);
  src(s, [
    { t: "Giả định minh họa của nhóm: 1 cơ sở 2 giường trị liệu; 4 kỹ thuật viên (lương cứng ~8 tr. + phụ cấp ca tại nhà), nhóm sáng lập trực tiếp điều phối; giá theo bảng giá đề xuất (slide 6)." },
  ]);
  s.addNotes("Tổng vốn khởi đầu 165 triệu = 95 triệu đầu tư + 70 triệu vốn lưu động bù lỗ 4 tháng đầu (26+18+10+4=58 triệu, dự trù dư 12 triệu). Doanh thu vượt chi phí cố định 48 triệu từ tháng 5; lãi ròng lũy kế T5-T12 = 174 triệu ≥ 165 triệu nên hoàn vốn đủ trong khoảng 12 tháng. Kịch bản thận trọng (quay lại 55%): 12-14 tháng. Nhân sự: 4 KTV + nhóm sáng lập điều phối.");
}

/* ============================== S11 · LỘ TRÌNH ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Lộ trình triển khai");
  title(s, "Từ một cơ sở kiểm chứng đến chuỗi CareTouch");
  motif(s); pageNo(s, 11);

  const ph = [
    ["GĐ 1 · Tháng 1–6", "Kiểm chứng", "Vận hành cơ sở đầu tiên; chuẩn hóa quy trình sàng lọc & đào tạo; đo tỷ lệ khách quay lại từ tháng đầu."],
    ["GĐ 2 · Tháng 7–12", "Tăng trưởng", "Đạt 88 tr./th. doanh thu; ≥55% khách quay lại; hoàn tất hồ sơ pháp lý & chứng chỉ nghề."],
    ["GĐ 3 · Năm 2", "Mở rộng", "Cơ sở thứ 2; Zalo Mini App đặt lịch & nhắc liệu trình; huy động vốn mở rộng."],
    ["GĐ 4 · Năm 3", "Chuỗi hóa", "Nhượng quyền 3–5 cơ sở vệ tinh; chuẩn CareTouch Academy đào tạo kỹ thuật viên."],
  ];
  // timeline line
  s.addShape("line", { x: M + 0.4, y: 2.62, w: 11.2, h: 0, line: { color: BORDER, width: 2 } });
  let x = M;
  ph.forEach((p, i) => {
    s.addShape("ellipse", { x: x + 0.4, y: 2.34, w: 0.56, h: 0.56, fill: { color: i === 3 ? ACCENT : PRIMARY }, line: { color: "FFFFFF", width: 2.5 }, shadow: softShadow() });
    s.addText(String(i + 1), { x: x + 0.4, y: 2.34, w: 0.56, h: 0.56, fontSize: 16, bold: true, color: "FFFFFF", align: "center", valign: "middle", fontFace: FONT, margin: 0 });
    s.addText(p[0], { x: x, y: 3.1, w: 2.85, h: 0.3, fontSize: 11, bold: true, color: TEALMID, fontFace: FONT, margin: 0 });
    s.addText(p[1], { x: x, y: 3.4, w: 2.85, h: 0.42, fontSize: 18, bold: true, color: TEXT, fontFace: FONT, margin: 0 });
    s.addText(p[2], { x: x, y: 3.9, w: 2.8, h: 1.7, fontSize: 11.5, color: MUTED, fontFace: FONT, margin: 0, lineSpacingMultiple: 1.18 });
    x += 3.08;
  });
  band(s, [
    { text: "Nguyên tắc mở rộng:  ", options: { bold: true } },
    { text: "chỉ mở cơ sở mới khi cơ sở hiện tại đạt ≥55% khách quay lại và quy trình an toàn đã được kiểm chứng — tăng trưởng đi theo chất lượng, không đi trước.", options: {} },
  ], 6.1, 0.72);
  s.addNotes("Lộ trình 4 giai đoạn, mỗi giai đoạn có điều kiện vượt qua rõ ràng. Năm 3 là năm nhượng quyền — bản chất franchising của mô hình dịch vụ chuẩn hóa.");
}

/* ============================== S12 · RỦI RO ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Quản trị rủi ro");
  title(s, "Nhìn thẳng rủi ro — và giải pháp ứng phó cho từng cái");
  motif(s); pageNo(s, 12);

  const risks = [
    ["Ranh giới pháp lý với “khám chữa bệnh”", "CAO",
      "Định vị chăm sóc – thư giãn – hỗ trợ vận động; không khám bệnh, không kê thuốc, không thực hiện kỹ thuật thuộc phạm vi khám chữa bệnh (Luật Khám bệnh, chữa bệnh 2023; Nghị định 96/2023/NĐ-CP); bác sĩ y học cổ truyền cộng tác; kỹ thuật viên thi chứng chỉ nghề."],
    ["Sự cố với khách có bệnh nền", "CAO",
      "Khảo sát bắt buộc buổi đầu (hỏi bệnh nền, đo huyết áp); từ chối hoặc khuyên đi khám khi có chống chỉ định; mua bảo hiểm trách nhiệm dịch vụ."],
    ["Nhân lực và tay nghề", "TRUNG BÌNH",
      "Đào tạo nội bộ theo giáo trình chuẩn, kèm cặp thực tế; chất lượng đo bằng đánh giá sau mỗi buổi; người dạy nghề giữ cổ phần nhỏ."],
    ["Cạnh tranh về giá", "TRUNG BÌNH",
      "Không đua giá rẻ nhất; cạnh tranh bằng quy trình an toàn và niềm tin; gói combo & CarePlus giữ khách dài hạn."],
  ];
  let y = 2.0;
  risks.forEach((r, i) => {
    const hi = r[1] === "CAO";
    card(s, M, y, 12.23, 1.06, i % 2 ? TINT2 : "FFFFFF", false, 0.07);
    s.addShape("roundRect", { x: M + 0.22, y: y + 0.33, w: hi ? 0.85 : 1.55, h: 0.4, rectRadius: 0.2, fill: { color: hi ? ACCENT : TEALMID }, line: { type: "none" } });
    s.addText(r[1], { x: M + 0.22, y: y + 0.33, w: hi ? 0.85 : 1.55, h: 0.4, fontSize: 10.5, bold: true, color: "FFFFFF", align: "center", valign: "middle", fontFace: FONT, margin: 0 });
    s.addText(r[0], { x: M + 2.0, y: y + 0.12, w: 3.6, h: 0.85, fontSize: 13.5, bold: true, color: TEXT, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.05 });
    s.addText([
      { text: "Giải pháp:  ", options: { bold: true, color: PRIMARY } },
      { text: r[2], options: { color: MUTED } },
    ], { x: M + 5.85, y: y + 0.08, w: 6.15, h: 0.94, fontSize: 10.8, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.1 });
    y += 1.18;
  });
  src(s, [{ t: "Đánh giá rủi ro của nhóm dựa trên quy định hiện hành về massage – xoa bóp (Luật Khám bệnh, chữa bệnh 2023; Nghị định 96/2023/NĐ-CP) và khảo sát thực địa." }]);
  s.addNotes("Hai rủi ro mức cao đều thuộc pháp lý - an toàn. Chiến lược: định vị 'chăm sóc - thư giãn - hỗ trợ vận động', không xâm phạm phạm vi khám chữa bệnh theo Luật KBBC 2023 và NĐ 96/2023/NĐ-CP. Nhân sự vận hành: 4 KTV lương cứng + nhóm sáng lập trực tiếp điều phối giai đoạn đầu; năm 2 tuyển thêm điều phối viên toàn thời gian.");
}

/* ============================== S13 · ĐỘI NGŨ ============================== */
{
  const s = pres.addSlide();
  s.background = { color: "FFFFFF" };
  kicker(s, "Nhóm sáng lập");
  title(s, "Nhóm 3 — Tổ 14 — Lớp A1K79");
  motif(s); pageNo(s, 13);

  const team = [
    ["Trịnh Hương Thảo", "2401634"],
    ["Bạch Thùy Trâm", "2401679"],
    ["Bùi Thị Anh Trâm", "2401681"],
    ["Phạm Quốc Việt", "2401742"],
    ["Viengkeo Shengchan", "2401761"],
  ];
  let x = M;
  team.forEach((t, i) => {
    card(s, x, 2.5, 2.28, 2.75, i === 0 ? TINT : "FFFFFF", true, 0.1);
    const init = t[0].split(" ").slice(-2).map(w => w[0]).join("");
    s.addShape("ellipse", { x: x + 0.74, y: 2.8, w: 0.8, h: 0.8, fill: { color: i % 2 ? ACCENT : PRIMARY }, line: { type: "none" } });
    s.addText(init, { x: x + 0.74, y: 2.8, w: 0.8, h: 0.8, fontSize: 17, bold: true, color: "FFFFFF", align: "center", valign: "middle", fontFace: FONT, margin: 0 });
    s.addText(t[0], { x: x + 0.12, y: 3.78, w: 2.04, h: 0.65, fontSize: 12.5, bold: true, color: TEXT, align: "center", fontFace: FONT, margin: 0, valign: "top", lineSpacingMultiple: 1.05 });
    s.addText(t[1], { x: x + 0.12, y: 4.48, w: 2.04, h: 0.3, fontSize: 10.5, color: MUTED, align: "center", fontFace: FONT, margin: 0 });
    x += 2.49;
  });
  band(s, [
    { text: "Cam kết vận hành:  ", options: { bold: true } },
    { text: "kỹ thuật viên thi chứng chỉ nghề · bác sĩ y học cổ truyền cộng tác · bảo hiểm trách nhiệm dịch vụ · quy trình sàng lọc an toàn cho mọi buổi trị liệu.", options: {} },
  ], 5.95, 0.75, TINT);
  s.addNotes("Dự án được phát triển trong môn Thực tập NCKH Đổi mới sáng tạo Khởi nghiệp — Buổi thực tập: ca chiều 1, thứ 7, ngày 26/9/2026.");
}

/* ============================== S14 · CLOSING ============================== */
{
  const s = pres.addSlide();
  s.background = { path: A("bg_close.png") };
  s.addText("Cùng CareTouch chạm đến\nsức khỏe Việt Nam.", {
    x: 0.9, y: 1.7, w: 8.6, h: 1.9, fontSize: 44, bold: true, color: "FFFFFF", fontFace: FONT, margin: 0, lineSpacingMultiple: 1.08,
  });
  s.addText("Healing Touch, Healthy Life.", {
    x: 0.9, y: 3.7, w: 8, h: 0.55, fontSize: 22, italic: true, color: "F7965A", fontFace: "Georgia", margin: 0,
  });
  // ask card
  s.addShape("roundRect", { x: 0.9, y: 4.5, w: 7.6, h: 1.5, rectRadius: 0.12, fill: { color: "0E4A41", transparency: 18 }, line: { color: TEALMID, width: 1 } });
  s.addText([
    { text: "Vốn gọi đề xuất (minh họa): 500 triệu đồng", options: { bold: true, fontSize: 17, color: "FFFFFF", breakLine: true } },
    { text: "mở rộng lên 2 cơ sở + Zalo Mini App đặt lịch — mục tiêu doanh thu năm 2 gấp đôi năm 1.", options: { fontSize: 12.5, color: "CFE6DF" } },
  ], { x: 1.2, y: 4.68, w: 7.0, h: 1.2, fontFace: FONT, margin: 0, paraSpaceAfter: 5 });
  s.addText([
    { text: "Nhóm 3 · Tổ 14 · Lớp A1K79", options: { bold: true, color: "FFFFFF", breakLine: true } },
    { text: "Trịnh Hương Thảo · Bạch Thùy Trâm · Bùi Thị Anh Trâm · Phạm Quốc Việt · Viengkeo Shengchan", options: { color: "9EC9BF" } },
  ], { x: 0.9, y: 6.5, w: 11.5, h: 0.8, fontSize: 12, fontFace: FONT, margin: 0, paraSpaceAfter: 4 });
  s.addNotes("Kết: khoản vốn 500 triệu (minh họa) rót vào 2 cơ sở mới + nền tảng đặt lịch. Cảm ơn thầy/cô và nhà đầu tư.");
}

pres.writeFile({ fileName: path.join(__dirname, "CareTouch-deck.pptx") }).then(() => console.log("PPTX written OK"));
