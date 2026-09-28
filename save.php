<?php
// CareTouch — lưu chỉnh sửa nội dung slide (đặt cùng thư mục với CareTouch-slides.html)
// Mã bảo mật mặc định: 1234321 — hãy đổi thành mã khác trước khi đưa lên host thật.
header('Content-Type: application/json; charset=utf-8');

$CODE = '1234321';
$FILE = __DIR__ . '/content.json';

$in = json_decode(file_get_contents('php://input'), true);
if (!$in || !isset($in['code']) || $in['code'] !== $CODE) {
    http_response_code(403);
    echo json_encode(['ok' => false, 'error' => 'Sai mã bảo mật']);
    exit;
}
$edits = isset($in['edits']) && is_array($in['edits']) ? $in['edits'] : null;
if ($edits === null) {
    http_response_code(400);
    echo json_encode(['ok' => false, 'error' => 'Dữ liệu không hợp lệ']);
    exit;
}

// làm sạch tối thiểu: loại script/iframe và thuộc tính on*
$clean = [];
foreach ($edits as $k => $v) {
    if (!is_string($k) || !preg_match('/^s\d+-\d+$/', $k) || !is_string($v)) continue;
    $v = preg_replace('#<script[^>]*>.*?</script>#is', '', $v);
    $v = preg_replace('#<iframe[^>]*>.*?</iframe>#is', '', $v);
    $v = preg_replace('#\son\w+\s*=\s*("[^"]*"|\'[^\']*\')#i', '', $v);
    $v = trim($v);
    if ($v !== '') $clean[$k] = $v;
}

$cur = is_file($FILE) ? json_decode(file_get_contents($FILE), true) : [];
if (!is_array($cur)) $cur = [];
$cur = array_merge($cur, $clean);

$ok = file_put_contents($FILE, json_encode($cur, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT), LOCK_EX);
if ($ok === false) {
    http_response_code(500);
    echo json_encode(['ok' => false, 'error' => 'Không ghi được content.json — kiểm tra quyền ghi (permission 644/664) của thư mục trên cPanel']);
    exit;
}
echo json_encode(['ok' => true, 'saved' => count($clean)]);
