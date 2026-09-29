/* CareTouch — lõi sửa nội dung dùng chung cho index.html và admin.html
   units(): xác định khối chữ được phép sửa, gán data-ct-eid theo thứ tự
   applyTo(): nạp nội dung đã lưu (content.json) lên trang
*/
(function (global) {
  'use strict';

  /* Chỉ các khối chữ này được sửa. Nếu một khối nằm trong khối khác đã chọn
     thì bỏ khối con (chỉ sửa từ ngoài vào) → giữ nguyên định dạng khi nạp lại. */
  var SEL = ['h1', 'h2', 'h3', 'p', 'li', 'summary', '.ans', 'select', 'label',
             '.stat', '.chip', '.price', '.who', '.tag', '.wv-quote', '.wv-author',
             '.hero-badge', '.contact-line', '.foot-bottom span', 'footer h4'].join(',');

  var FILE = 'content.json';

  /* id của các phần tử do JS quản lý — không cho sửa để không bị ghi đè.
     Các id khác (ví dụ #fSvc) vẫn sửa bình thường. */
  var PROTECTED = { year: 1, bookForm: 1, toast: 1 };

  function units(root) {
    var out = [], chosen = [];
    root.querySelectorAll(SEL).forEach(function (el) {
      if (el.id && PROTECTED[el.id]) return;
      for (var p = el.parentElement; p && p !== root; p = p.parentElement) {
        if (chosen.indexOf(p) !== -1) return;  // đã nằm trong khối cha
      }
      el.setAttribute('data-ct-eid', 'e' + (out.length + 1));
      chosen.push(el);
      out.push(el);
    });
    return out;
  }

  /* Làm sạch HTML người dùng nhập: bỏ script/khung, thuộc tính sự kiện,
     link javascript: và các thuộc tính kỹ thuật của trình soạn thảo. */
  function clean(html) {
    var box = document.createElement('div');
    box.innerHTML = html;
    box.querySelectorAll('script,style,iframe,object,embed,link,meta').forEach(function (n) { n.remove(); });
    [].slice.call(box.querySelectorAll('*')).forEach(function (n) {
      [].slice.call(n.attributes).forEach(function (a) {
        var nm = a.name.toLowerCase();
        if (nm.indexOf('on') === 0) n.removeAttribute(a.name);
        else if (nm === 'href' && /^\s*javascript:/i.test(a.value)) n.removeAttribute(a.name);
        else if (nm === 'data-ct-eid' || nm === 'contenteditable' || nm === 'spellcheck') n.removeAttribute(a.name);
      });
    });
    return box.innerHTML;
  }

  /* Đặt lại nội dung đã lưu. Bỏ data-count để những con số người dùng sửa
     không bị script đếm số ghi đè khi trang vừa mở. */
  function applyTo(root, data) {
    if (!data) return 0;
    var n = 0;
    units(root).forEach(function (el) {
      var rec = data[el.getAttribute('data-ct-eid')];
      if (!rec || typeof rec.h !== 'string') return;
      el.innerHTML = clean(rec.h);
      el.querySelectorAll('[data-count]').forEach(function (s) {
        s.removeAttribute('data-count');
        s.removeAttribute('data-dec');
      });
      n++;
    });
    var y = root.getElementById ? root.getElementById('year') : root.querySelector('#year');
    if (y) y.textContent = new Date().getFullYear();
    return n;
  }

  function load(cb) {
    fetch(FILE + '?t=' + Date.now())
      .then(function (r) { return r.ok ? r.json() : {}; })
      .then(function (d) { cb(d && typeof d === 'object' ? d : {}); })
      .catch(function () { cb({}); });
  }

  global.CareTouchEdit = { units: units, applyTo: applyTo, clean: clean, load: load, FILE: FILE };
})(window);
