/**
 * Nhận dữ liệu từ trang đăng ký và ghi vào Google Sheet.
 * Cách dùng: mở Google Sheet "Đăng ký khóa AI Ứng Dụng Cho Khối Văn Phòng"
 * → Tiện ích mở rộng → Apps Script → xóa hết code mẫu → dán toàn bộ file này → Lưu
 * → Triển khai → Tùy chọn triển khai mới → Loại: Ứng dụng web
 *   - Thực thi dưới dạng: Tôi
 *   - Người có quyền truy cập: Bất kỳ ai
 * → Triển khai → Cấp quyền → copy "URL ứng dụng web" (kết thúc bằng /exec).
 */

var HEADERS = ['Thời gian đăng ký', 'Họ và tên', 'SĐT / Zalo', 'Email', 'Đơn vị công tác',
  'Chức vụ hiện tại', 'Mức độ dùng AI hiện tại', 'Mong muốn sau khóa học',
  'Mong muốn khác / ghi chú', 'Biết đến khóa học qua'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var p = (e && e.parameter) || {};
    if (p.website) return json_({ ok: true });            // bẫy chống spam
    if (!p.hoTen || !p.email || !p.sdt) return json_({ ok: false, error: 'Thiếu thông tin bắt buộc' });

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);

    sh.appendRow([
      new Date(),
      clean_(p.hoTen, 80),
      "'" + clean_(p.sdt, 20).replace(/^'/, ''),      // giữ số 0 ở đầu
      clean_(p.email, 100),
      clean_(p.donVi, 120),
      clean_(p.chucVu, 80),
      clean_(p.mucDoAI, 60),
      clean_(p.mongMuon, 600),
      clean_(p.ghiChu, 800),
      clean_(p.nguon, 60)
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, msg: 'Form đăng ký đang hoạt động' });
}

// Chạy 1 lần (tùy chọn) để định dạng hàng tiêu đề
function dinhDangTieuDe() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#0B1028').setFontColor('#FFFFFF').setWrap(true);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  sh.setColumnWidths(1, HEADERS.length, 180);
}

function clean_(v, max) {
  v = String(v || '').trim().slice(0, max);
  if (/^[=+\-@]/.test(v)) v = "'" + v;                     // chống chèn công thức
  return v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
