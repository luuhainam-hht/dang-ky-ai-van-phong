// Nhận dữ liệu từ trang đăng ký và ghi vào Google Sheet
// Sheet "Cài đặt": B1 = số chỗ tối đa, B2 = số người đã đăng ký qua kênh khác (chỉ ghi số thật)

var SHEET_ID = '1hV8YCzfM245_lk5fsyCS2NX3sgrHbyCKzlA82sL0UlQ';

var HEADERS = ['Thời gian đăng ký', 'Họ và tên', 'SĐT / Zalo', 'Email', 'Đơn vị công tác',
  'Chức vụ hiện tại', 'Mức độ dùng AI hiện tại', 'Mong muốn sau khóa học',
  'Mong muốn khác / ghi chú', 'Biết đến khóa học qua', 'Học phí áp dụng'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var p = (e && e.parameter) || {};
    if (p.website) return json_({ ok: true });            // bẫy chống spam
    if (!p.hoTen || !p.email || !p.sdt) return json_({ ok: false, error: 'Thiếu thông tin bắt buộc' });

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var s = seats_(ss);
    if (s.taken >= s.max) return json_({ ok: false, full: true, error: 'Lớp đã đủ chỗ' });

    var sh = ss.getSheets()[0];
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
      clean_(p.nguon, 60),
      clean_(p.hocPhi, 20)
    ]);
    return json_({ ok: true, taken: s.taken + 1, max: s.max });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  var s = seats_(SpreadsheetApp.openById(SHEET_ID));
  return json_({ ok: true, max: s.max, taken: s.taken, left: s.max - s.taken });
}

// Số chỗ đã đăng ký = số dòng đăng ký trên Sheet + số ở ô B2 sheet "Cài đặt"
function seats_(ss) {
  var st = settings_(ss);
  var max = Number(st.getRange('B1').getValue()) || 30;
  var extra = Number(st.getRange('B2').getValue()) || 0;
  var n = Math.max(0, ss.getSheets()[0].getLastRow() - 1);
  return { max: max, taken: Math.min(max, n + extra) };
}

function settings_(ss) {
  var st = ss.getSheetByName('Cài đặt');
  if (!st) {
    st = ss.insertSheet('Cài đặt', ss.getNumSheets());
    st.getRange('A1:B3').setValues([
      ['Số chỗ tối đa', 30],
      ['Đã đăng ký qua kênh khác (Zalo, điện thoại…) – chỉ ghi số thật', 0],
      ['Ghi chú', 'Web hiển thị: (số dòng đăng ký ở sheet đầu tiên + ô B2) / ô B1']
    ]);
    st.setColumnWidth(1, 460);
    st.getRange('A1:A3').setFontWeight('bold');
  }
  return st;
}

// Chạy 1 lần: múi giờ Việt Nam + tiêu đề + sheet Cài đặt
function caiDatSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  ss.setSpreadsheetTimeZone('Asia/Ho_Chi_Minh');
  dinhDangTieuDe();
  settings_(ss);
  Logger.log(JSON.stringify(seats_(ss)));
}

function dinhDangTieuDe() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
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
