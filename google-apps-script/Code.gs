/**
 * XÓM TRỌ VĂN HÓA – Cầu nối 2 chiều giữa webapp (Netlify) và Google Sheet, có đăng nhập.
 * Dùng cùng file Excel mẫu google-sheet/xom-tro-van-hoa.xlsx (đã có sẵn 5 sheet, đúng cột).
 * Chưa cài Apps Script: webapp vẫn đăng nhập + xem dữ liệu từ link "Xuất bản lên web" (chỉ đọc).
 *
 * CÀI ĐẶT (làm 1 lần):
 *  1. Mở Google Sheet → Tiện ích mở rộng → Apps Script. Xóa mã mặc định, dán toàn bộ tệp này, bấm Lưu.
 *  2. Chọn hàm "caiDat" ở thanh công cụ → bấm Chạy → cấp quyền.
 *     Hàm tạo sheet "Tài khoản" (nếu chưa có) với 1 tài khoản quản trị mẫu và khóa sheet chỉ chủ file sửa được.
 *  3. Vào sheet "Tài khoản": đổi số điện thoại + mật khẩu của tài khoản mẫu, thêm tài khoản khác nếu cần.
 *  4. Triển khai → Tùy chọn triển khai mới → Loại: Ứng dụng web
 *       Thực thi dưới dạng: Tôi   ·   Người có quyền truy cập: Bất kỳ ai
 *     Sao chép "URL ứng dụng web" (…/exec) → đặt vào biến môi trường GOOGLE_SCRIPT_URL trên Netlify.
 *  5. Lưu ý: nếu file đang "Xuất bản lên web" (pubhtml) thì ai có link cũng xem được dữ liệu và mật khẩu.
 *     Khi đã đặt GOOGLE_SCRIPT_URL, webapp không cần bản xuất bản nữa – có thể "Dừng xuất bản" cho an toàn.
 *  Sửa mã sau này: Triển khai → Quản lý triển khai → Sửa (bút chì) → Phiên bản mới (giữ nguyên URL).
 *
 * SHEET "Tài khoản" (dữ liệu từ dòng 5):
 *  B Số điện thoại (ID đăng nhập) · C Mật khẩu · D Họ tên · E Quyền (Quản trị / Phó Quản trị / Nhập liệu / Chỉ xem / Khách)
 *    Quản trị, Phó Quản trị: toàn quyền. Các quyền còn lại (Nhập liệu, Khách, Chỉ xem): chỉ xem số liệu tổng quan, không xem thông tin cá nhân; Nhập liệu / Khách được nhập hồ sơ mới.
 *  F Trạng thái (Hoạt động / Khóa) · G Ghi chú
 *  Đổi mật khẩu, đổi quyền, khóa hoặc xóa dòng → phiên đăng nhập cũ bị đăng xuất trong vòng ~20 giây.
 */

var SH_RES = 'Người cư trú';
var SH_OWN = 'Cơ sở';
var SH_CFG = 'Tổng quan';
var SH_ACC = 'Tài khoản';
var FIRST = 5; // dòng dữ liệu đầu tiên (dòng 4 là tiêu đề) – dùng cho mọi sheet

// Cột nhập liệu (1 = A). Cột tự động (công thức) không bao giờ bị ghi đè.
var RES = { fac: 2, name: 3, photo: 5, dob: 6, cccd: 7, addr: 8, job: 9, income: 10, moveIn: 12, status: 14, regDate: 15, expiry: 16, fine: 18, note: 19 };
var OWN = { fac: 2, name: 3, photo: 5, dob: 6, cccd: 7, type: 8, phone: 9, rooms: 10, note: 14 };
var ACC = { phone: 2, pass: 3, name: 4, role: 5, state: 6, note: 7 };
var RES_W = 19, OWN_W = 14, ACC_W = 7;
var DATES = { dob: 1, moveIn: 1, regDate: 1, expiry: 1 };
var TEXTS = { cccd: 1, phone: 1 };

var ROLES = ['Quản trị', 'Phó Quản trị', 'Nhập liệu', 'Chỉ xem', 'Khách'];
var ADMIN = { 'Quản trị': 1, 'Phó Quản trị': 1 };
function limited_(role) { return !ADMIN[role]; } // mọi quyền khác Quản trị / Phó Quản trị

var SESSION_SEC = 6 * 3600;   // phiên đăng nhập 6 giờ (giới hạn của CacheService)
var MAX_FAIL = 5;             // sai mật khẩu 5 lần → khóa tạm
var FAIL_LOCK_SEC = 15 * 60;  // 15 phút

// ================= cài đặt =================
function caiDat() {
  var ss = ss_();
  var sh = ss.getSheetByName(SH_ACC);
  if (!sh) {
    sh = ss.insertSheet(SH_ACC);
    sh.getRange('A1').setValue('TÀI KHOẢN ĐĂNG NHẬP WEBAPP XÓM TRỌ VĂN HÓA').setFontWeight('bold').setFontSize(13);
    sh.getRange('A2').setValue('ID đăng nhập là số điện thoại. Quyền: Quản trị / Phó Quản trị / Nhập liệu / Chỉ xem / Khách. Trạng thái "Khóa" để chặn đăng nhập. Đổi mật khẩu ở đây → webapp tự đăng xuất phiên cũ.');
    sh.getRange(4, 1, 1, ACC_W).setValues([['STT', 'Số điện thoại (ID)', 'Mật khẩu', 'Họ tên', 'Quyền', 'Trạng thái', 'Ghi chú']])
      .setFontWeight('bold').setBackground('#0f5132').setFontColor('#ffffff');
    sh.getRange(FIRST, 1, 1, ACC_W).setValues([[1, '0900000000', 'DoiMatKhau@2026', 'Quản trị viên', 'Quản trị', 'Hoạt động', 'TÀI KHOẢN MẪU – đổi số điện thoại và mật khẩu ngay']]);
    sh.setFrozenRows(4);
    sh.setColumnWidths(1, ACC_W, 150);
  }
  sh.getRange(FIRST, ACC.phone, 500, 1).setNumberFormat('@');
  sh.getRange(FIRST, ACC.pass, 500, 1).setNumberFormat('@');
  sh.getRange(FIRST, ACC.role, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ROLES, true).build());
  sh.getRange(FIRST, ACC.state, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Hoạt động', 'Khóa'], true).build());
  // Chỉ chủ file được sửa sheet tài khoản.
  var p = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0] || sh.protect();
  p.setDescription('Tài khoản đăng nhập webapp');
  p.removeEditors(p.getEditors());
  if (p.canDomainEdit()) p.setDomainEdit(false);
  SpreadsheetApp.flush();
  return 'Đã cài đặt sheet "' + SH_ACC + '". Hãy đổi tài khoản mẫu rồi triển khai Ứng dụng web.';
}

// ================= điểm vào =================
function doGet() {
  return json_({ ok: true, app: 'xtvh', note: 'Dùng POST' });
}

function doPost(e) {
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (req.action) {
      case 'login': return json_(login_(req.phone, req.password));
      case 'logout': if (req.token) CacheService.getScriptCache().remove('s_' + req.token); return json_({ ok: true });
      case 'read': { var u = auth_(req.token); return json_({ ok: true, user: pub_(u), data: dataFor_(u) }); }
      case 'write': return json_(write_(auth_(req.token), req.ops || []));
      default: return json_(err_(400, 'Yêu cầu không hợp lệ'));
    }
  } catch (x) {
    if (x && x.code) return json_(err_(x.code, x.message));
    return json_(err_(500, String(x && x.message || x)));
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function err_(code, msg) { return { ok: false, code: code, error: msg }; }
function fail_(code, msg) { var e = new Error(msg); e.code = code; return e; }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(n) {
  var s = ss_().getSheetByName(n);
  if (!s) throw fail_(500, 'Không tìm thấy sheet "' + n + '"' + (n === SH_ACC ? ' – hãy chạy hàm caiDat trong Apps Script' : ''));
  return s;
}

// ================= đăng nhập =================
function normPhone_(s) {
  var d = String(s || '').replace(/\D/g, '');
  if (d.indexOf('84') === 0 && d.length === 11) d = '0' + d.slice(2);
  if (d.length === 9) d = '0' + d; // ô số bị mất số 0 đầu
  return d;
}

function hash_(s) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8));
}

function accounts_() {
  var sh = sheet_(SH_ACC), last = sh.getLastRow();
  if (last < FIRST) return [];
  var v = sh.getRange(FIRST, 1, last - FIRST + 1, ACC_W).getDisplayValues(), out = [];
  for (var i = 0; i < v.length; i++) {
    var phone = normPhone_(v[i][ACC.phone - 1]), pass = String(v[i][ACC.pass - 1]);
    if (!phone || !pass) continue;
    var role = String(v[i][ACC.role - 1]).trim();
    if (ROLES.indexOf(role) < 0) role = 'Chỉ xem';
    var active = String(v[i][ACC.state - 1]).trim() !== 'Khóa';
    var a = { phone: phone, pass: pass, name: String(v[i][ACC.name - 1]).trim(), role: role, active: active };
    // Dấu vân tay: đổi mật khẩu / quyền / trạng thái → phiên cũ hết hiệu lực.
    a.fp = hash_([phone, pass, role, active].join('|'));
    out.push(a);
  }
  return out;
}

function findAcc_(phone) {
  var list = accounts_();
  for (var i = 0; i < list.length; i++) if (list[i].phone === phone) return list[i];
  return null;
}

function login_(phone, password) {
  phone = normPhone_(phone);
  password = String(password || '');
  if (!phone || !password) return err_(400, 'Nhập số điện thoại và mật khẩu');
  var cache = CacheService.getScriptCache(), fk = 'f_' + phone;
  var fails = Number(cache.get(fk) || 0);
  if (fails >= MAX_FAIL) return err_(429, 'Sai mật khẩu quá ' + MAX_FAIL + ' lần. Thử lại sau 15 phút.');
  var a = findAcc_(phone);
  if (!a || a.pass !== password) {
    cache.put(fk, String(fails + 1), FAIL_LOCK_SEC);
    Utilities.sleep(800);
    return err_(401, 'Số điện thoại hoặc mật khẩu không đúng');
  }
  if (!a.active) return err_(403, 'Tài khoản đang bị khóa');
  cache.remove(fk);
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('s_' + token, JSON.stringify({ phone: a.phone, fp: a.fp }), SESSION_SEC);
  return { ok: true, token: token, user: pub_(a), data: dataFor_(a) };
}

function auth_(token) {
  if (!token) throw fail_(401, 'Chưa đăng nhập');
  var raw = CacheService.getScriptCache().get('s_' + token);
  if (!raw) throw fail_(401, 'Phiên đăng nhập đã hết hạn, hãy đăng nhập lại');
  var s = JSON.parse(raw), a = findAcc_(s.phone);
  if (!a || a.fp !== s.fp || !a.active) {
    CacheService.getScriptCache().remove('s_' + token);
    throw fail_(401, 'Thông tin đăng nhập đã thay đổi trên Google Sheet, hãy đăng nhập lại');
  }
  return a;
}

function pub_(a) { return { phone: a.phone, name: a.name, role: a.role }; }

// ================= đọc dữ liệu =================
function iso_(s) {
  s = String(s || '').trim();
  var m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return /^\d{4}-\d\d-\d\d$/.test(s) ? s : '';
}

function photo_(v) {
  try { if (v && typeof v.getContentUrl === 'function') return v.getContentUrl() || ''; } catch (e) {}
  return '';
}

function readRows_(sh, map, width) {
  var last = sh.getLastRow();
  if (last < FIRST) return [];
  var n = last - FIRST + 1, rng = sh.getRange(FIRST, 1, n, width);
  var disp = rng.getDisplayValues(), vals = null;
  try { vals = rng.getValues(); } catch (e) {}
  var out = [];
  for (var i = 0; i < n; i++) {
    var d = disp[i];
    if (!String(d[map.fac - 1]).trim() && !String(d[map.name - 1]).trim()) continue;
    var o = { row: FIRST + i };
    for (var k in map) {
      if (k === 'photo') { o.photo = vals ? photo_(vals[i][map.photo - 1]) : ''; continue; }
      var x = String(d[map[k] - 1] || '').trim();
      o[k] = DATES[k] ? iso_(x) : x;
    }
    out.push(o);
  }
  return out;
}

function readAll_() {
  var days = 20;
  try { days = Number(sheet_(SH_CFG).getRange('C4').getValue()) || 20; } catch (e) {}
  return {
    cfg: { days: days },
    own: readRows_(sheet_(SH_OWN), OWN, OWN_W),
    res: readRows_(sheet_(SH_RES), RES, RES_W),
    at: new Date().toISOString()
  };
}

// Khách chỉ nhận tên cơ sở và tình trạng tạm trú (để đếm số liệu) – không có thông tin cá nhân.
function dataFor_(user) {
  var d = readAll_();
  if (!limited_(user.role)) return d;
  return {
    cfg: d.cfg, at: d.at,
    own: d.own.map(function (o, i) { return { row: i + 1, fac: o.fac }; }),
    res: d.res.map(function (r, i) { return { row: i + 1, status: r.status, expiry: r.expiry }; })
  };
}

// ================= ghi dữ liệu =================
function write_(user, ops) {
  if (user.role === 'Chỉ xem') return err_(403, 'Tài khoản chỉ có quyền xem');
  for (var i = 0; i < ops.length; i++)
  {
    if ((ops[i].op === 'wipe' || ops[i].op === 'cfg') && !ADMIN[user.role])
      return err_(403, 'Chỉ tài khoản Quản trị hoặc Phó Quản trị được thực hiện thao tác này');
    if (limited_(user.role) && !(ops[i].op === 'add' && (ops[i].kind === 'res' || ops[i].kind === 'own')))
      return err_(403, 'Tài khoản của bạn chỉ được nhập thêm hồ sơ mới');
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var results = ops.map(apply_);
    SpreadsheetApp.flush();
    if (limited_(user.role)) return { ok: true, user: pub_(user), data: dataFor_(user) };
    return { ok: true, results: results, user: pub_(user), data: readAll_() };
  } catch (x) {
    return err_(x.code || 409, String(x && x.message || x));
  } finally {
    lock.releaseLock();
  }
}

function toDate_(s) {
  var m = String(s || '').match(/^(\d{4})-(\d\d)-(\d\d)$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : '';
}

function writeRow_(sh, row, map, rec) {
  for (var k in map) {
    if (!(k in rec)) continue;
    var cell = sh.getRange(row, map[k]);
    var v = rec[k] == null ? '' : rec[k];
    if (k === 'photo') {
      if (!v) { cell.clearContent(); continue; }
      if (String(v).indexOf('data:image') === 0) {
        try { cell.setValue(SpreadsheetApp.newCellImage().setSourceUrl(v).setAltTextTitle(String(rec.name || '')).build()); } catch (e) {}
      }
      continue;
    }
    if (DATES[k]) { var d = toDate_(v); cell.setValue(d); if (d) cell.setNumberFormat('dd/mm/yyyy'); continue; }
    if (TEXTS[k]) { cell.setNumberFormat('@'); cell.setValue(String(v)); continue; }
    if (k === 'rooms' && v !== '' && !isNaN(v)) { cell.setValue(Number(v)); continue; }
    cell.setValue(String(v));
  }
}

function clearRow_(sh, row, map) {
  for (var k in map) sh.getRange(row, map[k]).clearContent();
}

function emptyRow_(sh, map) {
  var last = Math.max(sh.getLastRow(), FIRST), n = last - FIRST + 1;
  var v = sh.getRange(FIRST, 1, n, Math.max(map.fac, map.name)).getDisplayValues();
  for (var i = 0; i < n; i++)
    if (!String(v[i][map.fac - 1]).trim() && !String(v[i][map.name - 1]).trim()) return FIRST + i;
  // Hết dòng trống: thêm dòng và chép công thức của dòng trên.
  var row = last + 1, w = sh.getLastColumn();
  sh.insertRowAfter(last);
  sh.getRange(last, 1, 1, w).copyTo(sh.getRange(row, 1, 1, w));
  clearRow_(sh, row, map);
  return row;
}

// Dòng vẫn là bản ghi webapp đang sửa (phòng khi dòng bị chèn/xóa trên sheet).
function checkRow_(sh, map, row, expect) {
  if (!row || row < FIRST) throw fail_(409, 'Dòng không hợp lệ');
  if (expect == null) return;
  var cur = String(sh.getRange(row, map.name).getDisplayValue()).trim();
  if (cur !== String(expect).trim()) throw fail_(409, 'Dữ liệu trên Google Sheet vừa thay đổi. Đã tải lại, hãy thử lại.');
}

function apply_(op) {
  var isRes = op.kind === 'res';
  var sh = sheet_(isRes ? SH_RES : SH_OWN), map = isRes ? RES : OWN;
  switch (op.op) {
    case 'add': {
      var row = emptyRow_(sh, map);
      writeRow_(sh, row, map, op.rec || {});
      return { row: row };
    }
    case 'update': {
      checkRow_(sh, map, op.row, op.expect);
      var old = String(sh.getRange(op.row, map.fac).getDisplayValue()).trim();
      writeRow_(sh, op.row, map, op.rec || {});
      if (!isRes && op.rec && op.rec.fac && old && old !== op.rec.fac) renameFac_(old, op.rec.fac);
      return { row: op.row };
    }
    case 'delete': {
      checkRow_(sh, map, op.row, op.expect);
      if (!isRes) {
        var fac = String(sh.getRange(op.row, map.fac).getDisplayValue()).trim();
        if (fac) deleteResOf_(fac);
      }
      clearRow_(sh, op.row, map);
      return { row: op.row };
    }
    case 'cfg':
      sheet_(SH_CFG).getRange('C4').setValue(Number(op.days) || 20);
      return {};
    case 'wipe':
      [[SH_RES, RES], [SH_OWN, OWN]].forEach(function (p) {
        var s = sheet_(p[0]), last = s.getLastRow();
        for (var r = FIRST; r <= last; r++) clearRow_(s, r, p[1]);
      });
      return {};
    default:
      throw fail_(400, 'Thao tác không hỗ trợ: ' + op.op);
  }
}

function renameFac_(from, to) {
  var sh = sheet_(SH_RES), last = sh.getLastRow();
  if (last < FIRST) return;
  var rng = sh.getRange(FIRST, RES.fac, last - FIRST + 1, 1), v = rng.getValues(), changed = false;
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === from) { v[i][0] = to; changed = true; }
  if (changed) rng.setValues(v);
}

function deleteResOf_(fac) {
  var sh = sheet_(SH_RES), last = sh.getLastRow();
  if (last < FIRST) return;
  var v = sh.getRange(FIRST, RES.fac, last - FIRST + 1, 1).getDisplayValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === fac) clearRow_(sh, FIRST + i, RES);
}
