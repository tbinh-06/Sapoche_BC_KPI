/**
 * SAPOCHE — Báo cáo KPI nhân sự
 * Apps Script Web App đọc dữ liệu từ view kpi_nhansu trên Google Sheet.
 *
 * ⚠️ BẮT BUỘC LÀM TRƯỚC (vì Sheet đang dùng Connected Sheets / BigQuery):
 *   Trong trình soạn Apps Script → cột trái mục "Services" (dấu +)
 *   → chọn "Google Sheets API" → Add.  Identifier để nguyên là "Sheets".
 *   Không bật cái này thì tab DATASOURCE sẽ báo:
 *   "Hành động này không được hỗ trợ cho trang tính DATASOURCE".
 *
 * Triển khai:
 *   1. Mở Sheet → Extensions → Apps Script
 *   2. Dán file này vào Code.gs, tạo file HTML tên "index" và dán index.html
 *   3. Bật Sheets API như trên
 *   4. Chạy hàm kiemTra() một lần để cấp quyền và xem log
 *   5. Deploy → New deployment → Web app
 *        Execute as: Me
 *        Who has access: Anyone with Google account
 */

var CONFIG = {
  TAB: '',                // để trống = tự dò. Điền tên tab nếu muốn ép cứng
  BO_TRANG_THAI: ['Đã huỷ'],
  FIELD_LY_DO_LUI: '',    // để trống = tự dò key chứa "lui"/"ly_do" trong form_json
  CACHE_PHUT: 0,          // >0 để bật cache (VD 10). 0 = luôn đọc mới
  COT_TOI_DA: 'BZ'        // vùng cột tối đa khi đọc qua Sheets API
};

var COT_BAT_BUOC = ['ten_cv', 'nguoi_phu_trach', 'trang_thai', 'diem_task'];

/*
 * Thứ tự trường trong mảng trả về cho giao diện.
 * PHẢI khớp biến F trong index.html.
 *  0 ten_cv · 1 nguoi_phu_trach · 2 phong_ban · 3 trang_thai · 4 ngay_giao
 *  5 so_lan_doi_han · 6 diem_HT · 7 diem_TN_w · 8 diem_NL_w · 9 diem_TN · 10 diem_NL
 * 11 diem_task · 12 ket_qua_text · 13 ly_do_lui · 14 link_cv · 15 avatar
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SAPOCHE — Báo cáo KPI nhân sự')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Hàm duy nhất mà giao diện gọi. */
function getTasks() {
  try {
    if (CONFIG.CACHE_PHUT > 0) {
      var hit = docCacheGet_();
      if (hit) return hit;
    }
    var res = docDoc_();
    if (CONFIG.CACHE_PHUT > 0) docCacheSet_(res);
    return res;
  } catch (e) {
    return { ok: false, error: e.message, rows: [], warn: [] };
  }
}

/* ================= ĐỌC Ô: hỗ trợ cả GRID lẫn DATASOURCE ================= */

/** Tab có phải kiểu DATASOURCE (Connected Sheets) không. */
function laDatasource_(sh) {
  try { return sh.getType() === SpreadsheetApp.SheetType.DATASOURCE; }
  catch (e) { return false; }
}

/**
 * Đọc vùng theo GridRange (sheetId + chỉ số dòng/cột).
 * Tab DATASOURCE của Connected Sheets KHÔNG nhận range kiểu A1 ('ten'!A1:BZ5),
 * nên phải đi bằng dataFilter. Thử lần lượt:
 *   1. Advanced Service — values.batchGetByDataFilter
 *   2. REST POST values:batchGetByDataFilter (không cần Advanced Service)
 *   3. Range A1 thường (cho tab GRID)
 */
function docQuaAPI_(sh, r0, r1, c0, c1) {
  var id = SpreadsheetApp.getActive().getId();
  var loi = [];

  var g = { sheetId: sh.getSheetId(), startRowIndex: r0, startColumnIndex: c0 };
  if (r1 != null) g.endRowIndex = r1;
  if (c1 != null) g.endColumnIndex = c1;

  var req = {
    dataFilters: [{ gridRange: g }],
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER'
  };

  // --- 1. Advanced Service, dataFilter ---
  if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets &&
      Sheets.Spreadsheets.Values && Sheets.Spreadsheets.Values.batchGetByDataFilter) {
    try {
      var v1 = layValues_(Sheets.Spreadsheets.Values.batchGetByDataFilter(req, id));
      if (v1.length) return v1;
      loi.push('AdvService/dataFilter: 0 dòng');
    } catch (e) { loi.push('AdvService/dataFilter: ' + e.message); }
  }

  // --- 2. REST, dataFilter ---
  try {
    var v2 = postREST_(id, req);
    if (v2.length) return v2;
    loi.push('REST/dataFilter: 0 dòng');
  } catch (e) { loi.push('REST/dataFilter: ' + e.message); }

  // --- 3. Range A1 (tab GRID thường) ---
  var a1 = 'A' + (r0 + 1) + ':' + chuCot_(c1 || 78) + (r1 != null ? r1 : '');
  var range = "'" + String(sh.getName()).replace(/'/g, "''") + "'!" + a1;
  if (typeof Sheets !== 'undefined') {
    try {
      var res3 = Sheets.Spreadsheets.Values.get(id, range, {
        valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'SERIAL_NUMBER'
      });
      if (res3 && res3.values && res3.values.length) return res3.values;
    } catch (e) { loi.push('AdvService/A1: ' + e.message); }
  }

  throw new Error(
    'Không đọc được "' + sh.getName() + '". ' + loi.join(' | ') +
    '  ⇒ CÁCH ĐI VÒNG: tạo tab thường tên BAOCAO_DATA, ô A1 điền  =' +
    sh.getName() + '!A1:BZ50000  rồi đặt CONFIG.TAB = "BAOCAO_DATA".'
  );
}

/** POST values:batchGetByDataFilter bằng token của chính script. */
function postREST_(id, req) {
  var r = UrlFetchApp.fetch(
    'https://sheets.googleapis.com/v4/spreadsheets/' + id + '/values:batchGetByDataFilter',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: JSON.stringify(req),
      muteHttpExceptions: true
    }
  );
  var code = r.getResponseCode(), body = r.getContentText();
  if (code === 401 || code === 403) {
    throw new Error('HTTP ' + code + ' — thiếu quyền. Mở appsscript.json, thêm oauthScopes ' +
      '"https://www.googleapis.com/auth/spreadsheets" và ' +
      '"https://www.googleapis.com/auth/script.external_request", chạy lại để cấp quyền.');
  }
  if (code !== 200) throw new Error('HTTP ' + code + ' — ' + body.slice(0, 180));
  return layValues_(JSON.parse(body));
}

/** Bóc mảng giá trị từ phản hồi batchGetByDataFilter. */
function layValues_(res) {
  var vr = res && res.valueRanges && res.valueRanges[0];
  if (!vr) return [];
  var v = vr.valueRange ? vr.valueRange.values : vr.values;
  return v || [];
}

/** 1 → A, 27 → AA */
function chuCot_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}

/** Đọc n dòng đầu để dò header. */
function docDauTab_(sh) {
  if (!laDatasource_(sh)) {
    try {
      var lr = Math.min(5, sh.getLastRow());
      if (lr < 1) return [];
      return sh.getRange(1, 1, lr, sh.getLastColumn()).getValues();
    } catch (e) { /* rơi xuống API */ }
  }
  return docQuaAPI_(sh, 0, 5, 0, 78);
}

/** Đọc toàn bộ tab. */
function docCaTab_(sh) {
  if (!laDatasource_(sh)) {
    try { return sh.getDataRange().getValues(); }
    catch (e) { /* rơi xuống API */ }
  }
  return docQuaAPI_(sh, 0, null, 0, 78);
}

/* ================= DÒ TAB & HEADER ================= */

function timTab_() {
  var ss = SpreadsheetApp.getActive();
  var tatCa = ss.getSheets();

  var uuTien = CONFIG.TAB ? ss.getSheetByName(CONFIG.TAB) : null;
  var ds = uuTien ? [uuTien].concat(tatCa) : tatCa;
  var daXet = {}, lyDo = [];

  for (var i = 0; i < ds.length; i++) {
    var s = ds[i];
    if (!s || daXet[s.getName()]) continue;
    daXet[s.getName()] = 1;
    var ten = s.getName() + (laDatasource_(s) ? ' [DATASOURCE]' : '');

    var dau;
    try {
      dau = docDauTab_(s);
    } catch (e) {
      lyDo.push('"' + ten + '": ' + e.message);
      continue;
    }
    if (!dau || !dau.length) { lyDo.push('"' + ten + '": đọc ra 0 dòng'); continue; }

    var thieuNhat = null;
    for (var r = 0; r < Math.min(5, dau.length); r++) {
      var head = (dau[r] || []).map(function (x) { return String(x).trim(); });
      var thieu = COT_BAT_BUOC.filter(function (c) { return head.indexOf(c) === -1; });
      if (!thieu.length) return { sheet: s, headerRow: r + 1 };
      if (thieuNhat === null || thieu.length < thieuNhat.length) thieuNhat = thieu;
    }
    lyDo.push('"' + ten + '": 5 dòng đầu không có header khớp (thiếu ' +
              (thieuNhat || COT_BAT_BUOC).join(', ') + '). Dòng 1 đọc được: ' +
              JSON.stringify((dau[0] || []).slice(0, 8)));
  }

  throw new Error('Không dò được tab dữ liệu. Chi tiết từng tab — ' + lyDo.join('  ||  '));
}

/**
 * CHẠY HÀM NÀY TRƯỚC khi có lỗi dò tab.
 * In ra: Sheets API bật chưa, từng tab đọc được bao nhiêu dòng,
 * header thật đang là gì — kèm gợi ý cột nào lệch tên.
 */
function chanDoan() {
  var ss = SpreadsheetApp.getActive();
  Logger.log('Advanced Service "Sheets": ' + (typeof Sheets === 'undefined' ? 'chưa bật (không sao, sẽ dùng REST)' : 'đã bật ✓'));
  try {
    var t = ScriptApp.getOAuthToken();
    Logger.log('OAuth token: lấy được ✓ (' + t.length + ' ký tự)');
  } catch (e) { Logger.log('OAuth token: ✗ ' + e.message); }
  Logger.log('Cột bắt buộc: ' + COT_BAT_BUOC.join(', '));
  Logger.log('');

  ss.getSheets().forEach(function (sh) {
    var kieu = laDatasource_(sh) ? 'DATASOURCE' : 'GRID';
    Logger.log('===== "' + sh.getName() + '" [' + kieu + '] =====');
    var dau;
    try { dau = docDauTab_(sh); }
    catch (e) { Logger.log('  ✗ ' + e.message); Logger.log(''); return; }

    if (!dau || !dau.length) { Logger.log('  ✗ đọc ra 0 dòng'); Logger.log(''); return; }

    for (var r = 0; r < Math.min(3, dau.length); r++) {
      var h = (dau[r] || []).map(function (x) { return String(x).trim(); }).filter(String);
      Logger.log('  Dòng ' + (r + 1) + ' (' + h.length + ' ô): ' + h.slice(0, 30).join(' | '));
    }

    var head = (dau[0] || []).map(function (x) { return String(x).trim(); });
    var thieu = COT_BAT_BUOC.filter(function (c) { return head.indexOf(c) === -1; });
    if (!thieu.length) { Logger.log('  ✓ khớp đủ cột bắt buộc'); }
    else {
      Logger.log('  ✗ thiếu: ' + thieu.join(', '));
      thieu.forEach(function (c) {
        var gan = head.filter(function (h2) {
          return h2 && (h2.toLowerCase().replace(/[\s_]/g, '') === c.toLowerCase().replace(/[\s_]/g, '') ||
                        h2.toLowerCase().indexOf(c.toLowerCase().split('_')[0]) > -1);
        });
        if (gan.length) Logger.log('     "' + c + '" có thể đang là: ' + gan.join(' / '));
      });
    }
    Logger.log('');
  });
}

/** Chạy tay để xem tên tab thật, kiểu tab và tab nào được nhận diện. */
function lietKeTab() {
  Logger.log('Sheets API: ' + (typeof Sheets === 'undefined' ? 'CHƯA BẬT ✗' : 'đã bật ✓'));
  SpreadsheetApp.getActive().getSheets().forEach(function (s) {
    var kt = laDatasource_(s) ? 'DATASOURCE' : 'GRID';
    var kt2 = '';
    try { kt2 = s.getLastRow() + ' dòng × ' + s.getLastColumn() + ' cột'; }
    catch (e) { kt2 = 'không đọc được kích thước (bình thường với DATASOURCE)'; }
    Logger.log('"' + s.getName() + '"  [' + kt + ']  ' + kt2);
  });
  try {
    var t = timTab_();
    Logger.log('→ Tab dữ liệu: "' + t.sheet.getName() + '", header ở dòng ' + t.headerRow);
  } catch (e) {
    Logger.log('→ ' + e.message);
  }
}

/* ================= ĐỌC DỮ LIỆU ================= */

function docDoc_() {
  var tim = timTab_();
  var sh = tim.sheet;
  var headerRow = tim.headerRow;

  var all = docCaTab_(sh);
  if (!all || all.length <= headerRow) {
    throw new Error('Tab "' + sh.getName() + '" chưa có dòng dữ liệu nào.');
  }

  var head = (all[headerRow - 1] || []).map(function (h) { return String(h).trim(); });
  var data = all.slice(headerRow);

  var idx = {};
  head.forEach(function (h, i) { if (h) idx[h] = i; });

  var thieu = [];
  ['ten_cv', 'nguoi_phu_trach', 'phong_ban', 'trang_thai', 'ngay_giao', 'diem_task']
    .forEach(function (c) { if (idx[c] === undefined) thieu.push(c); });
  if (thieu.length) {
    throw new Error('Tab "' + sh.getName() + '" thiếu cột: ' + thieu.join(', ') +
                    '. Header đang có: ' + head.filter(String).join(', '));
  }

  var tz = Session.getScriptTimeZone();
  var iForm = idx['form_json'];
  var bo = {};
  CONFIG.BO_TRANG_THAI.forEach(function (t) { bo[t] = 1; });

  var rows = [], coLui = 0, coLyDo = 0, boQuaNgay = 0, boQuaTT = 0;

  for (var r = 0; r < data.length; r++) {
    var row = data[r] || [];
    if (!row[idx['ten_cv']] && !row[idx['nguoi_phu_trach']]) continue;

    var tt = String(row[idx['trang_thai']] || '').trim();
    if (bo[tt]) { boQuaTT++; continue; }                    // So_CV: trang_thai <> "Đã huỷ"

    var ngay = ngayChuoi_(row[idx['ngay_giao']], tz);
    if (!ngay) { boQuaNgay++; continue; }

    var lui = toNum_(get_(row, idx, 'so_lan_doi_han')) || 0;
    if (lui) coLui++;

    var lyDo = iForm === undefined ? '' : bocLyDo_(row[iForm]);
    if (lyDo) coLyDo++;

    rows.push([
      String(row[idx['ten_cv']] || ''),
      String(row[idx['nguoi_phu_trach']] || ''),
      String(row[idx['phong_ban']] || ''),
      tt,
      ngay,
      lui,
      toNum_(get_(row, idx, 'diem_HT')),
      toNum_(get_(row, idx, 'diem_TN_w')),
      toNum_(get_(row, idx, 'diem_NL_w')),
      toNum_(get_(row, idx, 'diem_TN')),
      toNum_(get_(row, idx, 'diem_NL')),
      toNum_(get_(row, idx, 'diem_task')),
      String(get_(row, idx, 'ket_qua_text') || ''),
      lyDo,
      String(get_(row, idx, 'link_cv') || ''),
      String(get_(row, idx, 'avatar_nguoi_phu_trach') || '')
    ]);
  }

  var w = [];
  if (!coLui) w.push('Cột so_lan_doi_han rỗng toàn bộ — view BigQuery đang CAST(NULL). Cột "Số lần lùi" sẽ trắng.');
  if (iForm === undefined) w.push('Không có cột form_json — không bóc được Lý do lùi.');
  else if (!coLyDo) w.push('Không bóc được Lý do lùi từ form_json. Mở 1 ô form_json xem tên field thật rồi set CONFIG.FIELD_LY_DO_LUI.');
  if (boQuaNgay) w.push(boQuaNgay + ' dòng bị bỏ vì ngay_giao trống hoặc sai định dạng.');

  return {
    ok: true,
    rows: rows,
    warn: w,
    tong: rows.length,
    tab: sh.getName(),
    kieuTab: laDatasource_(sh) ? 'DATASOURCE' : 'GRID',
    headerRow: headerRow,
    boQuaTT: boQuaTT
  };
}

/* ================= TIỆN ÍCH ================= */

function get_(row, idx, ten) { return idx[ten] === undefined ? '' : row[idx[ten]]; }

function toNum_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(v);
  return isNaN(n) ? null : n;
}

/**
 * Chuẩn hoá ngày về chuỗi yyyy-MM-dd.
 * Xử lý 3 kiểu: Date (đọc qua SpreadsheetApp), số serial (đọc qua Sheets API),
 * và chuỗi (BigQuery trả text).
 */
function ngayChuoi_(v, tz) {
  if (v === '' || v === null || v === undefined) return '';

  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }

  if (typeof v === 'number') {
    if (v < 1 || v > 100000) return '';
    // serial của Sheets: 0 = 30/12/1899. Dùng getUTC* để không lệch múi giờ.
    var d = new Date(Math.round((v - 25569) * 86400000));
    if (isNaN(d.getTime())) return '';
    return d.getUTCFullYear() + '-' +
           ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' +
           ('0' + d.getUTCDate()).slice(-2);
  }

  var s = String(v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);           // 2026-05-12 hoặc 2026-05-12T…
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/); // 12/05/2026
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);

  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? '' : Utilities.formatDate(d2, tz, 'yyyy-MM-dd');
}

/** Bóc "Lý do lùi" từ form_json của Base WeWork. */
function bocLyDo_(raw) {
  if (!raw) return '';
  var s = String(raw);
  var obj;
  try { obj = JSON.parse(s); } catch (e) { obj = null; }

  if (obj) {
    var found = '';
    var duyet = function (o) {
      if (found || !o) return;
      if (Array.isArray(o)) { o.forEach(duyet); return; }
      if (typeof o !== 'object') return;
      for (var k in o) {
        var kl = String(k).toLowerCase();
        var hop = CONFIG.FIELD_LY_DO_LUI
          ? (k === CONFIG.FIELD_LY_DO_LUI)
          : (kl.indexOf('ly_do') > -1 || kl.indexOf('lydo') > -1 ||
             (kl.indexOf('lui') > -1 && kl.indexOf('so_lan') === -1));
        if (hop && o[k] && typeof o[k] !== 'object') { found = String(o[k]); return; }
        duyet(o[k]);
      }
    };
    duyet(obj);
    if (found) return lamSach_(found);
  }

  var m = s.match(/"[^"]*l(?:y_?do|ui)[^"]*"\s*:\s*"([^"]*)"/i);
  return m ? lamSach_(m[1]) : '';
}

function lamSach_(t) {
  return String(t).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

/* ================= CACHE (tuỳ chọn) ================= */

function docCacheGet_() {
  try {
    var c = CacheService.getScriptCache();
    var n = Number(c.get('SPC_N') || 0);
    if (!n) return null;
    var parts = [];
    for (var i = 0; i < n; i++) {
      var p = c.get('SPC_' + i);
      if (p === null) return null;
      parts.push(p);
    }
    return JSON.parse(parts.join(''));
  } catch (e) { return null; }
}

function docCacheSet_(res) {
  try {
    var s = JSON.stringify(res), size = 90000, n = Math.ceil(s.length / size);
    if (n > 40) return;
    var c = CacheService.getScriptCache(), o = {}, sec = CONFIG.CACHE_PHUT * 60;
    for (var i = 0; i < n; i++) o['SPC_' + i] = s.substr(i * size, size);
    o['SPC_N'] = String(n);
    c.putAll(o, sec);
  } catch (e) { /* bỏ qua */ }
}

function xoaCache() {
  var c = CacheService.getScriptCache();
  var n = Number(c.get('SPC_N') || 0);
  var keys = ['SPC_N'];
  for (var i = 0; i < n; i++) keys.push('SPC_' + i);
  c.removeAll(keys);
}

/* ================= LÀM MỚI DỮ LIỆU TỪ BIGQUERY ================= */

/**
 * Làm mới toàn bộ Connected Sheets (kpi_nhansu) → BAOCAO_DATA cập nhật theo.
 * Gắn trigger: đồng hồ bên trái trình soạn (Triggers) → Add Trigger
 *   Function: lamMoiDuLieu · Event source: Time-driven · Day timer · 6–7am
 */
function lamMoiDuLieu() {
  var ss = SpreadsheetApp.getActive();
  SpreadsheetApp.enableAllDataSourcesExecution();

  var ds = ss.getDataSources();
  if (!ds.length) { Logger.log('Không có data source nào trong file.'); return; }

  ds.forEach(function (d, i) {
    try {
      d.refreshAllLinkedDataSourceObjects();
      Logger.log('Đã yêu cầu làm mới data source #' + (i + 1));
    } catch (e) {
      Logger.log('Lỗi làm mới #' + (i + 1) + ': ' + e.message);
    }
  });

  Utilities.sleep(30000);          // chờ BigQuery trả kết quả rồi mới đọc lại
  xoaCache();

  var r = getTasks();
  Logger.log(r.ok
    ? ('Sau làm mới: ' + r.tong + ' dòng · tab "' + r.tab + '"')
    : ('Đọc lại lỗi: ' + r.error));
}

/* ================= CHẠY TAY ĐỂ KIỂM TRA ================= */

function kiemTra() {
  var r = getTasks();
  if (!r.ok) { Logger.log('LỖI: ' + r.error); return; }

  Logger.log('Tab: "' + r.tab + '" [' + r.kieuTab + '] · header dòng ' + r.headerRow);
  Logger.log('Số dòng lên báo cáo: ' + r.tong + (r.boQuaTT ? ' (đã loại ' + r.boQuaTT + ' dòng Đã huỷ)' : ''));
  Logger.log('Cảnh báo: ' + (r.warn.length ? r.warn.join(' | ') : 'không'));
  Logger.log('Dòng đầu: ' + JSON.stringify(r.rows[0]));

  var byTT = {}, diem = [], thang = {};
  r.rows.forEach(function (x) {
    byTT[x[3]] = (byTT[x[3]] || 0) + 1;
    if (x[11] !== null) diem.push(x[11]);
    var k = String(x[4]).substr(0, 7);
    thang[k] = (thang[k] || 0) + 1;
  });
  Logger.log('Theo trạng thái: ' + JSON.stringify(byTT));
  Logger.log('Theo tháng: ' + JSON.stringify(thang));
  Logger.log('Diem_TB toàn bộ: ' + (diem.length
    ? (diem.reduce(function (a, b) { return a + b; }, 0) / diem.length).toFixed(2)
    : 'không có dòng nào đã chấm điểm'));
}
