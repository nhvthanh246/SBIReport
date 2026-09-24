/* Đọc & chuẩn hoá dữ liệu từ workbook SheetJS.
 * Nguyên tắc: luôn lấy GIÁ TRỊ ĐÃ TÍNH của ô (cell.v), không lấy công thức (cell.f) —
 * file báo cáo tổng chứa kết quả paste-values chứ không phải công thức gốc của file OP.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Parse = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var MS_PER_DAY = 86400000;
  /* Excel coi 1900 là năm nhuận (sai) nên mốc quy đổi là 1899-12-30 */
  var EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

  /* ---------- chuỗi ---------- */

  function nfc(s) {
    return typeof s.normalize === 'function' ? s.normalize('NFC') : s;
  }

  /** Giá trị ô -> chuỗi, giữ nguyên nội dung (không trim) để so sánh từng ô chính xác. */
  function cellToString(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.toISOString();
    return nfc(String(v));
  }

  /** Chuẩn hoá để SO SÁNH: NFC + bỏ khoảng trắng thừa + thường hoá. */
  function normText(v) {
    return cellToString(v).replace(/\s+/g, ' ').trim().toLowerCase();
  }

  /** Chuẩn hoá TÊN người để tra email — giống VLOOKUP của Excel (không phân biệt hoa thường). */
  function normName(v) {
    return cellToString(v).replace(/\s+/g, ' ').trim().toUpperCase();
  }

  function isBlank(v) {
    return v === null || v === undefined || String(v).trim() === '';
  }

  /* ---------- ngày ---------- */

  function serialToDate(serial) {
    return new Date(EXCEL_EPOCH_UTC + Math.round(serial) * MS_PER_DAY);
  }

  function dateToSerial(y, m, d) {
    return Math.round((Date.UTC(y, m - 1, d) - EXCEL_EPOCH_UTC) / MS_PER_DAY);
  }

  function serialToISO(serial) {
    return serialToDate(serial).toISOString().slice(0, 10);
  }

  /**
   * Chuẩn hoá ô ngày-giờ của Salesforce về dạng `yyyy/mm/dd hh:mm`.
   *
   * Export đổi định dạng giữa chừng: bản 13/9 ra `2026/09/13 11:33`, còn bản 14/9
   * và 15/9 ra `10:09 14/09/2026`. File tổng luôn dùng dạng đầu, nên trước đây phải
   * sửa tay. Hàm này nhận cả hai và luôn trả về dạng đầu; giá trị lạ thì giữ nguyên.
   */
  function normalizeDateTime(v) {
    var s = cellToString(v).trim();
    if (!s) return s;
    function pad(n) { return (n < 10 ? '0' : '') + n; }

    /* hh:mm dd/mm/yyyy */
    var m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s+(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
    if (m) {
      return m[6] + '/' + pad(+m[5]) + '/' + pad(+m[4]) + ' ' + pad(+m[1]) + ':' + m[2];
    }
    /* yyyy/mm/dd hh:mm — đã đúng dạng, chỉ chuẩn hoá lại số 0 đứng đầu */
    m = s.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (m) {
      return m[1] + '/' + pad(+m[2]) + '/' + pad(+m[3]) + ' ' + pad(+m[4]) + ':' + m[5];
    }
    return s;
  }

  /** Tách giá trị ngày thành {y, m, d}, trả null nếu không đọc được. */
  function toDateParts(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) {
      return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
    }
    function fromSerial(n) {
      var dt = serialToDate(n);
      return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
    }
    if (typeof v === 'number') return isFinite(v) && v > 0 ? fromSerial(v) : null;

    var s = cellToString(v).trim();
    if (!s) return null;
    var m = s.match(/^(\d{4})\s*[-/.]+\s*(\d{1,2})\s*[-/.]+\s*(\d{1,2})/);
    if (m) return { y: +m[1], m: +m[2], d: +m[3] };
    m = s.match(/^(\d{1,2})\s*[-/.]+\s*(\d{1,2})\s*[-/.]+\s*(\d{4})/);
    if (m) return { y: +m[3], m: +m[2], d: +m[1] };
    if (/^\d+(\.\d+)?$/.test(s)) {
      var n = parseFloat(s);
      return n > 0 ? fromSerial(n) : null;
    }
    return null;
  }

  /**
   * Serial của ngày sau khi ĐẢO ngày <-> tháng, null nếu đảo không hợp lệ.
   *
   * Dùng để bắt lỗi locale trong file OP: gõ "9/1/2026" với ý 01/9 nhưng Excel
   * kiểu Mỹ đọc thành 09/01 rồi đổi luôn ra số. Ngày > 12 thì đảo không được
   * (không có tháng 13) — chính vì vậy các ngày 13-31 không dính lỗi này.
   */
  function swappedSerial(v) {
    var p = toDateParts(v);
    if (!p || p.d < 1 || p.d > 12 || p.m < 1 || p.m > 12) return null;
    return dateToSerial(p.y, p.d, p.m);
  }

  /**
   * Quy mọi kiểu ngày về serial Excel (số nguyên), trả null nếu không đọc được.
   * Cột DATE trong file OP lẫn cả hai kiểu: serial (46278) và text ("13/9/2026").
   * Bỏ qua một kiểu là mất dữ liệu — mẫu có 441 dòng serial + 222 dòng text.
   */
  function toDateSerial(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) {
      return dateToSerial(v.getFullYear(), v.getMonth() + 1, v.getDate());
    }
    if (typeof v === 'number') {
      return isFinite(v) && v > 0 ? Math.floor(v) : null;
    }

    var s = cellToString(v).trim();
    if (!s) return null;

    /* Dấu phân cách cho phép lặp và có khoảng trắng: file OP thật có lỗi gõ
       "13/09//2026" (8 dòng của 1 supporter). Nếu đòi đúng 1 dấu thì cả khối
       giờ làm của người đó bị bỏ qua mà không ai biết. */
    /* yyyy-mm-dd hoặc yyyy/mm/dd, có thể kèm giờ và offset */
    var m = s.match(/^(\d{4})\s*[-/.]+\s*(\d{1,2})\s*[-/.]+\s*(\d{1,2})/);
    if (m) return dateToSerial(+m[1], +m[2], +m[3]);

    /* d/m/yyyy, dd-mm-yyyy, d.m.yyyy — định dạng VN, ngày trước tháng */
    m = s.match(/^(\d{1,2})\s*[-/.]+\s*(\d{1,2})\s*[-/.]+\s*(\d{4})/);
    if (m) return dateToSerial(+m[3], +m[2], +m[1]);

    /* chuỗi toàn số = serial lưu dạng text */
    if (/^\d+(\.\d+)?$/.test(s)) {
      var n = parseFloat(s);
      return n > 0 ? Math.floor(n) : null;
    }
    return null;
  }

  /* ---------- đọc sheet ---------- */

  /**
   * Sheet -> ma trận 2 chiều các giá trị thô.
   * Duyệt theo các ô thực có thay vì theo `!ref`: sheet `1. CSKH` của file OP khai
   * dimension A1:IR49633 nhưng chỉ có ~8700 dòng dữ liệu thật.
   */
  function sheetToMatrix(ws) {
    var rows = [];
    var maxCol = -1;
    if (!ws) return { rows: rows, maxCol: maxCol };

    Object.keys(ws).forEach(function (key) {
      if (key.charCodeAt(0) === 33 /* '!' */) return;
      var addr = XLSX.utils.decode_cell(key);
      var cell = ws[key];
      if (!cell || cell.v === undefined || cell.v === null) return;
      var r = addr.r, c = addr.c;
      if (!rows[r]) rows[r] = [];
      rows[r][c] = cell.v;
      if (c > maxCol) maxCol = c;
    });

    for (var i = 0; i < rows.length; i++) {
      if (!rows[i]) rows[i] = [];
    }
    return { rows: rows, maxCol: maxCol };
  }

  function rowIsEmpty(row, maxCol) {
    if (!row) return true;
    for (var c = 0; c <= maxCol; c++) {
      if (!isBlank(row[c])) return false;
    }
    return true;
  }

  /**
   * Tìm cột theo tên header. Ưu tiên khớp chính xác, sau đó mới khớp chứa.
   * Khớp chính xác là bắt buộc ở Session Report: "Customer" phải ra cột C chứ không
   * phải "Customer wait time"; "First response time" chứ không phải "Agent first
   * response time". Khớp chứa dùng cho header dài của Salesforce
   * ("Case information/案件情報: Case Number/案件番号" <- "Case Number").
   */
  function findColumn(headers, name) {
    var want = normText(name);
    var i;
    for (i = 0; i < headers.length; i++) {
      if (normText(headers[i]) === want) return i;
    }
    for (i = 0; i < headers.length; i++) {
      if (normText(headers[i]).indexOf(want) !== -1) return i;
    }
    return -1;
  }

  /** Dò dòng header bằng chuỗi mốc, trả index 0-based; -1 nếu không thấy. */
  function findHeaderRow(rows, marker, maxScan) {
    var want = normText(marker);
    var limit = Math.min(rows.length, maxScan || 50);
    for (var r = 0; r < limit; r++) {
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c++) {
        if (normText(row[c]).indexOf(want) !== -1) return r;
      }
    }
    return -1;
  }

  /* ---------- luật lọc ---------- */

  function matchRule(value, rule) {
    var v = cellToString(value);
    switch (rule.op) {
      case 'equals': return v.trim() === cellToString(rule.value).trim();
      case 'equalsIgnoreCase': return normText(v) === normText(rule.value);
      case 'contains': return normText(v).indexOf(normText(rule.value)) !== -1;
      case 'isEmpty': return v.trim() === '';
      case 'regex':
        try { return new RegExp(rule.value, 'i').test(v); } catch (e) { return false; }
      default: return false;
    }
  }

  return {
    nfc: nfc,
    cellToString: cellToString,
    normText: normText,
    normName: normName,
    isBlank: isBlank,
    serialToDate: serialToDate,
    dateToSerial: dateToSerial,
    serialToISO: serialToISO,
    toDateSerial: toDateSerial,
    toDateParts: toDateParts,
    swappedSerial: swappedSerial,
    normalizeDateTime: normalizeDateTime,
    sheetToMatrix: sheetToMatrix,
    rowIsEmpty: rowIsEmpty,
    findColumn: findColumn,
    findHeaderRow: findHeaderRow,
    matchRule: matchRule
  };
});
