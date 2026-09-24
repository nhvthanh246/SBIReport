/* Ghi kết quả ra .xlsx và TSV.
 * Điểm mấu chốt: ép kiểu ô khi ghi. Nếu để SheetJS tự suy, `Sender id`
 * 26793145563657353 (> 2^53) sẽ mất chính xác và timestamp sẽ bị đổi định dạng.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(root.Parse);
  else root.Writer = factory(root.Parse);
})(typeof self !== 'undefined' ? self : this, function (P) {
  'use strict';

  var DATE_FMT = 'm/d/yyyy';

  function makeCell(value, type) {
    if (value === null || value === undefined || value === '') return null;
    if (type === 'date') {
      var d = Number(value);
      return isFinite(d) ? { t: 'n', v: d, z: DATE_FMT } : { t: 's', v: String(value) };
    }
    if (type === 'number') {
      var n = Number(value);
      return isFinite(n) ? { t: 'n', v: n } : { t: 's', v: String(value) };
    }
    if (type === 'text') return { t: 's', v: P.cellToString(value) };
    /* auto */
    if (typeof value === 'number') return { t: 'n', v: value };
    if (typeof value === 'boolean') return { t: 'b', v: value };
    if (value instanceof Date) return { t: 'd', v: value };
    return { t: 's', v: P.cellToString(value) };
  }

  /** Đánh dấu cột nào phải ghi kiểu text, dựa trên tên header. */
  function typesForHeaders(headers, textColumns) {
    var wants = (textColumns || []).map(function (x) { return P.normText(x); });
    return headers.map(function (h) {
      var hh = P.normText(h);
      for (var i = 0; i < wants.length; i++) {
        if (hh === wants[i] || (wants[i] && hh.indexOf(wants[i]) !== -1)) return 'text';
      }
      return 'auto';
    });
  }

  /** Mảng 2 chiều -> worksheet, tôn trọng mảng kiểu `types` theo cột. */
  function aoaToSheet(rows, types, opts) {
    opts = opts || {};
    var ws = {};
    var maxCol = 0;
    rows.forEach(function (row) { if (row && row.length > maxCol) maxCol = row.length; });

    for (var r = 0; r < rows.length; r++) {
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c++) {
        var cell = makeCell(row[c], (types && types[c]) || 'auto');
        if (cell) ws[XLSX.utils.encode_cell({ r: r, c: c })] = cell;
      }
    }
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 },
      e: { r: Math.max(rows.length - 1, 0), c: Math.max(maxCol - 1, 0) } });
    if (opts.autofilter && rows.length > 0) {
      ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 },
        e: { r: rows.length - 1, c: Math.max(maxCol - 1, 0) } }) };
    }
    if (opts.colWidths) ws['!cols'] = opts.colWidths;
    return ws;
  }

  function bookToBuffer(wb) {
    return XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
  }

  function newBook() { return { SheetNames: [], Sheets: {} }; }

  function addSheet(wb, name, ws) {
    /* Excel giới hạn tên sheet 31 ký tự */
    var safe = String(name).slice(0, 31).replace(/[\\/?*[\]:]/g, '_');
    var n = safe, i = 2;
    while (wb.Sheets[n]) { n = safe.slice(0, 28) + '_' + (i++); }
    wb.SheetNames.push(n);
    wb.Sheets[n] = ws;
    return n;
  }

  /* ---------- workbook đầu ra ---------- */

  /** Session Report đã làm sạch: sheet gốc (bớt dòng bị lọc) + sheet staging. */
  function buildSessionWorkbook(result, cfg) {
    var wb = newBook();
    var t1 = typesForHeaders(result.headers, cfg.textColumns);
    var rows1 = [result.headers].concat(result.keptRows.map(function (row) {
      var out = [];
      for (var i = 0; i < result.headers.length; i++) {
        out.push(row[i] === undefined || row[i] === null ? '' : row[i]);
      }
      return out;
    }));
    addSheet(wb, result.sheetName, aoaToSheet(rows1, t1, { autofilter: true }));

    var t2 = typesForHeaders(result.staging.headers, cfg.textColumns);
    var rows2 = [result.staging.headers].concat(result.staging.rows);
    addSheet(wb, result.staging.name, aoaToSheet(rows2, t2, { autofilter: true }));
    return bookToBuffer(wb);
  }

  /** Daily Report CCVN đã bỏ cột rỗng — giữ nguyên metadata, header, dữ liệu, dòng Total. */
  function buildCcvnWorkbook(result, cfg) {
    var wb = newBook();
    var types = typesForHeaders(result.headers, cfg.textColumns);
    addSheet(wb, result.sheetName, aoaToSheet(result.cleaned, types));
    return bookToBuffer(wb);
  }

  /** Một khối dán -> workbook 1 sheet (có dòng header để dễ soát, khi dán thì bỏ dòng đầu). */
  function buildBlockWorkbook(block) {
    var wb = newBook();
    var rows = [block.headers].concat(block.rows);
    addSheet(wb, block.title, aoaToSheet(rows, block.types, { autofilter: true }));
    return bookToBuffer(wb);
  }

  /** Cả 3 khối trong 1 file, mỗi khối 1 sheet. */
  function buildAllBlocksWorkbook(blocks) {
    var wb = newBook();
    ['block1', 'block2', 'block3'].forEach(function (k) {
      var b = blocks[k];
      if (!b) return;
      addSheet(wb, b.title, aoaToSheet([b.headers].concat(b.rows), b.types, { autofilter: true }));
    });
    return bookToBuffer(wb);
  }

  /* ---------- TSV cho clipboard ---------- */

  /**
   * Excel khi dán TSV vẫn hiểu quy tắc nháy kép, nên trường chứa tab/xuống dòng/nháy
   * được bọc lại thay vì bị cắt — `Case Log` và `Description` có xuống dòng thật.
   */
  function tsvField(v) {
    if (v === null || v === undefined) return '';
    var s = P.cellToString(v);
    if (/[\t\r\n"]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function blockToTsv(block, includeHeader) {
    var rows = includeHeader ? [block.headers].concat(block.rows) : block.rows;
    return rows.map(function (row) {
      return row.map(tsvField).join('\t');
    }).join('\r\n');
  }

  /* ---------- tải file ---------- */

  function download(buffer, filename) {
    var blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  }

  return {
    makeCell: makeCell,
    typesForHeaders: typesForHeaders,
    aoaToSheet: aoaToSheet,
    buildSessionWorkbook: buildSessionWorkbook,
    buildCcvnWorkbook: buildCcvnWorkbook,
    buildBlockWorkbook: buildBlockWorkbook,
    buildAllBlocksWorkbook: buildAllBlocksWorkbook,
    blockToTsv: blockToTsv,
    download: download
  };
});
