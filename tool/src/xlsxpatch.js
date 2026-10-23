/* Ghi dữ liệu vào File báo cáo tổng bằng cách sửa thẳng XML trong file zip.
 *
 * Vì sao không dùng SheetJS để ghi cả workbook: file tổng có 4 PivotTable,
 * 4 pivotCache (~1.7MB), externalLinks, conditionalFormatting, customXml và hàng
 * nghìn công thức ở các sheet khác. Đọc rồi ghi lại bằng bất kỳ thư viện JS nào
 * cũng làm mất phần lớn số đó.
 *
 * Cách làm ở đây: giải nén zip, CHỈ thay <sheetData> của 3 sheet dữ liệu, giữ
 * nguyên byte của 54 part còn lại, rồi nén lại. Định dạng ô được giữ bằng cách
 * dùng lại đúng thuộc tính `s` của ô cũ ở cùng vị trí.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(root.Parse, root.Pipeline);
  else root.XlsxPatch = factory(root.Parse, root.Pipeline);
})(typeof self !== 'undefined' ? self : this, function (P, Pipeline) {
  'use strict';

  var ROW_RE = /<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g;
  var CELL_RE = /<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g;
  /* XML 1.0 không cho phép các ký tự điều khiển này; \t \n \r thì hợp lệ */
  var BAD_CHARS = /[\x00-\x08\x0B\x0C\x0E-\x1F]/g;

  function esc(s) {
    return String(s).replace(BAD_CHARS, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function unesc(s) {
    return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  }
  function getAttr(tag, name) {
    var m = tag.match(new RegExp('\\s' + name + '="([^"]*)"'));
    return m ? m[1] : null;
  }

  function colToNum(s) {
    var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }
  function numToCol(n) {
    var s = '';
    n += 1;
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
    return s;
  }
  function parseAddr(a) {
    var m = a.match(/^([A-Z]+)(\d+)$/);
    return { c: colToNum(m[1]), r: parseInt(m[2], 10) };
  }
  function setRangeEndRow(ref, endRow) {
    var m = ref.match(/^(\$?[A-Z]+\$?\d+):(\$?[A-Z]+)\$?\d+$/);
    return m ? m[1] + ':' + m[2] + endRow : ref;
  }

  /* ---------- sinh XML cho một ô ---------- */
  function emitCell(addr, style, value, type, formula) {
    var s = (style === null || style === undefined) ? '' : ' s="' + style + '"';
    var blank = (value === '' || value === null || value === undefined);

    if (formula) {
      var numeric = !blank && isFinite(Number(value)) &&
                    (type === 'number' || type === 'date' || typeof value === 'number');
      var vTag = blank ? '' : '<v>' + esc(numeric ? Number(value) : value) + '</v>';
      return '<c r="' + addr + '"' + s + (numeric || blank ? '' : ' t="str"') +
             '><f>' + esc(formula) + '</f>' + vTag + '</c>';
    }
    if (blank) return '<c r="' + addr + '"' + s + '/>';

    if (type === 'date' || type === 'number' || (type === 'auto' && typeof value === 'number')) {
      var n = Number(value);
      if (isFinite(n)) return '<c r="' + addr + '"' + s + '><v>' + n + '</v></c>';
    }
    /* Dùng inlineStr để khỏi phải đụng vào sharedStrings.xml (660KB, dùng chung
       với mọi sheet khác). xml:space="preserve" để giữ khoảng trắng/tab đầu chuỗi
       như giá trị ma_kh "\t A0011242503". */
    return '<c r="' + addr + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' +
           esc(P.cellToString(value)) + '</t></is></c>';
  }

  /* ---------- vá <sheetData> của một sheet ---------- */
  function patchSheet(xml, spec, report) {
    var open = xml.indexOf('<sheetData');
    if (open === -1) throw new Error('Sheet "' + spec.sheetName + '": không có <sheetData>');
    var gt = xml.indexOf('>', open);
    var selfClosing = xml[gt - 1] === '/';
    var headEnd = gt + 1;
    var tailStart, body;
    if (selfClosing) {
      body = '';
      tailStart = headEnd;
    } else {
      tailStart = xml.lastIndexOf('</sheetData>');
      body = xml.slice(headEnd, tailStart);
      tailStart += '</sheetData>'.length;
    }

    /* tách các <row> */
    var rows = [], m;
    ROW_RE.lastIndex = 0;
    while ((m = ROW_RE.exec(body)) !== null) {
      var raw = m[0];
      var ot = raw.slice(0, raw.indexOf('>') + 1);
      var num = parseInt(getAttr(ot, 'r'), 10);
      var innerStart = raw.indexOf('>') + 1;
      var inner = raw.slice(-2) === '/>' ? '' : raw.slice(innerStart, raw.lastIndexOf('</row>'));
      rows.push({ num: num, openTag: ot, inner: inner, raw: raw });
    }
    var byNum = {};
    rows.forEach(function (r) { byNum[r.num] = r; });

    var a = parseAddr(spec.anchor);
    var anchorRow = a.r, startCol = a.c;
    var nCols = spec.types.length;
    var endCol = startCol + nCols - 1;

    var tpl = byNum[anchorRow];
    if (!tpl) throw new Error('Sheet "' + spec.sheetName + '": không có dòng ' + anchorRow +
      ' để lấy mẫu định dạng');

    /* bản đồ style theo cột, lấy từ dòng mẫu */
    function cellsOf(row) {
      var out = {};
      if (!row) return out;
      var cm;
      CELL_RE.lastIndex = 0;
      while ((cm = CELL_RE.exec(row.inner)) !== null) {
        var craw = cm[0];
        var ct = craw.slice(0, craw.indexOf('>') + 1);
        var ref = getAttr(ct, 'r');
        if (!ref) continue;
        out[parseAddr(ref).c] = { raw: craw, s: getAttr(ct, 's') };
      }
      return out;
    }
    var tplCells = cellsOf(tpl);

    var lastExisting = rows.length ? rows[rows.length - 1].num : anchorRow;
    var dataCount = spec.rows.length;
    var lastData = anchorRow + dataCount - 1;
    var lastOut = Math.max(lastData, lastExisting);

    var out = [];
    rows.forEach(function (r) { if (r.num < anchorRow) out.push(r.raw); });

    var cleared = 0, created = 0;
    for (var r = anchorRow; r <= lastOut; r++) {
      var idx = r - anchorRow;
      var orig = byNum[r];
      var hasData = idx < dataCount;

      /* Dòng trống sẵn, nằm ngoài vùng dữ liệu mới -> giữ nguyên, khỏi dựng lại */
      if (!hasData && orig && orig.inner.indexOf('<v>') === -1 && orig.inner.indexOf('<f') === -1 && orig.inner.indexOf('<is>') === -1) {
        out.push(orig.raw);
        continue;
      }
      if (!hasData && !orig) continue;
      if (!hasData) cleared++;
      if (!orig) created++;

      var oc = cellsOf(orig);
      var cols = {};
      Object.keys(oc).forEach(function (k) { cols[k] = true; });
      for (var c = startCol; c <= endCol; c++) cols[c] = true;
      var list = Object.keys(cols).map(Number).sort(function (x, y) { return x - y; });

      var inner = list.map(function (c) {
        if (c < startCol || c > endCol) {
          /* ngoài vùng khối -> giữ nguyên ô cũ */
          return oc[c] ? oc[c].raw : '';
        }
        var bi = c - startCol;
        var style = oc[c] ? oc[c].s : (tplCells[c] ? tplCells[c].s : null);
        var value = hasData ? spec.rows[idx][bi] : '';
        var f = null;
        if (hasData && spec.formulas && spec.formulas[bi]) {
          f = spec.formulas[bi](r, value, idx);
        }
        return emitCell(numToCol(c) + r, style, value, spec.types[bi], f);
      }).join('');

      var ot2 = orig ? orig.openTag : tpl.openTag.replace(/\sr="\d+"/, ' r="' + r + '"');
      out.push(ot2.slice(-2) === '/>' ? ot2.slice(0, -2) + '>' + inner + '</row>'
                                      : ot2 + inner + '</row>');
    }

    var head = xml.slice(0, headEnd);
    if (selfClosing) head = head.slice(0, head.lastIndexOf('/>')) + '>';
    var tail = xml.slice(tailStart);

    /* dimension bao hết dòng đang có; autoFilter bám đúng vùng dữ liệu */
    head = head.replace(/<dimension ref="([^"]*)"\/>/, function (mm, ref) {
      return '<dimension ref="' + setRangeEndRow(ref, lastOut) + '"/>';
    });
    tail = tail.replace(/<autoFilter ref="([^"]*)"/, function (mm, ref) {
      return '<autoFilter ref="' + setRangeEndRow(ref, Math.max(lastData, anchorRow)) + '"';
    });

    /* Bỏ hyperlink cũ: Excel tự tạo link mailto theo địa chỉ ô, sau khi thay dữ
       liệu thì chúng trỏ sang email của dòng khác. Phần .rels để nguyên (quan hệ
       không được tham chiếu là hợp lệ). */
    var hl = tail.match(/<hyperlinks>[\s\S]*?<\/hyperlinks>/);
    if (hl) {
      tail = tail.replace(hl[0], '');
      report.push('  ' + spec.sheetName + ': bỏ ' +
        (hl[0].match(/<hyperlink /g) || []).length + ' hyperlink cũ đã trỏ sai dòng');
    }

    report.push('  ' + spec.sheetName + ': ghi ' + dataCount + ' dòng (' + spec.anchor +
      ' → ' + numToCol(endCol) + lastData + ')' +
      (cleared ? ', xoá dữ liệu cũ ở ' + cleared + ' dòng' : '') +
      (created ? ', thêm mới ' + created + ' dòng' : ''));

    return head + out.join('') + '</sheetData>' + tail;
  }

  /* ---------- vá từng ô lẻ (dùng cho cột ngày ở sheet tổng hợp) ---------- */

  function rowInner(xml, rowNum) {
    /* dòng rỗng được ghi dạng tự đóng <row r="7"/> nên phải bắt cả hai kiểu */
    var re = new RegExp('<row\\b[^>]*\\sr="' + rowNum + '"(?:[^>]*\\/>|[^>]*>[\\s\\S]*?<\\/row>)');
    var m = xml.match(re);
    if (!m) return null;
    var raw = m[0];
    var inner = raw.slice(-2) === '/>' ? ''
      : raw.slice(raw.indexOf('>') + 1, raw.lastIndexOf('</row>'));
    return { raw: raw, inner: inner };
  }

  /** Chèn một <row> trống vào đúng vị trí thứ tự trong <sheetData>. */
  function insertRow(xml, rowNum) {
    var open = xml.indexOf('<sheetData');
    if (open === -1) return null;
    var gt = xml.indexOf('>', open);
    if (xml[gt - 1] === '/') {                       /* <sheetData/> rỗng */
      return xml.slice(0, gt - 1) + '><row r="' + rowNum + '"/></sheetData>' + xml.slice(gt + 1);
    }
    var body = xml.slice(gt + 1, xml.lastIndexOf('</sheetData>'));
    var at = -1, m;
    ROW_RE.lastIndex = 0;
    while ((m = ROW_RE.exec(body)) !== null) {
      var n = parseInt(getAttr(m[0].slice(0, m[0].indexOf('>') + 1), 'r'), 10);
      if (n > rowNum) { at = gt + 1 + m.index; break; }
    }
    if (at === -1) at = xml.lastIndexOf('</sheetData>');
    return xml.slice(0, at) + '<row r="' + rowNum + '"/>' + xml.slice(at);
  }

  /** Tìm cột có giá trị số bằng `serial` trên dòng header ngày. */
  function findDayColumn(xml, rowNum, serial) {
    var row = rowInner(xml, rowNum);
    if (!row) return -1;
    var cm;
    CELL_RE.lastIndex = 0;
    while ((cm = CELL_RE.exec(row.inner)) !== null) {
      var craw = cm[0];
      var ct = craw.slice(0, craw.indexOf('>') + 1);
      var t = getAttr(ct, 't');
      if (t === 's' || t === 'inlineStr' || t === 'str') continue;
      var vm = craw.match(/<v>([^<]*)<\/v>/);
      if (vm && Math.round(parseFloat(vm[1])) === serial) {
        return parseAddr(getAttr(ct, 'r')).c;
      }
    }
    return -1;
  }

  /**
   * edits: [{ row, col, value, type, formula, style }] — giữ nguyên style ô cũ;
   * `style` chỉ dùng khi ô chưa tồn tại. Dòng chưa tồn tại thì bỏ qua và đếm lại.
   */
  function patchCells(xml, edits, opts) {
    opts = opts || {};
    var byRow = {}, skipped = 0, added = 0, maxRow = 0;
    edits.forEach(function (e) { (byRow[e.row] = byRow[e.row] || []).push(e); });

    Object.keys(byRow).map(Number).sort(function (a, b) { return a - b; }).forEach(function (rowNum) {
      var rk = String(rowNum);
      var row = rowInner(xml, rowNum);
      if (!row && opts.createMissing) {
        var ins = insertRow(xml, rowNum);
        if (ins) { xml = ins; row = rowInner(xml, rowNum); added++; }
      }
      if (!row) { skipped += byRow[rk].length; return; }
      if (rowNum > maxRow) maxRow = rowNum;

      var cells = [], cm;
      CELL_RE.lastIndex = 0;
      while ((cm = CELL_RE.exec(row.inner)) !== null) {
        var craw = cm[0];
        var ct = craw.slice(0, craw.indexOf('>') + 1);
        cells.push({ col: parseAddr(getAttr(ct, 'r')).c, raw: craw, s: getAttr(ct, 's') });
      }

      byRow[rk].forEach(function (e) {
        var found = null;
        for (var i = 0; i < cells.length; i++) { if (cells[i].col === e.col) { found = cells[i]; break; } }
        var style = found ? found.s : (e.style !== undefined ? e.style : null);
        var xmlCell = emitCell(numToCol(e.col) + rowNum, style, e.value, e.type, e.formula);
        if (found) found.raw = xmlCell;
        else cells.push({ col: e.col, raw: xmlCell, s: null });
      });

      cells.sort(function (a, b) { return a.col - b.col; });
      var open = row.raw.slice(0, row.raw.indexOf('>') + 1);
      if (open.slice(-2) === '/>') open = open.slice(0, -2) + '>';
      xml = xml.replace(row.raw, function () {
        return open + cells.map(function (c) { return c.raw; }).join('') + '</row>';
      });
    });
    if (added) {
      /* dimension phải bao hết dòng vừa thêm, nếu không Excel coi là file hỏng */
      xml = xml.replace(/<dimension ref="([^"]*)"\/>/, function (mm, ref) {
        var m2 = ref.match(/^(\$?[A-Z]+\$?\d+):(\$?[A-Z]+)\$?(\d+)$/);
        if (!m2 || parseInt(m2[3], 10) >= maxRow) return mm;
        return '<dimension ref="' + m2[1] + ':' + m2[2] + maxRow + '"/>';
      });
    }
    if (skipped && opts.report) {
      opts.report.push('  (bỏ qua ' + skipped + ' ô vì dòng chưa tồn tại trong sheet)');
    }
    return xml;
  }

  /** Điền các chỉ tiêu tính được vào cột ngày của sheet Summarize_team VietNam. */
  function patchSummary(xml, summary, cfg, reportSerial, report) {
    var sc = cfg.summary;
    if (!sc || !summary) return xml;

    var col = findDayColumn(xml, sc.dayHeaderRow, reportSerial);
    if (col === -1) {
      report.push('  ' + sc.sheet + ': KHÔNG tìm thấy cột cho ngày ' + reportSerial +
        ' (dòng ' + sc.dayHeaderRow + ') — bỏ qua phần tổng hợp');
      return xml;
    }
    var colName = numToCol(col);

    var edits = [], filled = [];
    (sc.cells || []).forEach(function (spec) {
      if (spec.from === 'customerFormula') {
        var c = summary.counts;
        var cached = summary.distinct + (c.zalo || 0) + (c.viber || 0) + (c.line || 0) + (c.tel || 0);
        var f = (sc.customerFormula || '{distinct}+SUM({col}13:{col}16)')
          .replace(/\{distinct\}/g, summary.distinct)
          .replace(/\{col\}/g, colName);
        edits.push({ row: spec.row, col: col, value: cached, type: 'number', formula: f });
        filled.push(spec.label + '=' + cached + ' (=' + f + ')');
      } else {
        var v = summary.counts[spec.from];
        if (v === undefined) return;
        edits.push({ row: spec.row, col: col, value: v, type: 'number', formula: null });
        filled.push(spec.label + '=' + v);
      }
    });

    /* Khối "Nội dung hỗ trợ" — đếm theo Topic Level 1 */
    var topicFilled = [];
    (sc.topicRows || []).forEach(function (spec) {
      var v = (summary.topics || {})[String(spec.topic)] || 0;
      edits.push({ row: spec.row, col: col, value: v, type: 'number', formula: null });
      topicFilled.push(spec.label + '=' + v);
    });

    xml = patchCells(xml, edits);
    report.push('  ' + sc.sheet + ': điền cột ' + colName + ' (ngày ' + reportSerial + ') — ' +
      filled.join(', '));
    if (topicFilled.length) {
      report.push('  ' + sc.sheet + ': khối "Nội dung hỗ trợ" (dòng ' +
        sc.topicRows[0].row + '-' + sc.topicRows[sc.topicRows.length - 1].row + ') — ' +
        topicFilled.join(', '));
    }
    return xml;
  }

  /** Đọc cột email của một sheet trong workbook đã parse -> {emailChuẩnHoá: dòng}. */
  function emailRows(wb, sheetName, emailCol, r0, r1) {
    var out = {};
    if (!wb) return out;
    var ws = wb.Sheets[sheetName];
    if (!ws) return out;
    var m = P.sheetToMatrix(ws);
    for (var r = r0 - 1; r < Math.min(m.rows.length, r1); r++) {
      var v = P.cellToString((m.rows[r] || [])[emailCol]).trim();
      if (v.indexOf('@') !== -1) out[P.normText(v)] = r + 1;
    }
    return out;
  }

  /**
   * Ghi hiệu suất công việc vào sheet KPI tháng, và cập nhật ô ngày của cả 2 sheet KPI.
   * Chỉ ghi cho người có trong danh sách OP của "Daily KPI Result" — mấy dòng
   * JP Staff / tài khoản khác trong sheet tháng do nguồn khác điền, không đụng vào.
   */
  /**
   * Dòng nào trong "Daily KPI Result" đang dùng định mức KHÁC với `kpi.tracks`?
   *
   * Mỗi cột hiệu suất mảng có dạng `=IFERROR(N54/(O54*10),)` — con số cuối là định
   * mức. Tool tính theo định mức chuẩn (15/30/17/20/20), nên với dòng dùng định mức
   * riêng thì kết quả sẽ sai. Thà để trống cho người nhập tay còn hơn ghi số sai.
   *
   * Trả { dòng: 'mô tả lệch' } cho những dòng lệch chuẩn.
   */
  function oddRateRows(files, sheetPart, cfg, readText) {
    var kc = cfg.kpi, dc = kc.daily;
    var out = {};
    if (!dc.rateColumns || !dc.rateColumns.length) return out;
    var part = sheetPart[dc.sheet];
    if (!part || !files[part]) return out;

    var xml = readText(part);
    var masters = sharedMasters(xml);
    var want = kc.tracks.map(function (t) { return String(t.rate); });

    for (var r = dc.dataStartRow; r <= (dc.peopleEndRow || dc.dataEndRow); r++) {
      var row = rowInner(xml, r);
      if (!row) continue;
      var byCol = {}, cm;
      CELL_RE.lastIndex = 0;
      while ((cm = CELL_RE.exec(row.inner)) !== null) {
        var craw = cm[0];
        var ref = getAttr(craw.slice(0, craw.indexOf('>') + 1), 'r');
        if (ref) byCol[parseAddr(ref).c] = craw;
      }
      var diff = [];
      dc.rateColumns.forEach(function (ci, i) {
        var f = byCol[ci] ? realFormula(byCol[ci], r, masters) : null;
        if (!f) return;                               /* ô trống -> không kết luận */
        var m = f.match(/\*\s*(\d+(?:\.\d+)?)\s*\)/);
        if (!m || m[1] === want[i]) return;
        diff.push(numToCol(ci) + ': ' + m[1] + ' thay vì ' + want[i]);
      });
      if (diff.length) out[r] = diff.join(', ');
    }
    return out;
  }

  /* Các tổng số lượng/giờ phải cùng tập dòng để tính KPI nhóm bằng team(emails).
     Chỉ nhận SUM các ô/vùng cùng cột trong sheet hiện tại; không suy đoán công thức lạ. */
  function teamRowsFromFormulas(formulas, dc) {
    var columns = ['E', 'F', 'I', 'K', 'L', 'N', 'O', 'Q', 'R', 'T', 'U'];
    var chosen = null;
    columns.forEach(function (col) {
      var f = String(formulas[col] || '').replace(/\s+/g, '').replace(/^=/, '').toUpperCase();
      var sum = f.match(/^SUM\((.+)\)$/);
      function invalid() {
        throw new Error('KPI nhóm: không xác định được phạm vi SUM tại ' + dc.sheet + '!' +
          col + dc.totalRow + '. Kiểm tra công thức tổng và giới hạn peopleEndRow; tool không tự dùng phạm vi cũ.');
      }
      if (!sum) invalid();
      var seen = {};
      sum[1].split(/[,;]/).forEach(function (term) {
        var m = term.match(/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/);
        if (!m || m[1] !== col || (m[3] && m[3] !== col)) invalid();
        var first = +m[2], last = +(m[4] || m[2]);
        if (first < dc.dataStartRow || last < first || last > (dc.peopleEndRow || dc.dataEndRow)) invalid();
        for (var r = first; r <= last; r++) {
          if (r === dc.totalRow || seen[r]) invalid();
          seen[r] = true;
        }
      });
      var rows = Object.keys(seen).map(Number).sort(function (a, b) { return a - b; });
      if (chosen && rows.join(',') !== chosen.join(',')) {
        throw new Error('KPI nhóm: các công thức SUM số lượng và giờ không cùng phạm vi (' +
          col + dc.totalRow + '). Cần thống nhất công thức trong file nền.');
      }
      chosen = rows;
    });
    return chosen;
  }

  /* Số lượng không tìm thấy = 0, không phải chuỗi "" làm E-F / F*1.5 bị #VALUE.
     Sửa nội dung công thức tại chỗ để giữ nguyên nhóm shared và định dạng ô. */
  function normalizeKpiCountLookups(xml, dc, report) {
    var changed = 0;
    var result = xml.replace(CELL_RE, function (raw) {
      var addr = getAttr(raw, 'r');
      if (!addr) return raw;
      var at = parseAddr(addr);
      if (at.r < dc.dataStartRow || at.r > (dc.peopleEndRow || dc.dataEndRow) ||
          [4, 5, 10, 13, 16, 19].indexOf(at.c) === -1) return raw;
      return raw.replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/, function (tag, attrs, content) {
        var f = unesc(content);
        var m = f.match(/^IFERROR\((VLOOKUP\(\$?D\d+,(?:'Sheet5'|Sheet5)![^,]+,\d+,0\)),\s*""\)$/i);
        if (!m) return tag;
        changed++;
        return '<f' + attrs + '>' + esc('IFERROR(' + m[1] + ',0)') + '</f>';
      });
    });
    if (changed) report.push('  ' + dc.sheet + ': ' + changed + ' công thức tra số lượng dùng 0 khi không tìm thấy, không làm mất phần chat khi chưa có case.');
    return result;
  }

  function patchKpi(files, sheetPart, readText, writeText, blocks, cfg, masterWb, report, info) {
    var kc = cfg.kpi;
    if (!kc || !blocks.kpi) return;
    var serial = blocks.reportSerial;

    /* ô ngày của Daily KPI Result do patchDateHeaders lo */

    /* Đọc công thức thật từ XML (UI có thể parse workbook không giữ cell.f).
       Tổng OP theo mẫu; KPI cá nhân vẫn ghi cho cả người nằm ngoài tổng OP. */
    var dailyPart = sheetPart[kc.daily.sheet];
    if (!dailyPart || !masterWb || !masterWb.Sheets[kc.daily.sheet]) {
      throw new Error('KPI: thiếu sheet ' + kc.daily.sheet + ' trong file nền');
    }
    var dailyXml = readText(dailyPart);
    var dailyMasters = sharedMasters(dailyXml);
    var total = rowInner(dailyXml, kc.daily.totalRow), totalFormulas = {};
    ((total && total.inner.match(CELL_RE)) || []).forEach(function (cell) {
      var addr = getAttr(cell, 'r');
      if (addr) totalFormulas[numToCol(parseAddr(addr).c)] = realFormula(cell, kc.daily.totalRow, dailyMasters);
    });
    var teamRows = teamRowsFromFormulas(totalFormulas, kc.daily), opEmails = [], seenEmails = {};
    var dailyWs = masterWb.Sheets[kc.daily.sheet];
    teamRows.forEach(function (r) {
      var cell = dailyWs[numToCol(kc.daily.emailColumn) + r];
      var email = P.normText(cell && cell.v);
      if (email.indexOf('@') === -1) return;
      if (seenEmails[email]) throw new Error('KPI nhóm: email bị trùng tại dòng ' + seenEmails[email] + ' và ' + r + ': ' + email);
      seenEmails[email] = r;
      opEmails.push(email);
    });
    if (info) { info.teamRows = teamRows; info.teamEmails = opEmails; }
    report.push('  KPI: nhóm OP theo công thức SUM của file nền, dòng ' + teamRows.join(', ') +
      ' (' + opEmails.length + ' email); không dùng dataEndRow cố định.');
    var allRows = emailRows(masterWb, kc.daily.sheet, kc.daily.emailColumn,
      kc.daily.dataStartRow, kc.daily.peopleEndRow || kc.daily.dataEndRow);

    /* Bỏ những người dùng định mức KPI riêng — tính theo định mức chuẩn sẽ ra số sai */
    var odd = oddRateRows(files, sheetPart, cfg, readText);
    /* Đọc định mức và cột số lượng của JP từ chính mẫu; không áp định mức OP.
       Chỉ nhận mẫu công thức đã hiểu, mẫu khác vẫn giữ cảnh báo nhập tay. */
    function profileFor(row) {
      var raw = rowInner(dailyXml, row);
      if (!raw) return null;
      var cells = {}, matches = raw.inner.match(CELL_RE) || [];
      matches.forEach(function (cell) { cells[getAttr(cell, 'r')] = cell; });
      function formula(col) { return realFormula(cells[col + row] || '', row, dailyMasters) || ''; }
      var profile = { counts: ['H'], rates: [] };
      var countKeys = {16:'OB',17:'case',18:'CMT',19:'DKM',20:'XL',21:'JP',22:'SETS'};
      var countCols = ['H','K','N','Q','T'];
      var rateCols = ['J','M','P','S','V'];
      for (var i = 0; i < 5; i++) {
        var rate = formula(rateCols[i]).match(/\*\s*(\d+(?:\.\d+)?)\s*\)/);
        if (!rate || +rate[1] <= 0) return null;
        profile.rates.push(+rate[1]);
        if (!i) continue;
        var source = formula(countCols[i]).match(/VLOOKUP\(\$D\d+,Sheet5![^,]+,(\d+),0\)/i);
        if (!source || !countKeys[source[1]]) return null;
        profile.counts.push(countKeys[source[1]]);
      }
      return profile;
    }
    var profiles = {};
    var oddEmails = {};
    Object.keys(allRows).forEach(function (e) {
      if (odd[allRows[e]]) {
        var profile = profileFor(allRows[e]);
        if (profile) profiles[e] = profile;
        else oddEmails[e] = odd[allRows[e]];
      }
    });
    if (info) info.oddRateEmails = oddEmails;
    writeText(dailyPart, normalizeKpiCountLookups(dailyXml, kc.daily, report));
    if (!opEmails.length) {
      report.push('  KPI: không đọc được danh sách OP từ "' + kc.daily.sheet + '" — bỏ qua sheet KPI tháng');
      return;
    }

    /* --- sheet KPI tháng, dò theo mẫu tên vì tên đổi theo tháng --- */
    var re = new RegExp(kc.monthly.sheetPattern, 'i');
    var mName = Object.keys(sheetPart).filter(function (n) { return re.test(n); })[0];
    if (!mName) {
      report.push('  KPI: không tìm thấy sheet khớp mẫu /' + kc.monthly.sheetPattern + '/ — bỏ qua');
      return;
    }
    var mPart = sheetPart[mName];
    var xml = readText(mPart);

    var col = findDayColumn(xml, kc.monthly.dayHeaderRow, serial);
    if (col === -1) {
      report.push('  ' + mName + ': không có cột cho ngày ' + serial +
        ' (dòng ' + kc.monthly.dayHeaderRow + ') — bỏ qua');
      return;
    }

    var rowsOf = emailRows(masterWb, mName, kc.monthly.emailColumn,
      kc.monthly.dataStartRow, kc.monthly.dataEndRow);
    var edits = [], written = 0, skipped = [], noHours = [], oddSkipped = [];
    Object.keys(rowsOf).forEach(function (e) {
      if (!allRows[e]) { skipped.push(e); return; }
      if (oddEmails[e]) { oddSkipped.push(e + ' (' + oddEmails[e] + ')'); return; }
      var k = blocks.kpi.forEmail(e, profiles[e]);
      /* W = 0 mà người này vẫn có việc trong ngày nghĩa là thiếu giờ làm ở
         "3. Other Tasks" — phải nói ra, nếu không số 0 đó trông như thật. */
      if (k && !k.W && k.hasWork && !k.hours) noHours.push(e);
      edits.push({ row: rowsOf[e], col: col, value: k.W, type: 'number', formula: null });
      written++;
    });

    /* dòng tổng team = tính từ SUM của cả nhóm OP, không phải trung bình các W */
    if (kc.monthly.teamRow) {
      edits.push({ row: kc.monthly.teamRow, col: col,
        value: blocks.kpi.team(opEmails), type: 'number', formula: null });
    }
    if (kc.monthly.dateCell) {
      var b = parseAddr(kc.monthly.dateCell);
      edits.push({ row: b.r, col: b.c, value: serial, type: 'date', formula: null });
    }

    writeText(mPart, patchCells(xml, edits));
    report.push('  ' + mName + ': điền cột ' + numToCol(col) + ' (ngày ' + serial + ') — ' +
      written + ' người' + (kc.monthly.teamRow ? ' + dòng tổng' : '') +
      (skipped.length ? ', bỏ qua ' + skipped.length + ' dòng không có trong "' +
        kc.daily.sheet + '" (' + skipped.join(', ') + ')' : ''));
    /* Hướng NGUY HIỂM khi có nhân viên mới: đã thêm vào "Daily KPI Result" nhưng
       quên thêm dòng ở sheet KPI tháng -> tool không có ô nào để ghi, KPI người đó
       lặng lẽ không bao giờ xuất hiện. Phải báo. */
    var noMonthlyRow = Object.keys(allRows).filter(function (e) {
      return !rowsOf[e] && !oddEmails[e];
    });
    if (noMonthlyRow.length) {
      report.push('  ⚠ ' + mName + ': ' + noMonthlyRow.length + ' người có trong "' +
        kc.daily.sheet + '" nhưng CHƯA có dòng ở sheet KPI tháng — KPI của họ không được ' +
        'ghi vào đâu cả. Thêm dòng cho họ trong "' + mName + '" (cột ' +
        numToCol(kc.monthly.emailColumn) + ' là email) rồi chạy lại: ' + noMonthlyRow.join(', '));
    }

    if (oddSkipped.length) {
      report.push('  ⚠ ' + mName + ': để TRỐNG ' + oddSkipped.length + ' người dùng định mức ' +
        'KPI riêng trong "' + kc.daily.sheet + '" — tính theo định mức chuẩn sẽ ra số sai, ' +
        'nên phải nhập tay: ' + oddSkipped.join(' · '));
    }
    if (noHours.length) {
      report.push('  ⚠ ' + mName + ': ' + noHours.length + ' người có case/phiên chat hôm nay' +
        ' nhưng KHÔNG có khối giờ trong "3. Other Tasks", nên hiệu suất buộc phải ra 0 ' +
        '(công thức là tổng-việc chia tổng-giờ). Soát lại xem file OP có thiếu khối giờ ' +
        'của họ không — nếu họ thật sự không khai giờ thì 0 là đúng. Danh sách: ' +
        noHours.join(', '));
    }
  }

  /** Dịch cột trong công thức sang phải `delta` cột: R11 -> S11, nhưng $R11 giữ nguyên.
   *  Dấu `!` PHẢI nằm trong tập ký tự đứng trước hợp lệ, vì tham chiếu liên sheet
   *  luôn có dạng `'Tên sheet'!R11` — loại `!` ra là không dịch được gì cả. */
  function shiftFormulaCols(f, delta) {
    return f.replace(/(^|[^A-Za-z0-9_$])(\$?)([A-Z]{1,3})(\$?)(\d+)/g,
      function (m, pre, d1, col, d2, num) {
        if (d1) return m;                      /* $R = tuyệt đối, giữ nguyên */
        var n = colToNum(col) + delta;
        if (n < 0 || n > 16383) return m;
        return pre + numToCol(n) + d2 + num;
      });
  }

  /**
   * Sheet "Summarize" chỉ là bản rút gọn của "Summarize_team VietNam": dòng 7-27
   * toàn là công thức gương kiểu `='Summarize_team VietNam'!S11`. Nhưng công thức
   * chỉ có ĐẾN CỘT CỦA NGÀY CUỐI CÙNG — cột của ngày mới hoàn toàn trống, mỗi ngày
   * người làm phải kéo sang phải một cột.
   *
   * Không kéo thì `Summarize` trắng gần hết dù `Summarize_team` đã đủ số. Tool
   * chép công thức từ cột liền trước rồi dịch tham chiếu sang phải một cột — đúng
   * bằng thao tác kéo tay, và không đụng vào ô nào đã có công thức sẵn.
   */
  function patchSummarizeMirror(files, sheetPart, readText, writeText, cfg, serial, report) {
    var mc = cfg.summarize;
    if (!mc) return;
    var part = sheetPart[mc.sheet];
    if (!part || !files[part]) { report.push('  ' + mc.sheet + ': không có sheet — bỏ qua'); return; }

    var xml = readText(part);
    var col = findDayColumn(xml, mc.dayHeaderRow, serial);
    if (col === -1) {
      report.push('  ' + mc.sheet + ': không tìm thấy cột cho ngày ' + serial +
        ' (dòng ' + mc.dayHeaderRow + ') — bỏ qua');
      return;
    }
    if (col <= 0) return;

    var masters = sharedMasters(xml);
    function cellsOfRow(r) {
      var row = rowInner(xml, r);
      if (!row) return null;
      var out = {}, cm;
      CELL_RE.lastIndex = 0;
      while ((cm = CELL_RE.exec(row.inner)) !== null) {
        var craw = cm[0];
        var ct = craw.slice(0, craw.indexOf('>') + 1);
        var ref = getAttr(ct, 'r');
        if (ref) out[parseAddr(ref).c] = { raw: craw, s: getAttr(ct, 's') };
      }
      return out;
    }

    var edits = [], done = [], already = 0;
    for (var r = mc.firstRow; r <= mc.lastRow; r++) {
      var cs = cellsOfRow(r);
      if (!cs) continue;
      if (cs[col] && realFormula(cs[col].raw, r, masters)) { already++; continue; }
      var src = cs[col - 1];
      if (!src) continue;
      var f = realFormula(src.raw, r, masters);
      if (!f) continue;
      edits.push({ row: r, col: col, value: '', type: 'auto',
                   formula: shiftFormulaCols(f, 1), style: src.s });
      done.push(r);
    }
    if (!edits.length) {
      report.push('  ' + mc.sheet + ': cột ngày ' + numToCol(col) + ' đã có đủ công thức');
      return;
    }
    writeText(part, patchCells(xml, edits, { report: report }));
    report.push('  ' + mc.sheet + ': kéo công thức sang cột ' + numToCol(col) +
      ' (ngày ' + serial + ') cho ' + done.length + ' dòng — ' +
      (already ? already + ' dòng đã có sẵn. ' : '') +
      'Sheet này chỉ là gương của "' + cfg.summary.sheet + '", không có dữ liệu riêng.');
  }

  /**
   * 6 dòng fanpage (61-66) của sheet tổng hợp.
   *
   * FOLLOW: số thật do extension lấy về, người dùng dán vào tool.
   * LIKE:   phía SBI chỉ cần ước lượng theo các ngày trước. Đối chiếu dữ liệu
   *         thật thì 3 giá trị này giữ nguyên cả tháng, nên chép lại giá trị
   *         gần nhất bên trái đúng là việc người làm vẫn đang làm.
   *
   * Đọc giá trị cũ từ `masterWb` chứ không từ XML: Like là chữ ("382k") nên
   * trong XML nó là chỉ số vào bảng chuỗi dùng chung, tự giải rất phiền.
   */
  function patchFanpage(files, sheetPart, readText, writeText, cfg, serial, fanpage, masterWb, report) {
    var fc = cfg.fanpage;
    if (!fc || !fc.pages || !fc.pages.length) return [];
    var sheetName = (cfg.summary && cfg.summary.sheet) || fc.sheet;
    var part = sheetPart[sheetName];
    if (!part || !files[part]) return [];

    var xml = readText(part);
    var headerRow = (cfg.summary && cfg.summary.dayHeaderRow) || 4;
    var col = findDayColumn(xml, headerRow, serial);
    if (col === -1) {
      report.push('  fanpage: không tìm thấy cột ngày ' + serial + ' — bỏ qua');
      return [];
    }

    /* Ma trận của sheet để đọc giá trị các ngày trước */
    var m = null, hdr = [];
    if (masterWb && masterWb.Sheets[sheetName]) {
      m = P.sheetToMatrix(masterWb.Sheets[sheetName]);
      hdr = m.rows[headerRow - 1] || [];
    }
    function prevValue(row, fromCol) {
      if (!m) return null;
      var lookback = fc.carryLookback || 40;
      var r = m.rows[row - 1] || [];
      for (var c = fromCol - 1; c >= 0 && fromCol - c <= lookback; c--) {
        /* chỉ lùi qua các cột thực sự là cột ngày */
        var h = parseFloat(String(hdr[c] === undefined ? '' : hdr[c]).trim());
        if (!isFinite(h)) continue;
        var v = P.cellToString(r[c]).trim();
        if (v) return { value: v, col: c, serial: Math.round(h) };
      }
      return null;
    }

    var edits = [], wrote = [], notes = [], oddLikes = [];
    fc.pages.forEach(function (pg) {
      /* --- Follow --- */
      var n = fanpage && fanpage.follows ? fanpage.follows[pg.key] : undefined;
      if (n !== undefined && n !== null) {
        edits.push({ row: pg.followRow, col: col, value: n, type: 'number', formula: null });
        wrote.push(pg.followRow);
        notes.push(pg.label + ' follow=' + n);
      }

      /* --- Like: chép của ngày gần nhất --- */
      if (fc.carryLike) {
        var prev = prevValue(pg.likeRow, col);
        if (prev) {
          edits.push({ row: pg.likeRow, col: col, value: prev.value, type: 'auto', formula: null });
          wrote.push(pg.likeRow);
          notes.push(pg.label + ' like=' + prev.value + ' (chép ngày ' + prev.serial + ')');
          /* Giá trị này lẽ ra đứng yên cả tháng. Vừa đổi = đáng ngờ. */
          var before = prevValue(pg.likeRow, prev.col);
          if (before && P.normText(before.value) !== P.normText(prev.value)) {
            oddLikes.push(pg.label + ': ngày ' + before.serial + ' là "' + before.value +
              '" nhưng ngày ' + prev.serial + ' đổi thành "' + prev.value + '"');
          }
        }
      }
    });

    if (!edits.length) return [];
    writeText(part, patchCells(xml, edits, { report: report }));
    report.push('  ' + sheetName + ': 6 dòng fanpage, cột ' + numToCol(col) + ' — ' + notes.join(', '));
    if (oddLikes.length) {
      report.push('  ⚠ fanpage: số Like lẽ ra giữ nguyên cả tháng nhưng vừa đổi — ' +
        oddLikes.join(' · ') + '. Kiểm tra xem có gõ nhầm không trước khi chép tiếp.');
    }
    return wrote;
  }

  /** Ô ngày ở phần đầu mỗi tab — nằm ngoài vùng khối nên phải ghi riêng. */
  function patchDateHeaders(files, sheetPart, readText, writeText, cfg, serial, report) {
    var list = cfg.dateHeaderCells || [];
    var done = [], missing = [];
    list.forEach(function (spec) {
      var part = sheetPart[spec.sheet];
      if (!part || !files[part]) { missing.push(spec.sheet); return; }
      var a = parseAddr(spec.cell);
      writeText(part, patchCells(readText(part),
        [{ row: a.r, col: a.c, value: serial, type: 'date', formula: null }]));
      done.push(spec.sheet + '!' + spec.cell);
    });
    if (done.length) {
      report.push('  ô ngày đầu tab: cập nhật ' + done.length + ' ô = ' + serial +
        ' (' + done.join(', ') + ')');
    }
    if (missing.length) {
      report.push('  ô ngày đầu tab: không tìm thấy sheet ' + missing.join(', '));
    }
  }

  /**
   * Kéo công thức cột phụ của Sheet5 xuống cho dư chỗ.
   * Pivot nở ra khi có thêm supporter; cột phụ không tự nở theo, nên người mới
   * sẽ không có giá trị HTKH và "Daily KPI Result" tính sai cho họ.
   */
  /** Đổi số dòng trong công thức: A5 -> A12, nhưng $A$5 (tuyệt đối) giữ nguyên. */
  function rebaseFormula(f, from, to) {
    return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, function (m, d1, col, d2, num) {
      if (d2) return m;
      if (parseInt(num, 10) !== from) return m;
      return d1 + col + to;
    });
  }

  /**
   * Gom công thức chia sẻ: <f t="shared" ref="O6:O33" si="1">…</f> chỉ ghi nội dung
   * ở ô chủ, các ô còn lại là <f t="shared" si="1"/> rỗng. Muốn đọc được công thức
   * của một ô bất kỳ phải tra ngược về ô chủ rồi đổi số dòng.
   */
  function sharedMasters(xml) {
    var out = {}, rm;
    ROW_RE.lastIndex = 0;
    while ((rm = ROW_RE.exec(xml)) !== null) {
      var raw = rm[0];
      var rn = parseInt(getAttr(raw.slice(0, raw.indexOf('>') + 1), 'r'), 10);
      var cm;
      CELL_RE.lastIndex = 0;
      while ((cm = CELL_RE.exec(raw)) !== null) {
        var fm = cm[0].match(/<f\b([^>]*)>([\s\S]+?)<\/f>/);
        if (!fm || fm[1].indexOf('t="shared"') === -1) continue;
        var si = getAttr('<f' + fm[1] + '>', 'si');
        if (si !== null && !out[si]) out[si] = { f: unesc(fm[2]), row: rn };
      }
    }
    return out;
  }

  /** Công thức thực của một ô, đã giải shared và đổi về đúng dòng của nó. */
  function realFormula(cellXml, rowNum, masters) {
    var fm = cellXml.match(/<f\b([^>]*)>([\s\S]*?)<\/f>|<f\b([^>]*)\/>/);
    if (!fm) return null;
    var at = fm[1] !== undefined ? fm[1] : fm[3];
    var txt = fm[2] || '';
    if (txt) return unesc(txt);
    var si = getAttr('<f' + at + '>', 'si');
    if (si === null || !masters[si]) return null;
    return rebaseFormula(masters[si].f, masters[si].row, rowNum);
  }

  function patchSheet5(files, sheetPart, readText, writeText, cfg, report) {
    var sc = cfg.sheet5;
    if (!sc) return;
    var part = sheetPart[sc.sheet];
    if (!part || !files[part]) { report.push('  ' + sc.sheet + ': không có sheet — bỏ qua'); return; }

    var xml = readText(part);
    /* Pivot đổi thứ tự cột khi có kênh mới. Tra theo nhãn đã trim thay vì
       sao chép D/G/H của hôm trước (04/10: JP khác ở H, XL DKM ở I).
       B:N là vùng pivot, O:V là vùng phụ, không được tự tham chiếu. */
    var groups = [
      ['Outbound'], ['Xử lý case (BO Team)'], ['Trả lời Comment'],
      ['Đăng ký mới'], ['Xử lý đăng ký mới'], ['JP khác'], ['Setsumekai']
    ];
    function channelFormula(row, labels) {
      return labels.map(function (label) {
        return 'SUMPRODUCT(--(TRIM($B$4:$N$4)="' + label + '"),B' + row + ':N' + row + ')';
      }).join('+');
    }
    function whenPivotRowExists(row, formula) {
      /* Giữ công thức dự phòng cho pivot nở thêm, nhưng không hiện số 0
         ở các dòng trống bên dưới Grand Total. Số 0 của người thật vẫn giữ. */
      return 'IF($A' + row + '="","",' + formula + ')';
    }
    var edits = [];
    for (var rr = sc.templateRow; rr <= sc.extendToRow; rr++) {
      edits.push({ row: rr, col: 14, value: '', type: 'number',
        formula: whenPivotRowExists(rr, channelFormula(rr, ['Grand Total']) + '-SUM(P' + rr + ':U' + rr + ')+V' + rr) });
      groups.forEach(function (labels, i) {
        var formula = channelFormula(rr, labels);
        if (i === 6) {
          var weight = cfg.kpi && typeof cfg.kpi.setsumekaiWeight === 'number' ? cfg.kpi.setsumekaiWeight : 45;
          formula = '(' + formula + ')*' + weight;
        }
        edits.push({ row: rr, col: 15 + i, value: '', type: 'number',
          formula: whenPivotRowExists(rr, formula) });
      });
    }
    writeText(part, patchCells(xml, edits, { createMissing: true, report: report }));
    report.push('  ' + sc.sheet + ': cập nhật O:V dòng ' + sc.templateRow + '-' + sc.extendToRow +
      ' theo tên CHANNEL; tự theo đúng kênh khi pivot đổi thứ tự cột.');
  }

  /**
   * Bật "Include new items in manual filter" cho mọi trường pivot có bộ lọc.
   *
   * VÌ SAO CẦN: pivot đếm CHANNEL trong Sheet5 lấy `Date/日付` làm page field, và
   * trường này đang dùng bộ lọc tay (`multipleItemSelectionAllowed="1"` +
   * `showAll="0"` + danh sách <item> liệt kê từng ngày). Mỗi lần chạy, tool thay
   * toàn bộ "1. CSKH" bằng dữ liệu của ngày MỚI — ngày đó chưa có trong danh sách
   * item, nên khi Excel refresh lúc mở file, nó coi là mục mới và KHÔNG chọn:
   * pivot bị lọc sạch, chỉ còn dòng Grand Total rỗng.
   *
   * Hậu quả dây chuyền: cột phụ Sheet5 = 0 -> VLOOKUP của "Daily KPI Result"
   * không tra được -> toàn bộ sheet KPI ngày ra 0. Bộ test JS không thấy vì nó so
   * số tool tự tính, không đi qua công thức Excel; chỉ mở bằng Excel thật mới lộ.
   *
   * Đã kiểm chứng trên Excel: thêm thuộc tính này thì pivot giữ nguyên A1:J33 và
   * Grand Total = 2741 (đúng số dòng ngày 15/9); không có thì co lại còn A1:B5.
   */
  function fixPivotFilters(files, names, readText, writeText) {
    var PF = /<pivotField\b[^>]*\/>|<pivotField\b[^>]*>[\s\S]*?<\/pivotField>/g;
    var count = 0;
    names.forEach(function (n) {
      if (!/^xl\/pivotTables\/pivotTable\d+\.xml$/.test(n) || !files[n]) return;
      var changed = 0;
      var x = readText(n).replace(PF, function (c) {
        var head = c.slice(0, c.indexOf('>') + 1);
        /* chỉ trường nằm trên một trục mới có bộ lọc; trường không dùng thì bỏ qua */
        if (head.indexOf(' axis="') === -1) return c;
        if (head.indexOf('includeNewItemsInFilter') !== -1) return c;
        changed++;
        var selfClose = head.slice(-2) === '/>';
        var open = (selfClose ? head.slice(0, -2) : head.slice(0, -1)).replace(/\s+$/, '');
        return open + ' includeNewItemsInFilter="1"' + (selfClose ? '/>' : '>') +
               c.slice(head.length);
      });
      if (changed) { writeText(n, x); count += changed; }
    });
    return count;
  }

  /**
   * Gỡ sheet liệt kê lỗi mà `markup.js` chèn vào bản đánh dấu.
   *
   * Người dùng tải bản đánh dấu của File báo cáo tổng, sửa `mail SF` rồi nạp lại
   * làm file nền. Không gỡ thì sheet đó đi thẳng vào báo cáo gửi khách hàng.
   */
  function stripMarkupSheets(files, readText, writeText, report) {
    var wbXml = readText('xl/workbook.xml');
    var relsXml = readText('xl/_rels/workbook.xml.rels');
    var rel = {};
    (relsXml.match(/<Relationship\b[^>]*\/>/g) || []).forEach(function (t) {
      var id = getAttr(t, 'Id'), tgt = getAttr(t, 'Target');
      if (id && tgt) rel[id] = tgt.charAt(0) === '/' ? tgt.slice(1) : 'xl/' + tgt.replace(/^\.\//, '');
    });

    var dropped = [];
    (wbXml.match(/<sheet\b[^>]*\/>/g) || []).forEach(function (tag) {
      var nm = unesc(getAttr(tag, 'name') || '');
      if (!/^Loi can sua( \d+)?$/.test(nm)) return;
      var rid = getAttr(tag, 'r:id') || getAttr(tag, 'id');
      var part = rel[rid];
      wbXml = wbXml.replace(tag, '');
      relsXml = relsXml.replace(new RegExp('<Relationship[^>]*Id="' + rid + '"[^>]*\\/>'), '');
      if (part && files[part]) {
        delete files[part];
        writeText('[Content_Types].xml', readText('[Content_Types].xml')
          .replace(new RegExp('<Override[^>]*PartName="/' + part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"[^>]*\\/>'), ''));
      }
      dropped.push(nm);
    });

    if (!dropped.length) return;
    wbXml = wbXml.replace(/\s*activeTab="\d+"/, '');
    writeText('xl/workbook.xml', wbXml);
    writeText('xl/_rels/workbook.xml.rels', relsXml);
    report.push('  gỡ ' + dropped.length + ' sheet đánh dấu lỗi khỏi file nền (' +
      dropped.join(', ') + ') — không để lọt vào báo cáo gửi khách');
  }

  /* ---------- điểm vào ---------- */
  function patch(masterBytes, blocks, cfg, opts) {
    opts = opts || {};
    if (typeof fflate === 'undefined') throw new Error('Thiếu thư viện fflate');
    var report = [];
    (blocks.appliedAdjustments || []).forEach(function (a) {
      report.push('  Điều chỉnh đã xác nhận [' + a.date + '] ' + a.kind + ' ' + a.key +
        ': ' + a.from + ' → ' + a.to + '; nguồn: ' + a.source);
    });
    var files = fflate.unzipSync(masterBytes);
    var names = Object.keys(files);
    var originalCount = names.length;

    function readText(n) { return fflate.strFromU8(files[n]); }
    function writeText(n, s) { files[n] = fflate.strToU8(s); }

    /* sheet name -> đường dẫn part */
    var wbXml = readText('xl/workbook.xml');
    var relsXml = readText('xl/_rels/workbook.xml.rels');
    var rel = {};
    (relsXml.match(/<Relationship\b[^>]*\/>/g) || []).forEach(function (t) {
      var id = getAttr(t, 'Id'), tgt = getAttr(t, 'Target');
      if (id && tgt) rel[id] = tgt.charAt(0) === '/' ? tgt.slice(1) : 'xl/' + tgt.replace(/^\.\//, '');
    });
    var sheetPart = {};
    (wbXml.match(/<sheet\b[^>]*\/>/g) || []).forEach(function (t) {
      var nm = unesc(getAttr(t, 'name') || '');
      var id = getAttr(t, 'r:id') || getAttr(t, 'id');
      if (nm && rel[id]) sheetPart[nm] = rel[id];
    });

    /* Bản đánh dấu lỗi có thêm một sheet liệt kê — gỡ trước khi làm gì khác */
    stripMarkupSheets(files, readText, writeText, report);
    wbXml = readText('xl/workbook.xml');
    sheetPart = {};
    (wbXml.match(/<sheet\b[^>]*\/>/g) || []).forEach(function (t) {
      var nm = unesc(getAttr(t, 'name') || '');
      var id = getAttr(t, 'r:id') || getAttr(t, 'id');
      if (nm && rel[id]) sheetPart[nm] = rel[id];
    });

    /* Bảng tra trong XML phải giống bảng Pipeline đã dùng để tính KPI/cache. */
    var mailEdits = Pipeline.mailSfCorrections(opts.masterWb, cfg);
    if (mailEdits.length) {
      var mailPart = sheetPart[cfg.mailSf.sheet];
      if (!mailPart || !files[mailPart]) throw new Error('Không tìm thấy sheet mail SF để cập nhật tên');
      writeText(mailPart, patchCells(readText(mailPart), mailEdits));
      report.push('  mail SF: chuẩn hóa ' + mailEdits.length + ' tên theo email đã biết để VLOOKUP tính lại đúng.');
    }

    /* công thức giữ nguyên như file gốc */
    var fm = (cfg.masterFormulas || {});
    var emailF = fm.sfcaseEmail || "VLOOKUP($F{row},'mail SF'!$C$3:$D$69,2,0)";
    /* Email chưa tra được phải vẫn trống sau Excel refresh, như Pipeline.
       Cảnh báo nguồn đã chỉ rõ case thiếu người; không để #N/A lọt vào Pivot. */
    if (!/^\s*IFERROR\s*\(/i.test(emailF)) emailF = 'IFERROR(' + emailF + ',"")';
    var totalF = fm.otherTotal || 'SUM(G{row}:M{rowEnd})';
    var blockSize = (cfg.op && cfg.op.other && cfg.op.other.blockSize) || 8;

    var specs = [
      { block: blocks.block1, formulas: null },
      { block: blocks.block2, formulas: {
          5: function (row) { return emailF.replace(/\{row\}/g, row); }
        } },
      { block: blocks.block3, formulas: {
          /* SUM chỉ đặt ở dòng đầu mỗi khối supporter, giống file gốc */
          3: function (row, value, idx) {
            if (idx % blockSize !== 0) return null;
            return totalF.replace(/\{row\}/g, row).replace(/\{rowEnd\}/g, row + blockSize - 1);
          }
        } }
    ];

    specs.forEach(function (sp) {
      var b = sp.block;
      if (!b) return;
      var part = sheetPart[b.sheet];
      if (!part || !files[part]) {
        throw new Error('File báo cáo tổng không có sheet "' + b.sheet + '"');
      }
      writeText(part, patchSheet(readText(part), {
        sheetName: b.sheet, anchor: b.anchor, rows: b.rows,
        types: b.types, formulas: sp.formulas
      }, report));
    });

    /* Cột ngày ở sheet tổng hợp (cộng dồn cả tháng, không bị thay như 3 sheet dữ liệu) */
    if (cfg.summary && blocks.summary) {
      var sPart = sheetPart[cfg.summary.sheet];
      if (!sPart || !files[sPart]) {
        report.push('  ' + cfg.summary.sheet + ': không có sheet này — bỏ qua phần tổng hợp');
      } else {
        writeText(sPart, patchSummary(readText(sPart), blocks.summary, cfg,
          blocks.reportSerial, report));
      }
    }

    /* KPI ngày + KPI tháng */
    var kpiInfo = {};
    patchKpi(files, sheetPart, readText, writeText, blocks, cfg, opts.masterWb, report, kpiInfo);

    /* 6 dòng fanpage: Follow từ extension, Like chép của ngày gần nhất */
    var fanpageRows = patchFanpage(files, sheetPart, readText, writeText, cfg,
      blocks.reportSerial, blocks.fanpage, opts.masterWb, report);

    /* Sheet Summarize: kéo công thức gương sang cột của ngày mới */
    patchSummarizeMirror(files, sheetPart, readText, writeText, cfg, blocks.reportSerial, report);

    /* Bước 2 quy trình: ngày ở đầu mỗi tab (nằm ngoài vùng 3 khối dữ liệu) */
    patchDateHeaders(files, sheetPart, readText, writeText, cfg, blocks.reportSerial, report);

    /* Bước 14 quy trình: kéo cột phụ Sheet5 xuống cho dư chỗ */
    patchSheet5(files, sheetPart, readText, writeText, cfg, report);

    /* Bắt Excel tính lại toàn bộ công thức khi mở */
    if (/<calcPr\b[^>]*\/>/.test(wbXml)) {
      wbXml = wbXml.replace(/<calcPr\b([^>]*?)\/>/, function (mm, at) {
        return '<calcPr' + at.replace(/\s*fullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"/>';
      });
    } else {
      wbXml = wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    }
    writeText('xl/workbook.xml', wbXml);
    report.push('  workbook: bật tính lại công thức khi mở (fullCalcOnLoad)');

    /* Pivot tự refresh khi mở, vì pivotCache đang giữ dữ liệu của ngày cũ */
    var nPivot = 0;
    names.forEach(function (n) {
      if (n.indexOf('pivotCacheDefinition') === -1 || n.indexOf('.rels') !== -1) return;
      var x = readText(n).replace(/<pivotCacheDefinition\b([^>]*?)>/, function (mm, at) {
        return '<pivotCacheDefinition' + at.replace(/\s*refreshOnLoad="[^"]*"/, '') +
               ' refreshOnLoad="1">';
      });
      writeText(n, x);
      nPivot++;
    });
    if (nPivot) report.push('  pivot: bật refreshOnLoad cho ' + nPivot + ' pivotCache');

    var nField = fixPivotFilters(files, names, readText, writeText);
    if (nField) {
      report.push('  pivot: bật "nhận mục mới vào bộ lọc" cho ' + nField +
        ' trường — nếu không, ngày báo cáo mới bị bộ lọc tay loại sạch khi refresh');
    }

    /* calcChain trỏ tới ô theo thứ tự cũ; giữ lại sẽ làm Excel báo file hỏng.
       Xoá đi thì Excel tự dựng lại. */
    if (files['xl/calcChain.xml']) {
      delete files['xl/calcChain.xml'];
      writeText('[Content_Types].xml',
        readText('[Content_Types].xml').replace(/<Override[^>]*calcChain\.xml[^>]*\/>/g, ''));
      writeText('xl/_rels/workbook.xml.rels',
        readText('xl/_rels/workbook.xml.rels')
          .replace(/<Relationship\b[^>]*Target="calcChain\.xml"[^>]*\/>/g, ''));
      report.push('  calcChain.xml: đã xoá (Excel tự dựng lại khi mở)');
    }

    /* Nén lại, giữ nguyên thứ tự part gốc */
    var ordered = {};
    names.forEach(function (n) { if (files[n]) ordered[n] = files[n]; });
    var bytes = fflate.zipSync(ordered, { level: 6, mtime: new Date() });

    report.push('  zip: giữ ' + Object.keys(ordered).length + '/' + originalCount + ' part');
    return { bytes: bytes, report: report, partCount: Object.keys(ordered).length,
             oddRateEmails: kpiInfo.oddRateEmails || {},
             teamRows: kpiInfo.teamRows || [], teamEmails: kpiInfo.teamEmails || [],
             /* dòng fanpage tool THỰC SỰ ghi — phần đối chứng chỉ so những dòng này,
                vì Follow chỉ có khi người dùng dán số từ extension vào */
             fanpageRows: fanpageRows };
  }

  return { patch: patch, numToCol: numToCol, colToNum: colToNum, parseAddr: parseAddr,
           teamRowsFromFormulas: teamRowsFromFormulas };
});
