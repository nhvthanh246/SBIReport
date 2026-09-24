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
  if (typeof module === 'object' && module.exports) module.exports = factory(root.Parse);
  else root.XlsxPatch = factory(root.Parse);
})(typeof self !== 'undefined' ? self : this, function (P) {
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
      if (!hasData && orig && orig.inner.indexOf('<v>') === -1 && orig.inner.indexOf('<f') === -1) {
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
      xml = xml.replace(row.raw, open + cells.map(function (c) { return c.raw; }).join('') + '</row>');
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

  function patchKpi(files, sheetPart, readText, writeText, blocks, cfg, masterWb, report, info) {
    var kc = cfg.kpi;
    if (!kc || !blocks.kpi) return;
    var serial = blocks.reportSerial;

    /* ô ngày của Daily KPI Result do patchDateHeaders lo */

    /* --- danh sách người lấy từ Daily KPI Result ---
       `opRows` = khối OP, chỉ dùng cho dòng tổng team (đúng vùng của =SUM(E15:E43)).
       `allRows` = mọi người có dòng KPI, kể cả những người nằm ngoài tổng OP. */
    var opRows = emailRows(masterWb, kc.daily.sheet, kc.daily.emailColumn,
      kc.daily.dataStartRow, kc.daily.dataEndRow);
    var allRows = emailRows(masterWb, kc.daily.sheet, kc.daily.emailColumn,
      kc.daily.dataStartRow, kc.daily.peopleEndRow || kc.daily.dataEndRow);
    var opEmails = Object.keys(opRows);

    /* Bỏ những người dùng định mức KPI riêng — tính theo định mức chuẩn sẽ ra số sai */
    var odd = oddRateRows(files, sheetPart, cfg, readText);
    var oddEmails = {};
    Object.keys(allRows).forEach(function (e) {
      if (odd[allRows[e]]) oddEmails[e] = odd[allRows[e]];
    });
    if (info) info.oddRateEmails = oddEmails;
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
      var k = blocks.kpi.forEmail(e);
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
    var masters = sharedMasters(xml);

    /* Công thức cột phụ ở từng dòng, đã giải shared. */
    function helperRow(r) {
      var row = rowInner(xml, r);
      if (!row) return null;
      var got = null, cm;
      CELL_RE.lastIndex = 0;
      while ((cm = CELL_RE.exec(row.inner)) !== null) {
        var craw = cm[0];
        var ct = craw.slice(0, craw.indexOf('>') + 1);
        var ref = getAttr(ct, 'r');
        if (!ref) continue;
        var ci = parseAddr(ref).c;
        if (sc.columns.indexOf(ci) === -1) continue;
        var f = realFormula(craw, r, masters);
        if (!f) continue;
        (got = got || {})[ci] = { f: f, s: getAttr(ct, 's') };
      }
      return got;
    }

    /* Mẫu = dòng CUỐI đang có công thức, đúng như người ta kéo từ đáy bảng xuống.
       Lấy dòng đầu sẽ sai: trong file 15/9 dòng 5-9 còn giữ biến thể cũ
       (O = J-…+I*45) còn dòng 10-33 đã đổi sang O = J-…-I. */
    var last = 0, tplCells = null;
    for (var r = sc.templateRow; r <= sc.extendToRow; r++) {
      var h = helperRow(r);
      if (h) { last = r; tplCells = h; }
    }
    if (!tplCells) {
      report.push('  ' + sc.sheet + ': không tìm thấy công thức cột phụ trong dòng ' +
        sc.templateRow + '-' + sc.extendToRow + ' — bỏ qua');
      return;
    }

    /* Cảnh báo nếu các dòng hiện có không cùng một công thức (file gốc đang lệch). */
    var variants = {};
    for (var r2 = sc.templateRow; r2 <= last; r2++) {
      var h2 = helperRow(r2);
      if (!h2) continue;
      var key = Object.keys(h2).sort().map(function (k) {
        return numToCol(Number(k)) + '=' + rebaseFormula(h2[k].f, r2, 0);
      }).join(' | ');
      (variants[key] = variants[key] || []).push(r2);
    }
    var vKeys = Object.keys(variants);
    if (vKeys.length > 1) {
      vKeys.sort(function (a, b) { return variants[b].length - variants[a].length; });
      report.push('  ⚠ ' + sc.sheet + ': cột phụ đang có ' + vKeys.length +
        ' kiểu công thức khác nhau — ' + vKeys.map(function (k) {
          var g = variants[k];
          return g.length + ' dòng (' + g[0] + '-' + g[g.length - 1] + ')';
        }).join(', ') + '. Tool kéo theo dòng ' + last +
        '; nên kiểm tra lại bằng tay các dòng dùng kiểu cũ.');
    }

    if (last >= sc.extendToRow) {
      report.push('  ' + sc.sheet + ': cột phụ đã phủ tới dòng ' + last + ' — không cần kéo thêm');
      return;
    }

    var edits = [];
    for (var rr = last + 1; rr <= sc.extendToRow; rr++) {
      Object.keys(tplCells).forEach(function (k) {
        var ci2 = Number(k);
        edits.push({ row: rr, col: ci2, value: '', type: 'auto',
                     formula: rebaseFormula(tplCells[k].f, last, rr), style: tplCells[k].s });
      });
    }
    writeText(part, patchCells(xml, edits, { createMissing: true, report: report }));
    report.push('  ' + sc.sheet + ': kéo công thức cột phụ ' +
      Object.keys(tplCells).map(function (k) { return numToCol(Number(k)); }).join('/') +
      ' từ dòng ' + (last + 1) + ' đến ' + sc.extendToRow +
      ' (mẫu lấy ở dòng ' + last + ') để pivot nở thêm người vẫn có số');
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

  /* ---------- điểm vào ---------- */
  function patch(masterBytes, blocks, cfg, opts) {
    opts = opts || {};
    if (typeof fflate === 'undefined') throw new Error('Thiếu thư viện fflate');
    var report = [];
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

    /* công thức giữ nguyên như file gốc */
    var fm = (cfg.masterFormulas || {});
    var emailF = fm.sfcaseEmail || "VLOOKUP($F{row},'mail SF'!$C$3:$D$69,2,0)";
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
      if (!b || !b.rows.length) {
        report.push('  ' + b.sheet + ': bỏ qua (không có dữ liệu)');
        return;
      }
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
             oddRateEmails: kpiInfo.oddRateEmails || {} };
  }

  return { patch: patch, numToCol: numToCol, colToNum: colToNum, parseAddr: parseAddr };
});
