/* Các phép biến đổi thuần (không đụng DOM) — dùng chung cho UI và cho self-test. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(root.Parse);
  else root.Pipeline = factory(root.Parse);
})(typeof self !== 'undefined' ? self : this, function (P) {
  'use strict';

  /* Kiểu ghi của từng cột trong khối dán.
     'date'  -> số serial + định dạng ngày
     'text'  -> ép kiểu chuỗi (ID 16-17 chữ số, timestamp)
     'auto'  -> để SheetJS tự suy (giờ làm là số, còn lại là chuỗi) */
  var BLOCK1_TYPES = ['number', 'date', 'auto', 'auto', 'auto', 'text', 'text',
                      'auto', 'auto', 'auto', 'text', 'text', 'auto', 'auto'];
  var BLOCK1_HEADERS = ['No/順序数', 'Date/日付', 'CHANNEL/連絡チャネル', 'Customer/顧客',
    'Supporter/サポーター', 'Start Time/始まる時間', 'End Time/終了時間',
    'Request Description/リクエストの説明', 'Session Note/セッションノート',
    'Topic Level 1/トピックレベル', 'sender_id/連絡先', 'ma_kh/顧客コード',
    'DETAIL of Outbound Reply', 'TEAM'];

  var BLOCK2_TYPES = ['text', 'auto', 'auto', 'text', 'auto', 'auto', 'date'];
  var BLOCK3_TYPES = ['date', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto'];
  var BLOCK3_HEADERS = ['DATE', 'TEAM', 'Supporter Email', 'TỔNG GIỜ LÀM',
    'Tổng số giờ đã hỗ trợ KH', 'Tổng số giờ đã Outbound', 'Tổng số giờ đã xử lý case',
    'Trả lời Comment', 'Đăng ký mới + webform + Gọi OB', 'Khác', 'Số giờ khác'];

  /* `spots` = danh sách ô gây ra cảnh báo, để UI chỉ đúng tệp/dòng/ô và để dựng
     bản Excel đánh dấu. Mỗi spot: { file, sheet, row (1-based), col (0-based|null),
     value, why }. Cắt bớt ở SPOT_CAP để một cảnh báo hàng nghìn dòng không làm
     phình bộ nhớ; `spotTotal` giữ lại con số thật. */
  var SPOT_CAP = 400;

  function warn(list, level, message, detail, spots, display) {
    var item = { level: level, message: message, detail: detail || null };
    /* Lời nhắc ngắn cho giao diện; giữ nguyên giải thích và vị trí đầy đủ để tra cứu. */
    if (display) item.display = display;
    if (spots && spots.length) {
      item.spotTotal = spots.length;
      item.spots = spots.slice(0, SPOT_CAP);
    }
    list.push(item);
  }

  function spot(file, sheet, row, col, value, why) {
    return { file: file, sheet: sheet, row: row,
             col: (col === undefined || col === null || col < 0) ? null : col,
             value: P.cellToString(value), why: why || '' };
  }

  /* --- Cột luôn ghi trống (xem `blankColumns` trong config) ---
   * Không im lặng vứt dữ liệu: đếm lại từng giá trị bị bỏ rồi báo cho người dùng,
   * để nếu một ngày nào đó cột ấy có dữ liệu thật thì còn biết mà sửa config. */
  function blankSet(list) {
    var out = {};
    (list || []).forEach(function (ci) { out[ci] = true; });
    return out;
  }
  function noteBlanked(store, ci, value) {
    if (P.isBlank(value)) return;
    var v = P.cellToString(value).trim();
    if (!v) return;
    var bucket = store[ci] || (store[ci] = {});
    bucket[v] = (bucket[v] || 0) + 1;
  }
  function warnBlanked(warnings, store, sheetName) {
    Object.keys(store).forEach(function (ci) {
      var bucket = store[ci];
      var parts = Object.keys(bucket).sort(function (a, b) { return bucket[b] - bucket[a]; })
        .slice(0, 8).map(function (v) { return JSON.stringify(v) + ' ×' + bucket[v]; });
      var total = Object.keys(bucket).reduce(function (a, v) { return a + bucket[v]; }, 0);
      warn(warnings, 'warn', 'Sheet "' + sheetName + '": bỏ ' + total +
        ' giá trị ở cột luôn-để-trống (cột nguồn thứ ' + (Number(ci) + 1) + ')',
        parts.join(' · ') + ' — cột này trống ở mọi dòng của File báo cáo tổng nên ' +
        'tool ghi trống. Nếu nay cột đó có nghĩa thật, bỏ nó khỏi `blankColumns` trong config.', null,
        { title: total + ' giá trị không được đưa vào báo cáo',
          summary: 'Cấu hình hiện tại yêu cầu để trống cột này. Nếu cần giữ dữ liệu, hãy sửa cấu hình rồi chạy lại.',
          impact: 'Các giá trị dưới đây không được ghi vào báo cáo.',
          action: 'Nếu đây là dữ liệu cần giữ, kiểm tra danh sách cột để trống trong Cấu hình trước khi chạy lại.',
          note: 'Sheet: ' + sheetName + ' · Cột nguồn: ' + XLSX.utils.encode_col(Number(ci)),
          tables: [{ title: 'Giá trị bị bỏ qua', headers: ['Giá trị nguồn', 'Số lần xuất hiện'],
            rows: Object.keys(bucket).map(function (v) { return [v, bucket[v]]; }) }] });
    });
  }

  /**
   * Đoán file vừa nạp là loại nào, để người dùng thả cả 4 file một lượt
   * mà không phải tự phân loại. Dựa vào tên sheet trước (rẻ), sau đó mới
   * đọc nội dung mấy dòng đầu.
   */
  function detectKind(wb) {
    var sheets = (wb.SheetNames || []).map(function (n) { return P.normText(n); });
    function has(sub) {
      return sheets.some(function (n) { return n.indexOf(P.normText(sub)) !== -1; });
    }
    /* Phải kiểm tra file tổng TRƯỚC: nó cũng có "1. CSKH" và "3. Other Tasks" */
    if (has('summarize')) return 'master';
    if (has('1. cskh') && has('other tasks')) return 'ops';

    var ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) return null;
    var m = P.sheetToMatrix(ws);
    var top = '';
    for (var r = 0; r < Math.min(m.rows.length, 20); r++) {
      top += ' ' + (m.rows[r] || []).map(P.cellToString).join(' ');
    }
    top = P.normText(top);
    if (top.indexOf('conversation id') !== -1) return 'session';
    if (top.indexOf('case number') !== -1) return 'ccvn';
    return null;
  }

  /* =====================================================================
   * Nguồn 1 — Session Report (chat bot)
   * ===================================================================== */
  function processSession(wb, cfg, fileName) {
    var c = cfg.session;
    var warnings = [];
    var name = typeof c.sheet === 'number' ? wb.SheetNames[c.sheet] : c.sheet;
    var ws = wb.Sheets[name];
    if (!ws) throw new Error('Không tìm thấy sheet "' + name + '" trong ' + fileName);

    var m = P.sheetToMatrix(ws);
    var headerIdx = c.headerRow - 1;
    var headers = (m.rows[headerIdx] || []).slice(0, m.maxCol + 1);
    if (!headers.length) throw new Error('Sheet "' + name + '" không có dòng header ở dòng ' + c.headerRow);

    /* Vị trí cột cho từng luật lọc */
    var rules = (c.dropRules || []).map(function (r) {
      var idx = P.findColumn(headers, r.column);
      if (idx === -1) warn(warnings, 'error', 'Luật lọc trỏ tới cột không tồn tại: "' + r.column + '"', null, null,
        { title: 'Thiếu cột dùng để lọc dữ liệu',
          summary: 'Một quy tắc lọc chưa được áp dụng. Kiểm tra tên cột trong file nguồn và cấu hình.' });
      return { rule: r, col: idx };
    }).filter(function (r) { return r.col !== -1; });

    /* Vị trí cột cho sheet staging */
    var stagingCols = c.stagingColumns.map(function (sc) {
      if (sc.from === 'const') return { spec: sc, col: -1 };
      var idx = P.findColumn(headers, sc.from);
      if (idx === -1) warn(warnings, 'error', 'Cột staging "' + sc.out + '" không tìm thấy nguồn "' + sc.from + '"', null, null,
        { title: 'Thiếu cột dữ liệu chatbot',
          summary: 'Cột tương ứng trong kết quả sẽ trống. Kiểm tra tên cột trong file nguồn và cấu hình.' });
      return { spec: sc, col: idx };
    });

    var kept = [], keptRowNumbers = [], dropped = [];
    for (var r = headerIdx + 1; r < m.rows.length; r++) {
      var row = m.rows[r];
      if (P.rowIsEmpty(row, m.maxCol)) continue;
      var hit = null;
      for (var i = 0; i < rules.length; i++) {
        if (P.matchRule(row[rules[i].col], rules[i].rule)) { hit = rules[i]; break; }
      }
      if (hit) {
        dropped.push({
          excelRow: r + 1,
          col: hit.col,
          value: P.cellToString(row[hit.col]),
          reason: hit.rule.column + ' ' + hit.rule.op + ' "' + hit.rule.value + '"',
          preview: P.cellToString(row[P.findColumn(headers, 'Customer')]) + ' — ' +
                   P.cellToString(row[hit.col])
        });
      } else {
        kept.push(row);
        keptRowNumbers.push(r + 1);
      }
    }

    var stagingHeaders = c.stagingColumns.map(function (sc) { return sc.out; });
    if (c.stagingColumns[3]) warn(warnings, c.stagingColumns[3].from === 'Requested time' ? 'info' : 'warn',
      'Start Time chatbot lấy từ ' + c.stagingColumns[3].from,
      'SBI đã thống nhất dùng Requested time trong docs/VẤN ĐỀ.docx. Start Time không dùng để suy ra giờ KPI.', null,
      { title: 'Nguồn giờ bắt đầu chatbot',
        summary: c.stagingColumns[3].from === 'Requested time'
          ? 'Lấy thời điểm khách gửi yêu cầu theo thống nhất của SBI. Giờ KPI vẫn lấy từ bảng giờ nhân viên.'
          : 'Đang dùng “' + c.stagingColumns[3].from + '”, khác quy chuẩn SBI là Requested time. Kiểm tra lại Cấu hình.' });
    var stagingRows = kept.map(function (row) {
      return stagingCols.map(function (sc) {
        if (sc.spec.from === 'const') return sc.spec.value;
        var v = row[sc.col];
        return v === undefined || v === null ? '' : v;
      });
    });

    return {
      kind: 'session',
      fileName: fileName,
      sheetName: name,
      headers: headers,
      keptRows: kept,
      keptRowNumbers: keptRowNumbers,
      droppedRows: dropped,
      staging: { name: name + c.stagingSuffix, headers: stagingHeaders, rows: stagingRows },
      stats: { total: kept.length + dropped.length, dropped: dropped.length, kept: kept.length },
      warnings: warnings
    };
  }

  /** Ngày báo cáo = ngày xuất hiện nhiều nhất ở cột First response time. */
  function detectReportDate(sessionResult, cfg) {
    var col = P.findColumn(sessionResult.headers, cfg.session.dateColumn);
    if (col === -1) return null;
    var counts = {};
    sessionResult.keptRows.forEach(function (row) {
      var s = P.toDateSerial(row[col]);
      if (s) counts[s] = (counts[s] || 0) + 1;
    });
    var best = null;
    Object.keys(counts).forEach(function (k) {
      if (!best || counts[k] > counts[best]) best = k;
    });
    return best === null ? null : { serial: +best, iso: P.serialToISO(+best), count: counts[best] };
  }

  /* =====================================================================
   * Nguồn 2 — Daily Report CCVN (Salesforce)
   * ===================================================================== */
  function processCcvn(wb, cfg, fileName) {
    var c = cfg.ccvn;
    var warnings = [];
    var name = typeof c.sheet === 'number' ? wb.SheetNames[c.sheet] : c.sheet;
    var ws = wb.Sheets[name];
    if (!ws) throw new Error('Không tìm thấy sheet "' + name + '" trong ' + fileName);

    var m = P.sheetToMatrix(ws);
    var headerIdx = P.findHeaderRow(m.rows, c.headerMarker, 40);
    if (headerIdx === -1) throw new Error('Không dò được dòng header (mốc "' + c.headerMarker + '") trong ' + fileName);

    var headerRow = m.rows[headerIdx] || [];

    /* Xác định vùng cột của bảng = từ ô header đầu tiên tới ô header cuối cùng.
       Cột A rỗng nằm NGOÀI vùng này nên được giữ (đúng như file mẫu); cột rỗng
       nằm TRONG vùng (cột C) mới bị xoá. */
    var first = -1, last = -1;
    for (var c2 = 0; c2 < headerRow.length; c2++) {
      if (!P.isBlank(headerRow[c2])) { if (first === -1) first = c2; last = c2; }
    }
    if (first === -1) throw new Error('Dòng header rỗng trong ' + fileName);

    /* Vị trí dòng Total, để không tính nhãn chân bảng là "có dữ liệu".
       Trong file mẫu dòng Total là `Total | Count | 379`: nhãn "Count" nằm ngay ở
       cột rỗng cần xoá và vẫn bị xoá cùng cột đó. */
    var totalIdx = -1;
    for (var rt = m.rows.length - 1; rt > headerIdx; rt--) {
      var rrow = m.rows[rt] || [];
      for (var kk = 0; kk < rrow.length; kk++) {
        if (P.isBlank(rrow[kk])) continue;
        if (P.normText(rrow[kk]) === P.normText(c.totalRowMarker)) totalIdx = rt;
        break;
      }
      if (totalIdx !== -1) break;
    }
    var dataEnd = totalIdx === -1 ? m.rows.length : totalIdx;

    var removed = [];
    if (c.dropEmptyColumns) {
      for (var col = first; col <= last; col++) {
        if (!P.isBlank(headerRow[col])) continue;
        var hasData = false;
        for (var r2 = headerIdx + 1; r2 < dataEnd && !hasData; r2++) {
          if (!P.isBlank((m.rows[r2] || [])[col])) hasData = true;
        }
        if (hasData) {
          warn(warnings, 'warn', 'Cột ' + XLSX.utils.encode_col(col) + ' không có header nhưng có dữ liệu — giữ lại');
        } else {
          removed.push(col);
        }
      }
    }
    var removedSet = {};
    removed.forEach(function (x) { removedSet[x] = true; });

    function squeeze(row) {
      var out = [];
      for (var i = 0; i <= m.maxCol; i++) {
        if (removedSet[i]) continue;
        out.push(row && row[i] !== undefined ? row[i] : '');
      }
      while (out.length && P.isBlank(out[out.length - 1])) out.pop();
      return out;
    }

    /* Bỏ cột làm lệch chỉ số cột, nên giữ bảng tra ngược: cột thứ i của bảng đã
       bỏ cột ứng với cột nào trong FILE GỐC. Cần khi báo lỗi để chỉ đúng ô. */
    var srcCol = [];
    for (var sc = 0; sc <= m.maxCol; sc++) { if (!removedSet[sc]) srcCol.push(sc); }

    /* Ma trận sạch: giữ nguyên metadata, header, dữ liệu và dòng Total */
    var cleaned = [];
    for (var r3 = 0; r3 < m.rows.length; r3++) {
      cleaned.push(squeeze(m.rows[r3]));
    }
    while (cleaned.length && !cleaned[cleaned.length - 1].length) cleaned.pop();

    var headers = cleaned[headerIdx] || [];
    var dataRows = [], dataRowNumbers = [], totalRow = null;
    for (var r4 = headerIdx + 1; r4 < cleaned.length; r4++) {
      var row = cleaned[r4];
      if (!row.length) continue;
      var firstCell = '';
      for (var k = 0; k < row.length; k++) {
        if (!P.isBlank(row[k])) { firstCell = P.normText(row[k]); break; }
      }
      if (firstCell === P.normText(c.totalRowMarker)) { totalRow = row; continue; }
      dataRows.push(row);
      dataRowNumbers.push(r4 + 1);
    }

    return {
      kind: 'ccvn',
      fileName: fileName,
      sheetName: name,
      cleaned: cleaned,
      headerIndex: headerIdx,
      headers: headers,
      dataRows: dataRows,
      /* số dòng trong FILE GỐC của từng dòng dữ liệu, cùng thứ tự với dataRows */
      dataRowNumbers: dataRowNumbers,
      srcCol: srcCol,
      totalRow: totalRow,
      removedColumns: removed.map(function (x) { return XLSX.utils.encode_col(x); }),
      stats: { total: dataRows.length, removedColumns: removed.length },
      warnings: warnings
    };
  }

  /* =====================================================================
   * Nguồn 3 — file báo cáo ngày của nhân viên (SBI DAILY REPORT)
   * ===================================================================== */
  function readOpSheet(wb, spec, reportSerial, fileName, warnings, acceptSwapped) {
    var ws = wb.Sheets[spec.sheet];
    if (!ws) {
      warn(warnings, 'warn', 'File ' + fileName + ' không có sheet "' + spec.sheet + '" — bỏ qua', null, null,
        { title: 'Thiếu sheet trong báo cáo nhân viên',
          summary: 'Chưa lấy được phần dữ liệu này. Kiểm tra tên sheet hoặc nạp lại file đầy đủ.' });
      return { rows: [], scanned: 0, headers: [] };
    }
    var m = P.sheetToMatrix(ws);
    var headers = m.rows[spec.headerRow - 1] || [];
    var dateCol = P.findColumn(headers, spec.dateColumn);
    if (dateCol === -1) {
      warn(warnings, 'error', 'Sheet "' + spec.sheet + '" trong ' + fileName +
        ' không có cột ngày "' + spec.dateColumn + '"', null, null,
        { title: 'Không tìm thấy cột ngày',
          summary: 'Chưa lấy được dữ liệu từ sheet này. Kiểm tra tên cột ngày trong file và cấu hình.' });
      return { rows: [], scanned: 0, headers: headers };
    }

    /* Một dòng chỉ có ô ngày mà không có nội dung nào ở các cột được copy thì
       không mang thông tin gì. Việc kéo ô ngày xuống hàng chục nghìn dòng trống
       là chuyện có thật trong file nhân viên — không lọc thì báo cáo phình ra
       hàng chục nghìn dòng rác. */
    var copyCols = spec.copyColumns || [];
    function hasContent(row) {
      for (var i = 0; i < copyCols.length; i++) {
        if (!P.isBlank(row[copyCols[i]])) return true;
      }
      return false;
    }

    var picked = [], otherDayRows = [], scanned = 0, unparsed = [], swapped = [], dateOnly = 0;
    var noDateRows = [], unparsedRows = [];
    /* Đếm số phận của TỪNG dòng để trả lời được câu "dòng nào bị bỏ, vì sao".
       Không có bảng này thì chỉ thấy kết quả cuối, không biết đã bỏ cái gì. */
    var tally = { emptyRow: 0, otherDay: 0, scannedAll: 0 };
    for (var r = spec.dataStartRow - 1; r < m.rows.length; r++) {
      var row = m.rows[r];
      tally.scannedAll++;
      if (P.rowIsEmpty(row, m.maxCol)) { tally.emptyRow++; continue; }
      if (!hasContent(row)) { dateOnly++; continue; }
      scanned++;
      var serial = P.toDateSerial(row[dateCol]);
      if (serial === null) {
        /* Dòng có nội dung nhưng không xác định được ngày sẽ bị bỏ qua hoàn toàn —
           phải báo ra, nếu không sẽ thiếu dữ liệu mà không ai biết. */
        if (P.isBlank(row[dateCol])) {
          noDateRows.push(spot(fileName, spec.sheet, r + 1, dateCol, '',
            'Ô ngày để trống nên không biết dòng này thuộc ngày nào — dòng bị bỏ qua'));
        } else {
          unparsedRows.push(spot(fileName, spec.sheet, r + 1, dateCol, row[dateCol],
            'Giá trị ngày không đọc được — dòng bị bỏ qua'));
          if (unparsed.indexOf(P.cellToString(row[dateCol])) === -1) {
            unparsed.push(P.cellToString(row[dateCol]));
          }
        }
        continue;
      }
      if (serial === reportSerial) {
        picked.push({ excelRow: r + 1, cells: row });
        continue;
      }
      /* Chỉ là dấu hiệu có thể đảo ngày/tháng, không chứng minh nhập sai. */
      if (P.swappedSerial(row[dateCol]) === reportSerial) {
        swapped.push({ excelRow: r + 1, cells: row, raw: P.cellToString(row[dateCol]) });
        if (acceptSwapped) picked.push({ excelRow: r + 1, cells: row, swapped: true });
        continue;
      }
      tally.otherDay++;
      if (spec.blockSize) otherDayRows.push({ excelRow: r + 1, cells: row, serial: serial });
    }
    if (unparsed.length) {
      warn(warnings, 'error', 'Sheet "' + spec.sheet + '" trong ' + fileName +
        ': ' + unparsedRows.length + ' dòng có giá trị ngày không đọc được — bị bỏ qua',
        'Các giá trị gặp phải: ' + unparsed.slice(0, 10).join(' · '), unparsedRows,
        { title: unparsedRows.length + ' dòng có ngày không đọc được',
          summary: 'Các dòng này chưa được tính vào báo cáo. Sửa ô ngày trong file nguồn rồi chạy lại.' });
    }
    if (noDateRows.length) {
      warn(warnings, 'warn', 'Sheet "' + spec.sheet + '" trong ' + fileName + ': ' +
        noDateRows.length + ' dòng có nội dung nhưng bỏ trống ô ngày — bị bỏ qua',
        null, noDateRows,
        { title: noDateRows.length + ' dòng chưa điền ngày',
          summary: 'Các dòng này chưa được tính vào báo cáo. Điền đúng ngày trong file nguồn rồi chạy lại.' });
    }
    if (dateOnly > 50) {
      warn(warnings, 'warn', 'Sheet "' + spec.sheet + '" trong ' + fileName + ': bỏ qua ' +
        dateOnly + ' dòng chỉ có ô ngày, không có nội dung',
        'Thường do kéo ô ngày xuống quá nhiều dòng trống. Nên xoá bớt ở file nguồn cho nhẹ file.', null,
        { title: 'Đã bỏ qua ' + dateOnly + ' dòng chỉ có ngày',
          summary: 'Các dòng này không có nội dung công việc. Có thể xóa bớt trong file nguồn để file nhẹ hơn.' });
    }
    if (swapped.length) {
      var samples = {};
      swapped.forEach(function (s) {
        var shown = P.sourceCellText(wb, spec.sheet, s.excelRow, dateCol, s.raw);
        samples[shown] = (samples[shown] || 0) + 1;
      });
      warn(warnings, acceptSwapped ? 'warn' : 'error',
        'Sheet "' + spec.sheet + '" trong ' + fileName + ': ' + swapped.length +
        ' dòng có thể bị đảo ngày/tháng — ' +
        (acceptSwapped ? 'đã tính theo tùy chọn' : 'chưa tính vào báo cáo'),
        'Giá trị ghi trong file: ' +
        Object.keys(samples).map(function (k) { return k + ' (' + samples[k] + ' dòng)'; }).join(' · ') +
        ' — xác nhận ngày làm việc trước khi sửa nguồn hoặc thay đổi tùy chọn nhận ngày đảo.',
        swapped.map(function (sw) {
          var parts = P.toDateParts(sw.raw), wanted = P.toDateParts(reportSerial);
          function readable(d) { return 'ngày ' + d.d + ' tháng ' + d.m + ' năm ' + d.y; }
          return spot(fileName, spec.sheet, sw.excelRow, dateCol, sw.raw,
            'Ô này được đọc là ' + readable(parts) + ', còn báo cáo đang làm là ' + readable(wanted) + '. ' +
            (acceptSwapped ? 'Tool đã nhận dòng này theo tùy chọn đảo ngày/tháng. ' : 'Tool chưa lấy dòng này. ') +
            'Xác nhận ngày làm việc; nếu nhập nhầm, sửa lại ngày trong ô gốc rồi chạy lại. Chỉ đổi định dạng hiển thị không sửa được ngày đã nhập sai.');
        }), { title: swapped.length + ' dòng có thể bị đảo ngày và tháng',
          summary: acceptSwapped
            ? 'Đã tính vào báo cáo theo tùy chọn hiện tại. Hãy kiểm tra các ô ngày để tránh nhận nhầm dữ liệu.'
            : 'Chưa tính vào báo cáo. Kiểm tra ngày trước khi sửa nguồn hoặc bật tùy chọn nhận ngày đảo.' });
    }
    /* Bảng kê số phận từng dòng — luôn hiện, để trả lời ngay được câu
       "tool bỏ dòng nào, vì sao" mà không cần ai đi dò file. */
    var emailCol = spec.emailColumn === undefined ? 4 : spec.emailColumn;
    var people = {};
    picked.forEach(function (x) {
      var e = P.cellToString(x.cells[emailCol]).trim();
      if (e) people[e] = true;
    });
    var parts = [
      'lấy ' + picked.length + ' dòng' +
        (Object.keys(people).length ? ' (' + Object.keys(people).length + ' người)' : ''),
      'ngày khác: ' + tally.otherDay
    ];
    if (tally.emptyRow) parts.push('dòng trống: ' + tally.emptyRow);
    if (dateOnly) parts.push('chỉ có ô ngày, không nội dung: ' + dateOnly);
    if (noDateRows.length) parts.push('BỎ TRỐNG ô ngày: ' + noDateRows.length);
    if (unparsedRows.length) parts.push('ngày KHÔNG đọc được: ' + unparsedRows.length);
    if (swapped.length) parts.push('ĐẢO ngày/tháng: ' + swapped.length +
      (acceptSwapped ? ' (đã tính)' : ' (đang bỏ sót)'));
    warn(warnings, 'info', 'Sheet "' + spec.sheet + '" trong ' + fileName + ': quét ' +
      tally.scannedAll + ' dòng', parts.join(' · ') +
      '. Cột ngày đọc ở "' + spec.dateColumn + '". Những dòng ghi BỎ TRỐNG / KHÔNG đọc được / ' +
      'ĐẢO ngày đều có địa chỉ ô cụ thể ở cảnh báo riêng bên trên.', null,
      { title: 'Đã lấy ' + picked.length + ' dòng từ báo cáo nhân viên',
        summary: 'Sheet “' + spec.sheet + '”. Xem chi tiết để biết những dòng được lấy hoặc bỏ qua.',
        note: 'File: ' + fileName + ' · Sheet: ' + spec.sheet,
        tables: [{ title: 'Kết quả đọc dữ liệu', headers: ['Nhóm dòng', 'Số dòng'],
          rows: [['Được lấy vào báo cáo', picked.length], ['Thuộc ngày khác', tally.otherDay],
            ['Dòng trống', tally.emptyRow], ['Chỉ có ngày, không có nội dung', dateOnly],
            ['Chưa điền ngày', noDateRows.length], ['Không đọc được ngày', unparsedRows.length],
            ['Đảo ngày/tháng (' + (acceptSwapped ? 'đã tính trong dòng được lấy' : 'chưa được lấy') + ')', swapped.length]] }] });

    return { rows: picked, otherDayRows: otherDayRows, scanned: scanned, headers: headers, swapped: swapped,
             tally: tally, people: Object.keys(people) };
  }

  function processOpFile(wb, cfg, reportSerial, fileName) {
    var warnings = [];
    var acceptSwapped = !!(cfg.op && cfg.op.acceptSwappedDates);
    var cskh = readOpSheet(wb, cfg.op.cskh, reportSerial, fileName, warnings, acceptSwapped);
    var other = readOpSheet(wb, cfg.op.other, reportSerial, fileName, warnings, acceptSwapped);
    if (!cskh.rows.length && !other.rows.length) {
      warn(warnings, 'warn', 'File ' + fileName + ' không có dòng nào đúng ngày báo cáo', null, null,
        { title: 'Chưa có dữ liệu nhân viên cho ngày này',
          summary: 'File vừa nạp không có dòng phù hợp. Kiểm tra ngày báo cáo hoặc chọn lại file nhân viên.' });
    }
    return {
      kind: 'op',
      fileName: fileName,
      cskh: cskh,
      other: other,
      stats: {
        cskh: cskh.rows.length, other: other.rows.length, scanned: cskh.scanned,
        swapped: cskh.swapped.length + other.swapped.length
      },
      warnings: warnings
    };
  }

  /* =====================================================================
   * Số tổng hợp cho cột ngày của sheet Summarize_team VietNam
   * ===================================================================== */
  var B1_CHANNEL = 2, B1_SENDER = 10;

  function computeSummary(block1, cfg, warnings) {
    var sc = cfg.summary;
    if (!sc || !sc.channelGroups) return null;

    var groupOf = {};
    Object.keys(sc.channelGroups).forEach(function (g) {
      sc.channelGroups[g].forEach(function (name) { groupOf[P.normText(name)] = g; });
    });
    var ignored = {};
    (sc.ignoredChannels || []).forEach(function (n) { ignored[P.normText(n)] = true; });

    var counts = {};
    Object.keys(sc.channelGroups).forEach(function (g) { counts[g] = 0; });

    /* Khối "Nội dung hỗ trợ": đếm theo SỐ ĐẦU của cột Topic Level 1
       ("7. Cập nhật thẻ ngoại kiều/在留カード更新" -> "7") */
    var topicRow = {};
    (sc.topicRows || []).forEach(function (t) { topicRow[String(t.topic)] = t; });
    var ignoredTopic = {};
    (sc.ignoredTopics || []).forEach(function (t) { ignoredTopic[String(t)] = true; });
    var topics = {}, unknownTopics = {};
    var tCol = sc.topicColumn;

    var seen = {}, distinct = 0, unknown = {};
    block1.rows.forEach(function (row) {
      var raw = P.cellToString(row[B1_CHANNEL]).trim();
      var g = groupOf[P.normText(raw)];
      if (g) counts[g]++;
      else if (raw && !ignored[P.normText(raw)]) unknown[raw] = (unknown[raw] || 0) + 1;

      /* ①CUSTOMER đếm KHÁCH chứ không đếm phiên, nên gom sender_id duy nhất */
      if (g === 'chatbot') {
        var s = P.cellToString(row[B1_SENDER]).trim();
        if (s && !seen[s]) { seen[s] = 1; distinct++; }
      }

      if (tCol !== undefined) {
        var tv = P.cellToString(row[tCol]).trim();
        if (tv) {
          var m = tv.match(/^\s*(\d+)\s*\./);
          var key = m ? m[1] : null;
          if (key && topicRow[key]) topics[key] = (topics[key] || 0) + 1;
          else if (!key || !ignoredTopic[key]) unknownTopics[tv] = (unknownTopics[tv] || 0) + 1;
        }
      }
    });

    var un = Object.keys(unknown);
    if (un.length && warnings) {
      warn(warnings, 'warn', 'Có ' + un.length + ' giá trị CHANNEL chưa gán nhóm tổng hợp — ' +
        'không được cộng vào cột ngày',
        un.map(function (k) { return k + ' (' + unknown[k] + ')'; }).join(' · ') +
        ' — thêm vào `summary.channelGroups` hoặc `summary.ignoredChannels` trong Cấu hình.', null,
        { title: un.length + ' loại kênh chưa được tổng hợp',
          summary: 'Các kênh này chưa được cộng vào chỉ tiêu tương ứng. Bổ sung cách phân nhóm hoặc xác nhận bỏ qua trong Cấu hình.',
          impact: 'Chưa cộng các kênh dưới đây vào chỉ tiêu tổng hợp tương ứng.',
          action: 'Kiểm tra nhóm kênh trong Cấu hình và quyết định nhóm cần cộng hoặc bỏ qua.',
          tables: [{ title: 'Kênh chưa có nhóm', headers: ['Kênh', 'Số dòng'],
            rows: un.map(function (k) { return [k, unknown[k]]; }) }] });
    }
    var ut = Object.keys(unknownTopics);
    if (ut.length && warnings) {
      warn(warnings, 'warn', 'Có ' + ut.length + ' giá trị Topic Level 1 chưa gán dòng — ' +
        'không được cộng vào khối "Nội dung hỗ trợ"',
        ut.map(function (k) { return k + ' (' + unknownTopics[k] + ')'; }).join(' · ') +
        ' — thêm vào `summary.topicRows` hoặc `summary.ignoredTopics` trong Cấu hình.', null,
        { title: ut.length + ' loại nội dung chưa được tổng hợp',
          summary: 'Các nội dung này chưa được cộng vào mục “Nội dung hỗ trợ”. Bổ sung cách phân nhóm hoặc xác nhận bỏ qua trong Cấu hình.',
          impact: 'Chưa cộng các loại nội dung dưới đây vào mục “Nội dung hỗ trợ”.',
          action: 'Kiểm tra nhóm nội dung trong Cấu hình và quyết định nhóm cần cộng hoặc bỏ qua.',
          tables: [{ title: 'Nội dung chưa có nhóm', headers: ['Nội dung trong nguồn', 'Số dòng'],
            rows: ut.map(function (k) { return [k, unknownTopics[k]]; }) }] });
    }

    return { counts: counts, distinct: distinct, unknownChannels: unknown,
             topics: topics, unknownTopics: unknownTopics };
  }

  /* =====================================================================
   * Soát bảng `mail SF` — tên trùng làm VLOOKUP gán case sai người
   * ===================================================================== */

  /**
   * VLOOKUP luôn lấy dòng khớp ĐẦU TIÊN. Nếu `mail SF` có một tên ở hai dòng với
   * hai email khác nhau thì email thứ hai không bao giờ được chọn — nhưng những ô
   * đã tính từ trước khi thêm dòng trùng vẫn giữ giá trị cũ, làm file tự mâu thuẫn.
   *
   * Trả { total, duplicates: [{name, entries}], affected: {email đã chuẩn hoá: true} }
   */
  function mailSfCorrections(masterWb, cfg) {
    var spec = cfg.mailSf;
    var ws = masterWb && spec && masterWb.Sheets[spec.sheet];
    if (!ws) return [];
    var rows = P.sheetToMatrix(ws).rows, edits = [];
    (spec.nameCorrections || []).forEach(function (rule) {
      /* Tên đích đã gắn email khác: để người dùng xử lý, không tự chọn. */
      var conflict = rows.slice(spec.startRow - 1).some(function (row) {
        return P.normName(row[spec.nameColumn]) === P.normName(rule.to) &&
          P.normText(row[spec.emailColumn]) !== P.normText(rule.email);
      });
      if (conflict) return;
      rows.forEach(function (row, i) {
        if (i < spec.startRow - 1 || P.normName(row[spec.nameColumn]) !== P.normName(rule.from) ||
            P.normText(row[spec.emailColumn]) !== P.normText(rule.email)) return;
        edits.push({ row: i + 1, col: spec.nameColumn, value: rule.to, type: 'text',
          from: row[spec.nameColumn], email: row[spec.emailColumn] });
      });
    });
    return edits;
  }

  function inspectMailSf(masterWb, cfg, warnings, fileName) {
    var spec = cfg.mailSf;
    if (!spec || !masterWb) return null;
    var ws = masterWb.Sheets[spec.sheet];
    if (!ws) {
      warn(warnings, 'warn', 'File báo cáo tổng không có sheet "' + spec.sheet + '"',
        'Tool dùng bảng email trong Cấu hình thay thế.', null,
        { title: 'File tổng chưa có bảng tên và email',
          summary: 'Tool đang dùng bảng email trong Cấu hình. Kiểm tra bảng này còn đúng với danh sách nhân viên hiện tại.' });
      return null;
    }

    var m = P.sheetToMatrix(ws);
    var corrections = mailSfCorrections(masterWb, cfg);
    corrections.forEach(function (edit) {
      m.rows[edit.row - 1][spec.nameColumn] = edit.value;
    });
    if (corrections.length) warn(warnings, 'info', 'Đã chuẩn hóa tên trong mail SF theo quy tắc đã cấu hình',
      corrections.map(function (e) { return 'dòng ' + e.row + ': ' + e.from + ' → ' + e.value + ' (' + e.email + ')'; }).join(' · ') +
      '. Thay đổi được ghi vào file xuất; file nền giữ nguyên.', null,
      { title: 'Đã chuẩn hóa ' + corrections.length + ' tên trong bảng email',
        summary: 'Áp dụng quy tắc sửa tên đã cấu hình. File xuất dùng tên đã sửa; file nền không thay đổi.',
        tables: [{ title: 'Tên đã được chuẩn hóa', headers: ['Dòng', 'Tên trong nguồn', 'Tên dùng trong báo cáo', 'Email'],
          rows: corrections.map(function (e) { return [e.row, e.from, e.value, e.email]; }) }] });
    var first = {}, groups = {}, total = 0;
    for (var r = spec.startRow - 1; r < m.rows.length; r++) {
      var row = m.rows[r] || [];
      var name = P.cellToString(row[spec.nameColumn]).trim();
      var mail = P.cellToString(row[spec.emailColumn]).trim();
      if (!name || !mail) continue;
      total++;
      var key = P.normName(name);
      if (!first[key]) { first[key] = { row: r + 1, email: mail, name: name }; continue; }
      (groups[key] = groups[key] || { name: name, entries: [first[key]] })
        .entries.push({ row: r + 1, email: mail, name: name });
    }

    var dupKeys = Object.keys(groups);
    var affected = {};
    var duplicates = dupKeys.map(function (k) {
      var g = groups[k];
      g.entries.forEach(function (e) { affected[P.normText(e.email)] = true; });
      return g;
    });

    /* Bảng tra tên -> email lấy luôn từ file nền: nó mới là bản chuẩn đang dùng,
       còn bảng seed trong cấu hình thì cũ dần theo tháng. Giữ dòng khớp ĐẦU TIÊN
       cho đúng ngữ nghĩa VLOOKUP. */
    var table = {};
    Object.keys(first).forEach(function (k) { table[first[k].name] = first[k].email; });

    if (duplicates.length) {
      warn(warnings, 'error',
        'Sheet "' + spec.sheet + '" có ' + duplicates.length + ' tên bị ghi trùng — ' +
        'chưa đủ thông tin để phân công case chính xác',
        duplicates.map(function (g) {
          return g.name + ': ' +
            g.entries.map(function (e) { return 'dòng ' + e.row + ' → ' + e.email; }).join(' · ');
        }).join(' | ') +
        '. VLOOKUP chỉ lấy dòng ĐẦU TIÊN, nên các email sau không bao giờ nhận được case. ' +
        'Xác nhận email nhận toàn bộ case rồi sửa bảng tra đúng phân công. Nếu cần chia từng case, ' +
        'phải có thông tin phân công riêng; tool không tự chia đôi theo bản làm tay.',
        duplicates.reduce(function (acc, g) {
          g.entries.forEach(function (e, i) {
            acc.push(spot(fileName || 'File báo cáo tổng', spec.sheet, e.row, spec.nameColumn, e.name,
              i === 0
                ? 'Tên này bị ghi ' + g.entries.length + ' lần. Dòng này là dòng VLOOKUP lấy (' + e.email + ').'
                : 'Trùng với dòng ' + g.entries[0].row + ' (' + e.email + ') — dòng này không bao giờ nhận được case.'));
          });
          return acc;
        }, []), { title: duplicates.length + ' tên bị trùng trong bảng email',
          summary: 'Tool lấy email ở dòng đầu tiên nên có thể tính case nhầm người. Kiểm tra bảng “mail SF” và xác nhận dòng nào cần sửa.',
          impact: 'Tool chỉ lấy email ở dòng đầu tiên của mỗi tên trùng. Những email ở dòng sau không được nhận case qua bảng tra này.',
          action: 'Xác nhận email nhận toàn bộ case và sửa bảng tra đúng phân công. Bổ sung mã case → email nếu cần tách từng case, cùng bảng giờ từng người; tool không tự chia đôi.',
          tables: [{ title: 'Các tên trùng trong mail SF', headers: ['Tên', 'Dòng', 'Email', 'Cách tool đang dùng'],
            rows: duplicates.reduce(function (rows, g) {
              return rows.concat(g.entries.map(function (e, i) { return [g.name, e.row, e.email, i === 0 ? 'Dòng được tra đầu tiên' : 'Không được chọn']; }));
            }, []) }] });
    }
    return { total: total, duplicates: duplicates, affected: affected, table: table, corrections: corrections };
  }

  /* =====================================================================
   * KPI — dựng lại chuỗi Sheet5 -> Daily KPI Result -> KPI tháng
   * ===================================================================== */
  var B1_SUPPORTER = 4;   /* cột Supporter của khối 1 */
  var B2_EMAIL = 5;       /* cột Supporter email của khối 2 */
  var B3_EMAIL = 2;       /* cột Supporter Email của khối 3 */

  function toNum(v) {
    var n = P.normalizeHours(v);
    return typeof n === 'number' && isFinite(n) ? n : 0;
  }

  /**
   * Hiệu suất công việc từng người, tính lại đúng chuỗi công thức của file tổng.
   * Trả { byEmail: {email: {W, E, F, ...}}, team(emails) }.
   */
  function computeKpi(blocks, cfg) {
    var kc = cfg.kpi;
    if (!kc) return null;

    function setOf(list) {
      var o = {};
      (list || []).forEach(function (x) { o[P.normText(x)] = true; });
      return o;
    }
    var isCase = setOf(kc.caseChannels),
        isCmt = setOf(kc.commentChannels),
        isOther = setOf(kc.otherChannels);

    /* pivot 1 — đếm CHANNEL theo email, từ khối 1 */
    var p1 = {};
    blocks.block1.rows.forEach(function (row) {
      var e = P.normText(row[B1_SUPPORTER]);
      if (!e) return;
      var a = p1[e] || (p1[e] = { total: 0, cs: 0, cmt: 0, other: 0, ob: 0, dkm: 0, xl: 0, sets: 0 });
      var ch = P.normText(row[2]);
      a.total++;
      if (isCase[ch]) a.cs++;
      else if (isCmt[ch]) a.cmt++;
      else if (isOther[ch]) a.other++;
      else if (ch === 'outbound') a.ob++;
      else if (ch === 'đăng ký mới') a.dkm++;
      else if (ch === 'xử lý đăng ký mới') a.xl++;
      else if (ch === 'setsumekai') a.sets++;
    });

    /* pivot 2 — đếm case theo Supporter email, từ khối 2 */
    var p2 = {};
    blocks.block2.rows.forEach(function (row) {
      var e = P.normText(row[B2_EMAIL]);
      if (e) p2[e] = (p2[e] || 0) + 1;
    });

    /* pivot 3 — tổng giờ theo email, từ khối 3 */
    var p3 = {};
    blocks.block3.rows.forEach(function (row) {
      var e = P.normText(row[B3_EMAIL]);
      if (!e) return;
      var a = p3[e] || (p3[e] = {});
      kc.tracks.forEach(function (t) {
        a[t.key] = (a[t.key] || 0) + toNum(row[t.hoursColumn]);
      });
    });

    function supportCount(a) {
      var weight = typeof kc.setsumekaiWeight === 'number' ? kc.setsumekaiWeight : 45;
      return a.total - a.cs - a.cmt - a.other - (a.ob || 0) - (a.dkm || 0) - (a.xl || 0) + (a.sets || 0) * weight;
    }

    function forEmail(email, profile) {
      var e = P.normText(email);
      var a = p1[e] || { total: 0, cs: 0, cmt: 0, other: 0 };
      var hrs = p3[e] || {};
      var E = supportCount(a);   /* cùng công thức O của Sheet5 */
      var F = p2[e] || 0;
      var G = Math.max(0, E - F);
      var counts = { H: F * 1.5 + G, OB: a.ob || 0, 'case': a.cs, CMT: a.cmt,
        DKM: a.dkm || 0, XL: a.xl || 0, JP: a.other, SETS: a.sets || 0 };

      var numer = 0, denom = 0;
      kc.tracks.forEach(function (t, i) {
        var h = hrs[t.key] || 0;
        denom += h;
        var countKey = profile ? profile.counts[i] : t.count;
        var rate = profile ? profile.rates[i] : t.rate;
        if (h > 0) numer += (counts[countKey] || 0) / rate;
      });
      return {
        W: denom > 0 ? numer / denom : 0,
        E: E, F: F, G: G, H: counts.H, cases: a.cs, cmt: a.cmt,
        hours: denom, found: !!p1[e] || !!p3[e],
        /* Có việc hôm nay (phiên chat hoặc case) nhưng chưa chắc có giờ làm.
           `hasWork && !hours` = W ra 0 vì thiếu dòng trong "3. Other Tasks". */
        hasWork: a.total > 0 || F > 0
      };
    }

    /* Dòng tổng OP: cộng dồn số lượng và giờ TRƯỚC rồi mới chia,
       đúng như W14 = f(SUM(...)) chứ không phải trung bình các W. */
    function team(emails) {
      var cnt = { H: 0, OB: 0, 'case': 0, CMT: 0, DKM: 0 };
      var hrs = {};
      kc.tracks.forEach(function (t) { hrs[t.key] = 0; });
      emails.forEach(function (email) {
        var e = P.normText(email);
        var a = p1[e] || { total: 0, cs: 0, cmt: 0, other: 0 };
        var h = p3[e] || {};
        var E = supportCount(a);
        var F = p2[e] || 0;
        cnt.H += 0; /* H của dòng tổng tính từ SUM(E) và SUM(F), không phải tổng các H */
        cnt.__E = (cnt.__E || 0) + E;
        cnt.__F = (cnt.__F || 0) + F;
        cnt['case'] += a.cs;
        cnt.CMT += a.cmt;
        cnt.OB += a.ob || 0;
        cnt.DKM += a.dkm || 0;
        kc.tracks.forEach(function (t) { hrs[t.key] += (h[t.key] || 0); });
      });
      var G = Math.max(0, (cnt.__E || 0) - (cnt.__F || 0));
      cnt.H = (cnt.__F || 0) * 1.5 + G;

      var numer = 0, denom = 0;
      kc.tracks.forEach(function (t) {
        var h = hrs[t.key] || 0;
        denom += h;
        if (h > 0) numer += (cnt[t.count] || 0) / t.rate;
      });
      return denom > 0 ? numer / denom : 0;
    }

    return { forEmail: forEmail, team: team, p1: p1, p2: p2, p3: p3 };
  }

  /**
   * Gom mọi vị trí lỗi trong một danh sách cảnh báo lại theo TỆP.
   * Trả [{ file, sheets: { <tên sheet>: [spot...] }, count, levels }]
   * UI dùng để hiện "tệp nào, dòng nào"; phần xuất file dùng để đánh dấu ô.
   */
  function collectSpots(warningLists) {
    var byFile = {};
    (warningLists || []).forEach(function (list) {
      (list || []).forEach(function (w) {
        (w.spots || []).forEach(function (sp) {
          if (!sp || !sp.file) return;
          var f = byFile[sp.file] || (byFile[sp.file] = {
            file: sp.file, sheets: {}, count: 0, levels: {}, messages: {}
          });
          var arr = f.sheets[sp.sheet] || (f.sheets[sp.sheet] = []);
          arr.push({ row: sp.row, col: sp.col, value: sp.value, why: sp.why,
                     level: w.level, message: w.message });
          f.count++;
          f.levels[w.level] = (f.levels[w.level] || 0) + 1;
          f.messages[w.message] = (f.messages[w.message] || 0) + 1;
        });
      });
    });
    return Object.keys(byFile).map(function (k) { return byFile[k]; })
      .sort(function (a, b) { return b.count - a.count; });
  }

  /* =====================================================================
   * Soát danh sách nhân viên: ai có việc hôm nay mà chưa có dòng KPI
   * ===================================================================== */

  /**
   * Nhân viên mới làm xong một ngày nhưng chưa được thêm dòng ở "Daily KPI Result"
   * thì KPI của họ KHÔNG đi đâu cả — và không có dấu hiệu gì cho thấy đang thiếu.
   * Hàm này bắt đúng trường hợp đó, kèm tên (tra ngược từ `mail SF`) và số dòng
   * cần chèn, để người làm thêm một lần là xong.
   *
   * KHÔNG tự chèn dòng: hai sheet KPI có dòng tổng `=SUM(...)` ngay dưới và hàng
   * trăm công thức tra Sheet5; chèn dòng bằng cách sửa XML thì mọi tham chiếu phải
   * tự tay dịch lại, sai một chỗ là số liệu lệch mà không ai thấy. Đổi lại, báo
   * đúng chỗ cần thêm thì việc tay chỉ còn vài chục giây.
   */
  function checkRoster(input, cfg, masterWb, blocks, warnings) {
    var kc = cfg.kpi;
    if (!kc || !masterWb) return null;
    var ws = masterWb.Sheets[kc.daily.sheet];
    if (!ws) return null;

    /* --- ai có việc hôm nay --- */
    var active = {};
    function scan(rows, col, label) {
      (rows || []).forEach(function (r) {
        var e = P.cellToString(r[col]).trim();
        if (e.indexOf('@') === -1) return;
        var k = P.normText(e);
        var a = active[k] || (active[k] = { email: e, where: {} });
        a.where[label] = (a.where[label] || 0) + 1;
      });
    }
    scan(blocks.block1.rows, B1_SUPPORTER, 'phiên chat/CSKH');
    scan(blocks.block2.rows, B2_EMAIL, 'case Salesforce');
    scan(blocks.block3.rows, B3_EMAIL, 'giờ làm');

    /* --- ai đã có dòng KPI --- */
    var m = P.sheetToMatrix(ws);
    var roster = {}, lastRow = 0;
    var end = kc.daily.peopleEndRow || kc.daily.dataEndRow;
    for (var r = kc.daily.dataStartRow - 1; r < end; r++) {
      var e2 = P.cellToString((m.rows[r] || [])[kc.daily.emailColumn]).trim();
      if (e2.indexOf('@') !== -1) { roster[P.normText(e2)] = r + 1; lastRow = r + 1; }
    }
    if (!lastRow) return null;

    /* --- tra ngược email -> tên, từ bảng `mail SF` --- */
    var byEmail = {};
    var tbl = (input.emailTable && Object.keys(input.emailTable).length)
      ? input.emailTable : (cfg.supporterEmails || {});
    Object.keys(tbl).forEach(function (nm) { byEmail[P.normText(tbl[nm])] = nm; });
    /* Bảng tra case chỉ giữ email khớp đầu tiên. Danh sách cảnh báo giờ cần cả
       những email phía sau của tên trùng; không được báo nhầm là chưa có tên. */
    var mailSpec = cfg.mailSf, mailWs = mailSpec && masterWb.Sheets[mailSpec.sheet];
    if (mailWs) P.sheetToMatrix(mailWs).rows.slice(mailSpec.startRow - 1).forEach(function (row) {
      var email = P.normText(row[mailSpec.emailColumn]), name = P.cellToString(row[mailSpec.nameColumn]).trim();
      if (email.indexOf('@') !== -1 && name && !byEmail[email]) byEmail[email] = name;
    });

    /* Báo ngay khi Xử lý, không đợi người dùng tải file rồi mới thấy trong log.
       Chỉ kiểm tra người có dòng KPI; không tự lấy giờ của ngày khác để bù. */
    var noHours = blocks.kpi ? Object.keys(active).filter(function (k) {
      if (!roster[k]) return false;
      var perf = blocks.kpi.forEmail(active[k].email);
      return perf.hasWork && perf.hours <= 0;
    }) : [];
    if (noHours.length) {
      warn(warnings, 'warn', noHours.length + ' người có công việc nhưng thiếu giờ hợp lệ để tính KPI',
        noHours.map(function (k) {
          return (byEmail[k] ? byEmail[k] + ' — ' : '') + active[k].email;
        }).join(' · ') + '. KPI hiện bằng 0 vì tổng giờ dùng tính KPI không lớn hơn 0, ' +
        'không có nghĩa là những người này không làm việc. Kiểm tra email, ngày và các cột giờ ' +
        'trong file OP của ngày ' + blocks.reportISO + '. Tool không tự đổi ngày, ' +
        'lấy giờ ngày khác hoặc điền giờ theo bản chuẩn. Cảnh báo này không xác định được ô ' +
        'cần sửa nên không có vị trí tô vàng riêng.', null,
        { title: noHours.length + ' người chưa có giờ làm hợp lệ',
          summary: 'KPI đang bằng 0 dù có công việc. Kiểm tra ngày, email và giờ làm trong báo cáo nhân viên.',
          impact: 'KPI của những người dưới đây đang bằng 0 vì tổng giờ dùng tính KPI không lớn hơn 0; không có nghĩa là họ không làm việc.',
          action: 'Kiểm tra ngày, email và các cột giờ trong báo cáo nhân viên. Bổ sung đúng dữ liệu rồi chạy lại; không lấy giờ ngày khác để bù.',
          note: 'Ngày báo cáo: ' + blocks.reportISO + '. Chưa xác định được ô nguồn cần sửa nên cảnh báo này không có vị trí tô vàng riêng.',
          tables: [{ title: 'Người cần kiểm tra giờ làm', headers: ['Nhân viên', 'Email', 'Công việc ghi nhận', 'Giờ dùng tính KPI'],
            rows: noHours.map(function (k) {
              var a = active[k];
              return [byEmail[k] || 'Chưa có tên trong mail SF', a.email,
                Object.keys(a.where).map(function (label) { return a.where[label] + ' dòng ' + label; }).join('\n'),
                blocks.kpi.forEmail(a.email).hours];
            }) }] });

      /* Gợi ý vị trí cần kiểm tra, không coi dữ liệu lịch sử là lỗi và không
         kéo giờ của ngày khác sang. Chỉ xét người đang có việc mà thiếu giờ. */
      var candidates = [], wantedEmails = {};
      noHours.forEach(function (k) { wantedEmails[k] = true; });
      (input.ops || []).forEach(function (op) {
        var prev = null;
        ((op.other && op.other.otherDayRows) || []).forEach(function (rec) {
          var email = P.normText(rec.cells[B3_EMAIL + 2]);
          if (!wantedEmails[email]) { prev = null; return; }
          if (prev && prev.email === email && prev.serial === rec.serial && prev.last + 1 === rec.excelRow) {
            prev.last = rec.excelRow; prev.records.push(rec);
          } else {
            prev = { file: op.fileName, email: email, serial: rec.serial, first: rec.excelRow,
              last: rec.excelRow, records: [rec] };
            candidates.push(prev);
          }
        });
      });
      var nearest = {};
      candidates.forEach(function (c) {
        var distance = Math.abs(c.serial - blocks.reportSerial);
        nearest[c.email] = nearest[c.email] === undefined ? distance : Math.min(nearest[c.email], distance);
      });
      candidates = candidates.filter(function (c) { return Math.abs(c.serial - blocks.reportSerial) === nearest[c.email]; });
      if (candidates.length) {
        var dateColumn = cfg.op.other.copyColumns[0], dateLetter = XLSX.utils.encode_col(dateColumn);
        warn(warnings, 'warn', Object.keys(nearest).length + ' người thiếu giờ hôm nay nhưng có giờ khai ở ngày khác',
          'Chỉ hiển thị các dòng gần ngày báo cáo nhất. Có thể là dữ liệu lịch sử đúng; không kết luận nhân viên nhập sai và không tự đổi ngày.',
          candidates.reduce(function (list, c) {
            return list.concat(c.records.map(function (r) { return spot(c.file, cfg.op.other.sheet,
              r.excelRow, dateColumn, r.cells[dateColumn],
              'Email này có công việc ngày ' + blocks.reportISO + ' nhưng dòng giờ đang ghi ' + P.serialToISO(c.serial) +
              '. Kiểm tra có quên sửa ngày khi copy không; nếu là giờ lịch sử đúng thì giữ nguyên. Tool không tự nhận dòng này.'); }));
          }, []), { title: 'Có giờ khai ở ngày khác, cần kiểm tra ngày',
            summary: 'Những người này thiếu giờ của ngày báo cáo. Kiểm tra các dòng gần nhất xem có quên sửa ngày khi copy không; tool chưa lấy các giờ này.',
            impact: 'KPI vẫn thiếu giờ ngày ' + blocks.reportISO + '. Các dòng dưới đây không được đưa vào báo cáo.',
            action: 'Chỉ sửa ngày nếu xác nhận đây là giờ của ngày đang báo cáo. Nếu là dữ liệu lịch sử đúng, giữ nguyên và bổ sung giờ còn thiếu.',
            tables: [{ title: 'Các dòng giờ gần ngày báo cáo nhất', headers: ['Email', 'File nguồn', 'Ngày đang ghi', 'Ô ngày cần kiểm tra'],
              rows: candidates.map(function (c) { return [c.email, c.file, P.serialToISO(c.serial),
                dateLetter + c.first + (c.last === c.first ? '' : ':' + dateLetter + c.last)]; }) }] });
      }
    }

    var missing = Object.keys(active).filter(function (k) { return !roster[k]; });
    if (!missing.length) return { missing: [], lastRow: lastRow };

    var items = missing.map(function (k) {
      var a = active[k];
      return { email: a.email, name: byEmail[k] || null,
               where: Object.keys(a.where).map(function (w) { return w + ' (' + a.where[w] + ')'; }).join(', ') };
    });

    warn(warnings, 'error',
      items.length + ' người có việc hôm nay nhưng CHƯA có dòng trong "' + kc.daily.sheet + '"',
      items.map(function (it) {
        return (it.name ? it.name + ' <' + it.email + '>' : it.email + ' (chưa có tên trong `mail SF`)') +
               ' — thấy ở: ' + it.where;
      }).join(' · ') +
      '. KPI của họ sẽ KHÔNG được tính. Cách thêm, làm một lần trong File báo cáo tổng: ' +
      'chèn dòng mới ở "' + kc.daily.sheet + '" ngay sau dòng ' + lastRow +
      ' (cột ' + XLSX.utils.encode_col(kc.daily.emailColumn) + ' là email, copy công thức từ ' +
      'một dòng nhân viên bình thường ở trên), thêm dòng tương ứng ở sheet KPI tháng, ' +
      'và thêm tên + email vào tab `mail SF` nếu chưa có.', null,
      { title: items.length + ' người chưa có dòng trong bảng KPI',
        summary: 'KPI của họ chưa được ghi vào báo cáo. Bổ sung dòng nhân viên và công thức trong bảng KPI ngày, tháng.',
        impact: 'Chưa có dòng nhận kết quả KPI cho những nhân viên dưới đây.',
        action: 'Thêm dòng sau dòng ' + lastRow + ' trong “' + kc.daily.sheet + '”, sao chép công thức từ dòng nhân viên bình thường. Thêm dòng tương ứng ở KPI tháng và tên/email trong mail SF nếu chưa có.',
        tables: [{ title: 'Nhân viên cần thêm vào bảng KPI', headers: ['Nhân viên', 'Email', 'Có dữ liệu tại'],
          rows: items.map(function (it) { return [it.name || 'Chưa có tên trong mail SF', it.email, it.where]; }) }] });

    return { missing: items, lastRow: lastRow };
  }

  /* =====================================================================
   * 6 dòng fanpage — đọc số Follow dán từ extension
   * ===================================================================== */

  /**
   * Extension "Facebook Fanpage Stats" chép ra đúng 6 dòng theo thứ tự
   * SBIR, SMILES, DCOM, trong đó 3 dòng Like để trống:
   *
   *     (trống)
   *     431722
   *     (trống)
   *     149874
   *     (trống)
   *     393729
   *
   * Nhưng người dùng có thể dán thiếu dòng trống, dán 3 số, hoặc gõ tay kèm
   * tên trang. Nhận hết các dạng đó, miễn là suy ra được 3 số theo đúng thứ tự.
   *
   * Trả { follows: {KEY: số}, order: [KEY...], warnings: [...] } hoặc null nếu
   * ô dán còn trống.
   */
  function parseFanpage(text, cfg, warnings) {
    var fc = cfg.fanpage;
    if (!fc || !fc.pages || !fc.pages.length) return null;
    var raw = P.cellToString(text);
    if (!raw.trim()) return null;

    var keys = fc.pages.map(function (p) { return p.key; });
    var lines = raw.split(/\r?\n/).map(function (l) { return l.trim(); });

    function toNum(s) {
      /* 431.722 / 431,722 / 431 722 / 431722 — dấu phân cách hàng nghìn kiểu nào cũng nhận */
      var t = String(s).replace(/[.,\s ]/g, '');
      if (!/^\d+$/.test(t)) return null;
      var n = Number(t);
      return isFinite(n) && n >= 100 ? n : null;
    }

    var follows = {}, found = 0;

    /* Dạng 1: có tên trang trên cùng dòng với số — tin cậy nhất, không phụ thuộc thứ tự */
    lines.forEach(function (line) {
      if (!line) return;
      var up = P.normText(line).toUpperCase();
      for (var i = 0; i < keys.length; i++) {
        if (up.indexOf(keys[i]) === -1) continue;
        var m = line.match(/([\d][\d.,\s ]*)/g);
        if (!m) return;
        for (var j = m.length - 1; j >= 0; j--) {
          var n = toNum(m[j]);
          if (n !== null && follows[keys[i]] === undefined) { follows[keys[i]] = n; found++; return; }
        }
        return;
      }
    });

    /* Dạng 2: chỉ có số — gán theo thứ tự SBIR, SMILES, DCOM */
    if (found < keys.length) {
      var nums = [];
      lines.forEach(function (line) {
        if (!line) return;
        var n = toNum(line);
        if (n !== null) nums.push(n);
      });
      if (nums.length === keys.length) {
        follows = {};
        keys.forEach(function (k, i) { follows[k] = nums[i]; });
        found = keys.length;
      }
    }

    var missing = keys.filter(function (k) { return follows[k] === undefined; });
    if (missing.length && warnings) {
      warn(warnings, 'warn', 'Chưa đọc được số follow của ' + missing.join(', '),
        'Không đoán bừa thứ tự khi thiếu số — những dòng đó sẽ để trống trong báo cáo. ' +
        'Hãy bấm Copy trong extension rồi dán lại nguyên vẹn, hoặc gõ mỗi dòng một số ' +
        'theo đúng thứ tự ' + keys.join(' → ') + '.', null,
        { title: 'Thiếu số theo dõi fanpage: ' + missing.join(', '),
          summary: 'Các ô còn thiếu sẽ để trống. Lấy lại số từ extension hoặc dán đủ số cho từng trang.' });
    }
    if (!Object.keys(follows).length) return null;
    return { follows: follows, order: keys };
  }

  /* =====================================================================
   * Dựng 3 khối dán
   * ===================================================================== */
  function buildBlocks(input, cfg, reportSerial) {
    var warnings = [];
    var session = input.session, ccvn = input.ccvn, ops = input.ops || [];
    var adjustments = cfg.approvedAdjustments;
    var appliedAdjustments = [];
    var reportDate = P.serialToISO(reportSerial);
    var caseAdjustments = adjustments && adjustments.enabled ? (adjustments.cases || []).filter(function (a) { return a.date === reportDate; }) : [];
    var hourAdjustments = adjustments && adjustments.enabled ? (adjustments.hours || []).filter(function (a) { return a.date === reportDate; }) : [];
    var matchedCaseAdjustments = {};
    /* Số follow fanpage dán từ extension — không bắt buộc */
    var fanpage = parseFanpage(input.fanpageText, cfg, warnings);

    /* --- Khối 1: 1. CSKH --- */
    var rows1 = [];
    var feedbackSpots = [];
    function reviewTopic(value, file, sheet, row, col) {
      if (!cfg.sourceReviews || !cfg.sourceReviews.customerFeedbackTopic) return;
      var text = P.normText(value);
      if (/^12\s*\./.test(text) || /đánh giá (khách hàng|kh)(?:\b|\s|[(/]|$)/.test(text)) {
        feedbackSpots.push(spot(file, sheet, row, col, value,
          'SBI yêu cầu kiểm tra khi chọn “Đánh giá khách hàng”. Giữ nội dung nguồn; xác nhận chủ đề đúng rồi sửa nguồn, không tự đổi sang “Gặp nhân viên SBI”.'));
      }
    }
    if (session) {
      session.staging.rows.forEach(function (s, i) {
        rows1.push([null, reportSerial].concat(s, ['', '']));
        var topic = cfg.session.stagingColumns[7];
        reviewTopic(s[7], session.fileName, session.sheetName,
          (session.keptRowNumbers || [])[i] || null,
          topic ? P.findColumn(session.headers, topic.from) : -1);
      });
    }
    var copy = cfg.op.cskh.copyColumns;
    var blank1 = blankSet(cfg.op.cskh.blankColumns);
    var dropped1 = {};
    ops.forEach(function (op) {
      op.cskh.rows.forEach(function (rec) {
        var vals = copy.map(function (ci) {
          if (blank1[ci]) { noteBlanked(dropped1, ci, rec.cells[ci]); return ''; }
          var v = rec.cells[ci];
          return v === undefined || v === null ? '' : v;
        });
        rows1.push([null, reportSerial].concat(vals));
        reviewTopic(vals[7], op.fileName, cfg.op.cskh.sheet, rec.excelRow, copy[7]);
      });
    });
    warnBlanked(warnings, dropped1, cfg.op.cskh.sheet);
    if (feedbackSpots.length) warn(warnings, 'warn', feedbackSpots.length + ' dòng chọn Đánh giá khách hàng cần kiểm tra',
      'Theo xác nhận SBI trong docs/VẤN ĐỀ.docx, mục này về cơ bản không chọn. Tool giữ nguyên chủ đề và số đếm cho đến khi nguồn được sửa.',
      feedbackSpots, { title: feedbackSpots.length + ' dòng chọn “Đánh giá khách hàng”',
        summary: 'SBI yêu cầu kiểm tra chủ đề này. Tool chưa đổi nội dung; hãy xác nhận và sửa đúng chủ đề trong file nguồn.',
        impact: 'Chủ đề nguồn vẫn được giữ khi tổng hợp, không tự chuyển sang “Gặp nhân viên SBI”.',
        action: 'Kiểm tra nội dung từng phiên tại các ô bên dưới, sửa chủ đề đúng trong nguồn rồi chạy lại.' });
    rows1.forEach(function (row, i) {
      row[0] = i + 1;
      while (row.length < BLOCK1_HEADERS.length) row.push('');
    });

    /* --- Khối 2: 2. SF Case info --- */
    var rows2 = [];
    var missingEmails = {};
    /* đếm số case đã đổi tên cho từng luật — panel Đổi tên hiển thị lại cho người dùng */
    var aliasUsed = {}, aliasHits = {};
    /* tên người lập case thấy trong export, và tên có trong bảng `mail SF`:
       panel dùng làm gợi ý nhập, khỏi gõ sai chính tả */
    var namesSeen = {}, namesInMailSf = [];
    if (ccvn) {
      var cols = cfg.ccvn.caseColumns.map(function (cc) {
        var idx = P.findColumn(ccvn.headers, cc.from);
        if (idx === -1) warn(warnings, 'error', 'Khối 2: không tìm thấy cột nguồn "' + cc.from + '"', null, null,
          { title: 'Thiếu cột dữ liệu Salesforce',
            summary: 'Cột tương ứng trong kết quả sẽ trống. Kiểm tra tên cột trong file Salesforce và cấu hình.' });
        return idx;
      });
      var nameIdx = P.findColumn(ccvn.headers, cfg.ccvn.supporterNameColumn);
      var nameOutIdx = cfg.ccvn.caseColumns.length - 1; /* cột Requested By trong khối */
      /* Ưu tiên bảng `mail SF` đọc từ file nền — bảng trong cấu hình chỉ là dự phòng */
      var src = (input.emailTable && Object.keys(input.emailTable).length)
        ? input.emailTable : cfg.supporterEmails;
      var emails = {};
      Object.keys(src).forEach(function (k) { emails[P.normName(k)] = src[k]; });
      namesInMailSf = Object.keys(src);
      var aliases = {};
      Object.keys(cfg.requestedByAliases || {}).forEach(function (k) {
        aliases[P.normName(k)] = cfg.requestedByAliases[k];
      });
      aliasUsed = {}; aliasHits = {};
      /* Cắt ký tự thừa cuối tên (vd dấu chấm) — ảnh hưởng cả việc tra email */
      var trimRe = null;
      if (cfg.ccvn.nameTrimPattern) {
        try { trimRe = new RegExp(cfg.ccvn.nameTrimPattern); }
        catch (e) { warn(warnings, 'warn', 'nameTrimPattern không hợp lệ, bỏ qua', null, null,
          { title: 'Chưa áp dụng được quy tắc làm sạch tên',
            summary: 'Cấu hình quy tắc chưa hợp lệ nên tên được giữ nguyên. Kiểm tra lại quy tắc trong Cấu hình.' }); }
      }

      /* Cột `Created By` của bảng đã bỏ cột ứng với cột nào trong FILE GỐC */
      var nameSrcCol = (ccvn.srcCol && nameIdx >= 0) ? ccvn.srcCol[nameIdx] : nameIdx;
      var missSpots = [];

      ccvn.dataRows.forEach(function (row, rowIdx) {
        var vals = cols.map(function (ci, k) {
          if (ci === -1) return '';
          var v = row[ci];
          if (v === undefined || v === null) return '';
          /* cột khai báo format: đưa về đúng dạng file tổng đang dùng */
          if (cfg.ccvn.caseColumns[k].format === 'datetime') return P.normalizeDateTime(v);
          return v;
        });
        var rawName = nameIdx === -1 ? '' : P.cellToString(row[nameIdx]).trim();
        if (trimRe) rawName = rawName.replace(trimRe, '');
        if (rawName) namesSeen[rawName] = (namesSeen[rawName] || 0) + 1;
        /* Tên tài khoản SF -> tên OP thật; email tra theo tên đã đổi */
        var alias = aliases[P.normName(rawName)];
        var effName = alias || rawName;
        var approved = caseAdjustments.filter(function (a) { return String(a.caseNumber) === P.cellToString(vals[0]).trim(); });
        if (approved.length > 1) throw new Error('Nhiều điều chỉnh cùng Case Number ' + vals[0] + ' ngày ' + reportDate);
        if (approved.length) {
          var change = approved[0];
          if (P.normName(change.from) !== P.normName(rawName)) throw new Error('Điều chỉnh case ' + vals[0] + ' không khớp người lập trong nguồn; cần kiểm tra lại');
          if (!emails[P.normName(change.to)]) throw new Error('Điều chỉnh case ' + vals[0] + ': không có email của ' + change.to + ' trong mail SF');
          if (matchedCaseAdjustments[change.caseNumber]) throw new Error('Case Number bị trùng trong nguồn: ' + change.caseNumber);
          matchedCaseAdjustments[change.caseNumber] = true;
          effName = change.to;
          appliedAdjustments.push({kind:'case',date:reportDate,key:change.caseNumber,from:rawName,to:effName,source:adjustments.source});
        }
        if (alias && !approved.length) {
          var key = rawName + ' → ' + alias;
          aliasUsed[key] = (aliasUsed[key] || 0) + 1;
          /* đếm thêm theo tên đã chuẩn hoá, để panel tra đúng luật trong cấu hình
             dù export viết hoa/thường khác nhau */
          var nk = P.normName(rawName);
          aliasHits[nk] = (aliasHits[nk] || 0) + 1;
        }
        if (nameOutIdx >= 0) vals[nameOutIdx] = effName;

        var email = emails[P.normName(effName)];
        if (!email && effName) {
          missingEmails[effName] = (missingEmails[effName] || 0) + 1;
          missSpots.push(spot(ccvn.fileName, ccvn.sheetName,
            (ccvn.dataRowNumbers || [])[rowIdx] || null, nameSrcCol, rawName,
            'Tên này không có trong bảng `mail SF` của File báo cáo tổng' +
            (alias ? ' (sau khi đổi thành "' + alias + '")' : '') +
            ' nên case không tra được email supporter. Xác nhận người phụ trách trước khi bổ sung; ' +
            'tài khoản dùng chung cần xác nhận tên người được gán cho ngày báo cáo hoặc phân công từng case.'));
          email = '';
        }
        rows2.push(vals.concat([email || '', reportSerial]));
      });

      var aliasKeys = Object.keys(aliasUsed);
      if (aliasKeys.length) {
        warn(warnings, 'info', 'Đã đổi tên người lập case theo bảng `requestedByAliases`',
          aliasKeys.map(function (k) { return k + ' (' + aliasUsed[k] + ' case)'; }).join(' · ') +
          ' — kiểm tra lại nếu quy tắc này đã thay đổi.', null,
          { title: 'Đã đổi người nhận case theo cấu hình',
            summary: 'Theo SBI, gán tên ai thì toàn bộ case của tài khoản được tính cho người đó. Tool không tự chia đôi case.',
            action: 'Kiểm tra tên được gán cho ngày báo cáo trong “Đổi tên người lập case”. Quy tắc này không chuyển giờ làm hoặc lượt chat.',
            tables: [{ title: 'Quy tắc đổi tên đã áp dụng', headers: ['Tên nguồn → tên nhận case', 'Số case'],
              rows: aliasKeys.map(function (k) { return [k, aliasUsed[k]]; }) }] });
      }
    }
    var missNames = Object.keys(missingEmails);
    if (!missSpots) missSpots = [];
    if (missNames.length) {
      warn(warnings, 'error', missNames.length + ' tên không tra được email trong bảng `mail SF`',
        missNames.map(function (n) { return n + ' (' + missingEmails[n] + ' dòng)'; }).join(', ') +
        ' — kiểm tra tên và email trong tab ẩn `mail SF` của File báo cáo tổng. ' +
        'Chỉ bổ sung khi đã xác nhận đúng người phụ trách. Nếu đây là tài khoản dùng chung, ' +
        'xác nhận người được gán cho ngày báo cáo hoặc phân công từng case. Không mặc định CC-VN luôn là Thủy; tài khoản mượn có thể đổi người dùng.',
        missSpots, { title: missNames.length + ' tên chưa tìm được email',
          summary: 'Chưa có email để tính case. Xác nhận người dùng tài khoản trong ngày rồi gán tên, hoặc bổ sung phân công từng case.',
          impact: 'Email supporter của các case liên quan đang trống, nên chưa được tính cho nhân viên.',
          action: 'Xác nhận người đang mượn tài khoản rồi thêm vào “Đổi tên người lập case”. Tool tính toàn bộ case cho tên được gán, không tự chia đôi; không mặc định CC-VN là Thủy.',
          tables: [{ title: 'Tên chưa có email', headers: ['Tên cần đối chiếu', 'Số case', 'Kết quả tra email'],
            rows: missNames.map(function (n) { return [n, missingEmails[n], 'Chưa tìm thấy']; }) }] });
    }

    /* --- Khối 3: 3. Other Tasks & Working time --- */
    var rows3 = [];
    var copy3 = cfg.op.other.copyColumns;
    /* Cấu hình v1 đã lưu trong trình duyệt cũng phải nhận đủ giờ ở cột M. */
    if (JSON.stringify(copy3) === '[2,3,4,5,6,7,8,9,10,11]') copy3 = copy3.concat([12]);
    var blank3 = blankSet(cfg.op.other.blankColumns);
    var zeroBlank3 = blankSet(cfg.op.other.zeroAsBlankColumns);
    var dropped3 = {}, convertedHours = [], invalidHours = [], blankedZeros = [], reviewedHours = [];
    ops.forEach(function (op) {
      op.other.rows.forEach(function (rec) {
        var vals = copy3.map(function (ci) {
          if (blank3[ci]) { noteBlanked(dropped3, ci, rec.cells[ci]); return ''; }
          var v = rec.cells[ci];
          return v === undefined || v === null ? '' : v;
        });
        vals[0] = reportSerial; /* cột DATE luôn ghi serial ngày báo cáo */
        /* G:K và M là giờ; L là nhãn công việc, không được chuyển thành số. */
        [4, 5, 6, 7, 8, 10].forEach(function (ci) {
          var old = vals[ci], value = P.normalizeHours(old);
          if (typeof old === 'string' && typeof value === 'number') {
            vals[ci] = value;
            convertedHours.push(spot(op.fileName, cfg.op.other.sheet, rec.excelRow,
              copy3[ci], old, 'Giờ lưu dạng chữ đã chuyển thành số ' + value));
          } else if (!P.isBlank(old) && (typeof old !== 'number' || !isFinite(old))) {
            invalidHours.push(spot(op.fileName, cfg.op.other.sheet, rec.excelRow,
              copy3[ci], old, 'Không đọc được giờ; giữ nguyên nội dung và không cộng như một số. Sửa ô nguồn rồi chạy lại.'));
          }
          if (zeroBlank3[copy3[ci]] && value === 0) {
            vals[ci] = '';
            blankedZeros.push(spot(op.fileName, cfg.op.other.sheet, rec.excelRow,
              copy3[ci], old, 'Hiển thị số 0 ở cột giờ việc khác thành ô trống theo DataQuyTac; tổng giờ không đổi.'));
          }
        });
        ((cfg.sourceReviews && cfg.sourceReviews.hours) || []).forEach(function (rule) {
          if (rule.date !== reportDate || P.normText(vals[2]) !== P.normText(rule.email)) return;
          if (P.normalizeHours(rec.cells[rule.column]) !== rule.from) return;
          reviewedHours.push({ rule: rule, file: op.fileName, row: rec.excelRow,
            value: rec.cells[rule.column] });
        });
        rows3.push(vals);
      });
    });
    if (reviewedHours.length) warn(warnings, 'warn', reviewedHours.length + ' ô giờ đã được SBI xác nhận cần sửa nguồn',
      'Đây là lỗi nhập liệu đã được xác nhận trong docs/VẤN ĐỀ.docx, không phải quy tắc tự bù giờ cho mọi ngày.',
      reviewedHours.map(function (x) { return spot(x.file, cfg.op.other.sheet, x.row, x.rule.column, x.value, x.rule.reason); }),
      { title: reviewedHours.length + ' ô giờ cần sửa theo xác nhận SBI',
        summary: 'Số giờ nguồn đã được xác nhận nhập sai. Tool vẫn giữ nguồn, không tự lấy số trong bản làm tay để thay vào.',
        impact: 'KPI vẫn tính theo số giờ đang có trong nguồn cho đến khi bạn sửa và nạp lại.',
        action: 'Xác nhận số giờ đúng với nhân viên, sửa và lưu file OP rồi chạy lại. Không áp số đã đối chiếu cho ngày hoặc người khác.',
        tables: [{ title: 'Giờ SBI yêu cầu kiểm tra', headers: ['Ngày', 'Email', 'Ô nguồn', 'Giờ đang giữ', 'Giờ trong bản đã đối chiếu'],
          rows: reviewedHours.map(function (x) { return [x.rule.date, x.rule.email,
            XLSX.utils.encode_col(x.rule.column) + x.row, x.value, x.rule.reference]; }) }] });
    warnBlanked(warnings, dropped3, cfg.op.other.sheet);
    if (blankedZeros.length) warn(warnings, 'info', 'Đã để trống ' + blankedZeros.length + ' ô có 0 giờ việc khác',
      'Theo quy tắc trình bày DataQuyTac. Không đổi tổng giờ, KPI hoặc file OP gốc.', blankedZeros,
      { title: blankedZeros.length + ' ô 0 giờ việc khác được để trống',
        summary: 'Chỉ đổi cách hiển thị trong báo cáo. Tổng giờ và file nguồn vẫn giữ nguyên.' });
    if (convertedHours.length) warn(warnings, 'info', 'Đã chuyển ' + convertedHours.length + ' ô giờ dạng chữ thành số',
      'Chuẩn hóa trước khi cộng tổng và tính KPI; không sửa file OP gốc.', convertedHours,
      { title: 'Đã đọc được ' + convertedHours.length + ' ô giờ lưu dạng chữ',
        summary: 'Tool đã chuyển thành số để tính giờ và KPI. Không cần sửa nếu số giờ đã đúng; file nguồn không thay đổi.' });
    if (invalidHours.length) warn(warnings, 'warn', invalidHours.length + ' ô giờ không phải số hợp lệ',
      'Không tự đoán đơn vị, dấu phẩy hay giờ:phút. Các ô này không được tính vào tổng giờ/KPI.', invalidHours,
      { title: invalidHours.length + ' ô giờ chưa đúng dạng số',
        summary: 'Các ô này chưa được cộng vào giờ làm và KPI. Nhập số giờ, ví dụ 6.5, thay cho chữ hoặc giờ:phút rồi chạy lại.',
        impact: 'Các ô này được giữ nguyên nội dung nhưng không được cộng như số khi tính tổng giờ và KPI.',
        action: 'Kiểm tra giá trị trong bảng bên dưới. Nhập đúng số giờ vào file nguồn rồi chạy lại; không chỉ xóa ký tự khi chưa rõ ý nghĩa.' });
    var bs = cfg.op.other.blockSize;
    hourAdjustments.forEach(function (change) {
      var candidates = [];
      rows3.forEach(function (row, i) {
        if (i % bs === 0 && P.normText(row[2]) === P.normText(change.email)) candidates.push(row);
      });
      if (!candidates.length) {
        warn(warnings, 'error', 'Thiếu khối giờ cần đối chiếu: ' + change.email,
          'Chưa áp dụng điều chỉnh đã xác nhận cho ' + reportDate + '; nạp đủ file OP trước khi xuất báo cáo.', null,
          { title: 'Thiếu giờ làm để áp dụng điều chỉnh',
            summary: 'Chưa áp dụng được điều chỉnh đã xác nhận. Nạp đủ báo cáo nhân viên cho ngày đang làm.' });
        return;
      }
      if (candidates.length !== 1) throw new Error('Điều chỉnh giờ: cần đúng một khối của ' + change.email + ' ngày ' + reportDate);
      var row = candidates[0], old = row[change.column];
      if (!Number.isInteger(change.column) || change.column < 4 || change.column > 8 ||
          typeof old !== 'number' || !isFinite(old) || typeof change.from !== 'number' || !isFinite(change.from) ||
          typeof change.to !== 'number' || !isFinite(change.to) || change.to < 0 ||
          Math.abs(old - change.from) > 1e-9) throw new Error('Điều chỉnh giờ của ' + change.email + ' không khớp giá trị nguồn; cần kiểm tra lại');
      row[change.column] = change.to;
      appliedAdjustments.push({kind:'hours',date:reportDate,key:change.email,column:change.column,from:old,to:change.to,source:adjustments.source});
    });
    /* Cache F phải là cùng SUM(G:M) với công thức xuất ra, không dùng F đã cũ
       của nguồn. Excel SUM(range) bỏ qua text và ô trống. */
    if (bs) for (var start = 0; start < rows3.length; start += bs) {
      var totalHours = 0;
      for (var hr = start; hr < Math.min(start + bs, rows3.length); hr++) {
        for (var hc = 4; hc <= 10; hc++) {
          var hv = rows3[hr][hc];
          if (typeof hv === 'number' && isFinite(hv)) totalHours += hv;
        }
      }
      rows3[start][3] = totalHours;
    }
    caseAdjustments.forEach(function (change) {
      if (!matchedCaseAdjustments[change.caseNumber]) warn(warnings, 'warn', 'Không tìm thấy case đã có điều chỉnh: ' + change.caseNumber,
        'Ngày ' + reportDate + '; không tạo thêm case. Kiểm tra tính đầy đủ của export.', null,
        { title: 'Không tìm thấy case cần điều chỉnh',
          summary: 'Điều chỉnh này chưa được áp dụng. Kiểm tra file Salesforce đã xuất đủ dữ liệu của ngày báo cáo chưa.' });
    });
    if (appliedAdjustments.length) warn(warnings, 'warn', 'Đã áp dụng ' + appliedAdjustments.length + ' điều chỉnh có xác nhận từ output chuẩn',
      'Nguồn: ' + adjustments.source + '. ' + appliedAdjustments.map(function (a) { return a.key + ': ' + a.from + ' → ' + a.to; }).join(' · ') +
      '. Chỉ áp dụng đúng ngày và khóa dữ liệu; có thể tắt approvedAdjustments.enabled trong Cấu hình.', null,
      { title: 'Đã áp dụng ' + appliedAdjustments.length + ' điều chỉnh được xác nhận',
        summary: 'Chỉ áp dụng cho đúng ngày và dữ liệu đã duyệt. Xem chi tiết để biết case hoặc giờ làm nào được thay đổi.',
        note: 'Ngày áp dụng: ' + reportDate + '. Nguồn xác nhận: ' + adjustments.source,
        tables: [{ title: 'Các thay đổi đã áp dụng', headers: ['Loại điều chỉnh', 'Case / Email', 'Trước điều chỉnh', 'Sau điều chỉnh'],
          rows: appliedAdjustments.map(function (a) { return [a.kind === 'hours' ? 'Giờ làm' : 'Người nhận case', a.key, a.from, a.to]; }) }] });
    if (bs && rows3.length % bs !== 0) {
      warn(warnings, 'warn', 'Khối 3 có ' + rows3.length + ' dòng, không chia hết cho ' + bs +
        ' — kiểm tra lại khối dòng của từng supporter trong file OP', null, null,
        { title: 'Số dòng giờ làm chưa đúng mẫu',
          summary: 'Mỗi nhân viên cần đủ ' + bs + ' dòng. Kiểm tra xem báo cáo nhân viên có bị thiếu hoặc thừa dòng không.' });
    }

    /* SBI: case theo tên được gán; giờ vẫn theo nguồn. Chủ tài khoản có giờ
       không chứng minh luật sai, cũng không cho phép tự chia đôi case. */
    if (cfg.requestedByAliases && rows3.length) {
      var hoursEmails = {};
      rows3.forEach(function (r) {
        var e = P.normText(r[B3_EMAIL]);
        if (e) hoursEmails[e] = true;
      });
      var srcTable = (input.emailTable && Object.keys(input.emailTable).length)
        ? input.emailTable : (cfg.supporterEmails || {});
      var byName = {};
      Object.keys(srcTable).forEach(function (n) { byName[P.normName(n)] = srcTable[n]; });

      var clash = [], clashRows = [];
      Object.keys(cfg.requestedByAliases).forEach(function (src) {
        if (!cfg.requestedByAliases[src] || !aliasHits[P.normName(src)]) return;
        var em = byName[P.normName(src)];
        if (em && hoursEmails[P.normText(em)]) {
          clashRows.push([src, em, cfg.requestedByAliases[src]]);
          clash.push(src + ' <' + em + '> → đang gán hết case sang "' +
            cfg.requestedByAliases[src] + '"');
        }
      });
      if (clash.length) {
        warn(warnings, 'info',
          clash.length + ' tài khoản đã gán case nhưng giờ vẫn đứng tên chủ tài khoản',
          clash.join(' · ') + '. Theo SBI, case tính cho tên đã gán; tool không chia đôi hoặc chuyển giờ theo case. ' +
          'Nếu phân công thay đổi, cập nhật tên người nhận cho ngày báo cáo.', null,
          { title: 'Case và giờ đang ghi cho hai người khác nhau',
            summary: 'Case theo người đã gán, còn giờ giữ theo email trong OP. Có giờ của chủ tài khoản không đồng nghĩa quy tắc gán case bị sai.',
            impact: 'Giờ và lượt chat không tự chuyển sang người nhận case; nếu người nhận chưa có giờ KPI sẽ có cảnh báo riêng.',
            action: 'Kiểm tra tên được gán và bảng giờ của từng người. Không tự chia đôi case hoặc sao chép giờ giữa hai người.',
            tables: [{ title: 'Người nhận case và chủ tài khoản', headers: ['Tên tài khoản nguồn', 'Email chủ tài khoản', 'Tên đang nhận toàn bộ case'], rows: clashRows }] });
      }
    }

    var t = cfg.targets;
    var b1 = { rows: rows1 };
    var summary = computeSummary(b1, cfg, warnings);

    if (summary && cfg.summary && cfg.summary.manualRows) {
      var mr = cfg.summary.manualRows;
      var keys = Object.keys(mr);
      var nAuto = (cfg.summary.cells || []).length + (cfg.summary.topicRows || []).length;
      warn(warnings, 'info', 'Cột ngày trong "' + cfg.summary.sheet + '": tool điền được ' +
        nAuto + ' chỉ tiêu, còn ' + keys.length + ' chỉ tiêu phải nhập tay',
        keys.map(function (k) { return 'dòng ' + k + ' ' + mr[k]; }).join(' · ') +
        ' — số tồn đọng Salesforce lấy từ 4 báo cáo SF khác; Head Count, khung giờ và ' +
        'số lượng OP là số khai tay; 6 dòng fanpage do tool bên khác cung cấp. ' +
        'Sheet "Summarize" không cần điền: mọi ô ở đó đều là công thức rút từ sheet này.', null,
        { title: keys.length + ' chỉ tiêu cần bổ sung ngoài dữ liệu đang nạp',
        summary: 'Tool tự điền ' + nAuto + ' chỉ tiêu. Xem danh sách còn thiếu để nhập tay hoặc lấy từ nguồn khác.',
        note: 'Sheet: ' + cfg.summary.sheet + '. Kiểm tra các dòng dưới đây; một số số liệu lấy từ extension hoặc nguồn khác.',
        tables: [{ title: 'Chỉ tiêu cần kiểm tra thêm', headers: ['Dòng trong sheet', 'Chỉ tiêu'],
          rows: keys.map(function (k) { return [k, mr[k]]; }) }] });
    }

    return {
      fanpage: fanpage,
      /* cho panel Đổi tên người lập case */
      aliasUsed: aliasUsed,
      aliasHits: aliasHits,
      namesSeen: namesSeen,
      namesInMailSf: namesInMailSf,
      reportSerial: reportSerial,
      reportISO: P.serialToISO(reportSerial),
      warnings: warnings,
      summary: summary,
      kpi: computeKpi({ block1: b1, block2: { rows: rows2 }, block3: { rows: rows3 } }, cfg),
      block1: { key: 'block1', title: '1. CSKH', sheet: t.cskh.sheet, anchor: t.cskh.anchor,
                headers: BLOCK1_HEADERS, types: BLOCK1_TYPES, rows: rows1,
                note: (session ? session.staging.rows.length : 0) + ' dòng CHAT BOT + ' +
                      (rows1.length - (session ? session.staging.rows.length : 0)) + ' dòng từ file OP' },
      block2: { key: 'block2', title: '2. SF Case info', sheet: t.sfcase.sheet, anchor: t.sfcase.anchor,
                headers: (cfg.ccvn.caseColumns.map(function (x) { return x.out; }))
                          .concat(['Supporter email', 'Date']),
                types: BLOCK2_TYPES, rows: rows2,
                note: rows2.length + ' case (đã bỏ dòng Total)' },
      appliedAdjustments: appliedAdjustments,
      block3: { key: 'block3', title: '3. Other Tasks & Working time', sheet: t.other.sheet,
                anchor: t.other.anchor, headers: BLOCK3_HEADERS, types: BLOCK3_TYPES, rows: rows3,
                note: rows3.length + ' dòng' + (bs ? ' (' + Math.round(rows3.length / bs) + ' supporter)' : '') }
    };
  }

  return {
    detectKind: detectKind,
    processSession: processSession,
    processCcvn: processCcvn,
    processOpFile: processOpFile,
    detectReportDate: detectReportDate,
    computeSummary: computeSummary,
    computeKpi: computeKpi,
    inspectMailSf: inspectMailSf,
    mailSfCorrections: mailSfCorrections,
    parseFanpage: parseFanpage,
    collectSpots: collectSpots,
    checkRoster: checkRoster,
    buildBlocks: buildBlocks,
    BLOCK1_HEADERS: BLOCK1_HEADERS,
    BLOCK3_HEADERS: BLOCK3_HEADERS
  };
});
