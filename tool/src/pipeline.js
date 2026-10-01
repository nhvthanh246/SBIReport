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
  var BLOCK3_TYPES = ['date', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto'];
  var BLOCK3_HEADERS = ['DATE', 'TEAM', 'Supporter Email', 'TỔNG GIỜ LÀM',
    'Tổng số giờ đã hỗ trợ KH', 'Tổng số giờ đã Outbound', 'Tổng số giờ đã xử lý case',
    'Trả lời Comment', 'Đăng ký mới + webform + Gọi OB', 'Khác'];

  /* `spots` = danh sách ô gây ra cảnh báo, để UI chỉ đúng tệp/dòng/ô và để dựng
     bản Excel đánh dấu. Mỗi spot: { file, sheet, row (1-based), col (0-based|null),
     value, why }. Cắt bớt ở SPOT_CAP để một cảnh báo hàng nghìn dòng không làm
     phình bộ nhớ; `spotTotal` giữ lại con số thật. */
  var SPOT_CAP = 400;

  function warn(list, level, message, detail, spots) {
    var item = { level: level, message: message, detail: detail || null };
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
        'tool ghi trống. Nếu nay cột đó có nghĩa thật, bỏ nó khỏi `blankColumns` trong config.');
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
      if (idx === -1) warn(warnings, 'error', 'Luật lọc trỏ tới cột không tồn tại: "' + r.column + '"');
      return { rule: r, col: idx };
    }).filter(function (r) { return r.col !== -1; });

    /* Vị trí cột cho sheet staging */
    var stagingCols = c.stagingColumns.map(function (sc) {
      if (sc.from === 'const') return { spec: sc, col: -1 };
      var idx = P.findColumn(headers, sc.from);
      if (idx === -1) warn(warnings, 'error', 'Cột staging "' + sc.out + '" không tìm thấy nguồn "' + sc.from + '"');
      return { spec: sc, col: idx };
    });

    var kept = [], dropped = [];
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
      }
    }

    var stagingHeaders = c.stagingColumns.map(function (sc) { return sc.out; });
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
      warn(warnings, 'warn', 'File ' + fileName + ' không có sheet "' + spec.sheet + '" — bỏ qua');
      return { rows: [], scanned: 0, headers: [] };
    }
    var m = P.sheetToMatrix(ws);
    var headers = m.rows[spec.headerRow - 1] || [];
    var dateCol = P.findColumn(headers, spec.dateColumn);
    if (dateCol === -1) {
      warn(warnings, 'error', 'Sheet "' + spec.sheet + '" trong ' + fileName +
        ' không có cột ngày "' + spec.dateColumn + '"');
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

    var picked = [], scanned = 0, unparsed = [], swapped = [], dateOnly = 0;
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
      /* Ngày không khớp, nhưng đảo ngày<->tháng lại ra đúng ngày báo cáo:
         gần như chắc chắn là lỗi locale lúc nhập liệu. */
      if (P.swappedSerial(row[dateCol]) === reportSerial) {
        swapped.push({ excelRow: r + 1, cells: row, raw: P.cellToString(row[dateCol]) });
        if (acceptSwapped) picked.push({ excelRow: r + 1, cells: row, swapped: true });
        continue;
      }
      tally.otherDay++;
    }
    if (unparsed.length) {
      warn(warnings, 'error', 'Sheet "' + spec.sheet + '" trong ' + fileName +
        ': ' + unparsedRows.length + ' dòng có giá trị ngày không đọc được — bị bỏ qua',
        'Các giá trị gặp phải: ' + unparsed.slice(0, 10).join(' · '), unparsedRows);
    }
    if (noDateRows.length) {
      warn(warnings, 'warn', 'Sheet "' + spec.sheet + '" trong ' + fileName + ': ' +
        noDateRows.length + ' dòng có nội dung nhưng bỏ trống ô ngày — bị bỏ qua',
        null, noDateRows);
    }
    if (dateOnly > 50) {
      warn(warnings, 'warn', 'Sheet "' + spec.sheet + '" trong ' + fileName + ': bỏ qua ' +
        dateOnly + ' dòng chỉ có ô ngày, không có nội dung',
        'Thường do kéo ô ngày xuống quá nhiều dòng trống. Nên xoá bớt ở file nguồn cho nhẹ file.');
    }
    if (swapped.length) {
      var samples = {};
      swapped.forEach(function (s) {
        /* ô lưu dạng serial thì in ra ngày cho dễ đọc, kèm giá trị thô */
        var shown = /^\d+(\.\d+)?$/.test(s.raw)
          ? P.serialToISO(parseFloat(s.raw)) + ' (ô ghi ' + s.raw + ')'
          : s.raw;
        samples[shown] = (samples[shown] || 0) + 1;
      });
      warn(warnings, acceptSwapped ? 'warn' : 'error',
        'Sheet "' + spec.sheet + '" trong ' + fileName + ': ' + swapped.length +
        ' dòng bị ĐẢO ngày/tháng — ' +
        (acceptSwapped ? 'đã tính vào báo cáo' : 'ĐANG BỊ BỎ SÓT'),
        'Giá trị ghi trong file: ' +
        Object.keys(samples).map(function (k) { return k + ' (' + samples[k] + ' dòng)'; }).join(' · ') +
        (acceptSwapped ? '' : ' — bật "Nhận dòng bị đảo ngày/tháng" ở bước 2 để tính vào.'),
        swapped.map(function (sw) {
          return spot(fileName, spec.sheet, sw.excelRow, dateCol, sw.raw,
            'Ngày và tháng bị đảo. Sửa ô này thành ngày báo cáo, hoặc đổi định dạng cột sang dd/mm/yyyy.');
        }));
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
      'ĐẢO ngày đều có địa chỉ ô cụ thể ở cảnh báo riêng bên trên.');

    return { rows: picked, scanned: scanned, headers: headers, swapped: swapped,
             tally: tally, people: Object.keys(people) };
  }

  function processOpFile(wb, cfg, reportSerial, fileName) {
    var warnings = [];
    var acceptSwapped = !!(cfg.op && cfg.op.acceptSwappedDates);
    var cskh = readOpSheet(wb, cfg.op.cskh, reportSerial, fileName, warnings, acceptSwapped);
    var other = readOpSheet(wb, cfg.op.other, reportSerial, fileName, warnings, acceptSwapped);
    if (!cskh.rows.length && !other.rows.length) {
      warn(warnings, 'warn', 'File ' + fileName + ' không có dòng nào đúng ngày báo cáo');
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
        ' — thêm vào `summary.channelGroups` hoặc `summary.ignoredChannels` trong Cấu hình.');
    }
    var ut = Object.keys(unknownTopics);
    if (ut.length && warnings) {
      warn(warnings, 'warn', 'Có ' + ut.length + ' giá trị Topic Level 1 chưa gán dòng — ' +
        'không được cộng vào khối "Nội dung hỗ trợ"',
        ut.map(function (k) { return k + ' (' + unknownTopics[k] + ')'; }).join(' · ') +
        ' — thêm vào `summary.topicRows` hoặc `summary.ignoredTopics` trong Cấu hình.');
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
  function inspectMailSf(masterWb, cfg, warnings, fileName) {
    var spec = cfg.mailSf;
    if (!spec || !masterWb) return null;
    var ws = masterWb.Sheets[spec.sheet];
    if (!ws) {
      warn(warnings, 'warn', 'File báo cáo tổng không có sheet "' + spec.sheet + '"',
        'Tool dùng bảng email trong Cấu hình thay thế.');
      return null;
    }

    var m = P.sheetToMatrix(ws);
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
        'số case sẽ bị gán sai người',
        duplicates.map(function (g) {
          return g.name + ': ' +
            g.entries.map(function (e) { return 'dòng ' + e.row + ' → ' + e.email; }).join(' · ');
        }).join(' | ') +
        '. VLOOKUP chỉ lấy dòng ĐẦU TIÊN, nên các email sau không bao giờ nhận được case. ' +
        'Hãy xoá dòng thừa trong "' + spec.sheet + '", hoặc nếu đúng là hai người khác nhau ' +
        'thì phải thêm cách phân biệt vì tra theo tên không đủ.',
        duplicates.reduce(function (acc, g) {
          g.entries.forEach(function (e, i) {
            acc.push(spot(fileName || 'File báo cáo tổng', spec.sheet, e.row, spec.nameColumn, e.name,
              i === 0
                ? 'Tên này bị ghi ' + g.entries.length + ' lần. Dòng này là dòng VLOOKUP lấy (' + e.email + ').'
                : 'Trùng với dòng ' + g.entries[0].row + ' (' + e.email + ') — dòng này không bao giờ nhận được case.'));
          });
          return acc;
        }, []));
    }
    return { total: total, duplicates: duplicates, affected: affected, table: table };
  }

  /* =====================================================================
   * KPI — dựng lại chuỗi Sheet5 -> Daily KPI Result -> KPI tháng
   * ===================================================================== */
  var B1_SUPPORTER = 4;   /* cột Supporter của khối 1 */
  var B2_EMAIL = 5;       /* cột Supporter email của khối 2 */
  var B3_EMAIL = 2;       /* cột Supporter Email của khối 3 */

  function toNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = P.cellToString(v).trim();
    if (!s) return 0;
    var n = parseFloat(s);
    return isFinite(n) ? n : 0;
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
      var a = p1[e] || (p1[e] = { total: 0, cs: 0, cmt: 0, other: 0 });
      var ch = P.normText(row[2]);
      a.total++;
      if (isCase[ch]) a.cs++;
      else if (isCmt[ch]) a.cmt++;
      else if (isOther[ch]) a.other++;
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

    function forEmail(email) {
      var e = P.normText(email);
      var a = p1[e] || { total: 0, cs: 0, cmt: 0, other: 0 };
      var hrs = p3[e] || {};
      var E = a.total - a.cs - a.cmt - a.other;   /* HTKH */
      var F = p2[e] || 0;
      var G = Math.max(0, E - F);
      var counts = { H: F * 1.5 + G, OB: 0, 'case': a.cs, CMT: a.cmt, DKM: 0 };

      var numer = 0, denom = 0;
      kc.tracks.forEach(function (t) {
        var h = hrs[t.key] || 0;
        denom += h;
        if (h > 0) numer += (counts[t.count] || 0) / t.rate;
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
        var E = a.total - a.cs - a.cmt - a.other;
        var F = p2[e] || 0;
        cnt.H += 0; /* H của dòng tổng tính từ SUM(E) và SUM(F), không phải tổng các H */
        cnt.__E = (cnt.__E || 0) + E;
        cnt.__F = (cnt.__F || 0) + F;
        cnt['case'] += a.cs;
        cnt.CMT += a.cmt;
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
      'và thêm tên + email vào tab `mail SF` nếu chưa có.');

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
        'theo đúng thứ tự ' + keys.join(' → ') + '.');
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
    /* Số follow fanpage dán từ extension — không bắt buộc */
    var fanpage = parseFanpage(input.fanpageText, cfg, warnings);

    /* --- Khối 1: 1. CSKH --- */
    var rows1 = [];
    if (session) {
      session.staging.rows.forEach(function (s) {
        rows1.push([null, reportSerial].concat(s, ['', '']));
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
      });
    });
    warnBlanked(warnings, dropped1, cfg.op.cskh.sheet);
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
        if (idx === -1) warn(warnings, 'error', 'Khối 2: không tìm thấy cột nguồn "' + cc.from + '"');
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
        catch (e) { warn(warnings, 'warn', 'nameTrimPattern không hợp lệ, bỏ qua'); }
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
        if (alias) {
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
            ' nên case không tra được email supporter.'));
          email = '';
        }
        rows2.push(vals.concat([email || '', reportSerial]));
      });

      var aliasKeys = Object.keys(aliasUsed);
      if (aliasKeys.length) {
        warn(warnings, 'warn', 'Đã đổi tên người lập case theo bảng `requestedByAliases`',
          aliasKeys.map(function (k) { return k + ' (' + aliasUsed[k] + ' case)'; }).join(' · ') +
          ' — kiểm tra lại nếu quy tắc này đã thay đổi.');
      }
    }
    var missNames = Object.keys(missingEmails);
    if (!missSpots) missSpots = [];
    if (missNames.length) {
      warn(warnings, 'error', missNames.length + ' tên không tra được email trong bảng `mail SF`',
        missNames.map(function (n) { return n + ' (' + missingEmails[n] + ' dòng)'; }).join(', ') +
        ' — mở tab ẩn `mail SF` trong File báo cáo tổng, thêm dòng cho từng tên rồi chạy lại.',
        missSpots);
    }

    /* --- Khối 3: 3. Other Tasks & Working time --- */
    var rows3 = [];
    var copy3 = cfg.op.other.copyColumns;
    var blank3 = blankSet(cfg.op.other.blankColumns);
    var dropped3 = {};
    ops.forEach(function (op) {
      op.other.rows.forEach(function (rec) {
        var vals = copy3.map(function (ci) {
          if (blank3[ci]) { noteBlanked(dropped3, ci, rec.cells[ci]); return ''; }
          var v = rec.cells[ci];
          return v === undefined || v === null ? '' : v;
        });
        vals[0] = reportSerial; /* cột DATE luôn ghi serial ngày báo cáo */
        rows3.push(vals);
      });
    });
    warnBlanked(warnings, dropped3, cfg.op.other.sheet);
    var bs = cfg.op.other.blockSize;
    if (bs && rows3.length % bs !== 0) {
      warn(warnings, 'warn', 'Khối 3 có ' + rows3.length + ' dòng, không chia hết cho ' + bs +
        ' — kiểm tra lại khối dòng của từng supporter trong file OP');
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
        'Sheet "Summarize" không cần điền: mọi ô ở đó đều là công thức rút từ sheet này.');
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
    parseFanpage: parseFanpage,
    collectSpots: collectSpots,
    checkRoster: checkRoster,
    buildBlocks: buildBlocks,
    BLOCK1_HEADERS: BLOCK1_HEADERS,
    BLOCK3_HEADERS: BLOCK3_HEADERS
  };
});
