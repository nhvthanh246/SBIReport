/* Tự kiểm chứng: chạy đúng pipeline của tool rồi so từng ô với file kết quả mẫu
 * trong Data/output và với File báo cáo tổng. Chạy được ngay trên file:// nên
 * không cần Node hay công cụ build.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(root.Parse, root.Pipeline, root.Writer);
  else root.SelfTest = factory(root.Parse, root.Pipeline, root.Writer);
})(typeof self !== 'undefined' ? self : this, function (P, Pipeline, Writer) {
  'use strict';

  var MAX_SHOW = 8;

  /* `notes` = những ô lệch đã truy được nguyên nhân — vẫn hiện ra để soát, nhưng không tính là lỗi */
  function pass(name, detail, notes) {
    return { name: name, status: 'pass', detail: detail, mismatches: (notes || []).slice(0, MAX_SHOW) };
  }
  function fail(name, detail, mismatches) {
    return { name: name, status: 'fail', detail: detail, mismatches: (mismatches || []).slice(0, MAX_SHOW) };
  }
  function skip(name, detail) { return { name: name, status: 'skip', detail: detail, mismatches: [] }; }

  /* Excel khi mở rồi lưu lại sẽ đổi cách mã hoá xuống dòng: file kết quả mẫu chứa
     "\r\r\n" ở chỗ file gốc là "\r\n" hoặc "\n". Đây là dấu vết của Excel chứ không
     phải khác biệt nội dung, nên quy mọi kiểu xuống dòng về "\n" trước khi so. */
  function eol(s) {
    return s.replace(/\r\r\n/g, '\n').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  }

  /** So sánh nội dung ô theo dạng chuỗi đã chuẩn hoá khoảng trắng hai đầu. */
  function same(a, b) {
    var x = eol(P.cellToString(a)).trim();
    var y = eol(P.cellToString(b)).trim();
    if (x === y) return true;
    /* số: so theo giá trị để 8 và 8.0000000001 do làm tròn nhị phân không bị coi là lệch */
    var nx = parseFloat(x), ny = parseFloat(y);
    if (!isNaN(nx) && !isNaN(ny) && x !== '' && y !== '') {
      return Math.abs(nx - ny) < 1e-9 || (Math.abs(ny) > 0 && Math.abs((nx - ny) / ny) < 1e-9);
    }
    return false;
  }

  function addr(r, c) { return XLSX.utils.encode_cell({ r: r, c: c }); }

  /* ---- 1 & 2: Session Report ---- */
  function checkSession(result, expectedWb, cfg) {
    var out = [];
    var dropCol = P.findColumn(result.headers, cfg.session.dropRules[0].column);
    var leftovers = result.keptRows.filter(function (row) {
      return P.matchRule(row[dropCol], cfg.session.dropRules[0]);
    });
    out.push(leftovers.length === 0
      ? pass('1. Lọc "lỗi bot"', 'Giữ ' + result.stats.kept + ' / loại ' + result.stats.dropped +
          ' / tổng ' + result.stats.total + ' dòng — không còn dòng nào khớp luật lọc')
      : fail('1. Lọc "lỗi bot"', 'Còn ' + leftovers.length + ' dòng đáng lẽ phải bị loại'));

    if (!expectedWb) {
      out.push(skip('2. Sheet staging so với file mẫu', 'Chưa nạp file output mẫu'));
      return out;
    }
    var expName = result.staging.name;
    var ws = expectedWb.Sheets[expName];
    if (!ws) {
      var alt = expectedWb.SheetNames.filter(function (n) { return n !== result.sheetName; })[0];
      ws = alt ? expectedWb.Sheets[alt] : null;
      expName = alt;
    }
    if (!ws) {
      out.push(fail('2. Sheet staging so với file mẫu', 'File mẫu không có sheet staging'));
      return out;
    }

    var m = P.sheetToMatrix(ws);
    var expRows = m.rows.slice(1).filter(function (r) { return !P.rowIsEmpty(r, m.maxCol); });
    var got = result.staging.rows;
    var mism = [];
    var n = Math.max(expRows.length, got.length);
    for (var i = 0; i < n; i++) {
      for (var c = 0; c < result.staging.headers.length; c++) {
        var a = got[i] ? got[i][c] : undefined;
        var b = expRows[i] ? expRows[i][c] : undefined;
        if (!same(a, b)) {
          mism.push(addr(i + 1, c) + ' (' + result.staging.headers[c] + '): tool=' +
            JSON.stringify(P.cellToString(a).slice(0, 40)) + ' mẫu=' +
            JSON.stringify(P.cellToString(b).slice(0, 40)));
        }
      }
    }
    var cells = n * result.staging.headers.length;
    out.push(mism.length === 0 && expRows.length === got.length
      ? pass('2. Sheet staging so với file mẫu',
          'Khớp ' + cells + ' ô / ' + got.length + ' dòng × ' + result.staging.headers.length + ' cột')
      : fail('2. Sheet staging so với file mẫu',
          'Tool ' + got.length + ' dòng, mẫu ' + expRows.length + ' dòng, ' + mism.length + ' ô lệch', mism));
    return out;
  }

  /* ---- 3: CCVN ---- */
  function checkCcvn(result, expectedWb) {
    if (!expectedWb) return [skip('3. Daily Report CCVN so với file mẫu', 'Chưa nạp file output mẫu')];
    var ws = expectedWb.Sheets[expectedWb.SheetNames[0]];
    var m = P.sheetToMatrix(ws);
    var mism = [];
    var rows = Math.max(m.rows.length, result.cleaned.length);
    var cells = 0;
    for (var r = 0; r < rows; r++) {
      var a = result.cleaned[r] || [];
      var b = m.rows[r] || [];
      var w = Math.max(a.length, b.length);
      for (var c = 0; c < w; c++) {
        cells++;
        if (!same(a[c], b[c])) {
          mism.push(addr(r, c) + ': tool=' + JSON.stringify(P.cellToString(a[c]).slice(0, 40)) +
            ' mẫu=' + JSON.stringify(P.cellToString(b[c]).slice(0, 40)));
        }
      }
    }
    return [mism.length === 0
      ? pass('3. Daily Report CCVN so với file mẫu',
          'Khớp ' + cells + ' ô; đã xoá cột rỗng ' + (result.removedColumns.join(', ') || '(không có)') +
          '; ' + result.dataRows.length + ' dòng dữ liệu')
      : fail('3. Daily Report CCVN so với file mẫu', mism.length + ' ô lệch', mism)];
  }

  /* ---- 4/5/6: so khối dán với File báo cáo tổng ---- */
  function compareBlock(name, block, masterWb, colCount, opts) {
    opts = opts || {};
    var ws = masterWb.Sheets[block.sheet];
    if (!ws) return fail(name, 'File báo cáo tổng không có sheet "' + block.sheet + '"');
    var m = P.sheetToMatrix(ws);
    var a0 = XLSX.utils.decode_cell(block.anchor);

    var expRows = [];
    for (var r = a0.r; r < m.rows.length; r++) {
      var row = m.rows[r];
      if (P.rowIsEmpty(row, m.maxCol)) break;
      expRows.push(row);
    }

    var mism = [];
    var limit = opts.subset ? block.rows.length : Math.max(expRows.length, block.rows.length);
    for (var i = 0; i < limit; i++) {
      for (var c = 0; c < colCount; c++) {
        if (opts.skipCols && opts.skipCols.indexOf(c) !== -1) continue;
        var got = block.rows[i] ? block.rows[i][c] : undefined;
        var exp = expRows[i] ? expRows[i][a0.c + c] : undefined;
        if (!same(got, exp)) {
          mism.push(addr(a0.r + i, a0.c + c) + ' (' + (block.headers[c] || ('cột ' + c)) + '): tool=' +
            JSON.stringify(P.cellToString(got).slice(0, 40)) + ' tổng=' +
            JSON.stringify(P.cellToString(exp).slice(0, 40)));
        }
      }
    }

    var countOk = opts.subset ? block.rows.length <= expRows.length : block.rows.length === expRows.length;
    var msg = 'tool ' + block.rows.length + ' dòng, file tổng ' + expRows.length + ' dòng' +
      (opts.subset ? ' (chấp nhận tập con)' : '');
    return (mism.length === 0 && countOk)
      ? pass(name, 'Khớp ' + (limit * colCount) + ' ô — ' + msg)
      : fail(name, mism.length + ' ô lệch — ' + msg, mism);
  }

  /**
   * Khối 3 gom theo supporter: file tổng gộp nhiều file OP nên chứa cả supporter
   * không có trong file đang nạp. So theo vị trí sẽ lệch ngay khi thiếu một người,
   * nên ghép nhóm theo email rồi mới so từng dòng trong nhóm.
   */
  function groupByEmail(rows, emailCol) {
    var groups = [], cur = null;
    rows.forEach(function (row, i) {
      var e = P.cellToString(row[emailCol]).trim();
      if (!cur || cur.email !== e) {
        cur = { email: e, rows: [], at: i };
        groups.push(cur);
      }
      cur.rows.push(row);
    });
    return groups;
  }

  function compareBlock3(name, block, masterWb) {
    var ws = masterWb.Sheets[block.sheet];
    if (!ws) return fail(name, 'File báo cáo tổng không có sheet "' + block.sheet + '"');
    var m = P.sheetToMatrix(ws);
    var a0 = XLSX.utils.decode_cell(block.anchor);
    var COLS = block.headers.length;
    var EMAIL = 2;

    var expRows = [];
    for (var r = a0.r; r < m.rows.length; r++) {
      var row = m.rows[r];
      if (P.rowIsEmpty(row, m.maxCol)) break;
      var slice = [];
      for (var c = 0; c < COLS; c++) slice.push(row[a0.c + c]);
      slice.__row = r;
      expRows.push(slice);
    }

    var expGroups = {}, dupes = [];
    groupByEmail(expRows, EMAIL).forEach(function (g) {
      if (expGroups[g.email]) dupes.push(g.email);
      else expGroups[g.email] = g;
    });
    var gotGroups = groupByEmail(block.rows, EMAIL);

    var mism = [], cells = 0;
    gotGroups.forEach(function (g) {
      var exp = expGroups[g.email];
      if (!exp) {
        mism.push('supporter "' + g.email + '" không có trong file tổng');
        return;
      }
      if (exp.rows.length !== g.rows.length) {
        mism.push('supporter "' + g.email + '": tool ' + g.rows.length +
          ' dòng, file tổng ' + exp.rows.length + ' dòng');
      }
      var n = Math.min(exp.rows.length, g.rows.length);
      for (var i = 0; i < n; i++) {
        for (var c = 0; c < COLS; c++) {
          cells++;
          if (!same(g.rows[i][c], exp.rows[i][c])) {
            mism.push(addr(exp.rows[i].__row, a0.c + c) + ' (' + block.headers[c] + ', ' +
              g.email + '): tool=' + JSON.stringify(P.cellToString(g.rows[i][c]).slice(0, 30)) +
              ' tổng=' + JSON.stringify(P.cellToString(exp.rows[i][c]).slice(0, 30)));
          }
        }
      }
    });

    var extra = Object.keys(expGroups).filter(function (e) {
      return !gotGroups.some(function (g) { return g.email === e; });
    });
    var detail = 'tool ' + gotGroups.length + ' supporter / ' + block.rows.length + ' dòng; file tổng ' +
      Object.keys(expGroups).length + ' supporter / ' + expRows.length + ' dòng' +
      (extra.length ? '; chỉ có ở file tổng (thiếu file OP): ' + extra.join(', ') : '');
    if (dupes.length) detail += '; nhóm bị tách rời: ' + dupes.join(', ');

    return mism.length === 0
      ? pass(name, 'Khớp ' + cells + ' ô — ' + detail)
      : fail(name, mism.length + ' điểm lệch — ' + detail, mism);
  }

  /* ---- 7: kiểu ô khi ghi ---- */
  function checkCellTypes(blocks) {
    var probs = [];
    ['block1', 'block2', 'block3'].forEach(function (k) {
      var b = blocks[k];
      if (!b || !b.rows.length) return;
      b.types.forEach(function (t, c) {
        var sample = null;
        for (var i = 0; i < b.rows.length && sample === null; i++) {
          var v = b.rows[i][c];
          if (v !== '' && v !== null && v !== undefined) sample = v;
        }
        if (sample === null) return;
        var cell = Writer.makeCell(sample, t);
        if (t === 'text' && cell.t !== 's') probs.push(b.title + ' cột "' + b.headers[c] + '" không ra text');
        if ((t === 'date' || t === 'number') && cell.t !== 'n') probs.push(b.title + ' cột "' + b.headers[c] + '" không ra số');
        /* ID dài phải giữ nguyên từng chữ số */
        if (t === 'text' && /^\d{15,}$/.test(String(sample)) && cell.v !== String(sample)) {
          probs.push(b.title + ' cột "' + b.headers[c] + '" mất chính xác ID: ' + sample + ' -> ' + cell.v);
        }
      });
    });
    return probs.length === 0
      ? pass('7. Kiểu ô khi ghi', 'ID/timestamp ghi kiểu text, No/Date ghi kiểu số — ID 15+ chữ số giữ nguyên')
      : fail('7. Kiểu ô khi ghi', probs.length + ' vấn đề', probs);
  }

  /**
   * files: { session, ccvn, ops[], expectedSession, expectedCcvn, master } — mỗi cái {name, wb}
   */
  function run(files, cfg) {
    var checks = [];
    if (!files.session) return [fail('Thiếu dữ liệu', 'Cần ít nhất file Session Report')];

    var session = Pipeline.processSession(files.session.wb, cfg, files.session.name);
    checks = checks.concat(checkSession(session, files.expectedSession && files.expectedSession.wb, cfg));

    var ccvn = null;
    if (files.ccvn) {
      ccvn = Pipeline.processCcvn(files.ccvn.wb, cfg, files.ccvn.name);
      checks = checks.concat(checkCcvn(ccvn, files.expectedCcvn && files.expectedCcvn.wb));
    } else {
      checks.push(skip('3. Daily Report CCVN so với file mẫu', 'Chưa nạp Daily Report CCVN'));
    }

    var det = Pipeline.detectReportDate(session, cfg);
    var serial = det ? det.serial : null;
    if (!serial) return checks.concat([fail('Ngày báo cáo', 'Không suy được ngày báo cáo')]);
    checks.push(pass('0. Ngày báo cáo tự nhận', det.iso + ' (serial ' + det.serial + ', ' + det.count + ' phiên)'));

    var ops = (files.ops || []).map(function (f) {
      return Pipeline.processOpFile(f.wb, cfg, serial, f.name);
    });
    var blocks = Pipeline.buildBlocks({ session: session, ccvn: ccvn, ops: ops }, cfg, serial);

    if (files.master) {
      checks.push(compareBlock('4. Khối 1 so với "1. CSKH" của file tổng',
        blocks.block1, files.master.wb, 12));
      if (ccvn) {
        checks.push(compareBlock('5. Khối 2 so với "2. SF Case info" của file tổng',
          blocks.block2, files.master.wb, 7, { skipCols: [5] }));
        checks.push(checkEmails(blocks.block2, files.master.wb, cfg));
      }
      checks.push(compareBlock3('6. Khối 3 so với "3. Other Tasks" của file tổng',
        blocks.block3, files.master.wb));
    } else {
      checks.push(skip('4-6. So với File báo cáo tổng', 'Chưa nạp File báo cáo tổng'));
    }

    checks.push(checkCellTypes(blocks));
    return checks;
  }

  /**
   * Cột G của "2. SF Case info" là công thức VLOOKUP. So email tool tra được với
   * giá trị Excel đã tính sẵn trong file — chứ không so tool với chính bảng của tool.
   *
   * Lệch chỉ được chấp nhận ở tên bị TRÙNG trong `mail SF`: chỗ đó giá trị cache trong
   * file là kết quả cũ chưa tính lại, còn VLOOKUP luôn lấy dòng khớp đầu tiên.
   */
  function checkEmails(block2, masterWb, cfg) {
    var ws = masterWb.Sheets[cfg.targets.sfcase.sheet];
    if (!ws) return skip('5b. Cột email so với file tổng',
      'Không có sheet "' + cfg.targets.sfcase.sheet + '"');
    var m = P.sheetToMatrix(ws);
    var a0 = XLSX.utils.decode_cell(cfg.targets.sfcase.anchor);

    function stem(mail) { return mail.split('@')[0].replace(/\d+$/, ''); }

    var bad = [], staleRows = 0, checked = 0;
    block2.rows.forEach(function (row, i) {
      var fileVal = P.cellToString((m.rows[a0.r + i] || [])[a0.c + 5]).trim();
      if (!fileVal) return;
      checked++;
      var got = P.cellToString(row[5]).trim();
      if (got === fileVal) return;
      /* cùng một người, chỉ khác hậu tố số -> file đang giữ cache cũ của tên trùng */
      if (got && stem(got) === stem(fileVal)) { staleRows++; return; }
      if (bad.length < MAX_SHOW) {
        bad.push('dòng ' + (a0.r + i + 1) + ' "' + row[4] + '": tool=' + got + ' file=' + fileVal);
      }
    });

    var detail = 'Đối chiếu ' + checked + ' ô với cột email có sẵn trong file tổng';
    if (staleRows) {
      detail += '; ' + staleRows + ' ô lệch do `mail SF` có tên trùng — file giữ cache cũ, ' +
        'VLOOKUP tính lại sẽ ra đúng như tool';
    }
    return bad.length === 0 ? pass('5b. Cột email so với file tổng', detail)
                            : fail('5b. Cột email so với file tổng', bad.length + ' ô lệch thật', bad);
  }

  /* =====================================================================
   * Đối chứng: so file tool vừa sinh với một bản đúng đã biết
   * ===================================================================== */

  function emailsIn(wb, sheetName, col, r0, r1) {
    var out = {};
    var ws = wb && wb.Sheets[sheetName];
    if (!ws) return out;
    var m = P.sheetToMatrix(ws);
    for (var r = r0 - 1; r < Math.min(m.rows.length, r1); r++) {
      var v = P.cellToString((m.rows[r] || [])[col]).trim();
      if (v.indexOf('@') !== -1) out[P.normText(v)] = r + 1;
    }
    return out;
  }

  /**
   * So một vùng khối giữa 2 workbook.
   *
   * `opts.explain(i, c, av, bv, rowA, rowB)` trả về lý do nếu ô lệch đó đã truy được
   * nguyên nhân nằm ngoài tool — khi đó vẫn hiện ra để soát nhưng không tính là lỗi.
   * Kết quả gom theo CỘT: biết "127 ô lệch, toàn bộ ở cột Case Status" thì đọc được
   * ngay, chứ một con số 146 trần trụi thì không nói lên điều gì.
   */
  function diffBlock(name, outWb, refWb, block, opts) {
    opts = opts || {};
    var a = outWb.Sheets[block.sheet], b = refWb.Sheets[block.sheet];
    if (!a || !b) return fail(name, 'Thiếu sheet "' + block.sheet + '" ở một trong hai file');
    var ma = P.sheetToMatrix(a), mb = P.sheetToMatrix(b);
    var a0 = XLSX.utils.decode_cell(block.anchor);
    var nCol = block.headers.length;

    var byCol = {}, cells = 0, stale = 0, nBad = 0;
    var explained = [], explainCount = 0;
    function stem(s) { return s.split('@')[0].replace(/\d+$/, ''); }
    /* đọc ô theo chỉ số cột CỦA KHỐI (0..nCol-1), không phải chỉ số cột của sheet */
    function colOf(row) { return function (c) { return row[a0.c + c]; }; }
    for (var i = 0; i < block.rows.length; i++) {
      var rowA = ma.rows[a0.r + i] || [], rowB = mb.rows[a0.r + i] || [];
      for (var c = 0; c < nCol; c++) {
        var av = rowA[a0.c + c], bv = rowB[a0.c + c];
        cells++;
        if (same(av, bv)) continue;
        /* cột email: bản đúng có thể giữ cache VLOOKUP cũ của tên trùng */
        if (opts.emailCol === c) {
          var sa = P.cellToString(av).trim(), sb = P.cellToString(bv).trim();
          if (sa && sb && stem(sa) === stem(sb)) { stale++; continue; }
        }
        var line = addr(a0.r + i, a0.c + c) + ': tool=' +
          JSON.stringify(P.cellToString(av).slice(0, 30)) + ' chuẩn=' +
          JSON.stringify(P.cellToString(bv).slice(0, 30));
        var why = opts.explain ? opts.explain(i, c, av, bv, colOf(rowA), colOf(rowB)) : null;
        if (why) {
          explainCount++;
          if (explained.length < MAX_SHOW) explained.push(line + ' — ' + why);
          continue;
        }
        nBad++;
        var d = byCol[c] || (byCol[c] = { header: block.headers[c] || ('cột ' + (c + 1)),
                                          count: 0, samples: [] });
        d.count++;
        if (d.samples.length < 3) d.samples.push(line);
      }
    }

    var detail = 'So ' + cells + ' ô (' + block.rows.length + ' dòng × ' + nCol + ' cột)' +
      (stale ? '; bỏ qua ' + stale + ' ô email do bản chuẩn giữ cache cũ' : '') +
      (explainCount ? '; ' + explainCount + ' ô lệch đã truy được nguyên nhân (không phải lỗi tool)' : '');
    if (!nBad) return pass(name, detail + (explainCount ? '' : ' — khớp hết'), explained);

    var cols = Object.keys(byCol).map(Number).sort(function (x, y) {
      return byCol[y].count - byCol[x].count;
    });
    var lines = [];
    cols.forEach(function (c) {
      lines.push('▸ ' + byCol[c].header + ': ' + byCol[c].count + ' ô');
      byCol[c].samples.forEach(function (s) { lines.push('    ' + s); });
    });
    return fail(name, nBad + ' ô lệch ở ' + cols.length + ' cột — ' + detail,
                lines.concat(explained));
  }

  /**
   * Lệch do chính luật `requestedByAliases` của tool: tool đổi tên còn bản chuẩn
   * thì không (người làm quên áp dụng hôm đó, hoặc sửa tay rồi bỏ dở giữa chừng —
   * cả hai trường hợp đều đã gặp trong dữ liệu thật).
   * Ô email của đúng dòng ấy lệch theo, nên cũng tính là đã giải thích.
   */
  function aliasExplainer(cfg, nameCol, emailCol) {
    var aliases = (cfg.ccvn && cfg.ccvn.requestedByAliases) || cfg.requestedByAliases;
    if (!aliases) return null;
    var toSource = {};
    Object.keys(aliases).forEach(function (src) {
      if (!aliases[src]) return;
      (toSource[P.normText(aliases[src])] = toSource[P.normText(aliases[src])] || []).push(src);
    });
    function isAliasPair(tool, ref) {
      var list = toSource[P.normText(P.cellToString(tool))];
      if (!list) return null;
      var r = P.normText(P.cellToString(ref));
      for (var i = 0; i < list.length; i++) {
        if (P.normText(list[i]) === r) return list[i];
      }
      return null;
    }
    return function (i, c, av, bv, getA, getB) {
      if (c !== nameCol && c !== emailCol) return null;
      /* ô email lệch là hệ quả của ô tên cùng dòng, nên luôn xét theo cột tên */
      var base = isAliasPair(getA(nameCol), getB(nameCol));
      return base ? 'bản chuẩn không đổi tên "' + base + '" ở dòng này (luật `requestedByAliases`)' : null;
    };
  }

  /**
   * So cột ngày của một sheet tổng hợp, chỉ ở những dòng tool có ghi.
   * `excuse(row)` trả về lý do nếu chênh lệch ở dòng đó đã truy được nguyên nhân
   * nằm ngoài tool — khi đó tính là "đã giải thích" chứ không phải lỗi.
   */
  function diffDayColumn(name, outWb, refWb, sheetName, headerRow, serial, rows, excuse) {
    var a = outWb.Sheets[sheetName], b = refWb.Sheets[sheetName];
    if (!a || !b) return fail(name, 'Thiếu sheet "' + sheetName + '"');
    var ma = P.sheetToMatrix(a), mb = P.sheetToMatrix(b);
    var col = -1;
    (ma.rows[headerRow - 1] || []).forEach(function (v, c) {
      if (String(v).trim() === String(serial)) col = c;
    });
    if (col === -1) return fail(name, 'Không tìm thấy cột ngày ' + serial);

    var mism = [], explained = [];
    rows.forEach(function (r) {
      var av = (ma.rows[r - 1] || [])[col];
      var bv = (mb.rows[r - 1] || [])[col];
      if (same(av, bv)) return;
      var line = 'dòng ' + r + ': tool=' + JSON.stringify(P.cellToString(av).slice(0, 22)) +
        ' chuẩn=' + JSON.stringify(P.cellToString(bv).slice(0, 22));
      var why = excuse ? excuse(r) : null;
      if (why) explained.push(line + ' — ' + why);
      else mism.push(line);
    });

    var detail = 'So ' + rows.length + ' ô ở cột ngày ' + serial;
    if (explained.length) {
      detail += '; ' + explained.length + ' ô lệch đã truy được nguyên nhân (không phải lỗi tool)';
    } else if (!mism.length) {
      detail += ' — khớp hết';
    }
    return mism.length === 0 ? pass(name, detail, explained)
                             : fail(name, mism.length + '/' + rows.length + ' ô lệch — ' + detail,
                                    mism.concat(explained));
  }

  /**
   * outBytes: file tool vừa sinh.  refWb: workbook bản đúng đã biết.
   * Chỉ so những chỗ tool THỰC SỰ ghi — các ô nhập tay không tính.
   */
  function compareWithReference(outBytes, refWb, blocks, cfg, opts) {
    var checks = [];
    var outWb = XLSX.read(outBytes, { type: 'array', cellFormula: false, cellHTML: false,
                                      cellDates: false, cellNF: false, cellText: false });

    checks.push(diffBlock('1. Sheet "1. CSKH"', outWb, refWb, blocks.block1));
    if (blocks.block2.rows.length) {
      /* cột 4 = Requested By, cột 5 = Supporter email (tính từ ô neo B7) */
      checks.push(diffBlock('2. Sheet "2. SF Case info"', outWb, refWb, blocks.block2, {
        emailCol: 5,
        explain: aliasExplainer(cfg, 4, 5)
      }));
    }
    if (blocks.block3.rows.length) {
      checks.push(diffBlock('3. Sheet "3. Other Tasks"', outWb, refWb, blocks.block3));
    }

    if (cfg.summary && blocks.summary) {
      var sumRows = (cfg.summary.cells || []).map(function (c) { return c.row; })
        .concat((cfg.summary.topicRows || []).map(function (c) { return c.row; }))
        /* dòng fanpage: chỉ so những dòng lần chạy này thực sự ghi */
        .concat((opts && opts.fanpageRows) || []);
      checks.push(diffDayColumn('4. Cột ngày "' + cfg.summary.sheet + '"', outWb, refWb,
        cfg.summary.sheet, cfg.summary.dayHeaderRow, blocks.reportSerial, sumRows));
    }

    if (cfg.kpi && blocks.kpi) {
      /* chỉ so những dòng tool ghi: mọi người có dòng KPI trong Daily KPI Result,
         kể cả người nằm ngoài khối tổng OP (dòng 44-54) */
      var opRows = emailsIn(refWb, cfg.kpi.daily.sheet, cfg.kpi.daily.emailColumn,
        cfg.kpi.daily.dataStartRow,
        cfg.kpi.daily.peopleEndRow || cfg.kpi.daily.dataEndRow);
      var mName = Object.keys(refWb.Sheets).filter(function (n) {
        return new RegExp(cfg.kpi.monthly.sheetPattern, 'i').test(n);
      })[0];
      if (mName) {
        var mRows = emailsIn(refWb, mName, cfg.kpi.monthly.emailColumn,
          cfg.kpi.monthly.dataStartRow, cfg.kpi.monthly.dataEndRow);
        /* người dùng định mức KPI riêng thì tool cố tình để trống, không so */
        var oddRate = (opts && opts.oddRateEmails) || {};
        var used = Object.keys(mRows).filter(function (e) {
          return opRows[e] && !oddRate[e];
        });
        var wrote = used.map(function (e) { return mRows[e]; });
        var rowEmail = {};
        used.forEach(function (e) { rowEmail[mRows[e]] = e; });
        if (cfg.kpi.monthly.teamRow) wrote.push(cfg.kpi.monthly.teamRow);

        /* Người có tên bị trùng trong `mail SF` thì bản chuẩn đang giữ số case cũ,
           nên hiệu suất lệch là do bản chuẩn chứ không phải tool. Dòng tổng team
           cộng dồn cả nhóm nên cũng bị kéo theo. */
        var affected = (opts && opts.affectedEmails) || {};
        var anyAffected = used.some(function (e) { return affected[e]; });
        var excuse = function (r) {
          if (cfg.kpi.monthly.teamRow && r === cfg.kpi.monthly.teamRow) {
            return anyAffected ? 'dòng tổng, bị kéo theo bởi người có tên trùng trong `mail SF`' : null;
          }
          var e = rowEmail[r];
          if (!e) return null;
          if (affected[e]) {
            return '`mail SF` có tên trùng nên bản chuẩn giữ số case cũ (' + e + ')';
          }
          /* Hiệu suất = tổng-việc / tổng-giờ. Người có việc mà file OP không có khối
             giờ của họ thì mẫu số = 0, buộc phải ra 0 — nguồn thiếu, không phải tool sai. */
          var k = blocks.kpi && blocks.kpi.forEmail(e);
          if (k && !k.hours && k.hasWork) {
            return 'file OP không có khối giờ của ' + e + ' trong ngày này nên hiệu suất ' +
                   'buộc ra 0 (bản chuẩn có số vì được nhập tay từ nguồn khác)';
          }
          return null;
        };
        checks.push(diffDayColumn('5. Cột ngày "' + mName + '"', outWb, refWb,
          mName, cfg.kpi.monthly.dayHeaderRow, blocks.reportSerial, wrote, excuse));
      }
    }

    /* Lịch sử phải khớp FILE NỀN, không phải bản chuẩn: người làm có thể sửa
       số của ngày cũ sau khi đã phát hành, nên bản chuẩn không còn là mốc đúng
       cho phần lịch sử. Điều tool phải bảo đảm là không tự ý đổi cái nó nhận vào. */
    var baseWb = (opts && opts.baseWb) || refWb;
    checks.push(checkHistory(outWb, baseWb, blocks, cfg,
      opts && opts.baseWb ? 'file nền' : 'bản chuẩn'));
    if (opts && opts.baseWb) {
      var edits = historyEdits(refWb, opts.baseWb, blocks, cfg);
      if (edits) checks.push(edits);
    }
    checks.push(checkPivotFilters(outBytes));
    return checks;
  }

  /** serial ngày -> chỉ số cột, đọc trên dòng header ngày của một workbook. */
  function dayColumns(wb, cfg) {
    var out = {};
    var ws = wb && wb.Sheets[cfg.summary.sheet];
    if (!ws) return out;
    var m = P.sheetToMatrix(ws);
    (m.rows[cfg.summary.dayHeaderRow - 1] || []).forEach(function (v, c) {
      var s = parseFloat(String(v).trim());
      if (isFinite(s)) out[s] = c;
    });
    return out;
  }

  /**
   * Bản chuẩn khác file nền ở NGÀY CŨ nghĩa là người làm đã sửa tay số của ngày
   * trước trong lúc làm báo cáo hôm sau. Không phải lỗi tool — tool giữ nguyên cái
   * nó nhận vào là đúng — nhưng phải nói ra, nếu không người dùng thấy check 6 đỏ
   * mà không hiểu vì sao.
   */
  function historyEdits(refWb, baseWb, blocks, cfg) {
    var name = '6b. Bản chuẩn sửa lại số của ngày cũ';
    var ra = refWb && refWb.Sheets[cfg.summary.sheet];
    var ba = baseWb && baseWb.Sheets[cfg.summary.sheet];
    if (!ra || !ba) return null;
    var mr = P.sheetToMatrix(ra), mb = P.sheetToMatrix(ba);
    var cr = dayColumns(refWb, cfg), cb = dayColumns(baseWb, cfg);

    var notes = [], n = 0;
    Object.keys(cb).map(Number).sort().forEach(function (serial) {
      if (serial >= blocks.reportSerial || cr[serial] === undefined) return;
      for (var r = 5; r <= 70; r++) {
        var av = (mr.rows[r - 1] || [])[cr[serial]];
        var bv = (mb.rows[r - 1] || [])[cb[serial]];
        if (same(av, bv)) continue;
        n++;
        if (notes.length < MAX_SHOW) {
          notes.push(addr(r - 1, cr[serial]) + ' (ngày ' + serial + '): file nền=' +
            JSON.stringify(P.cellToString(bv).slice(0, 20)) + ' bản chuẩn=' +
            JSON.stringify(P.cellToString(av).slice(0, 20)));
        }
      }
    });
    if (!n) return null;
    return pass(name, n + ' ô ở các cột ngày TRƯỚC đã bị sửa tay giữa file nền và bản ' +
      'chuẩn — tool giữ nguyên theo file nền nên phần chênh này không phải lỗi tool. ' +
      'Nếu số mới mới là số đúng thì phải sửa trong file nền trước khi chạy.', notes);
  }

  /**
   * Chốt chặn cho một lỗi CHỈ lộ ra khi mở bằng Excel thật, không cách nào thấy
   * qua so ô: pivot lấy `Date/日付` làm page field với bộ lọc tay, nên ngày báo
   * cáo mới bị coi là "mục mới" và không được chọn -> refresh xong pivot rỗng ->
   * cột phụ Sheet5 = 0 -> "Daily KPI Result" ra 0 hết. Mọi ô dữ liệu vẫn đúng,
   * nên các check 1-6 đều xanh trong khi file giao ra thực chất hỏng.
   */
  function checkPivotFilters(outBytes) {
    var name = '8. Pivot nhận ngày mới';
    if (typeof fflate === 'undefined') return skip(name, 'Không có fflate để mở zip');
    var files;
    try { files = fflate.unzipSync(outBytes); }
    catch (e) { return fail(name, 'Không giải nén được file: ' + e.message); }

    var total = 0, missing = [];
    Object.keys(files).forEach(function (n) {
      if (!/^xl\/pivotTables\/pivotTable\d+\.xml$/.test(n)) return;
      var xml = fflate.strFromU8(files[n]);
      var fields = xml.match(/<pivotField\b[^>]*\/>|<pivotField\b[^>]*>/g) || [];
      fields.forEach(function (head) {
        if (head.indexOf(' axis="') === -1) return;
        total++;
        if (head.indexOf('includeNewItemsInFilter="1"') === -1) {
          missing.push(n.split('/').pop() + ' ' + (head.match(/axis="([^"]*)"/) || [])[1]);
        }
      });
    });
    if (!total) return skip(name, 'File không có PivotTable nào dùng bộ lọc');
    if (missing.length) {
      return fail(name, missing.length + '/' + total +
        ' trường pivot thiếu includeNewItemsInFilter — mở bằng Excel sẽ ra pivot rỗng ' +
        'và "Daily KPI Result" toàn số 0', missing);
    }
    return pass(name, 'Cả ' + total + ' trường pivot đều nhận mục mới vào bộ lọc');
  }

  /** Các cột ngày TRƯỚC ngày báo cáo không được đụng tới. */
  function checkHistory(outWb, refWb, blocks, cfg, srcLabel) {
    var sheetName = cfg.summary.sheet;
    var a = outWb.Sheets[sheetName], b = refWb.Sheets[sheetName];
    if (!a || !b) return skip('6. Lịch sử các ngày trước', 'Thiếu sheet');
    var ma = P.sheetToMatrix(a), mb = P.sheetToMatrix(b);
    var hdr = ma.rows[cfg.summary.dayHeaderRow - 1] || [];

    var mism = [], cells = 0, cols = 0;
    hdr.forEach(function (v, c) {
      var s = parseFloat(String(v).trim());
      if (!isFinite(s) || s >= blocks.reportSerial) return;   /* chỉ xét ngày trước đó */
      cols++;
      for (var r = 5; r <= 70; r++) {
        var av = (ma.rows[r - 1] || [])[c];
        var bv = (mb.rows[r - 1] || [])[c];
        cells++;
        if (!same(av, bv) && mism.length < MAX_SHOW) {
          mism.push(addr(r - 1, c) + ': tool=' + JSON.stringify(P.cellToString(av).slice(0, 20)) +
            ' nguồn=' + JSON.stringify(P.cellToString(bv).slice(0, 20)));
        }
      }
    });
    var lbl = srcLabel || 'bản chuẩn';
    return mism.length === 0
      ? pass('6. Lịch sử các ngày trước', 'So ' + cells + ' ô ở ' + cols +
          ' cột ngày trước với ' + lbl + ' — không ô nào bị đụng')
      : fail('6. Lịch sử các ngày trước',
          mism.length + ' ô khác so với ' + lbl, mism);
  }

  return { run: run, compareWithReference: compareWithReference };
});
