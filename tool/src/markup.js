/* Dựng bản Excel ĐÃ ĐÁNH DẤU từ chính file người dùng nạp vào.
 *
 * Mục đích: người làm thấy ngay lỗi nằm ở ô nào, sửa thẳng trong file đó rồi nạp
 * lại — không phải tự dò theo số dòng trong bảng cảnh báo.
 *
 * Cách làm giống `xlsxpatch.js`: giải nén zip, chỉ sửa đúng phần cần sửa, giữ
 * nguyên byte của mọi part còn lại. Cụ thể:
 *   1. Thêm một fill vàng + một cellXf vào `styles.xml`, lấy chỉ số xf mới.
 *   2. Đổi thuộc tính `s` của các ô lỗi sang xf đó (giữ giá trị, chỉ đổi màu).
 *   3. Thêm sheet liệt kê lỗi ở ĐẦU workbook, mỗi dòng một ô kèm lý do.
 *
 * Không đổi giá trị ô nào: file vẫn nạp lại vào tool được như bình thường, và
 * màu đánh dấu không ảnh hưởng kết quả xử lý.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(root.Parse);
  else root.Markup = factory(root.Parse);
})(typeof self !== 'undefined' ? self : this, function (P) {
  'use strict';

  var SHEET_TITLE = 'Loi can sua';
  var BAD_CHARS = /[\x00-\x08\x0B\x0C\x0E-\x1F]/g;

  function esc(s) {
    return String(s).replace(BAD_CHARS, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }
  function getAttr(tag, name) {
    var m = tag.match(new RegExp('\\s' + name + '="([^"]*)"'));
    return m ? m[1] : null;
  }
  function numToCol(n) {
    var s = '';
    n += 1;
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
    return s;
  }

  /* ---------- styles.xml: thêm một fill màu và một cellXf dùng fill đó ---------- */
  function addHighlightStyle(xml, argb) {
    var fillId = null;
    xml = xml.replace(/<fills\b([^>]*)>([\s\S]*?)<\/fills>/, function (all, at, body) {
      var n = (body.match(/<fill[ >\/]/g) || []).length;
      fillId = n;
      var add = '<fill><patternFill patternType="solid"><fgColor rgb="' + argb +
                '"/><bgColor indexed="64"/></patternFill></fill>';
      return '<fills' + at.replace(/\s*count="\d+"/, '') + ' count="' + (n + 1) + '">' +
             body + add + '</fills>';
    });
    if (fillId === null) return null;

    /* cellXf mới: dùng font/border mặc định, chỉ gắn fill vừa thêm */
    var xfId = null;
    xml = xml.replace(/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/, function (all, at, body) {
      var n = (body.match(/<xf[ >\/]/g) || []).length;
      xfId = n;
      var add = '<xf numFmtId="0" fontId="0" fillId="' + fillId +
                '" borderId="0" xfId="0" applyFill="1"/>';
      return '<cellXfs' + at.replace(/\s*count="\d+"/, '') + ' count="' + (n + 1) + '">' +
             body + add + '</cellXfs>';
    });
    if (xfId === null) return null;
    return { xml: xml, xfId: xfId, fillId: fillId };
  }

  /* Mỗi ô tô vàng giữ định dạng ngày/số, font, viền, căn lề và bảo vệ cũ.
     Dùng một style General cho tất cả sẽ biến ngày hiển thị thành số 46xxx. */
  function highlightVariants(xml, fillId) {
    var group = xml.match(/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/);
    var originals = group[2].match(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g) || [];
    var added = [], mapped = {};
    return {
      forStyle: function (oldId) {
        if (mapped[oldId] !== undefined) return mapped[oldId];
        var base = originals[oldId] || originals[0];
        var clone = base.replace(/<xf\b[^>]*>/, function (head) {
          function set(name, value) {
            var re = new RegExp('\\s' + name + '="[^"]*"');
            head = re.test(head) ? head.replace(re, ' ' + name + '="' + value + '"')
              : head.replace(/\/?>(?=$)/, function (end) { return ' ' + name + '="' + value + '"' + end; });
          }
          set('fillId', fillId); set('applyFill', 1); return head;
        });
        mapped[oldId] = originals.length + added.length;
        added.push(clone);
        return mapped[oldId];
      },
      xml: function () {
        return xml.replace(group[0], '<cellXfs' + group[1].replace(/\s*count="\d+"/, '') +
          ' count="' + (originals.length + added.length) + '">' + group[2] + added.join('') + '</cellXfs>');
      }
    };
  }

  /* ---------- đổi thuộc tính `s` của các ô cần đánh dấu ---------- */
  function paintCells(xml, byRow, styleFor) {
    var ROW_RE = /<row\b[^>]*\/>|<row\b[^>]*>[\s\S]*?<\/row>/g;
    var CELL_RE = /<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g;
    var painted = 0;

    var out = xml.replace(ROW_RE, function (rowXml) {
      var head = rowXml.slice(0, rowXml.indexOf('>') + 1);
      var rn = parseInt(getAttr(head, 'r'), 10);
      var want = byRow[rn];
      if (!want) return rowXml;

      return rowXml.replace(CELL_RE, function (cellXml) {
        var ch = cellXml.slice(0, cellXml.indexOf('>') + 1);
        var ref = getAttr(ch, 'r');
        if (!ref) return cellXml;
        var m = ref.match(/^([A-Z]+)(\d+)$/);
        if (!m) return cellXml;
        var ci = 0;
        for (var i = 0; i < m[1].length; i++) ci = ci * 26 + (m[1].charCodeAt(i) - 64);
        ci -= 1;
        if (want.indexOf(ci) === -1) return cellXml;
        painted++;
        var xfId = styleFor(parseInt(getAttr(ch, 's') || '0', 10));
        var nh = /\ss="\d+"/.test(ch)
          ? ch.replace(/\ss="\d+"/, ' s="' + xfId + '"')
          : (ch.slice(-2) === '/>' ? ch.slice(0, -2) + ' s="' + xfId + '"/>'
                                   : ch.slice(0, -1) + ' s="' + xfId + '">');
        return nh + cellXml.slice(ch.length);
      });
    });
    return { xml: out, painted: painted };
  }

  /* ---------- sheet liệt kê lỗi ---------- */
  function issueSheetXml(fileEntry, headerXfId, sourceWb) {
    var rows = [], r = 1;
    function cell(col, row, text, xf) {
      return '<c r="' + numToCol(col) + row + '"' + (xf !== null && xf !== undefined ? ' s="' + xf + '"' : '') +
             ' t="inlineStr"><is><t xml:space="preserve">' + esc(text) + '</t></is></c>';
    }
    function line(vals, xf) {
      var cells = vals.map(function (v, i) { return v === '' ? '' : cell(i, r, v, xf); }).join('');
      rows.push('<row r="' + r + '">' + cells + '</row>');
      r++;
    }

    line(['Các ô được đánh dấu trong file: ' + fileEntry.file], headerXfId);
    line(['Nội dung ô được hiển thị theo định dạng trong file gốc. Màu vàng không có nghĩa là ô nào cũng sai; xem cột Trạng thái trước khi sửa.'], null);
    line([''], null);
    line(['Sheet trong file gốc', 'Ô', 'Dòng', 'Nội dung ô trong file gốc', 'Trạng thái', 'Vì sao được đánh dấu', 'Tool xử lý / Bạn cần làm'], headerXfId);

    Object.keys(fileEntry.sheets).forEach(function (sheetName) {
      fileEntry.sheets[sheetName].forEach(function (sp) {
        var addr = (sp.col === null || sp.col === undefined || !sp.row) ? '' : numToCol(sp.col) + sp.row;
        var shown = P.sourceCellText(sourceWb, sheetName, sp.row, sp.col, sp.value);
        line([sheetName, addr, sp.row ? String(sp.row) : '', shown === '' ? '(trống)' : shown,
          sp.level === 'info' ? 'Đã xử lý — xem ghi chú' : 'Cần kiểm tra', sp.message, sp.why], null);
      });
    });

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetPr><tabColor rgb="FFFFC000"/></sheetPr>' +
      '<dimension ref="A1:G' + Math.max(1, r - 1) + '"/>' +
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<cols>' +
      '<col min="1" max="1" width="30" customWidth="1"/>' +
      '<col min="2" max="2" width="8" customWidth="1"/>' +
      '<col min="3" max="3" width="8" customWidth="1"/>' +
      '<col min="4" max="4" width="30" customWidth="1"/>' +
      '<col min="5" max="5" width="28" customWidth="1"/>' +
      '<col min="6" max="6" width="52" customWidth="1"/>' +
      '<col min="7" max="7" width="70" customWidth="1"/>' +
      '</cols>' +
      '<sheetData>' + rows.join('') + '</sheetData></worksheet>';
  }

  /**
   * fileEntry: một phần tử của Pipeline.collectSpots()
   * bytes:     nội dung gốc của đúng file đó
   * Trả { bytes, painted, listed, sheetTitle, report }
   */
  function build(bytes, fileEntry, opts) {
    opts = opts || {};
    if (typeof fflate === 'undefined') throw new Error('Thiếu thư viện fflate');
    var files = fflate.unzipSync(bytes);
    var names = Object.keys(files);
    var report = [];
    function readText(n) { return fflate.strFromU8(files[n]); }
    function writeText(n, s) { files[n] = fflate.strToU8(s); }

    /* --- 1. hai style: vàng nhạt cho ô lỗi, vàng đậm cho dòng tiêu đề --- */
    var xfId = null, headerXfId = null, variants = null;
    if (files['xl/styles.xml']) {
      var a = addHighlightStyle(readText('xl/styles.xml'), opts.argb || 'FFFFF2A8');
      if (a) {
        var b = addHighlightStyle(a.xml, 'FFFFD966');
        xfId = a.xfId;
        headerXfId = b ? b.xfId : a.xfId;
        variants = highlightVariants(b ? b.xml : a.xml, a.fillId);
        writeText('xl/styles.xml', variants.xml());
      }
    }
    if (xfId === null) report.push('Không thêm được màu đánh dấu — chỉ kèm sheet liệt kê.');

    /* --- 2. bản đồ tên sheet -> part --- */
    var wbXml = readText('xl/workbook.xml');
    var relsXml = readText('xl/_rels/workbook.xml.rels');
    var rel = {};
    (relsXml.match(/<Relationship\b[^>]*\/>/g) || []).forEach(function (t) {
      var id = getAttr(t, 'Id'), tgt = getAttr(t, 'Target');
      if (id && tgt) rel[id] = tgt.charAt(0) === '/' ? tgt.slice(1) : 'xl/' + tgt.replace(/^\.\//, '');
    });
    var sheetPart = {};
    (wbXml.match(/<sheet\b[^>]*\/>/g) || []).forEach(function (t) {
      var nm = (getAttr(t, 'name') || '')
        .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
      var id = getAttr(t, 'r:id') || getAttr(t, 'id');
      if (nm && rel[id]) sheetPart[nm] = rel[id];
    });

    /* --- 3. tô màu từng ô --- */
    var painted = 0, missSheets = [];
    if (xfId !== null) {
      Object.keys(fileEntry.sheets).forEach(function (sheetName) {
        var part = sheetPart[sheetName];
        if (!part || !files[part]) { missSheets.push(sheetName); return; }
        var byRow = {};
        fileEntry.sheets[sheetName].forEach(function (sp) {
          if (!sp.row || sp.col === null || sp.col === undefined) return;
          (byRow[sp.row] = byRow[sp.row] || []).push(sp.col);
        });
        if (!Object.keys(byRow).length) return;
        var res = paintCells(readText(part), byRow, variants.forStyle);
        writeText(part, res.xml);
        painted += res.painted;
      });
    }
    if (variants) writeText('xl/styles.xml', variants.xml());
    if (missSheets.length) report.push('Không tìm thấy sheet: ' + missSheets.join(', '));

    /* --- 4. thêm sheet liệt kê vào ĐẦU workbook --- */
    var newPart = 'xl/worksheets/sheetMarkup.xml', guard = 0;
    while (files[newPart] && guard++ < 50) newPart = 'xl/worksheets/sheetMarkup' + guard + '.xml';

    var newRid = 'rIdMarkup';
    guard = 0;
    while (relsXml.indexOf('"' + newRid + '"') !== -1 && guard++ < 50) newRid = 'rIdMarkup' + guard;

    var title = SHEET_TITLE;
    guard = 0;
    while (sheetPart[title] && guard++ < 50) title = SHEET_TITLE + ' ' + guard;

    var sourceWb = opts.sourceWb || XLSX.read(bytes, {
      type: 'array', cellDates: false, cellNF: true, cellText: false, cellFormula: false, cellHTML: false
    });
    writeText(newPart, issueSheetXml(fileEntry, headerXfId, sourceWb));
    writeText('xl/_rels/workbook.xml.rels', relsXml.replace('</Relationships>',
      '<Relationship Id="' + newRid + '" Type="http://schemas.openxmlformats.org/' +
      'officeDocument/2006/relationships/worksheet" Target="worksheets/' +
      newPart.split('/').pop() + '"/></Relationships>'));

    /* sheetId phải không trùng với sheet nào đang có */
    var maxId = 0;
    (wbXml.match(/<sheet\b[^>]*\/>/g) || []).forEach(function (t) {
      var v = parseInt(getAttr(t, 'sheetId') || '0', 10);
      if (v > maxId) maxId = v;
    });
    wbXml = wbXml.replace(/<sheets>/, '<sheets><sheet name="' + esc(title) +
      '" sheetId="' + (maxId + 1) + '" r:id="' + newRid + '"/>');
    /* bỏ activeTab: chèn sheet vào đầu làm chỉ số cũ lệch, và mở ra thấy ngay
       danh sách lỗi mới là điều mong muốn */
    wbXml = wbXml.replace(/\s*activeTab="\d+"/, '');
    writeText('xl/workbook.xml', wbXml);

    writeText('[Content_Types].xml', readText('[Content_Types].xml').replace('</Types>',
      '<Override PartName="/' + newPart + '" ContentType="application/vnd.openxmlformats-' +
      'officedocument.spreadsheetml.worksheet+xml"/></Types>'));

    /* calcChain đánh số theo thứ tự sheet cũ — xoá để Excel tự dựng lại */
    if (files['xl/calcChain.xml']) {
      delete files['xl/calcChain.xml'];
      writeText('[Content_Types].xml',
        readText('[Content_Types].xml').replace(/<Override[^>]*calcChain\.xml[^>]*\/>/g, ''));
      writeText('xl/_rels/workbook.xml.rels',
        readText('xl/_rels/workbook.xml.rels').replace(/<Relationship[^>]*calcChain\.xml[^>]*\/>/g, ''));
    }

    var ordered = {};
    names.forEach(function (n) { if (files[n]) ordered[n] = files[n]; });
    Object.keys(files).forEach(function (n) { if (!ordered[n]) ordered[n] = files[n]; });

    return {
      bytes: fflate.zipSync(ordered, { level: 6, mtime: new Date() }),
      painted: painted,
      listed: fileEntry.count,
      sheetTitle: title,
      report: report
    };
  }

  return { build: build, numToCol: numToCol };
});
