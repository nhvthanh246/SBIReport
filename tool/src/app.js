/* Điều phối giao diện. Mọi logic nghiệp vụ nằm ở pipeline.js / writer.js / xlsxpatch.js. */
(function () {
  'use strict';

  var P = window.Parse, Pipe = window.Pipeline, W = window.Writer,
      ST = window.SelfTest, C = window.AppConfig, XP = window.XlsxPatch;
  var CFG_KEY = 'ccvn-report-tool.config.v1';

  /* Bốn loại file tool cần, theo thứ tự hiển thị. `multi` = nhận nhiều file. */
  var SLOTS = [
    { key: 'master', title: 'File báo cáo tổng', desc: 'Bản của ngày hôm trước' },
    { key: 'session', title: 'Session Report', desc: 'Chat bot — tải từ Live Support' },
    { key: 'ccvn', title: 'Daily Report CCVN', desc: 'Export case từ Salesforce' },
    { key: 'ops', title: 'Báo cáo ngày nhân viên', desc: 'SBI DAILY REPORT', multi: true }
  ];
  var TEST_SLOTS = [
    { key: 'expectedSession', title: 'Mẫu: Session-Report', desc: 'trong Data/output/' },
    { key: 'expectedCcvn', title: 'Mẫu: Daily Report CCVN', desc: 'trong Data/output/' }
  ];
  var COMPARE_SLOTS = [
    { key: 'reference', title: 'File tổng chuẩn', desc: 'bản đúng của ngày đang làm' }
  ];

  var state = {
    files: { session: null, ccvn: null, ops: [], master: null,
             expectedSession: null, expectedCcvn: null, reference: null },
    cfg: null,
    session: null, ccvn: null, ops: [], blocks: null, reportSerial: null,
    /* chữ ký bộ file lúc bấm Xử lý; khác chữ ký hiện tại = kết quả đã cũ */
    ranSignature: null, stale: false
  };

  var $ = function (id) { return document.getElementById(id); };

  function el(tag, cls, text) {
    var d = document.createElement(tag);
    if (cls) d.className = cls;
    if (text !== undefined && text !== null) d.textContent = text;
    return d;
  }
  function esc(s) {
    return String(s).replace(/[&<>]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c];
    });
  }

  /* ================= cấu hình ================= */

  function loadConfig() {
    var stored = null;
    try { stored = localStorage.getItem(CFG_KEY); } catch (e) { /* storage bị chặn */ }
    if (stored) { try { return JSON.parse(stored); } catch (e) { /* hỏng -> mặc định */ } }
    return C.defaults();
  }
  function saveConfig(cfg) {
    try { localStorage.setItem(CFG_KEY, JSON.stringify(cfg)); return true; }
    catch (e) { return false; }
  }
  function cfgStatus(text, kind) {
    var e = $('cfg-status');
    e.textContent = text;
    e.style.color = kind === 'err' ? 'var(--err)' : kind === 'ok' ? 'var(--ok)' : 'var(--ink-2)';
  }
  function paintConfigEditor() { $('cfg-text').value = JSON.stringify(state.cfg, null, 2); }
  function readConfigFromEditor() {
    var raw = $('cfg-text').value.trim();
    if (!raw) return null;
    try { var p = JSON.parse(raw); cfgStatus('', ''); return p; }
    catch (e) { cfgStatus('JSON không hợp lệ — giữ cấu hình đã lưu', 'err'); return null; }
  }

  /* ================= nạp file ================= */

  var READ_OPTS = { type: 'array', cellFormula: false, cellHTML: false,
                    cellDates: false, cellNF: false, cellText: false };

  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onerror = function () { reject(new Error('Không đọc được ' + file.name)); };
      fr.onload = function () {
        try {
          /* giữ byte gốc: ghi vào file tổng là sửa thẳng zip gốc chứ không dựng lại */
          var bytes = new Uint8Array(fr.result);
          resolve({ name: file.name, bytes: bytes, wb: XLSX.read(bytes, READ_OPTS) });
        } catch (e) { reject(new Error(file.name + ': ' + e.message)); }
      };
      fr.readAsArrayBuffer(file);
    });
  }

  function busy(on, text) {
    $('busy').classList.toggle('hidden', !on);
    if (text) $('busy-text').textContent = text;
  }

  /** Nhận file bất kỳ, tự đoán loại rồi xếp vào đúng chỗ. */
  function acceptFiles(fileList, forcedSlot) {
    var files = Array.prototype.slice.call(fileList)
      .filter(function (f) { return /\.(xlsx|xlsm|xls)$/i.test(f.name); });
    if (!files.length) return;

    busy(true, 'Đang đọc ' + files.length + ' file…');
    Promise.all(files.map(readFile)).then(function (loaded) {
      var unknown = [];
      loaded.forEach(function (f) {
        var kind = forcedSlot || Pipe.detectKind(f.wb);
        if (!kind) { unknown.push(f.name); return; }
        if (kind === 'ops') state.files.ops = state.files.ops.concat([f]);
        else state.files[kind] = f;
      });
      renderNeeds();
      busy(false);
      if (unknown.length) {
        showTopIssue('Không nhận ra loại file: ' + unknown.join(', '),
          'Kiểm tra lại file — tool nhận biết qua tên sheet và dòng header.');
      }
    }).catch(function (e) { busy(false); showTopIssue('Lỗi đọc file', e.message); });
  }

  /**
   * Danh sách file đã đổi sau khi Xử lý -> kết quả đang hiện thuộc về bộ file cũ.
   * Khoá mọi nút sinh file và nói rõ lý do, thay vì để người dùng tải nhầm bản cũ.
   */
  function markStale() {
    if (state.stale) return;
    state.stale = true;
    ['sec-result', 'sec-export'].forEach(function (id) { $(id).classList.add('is-stale'); });
    Array.prototype.forEach.call(
      document.querySelectorAll('#cta-master button, #outputs button'),
      function (b) { b.disabled = true; });
    $('btn-all').disabled = true;
    var box = $('stale-note');
    box.classList.remove('hidden');
    box.innerHTML = '';
    box.appendChild(issueEl('warn', 'Danh sách file đã thay đổi — kết quả bên dưới là của bộ file cũ',
      'Bấm Xử lý lại để dựng theo bộ file hiện tại. Các nút tải đã được khoá để tránh tải nhầm bản cũ.'));
  }

  function clearStale() {
    state.stale = false;
    ['sec-result', 'sec-export'].forEach(function (id) { $(id).classList.remove('is-stale'); });
    $('stale-note').classList.add('hidden');
    $('btn-all').disabled = false;
  }

  function showTopIssue(title, detail) {
    $('sec-result').classList.remove('hidden');
    var box = $('issues');
    box.insertBefore(issueEl('error', title, detail), box.firstChild);
  }

  /** Thẻ cho từng loại file cần nạp — là <label> bọc input ẩn, bấm đâu cũng mở hộp chọn. */
  function needEl(slot, host) {
    var f = state.files[slot.key];
    var on = slot.multi ? (f && f.length) : !!f;
    var d = el('label', 'need' + (on ? ' on' : ''));
    var input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xlsm,.xls';
    input.setAttribute('data-slot-input', slot.key);
    if (slot.multi) input.multiple = true;
    input.addEventListener('change', function () {
      acceptFiles(input.files, slot.key);
      input.value = '';
    });
    d.appendChild(input);
    d.appendChild(el('span', 'tick', '✓'));
    var body = el('span', 'nd-body');
    body.appendChild(el('span', 'nd-t', slot.title));
    var names = slot.multi
      ? (on ? f.map(function (x) { return x.name; }).join('\n') : slot.desc)
      : (on ? f.name : slot.desc);
    body.appendChild(el('span', 'nd-d', names));
    d.appendChild(body);
    if (on) {
      var x = el('button', 'nd-x', '✕');
      x.title = 'Bỏ file này';
      x.addEventListener('click', function (ev) {
        ev.preventDefault(); ev.stopPropagation();
        if (slot.multi) state.files[slot.key] = [];
        else state.files[slot.key] = null;
        renderNeeds();
      });
      d.appendChild(x);
    }
    host.appendChild(d);
  }

  /** Chữ ký của bộ file đang nạp — đổi chữ ký nghĩa là kết quả cũ không còn dùng được. */
  function fileSignature() {
    return SLOTS.map(function (s) {
      var f = state.files[s.key];
      if (s.multi) return s.key + ':' + (f || []).map(function (x) { return x.name; }).join(',');
      return s.key + ':' + (f ? f.name : '');
    }).join('|');
  }

  function renderNeeds() {
    /* Đổi file sau khi đã Xử lý thì kết quả trên màn hình thuộc về bộ file cũ.
       Phải khoá lại: nếu không, bấm Tải sẽ ra file dựng từ dữ liệu cũ mà không ai biết. */
    if (state.blocks && state.ranSignature !== null &&
        fileSignature() !== state.ranSignature) {
      markStale();
    }

    var host = $('needs');
    host.innerHTML = '';
    SLOTS.forEach(function (s) { needEl(s, host); });
    var host2 = $('needs-test');
    host2.innerHTML = '';
    TEST_SLOTS.forEach(function (s) { needEl(s, host2); });
    var host3 = $('needs-compare');
    host3.innerHTML = '';
    COMPARE_SLOTS.forEach(function (s) { needEl(s, host3); });
    $('btn-run').disabled = !state.files.session;

    /* có file rồi thì thu gọn vùng kéo thả */
    var any = SLOTS.some(function (s) {
      var f = state.files[s.key];
      return s.multi ? (f && f.length) : !!f;
    });
    var dz = $('dz');
    dz.classList.toggle('compact', any);
    dz.querySelector('.dz-main').textContent = any
      ? '+ Kéo thêm file vào đây' : 'Kéo cả 4 file vào đây';
  }

  function setupDropzone() {
    var dz = $('dz'), input = $('dz-input');
    input.addEventListener('change', function () { acceptFiles(input.files); input.value = ''; });
    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('over'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('over'); });
    });
    dz.addEventListener('drop', function (e) {
      if (e.dataTransfer && e.dataTransfer.files) acceptFiles(e.dataTransfer.files);
    });
  }

  /* ================= chạy pipeline ================= */

  function run() {
    if (!state.files.session) return;
    busy(true, 'Đang xử lý…');
    setTimeout(function () {
      try {
        state.cfg = readConfigFromEditor() || state.cfg;
        state.session = Pipe.processSession(state.files.session.wb, state.cfg, state.files.session.name);
        state.ccvn = state.files.ccvn
          ? Pipe.processCcvn(state.files.ccvn.wb, state.cfg, state.files.ccvn.name) : null;

        var serial;
        if (state.cfg.reportDate && state.cfg.reportDate !== 'auto') {
          serial = P.toDateSerial(state.cfg.reportDate);
        }
        if (!serial) {
          var det = Pipe.detectReportDate(state.session, state.cfg);
          serial = det ? det.serial : null;
        }
        if (!serial) throw new Error('Không suy được ngày báo cáo. Hãy chọn ngày thủ công.');
        state.reportSerial = serial;

        rebuild();
        $('sec-result').classList.remove('hidden');
        $('sec-export').classList.remove('hidden');
        $('sec-preview').classList.remove('hidden');
      } catch (e) {
        $('sec-result').classList.remove('hidden');
        $('issues').innerHTML = '';
        $('issues').appendChild(issueEl('error', 'Không xử lý được', e.message));
      }
      busy(false);
    }, 20);
  }

  function rebuild() {
    state.ops = state.files.ops.map(function (f) {
      return Pipe.processOpFile(f.wb, state.cfg, state.reportSerial, f.name);
    });

    /* Đọc bảng `mail SF` từ file nền TRƯỚC khi dựng khối: nó là bảng đang dùng thật,
       còn bảng trong cấu hình cũ dần theo tháng. Đồng thời soát tên trùng. */
    state.extraWarnings = [];
    state.mailSf = null;
    if (state.files.master) {
      state.mailSf = Pipe.inspectMailSf(state.files.master.wb, state.cfg, state.extraWarnings);
    }

    state.blocks = Pipe.buildBlocks({
      session: state.session, ccvn: state.ccvn, ops: state.ops,
      emailTable: state.mailSf && state.mailSf.table
    }, state.cfg, state.reportSerial);

    state.ranSignature = fileSignature();
    clearStale();
    renderResult();
    renderExport();
    renderPreviewPicker();
  }

  /* ================= render: kết quả ================= */

  var VI_DATE = ['Chủ nhật', 'Thứ hai', 'Thứ ba', 'Thứ tư', 'Thứ năm', 'Thứ sáu', 'Thứ bảy'];

  function renderResult() {
    var iso = state.blocks.reportISO;
    var dt = P.serialToDate(state.reportSerial);
    var dd = String(dt.getUTCDate()).padStart(2, '0');
    var mm = String(dt.getUTCMonth() + 1).padStart(2, '0');
    $('rh-date').textContent = 'Báo cáo ngày ' + dd + '/' + mm + '/' + dt.getUTCFullYear();
    $('rh-sub').textContent = VI_DATE[dt.getUTCDay()] + ' · serial Excel ' + state.reportSerial;
    $('report-date').value = iso;

    /* ---- số liệu ---- */
    var box = $('metrics');
    box.innerHTML = '';
    function metric(v, k, s, cls) {
      var d = el('div', 'metric' + (cls ? ' ' + cls : ''));
      d.appendChild(el('div', 'm-v', v));
      d.appendChild(el('div', 'm-k', k));
      if (s) d.appendChild(el('div', 'm-s', s));
      box.appendChild(d);
    }
    var nChat = state.session.staging.rows.length;
    var nOp = state.blocks.block1.rows.length - nChat;
    metric(state.blocks.block1.rows.length, 'Dòng vào “1. CSKH”',
      nChat + ' chat bot + ' + nOp + ' nhân viên');
    metric(state.blocks.block2.rows.length || '—', 'Case Salesforce',
      state.ccvn ? 'đã bỏ dòng Total' : 'chưa nạp file');
    metric(state.blocks.block3.rows.length || '—', 'Dòng giờ làm',
      state.files.ops.length + ' file nhân viên');
    metric(state.session.stats.dropped, 'Phiên đã lọc', 'theo luật trong cấu hình', 'is-muted');

    var swapped = state.ops.reduce(function (a, o) { return a + o.stats.swapped; }, 0);
    var accepted = !!(state.cfg.op && state.cfg.op.acceptSwappedDates);
    if (swapped) {
      metric(swapped, 'Dòng đảo ngày/tháng',
        accepted ? 'đã tính vào báo cáo' : 'đang bị bỏ sót', accepted ? '' : 'is-alert');
    }

    /* ---- ô tích nhận dòng đảo ngày ---- */
    var opt = $('opt-swapped');
    opt.classList.toggle('hidden', !swapped);
    opt.classList.toggle('hot', !!swapped && !accepted);
    if (swapped) {
      $('opt-swapped-d').textContent = accepted
        ? swapped + ' dòng đã được tính vào báo cáo.'
        : swapped + ' dòng đang bị bỏ sót. Lỗi locale lúc nhập liệu — tích vào để tính.';
    }

    renderIssues();
  }

  function issueEl(kind, title, detail) {
    var d = el('div', 'issue ' + kind);
    d.appendChild(el('span', 'ic', '●'));
    var b = el('span', 'ib');
    b.appendChild(el('span', 'it', title));
    if (detail) b.appendChild(el('span', 'id', detail));
    d.appendChild(b);
    return d;
  }

  function renderIssues() {
    var all = (state.extraWarnings || []).slice();
    [state.session, state.ccvn].concat(state.ops).forEach(function (r) {
      if (r && r.warnings) all = all.concat(r.warnings);
    });
    if (state.blocks) all = all.concat(state.blocks.warnings);

    var box = $('issues');
    box.innerHTML = '';
    var errors = all.filter(function (w) { return w.level === 'error'; });
    var warns = all.filter(function (w) { return w.level === 'warn'; });
    var infos = all.filter(function (w) { return w.level === 'info'; });

    if (!errors.length && !warns.length) {
      box.appendChild(issueEl('ok', 'Không có vấn đề cần xử lý',
        'Mọi cột khớp cấu hình, mọi tên supporter tra được email.'));
    }
    errors.forEach(function (w) { box.appendChild(issueEl('error', w.message, w.detail)); });

    /* cảnh báo nhẹ và ghi chú thì gấp lại, khỏi che mất lỗi thật */
    function fold(label, list, kind) {
      if (!list.length) return;
      var d = el('details', 'fold');
      d.appendChild(el('summary', null, label + ' (' + list.length + ')'));
      list.forEach(function (w) { d.appendChild(issueEl(kind, w.message, w.detail)); });
      box.appendChild(d);
    }
    fold('Cảnh báo', warns, 'warn');
    fold('Ghi chú', infos, 'ok');

    if (state.session.droppedRows.length) {
      var d = el('details', 'fold');
      d.appendChild(el('summary', null,
        'Phiên bị lọc bỏ (' + state.session.droppedRows.length + ')'));
      var list = el('div', 'scrolllist');
      state.session.droppedRows.forEach(function (x) {
        list.appendChild(el('div', null, 'Dòng ' + x.excelRow + ' — ' + x.preview));
      });
      d.appendChild(list);
      box.appendChild(d);
    }
  }

  /* ================= render: xuất file ================= */

  function cleanName(n) {
    var base = n.replace(/\.[^.]+$/, '');
    if (/^download/i.test(base)) base = base.replace(/^[^_]*_/, '');
    return base + '.xlsx';
  }
  function masterOutName(orig, iso) {
    var base = orig.replace(/\.[^.]+$/, '').replace(/^File\s*báo\s*cáo\s*tổng_/i, '');
    var ymd = iso.replace(/-/g, '');
    base = /-\d{8}$/.test(base) ? base.replace(/-\d{8}$/, '-' + ymd) : base + '-' + ymd;
    return base + '.xlsx';
  }

  function renderPatchReport(lines, err) {
    var box = $('patch-report');
    box.classList.remove('hidden');
    box.innerHTML = '';
    if (err) { box.appendChild(issueEl('error', 'Không ghi được vào File báo cáo tổng', err)); return; }
    box.appendChild(issueEl('ok', 'Đã tải xong file báo cáo tổng',
      'Mở bằng Excel — pivot tự refresh và công thức tự tính lại, nên lần mở đầu hơi lâu.'));
    var d = el('details', 'fold');
    d.appendChild(el('summary', null, 'Tool đã đụng vào những gì (' + lines.length + ' mục)'));
    var list = el('div', 'scrolllist');
    lines.forEach(function (l) { list.appendChild(el('div', null, l.replace(/^\s+/, ''))); });
    d.appendChild(list);
    box.appendChild(d);
  }

  function renderExport() {
    var iso = state.blocks.reportISO;

    /* ---- hành động chính ---- */
    var host = $('cta-master');
    host.innerHTML = '';
    var cta = el('div', 'cta' + (state.files.master ? '' : ' disabled'));
    var body = el('div', 'cta-body');
    if (state.files.master) {
      body.appendChild(el('div', 'cta-t', masterOutName(state.files.master.name, iso)));
      var d = el('div', 'cta-d');
      d.innerHTML = 'Đã điền 3 sheet dữ liệu và cột ngày trong <code>Summarize_team VietNam</code>. ' +
        'Pivot, công thức và định dạng của bản gốc giữ nguyên.';
      body.appendChild(d);
      var btn = el('button', 'primary', 'Tải file báo cáo tổng');
      btn.addEventListener('click', function () {
        /* Chặn trước khi làm gì: file có thể đã bị bỏ ra, hoặc bộ file đã đổi
           sau lần Xử lý gần nhất — lúc đó kết quả trên màn hình không còn đúng. */
        if (!state.files.master) {
          renderPatchReport(null, 'File báo cáo tổng đã bị bỏ ra. Hãy nạp lại rồi bấm Xử lý.');
          return;
        }
        if (!state.blocks) {
          renderPatchReport(null, 'Chưa có kết quả. Hãy bấm Xử lý trước.');
          return;
        }
        if (state.stale || fileSignature() !== state.ranSignature) {
          markStale();
          renderPatchReport(null, 'Danh sách file đã đổi từ lần Xử lý gần nhất. Bấm Xử lý lại trước khi tải.');
          return;
        }
        var old = btn.textContent;
        btn.textContent = 'Đang ghi…'; btn.disabled = true;
        setTimeout(function () {
          try {
            var res = XP.patch(state.files.master.bytes, state.blocks, state.cfg,
              { masterWb: state.files.master.wb });
            W.download(res.bytes, masterOutName(state.files.master.name, iso));
            renderPatchReport(res.report);
          } catch (e) { renderPatchReport(null, e.message); }
          btn.textContent = old; btn.disabled = false;
        }, 20);
      });
      cta.appendChild(body); cta.appendChild(btn);
    } else {
      body.appendChild(el('div', 'cta-t', 'Chưa có File báo cáo tổng'));
      body.appendChild(el('div', 'cta-d',
        'Nạp bản của ngày hôm trước ở bước 1 để tool điền thẳng dữ liệu vào đó.'));
      cta.appendChild(body);
    }
    host.appendChild(cta);

    /* ---- tệp phụ ---- */
    var box = $('outputs');
    box.innerHTML = '';
    function fileRow(name, metaHtml, actions) {
      var r = el('div', 'file-row');
      var i = el('div', 'fr-i');
      i.appendChild(el('div', 'fr-n', name));
      var m = el('div', 'fr-d'); m.innerHTML = metaHtml; i.appendChild(m);
      var a = el('div', 'fr-a');
      actions.forEach(function (act) {
        var b = el('button', 'sm', act.label);
        b.addEventListener('click', function () { act.run(b); });
        a.appendChild(b);
      });
      r.appendChild(i); r.appendChild(a);
      box.appendChild(r);
    }

    fileRow(cleanName(state.files.session.name),
      'Đã lọc <code class="k">' + state.session.stats.dropped + '</code> dòng, kèm sheet staging ' +
      '<code class="k">' + esc(state.session.staging.name) + '</code>',
      [{ label: 'Tải .xlsx', run: function () {
        W.download(W.buildSessionWorkbook(state.session, state.cfg),
          cleanName(state.files.session.name));
      } }]);

    if (state.ccvn) {
      fileRow(cleanName(state.files.ccvn.name),
        'Đã xoá cột rỗng <code class="k">' +
          (state.ccvn.removedColumns.join(', ') || 'không có') + '</code>',
        [{ label: 'Tải .xlsx', run: function () {
          W.download(W.buildCcvnWorkbook(state.ccvn, state.cfg), cleanName(state.files.ccvn.name));
        } }]);
    }

    [state.blocks.block1, state.blocks.block2, state.blocks.block3].forEach(function (b, i) {
      if (!b.rows.length) return;
      fileRow('Khối ' + (i + 1) + ' → ' + b.title,
        'Dán vào <code class="k">' + esc(b.sheet) + '</code> ô <code class="k">' + b.anchor +
          '</code> · ' + esc(b.note),
        [
          { label: 'Copy', run: function (btn) { copyText(W.blockToTsv(b, false), btn); } },
          { label: 'Tải .xlsx', run: function () {
            W.download(W.buildBlockWorkbook(b),
              'Khoi-' + (i + 1) + '_' + b.title.replace(/[^\w]+/g, '-') + '_' + iso + '.xlsx');
          } }
        ]);
    });
  }

  function copyText(text, btn) {
    var done = function () {
      var old = btn.textContent;
      btn.textContent = '✓ Đã copy';
      setTimeout(function () { btn.textContent = old; }, 1600);
    };
    var fallback = function () {
      var ta = document.createElement('textarea');
      ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy'); done(); }
      catch (e) { alert('Trình duyệt chặn clipboard. Hãy dùng nút Tải .xlsx.'); }
      document.body.removeChild(ta);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done, fallback);
    } else fallback();
  }

  /* ================= xem trước ================= */

  function renderPreviewPicker() {
    var sel = $('preview-pick');
    sel.innerHTML = '';
    ['block1', 'block2', 'block3'].forEach(function (k) {
      var b = state.blocks[k];
      if (!b.rows.length) return;
      var o = document.createElement('option');
      o.value = k; o.textContent = b.title;
      sel.appendChild(o);
    });
    var st = document.createElement('option');
    st.value = 'staging'; st.textContent = 'Sheet staging (Session Report)';
    sel.appendChild(st);
    renderPreview();
  }

  function renderPreview() {
    var k = $('preview-pick').value;
    var headers, rows;
    if (k === 'staging') { headers = state.session.staging.headers; rows = state.session.staging.rows; }
    else { headers = state.blocks[k].headers; rows = state.blocks[k].rows; }
    $('preview-count').textContent = rows.length + ' dòng · hiện 50 dòng đầu';

    var t = $('preview-table');
    t.innerHTML = '';
    var thead = document.createElement('thead'), tr = document.createElement('tr');
    headers.forEach(function (h) { tr.appendChild(el('th', null, h)); });
    thead.appendChild(tr); t.appendChild(thead);
    var tb = document.createElement('tbody');
    rows.slice(0, 50).forEach(function (row) {
      var r = document.createElement('tr');
      for (var i = 0; i < headers.length; i++) {
        var td = el('td', null, P.cellToString(row[i]));
        td.title = td.textContent;
        r.appendChild(td);
      }
      tb.appendChild(r);
    });
    t.appendChild(tb);
  }

  /* ================= đối chứng với bản đúng ================= */

  function renderChecks(box, checks) {
    checks.forEach(function (c) {
      var d = el('div', 'check ' + c.status);
      d.appendChild(el('div', 'mk',
        c.status === 'pass' ? 'ĐẠT' : c.status === 'fail' ? 'LỖI' : 'BỎ'));
      var b = el('div', 'cb');
      b.appendChild(el('div', 'cn', c.name));
      b.appendChild(el('div', 'cd', c.detail || ''));
      if (c.mismatches && c.mismatches.length) {
        var ul = document.createElement('ul');
        c.mismatches.forEach(function (m) { ul.appendChild(el('li', null, m)); });
        b.appendChild(ul);
      }
      d.appendChild(b);
      box.appendChild(d);
    });
  }

  function runCompare() {
    var box = $('compare-results');
    box.innerHTML = '';
    if (!state.blocks) {
      box.appendChild(issueEl('warn', 'Chưa có kết quả', 'Bấm Xử lý ở trên trước đã.'));
      return;
    }
    if (!state.files.master) {
      box.appendChild(issueEl('warn', 'Chưa nạp File báo cáo tổng', 'Cần file nền ở bước 1 để sinh ra bản so.'));
      return;
    }
    if (!state.files.reference) {
      box.appendChild(issueEl('warn', 'Chưa nạp file tổng chuẩn', 'Bấm vào ô “File tổng chuẩn” để chọn.'));
      return;
    }
    if (state.stale || fileSignature() !== state.ranSignature) {
      markStale();
      box.appendChild(issueEl('warn', 'Danh sách file đã đổi từ lần Xử lý gần nhất',
        'Bấm Xử lý lại trước khi đối chứng, nếu không sẽ so bằng dữ liệu cũ.'));
      return;
    }
    busy(true, 'Đang sinh file và đối chứng…');
    setTimeout(function () {
      try {
        var res = XP.patch(state.files.master.bytes, state.blocks, state.cfg,
          { masterWb: state.files.master.wb });
        /* Tên trùng phát hiện ở file nền cũng áp cho bản chuẩn — cùng một bảng `mail SF`. */
        var refDup = Pipe.inspectMailSf(state.files.reference.wb, state.cfg, []);
        var affected = {};
        [state.mailSf, refDup].forEach(function (x) {
          if (x) Object.keys(x.affected).forEach(function (e) { affected[e] = true; });
        });
        /* Lịch sử phải so với FILE NỀN chứ không phải bản chuẩn: người làm hay sửa
           số của ngày cũ lúc làm báo cáo hôm sau, nên bản chuẩn không còn là mốc
           đúng cho phần lịch sử. Việc tool phải bảo đảm là không tự đổi cái nó nhận. */
        var checks = ST.compareWithReference(res.bytes, state.files.reference.wb,
          state.blocks, state.cfg,
          { affectedEmails: affected, baseWb: state.files.master.wb,
            oddRateEmails: res.oddRateEmails });
        var nFail = checks.filter(function (c) { return c.status === 'fail'; }).length;
        box.appendChild(issueEl(nFail ? 'error' : 'ok',
          nFail ? nFail + ' mục KHÔNG khớp bản chuẩn' : 'Khớp bản chuẩn ở mọi mục tool ghi',
          'So với: ' + state.files.reference.name));
        renderChecks(box, checks);
      } catch (e) {
        box.appendChild(issueEl('error', 'Đối chứng lỗi', e.message));
      }
      busy(false);
    }, 20);
  }

  /* ================= tự kiểm chứng ================= */

  function runSelfTest() {
    var box = $('test-results');
    box.innerHTML = '';
    if (!state.files.session) {
      box.appendChild(issueEl('warn', 'Chưa có dữ liệu', 'Nạp ít nhất Session Report ở bước 1.'));
      return;
    }
    busy(true, 'Đang kiểm chứng…');
    setTimeout(function () {
      var checks;
      try { checks = ST.run(state.files, state.cfg); }
      catch (e) { busy(false); box.appendChild(issueEl('error', 'Kiểm chứng lỗi', e.message)); return; }

      var nFail = checks.filter(function (c) { return c.status === 'fail'; }).length;
      var nPass = checks.filter(function (c) { return c.status === 'pass'; }).length;
      box.appendChild(issueEl(nFail ? 'error' : 'ok',
        nFail ? nFail + ' kiểm tra KHÔNG đạt' : 'Tất cả ' + nPass + ' kiểm tra đều đạt',
        nFail ? 'Xem chi tiết ô lệch bên dưới.' : 'Pipeline tái tạo đúng kết quả mẫu.'));

      renderChecks(box, checks);
      busy(false);
    }, 20);
  }

  /* ================= bảng email từ file tổng ================= */

  function loadMailSfFromMaster() {
    if (!state.files.master) { cfgStatus('Chưa nạp File báo cáo tổng', 'err'); return; }
    var spec = state.cfg.mailSf || C.MAIL_SF;
    var ws = state.files.master.wb.Sheets[spec.sheet];
    if (!ws) { cfgStatus('File tổng không có sheet "' + spec.sheet + '"', 'err'); return; }
    var m = P.sheetToMatrix(ws), map = {}, seen = {}, dups = [], n = 0;
    for (var r = spec.startRow - 1; r < m.rows.length; r++) {
      var row = m.rows[r] || [];
      var name = P.cellToString(row[spec.nameColumn]).trim();
      var mail = P.cellToString(row[spec.emailColumn]).trim();
      if (!name || !mail) continue;
      var key = P.normName(name);
      /* VLOOKUP lấy dòng khớp ĐẦU TIÊN — giữ nguyên thứ tự đó, đừng để dòng sau đè lên,
         nếu không số case sẽ bị gán sang email khác so với công thức trong file. */
      if (seen[key]) {
        dups.push(name + ' (dòng ' + (r + 1) + ' → ' + mail + ', đã có → ' + seen[key] + ')');
        continue;
      }
      seen[key] = mail;
      map[name] = mail;
      n++;
    }
    if (!n) { cfgStatus('Sheet "' + spec.sheet + '" không đọc được dòng nào', 'err'); return; }
    state.cfg.supporterEmails = map;
    paintConfigEditor();
    saveConfig(state.cfg);
    cfgStatus(dups.length
      ? 'Đã nạp ' + n + ' tên → email. BỎ QUA ' + dups.length + ' dòng trùng tên: ' + dups.join('; ')
      : 'Đã nạp ' + n + ' tên → email', dups.length ? 'err' : 'ok');
  }

  /* ================= khởi động ================= */

  function init() {
    state.cfg = loadConfig();
    paintConfigEditor();
    $('accept-swapped').checked = !!(state.cfg.op && state.cfg.op.acceptSwappedDates);
    setupDropzone();
    renderNeeds();

    $('btn-run').addEventListener('click', run);
    $('btn-reset').addEventListener('click', function () { location.reload(); });
    $('btn-test').addEventListener('click', runSelfTest);
    $('btn-compare').addEventListener('click', runCompare);
    $('preview-pick').addEventListener('change', renderPreview);
    $('btn-all').addEventListener('click', function () {
      W.download(W.buildAllBlocksWorkbook(state.blocks), 'Khoi-dan_' + state.blocks.reportISO + '.xlsx');
    });

    $('accept-swapped').addEventListener('change', function (e) {
      state.cfg.op.acceptSwappedDates = e.target.checked;
      paintConfigEditor();
      if (!state.blocks) return;
      busy(true, 'Đang dựng lại…');
      setTimeout(function () { rebuild(); busy(false); }, 20);
    });

    $('btn-redate').addEventListener('click', function () {
      var serial = P.toDateSerial($('report-date').value);
      if (!serial) { showTopIssue('Ngày không hợp lệ', ''); return; }
      state.reportSerial = serial;
      busy(true, 'Đang dựng lại…');
      setTimeout(function () { rebuild(); busy(false); }, 20);
    });

    $('btn-cfg-save').addEventListener('click', function () {
      var c2 = readConfigFromEditor();
      if (!c2) return;
      state.cfg = c2;
      var ok = saveConfig(c2);
      cfgStatus(ok ? 'Đã lưu' : 'Trình duyệt chặn lưu — chỉ áp dụng trong phiên này', ok ? 'ok' : 'err');
    });
    $('btn-cfg-reset').addEventListener('click', function () {
      state.cfg = C.defaults();
      paintConfigEditor(); saveConfig(state.cfg);
      $('accept-swapped').checked = !!state.cfg.op.acceptSwappedDates;
      cfgStatus('Đã khôi phục mặc định', 'ok');
    });
    $('btn-cfg-export').addEventListener('click', function () {
      var blob = new Blob([JSON.stringify(state.cfg, null, 2)], { type: 'application/json' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'cau-hinh-bao-cao.json';
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 2000);
    });
    $('btn-cfg-import-proxy').addEventListener('click', function () { $('cfg-import').click(); });
    $('cfg-import').addEventListener('change', function (e) {
      var f = e.target.files[0];
      if (!f) return;
      var fr = new FileReader();
      fr.onload = function () {
        try {
          state.cfg = JSON.parse(fr.result);
          paintConfigEditor(); saveConfig(state.cfg);
          cfgStatus('Đã nhập cấu hình', 'ok');
        } catch (err) { cfgStatus('File JSON không hợp lệ', 'err'); }
      };
      fr.readAsText(f);
      e.target.value = '';
    });
    $('btn-cfg-mailsf').addEventListener('click', loadMailSfFromMaster);
  }

  document.addEventListener('DOMContentLoaded', init);
})();
