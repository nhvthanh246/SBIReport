(function () {
  'use strict';

  var READ_OPTS = { type: 'array', cellFormula: false, cellHTML: false,
                    cellDates: false, cellNF: false, cellText: false };
  var DATA = '/Data/';

  var days = [
    {
      day: '2026-09-13',
      base: 'test/File báo cáo tổng_AltiuslinkVN_Daily Report_(New)-20260912.xlsx',
      session: 'input/13092026/Session-Report-1789355016614.xlsx',
      ccvn: 'input/13092026/Daily Report CCVN-2026-09-14-12-03-49.xlsx',
      ops: ['input/13092026/File báo cáo ngày của nhân viên_SBI DAILY REPORT_0926.xlsx'],
      reference: 'input/13092026/File báo cáo tổng_AltiuslinkVN_Daily Report_(New)-20260913.xlsx'
    },
    {
      day: '2026-09-14',
      base: 'input/13092026/File báo cáo tổng_AltiuslinkVN_Daily Report_(New)-20260913.xlsx',
      session: 'input/14092026/Session-Report-14.xlsx',
      ccvn: 'input/14092026/Daily Report CCVN-14.09.xlsx',
      ops: ['input/SBI DAILY REPORT_0926.xlsx'],
      reference: 'input/14092026/AltiuslinkVN_Daily Report_(New)-20260914.xlsx'
    },
    {
      day: '2026-09-15',
      base: 'input/14092026/AltiuslinkVN_Daily Report_(New)-20260914.xlsx',
      session: 'input/15092026/Session-Report-15.xlsx',
      ccvn: 'input/15092026/Daily Report CCVN-15.09.xlsx',
      ops: ['input/SBI DAILY REPORT_0926.xlsx'],
      reference: 'input/15092026/AltiuslinkVN_Daily Report_(New)-20260915.xlsx'
    }
  ];

  function load(path) {
    return fetch(encodeURI(DATA + path)).then(function (r) {
      if (!r.ok) throw new Error(path + ': HTTP ' + r.status);
      return r.arrayBuffer();
    }).then(function (buf) {
      var bytes = new Uint8Array(buf);
      return { name: path.split('/').pop(), bytes: bytes, wb: XLSX.read(bytes, READ_OPTS) };
    });
  }

  function xmlErrors(bytes) {
    var files = fflate.unzipSync(bytes), bad = [];
    Object.keys(files).forEach(function (name) {
      if (!/\.xml$/.test(name) && !/\.rels$/.test(name)) return;
      var xml = fflate.strFromU8(files[name]);
      var doc = new DOMParser().parseFromString(xml, 'application/xml');
      if (doc.getElementsByTagName('parsererror').length) bad.push(name);
    });
    return bad;
  }

  function diffSummary(block, refWb) {
    var ws = refWb.Sheets[block.sheet];
    if (!ws) return [];
    var m = Parse.sheetToMatrix(ws), a0 = XLSX.utils.decode_cell(block.anchor), byCol = {};
    block.rows.forEach(function (row, i) {
      block.headers.forEach(function (header, c) {
        var got = Parse.cellToString(row[c]);
        var want = Parse.cellToString((m.rows[a0.r + i] || [])[a0.c + c]);
        if (got === want || (got === '' && want === '')) return;
        var d = byCol[c] || (byCol[c] = { header: header, count: 0, pairs: {}, samples: [] });
        d.count++;
        var pair = JSON.stringify(got) + ' → ' + JSON.stringify(want);
        d.pairs[pair] = (d.pairs[pair] || 0) + 1;
        if (d.samples.length < 12) {
          d.samples.push('dòng ' + (a0.r + i + 1) + ', khóa ' +
            JSON.stringify(Parse.cellToString(row[0])) + ': ' + pair);
        }
      });
    });
    return Object.keys(byCol).map(function (c) {
      var d = byCol[c];
      d.pairs = Object.keys(d.pairs).map(function (p) { return p + ' (' + d.pairs[p] + ')'; });
      return d;
    });
  }

  function runDay(spec) {
    return Promise.all([
      load(spec.base), load(spec.session), load(spec.ccvn),
      Promise.all(spec.ops.map(load)), load(spec.reference)
    ]).then(function (all) {
      var base = all[0], sessionFile = all[1], ccvnFile = all[2], opFiles = all[3], ref = all[4];
      var cfg = AppConfig.defaults();
      cfg.reportDate = spec.day;

      var kinds = [
        [Pipeline.detectKind(base.wb), 'master', base.name],
        [Pipeline.detectKind(sessionFile.wb), 'session', sessionFile.name],
        [Pipeline.detectKind(ccvnFile.wb), 'ccvn', ccvnFile.name]
      ];
      opFiles.forEach(function (f) { kinds.push([Pipeline.detectKind(f.wb), 'ops', f.name]); });
      kinds.forEach(function (x) {
        if (x[0] !== x[1]) throw new Error(x[2] + ': nhận diện ' + x[0] + ', cần ' + x[1]);
      });

      var serial = Parse.toDateSerial(spec.day);
      var session = Pipeline.processSession(sessionFile.wb, cfg, sessionFile.name);
      var ccvn = Pipeline.processCcvn(ccvnFile.wb, cfg, ccvnFile.name);
      var ops = opFiles.map(function (f) {
        return Pipeline.processOpFile(f.wb, cfg, serial, f.name);
      });
      var warnings = [];
      var mail = Pipeline.inspectMailSf(base.wb, cfg, warnings);
      var blocks = Pipeline.buildBlocks({
        session: session, ccvn: ccvn, ops: ops,
        emailTable: mail && mail.table
      }, cfg, serial);
      var patched = XlsxPatch.patch(base.bytes, blocks, cfg, { masterWb: base.wb });
      var refMail = Pipeline.inspectMailSf(ref.wb, cfg, []);
      var affected = {};
      [mail, refMail].forEach(function (m) {
        if (!m) return;
        Object.keys(m.affected).forEach(function (e) { affected[e] = true; });
      });
      var checks = SelfTest.compareWithReference(patched.bytes, ref.wb, blocks, cfg, {
        affectedEmails: affected,
        baseWb: base.wb,
        oddRateEmails: patched.oddRateEmails
      });
      var badXml = xmlErrors(patched.bytes);
      checks.push({
        status: badXml.length ? 'fail' : 'pass',
        name: '7. XML hợp lệ',
        detail: badXml.length ? badXml.join(', ') : 'Mọi XML part đều parse được'
      });
      return {
        day: spec.day,
        rows: [blocks.block1.rows.length, blocks.block2.rows.length, blocks.block3.rows.length],
        dropped: session.droppedRows.length,
        warnings: warnings.concat(session.warnings, ccvn.warnings,
          ops.reduce(function (a, o) { return a.concat(o.warnings); }, []), blocks.warnings),
        checks: checks,
        diagnostics: {
          block1: diffSummary(blocks.block1, ref.wb),
          block2: diffSummary(blocks.block2, ref.wb),
          block3: diffSummary(blocks.block3, ref.wb)
        },
        patchReport: patched.report
      };
    });
  }

  Promise.all(days.map(runDay)).then(function (results) {
    var failures = [];
    results.forEach(function (r) {
      r.checks.forEach(function (c) {
        if (c.status === 'fail') failures.push(r.day + ' — ' + c.name + ': ' + c.detail);
      });
    });
    var out = { ok: failures.length === 0, failures: failures, days: results };
    document.getElementById('result').textContent = JSON.stringify(out, null, 2);
    document.title = out.ok ? 'PASS' : 'FAIL';
    document.body.setAttribute('data-done', '1');
  }).catch(function (err) {
    document.getElementById('result').textContent = JSON.stringify({
      ok: false, fatal: err && (err.stack || err.message) || String(err)
    }, null, 2);
    document.title = 'ERROR';
    document.body.setAttribute('data-done', '1');
  });
})();
