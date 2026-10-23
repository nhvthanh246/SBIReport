// Explicit HISTORICAL business profile, for old regression fixtures only.
// This is never loaded by the application. Current SBI rules have separate tests.
module.exports = function historicalV4(cfg) {
  cfg.rulesSource = 'historical-v4-test-only';
  cfg.session.stagingColumns[3] = {out:'First response time',from:'First response time'};
  // Keep the current schema version so intentional settings survive UI migration.
  cfg.requestedByAliases = {'VU NHU QUYNH':false,'LUU NGOC QUANG':false,'PHAM THI THU HOAI':false};
  cfg.op.other.zeroAsBlankColumns = [];
  cfg.mailSf.nameCorrections = [{email:'ly.ntk1@altius-link.com.vn',from:'NGUYEN THI KHANH LY',to:'NGUYEN THI KHANH LY (3M)'}];
  cfg.approvedAdjustments = {
    enabled:true, source:'ExtraData/Daily Report 11.10.2026/Output',
    cases:['20261012218289','20261012218452','20261012218454','20261012218613','20261012218655'].map(caseNumber =>
      ({date:'2026-10-11',caseNumber,from:'VU NHU QUYNH',to:'PHAM THANH DAO'})),
    hours:[{date:'2026-10-11',email:'cuc.ht@altius-link.com.vn',column:6,from:0.941176470588235,to:0.94}]
  };
  return cfg;
};
