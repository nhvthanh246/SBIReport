/* Cấu hình ánh xạ input -> output.
 * Mọi quy tắc nghiệp vụ nằm ở đây; sửa file này (hoặc màn hình Cấu hình trong tool)
 * là đủ khi web nguồn đổi tên cột / thêm luật lọc, không cần sửa code xử lý.
 *
 * Số dòng/cột trong file này đếm từ 1 giống Excel (headerRow: 10 = dòng 10).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.AppConfig = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var DEFAULT_CONFIG = {
    /* "auto" = suy ra từ ngày xuất hiện nhiều nhất ở cột First response time
       của Session Report. Hoặc đặt cứng "2026-09-13". */
    reportDate: 'auto',
    rulesVersion: 6,
    rulesSource: 'SBI 2026-10-23 — docs/VẤN ĐỀ.docx',

    /* --- Nguồn 1: Session Report (chat bot, tải từ FPT.AI Live Support) --- */
    session: {
      sheet: 0,
      headerRow: 1,
      /* SBI chốt Requested time cho Start Time. Ngày báo cáo vẫn nhận từ
         First response time; không suy ra giờ KPI từ độ dài phiên chat. */
      startTimeMappingVersion: 3,
      dateColumn: 'First response time',
      /* Dòng khớp BẤT KỲ luật nào dưới đây sẽ bị loại.
         op: equalsIgnoreCase | equals | contains | isEmpty | regex */
      /* Bắt cả "lỗi bot" lẫn "lỗi" — file tổng ngày 15 bỏ thêm 1 dòng ghi mỗi "lỗi".
         Kiểm chứng: ngày 13 -> 12 dòng, ngày 14 -> 3 dòng, ngày 15 -> 21 dòng,
         khớp đúng số dòng file tổng đã bỏ ở cả 3 ngày. */
      dropRules: [
        { column: 'Description', op: 'regex', value: '^\\s*lỗi\\s*(bot)?\\s*$' }
      ],
      /* Tên sheet staging = tên sheet gốc + hậu tố này */
      stagingSuffix: ' (2)',
      /* 10 cột của sheet staging. from: "const" -> dùng value; còn lại là tên cột nguồn. */
      stagingColumns: [
        { out: 'TYPE', from: 'const', value: 'CHAT BOT' },
        { out: 'Customer', from: 'Customer' },
        { out: 'Supporter', from: 'Supporter' },
        { out: 'Requested time', from: 'Requested time' },
        { out: 'End time', from: 'End time' },
        { out: 'Description', from: 'Description' },
        { out: 'Notes', from: 'Notes' },
        { out: 'Topic level 1', from: 'Topic level 1' },
        { out: 'Sender id', from: 'Sender id' },
        { out: 'Ma kh', from: 'Ma kh' }
      ]
    },

    /* --- Nguồn 2: Daily Report CCVN (export Salesforce) --- */
    ccvn: {
      sheet: 0,
      /* Dò dòng header bằng chuỗi này thay vì đặt cứng số dòng */
      headerMarker: 'Case Number',
      /* Xoá cột rỗng nằm trong vùng header (cột C trong mẫu) */
      dropEmptyColumns: true,
      /* Dòng tổng ở cuối: giữ trong file sạch, bỏ khi dựng khối dán */
      totalRowMarker: 'Total',
      /* Ánh xạ sang khối "2. SF Case info" */
      caseColumns: [
        { out: 'Case information/案件情報: Case Number/案件番号', from: 'Case Number' },
        { out: 'Customer information/顧客情報:取引先名', from: 'Customer information' },
        { out: 'Case Status/案件ステータス', from: 'Case Status' },
        /* Export Salesforce đổi định dạng giữa chừng (`10:09 14/09/2026` thay vì
           `2026/09/14 10:09`) — chuẩn hoá về dạng file tổng vẫn dùng. */
        { out: '案件作成日時', from: '案件作成日時', format: 'datetime' },
        { out: 'Requested By/案件登録者', from: 'Created By' }
      ],
      /* Cột dùng để tra email supporter */
      supporterNameColumn: 'Created By',
      /* Ký tự thừa ở cuối tên cần cắt trước khi tra email và trước khi ghi ra.
         Export có "Dinh Thi Kim Oanh." (dấu chấm cuối) trong khi `mail SF` ghi
         "DINH THI KIM OANH" — không cắt thì 35-50 case mỗi ngày không tra được email. */
      nameTrimPattern: '[.,;\\s]+$'
    },

    /* Bộ DataQuyTac được ưu tiên hơn các mẫu cũ: 19/19 case Quỳnh -> Anh,
       19/19 case Quang -> Nguyên; 6/6 case Hoài giữ nguyên. Không suy đoán
       cách chia 62 case Khánh Ly giữa hai email chỉ từ tên người lập. */
    requestedByAliases: {
      'VU NHU QUYNH': 'NGUYEN PHUONG ANH',
      'LUU NGOC QUANG': 'NGUYEN THAO NGUYEN',
      'PHAM THI THU HOAI': false
    },

    /* Điều chỉnh đã được xác nhận bằng output chuẩn mới, KHÔNG phải luật suy đoán.
       Khóa theo ngày + Case Number / email + cột + giá trị gốc, không theo số dòng.
       Có thể tắt toàn bộ hoặc sửa danh sách trong Cấu hình JSON. */
    approvedAdjustments: {
      enabled: false,
      source: '',
      cases: [],
      hours: []
    },

    /* Chỉ cảnh báo, KHÔNG sửa dữ liệu theo đáp án. Kết luận SBI ngày 23/10.
       Mục giờ được khóa theo ngày + email + cột + giá trị nguồn, không theo dòng. */
    sourceReviews: {
      customerFeedbackTopic: true,
      hours: [{ date: '2026-10-12', email: 'nguyen.nt@altius-link.com.vn',
        column: 6, from: 5, reference: 5.24,
        reason: 'SBI xác nhận nhân viên nhập sai giờ hỗ trợ trong bộ ngày 12/10. Cần sửa và lưu file OP đã xác nhận; tool không tự đổi 5 thành 5.24.' }]
    },

    /* --- Nguồn 3: file báo cáo ngày của nhân viên (SBI DAILY REPORT_MMYY) --- */
    op: {
      cskh: {
        sheet: '1. CSKH',
        headerRow: 10,
        dataStartRow: 12,
        dateColumn: 'Date/日付',
        /* Cột nguồn (0-based) đem sang khối 1, tương ứng cột C..N của đích */
        copyColumns: [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
        /* Cột luôn ghi trống, bất kể nguồn có gì (xem `blankColumns` ở `other`). */
        blankColumns: [13]
      },
      /* Bật mặc định theo bộ 11/10: 86 dòng CSKH + 16 dòng giờ bị đảo ngày/tháng.
       * Có thể tắt bằng checkbox nếu nguồn có dữ liệu ngày đảo hợp lệ khác.
       * Dù bật hay tắt, tool luôn báo số dòng bị ảnh hưởng để kiểm tra. */
      acceptSwappedDates: true,
      other: {
        sheet: '3. Other Tasks & Working time',
        headerRow: 5,
        dataStartRow: 7,
        dateColumn: 'DATE',
        /* Mỗi supporter chiếm đúng 1 khối liên tiếp bấy nhiêu dòng */
        blockSize: 8,
        /* Cột nguồn (0-based) đem sang khối 3, tương ứng cột C..M của đích.
           Index 2 là DATE - sẽ được ghi đè bằng serial ngày báo cáo. */
        copyColumns: [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
        /* Index 3 = cột TEAM. Cột này TRỐNG 100% trong cả 3 file tổng đã đối chiếu
           (152 + 192 + 192 dòng ở "3. Other Tasks", 2218 + 2577 + 2741 dòng ở
           "1. CSKH") — tức là cột chết, không ai dùng. Nhưng file OP nguồn lại có
           16 ô bị gõ nhầm vào đó (9 ô dấu ` và 7 ô số 6, đều do OP nhập nhầm phím).
           Chép nguyên si thì rác lọt vào báo cáo, nên luôn ghi trống.
           Bỏ index ra khỏi mảng này nếu sau này cột TEAM thật sự được dùng. */
        blankColumns: [3],
        /* DataQuyTac: số 0 ở cột M (giờ việc khác) được trình bày là ô trống.
           Chỉ đổi cách hiển thị 0 ở cột này, không xóa giờ hoặc điền giờ còn thiếu. */
        zeroAsBlankColumns: [12]
      }
    },

    /* --- Cột ngày trong sheet Summarize_team VietNam ---
     * Sheet này cộng dồn cả tháng, mỗi ngày một cột. Tool điền các chỉ tiêu tính
     * được từ "1. CSKH"; những dòng còn lại (tồn đọng Salesforce, Head Count)
     * phải nhập tay vì không suy ra được từ 3 nguồn đầu vào.
     * Đã đối chiếu với cột ngày 13/9 trong file mẫu: cả 8 giá trị đều khớp. */
    summary: {
      sheet: 'Summarize_team VietNam',
      dayHeaderRow: 4,
      /* Giá trị cột CHANNEL của "1. CSKH" gộp vào nhóm nào.
         So khớp không phân biệt hoa thường nên "(BO Team)" và "(BO team)" gộp chung. */
      channelGroups: {
        chatbot: ['CHAT BOT'],
        zalo: ['Zalo T', 'Zalo L'],
        viber: ['Viber T', 'Viber L'],
        line: ['Line T', 'Line L', 'Line'],
        tel: ['Tel/Hotline', 'Tel'],
        outbound: ['Outbound'],
        sfcase: ['Xử lý case (BO Team)'],
        comment: ['Trả lời Comment']
      },
      /* Kênh có thật nhưng không có dòng tương ứng trong bảng tổng hợp —
         liệt kê ở đây để khỏi bị cảnh báo "kênh lạ" mỗi lần chạy.
         `Setsumekai` (説明会 - buổi thuyết minh): file OP có 7 dòng, và dòng
         "Kênh hỗ trợ" của Summarize_team = đúng tổng 5 dòng 12-16, không cộng
         Setsumekai vào đâu cả -> đúng là kênh không tính. */
      ignoredChannels: ['JP khác', 'Outbound Reply', 'Gọi Outbound', 'Setsumekai'],
      /* Dòng nào lấy số nào. Dòng không có ở đây thì tool KHÔNG đụng tới. */
      cells: [
        { row: 11, from: 'customerFormula', label: '①CUSTOMER' },
        { row: 12, from: 'chatbot', label: '1. FACEBOOK (CHAT BOT)' },
        { row: 13, from: 'zalo', label: '2. ZALO' },
        { row: 14, from: 'viber', label: '3. VIBER' },
        { row: 15, from: 'line', label: '4. LINE' },
        { row: 16, from: 'tel', label: '5. Telephone' },
        { row: 18, from: 'outbound', label: '② OUTBOUND' },
        { row: 28, from: 'sfcase', label: '⑨CCVN Case check&process' },
        { row: 57, from: 'comment', label: '② Số lượng comment đã hỗ trợ' }
      ],

      /* --- Khối "Nội dung hỗ trợ" (dòng 33-45) ---
       * Đếm theo cột `Topic Level 1` của "1. CSKH". Giá trị trong cột đó luôn có
       * dạng "<số>. <tên tiếng Việt>/<tên tiếng Nhật>", nên khớp theo SỐ ĐẦU là
       * chắc chắn nhất — tên có thể đổi chính tả, số thì không.
       * Chú ý thứ tự KHÔNG trùng nhau: dòng 33 (① Rakuraku đổi sang APP) ứng với
       * topic số 13, còn dòng 34 mới là topic số 1.
       * Đã đối chiếu đủ 13 dòng × 3 ngày (13, 14, 15/9): khớp 39/39 ô.
       * Topic "14. NEOBANK - SSNB" không có dòng tương ứng nên không được cộng. */
      topicColumn: 9,
      topicRows: [
        { row: 33, topic: '13', label: '① KH Rakuraku đổi sang APP' },
        { row: 34, topic: '1', label: '② Đăng ký mới' },
        { row: 35, topic: '2', label: '③ Thêm người nhận' },
        { row: 36, topic: '3', label: '④ Tình trạng thẻ' },
        { row: 37, topic: '4', label: '⑤ Phát hành lại thẻ' },
        { row: 38, topic: '5', label: '⑥ Đổi thông tin thẻ' },
        { row: 39, topic: '6', label: '⑦ Tình trạng giao dịch' },
        { row: 40, topic: '7', label: '⑧ Cập nhật thẻ ngoại kiều' },
        { row: 41, topic: '8', label: '⑨ Tỷ giá/Lệ phí' },
        { row: 42, topic: '9', label: '⑩ Giấy CNCT' },
        { row: 43, topic: '10', label: '⑪ Gửi My number' },
        { row: 44, topic: '11', label: '⑫ Gặp nhân viên SBI' },
        { row: 45, topic: '12', label: '⑬ Đánh giá KH (claim/khen chê)' }
      ],
      /* Topic có thật nhưng không có dòng trong bảng — khỏi cảnh báo mỗi lần chạy. */
      ignoredTopics: ['14'],
      /* ①CUSTOMER = số sender_id DUY NHẤT của CHAT BOT, cộng các kênh dòng 13-16.
         Đúng như công thức có sẵn trong file: =1378+SUM(R13:R16) với 1378 là số
         sender_id duy nhất của ngày 13/9. */
      customerFormula: '{distinct}+SUM({col}13:{col}16)',
      /* Các dòng tool không điền được — hiện trong cảnh báo để nhập tay. */
      manualRows: {
        '6': 'Đăng ký mới — Tổng',
        '7': '① APP', '8': '② WEBFORM', '9': '③ Kênh khác (CCC)',
        '17': '②Unprocessed/未対応',
        '20': '①CCVN On-processed', '21': '②CCVN Processing request (BM)',
        '22': '③CCVN Processing request', '23': '④CCVN Processing request',
        '24': '⑤CCVN Processing request (Other)', '25': '⑥CCVN Unprocessed',
        '26': '⑦CCVN waiting-response', '27': '⑧CCVN Completed',
        '29': 'Order HC', '30': 'Actual HC',
        /* Đã thử suy ra từ dữ liệu đầu vào nhưng KHÔNG khớp, nên để nhập tay:
           56 Others  454/601/774 — không khớp số dòng thiếu topic (276/178/308)
              cũng không khớp "JP khác" (47/0/43)
           58 khung giờ  20:30/20:35/20:00 — End Time muộn nhất là 20:09/21:18/21:30,
              tức đây là giờ chốt ca do người khai, không phải số đo từ dữ liệu
           59 số lượng OP  18/24/23 — số supporter khác nhau trong "1. CSKH" là
              22/28/28; nhãn ghi "không bao gồm TL, JP, senior" nên phải có danh
              sách chức danh mới tính được, mà không nguồn nào có
           60 claim  cả 3 ngày đều = 0 nên không đủ dữ liệu để kết luận nguồn */
        '56': 'Others', '58': '③ khung giờ hỗ trợ', '59': '④ Số lượng OP',
        '60': '⑤ Số lượng claim/ý kiến',
        '46': 'Ý kiến KH — 6 dòng 46-51',
        '53': '① Số lượng OP hỗ trợ Outbound', '54': '② Số lượng KH Outbound',
        '55': '③ Nội dung Outbound',
        /* 61-66 fanpage: tool tự điền, xem khối `fanpage` bên dưới.
           Follow lấy từ extension, Like chép của ngày gần nhất. */
      }
    },

    /* --- 6 dòng fanpage (61-66) ---
     * FOLLOW (62/64/66): số thật, nhích lên mỗi ngày. Lấy bằng extension
     *   "Facebook Fanpage Stats - AltiuslinkVN" trong thư mục `extension/`:
     *   bấm "Lấy số liệu hôm nay" -> "Copy" -> dán vào ô trong tool.
     *   Extension chép ra đúng 6 dòng theo thứ tự SBIR, SMILES, DCOM, trong đó
     *   3 dòng Like để trống — khớp y hệt 6 dòng này của Excel.
     *
     * LIKE (61/63/65): phía SBI xác nhận chỉ cần ước lượng theo các ngày trước,
     *   không cần chính xác tuyệt đối. Đối chiếu dữ liệu thật: 3 giá trị này
     *   GIỮ NGUYÊN suốt cả tháng (382k / 111k / 311k), tức là người làm vẫn
     *   đang chép lại của hôm trước. Tool làm đúng như vậy: lấy giá trị gần
     *   nhất ở các cột ngày bên trái. Nếu giá trị đó vừa đổi so với ngày liền
     *   trước nữa thì tool cảnh báo — chính nhờ luật này mà phát hiện được
     *   DCOM bị gõ "3111k" (thừa số 1) từ ngày 14/9, trong khi cả tháng là "311k". */
    fanpage: {
      pages: [
        { key: 'SBIR', label: 'SBIR', likeRow: 61, followRow: 62 },
        { key: 'SMILES', label: 'SMILES', likeRow: 63, followRow: 64 },
        { key: 'DCOM', label: 'DCOM', likeRow: 65, followRow: 66 }
      ],
      /* Lấy Like của ngày gần nhất bên trái. Đặt false nếu muốn tự nhập tay. */
      carryLike: true,
      /* Quét tối đa bao nhiêu cột sang trái để tìm giá trị Like gần nhất */
      carryLookback: 40
    },

    /* --- Sheet "Summarize" ---
     * Không có dữ liệu riêng: dòng 7-27 toàn công thức gương
     * `='Summarize_team VietNam'!<cột ngày><dòng+4>`. Nhưng công thức chỉ có đến
     * cột của ngày CUỐI CÙNG đã làm — cột ngày mới trống trơn, phải kéo sang phải.
     * Không kéo thì sheet này trắng gần hết dù Summarize_team đã đủ số.
     * Dòng 6, 15, 25 đã trải hết chiều ngang sẵn nên tool bỏ qua (không ghi đè). */
    summarize: {
      sheet: 'Summarize',
      dayHeaderRow: 4,
      firstRow: 6,
      lastRow: 27
    },

    /* --- KPI: tự tính lại chuỗi Sheet5 -> Daily KPI Result -> 202609 KPI Result ---
     *
     * Cột "Hiệu suất công việc" (W) của Daily KPI Result là công thức, chỉ có giá trị
     * sau khi Excel tính lại — nên không copy sang sheet KPI tháng được. Tool tự dựng
     * lại cả chuỗi từ dữ liệu thô:
     *
     *   pivot1 (đếm CHANNEL theo email, từ "1. CSKH")
     *     HTKH = Grand Total - case - CMT - JP khác      (cột OB/DKM của pivot luôn rỗng)
     *   pivot2 (đếm case theo Supporter email, từ "2. SF Case info")
     *   pivot3 (tổng giờ theo email, từ "3. Other Tasks")
     *
     *   E=HTKH   F=số case   G=max(0,E-F)   H=F*1.5+G
     *   mỗi mảng: hiệu suất*giờ = count/rate khi giờ>0, ngược lại 0
     *   W = tổng(count/rate) / tổng(giờ)
     *
     * Đã đối chiếu ngày 13/9: khớp 32/32 người, mọi cột trung gian đúng. */
    kpi: {
      /* Sheet5!V = số buổi Setsumekai * 45 (bản chuẩn 14/9).
         O cộng V ngoài số dòng đã nằm trong Grand Total, theo công thức mẫu. */
      setsumekaiWeight: 45,
      /* 5 mảng KPI: [nguồn số lượng, cột giờ trong khối 3, định mức/giờ] */
      tracks: [
        { key: 'cskh', count: 'H', hoursColumn: 4, rate: 15 },
        { key: 'outbound', count: 'OB', hoursColumn: 5, rate: 30 },
        { key: 'case', count: 'case', hoursColumn: 6, rate: 17 },
        { key: 'comment', count: 'CMT', hoursColumn: 7, rate: 20 },
        { key: 'dkm', count: 'DKM', hoursColumn: 8, rate: 20 }
      ],
      /* Giá trị CHANNEL bị trừ khỏi HTKH (đúng công thức =I-P-Q-R-S-T-U+V của Sheet5) */
      caseChannels: ['Xử lý case (BO Team)'],
      commentChannels: ['Trả lời Comment'],
      otherChannels: ['JP khác'],
      /* Sheet KPI ngày — chỉ đọc để lấy danh sách người và cập nhật ô ngày.
       *
       * Hai danh sách khác nhau, đừng gộp làm một:
       *   KHỐI OP = tập dòng của các công thức SUM ở totalRow trong file nền.
       *       Mẫu tháng 09: 15–43; mẫu tháng 10: 15–42. Không lấy biên cố định.
       *   dataStartRow..peopleEndRow = TOÀN BỘ người có dòng KPI. Sheet còn người
       *       ở dòng 44-54 (DO PHUONG LINH, DO TU UYEN, LE VAN LONG, DUONG HUY,
       *       ly.ntk1, TRAN LAN HUONG, NGUYEN THAO NGUYEN, NGUYEN PHUONG ANH,
       *       NGO QUOC BAO) — họ nằm ngoài tổng OP nhưng vẫn có công thức KPI đầy
       *       đủ và vẫn phải được ghi sang sheet KPI tháng.
       *
       * Trước đây dùng chung một biên 43 nên 11 người này bị bỏ trắng ở sheet tháng.
       * Đã đối chiếu: với 4 người có đủ dữ liệu trong file OP (linh.dp1, uyen.dt,
       * long.lv, huy.d) tool tính ra ĐÚNG số của bản chuẩn ở cả 3 ngày 13/14/15. */
      daily: {
        sheet: 'Daily KPI Result',
        /* ô ngày của sheet này đã nằm trong `dateHeaderCells` */
        emailColumn: 3,
        dataStartRow: 15,
        /* Biên dự phòng cho danh sách cá nhân nếu thiếu peopleEndRow.
           Phạm vi tổng team luôn đọc từ công thức SUM trong file nền. */
        dataEndRow: 43,
        peopleEndRow: 60,
        totalRow: 14,
        /* Đọc định mức riêng từ J/M/P/S/V và nguồn số lượng H/K/N/Q/T.
         * JP dùng 10/10/10/15/20, K lấy JP khác và Q lấy XL DKM.
         * Mẫu công thức đã hiểu được tính riêng để ghi KPI tháng; mẫu lạ báo lại. */
        rateColumns: [9, 12, 15, 18, 21]
      },
      /* Sheet KPI tháng — tên đổi theo tháng (202609, 202610…) nên dò bằng mẫu */
      monthly: {
        sheetPattern: '^\\d{6} KPI Result$',
        dateCell: 'A2',
        dayHeaderRow: 5,
        emailColumn: 3,
        dataStartRow: 6,
        dataEndRow: 60,
        teamRow: 7
      }
    },

    /* Công thức giữ lại khi ghi thẳng vào File báo cáo tổng.
     * {row} = số dòng hiện tại, {rowEnd} = dòng cuối của khối supporter.
     * Lấy đúng công thức đang có sẵn trong file tổng; sửa ở đây nếu file đổi. */
    masterFormulas: {
      sfcaseEmail: "VLOOKUP($F{row},'mail SF'!$C$3:$D$69,2,0)",
      otherTotal: 'SUM(G{row}:M{rowEnd})'
    },

    /* --- Ô ngày ở phần ĐẦU mỗi tab ---
     * Mỗi tab có một ô ghi ngày của báo cáo, nằm phía trên vùng dữ liệu. Vì nằm
     * ngoài ô neo nên trước đây tool không đụng tới -> file xuất ra vẫn mang ngày
     * của hôm trước ở đầu tab. Đã đối chiếu vị trí trên cả 3 file tổng 13, 14, 15/9.
     * Sheet KPI tháng không nằm ở đây vì tên sheet đổi theo tháng (xem `kpi.monthly`). */
    dateHeaderCells: [
      { sheet: 'Summarize', cell: 'C2' },
      { sheet: 'Summarize_team VietNam', cell: 'C2' },
      { sheet: '1. CSKH', cell: 'B2' },
      { sheet: '2. SF Case info', cell: 'B3' },
      { sheet: '3. Other Tasks & Working time', cell: 'C2' },
      { sheet: 'Daily KPI Result', cell: 'A2' }
    ],

    /* Sheet5: O:V tra kênh theo nhãn ở B4:N4 và phủ tới dòng 70.
     * Pivot có thể đổi thứ tự cột khi thêm/bớt kênh mỗi ngày. */
    sheet5: {
      sheet: 'Sheet5',
      /* Dòng đầu dữ liệu pivot. Header kênh nằm ở dòng 4. */
      templateRow: 5,
      columns: [14, 15, 16, 17, 18, 19, 20, 21],   /* O..V */
      /* Trần an toàn: VLOOKUP của "Daily KPI Result" chỉ tra tới Sheet5!$A$2:$S$73,
         và pivot kế tiếp bắt đầu ở A74 — nên không được kéo quá 70. */
      extendToRow: 70
    },

    /* --- Ô neo khi dán vào File báo cáo tổng --- */
    targets: {
      cskh: { sheet: '1. CSKH', anchor: 'A11' },
      sfcase: { sheet: '2. SF Case info', anchor: 'B7' },
      other: { sheet: '3. Other Tasks & Working time', anchor: 'C6' }
    },

    /* Cột buộc ghi kiểu text. ID 16-17 chữ số ghi kiểu số sẽ mất chính xác
       (26793145563657353 > 2^53); timestamp ghi kiểu số sẽ bị Excel đổi định dạng. */
    textColumns: [
      'Sender id', 'Ma kh', 'Case Number', 'Conversation id',
      'First response time', 'End time', 'Requested time', 'Assign time',
      '案件作成日時'
    ],

    /* Vị trí bảng tra tên -> email trong File báo cáo tổng. Tool đọc bảng này mỗi
       lần chạy để soát tên trùng — tên trùng làm VLOOKUP gán case sai người. */
    mailSf: {
      sheet: 'mail SF', startRow: 3, nameColumn: 2, emailColumn: 3,
      /* Không tự đổi tên trùng thành tài khoản mới. Cần khóa phân công case
         từ nguồn để tách hai email; giữ bảng tra và cảnh báo khi chưa có. */
      nameCorrections: []
    },

    /* Bảng tra tên -> email, lấy từ sheet ẩn `mail SF` của File báo cáo tổng.
       Tra không phân biệt hoa thường và gộp khoảng trắng, giống VLOOKUP của Excel. */
    supporterEmails: {
      'DINH THI KIM OANH': 'oanh.dtk@altius-link.vn',
      'NGUYEN THI THUY TIEN': 'tien.ntt@altius-link.com.vn',
      'HOANG THE PHONG': 'phong.ht@altius-link.com.vn',
      'NGUYEN THUY TRANG': 'trang.nt@altius-link.com.vn',
      'DANG THI BICH': 'bich.dt@altius-link.com.vn',
      'NGUYEN THI NGOC THU': 'thu.ntn@altius-link.com.vn',
      'NGUYEN THI THANH THUY': 'thuy.ntt1@altius-link.com.vn',
      'TRAN TRUNG DUC': 'duc.tt@altius-link.com.vn',
      'NGUYEN TIEN THANG': 'thang.nt@altius-link.com.vn',
      'PHAM THI THU HOAI': 'hoai.ptt@altius-link.vn',
      'PHUNG LAN HUONG': 'huong.pl@altius-link.com.vn',
      'LE NGOC THUY': 'thuy.ln@altius-link.com.vn',
      'NGUYEN QUANG HUY': 'huy.nq1@altius-link.vn',
      'VU THI HONG VAN': 'van.vth1@altius-link.com.vn',
      'DANG NHAT MINH': 'minh.dn@altius-link.com.vn',
      'NGO QUANG THAI': 'thai.nq@altius-link.com.vn',
      'DINH THU THUY': 'thuy.dt@altius-link.com.vn',
      'LE QUANG HUY': 'huy.lq2@altius-link.com.vn',
      'VU NHU QUYNH': 'quynh.vn@altius-link.com.vn',
      'DINH THI QUYNH NHU': 'quynh.dtn1@altius-link.com.vn',
      'TRAN PHUONG THAO': 'thao.tp@altius-link.com.vn',
      'BUI THI KIM LINH': 'linh.btk@altius-link.com.vn',
      'NGUYEN HOANG LONG': 'long.nh@altius-link.com.vn',
      'PHAN THI HONG THUY': 'thuy.nth@altius-link.com.vn',
      'TRAN HUYEN TRANG': 'trang.th@altius-link.com.vn',
      'NGO QUOC BAO': 'bao.nq@altius-link.com.vn',
      'LUU NGOC QUANG': 'quang.ln@altius-link.com.vn',
      'BUI NGOC ANH': 'anh.bn1@altius-link.com.vn',
      'HOANG THI CUC': 'cuc.ht@altius-link.com.vn',
      'TA THI TRANG': 'trang.tt1@altius-link.com.vn',
      'TRAN THUY DUONG': 'duong.tt@altius-link.com.vn',
      /* `mail SF` có tên này 2 lần (dòng 34 -> ly.ntk@, dòng 43 -> ly.ntk1@).
         VLOOKUP lấy dòng khớp ĐẦU TIÊN nên phải là ly.ntk@. Nên dọn trùng ở file nguồn. */
      'NGUYEN THI KHANH LY': 'ly.ntk@altius-link.com.vn',
      'DANG QUANG MINH': 'minh.dq@altius-link.com.vn',
      'TRAN HUU QUANG LONG': 'long.thq@altius-link.com.vn',
      'LE PHUONG LINH': 'linh.lp@altius-link.com.vn',
      'DO PHUONG LINH': 'linh.dp1@altius-link.com.vn',
      'DO TU UYEN': 'uyen.dt@altius-link.com.vn',
      'LE VAN LONG': 'long.lv@altius-link.com.vn',
      'NGHIEM THU TRANG': 'trang.nt4@altius-link.com.vn',
      'DUONG HUY': 'huy.d@altius-link.com.vn',
      'TRAN LAN HUONG': 'huong.tl@altius-link.com.vn',
      'NGUYEN THAO NGUYEN': 'nguyen.nt@altius-link.com.vn',
      'NGUYEN PHUONG ANH': 'anh.np@altius-link.com.vn'
    }
  };

  /* Vị trí sheet `mail SF` trong File báo cáo tổng, dùng cho nút "nạp lại bảng email" */
  var MAIL_SF = { sheet: 'mail SF', startRow: 3, nameColumn: 2, emailColumn: 3 };

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  /* Nâng các ánh xạ mặc định cũ; giữ mọi tuỳ chỉnh khác của người dùng. */
  function migrate(config) {
    var cfg = clone(config);
    var startVersion = cfg.session && cfg.session.startTimeMappingVersion || 0;
    var rulesVersion = cfg.rulesVersion || 0;
    function addMissing(target, defaults) {
      Object.keys(defaults).forEach(function (key) {
        if (target[key] === undefined) target[key] = clone(defaults[key]);
        else if (target[key] && defaults[key] && typeof target[key] === 'object' &&
                 !Array.isArray(target[key]) && !Array.isArray(defaults[key])) {
          addMissing(target[key], defaults[key]);
        }
      });
    }
    addMissing(cfg, DEFAULT_CONFIG);
    var other = cfg.op && cfg.op.other;
    if (other && JSON.stringify(other.copyColumns) === '[2,3,4,5,6,7,8,9,10,11]') {
      other.copyColumns.push(12);
    }
    var start = cfg.session.stagingColumns[3];
    if (startVersion < 2 && start && start.out === 'Requested time' && start.from === 'Requested time') {
      /* Chỉ hoàn nguyên đúng ánh xạ mặc định của bản sửa ngày 14 đã phát hành.
         Cấu hình v2 hoặc ánh xạ tuỳ chỉnh khác không bị ghi đè. */
      start.out = 'First response time';
      start.from = 'First response time';
    }
    if (startVersion < 2) cfg.session.startTimeMappingVersion = 2;
    if (rulesVersion < 3) {
      /* Nâng mặc định cũ theo bộ 11/10. Các cấu hình v3 tự chỉnh giữ nguyên. */
      if (cfg.requestedByAliases['VU NHU QUYNH'] === 'NGUYEN PHUONG ANH') cfg.requestedByAliases['VU NHU QUYNH'] = false;
      if (cfg.requestedByAliases['PHAM THI THU HOAI'] === 'TRAN LAN HUONG') cfg.requestedByAliases['PHAM THI THU HOAI'] = false;
      cfg.op.acceptSwappedDates = true;
      cfg.rulesVersion = 3;
    }
    if (rulesVersion < 4) {
      /* Chỉ tắt đúng luật mặc định cũ; không ghi đè tên đích do người dùng tự đặt. */
      if (cfg.requestedByAliases['LUU NGOC QUANG'] === 'NGUYEN THAO NGUYEN') cfg.requestedByAliases['LUU NGOC QUANG'] = false;
      cfg.rulesVersion = 4;
    }
    if (rulesVersion < 5) {
      /* Chỉ thay các mặc định đã phát hành; giữ tên đích tùy chỉnh khác.
         Cấu hình v5 tự tắt alias sẽ không bị bật lại khi tải trang. */
      ['VU NHU QUYNH', 'LUU NGOC QUANG'].forEach(function (name) {
        if (cfg.requestedByAliases[name] === false) cfg.requestedByAliases[name] = DEFAULT_CONFIG.requestedByAliases[name];
      });
      if (cfg.requestedByAliases['PHAM THI THU HOAI'] === 'TRAN LAN HUONG') cfg.requestedByAliases['PHAM THI THU HOAI'] = false;
      if (cfg.approvedAdjustments.source === 'ExtraData/Daily Report 11.10.2026/Output') {
        /* Giữ danh sách để kiểm tra lại, nhưng không còn tự áp dụng ngoại lệ cũ. */
        cfg.approvedAdjustments.enabled = false;
      }
      cfg.mailSf.nameCorrections = cfg.mailSf.nameCorrections.filter(function (rule) {
        return !(rule.email === 'ly.ntk1@altius-link.com.vn' && rule.from === 'NGUYEN THI KHANH LY' && rule.to === 'NGUYEN THI KHANH LY (3M)');
      });
      cfg.rulesVersion = 5;
      cfg.rulesSource = 'DataQuyTac';
    }
    if (rulesVersion < 6) {
      /* Quy chuẩn SBI mới thay nguồn Start Time cũ, kể cả lựa chọn đã lưu.
         Sau khi nâng v6, lựa chọn chủ động của người dùng không bị đổi lại. */
      cfg.session.stagingColumns[3] = { out: 'Requested time', from: 'Requested time' };
      cfg.session.startTimeMappingVersion = 3;
      if (cfg.approvedAdjustments.source === 'ExtraData/Daily Report 11.10.2026/Output') {
        cfg.approvedAdjustments.enabled = false;
      }
      cfg.rulesVersion = 6;
      cfg.rulesSource = DEFAULT_CONFIG.rulesSource;
    }
    return cfg;
  }

  return {
    DEFAULT_CONFIG: DEFAULT_CONFIG,
    MAIL_SF: MAIL_SF,
    defaults: function () { return clone(DEFAULT_CONFIG); },
    migrate: migrate,
    clone: clone
  };
});
