# Tool ánh xạ báo cáo ngày — CCVN / Altius Link VN

Web tool chạy **100% trong trình duyệt**. Nhận file tải từ web + file báo cáo ngày của nhân viên,
sinh ra file đã làm sạch và 3 khối dữ liệu dán thẳng vào *File báo cáo tổng*.
Không có server, không cài đặt, không kết nối mạng — dữ liệu khách hàng không rời khỏi máy.

## Dùng thế nào

Mở `index.html` bằng Chrome hoặc Edge (double-click là được), rồi:

1. **Kéo cả 4 file vào vùng thả** — thả một lượt, thứ tự nào cũng được.
   Tool tự nhận biết từng loại qua tên sheet và dòng header, không cần phân loại tay.
2. Bấm **Xử lý**.
3. Liếc phần số liệu; nếu có ô đỏ thì xử lý trước.
4. Bấm **Tải file báo cáo tổng** — ra luôn file hoàn chỉnh.

Nếu muốn tự dán thì mở **Tệp phụ** để lấy 3 khối riêng (`Copy` hoặc `Tải .xlsx`):

| Khối | Dán vào sheet | Ô neo |
|---|---|---|
| Khối 1 | `1. CSKH` | `A11` |
| Khối 2 | `2. SF Case info` | `B7` |
| Khối 3 | `3. Other Tasks & Working time` | `C6` |

## Giao diện

Ba khu, đi từ trên xuống một lần là xong: **Nạp file → Kết quả → Xuất file**.

Nguyên tắc dựng:

- **Một hành động chính duy nhất.** Nút *Tải file báo cáo tổng* nằm riêng một khối nổi bật;
  5 tệp phụ gấp lại trong mục *Tệp phụ* vì đa số ngày không cần đến.
- **Hiện dần theo nhu cầu.** Khu Kết quả và Xuất file chỉ hiện sau khi bấm Xử lý;
  vùng kéo thả tự thu nhỏ khi đã có file.
- **Phân cấp theo mức nghiêm trọng.** Lỗi đỏ hiện thẳng; cảnh báo và ghi chú gấp lại kèm số đếm,
  để lỗi thật không chìm giữa những thứ chỉ để tham khảo.
- **Số liệu trước, giải thích sau.** Bốn con số lớn (dòng vào, case, giờ làm, phiên đã lọc),
  mỗi số kèm một dòng nhỏ nói số đó từ đâu ra.
- **Trạng thái nhìn là biết.** Mỗi loại file một thẻ, xanh khi đã nhận, kèm tên file để đối chiếu.
- **Không bắt người dùng làm việc máy làm được.** Thả cả 4 file một lượt, tool tự phân loại.

Theo chế độ sáng/tối của máy, và co về một cột ở bề rộng điện thoại (đã đo ở viewport 400px:
`scrollWidth` đúng 400, không phần tử nào tràn).

## Cách ghi vào File báo cáo tổng

File tổng có 4 PivotTable, 4 pivotCache (~1,7 MB), conditionalFormatting, customXml và hàng nghìn
công thức ở các sheet khác. Đọc rồi ghi lại bằng thư viện Excel thông thường sẽ **mất phần lớn số đó**.

Vì vậy tool không dựng lại workbook mà **sửa thẳng XML bên trong file zip**: giải nén, chỉ thay
`<sheetData>` của 3 sheet dữ liệu, giữ nguyên byte của các part còn lại, rồi nén lại.

Kết quả đo trên bộ mẫu: **46/56 part giữ nguyên từng byte**. Chỉ 10 part đổi, đều là cố ý:

| Part | Vì sao đổi |
|---|---|
| `sheet6/7/8.xml` | 3 sheet dữ liệu được ghi mới |
| `sheet2.xml` | cột ngày trong `Summarize_team VietNam` (chỉ 8 ô, các ô khác giữ nguyên) |
| `workbook.xml` | bật `fullCalcOnLoad` để Excel tính lại công thức khi mở |
| `pivotCacheDefinition1-4.xml` | bật `refreshOnLoad` để pivot tự refresh (mỗi file +18 byte) |
| `[Content_Types].xml`, `workbook.xml.rels` | gỡ tham chiếu tới `calcChain.xml` |
| `calcChain.xml` | **đã xoá** — nó ghi thứ tự tính theo ô cũ, giữ lại sẽ làm Excel báo file hỏng. Excel tự dựng lại khi mở. |

Những thứ được giữ nguyên:

- **Định dạng ô**: dùng lại đúng thuộc tính `s` của ô cũ ở cùng vị trí, nên màu/border/format ngày không đổi.
- **Công thức trong vùng dữ liệu**: cột `G` của `2. SF Case info` vẫn là
  `=VLOOKUP($F7,'mail SF'!$C$3:$D$69,2,0)`; cột `F` của `3. Other Tasks` vẫn là `=SUM(G6:M13)`,
  được đánh lại số dòng cho từng khối supporter.
- **Dữ liệu ngày cũ được dọn**: dòng thừa phía dưới bị xoá giá trị nhưng giữ định dạng.
- **Hyperlink cũ bị gỡ**: Excel tự tạo link `mailto:` bám theo địa chỉ ô; sau khi thay dữ liệu chúng
  trỏ sang email của dòng khác nên tool bỏ đi.

Chuỗi được ghi kiểu `inlineStr` để không phải đụng vào `sharedStrings.xml` (660 KB, dùng chung với
mọi sheet khác). Vì vậy file XML bên trong to hơn bản gốc, nhưng file zip cuối vẫn xấp xỉ kích thước cũ.

> Nên giữ lại bản gốc. Tool luôn xuất ra file mới, không ghi đè file bạn nạp vào.

## File tổng là file cộng dồn cả tháng

Hiểu đúng chỗ này thì mới dùng tool đúng:

| Tầng | Sheet | Hành vi |
|---|---|---|
| Dữ liệu thô | `1. CSKH`, `2. SF Case info`, `3. Other Tasks` | **Thay mới mỗi ngày**, chỉ giữ đúng 1 ngày |
| Tổng hợp | `Summarize`, `Summarize_team VietNam` | **Cộng dồn cả tháng**, mỗi ngày thêm 1 cột |

Vì vậy **mỗi ngày phải nạp vào tool file của ngày hôm trước**, không phải file rỗng — nếu không
sẽ mất các cột ngày đã làm trong tháng. Tool giữ nguyên toàn bộ phần tổng hợp của những ngày trước.

Tool tự tìm cột ứng với ngày báo cáo trong `Summarize_team VietNam` rồi điền **8 chỉ tiêu tính được
từ `1. CSKH`** (đã đối chiếu với cột ngày 13/9 trong file mẫu, khớp cả 8):

| Dòng | Chỉ tiêu | Nguồn |
|---|---|---|
| 11 | ①CUSTOMER | số `sender_id` **duy nhất** của CHAT BOT, giữ nguyên dạng công thức `=1378+SUM(R13:R16)` |
| 12 | 1. FACEBOOK (CHAT BOT) | đếm CHANNEL = `CHAT BOT` |
| 13 | 2. ZALO | `Zalo T` + `Zalo L` |
| 14 / 15 | 3. VIBER / 4. LINE | `Viber T/L` · `Line T/L` |
| 16 | 5. Telephone | `Tel/Hotline` |
| 18 | ② OUTBOUND | `Outbound` |
| 28 | ⑨CCVN Case check&process | `Xử lý case (BO Team)` (gộp cả biến thể viết thường) |

**15 chỉ tiêu còn lại tool không đụng tới** và hiện trong cảnh báo mỗi lần chạy: số tồn đọng Salesforce
(dòng 20–27) lấy từ 4 báo cáo SF khác, Head Count (dòng 29–30) và Đăng ký mới (dòng 6–9) nhập tay.

## KPI — tool tự tính lại cả chuỗi

`Summarize` và `Daily KPI Result` **không cần tool ghi vào**: chúng là công thức thuần
(420 và 756 công thức), tự chạy khi tool ghi đúng `1. CSKH` rồi bật tính lại + refresh pivot.

`202609 KPI Result` thì khác — 657 hằng số, **không tham chiếu sheet nào**. Cột ngày của nó là
paste-values từ cột `W` (`Hiệu suất công việc`) của `Daily KPI Result`. Mà `W` là công thức, chỉ có
giá trị sau khi Excel tính, nên tool phải **dựng lại cả chuỗi** từ dữ liệu thô:

```
pivot1  đếm CHANNEL theo email           (từ khối 1)
        HTKH = Grand Total − case − CMT − JP khác
pivot2  đếm case theo Supporter email     (từ khối 2)
pivot3  tổng giờ theo email               (từ khối 3)

E = HTKH      F = số case      G = max(0, E−F)      H = F×1,5 + G
mỗi mảng:  hiệu suất × giờ = count / định mức   (khi giờ > 0, ngược lại 0)
W = Σ(count/định mức) / Σ(giờ)
```

Định mức: CSKH 15, Outbound 30, Xử lý case 17, Trả lời Comment 20, ĐKM 20.

Tool chỉ ghi cho người **có trong khối OP của `Daily KPI Result`** (dòng 15–43). Các dòng JP Staff
hoặc tài khoản ngoài khối đó do nguồn khác điền, tool không đụng vào. Dòng tổng team tính từ
tổng số lượng và tổng giờ của cả nhóm, không phải trung bình các `W`.

Tool cũng cập nhật ô ngày `A2` của cả hai sheet KPI.

## `mail SF` có tên trùng — tool từng tra sai

Sheet `mail SF` có **`NGUYEN THI KHANH LY` hai lần**: dòng 34 → `ly.ntk@`, dòng 43 → `ly.ntk1@`.
`VLOOKUP` lấy dòng khớp **đầu tiên**, nhưng tool lúc nạp bảng lại để dòng sau ghi đè dòng trước —
ngược hẳn, khiến 62 case bị gán sang email khác và sai KPI của 2 người. Đã sửa: **lấy dòng đầu tiên**,
và cảnh báo khi phát hiện tên trùng.

Chính file mẫu cũng đang mâu thuẫn ở chỗ này: 40 dòng cache ra `ly.ntk@`, 22 dòng cache ra `ly.ntk1@`
dù cùng một công thức — là kết quả cũ chưa tính lại. Khi mở file đã vá, Excel tính lại và cả 62 dòng
về `ly.ntk@`, đúng như tool.

### Tool cảnh báo ngay mỗi lần chạy

Mỗi lần bấm **Xử lý**, tool đọc sheet `mail SF` của file tổng vừa nạp và soát tên trùng. Nếu có,
một **cảnh báo đỏ hiện ngay đầu danh sách**, ghi rõ tên nào, ở dòng nào, ra email nào.

Cảnh báo này cũng nói luôn việc cần quyết, vì đây không phải câu hỏi kỹ thuật:

- **Một người hai tài khoản** → xoá dòng thừa trong `mail SF`, giữ email đang dùng.
- **Hai người trùng tên** → bảng `mail SF` hỏng về thiết kế: tra theo tên thì không phân biệt được,
  phải thêm cột phân biệt.

Những người bị ảnh hưởng cũng được dùng trong mục **Đối chứng**: chênh lệch ở đúng các dòng đó được
ghi nhận kèm nguyên nhân thay vì báo lỗi, nên lỗi thật không bị lẫn vào nhiễu.

## Những bẫy dữ liệu tool đã xử lý

Phát hiện khi chạy thử trên bộ 3 ngày thật. Nếu không xử lý, báo cáo sẽ sai mà không ai biết.

### 37.142 dòng chỉ có ô ngày

Trong `SBI DAILY REPORT`, ngày 15/9 có **37.886 dòng** — nhưng chỉ **744 dòng có dữ liệu thật**,
còn **37.142 dòng chỉ có mỗi ô ngày** (do kéo ô ngày xuống quá tay, dòng 12.492 → 49.633).

Tool bỏ qua dòng không có nội dung nào ở các cột được copy, và báo số lượng đã bỏ.
Không lọc thì báo cáo phình từ 2.741 lên gần 40.000 dòng.

### Salesforce đổi định dạng ngày giờ giữa chừng

| | Cột `案件作成日時` |
|---|---|
| Export 13/9 | `2026/09/13 11:33` |
| Export 14/9, 15/9 | `10:09 14/09/2026` |
| File tổng (mọi ngày) | `2026/09/14 10:09` |

Trước đây phải sửa tay. Tool nhận cả hai dạng và luôn ghi ra dạng file tổng đang dùng.

### Tên có dấu chấm thừa

Export ghi `Dinh Thi Kim Oanh.` còn `mail SF` ghi `DINH THI KIM OANH` — lệch một dấu chấm làm
**35 case ngày 14 và 15 case ngày 15 không tra được email**. Tool cắt ký tự thừa cuối tên
(cấu hình `ccvn.nameTrimPattern`).

### Bảng email đọc từ file nền

`mail SF` thay đổi theo tháng (ngày 13 có 43 tên, ngày 14 trở đi có 44). Tool đọc bảng **từ chính
file tổng vừa nạp** thay vì dùng bảng seed trong cấu hình, nên không bị cũ dần.

### Luật lọc phiên chat

Ngoài `lỗi bot`, file tổng ngày 15 còn bỏ 1 dòng ghi mỗi `lỗi`. Luật đổi thành biểu thức
`^\s*lỗi\s*(bot)?\s*$`, kiểm chứng đúng số dòng bị bỏ ở cả 3 ngày: **12 / 3 / 21**.

## Lỗi đảo ngày/tháng trong file nhân viên

File OP có **1.188/8.677 dòng (13,7%) bị đảo ngày và tháng**. Nguyên nhân là locale Excel: gõ
`9/1/2026` với ý 01/9, Excel kiểu Mỹ đọc thành `m/d` = 09/01 rồi đổi luôn ra số.

Ngày > 12 không dính lỗi (không có tháng 13 nên Excel để nguyên dạng chuỗi), nên **báo cáo ngày 13–31
không ảnh hưởng**. Ngày 1–12 thì mất dòng — ví dụ ngày 07/9 mất 31%, ngày 04/9 mất 23%.

Tool phát hiện bằng cách: dòng nào có ngày không khớp, nhưng **đảo ngày↔tháng lại ra đúng ngày báo cáo**
thì đánh dấu. Mặc định chỉ **cảnh báo đỏ** kèm số dòng; tích ô ở bước 2 thì tính luôn vào báo cáo.

Nên sửa ở gốc: định dạng cột `Date/日付` trong file OP thành Text, hoặc thống nhất nhập `yyyy-mm-dd`.

## Quy tắc nghiệp vụ đang áp dụng

Tất cả nằm trong `src/config.js`, sửa được qua màn hình **Cấu hình ánh xạ** trong tool
(lưu vào trình duyệt, xuất/nhập được ra `.json`).

- **Lọc phiên chat**: bỏ dòng có `Description` = `lỗi bot`.
- **Sheet staging**: 10 cột, `TYPE` = hằng số `CHAT BOT`.
- **Daily Report CCVN**: xoá cột rỗng nằm trong vùng header; giữ metadata và dòng `Total`.
- **Lọc file nhân viên**: theo ngày báo cáo, chấp nhận cả serial Excel lẫn text `d/m/yyyy`.
- **Tra email supporter**: thay công thức `VLOOKUP` sang sheet `mail SF`, không phân biệt hoa thường.
- **Đổi tên người lập case** (`requestedByAliases`) — xem mục dưới.

## Hai điều phát hiện khi đối chiếu file mẫu — cần xác nhận

**1. Tên tài khoản Salesforce khác tên OP thật.**
Cột `Requested By` của *File báo cáo tổng* khác cột `Created By` của export ở 38/379 case,
theo đúng 2 cặp và áp dụng cho 100% số dòng của hai tên đó:

| Tên trên export Salesforce | Tên trong file báo cáo tổng | Số case |
|---|---|---|
| `VU NHU QUYNH` | `NGUYEN PHUONG ANH` | 19 |
| `LUU NGOC QUANG` | `NGUYEN THAO NGUYEN` | 19 |

Hai tên đích (`anh.np@`, `nguyen.nt@`) cũng đúng là 2 supporter chỉ có trong file tổng mà không có
trong file OP mẫu — khớp với giả thiết hai bạn này lập case bằng tài khoản SF khác.
Tool đang áp dụng quy tắc này và **luôn hiện cảnh báo mỗi lần chạy**.
Nếu quy tắc sai hoặc đã thay đổi, sửa `requestedByAliases` trong Cấu hình.

**2. Lỗi gõ ngày trong file báo cáo ngày của nhân viên.**
Sheet `3. Other Tasks & Working time` có 80 dòng ghi ngày kiểu `13/09//2026` (hai dấu gạch chéo),
trong đó 8 dòng thuộc ngày báo cáo — tức nguyên khối giờ làm của một supporter (`ly.ntk@`).
Tool đọc được các giá trị này, nhưng nên sửa lại ở file nguồn.

## Cấu trúc

```
tool/
  index.html            giao diện
  styles.css
  vendor/
    xlsx.full.min.js    SheetJS 0.18.5 — đọc file nguồn, ghi file khối
    fflate.min.js       fflate 0.8.2 — giải nén/nén zip cho việc vá file tổng
  src/
    config.js           toàn bộ quy tắc nghiệp vụ
    parse.js            đọc sheet, chuẩn hoá ngày/chuỗi, dò cột
    pipeline.js         3 phép biến đổi + dựng 3 khối (thuần, không đụng DOM)
    writer.js           ghi .xlsx đúng kiểu ô + TSV cho clipboard
    xlsxpatch.js        vá thẳng XML trong File báo cáo tổng
    selftest.js         đối chiếu với file mẫu
    app.js              điều phối giao diện
```

Script viết kiểu classic + UMD, không dùng ES module, nên mở trực tiếp bằng `file://` vẫn chạy.

## Kiểm chứng trên chuỗi ngày thật (13→14→15)

Phép thử mạnh nhất: lấy file tổng ngày N làm nền, nạp nguồn ngày N+1, rồi so với file tổng
ngày N+1 do team làm tay. Chạy hai chuỗi liên tiếp với dữ liệu thật:

| Mục | 13/9 → 14/9 | 14/9 → 15/9 |
|---|---|---|
| `1. CSKH` | ✅ **36.064 ô khớp** | 1 ô lệch / 38.374 |
| `2. SF Case info` | 182 ô lệch / 3.584 | 144 ô lệch / 3.836 |
| `3. Other Tasks` | 13 ô lệch / 1.600 | 4 ô lệch / 1.760 |
| Cột ngày `Summarize_team` | 2 ô lệch / 8 | ✅ **khớp hết** |
| Cột ngày `202609 KPI` | 1 ô lệch / 30 | 3 ô lệch / 30 |
| **Lịch sử ngày trước** | ✅ **858 ô, không đụng** | ✅ **924 ô, không đụng** |

### Chênh lệch còn lại — đã truy nguyên nhân, không phải lỗi tool

- **Case Status (~130 ô/ngày)**: export Salesforce tôi có được chụp **muộn hơn** bản người làm dùng.
  Ngày 14: 96 case `Processing request` → `Complete`, 27 case `On processed` → `Complete`.
  Chuyển dịch gần như một chiều, đúng quy luật case tiến triển theo thời gian.
- **`3. Other Tasks`**: người làm sửa tay sau khi dán — xoá dấu `` ` `` lạc, đổi 5 → 5,25, 8,35 → 8.
- **`Summarize_team` ngày 14 dòng VIBER**: tool đếm 3, file tổng ghi 0. Chính `1. CSKH` của
  file tổng có 3 dòng `Viber T`, nên **số 0 là sai sót khi nhập tay**, tool đúng.
- **`1. CSKH` một ô Topic**: người làm đổi chủ đề của 1 dòng sau khi dán.

## Đối chứng với một bản đúng đã biết

Mục **Đối chứng với file tổng chuẩn** dùng khi muốn chứng minh tool làm đúng: nạp bản đúng của
ngày đang làm (ví dụ file team làm tay hôm đó), tool sinh file rồi so từng ô.

Chỉ so **những chỗ tool thực sự ghi** — 15 chỉ tiêu nhập tay không tính vào. Có thêm một mục
kiểm tra riêng: **các cột ngày trước có bị đụng không**.

### Bộ dữ liệu thử sẵn có trong `Data/test/`

| File | Là gì |
|---|---|
| `...Daily Report_(New)-20260912.xlsx` | File tổng mô phỏng **cuối ngày 12** — lịch sử ngày 1–12 nguyên vẹn, cột ngày 13 để trống, 3 sheet dữ liệu mang dữ liệu ngày 12 |
| `DOI-CHUNG_ban dung ngay 13.xlsx` | Bản đúng của ngày 13 để đối chứng |

Cách thử: nạp file ngày 12 + 3 file nguồn → **Xử lý** → nạp bản đối chứng → **Chạy đối chứng**.

Kết quả đo được:

| Mục | Kết quả |
|---|---|
| Sheet `1. CSKH` | **31.052 ô khớp** (2218 dòng × 14 cột) |
| Sheet `2. SF Case info` | **2.653 ô khớp**; 22 ô email bỏ qua do bản chuẩn giữ cache VLOOKUP cũ |
| Sheet `3. Other Tasks` | **1.280 ô khớp** |
| Cột ngày `Summarize_team VietNam` | 8/8 ô khớp |
| Cột ngày `202609 KPI Result` | 30/30; 2 ô lệch đã truy được nguyên nhân |
| **Lịch sử 12 ngày trước** | **792 ô, không ô nào bị đụng** |

Kết luận: **khớp bản chuẩn ở mọi mục tool ghi**.

2 ô lệch ở KPI tháng là `ly.ntk@` (dòng 34) và dòng tổng team kéo theo. Tool **không giấu chúng đi**
mà hiện ra kèm nguyên nhân cụ thể và không tính là lỗi, vì đã truy được gốc: tên trùng trong
`mail SF` làm bản chuẩn giữ số case cũ. Đã chứng minh **tool đúng, bản chuẩn sai** — mở file đã vá
bằng Excel, chính Excel tính ra `1,97103398044654` giống tool.

> File ngày 12 là **mô phỏng**: phần chat và case dùng lại dữ liệu ngày 13 vì bộ mẫu không có
> bản tải về của ngày 12. Điều đó không ảnh hưởng phép thử, vì 3 sheet dữ liệu bị ghi đè hoàn toàn
> khi làm báo cáo ngày 13 — ngược lại còn kiểm tra được việc **dọn dòng thừa** (giờ làm đi từ
> 144 dòng xuống 128 dòng).

## Kiểm chứng

Mở mục **Tự kiểm chứng với dữ liệu mẫu**, nạp thêm 2 file trong `Data/output/` (và *File báo cáo tổng*
ở bước 1), rồi bấm **Chạy kiểm chứng**. Tool chạy lại đúng pipeline và so từng ô.

Kết quả với bộ dữ liệu mẫu ngày 2026-09-13 — **49.391 ô khớp, 0 sai lệch**:

| Kiểm tra | Kết quả |
|---|---|
| Lọc `lỗi bot` | giữ 1555 / loại 12 / tổng 1567 |
| Sheet staging vs `Data/output/Session-Report-*.xlsx` | 15.550 ô khớp |
| Daily Report CCVN vs `Data/output/Daily Report CCVN-*.xlsx` | 3.292 ô khớp |
| Khối 1 vs `1. CSKH` của file tổng | 26.616 ô khớp, 2218/2218 dòng |
| Khối 2 vs `2. SF Case info` của file tổng | 2.653 ô khớp, 379/379 dòng |
| Tra email supporter | 379/379 dòng đúng bảng `mail SF` |
| Khối 3 vs `3. Other Tasks` của file tổng | 1.280 ô khớp, 16 supporter |
| Kiểu ô khi ghi | ID 15+ chữ số và timestamp giữ kiểu text |

Khối 3 ra 16/19 supporter vì bộ mẫu chỉ có 1 file OP; 3 supporter còn lại nằm ở file OP khác.
Nạp đủ file OP thì đủ 19.

### File báo cáo tổng xuất ra đã được mở bằng Excel thật để kiểm tra

| Kiểm tra | Kết quả |
|---|---|
| Excel mở file | không báo lỗi, không hỏi repair |
| Số sheet / pivotCache / PivotTable | 11 / 4 / 4 — **đúng bằng bản gốc** |
| 3 sheet dữ liệu đọc lại | 34.985 ô khớp, 0 lệch |
| `K11` (sender_id) | kiểu String `5855010741228717` — không mất chữ số |
| `B11` (Date) | hiện `13/09/2026`, value2 = `46278` |
| `2. SF Case info` `G7` | `=VLOOKUP($F7,'mail SF'!$C$3:$D$69,2,0)` → `trang.th@altius-link.com.vn` |
| `3. Other Tasks` `F6` / `F14` | `=SUM(G6:M13)` / `=SUM(G14:M21)` — đánh lại số dòng đúng từng khối |
| Pivot `Sheet5` sau refresh | `van.vth1`=175, `thuy.ln`=131 — khớp số phiên CHAT BOT |
| Dòng ngay sau vùng dữ liệu | đã rỗng (dữ liệu ngày cũ được dọn) |
| Cột ngày `Summarize_team VietNam` | R11–R28: 8/8 giá trị khớp; `R11` giữ công thức `=1378+SUM(R13:R16)` → 1422 |
| Dòng nhập tay cùng cột | `R17`=61, `R26`=3026, `R29`=20 — **không bị đụng** |
| Cột ngày khác | `Q12`=1687, `F12`=1403 — **không bị đụng** |
| `Summarize` sau khi tính lại | tổng tháng `E6`=25.904, `E8`=25.108 — đúng như bản gốc |
| Ô ngày 2 sheet KPI | `Daily KPI Result!A2` và `202609 KPI Result!A2` đều hiện `13/09/2026` |

### KPI: đối chiếu với chính Excel

Phép thử mạnh nhất — mở file đã vá bằng Excel rồi so cột `W` **Excel tự tính** với cột ngày
**tool ghi** vào sheet KPI tháng:

| Kiểm tra | Kết quả |
|---|---|
| Số người khớp | **34/36** |
| 2 người còn lại | `ly.ntk1@`, `nguyen.nt@` — tool **cố ý không ghi** vì nằm ngoài khối OP (dòng 15–43) |
| Dòng tổng team | Daily `W14` = `1,3547322775264` = 202609 dòng 7 — **khớp tuyệt đối** |
| `ly.ntk@` | Excel tính `1,97103398044654`, tool tính `1,971033980446537` — khớp, xác nhận tool đúng còn file gốc giữ cache cũ |
| Cột ngày khác của sheet KPI | 225 ô giữ nguyên, **0 ô bị đụng** |

Trước khi ghi vào file, chuỗi tính còn được đối chiếu với cột `W` có sẵn: **32/32 người khớp,
mọi cột trung gian (E, F, N, Q và 5 cột giờ) đều đúng**.

Có thử thêm trường hợp dữ liệu **dài hơn** vùng có sẵn (mô phỏng 3 file OP → khối 3 cần 384 dòng
trong khi sheet chỉ có 152): tool tạo thêm 232 dòng nhân bản định dạng từ dòng mẫu, 56.109 ô khớp,
Excel vẫn mở bình thường.

### Vì sao phải ép kiểu ô

`Sender id` có giá trị như `26793145563657353` — lớn hơn 2^53, ghi kiểu số sẽ sai chữ số cuối.
`No` và `Date` thì ngược lại, phải là số để công thức và pivot của file tổng chạy được.
Tool ép kiểu theo cột thay vì để thư viện tự suy; mục kiểm tra số 7 canh đúng chỗ này.

## Giới hạn đã biết

- SheetJS bản community không ghi định dạng ô, nên file xuất ra là bảng dữ liệu thuần:
  đúng nội dung nhưng không có màu/border như file mẫu. Dán vào file tổng thì định dạng
  của file tổng được giữ nguyên, nên không ảnh hưởng.
- Khi so sánh, mọi kiểu xuống dòng (`\n`, `\r\n`, `\r\r\n`) được coi là như nhau —
  Excel đổi cách mã hoá xuống dòng mỗi lần mở rồi lưu lại, đó là dấu vết của Excel chứ
  không phải khác biệt nội dung.
