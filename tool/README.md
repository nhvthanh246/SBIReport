# Tool ánh xạ báo cáo ngày — CCVN / Altius Link VN

Xem [Luồng tính toán báo cáo](LUONG-TINH-TOAN.md) để biết nguồn của từng cột,
cách tính tổng hợp, giờ G–K, KPI cá nhân/nhóm và ví dụ cụ thể từ DataQuyTac.

## Quy tắc hiện hành theo xác nhận SBI ngày 23/10/2026

**Phiên bản 6** ưu tiên kết luận trong [VẤN ĐỀ.docx](<../docs/VẤN ĐỀ.docx>).
Phần không thay đổi tiếp tục theo DataQuyTac; không sửa file nguồn và không chép
output làm đáp án. [Luồng tính toán](LUONG-TINH-TOAN.md) đã được cập nhật tương ứng.

| Nội dung SBI chốt | Tool xử lý |
|---|---|
| Start Time là lúc khách gửi yêu cầu | Mặc định lấy `Requested time`; nâng cấu hình cũ sang nguồn này |
| Tài khoản chung gán tên ai thì tính cho người đó | Tính toàn bộ case cho tên được gán, không chia đôi hoặc chia theo thứ tự dòng; không chuyển giờ theo case |
| CC-VN là tài khoản mượn | Chưa biết người dùng thì cảnh báo và để trống email; chỉ gán sau khi xác nhận, không mặc định luôn là Thủy |
| Quên sửa ngày khi copy | Chỉ thêm vị trí dòng giờ ngày khác để kiểm tra; không tự kéo giờ sang ngày báo cáo |
| Giờ làm tròn khác nhau do làm tay | Giữ số giờ nguồn, không tự làm tròn; ngoại lệ Cúc ngày 11/10 vẫn tắt |
| Nguyên nhập sai giờ ngày 12/10 | Cảnh báo tại G2001 nếu nguồn còn là 5; hiển thị số 5,24 trong bản đối chiếu để kiểm tra, không tự thay |
| Ghi chú “ok” có thể có hoặc không | Giữ nội dung nguồn; không xóa để khớp bản tay |
| Chủ đề “Đánh giá khách hàng” về cơ bản không chọn | Cảnh báo và tô vàng ô chủ đề trong Session/OP; không tự đổi sang chủ đề khác |
| Tổng hợp Viber bị lệch do công thức bản tay | Đếm từ chi tiết Viber T/L; ngày 12/10 có 1 dòng thì tổng hợp là 1 |

Khi chủ tài khoản có giờ nhưng case được gán sang người khác, tool giải thích
**case và giờ đang tính cho hai người**, không kết luận luật gán sai chỉ từ việc có giờ.
Bảng gán Quỳnh → Phương Anh, Quang → Thảo Nguyên đang giữ từ DataQuyTac; cần cập nhật
nếu phân công ngày đang làm khác. SBI xác nhận cách tính theo tên đã gán, không xác nhận
các cặp này bất biến. Tên trùng Khánh Ly vẫn cảnh báo để xác nhận email nhận case.

### Cảnh báo lỗi nhập liệu

Ngày 12/10, các cảnh báo mới chỉ tới:

- `Session Report-0!O1070`: chọn Đánh giá khách hàng; giữ nguồn đến khi được sửa.
- `3. Other Tasks & Working time!G2001`: giờ hỗ trợ Nguyên còn là 5. Cảnh báo khóa theo
  ngày, email, cột và giá trị; không coi mọi ô 5 giờ là sai.
- `C2017:C2024`: giờ Đào đang ghi 11/10. SBI xác nhận khối này quên đổi ngày.
  Tool cũng chỉ khối lịch sử gần nhất `C1881:C1888`; không được sửa khối lịch sử đó
  chỉ vì cùng email. Cảnh báo ngày là gợi ý để xác minh, không tự sửa ngày.
- 11 case `CC-VN`: cần xác nhận người mượn tài khoản. Công thức tra email đã được
  bảo vệ bằng `IFERROR`, nên mở Excel tính lại vẫn trống thay vì sinh `#N/A`.

Một số nhập sai nhưng vẫn là số hợp lệ không thể tự nhận ra nếu chưa có căn cứ.
Tool không tự bù 5 thành 5,24, không chuyển giờ giữa người hoặc ngày, không tự sửa topic.
File tô vàng vẫn giữ giá trị, công thức và định dạng nguồn; cảnh báo nêu việc cần làm.

### Cập nhật và kiểm thử phiên bản 6

Tải lại `index.html` bằng **Ctrl+F5**, rồi bấm **Xử lý**. Cấu hình phiên bản 5 trở xuống
được nâng lên `rulesVersion: 6`, đổi Start Time đã lưu sang Requested time và tắt
ngoại lệ cũ ngày 11/10. Các bảng gán tùy chỉnh không bị thay bằng tên khác.
Sau khi đã nâng v6, lựa chọn Start Time mới do người dùng chủ động đặt vẫn được giữ,
nhưng có cảnh báo nếu khác Requested time. Ngày báo cáo tự động vẫn nhận theo
First response time; giờ KPI vẫn lấy từ OP, không tính từ thời lượng chat.

Kiểm thử trên dữ liệu thật ngày 12/10:

- **1.604/1.604 Start Time** lấy Requested time và khớp bản tay ngày này.
- CSKH còn 30 ô khác bản tay: 29 ghi chú “ok” được giữ và 1 chủ đề nguồn đang cảnh báo.
- Viber bằng 1; giờ Nguyên vẫn 5; giờ Đào ghi ngày 11 không bị tự chuyển sang ngày 12.
- Kiểm tra trình duyệt với cả ngày 12/10 và DataQuyTac: cảnh báo có địa chỉ ô,
  nút tải tô vàng và dữ liệu trong bản tô vàng được giữ nguyên.
- Excel thật refresh Pivot, tính lại: **33.118 ô khớp kết quả dự kiến từ nguồn và
  quy tắc hiện tại**, 344 ô phụ dư trống. Đây không phải số ô khớp toàn bộ output tay.
- Lịch sử ngày trước và cả 7 trường lọc Pivot vẫn được giữ đúng.

Các chênh lệch tên/email, giờ chưa sửa và chủ đề chưa sửa **vẫn được báo lệch** khi
đối chiếu; không đổi thành ĐẠT chỉ vì đã giải thích được nguyên nhân.

```powershell
node tool/tests/sbi-rules.cjs
node tool/tests/dataquytac.cjs
node tool/tests/app-config.cjs
node tool/tests/app-warnings.cjs
node tool/tests/browser-warnings.cjs
node tool/tests/browser-warnings.cjs --dataquytac
powershell -NoProfile -ExecutionPolicy Bypass -File tool/tests/verify-october11.ps1 -ReportFile SBI-v6-20261012.xlsx -ExpectedFile sbi-v6-excel-expected.json -Label SBI-v6-October12
```

File kiểm thử: `tool/tests/artifacts/SBI-v6-20261012.xlsx`, còn cảnh báo dữ liệu cần
xử lý trước khi gửi. Output cũ DataQuyTac dùng First response time nên nay có 1.555
ô Start Time khác theo quy chuẩn mới; 22 chỉ tiêu tổng hợp và 128 dòng giờ có nguồn
vẫn khớp. Kiểm thử kiểm tra riêng thay đổi có chủ đích này, không bỏ qua ô lệch.

## Lịch sử đối chiếu DataQuyTac phiên bản 5

Mục này lưu kết quả trước xác nhận SBI. Các quy tắc và số ô khớp dưới đây là của v5;
phần mâu thuẫn, đặc biệt Start Time và cách chia case làm tay, đã được thay bằng v6 ở trên.

Theo yêu cầu mới, **DataQuyTac được ưu tiên khi có xung đột với các bộ tháng 09,
11/10 hoặc 12/10 đã kiểm tra trước đây**. Bộ này chứa dữ liệu ngày **13/09/2026**,
gồm ba file thô, hai file đã xử lý và một báo cáo tổng chuẩn. Tool không đọc
output chuẩn để chép đáp án vào kết quả; output chỉ dùng trong kiểm thử.

| Phần xử lý | Quy tắc đang áp dụng | Căn cứ trong bộ mới |
|---|---|---|
| Session | Loại dòng lỗi bot; Start Time mặc định lấy First response time; giữ nguyên ID dạng chữ | Bỏ 12/1.567 dòng; 1.555 dòng còn lại và sheet staging khớp file đã xử lý |
| Salesforce | Bỏ cột rỗng; giữ trạng thái, ngày giờ và nội dung từ export | 379 case trong file thô và file đã xử lý không thay đổi nội dung |
| Người nhận case | Quỳnh → Phương Anh; Quang → Thảo Nguyên; Hoài giữ nguyên | Mỗi cặp đổi tên khớp 19/19 case; 6/6 case Hoài giữ nguyên |
| Khánh Ly | Giữ tên và bảng email; tra dòng đầu theo VLOOKUP, kèm cảnh báo chưa xác nhận phân công | Cùng một tên nhưng chuẩn chia 40 case cho ly.ntk và 22 case cho ly.ntk1; nguồn không có khóa để chia |
| Giờ làm | Lấy đúng dữ liệu OP, tính F bằng SUM(G:M) của khối 8 dòng; không làm tròn hay bù giờ | 128 dòng có trong nguồn khớp chuẩn sau chuẩn hóa cách hiển thị |
| Cột M — giờ việc khác | Số 0 hiển thị thành ô trống, không đổi tổng; các cột giờ khác không bị xóa số 0 | M2026, M2028, M2029 ở OP là 0; tương ứng M96, M98, M99 ở chuẩn để trống |
| Ngoại lệ riêng ngày 11/10 | Tắt mặc định năm phân công Quỳnh → Đào và một điều chỉnh giờ Cúc | Không mang ngoại lệ của bộ cũ sang quy chuẩn mới |

Đây là quy chuẩn làm việc theo bộ bạn chọn, **không phải bằng chứng rằng phân công
của hai tài khoản sẽ luôn cố định trong tương lai**. Mỗi lần áp dụng đổi người nhận,
tool vẫn ghi rõ số case và tên trước/sau trong phần chi tiết. Nếu phân công thay đổi,
cần cập nhật luật hoặc bổ sung bảng phân công theo mã case.

### Kết quả đối chiếu và dữ liệu còn thiếu

| Phạm vi | Kết quả với cấu hình mặc định mới |
|---|---|
| Hai file đã xử lý | 51.528 ô được so, khớp giá trị, kể cả metadata và sheet staging |
| 1. CSKH | 31.052/31.052 ô khớp |
| 2. SF Case info | Còn 22 ô email Khánh Ly khác chuẩn; 38 case Quỳnh/Quang đã khớp tên/email |
| 3. Other Tasks | 128 dòng có nguồn khớp; thiếu 24 dòng của 3 người, tương ứng 81 ô khác chuẩn |
| Summarize_team VietNam | 22/22 chỉ tiêu tool tự tính khớp |
| KPI tháng | 34/39 ô khớp; 5 ô còn khác liên quan phân công Khánh Ly, giờ thiếu và tổng nhóm |
| Lịch sử | 792 ô quản lý ở 12 ngày trước không đổi; kiểm tra thêm 3.120 giá trị/công thức lịch sử đều giữ nguyên |
| Pivot | Cả 7 trường lọc nhận ngày mới |

**Theo lựa chọn của bạn: giữ dữ liệu nguồn và cảnh báo, không tự bù theo output.**
Để giải quyết phần còn lệch, cần bổ sung:

- Bảng **mã case → email người xử lý** cho Khánh Ly. Không đổi tất cả 62 case sang
  `ly.ntk1@altius-link.com.vn`, không tách theo vị trí dòng và không tự thêm tên “(3M)”.
- Bảng giờ đúng ngày cho `ly.ntk1@altius-link.com.vn`, `nguyen.nt@altius-link.com.vn`
  và `anh.np@altius-link.com.vn`. File OP không có ba email này; không tự điền
  7 giờ hỗ trợ + 1 giờ khác dù output có mẫu đó.

Giao diện thực tế báo **6 người có công việc nhưng chưa có giờ KPI hợp lệ**,
bao gồm ba người trên, Hoài, Nguyễn Quang Huy và Đinh Thị Quỳnh Như. Ba người sau
cũng có KPI bằng 0 trong chuẩn, nên không làm tăng số ô KPI lệch nhưng vẫn cần cảnh báo.
Tên trùng có bảng riêng liệt kê cả hai email và vị trí `mail SF!C34`, `C43`.
Các sai khác chưa giải quyết vẫn hiện **LỖI** khi đối chiếu, không được đổi thành ĐẠT.

### Cập nhật cấu hình và kiểm thử

Tải lại `index.html` bằng **Ctrl+F5**, rồi bấm **Xử lý**. Cấu hình cũ được nâng lên
`rulesVersion: 5`, `rulesSource: "DataQuyTac"`: bật hai alias đã xác nhận, tắt các
ngoại lệ ngày 11/10 và bỏ luật tự đổi tên Khánh Ly. Danh sách ngoại lệ cũ đã lưu
được giữ để tra cứu nhưng không còn chạy. Tên đích tùy chỉnh khác, nguồn Start Time
đã chủ động chọn và lựa chọn ngày đảo từ bản v3 trở lên vẫn được giữ. Muốn dùng đúng
toàn bộ mặc định mới, chọn **Khôi phục mặc định**. Các luật tự tắt trong v5 không tự bật lại.

```powershell
node tool/tests/dataquytac.cjs
node tool/tests/dataquytac-audit.cjs
node tool/tests/app-config.cjs
node tool/tests/app-warnings.cjs
node tool/tests/browser-warnings.cjs --dataquytac
powershell -NoProfile -ExecutionPolicy Bypass -File tool/tests/verify-october11.ps1 -ReportFile DataQuyTac-20260913.xlsx -ExpectedFile dataquytac-excel-expected.json -Label DataQuyTac-v5
```

`DataQuyTac` không có file nền ngày trước, nên kiểm thử dùng file ngày 12/09 ở
`Data/test/`, không dùng output ngày 13 làm nền. File thử xuất:
`tool/tests/artifacts/DataQuyTac-20260913.xlsx`; còn thiếu dữ liệu như bảng trên,
**không phải bản hoàn thiện để gửi báo cáo**. Excel thật đã refresh Pivot và tính lại:
35.174 ô khớp giá trị dự kiến từ nguồn; 38 KPI cá nhân và tổng nhóm nhất quán với
tool; 344 ô phụ dư giữ trống. Đây không phải tuyên bố khớp 100% output chuẩn.

Các bài `october11.cjs`, `october12.cjs`, `chain14.cjs`, `chain15.cjs` kiểm tra hồi quy
với **cấu hình nghiệp vụ v4 được khai báo riêng trong tests/fixtures**, để tiếp tục
kiểm tra cơ chế xuất Excel, tính giờ, lịch sử và xử lý ngoại lệ cũ. Chúng không chứng
minh mặc định v5 khớp output cũ. Kiểm thử mặc định v5 là `dataquytac.cjs`, kiểm thử
giao diện, cùng `october.cjs`. Các kết luận lịch sử bên dưới chỉ còn áp dụng nếu
không xung đột với mục DataQuyTac này.

## Cảnh báo dữ liệu và tải file tô vàng

- Cột **Nội dung ô trong file gốc** hiển thị theo định dạng của ô nguồn: ngày
  `13/09/2026` không còn hiện thành `46278`; số `0.00` giữ cách hiển thị `0.00`.
  Đây chỉ là nội dung để tìm lại ô, không phải số tool yêu cầu bạn nhập vào.
  Cảnh báo ngày đảo nói rõ ngày đang được đọc và ngày báo cáo, không dùng mã số ngày.
- File tô vàng **giữ định dạng ngày/số, font, viền và căn lề của ô gốc**, chỉ đổi màu
  nền. Đã sửa lỗi trước đây khiến ô ngày bị chuyển sang định dạng General và hiện
  thành số `46xxx`. Để nhận bản sửa, tải lại trang, bấm Xử lý và tải lại file tô vàng.
- Sheet `Loi can sua` có cột **Trạng thái**: “Cần kiểm tra” hoặc “Đã xử lý — xem ghi
  chú”. Màu vàng không có nghĩa ô nào cũng sai hoặc cần bạn sửa; bảng chỉ rõ tool đã
  làm gì và bạn cần kiểm tra gì. Đổi định dạng ngày không tự sửa ngày đã nhập nhầm.
- Giữ các quy tắc xử lý đã kiểm chứng và những điều chỉnh riêng đã xác nhận trong
  cấu hình. Không thêm ngoại lệ chỉ để ép khớp output, không tự đoán người nhận
  case, đổi ngày hoặc bù giờ làm khi nguồn chưa đủ thông tin.
- Khối cảnh báo nằm ngay dưới số liệu kết quả, có viền đỏ/vàng và số mục cần kiểm tra.
  Mỗi mục hiện ngắn gọn vấn đề, ảnh hưởng và việc cần làm. Ví dụ: **“4 người chưa có
  giờ làm hợp lệ”** — “KPI đang bằng 0 dù có công việc. Kiểm tra ngày, email và giờ
  làm trong báo cáo nhân viên.”
- Bấm **Xem chi tiết** để mở danh sách người, email và giải thích đầy đủ. Nếu có
  địa chỉ ô nguồn, mục **Xem chi tiết và vị trí ô** giữ tên file, sheet và địa chỉ
  từng ô. Chỉ thu gọn phần giải thích, không giấu cảnh báo hay thay đổi mức độ lỗi.
  Ghi chú xử lý vẫn nằm riêng; quy tắc tính toán không thay đổi.
- Trong chi tiết, **Ảnh hưởng** và **Cần làm** có khung màu, nhãn đậm riêng.
  Danh sách người/email, tên trùng, số case, kênh chưa có nhóm và điều chỉnh đã áp
  dụng được trình bày bằng bảng. Vị trí nguồn có bảng **Ô trong file gốc — Nội dung ô
  trong file gốc — Tool xử lý / Bạn cần làm**; không cần rê chuột để đọc nội dung.
  Bảng có tiêu đề cột cố định khi cuộn, phân biệt số `0` với ô trống và cuộn ngang
  riêng trên màn hình nhỏ. Hiện đủ vị trí đã lưu thay vì chỉ 12 ô đầu; nếu vượt
  giới hạn lưu của cảnh báo thì ghi rõ số vị trí đang hiển thị/tổng số. Phần giải
  thích nguyên gốc vẫn giữ trong **Giải thích đầy đủ**.
- Có việc nhưng không có giờ hợp lệ để tính KPI được cảnh báo ngay khi bấm
  **Xử lý**, không phải đợi xuất file. Ví dụ ngày 12 có `dao.pt@altius-link.com.vn`:
  KPI bằng 0 do thiếu giờ phù hợp, không có nghĩa là không làm việc. Tool không
  tự lấy khối giờ ghi ngày 11 để bù. Khi chưa xác định được ô nguồn cần sửa,
  cảnh báo nói rõ là không có vị trí tô vàng riêng.
- Nút vàng **Tải file tô vàng** nằm **ngay sau khối cảnh báo**, trước ghi chú,
  fanpage và phần xuất báo cáo. Mỗi file có nút riêng; bản sao giữ giá trị nguồn,
  tô các vị trí được liệt kê và thêm sheet `Loi can sua` giải thích lý do.
  Bao gồm cả ô đã chuẩn hóa tự động; không phải ô vàng nào cũng cần sửa.
  File này **không phải báo cáo tổng đã xử lý**. Xác nhận dữ liệu, sửa nếu cần,
  rồi nạp lại và bấm Xử lý. Nếu danh sách vị trí bị giới hạn, giao diện có thông báo.
- Đổi file hoặc cấu hình sẽ khóa các nút tải kết quả cũ, kể cả nạp lại file trùng
  tên. Cảnh báo fanpage cập nhật ngay khi sửa ô dán. Không có cảnh báo không có
  nghĩa là đã xác nhận nguồn đầy đủ hoặc khớp 100% bản chuẩn.

Kiểm thử giao diện: `node tool/tests/app-warnings.cjs`;
kiểm thử định dạng cảnh báo và file tô vàng: `node tool/tests/warning-values.cjs`;
kiểm thử bằng Chrome/Edge ẩn với dữ liệu ngày 12: `node tool/tests/browser-warnings.cjs`.
Thử nghiệm trình duyệt dùng hồ sơ tạm riêng, kiểm tra cả giao diện sáng và màn hình
390px nền tối; xác nhận 2.967 ô của bản Salesforce tô vàng giữ nguyên giá trị,
kiểu dữ liệu và công thức. Các kiểm thử tính toán 04/10, 11/10, 12/10 và chuỗi
13→14→15/09 vẫn đạt theo phạm vi và giới hạn được nêu bên dưới.

## Lịch sử v4: sửa lỗi sau đối chiếu ngày 12/10/2026

Các kết quả và luật đổi tên trong mục này mô tả bản v4, không phải mặc định v5.

Đã sửa các lỗi xử lý có căn cứ; **không tự sửa dữ liệu nghiệp vụ để làm khớp bản chuẩn**.

- **Giờ lưu dạng chữ:** chuyển chuỗi số đầy đủ trong cột G:K và M thành số trước
  khi cộng giờ và tính KPI. Ví dụ OP `G1977 = "6.12"` của Cúc được xuất thành số
  `6.12`; tổng giờ F94 là `8.002352941176472`, khớp F86 của bản chuẩn (khác thứ tự
  nhân viên). Nhãn công việc cột L, email và ID không bị ép kiểu. Chuỗi như `6 giờ`,
  `6,12`, `6:12` không bị đọc nhầm thành 6: giữ nguyên, cảnh báo và không cộng như số.
- **Tổng KPI nhóm:** đọc phạm vi từ các công thức SUM số lượng/giờ trong file nền.
  Mẫu tháng 09 dùng dòng 15–43; các mẫu tháng 10 đã kiểm tra dùng 15–42. Không còn
  lấy cố định `kpi.daily.dataEndRow` làm nhóm. Người ngoài tổng nhóm vẫn được ghi KPI
  cá nhân. Nếu các tổng dùng phạm vi khác nhau, công thức không được hỗ trợ, vượt
  `peopleEndRow` hoặc email trong nhóm bị trùng, dừng xuất và yêu cầu kiểm tra mẫu
  thay vì tự chọn một phạm vi. Hỗ trợ SUM các ô/vùng cùng cột, kể cả tham chiếu `$`.
- **Tên người lập case:** tắt luật mặc định Quang → Nguyên. Cả 4 case Quang ngày 12
  giữ đúng tên/email như chuẩn. Cấu hình cũ tự nâng lên `rulesVersion: 4`, chỉ tắt
  đúng luật mặc định này; tên đích tùy chỉnh khác và các lựa chọn Start Time/ngày
  đảo đã lưu được giữ lại. Cấu hình v4 do người dùng chủ động chỉnh vẫn được giữ.
- **Nhất quán với Excel:** công thức tra số lượng ở Daily KPI trả **0** thay vì
  chuỗi rỗng khi không tìm thấy. Điều này tránh mất phần hiệu suất chat khi người
  đó chưa có case Salesforce được gán đúng tên. Ví dụ Thủy ngày 12 có 58 lượt hỗ trợ
  trong phép tính, nhưng công thức cũ làm H22 rỗng vì F22 rỗng. Không tự gán 11 case
  `CC-VN` cho Thủy; cảnh báo thiếu email vẫn còn. Giữ cấu trúc shared formula của mẫu.

### Kiểm chứng và giới hạn

- Ngày 12: giờ của Cúc, 4 tên/email Quang và phạm vi tổng nhóm đã được kiểm thử.
  1.430 phép kiểm tra ô/công thức lịch sử không thay đổi. Mở file bằng Excel thật,
  refresh Pivot và tính lại: **39 KPI cá nhân + 1 KPI nhóm khớp với số tool tính**.
  Đây là kiểm tra nhất quán Excel/tool, **không phải 40 ô đều khớp bản chuẩn**.
- Phép thử riêng dùng dữ liệu chi tiết của bản chuẩn tái tạo được **40/40 KPI chuẩn**.
  File `chain12-reference-fixture.xlsx` chỉ dùng chẩn đoán công thức; không dùng nó
  làm báo cáo thực tế và không đưa dữ liệu bản chuẩn vào luồng xuất thông thường.
- Ngày 11 vẫn khớp 33.757 ô chi tiết, 22 ô tổng hợp và 40 ô KPI; các ngoại lệ riêng
  đã xác nhận của ngày 11 vẫn giữ nguyên, không áp sang ngày 12.
- Ngày 12 còn 60 ô SF lệch do 19 case Quỳnh/Đào và 11 case CC-VN/Thủy; 5/40 KPI
  lệch do phân công và giờ đầu vào chưa khớp. Không tự sửa ngày 11 thành ngày 12
  cho khối giờ Đào, không đổi giờ Nguyên 5 thành 5,24, không xóa 29 mô tả “ok” hay
  đổi Topic. Bảng giờ khác thứ tự vẫn được bộ so ô báo lệch theo vị trí dòng.
- Start Time mặc định vẫn là `First response time`; bản chuẩn ngày 12 dùng
  `Requested time` cho 1.604 phiên. Có thể chọn nguồn này trong Cấu hình để tái lập
  thời gian của ngày 12, nhưng tool không tự đổi theo ngày.

Kiểm thử mới:

```powershell
node tool/tests/october12.cjs
powershell -NoProfile -ExecutionPolicy Bypass -File tool/tests/verify-october12.ps1
```

Tải lại trang bằng **Ctrl+F5** rồi bấm **Xử lý** để dựng lại kết quả. File thử xuất từ
nguồn thực tế là `tool/tests/artifacts/AltiuslinkVN-20261012.xlsx`; không sửa các file
trong `ExtraData`. Các ghi chú 11/10 và tháng 09 bên dưới là lịch sử, ưu tiên mục này
nếu có xung đột.

## Lịch sử v3: cập nhật theo bộ chuẩn 11/10/2026

Nguồn đối chiếu: `ExtraData/Daily Report 11.10.2026/Input` và `Output`.
Đây là quyết định tại thời điểm v3; phần có xung đột nay được thay bằng DataQuyTac.
Sáu ngoại lệ trong mục dưới đã tắt mặc định từ v5. Không sửa file nguồn.

| Hạng mục | Quyết định và bằng chứng |
|---|---|
| Ngày trong OP | Bật mặc định nhận ngày/tháng bị đảo. Bộ mới có 86 dòng CSKH và 16 dòng giờ cần nhận thêm để khớp chuẩn. Luôn báo số dòng; có thể tắt checkbox nếu ngày đảo thực sự là một ngày khác trong nguồn. |
| Người lập case Hoài | Giữ tên/email gốc: cả 12 case của Hoài trong chuẩn đều giữ nguyên. Tắt luật toàn bộ Hoài → Hương. |
| Người lập case Quỳnh | Tắt luật toàn bộ Quỳnh → Anh. Trong 14 case của Quỳnh, chuẩn giữ 9 case và giao 5 case cho Phạm Thanh Đào. Không suy ra quy tắc phân ca từ vị trí dòng. |
| Tổng giờ làm F | Tính lại `SUM(G:M)` của cả khối 8 dòng, đồng nhất cache với công thức Excel. Không chép tổng giờ tính sẵn đã cũ từ OP. Ví dụ F142 phải là `7.995294117647059`, không phải 8. |
| Giờ xử lý case | Chỉ điều chỉnh giá trị đã được chuẩn xác nhận; không làm tròn toàn bộ giờ. I110 của Cúc là `0.94`, các số lẻ của người khác giữ nguyên. |
| Bộ đối chiếu | Tên/email khác chuẩn vẫn báo lệch, kể cả khi do luật alias trong cấu hình; không miễn trừ tự động. |

Tại lần đối chiếu 11/10, luật Quang → Nguyên chưa đổi vì thiếu bằng chứng;
luật này đã được tắt sau đối chiếu 12/10 như mục đầu tài liệu.
Start Time mặc định vẫn là `First response time`, khớp 1.399 phiên chatbot giữ lại.

### Sáu điều chỉnh riêng, không phải quy tắc tổng quát

Đầu vào hiện có chưa đủ để suy ra việc chia case Quỳnh/Đào hoặc lý do chỉ một ô giờ
được làm tròn. Vì bạn yêu cầu ưu tiên output mới, cấu hình `approvedAdjustments`
lưu rõ nguồn bằng chứng và các điều chỉnh sau, chỉ áp dụng ngày **2026-10-11**:

- Năm Case Number chuyển từ `VU NHU QUYNH` sang `PHAM THANH DAO`:
  `20261012218289`, `20261012218452`, `20261012218454`, `20261012218613`, `20261012218655`.
- Khối giờ của `cuc.ht@altius-link.com.vn`: cột I, từ `0.941176470588235` thành `0.94`.
  `column: 6` là chỉ số bắt đầu từ 0 trong khối C:M, không phải số cột Excel.

Tool kiểm tra người lập/giá trị nguồn trước khi áp dụng, chặn khóa trùng hoặc dữ liệu
không còn khớp, báo thiếu case/khối giờ thay vì tạo dữ liệu mới. Các điều chỉnh được
ghi trong cảnh báo và log xuất file; không đọc file chuẩn trong luồng xử lý thực tế.
Tắt bằng `approvedAdjustments.enabled: false` trong Cấu hình JSON; khi đó các sai khác
nguồn tương ứng sẽ xuất hiện trở lại. Để áp dụng cho ngày khác cần bằng chứng phân công
mới hoặc dữ liệu nguồn đã sửa, không tự nhân rộng sáu ngoại lệ này.

Cấu hình trình duyệt cũ được nâng một lần lên `rulesVersion: 3`: tắt đúng hai alias
mặc định cũ, bật nhận ngày đảo và bổ sung danh sách điều chỉnh. Các tên đích tự chỉnh
khác được giữ lại. Sau khi nâng v3, lựa chọn tắt ngày đảo/điều chỉnh cũng được giữ nguyên.
Tải lại trang bằng Ctrl+F5; kiểm tra Cấu hình nếu trước đây đã tự sửa luật.

### Kết quả kiểm chứng bộ 11/10

| Phạm vi | Kết quả |
|---|---:|
| 1. CSKH | Khớp 29.358 ô — 2.097 dòng × 14 cột |
| 2. SF Case info | Khớp 2.639 ô — 377 dòng × 7 cột |
| 3. Other Tasks & Working time | Khớp 1.760 ô — 160 dòng × 11 cột |
| Summary ngày mới, phần tool quản lý | Khớp 22 ô |
| KPI tháng ngày mới | Khớp 40 ô, gồm tổng team |
| Lịch sử | 1.300 phép kiểm tra ô/công thức không đổi so với file nền |
| Excel thật sau refresh Pivot và CalculateFullRebuild | Khớp 33.859 ô, gồm thêm 40 ô KPI ngày; 344 ô phụ dư Sheet5 vẫn trống |

Kiểm thử còn kiểm tra đảo thứ tự case, tắt điều chỉnh, đổi ngày báo cáo, khóa trùng,
nguồn không khớp, nâng cấu hình và không sửa dữ liệu đầu vào trong bộ nhớ.
Đây là kết quả cho các vùng đã nêu, **không phải cam kết toàn bộ mọi ô/mọi ngày đều đúng**;
các số liệu nhập tay ngoài phạm vi, fanpage và quy tắc Setsumekai chưa có bằng chứng mới
trong bộ này không được coi là đã kiểm chứng.

```powershell
node tool/tests/october11.cjs
powershell -NoProfile -ExecutionPolicy Bypass -File tool/tests/verify-october11.ps1
node tool/tests/app-config.cjs
node tool/tests/chain14.cjs
node tool/tests/chain15.cjs
node tool/tests/october.cjs
```

File sinh để kiểm tra: `tool/tests/artifacts/AltiuslinkVN-20261011.xlsx`.
Script Excel mở file sinh ra ở chế độ chỉ đọc và không lưu thay đổi.
Các mục về tháng 09 bên dưới là lịch sử đối chiếu; kết luận nào xung đột phải theo mục này.

## Lịch sử quyết định Start Time sau đối chiếu tháng 09

**Quyết định cũ, đã được thay bằng Requested time từ v6.** Trước xác nhận SBI,
tool chọn First response time từ các mẫu sau; không thể dùng mẫu này phủ định quy chuẩn mới.

Đã quét đủ 30 file (52.985 phiên chatbot) trong `Tổng hợp dữ liệu tháng/Tháng 09` và file `full tháng`.
Ghép phiên bằng supporter + End time + sender ID, không dựa vào thứ tự dòng:

| Ngày có Session gốc | Số phiên khớp | Nguồn Start Time trong bản chuẩn |
|---|---:|---|
| 13/09 | 1.555/1.555 | First response time |
| 14/09 | 1.929/1.929 | Requested time |
| 15/09 | 1.997/1.997 | First response time |

File tổng hợp giữ nguyên Start Time của 30.796 phiên chatbot từ ngày 1–16,
chưa có phiên chatbot ngày 17–30. Dù tên file là `full tháng`, không dùng nó để
kết luận phần chưa có dữ liệu. 27 ngày ngoài 13–15 thiếu Session gốc để xác định
tên cột nguồn; thời lượng phiên dài/ngắn không chứng minh được cột nào đã dùng.

Trong **Cấu hình ánh xạ → Start Time của chatbot**, có thể chọn First response,
Requested hoặc Assign time. Lựa chọn được lưu, áp dụng sau khi bấm **Xử lý**;
đổi cấu hình sẽ khóa đầu ra cũ đến khi xử lý lại. Không tự đổi theo ngày hoặc
theo bản chuẩn. Theo quy chuẩn hiện tại, chọn Requested time cho mọi ngày.
Chỉ chọn nguồn cũ khi chủ động tái lập một báo cáo lịch sử. Ngày báo cáo tự động không thay đổi.

Trước đây cấu hình bản sửa ngày 14 từng được nâng về First response time.
Hiện v6 đổi nguồn Start Time trong cấu hình cũ sang Requested time theo SBI,
không tiếp tục giữ mốc cũ chỉ vì đã lưu trong trình duyệt. Không sửa file dữ liệu gốc.

Kiểm thử: `node tool/tests/audit-month-times.cjs` ghi thống kê không chứa thông tin
khách hàng vào `tool/tests/artifacts/month-times-audit.json`. Chạy tiếp
`node tool/tests/chain14.cjs`, `node tool/tests/chain15.cjs`, `node tool/tests/october.cjs`.
Kiểm tra bộ chọn cấu hình: `node tool/tests/app-config.cjs`. Kiểm tra Excel thật:
`powershell -NoProfile -ExecutionPolicy Bypass -File tool/tests/verify-chain15.ps1`
(và `verify-chain14.ps1`, `verify-excel.ps1` cho hai bộ dữ liệu còn lại).
Ngày 15 khớp đủ 1.997 Start Time; 1 Topic và các chênh lệch nguồn khác vẫn được báo,
không tự chép bản chuẩn vào đầu ra để làm xanh phép so sánh.

## Bản sửa chain 13/09 → 14/09/2026

- Bản sửa ban đầu dùng `Requested time` để khớp **1.929/1.929** phiên ngày 14.
  **Quyết định này đã được thay thế bởi mặc định First response time ở mục trên**;
  ngày 14 sẽ báo 1.929 ô thời gian khác bản chuẩn nếu dùng mặc định mới.
- Áp dụng bước 9 của flow: đổi đúng tên gắn với `ly.ntk1@altius-link.com.vn`
  thành `NGUYEN THI KHANH LY (3M)` trong bảng tra **và file xuất**. Sửa đủ 18 email;
  tại thời điểm bản sửa đó cả 512 email khớp bản chuẩn. Sau khi áp dụng quy tắc 11/10,
  có 31 case Hoài/Quỳnh giữ tên gốc khác chuẩn 14/09. Sau cập nhật 12/10 còn có
  20 case Quang giữ tên gốc: tổng 51 case, 102 ô tên/email, cộng 128 ô nguồn khác
  sẵn thành 230 ô SF lệch. Bộ đối chiếu báo đầy đủ, không miễn trừ. Không sửa file nền.
  Nếu tên đích đã thuộc email khác, không tự đổi; các tên trùng khác vẫn phải xử lý riêng.
- `Sheet5!V` tính Setsumekai × 45, đồng bộ với KPI JavaScript và công thức Excel.
  Hệ số nằm ở `kpi.setsumekaiWeight`; `O` giữ cách cộng `Grand Total-SUM(P:U)+V` của mẫu.
- Bộ đối chiếu không còn dùng `parseFloat` để so chuỗi thời gian/topic/ID,
  không coi hai email khác hậu tố là một, và kiểm tra cả dòng thừa/thiếu cuối bảng.
  KPI thiếu giờ làm vẫn báo lệch, không tự miễn trừ bằng phỏng đoán về bản chuẩn.
  Sai khác khớp nguồn được ghi rõ vị trí nguồn để kiểm tra.

**Vẫn còn chênh giữa bộ đầu vào hiện có và bản chuẩn, không tự chép bản chuẩn vào kết quả:**

| Vùng | Bằng chứng |
|---|---|
| SF | 127 trạng thái + 1 tên khách khác bản chuẩn; đầu ra giữ đúng export Salesforce đang nạp. |
| CSKH | Bản chuẩn có thêm dòng 2587: Setsumekai của `anh.bn1@`; file OP đang nạp không cung cấp dòng đó. |
| Other Tasks | 160 dòng đầu vào so với 192 dòng chuẩn. Thiếu 32 dòng giờ của `ly.ntk1@`, `huong.tl@`, `nguyen.nt@`, `anh.np@`. |
| Giờ làm | `G22` 5/6, `G126` 3/5, `I126` 4/2, `F38` 8/8.00352941176471; `H9` trống/dấu backtick. Các giá trị đầu ra khớp file OP hiện tại. |
| Summary | Chính CSKH của bản chuẩn có 3 Viber nhưng Summary ghi 0; chatbot có 1.618 sender khác nhau nhưng công thức mẫu dùng 1.619. Tool tính 1.710 thay vì 1.708. |
| KPI | Với đầu vào hiện có còn 6/39 ô lệch. Phép thử riêng dùng đúng các dòng/giờ trong bản chuẩn khớp **39/39**, gồm KPI Bùi Ngọc Anh 2.204761904761905 và tổng team 1.3095795162459634. |

Kiểm thử bảo toàn 1.690 ô lịch sử (giá trị và công thức) ở hai sheet tổng hợp.
Đã mở cả đầu ra thật và fixture chẩn đoán bằng Excel, refresh Pivot và tính lại:
18 email sửa đúng, 38 KPI cá nhân + tổng team khớp giữa ngày/tháng, dòng phụ trống không hiện 0.
Bộ kiểm thử 04/10 cũng chạy lại đạt.

```powershell
node tool/tests/chain14.cjs
node tool/tests/chain14.cjs --compare
powershell -NoProfile -ExecutionPolicy Bypass -File tool/tests/verify-chain14.ps1
```

`artifacts/AltiuslinkVN-20260914.xlsx` là đầu ra từ nguồn thực tế.
`artifacts/chain14-reference-fixture.xlsx` **chỉ là fixture kiểm thử bằng dữ liệu bản chuẩn,
không phải báo cáo đã hoàn thiện để gửi**. Muốn khớp bản chuẩn cần cung cấp bản OP/Salesforce
cùng thời điểm chốt, bổ sung các dòng còn thiếu và xác nhận các ô Summary tự mâu thuẫn.

## Bản sửa theo bộ đối chiếu 04/10/2026

- Khối `3. Other Tasks` chép đủ **C:M**, gồm giờ “Khác” ở M; xóa dữ liệu thừa của ngày cũ.
- Cột phụ `Sheet5!O:V` tra theo **tên kênh**, chịu được pivot thay đổi thứ tự cột.
  JP khác và Xử lý đăng ký mới vào đúng U/T; KPI tháng dùng định mức JP trong file nền.
- Cấu hình cũ lưu trong trình duyệt được bổ sung trường còn thiếu, bao gồm tổng hợp Topic;
  ánh xạ cũ C:L được nâng lên C:M. Các tuỳ chỉnh khác được giữ lại.
- Bộ nguồn 04/10 có **90 dòng đảo ngày/tháng** (82 dòng CSKH + 8 dòng giờ làm).
  Để tái tạo ảnh chuẩn, bật **Nhận dòng bị đảo ngày/tháng** trước khi xuất.

Đã mở đầu ra bằng Excel, refresh 4 pivot và tính lại: tổng giờ 4/8/8,
JP khác 24, XL ĐKM 1, KPI JP 113,08%, Tình trạng thẻ 621, tổng nội dung hỗ trợ 2.167.
KPI ngày/tháng khớp cả 38 người. Riêng số 615 trong ảnh lỗi không tái hiện với mã nguồn
hiện tại và bộ nguồn cung cấp; kiểm thử cố định kết quả 621 trước và sau khi Excel tính lại.

Chạy lại từ thư mục gốc dự án (cần Node.js, bước thứ hai cần Excel trên Windows):

```powershell
node tool/tests/october.cjs
powershell -NoProfile -ExecutionPolicy Bypass -File tool/tests/verify-excel.ps1
```

File kiểm thử nằm trong `tool/tests/artifacts/`; dữ liệu nguồn không bị sửa.
Các số nhập tay và số follow chưa được cung cấp vẫn cần hoàn thiện trước khi gửi báo cáo.

---

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

Tool ghi cho người có email trong `Daily KPI Result`, kể cả người ngoài khối OP.
JP được tính theo định mức và nguồn số lượng riêng đọc từ công thức trong file nền;
gặp mẫu công thức chưa hỗ trợ thì tool báo để kiểm tra. Dòng tổng team tính từ
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

Bộ OP tháng 09 đã khảo sát có **1.188/8.677 dòng (13,7%) nghi đảo ngày và tháng**. Một nguyên nhân có thể là locale Excel: gõ
`9/1/2026` với ý 01/9, Excel kiểu Mỹ đọc thành `m/d` = 09/01 rồi đổi luôn ra số.

Ngày > 12 không dính lỗi (không có tháng 13 nên Excel để nguyên dạng chuỗi), nên **báo cáo ngày 13–31
không ảnh hưởng**. Ngày 1–12 thì mất dòng — ví dụ ngày 07/9 mất 31%, ngày 04/9 mất 23%.

Tool phát hiện bằng cách: dòng nào có ngày không khớp, nhưng **đảo ngày↔tháng lại ra đúng ngày báo cáo**
thì đánh dấu. Từ bản cập nhật 11/10, mặc định **nhận các dòng này kèm cảnh báo**;
bỏ tích ô ở bước 2 để chỉ giữ dòng khớp ngày gốc. Việc đảo ra đúng ngày báo cáo
không tự chứng minh ngày nguồn bị sai; cần kiểm tra khi file chứa nhiều tháng.

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

## Hai điều phát hiện ở file mẫu tháng 09 — ghi nhận lịch sử

**1. Tên tài khoản Salesforce khác tên OP thật.**
Cột `Requested By` của *File báo cáo tổng* khác cột `Created By` của export ở 38/379 case,
theo đúng 2 cặp và áp dụng cho 100% số dòng của hai tên đó:

| Tên trên export Salesforce | Tên trong file báo cáo tổng | Số case |
|---|---|---|
| `VU NHU QUYNH` | `NGUYEN PHUONG ANH` | 19 |
| `LUU NGOC QUANG` | `NGUYEN THAO NGUYEN` | 19 |

Hai tên đích (`anh.np@`, `nguyen.nt@`) cũng đúng là 2 supporter chỉ có trong file tổng mà không có
trong file OP mẫu — khớp với giả thiết hai bạn này lập case bằng tài khoản SF khác.
Hai luật từng bị tắt trong v3/v4 theo bằng chứng ngày 11–12/10, nay **được bật lại
trong v5 theo yêu cầu ưu tiên DataQuyTac**. Mẫu này không chứng minh được phân công
cố định cho mọi ngày; tool vẫn ghi cảnh báo mỗi lần áp dụng để người dùng kiểm tra.

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
