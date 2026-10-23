# Luồng tính toán báo cáo ngày Excel Automation

Tài liệu này giải thích **mỗi số trong báo cáo lấy từ đâu, được tính thế nào và cần kiểm tra gì khi thấy lệch**. Các cột G–K trong bảng giờ và các cột G–K trong bảng KPI là những cột khác nhau; tài liệu luôn ghi kèm tên sheet để tránh nhầm.

Nội dung mô tả logic **SBI xác nhận ngày 23/10/2026, phiên bản 6**, căn cứ [VẤN ĐỀ.docx](<../docs/VẤN ĐỀ.docx>). Các kết luận mới được ưu tiên hơn suy luận từ output cũ; phần không thay đổi tiếp tục theo DataQuyTac. Nếu đã tự sửa cấu hình trong trình duyệt, kết quả có thể khác tài liệu này.

Đây là mô tả cách tool hiện tại hoạt động, không phải khẳng định mọi dữ liệu nguồn đều đúng. Với thông tin còn thiếu, tool giữ dữ liệu có căn cứ và cảnh báo; không lấy số từ output chuẩn để bù cho khớp.

### Những điểm SBI đã chốt

| Vấn đề | Quy tắc đang áp dụng |
|---|---|
| Mốc bắt đầu tư vấn | Start Time lấy `Requested time`, lúc khách gửi yêu cầu; không suy ra giờ KPI từ mốc này |
| Chia case của tài khoản dùng chung | Gán tên ai thì tính toàn bộ case cho người đó; không tự chia đôi theo cách làm tay |
| Tài khoản mượn như CC-VN | Phải xác nhận người dùng cho ngày báo cáo rồi gán tên; không mặc định luôn là Thủy |
| Nhân viên quên đổi ngày | Cảnh báo để sửa nguồn; không lấy giờ ngày hôm trước bù cho hôm nay |
| Làm tròn giờ | Giữ độ chính xác của số giờ trong nguồn; không tự làm tròn |
| Giờ nhân viên nhập sai | Giữ nguồn và cảnh báo trường hợp có căn cứ; không tự đổi 5 thành 5,24 |
| Ghi chú tình trạng hỗ trợ, như “ok” | Có thể có hoặc không; tool giữ nội dung nguồn, không tự xóa |
| Chủ đề “Đánh giá khách hàng” | Cảnh báo tại ô nguồn khi xuất hiện; không tự chuyển thành “Gặp nhân viên SBI” |
| Viber chi tiết có 1 nhưng tổng hợp ghi 0 | Tính từ các dòng Viber trong chi tiết; không làm theo số 0 do công thức bản tay chưa cập nhật |

## 1. Nhìn toàn bộ luồng báo cáo

```text
Session Report ── loại dòng lỗi bot ─────────────┐
                                              ├─→ 1. CSKH
SBI DAILY REPORT ── lấy CSKH đúng ngày ─────────┘
Daily Report CCVN ── chuẩn hóa, tra người/email ─→ 2. SF Case info
SBI DAILY REPORT ── lấy bảng giờ đúng ngày ──────→ 3. Other Tasks

Từ các sheet chi tiết, số liệu đi theo hai nhánh:

1. CSKH ─────────→ Summarize_team VietNam ──────→ Summarize

Cả 3 sheet ──────→ Sheet5 / Pivot ──────────────→ Daily KPI Result
                                                       │
                                                Cột ngày KPI tháng

File tổng ngày trước: cung cấp mẫu, công thức, mail SF, danh sách KPI và lịch sử.
Số Follow nhập thêm: đi vào các dòng fanpage của Summarize_team VietNam.
```

Tool tính KPI từ ba bảng chi tiết ngay khi xuất file. Các Pivot và công thức Excel cũng được giữ để Excel tính lại khi mở file. Vì vậy không phải đợi nhìn thấy Pivot mới có số KPI tháng.

## 2. Những file nào cung cấp thông tin gì

| Nguồn | Thông tin lấy từ nguồn này | Không được suy ra chỉ từ nguồn này |
|---|---|---|
| Session Report | Phiên chat, khách, email hỗ trợ, thời điểm, nội dung, chủ đề, Sender id | Giờ làm của nhân viên; người thật sự dùng tài khoản chung |
| Daily Report CCVN | Mã case, khách, trạng thái, ngày tạo, người tạo case | Giờ xử lý case; cách chia một tên trùng cho nhiều người |
| SBI DAILY REPORT của nhân viên, gọi tắt là file OP | Công việc trong `1. CSKH`; giờ khai trong `3. Other Tasks & Working time` | Giờ của người không có trong file; giờ của ngày khác |
| File báo cáo tổng làm nền | Cấu trúc báo cáo, `mail SF`, danh sách nhân viên, công thức KPI, lịch sử các ngày trước | Bằng chứng về công việc và giờ mới của ngày đang làm |
| Số Follow nhập vào tool | Lượng người theo dõi của SBIR, SMILES, DCOM | Số Like mới của ngày hôm nay |
| Output chuẩn, nếu nạp để đối chiếu | Cho biết ô nào khác kết quả tool | Không được dùng làm nguồn chép số vào báo cáo |

Một ô nguồn có công thức được đọc theo **kết quả đã lưu trong file Excel**, không phải tool tự chạy lại mọi công thức nguồn. Nếu vừa sửa dữ liệu liên quan, cần để Excel tính lại, lưu file rồi nạp lại tool.

## 3. Xác định ngày báo cáo và chọn dữ liệu

### Ngày báo cáo

Mặc định, tool lấy ngày xuất hiện nhiều nhất trong cột `First response time` của các dòng Session còn lại sau khi lọc lỗi bot. Có thể chọn ngày báo cáo trong cấu hình.

Ví dụ DataQuyTac: 1.554 dòng cùng ngày 13/09/2026, nên ngày báo cáo là **13/09/2026**.

### Không phải mọi nguồn đều được lọc ngày giống nhau

| Nguồn | Tool lấy những dòng nào |
|---|---|
| Session | Tất cả dòng còn lại sau lọc lỗi bot; không loại tiếp từng dòng theo ngày báo cáo |
| Salesforce | Tất cả dòng case trong export; không lọc theo ngày tạo case; không lấy dòng `Total` |
| CSKH trong file OP | Dòng có ngày khớp ngày báo cáo, cộng các dòng nghi đảo ngày nếu tùy chọn đó được bật |
| Giờ trong file OP | Các dòng giờ có ngày phù hợp, tổ chức theo khối 8 dòng của mẫu |

Vì vậy cần xuất đúng phạm vi Session và Salesforce. Ví dụ, Session của DataQuyTac có một dòng khác ngày nhưng vẫn được giữ, theo cách xử lý của bộ chuẩn. Tool không tự loại trùng theo mã phiên hoặc mã case; nạp dữ liệu trùng có thể làm tăng số đếm.

### Trường hợp có thể bị đảo ngày và tháng

Tool so sánh ngày đang lưu trong ô với ngày báo cáo. Nếu đổi chỗ ngày và tháng thì khớp, tool đánh dấu để kiểm tra. Mặc định hiện tại cho phép lấy các dòng này; có thể tắt tùy chọn.

Ví dụ, đang làm báo cáo **11/10/2026**, nhưng ô thực sự lưu **ngày 10 tháng 11 năm 2026**: đó chỉ là dấu hiệu *có thể* nhập đảo, không chứng minh người nhập đã sai. Nếu dòng này thực sự thuộc tháng 11 thì không được đưa vào báo cáo tháng 10.

Định dạng hiển thị như `m/d/yyyy` hay `dd/mm/yyyy` không tự chứng minh ý định nhập ngày. Với ô ngày Excel dạng số, có thể dùng `DAY(ô)` và `MONTH(ô)` để kiểm tra ngày, tháng đang lưu. Với ngày lưu dạng chữ, quy tắc đọc của tool là ngày/tháng/năm.

## 4. Tạo sheet 1. CSKH

Sheet này ghép hai nguồn: **Session sau xử lý ở trên, CSKH từ file OP ở dưới**, bắt đầu từ `A11`. Tool đánh lại số thứ tự cho toàn bộ danh sách.

### Xử lý Session trước khi ghép

Loại các dòng có `Description` chỉ là `lỗi` hoặc `lỗi bot`, không phân biệt hoa thường và cho phép khoảng trắng thừa. Không phải cứ nội dung chứa từ “lỗi” là loại.

Ví dụ DataQuyTac: **1.567 dòng − 12 dòng lỗi = 1.555 dòng giữ lại**.

### Nguồn của từng cột

| Cột trong `1. CSKH` | Nội dung | Dòng từ Session | Dòng từ OP |
|---|---|---|---|
| A | Số thứ tự | Tool đánh số | Tool đánh số tiếp |
| B | Ngày báo cáo | Ngày đã chọn | Ngày đã chọn |
| C | CHANNEL | Gán `CHAT BOT` | Cột C nguồn |
| D | Customer | `Customer` | Cột D nguồn |
| E | Supporter | `Supporter` | Cột E nguồn |
| F | Start Time | `Requested time` theo SBI | Cột F nguồn |
| G | End Time | `End time` | Cột G nguồn |
| H | Description | `Description` | Cột H nguồn |
| I | Notes | `Notes` | Cột I nguồn |
| J | Topic Level 1 | `Topic level 1` | Cột J nguồn |
| K | Sender id | `Sender id` | Cột K nguồn |
| L | Mã khách hàng | `Ma kh` | Cột L nguồn |
| M | Chi tiết Outbound Reply | Để trống | Cột M nguồn |
| N | TEAM | Để trống | Để trống theo quy tắc hiện hành |

Start Time của Session mặc định lấy Requested time cho mọi ngày. Khi nâng cấu hình cũ lên v6, tool đổi nguồn Start Time đã lưu sang Requested time. Nếu chủ động chọn nguồn khác sau đó, tool giữ lựa chọn nhưng cảnh báo khác quy chuẩn SBI. Ngày báo cáo tự động vẫn dựa vào First response time; đây là lựa chọn riêng, không phải phép tính giờ KPI. Mã khách và Sender id được giữ dạng định danh, không đổi thành số để tính toán.

Với DataQuyTac: **1.555 dòng Session + 663 dòng OP = 2.218 dòng CSKH**.

## 5. Tạo sheet 2. SF Case info và xác định người nhận case

Tool bỏ cột rỗng trong export, giữ nội dung case và ghi vào vùng `B7:H…` của báo cáo. Dòng `Total` được giữ trong file Salesforce đã xử lý nhưng không được tính là một case.

| Cột trong `2. SF Case info` | Nguồn và cách xử lý |
|---|---|
| B — Case Number | Mã case từ Salesforce |
| C — Customer information | Khách từ Salesforce |
| D — Case Status | Giữ trạng thái từ export; không tự đổi tất cả sang Complete |
| E — Ngày tạo case | Ngày giờ từ export, chuẩn hóa cách viết `yyyy/mm/dd hh:mm` |
| F — Requested By | Từ `Created By`, làm sạch dấu câu cuối tên và áp dụng quy tắc đổi người dưới đây |
| G — Supporter email | Tra tên sau xử lý trong `mail SF` của file nền |
| H — Ngày báo cáo | Ngày đang làm báo cáo, không phải ngày tạo case |

### Bảng gán người hiện tại

Bảng dưới được giữ từ DataQuyTac. SBI xác nhận **cách tính theo tên đã gán**, không xác nhận các cặp tên này cố định cho mọi ngày. Người lập báo cáo cần cập nhật bảng gán khi người dùng tài khoản thay đổi.

| Người tạo trên Salesforce | Người được ghi nhận case trong báo cáo |
|---|---|
| VU NHU QUYNH | NGUYEN PHUONG ANH |
| LUU NGOC QUANG | NGUYEN THAO NGUYEN |
| PHAM THI THU HOAI | Giữ nguyên |

**Quy tắc này chỉ đổi người nhận case Salesforce.** Nó không đổi email của phiên chat, không chuyển dòng CSKH và không chuyển giờ làm của người này sang người khác.

Ví dụ: case từ tài khoản Quỳnh được ghi cho Phương Anh, nhưng nếu giờ trong OP vẫn mang email Quỳnh thì Phương Anh chưa có giờ để tính KPI. Muốn chuyển cả giờ hoặc lượt chat, cần quy tắc nghiệp vụ bổ sung; tool hiện không tự làm việc đó.

### Tên trùng và case chưa xác định được người xử lý

Khi một tên có nhiều dòng trong `mail SF`, tool tra dòng đầu tiên như `VLOOKUP` và cảnh báo cần xác nhận email. Không tìm được tên thì email để trống, case chưa được cộng cho một nhân viên cụ thể. Công thức Excel có `IFERROR` để sau khi tính lại, email vẫn trống chứ không thành `#N/A`.

Trong DataQuyTac có **62 case cùng tên Khánh Ly**. Output chuẩn chia 40 case cho `ly.ntk` và 22 case cho `ly.ntk1`. Tool không tái tạo cách chia tay này: hiện tra dòng đầu là `ly.ntk` và cảnh báo tên trùng. Cần xác nhận một email nhận toàn bộ theo quy tắc gán của SBI; nếu nghiệp vụ yêu cầu tách từng case thì phải có bảng **mã case → email người xử lý**. Không tự thêm tên “(3M)” hoặc chia theo thứ tự dòng.

Các ngoại lệ riêng ngày 11/10 đang tắt mặc định. Mẫu tháng 10 chia case Quỳnh cho Quỳnh và Đào là cách làm tay, không phải quy tắc tự chia của tool. Nếu gán tài khoản Quỳnh cho Đào thì toàn bộ case của tài khoản sẽ tính cho Đào; giờ vẫn theo email trong OP. Việc chủ tài khoản cũng có giờ không tự chứng minh cách gán case sai.

Với `CC-VN`, SBI giải thích đây có thể là tài khoản mượn do tài khoản nhân viên bị khóa. Cần xác nhận ai đang dùng trong ngày rồi thêm luật gán tên. Chưa xác nhận thì giữ tên nguồn, email trống và cảnh báo; không lấy ví dụ Thủy ngày 12 để áp dụng cho mọi ngày.

## 6. Tạo sheet 3. Other Tasks & Working time

Đây là **bảng giờ làm**, lấy từ sheet cùng tên trong file OP, không lấy từ Session hay Salesforce. Tool lọc ngày, lấy các cột C–M rồi ghi vào báo cáo từ `C6` theo khối 8 dòng.

### Các cột G–K lấy từ đâu

| Cột trong `3. Other Tasks & Working time` | Ý nghĩa | Tool lấy hoặc tính thế nào | Có dùng tính KPI không? |
|---|---|---|---|
| C | Ngày | Ngày báo cáo | Dùng chọn dữ liệu đúng ngày |
| D | TEAM | Để trống | Không |
| E | Email nhân viên | Email ở cột E của file OP | Dùng ghép giờ với công việc |
| F | Tổng giờ làm | Tính lại tổng G–M của cả khối 8 dòng | Không lấy trực tiếp làm mẫu số KPI |
| G | Giờ hỗ trợ khách hàng | Cột G của OP | Có, cho mảng hỗ trợ khách |
| H | Giờ Outbound | Cột H của OP | Có, cho Outbound |
| I | Giờ xử lý case | Cột I của OP | Có, cho xử lý case BO |
| J | Giờ trả lời Comment | Cột J của OP | Có, cho Comment |
| K | Giờ đăng ký mới | Cột K của OP | Có, cho đăng ký mới |
| L | Tên việc khác | Cột L của OP | Không phải số giờ |
| M | Giờ việc khác | Cột M của OP | Có trong tổng F, không trong mẫu số KPI hiện hành |

**Tool không lấy số lượt tư vấn chia định mức để tự điền G–K.** Nếu ô nguồn có công thức, tool lấy kết quả đã lưu của công thức đó. Chẳng hạn nguồn tính G bằng tổng giờ trừ các việc còn lại thì G đã được tính trong file OP, không phải tool tự suy ra cách phân bổ giờ.

Ví dụ một khối ở báo cáo tổng bắt đầu dòng 6, kết thúc dòng 13:

```excel
F6 = SUM(G6:M13)
```

Vùng này có cả cột L chứa tên công việc; hàm `SUM` bỏ qua chữ. Giờ ở các dòng còn lại trong khối vẫn được cộng, không chỉ lấy dòng đầu.

### Ví dụ thật trong DataQuyTac

Khối của `phong.ht@altius-link.com.vn` ở đầu bảng kết quả có:

| Nội dung | Giá trị |
|---|---:|
| G6 — Giờ hỗ trợ khách | 6 |
| I6 — Giờ xử lý case | 1 |
| M13 — Giờ việc khác | 1 |
| F6 — Tổng giờ làm | 8 |
| Giờ dùng tính KPI, tổng G–K | 7 |

Vì vậy **8 giờ làm và 7 giờ tính KPI đều có thể đúng**: một giờ việc khác không nằm trong năm mảng KPI hiện hành. Không được tự nâng giờ hỗ trợ từ 6 lên 7 để đủ 8 giờ KPI.

### Số và chữ trong ô giờ

| Nội dung thực sự lưu trong ô | Cách tool xử lý |
|---|---|
| Số `6.12` | Dùng trực tiếp |
| Chuỗi chữ `"6.12"` | Chuyển thành số 6.12 trong báo cáo và thông báo đã xử lý |
| `6 giờ`, `6,12`, `6:12` lưu dạng chữ | Giữ nội dung, cảnh báo và không cộng như số giờ |
| Ô trống | Không đóng góp giờ; không tự bù |
| Số 0 ở cột M | Xuất thành ô trống theo mẫu DataQuyTac; tổng không đổi |

`6.12` là **6,12 giờ thập phân**, không mặc nhiên là 6 giờ 12 phút. Tool không tự làm tròn các số lẻ để khớp output.

Ctrl+1 và định dạng `0.00` chỉ nói cách trình bày ô, không xác định ô đang lưu số hay chữ. Có thể dùng `=ISNUMBER(G1977)` để kiểm tra đúng ô nguồn. File tô vàng giữ nguyên dữ liệu nguồn; việc chuyển chữ thành số được áp dụng trong báo cáo xuất ra, không sửa file OP gốc.

## 7. Tính Summarize_team VietNam

Tool tìm cột của ngày báo cáo ở dòng 4 rồi điền **22 chỉ tiêu tự tính**: 9 chỉ tiêu khách/kênh/công việc và 13 chỉ tiêu nội dung hỗ trợ. Các số này lấy từ `1. CSKH`, không lấy bằng cách đọc đáp án output chuẩn.

### Khách và kênh hỗ trợ

| Dòng | Chỉ tiêu | Cách tính từ `1. CSKH` | Ví dụ 13/09/2026 |
|---|---|---|---:|
| 11 | CUSTOMER | Sender id không trùng, không rỗng của CHAT BOT + tổng dòng 13–16 | 1.422 |
| 12 | FACEBOOK / CHAT BOT | Đếm dòng CHANNEL = `CHAT BOT` | 1.555 |
| 13 | ZALO | Đếm `Zalo T` và `Zalo L` | 31 |
| 14 | VIBER | Đếm `Viber T` và `Viber L` | 0 |
| 15 | LINE | Đếm `Line T`, `Line L`, `Line` | 0 |
| 16 | Telephone | Đếm `Tel/Hotline` và `Tel` | 13 |
| 18 | OUTBOUND | Đếm `Outbound` | 0 |
| 28 | CCVN Case check & process | Đếm `Xử lý case (BO Team)` | 524 |
| 57 | Comment đã hỗ trợ | Đếm `Trả lời Comment` | 48 |

Ví dụ CUSTOMER của ngày 13/09:

```text
1.378 Sender id CHAT BOT khác nhau + 31 Zalo + 0 Viber + 0 Line + 13 điện thoại
= 1.422
```

Do một Sender id có thể xuất hiện nhiều lần, **1.555 dòng CHAT BOT không có nghĩa là 1.555 khách khác nhau**. Việc loại trùng chỉ áp dụng cho Sender id của CHAT BOT trong công thức này; tool không xác định một người dùng cả Facebook và Zalo để loại trùng giữa các kênh.

Đặc biệt, **524 công việc xử lý case ở dòng 28 không phải 379 case Salesforce** của ngày này. Hai số lấy từ hai nguồn, phản ánh hai cách ghi nhận khác nhau.

Kênh được so khớp không phân biệt hoa thường. `JP khác`, `Outbound Reply`, `Gọi Outbound`, `Setsumekai` không có dòng kênh tương ứng ở bảng tổng hợp này. Điều đó không có nghĩa các dòng đó bị xóa khỏi CSKH hoặc đều bị loại khỏi mọi phép tính KPI.

### Nội dung hỗ trợ

Tool đọc số đứng đầu `Topic Level 1` ở cột J của `1. CSKH`, chẳng hạn `3. Tình trạng thẻ/...` được tính là topic 3. Đếm theo dòng công việc, không đếm khách duy nhất và không chỉ đếm CHAT BOT.

| Dòng tổng hợp | Số topic | Nội dung |
|---|---:|---|
| 33 | 13 | Khách Rakuraku đổi sang APP |
| 34 | 1 | Đăng ký mới |
| 35 | 2 | Thêm người nhận |
| 36 | 3 | Tình trạng thẻ |
| 37 | 4 | Phát hành lại thẻ |
| 38 | 5 | Đổi thông tin thẻ |
| 39 | 6 | Tình trạng giao dịch |
| 40 | 7 | Cập nhật thẻ ngoại kiều |
| 41 | 8 | Tỷ giá / lệ phí |
| 42 | 9 | Giấy chứng nhận chuyển tiền |
| 43 | 10 | My Number |
| 44 | 11 | Gặp nhân viên SBI |
| 45 | 12 | Đánh giá khách hàng |

Ví dụ DataQuyTac: topic 3 có 515 dòng nên dòng 36 nhận **515**; topic 13 có 38 dòng nên dòng 33 nhận **38**. Topic 14 chưa có dòng đích trong mẫu nên không cộng vào 13 mục này. Topic lạ sẽ được cảnh báo; ô thiếu topic không tự được phân loại thành “Others”.

Riêng topic 12 — **Đánh giá khách hàng**, SBI yêu cầu cảnh báo vì về cơ bản nhân viên không chọn mục này. Tool chỉ ra đúng ô ở Session hoặc OP, giữ nguyên nội dung và số đếm cho đến khi nguồn được sửa. Ví dụ ngày 12/10: `Session Report-0!O1070` đang là topic 12; tool không tự đổi thành topic 11 theo bản làm tay.

### Các chỉ tiêu chưa có đủ nguồn

| Dòng | Nội dung cần nhập hoặc kiểm tra từ nguồn khác |
|---|---|
| 6–9 | Tổng đăng ký mới và chia APP, WEBFORM, kênh khác |
| 17 | Unprocessed |
| 20–27 | Các chỉ tiêu tồn/trạng thái CCVN |
| 29–30 | Order HC, Actual HC |
| 46–51 | Ý kiến khách hàng |
| 53–55 | Người hỗ trợ Outbound, số khách Outbound, nội dung Outbound |
| 56 | Others |
| 58–60 | Khung giờ hỗ trợ, số lượng OP, số claim/ý kiến |

Tool không tự tính hay ghi đè các mục này bằng số đoán. Nếu cột ngày đã có giá trị trong file nền, giá trị đó không tự trở thành số đã được xác minh cho ngày đang làm. Người lập báo cáo vẫn cần kiểm tra.

Không lấy số người có chat làm số OP vì danh sách có thể gồm TL, JP, senior. Không lấy End Time muộn nhất làm giờ chốt hỗ trợ. Cũng không dùng số dòng trạng thái trong export để thay cho toàn bộ chỉ tiêu tồn CCVN khi chưa có quy tắc tương ứng.

### Fanpage và sheet Summarize

| Dòng | Nội dung | Nguồn |
|---|---|---|
| 61, 63, 65 | Like của SBIR, SMILES, DCOM | Mặc định chép giá trị gần nhất ở các cột ngày bên trái; không phải đo mới hôm nay |
| 62, 64, 66 | Follow của SBIR, SMILES, DCOM | Số người dùng dán vào tool, có thể lấy từ extension; thiếu thì cảnh báo |

`Summarize` là bản rút gọn, lấy số bằng công thức tham chiếu `Summarize_team VietNam`. Tool bổ sung công thức còn thiếu ở cột ngày mới bằng cách dịch công thức cột trước, không ghi đè công thức đã có. Đây không phải một nguồn dữ liệu độc lập.

## 8. Từ ba sheet chi tiết đến KPI

### Ghép dữ liệu theo email

| Nhóm thông tin | Nguồn | Cách gom |
|---|---|---|
| Số công việc theo kênh | `1. CSKH`, cột E và C | Mỗi dòng được đếm cho email hỗ trợ và kênh tương ứng |
| Số case Salesforce | `2. SF Case info`, cột G | Mỗi dòng case có email được tính cho email đó; không chỉ lấy Complete |
| Giờ từng mảng | `3. Other Tasks & Working time`, cột E và G–K | Cộng giờ theo email cho từng mảng |

`Sheet5` chứa các Pivot và cột phụ để thực hiện việc gom này trong Excel. Tool đồng thời thực hiện phép tính tương ứng để điền KPI tháng ngay khi xuất file.

Email là khóa ghép, không phải tên hiển thị. Chữ hoa/thường và khoảng trắng được chuẩn hóa; hai địa chỉ khác đuôi `.vn` và `.com.vn` vẫn là hai địa chỉ khác nhau. Tool không tự coi chúng là cùng người.

### Các cột số lượng trong Daily KPI Result

Các công thức sau là của **bảng KPI**, không phải cột cùng chữ trong bảng giờ.

```text
E = Tổng dòng CSKH của người đó
    − dòng Xử lý case (BO Team)
    − dòng Trả lời Comment
    − dòng JP khác
    − dòng Outbound
    − dòng Đăng ký mới
    − dòng Xử lý đăng ký mới
    + số dòng Setsumekai × 45

F = Số case Salesforce được gán cho email đó
G = MAX(0, E − F)
H = F × 1,5 + G
```

E là số hỗ trợ theo quy tắc mẫu; G là phần còn lại sau khi trừ case; H là số lượng hỗ trợ sau quy đổi trọng số 1,5 cho case. Không đọc H thành số giờ.

Lưu ý về Setsumekai: công thức hiện tại **cộng thêm 45 mỗi dòng** vào tổng đã chứa dòng đó. Nếu chỉ có một dòng Setsumekai và không có loại trừ khác, E bằng 46, không phải 45. Đây là cách code và công thức mẫu hiện tính; không nên diễn giải thành “đổi một buổi thành đúng 45 lượt” hoặc “45 phút”.

### Số lượng, giờ và định mức của từng mảng

Đây là bộ định mức OP mặc định. Định mức là số lượng quy chuẩn trên một giờ, không phải số giờ điền sẵn cho nhân viên.

| Mảng | Cột số lượng trong Daily KPI | Cột giờ trong Daily KPI | Giờ lấy từ Other Tasks | Định mức | Cột hiệu suất và phép tính |
|---|---|---|---|---:|---|
| Hỗ trợ khách | H — số lượng quy đổi | I | G | 15 | J = H / (I × 15) |
| Outbound | K — dòng Outbound | L | H | 30 | M = K / (L × 30) |
| Xử lý case BO | N — dòng BO trong CSKH | O | I | 17 | P = N / (O × 17) |
| Comment | Q — dòng Comment | R | J | 20 | S = Q / (R × 20) |
| Đăng ký mới | T — dòng Đăng ký mới | U | K | 20 | V = T / (U × 20) |

Không nhầm **F — case Salesforce** với **N — công việc BO ghi trong CSKH**. F tham gia quy đổi phần hỗ trợ khách; N tính mảng xử lý case BO riêng.

Mảng không có giờ dương không đóng góp vào tử số KPI tổng. Tool không tự bổ sung giờ chỉ vì đã có số lượng công việc ở mảng đó.

### KPI tổng của một người

```text
W = (J × I + M × L + P × O + S × R + V × U)
    / (I + L + O + R + U)
```

Đây là hiệu suất các mảng có xét số giờ, không phải trung bình cộng đơn giản của năm hiệu suất. Nếu tổng giờ bằng 0, kết quả W là 0. Số `1` tương ứng 100% định mức; `0,8` tương ứng 80%, tùy định dạng hiển thị của ô.

Nếu file nền có nhân viên dùng định mức riêng, tool đọc định mức và nguồn số lượng từ công thức của người đó khi nhận diện được mẫu công thức. Nếu không hiểu mẫu, tool cảnh báo và không tự ghi KPI tháng cho người đó; không áp bộ OP cho tất cả nhân viên.

### Ví dụ thật của Phong ngày 13/09/2026

Nguồn là bộ DataQuyTac, email `phong.ht@altius-link.com.vn`.

| Bước | Dữ liệu và phép tính | Kết quả |
|---|---|---:|
| Đếm CSKH | 93 dòng tổng, trong đó 21 dòng BO; các nhóm loại trừ khác bằng 0 | E = 72 |
| Đếm Salesforce | 26 case được gán cho email Phong | F = 26 |
| Phần hỗ trợ còn lại | 72 − 26 | G = 46 |
| Quy đổi hỗ trợ | 26 × 1,5 + 46 | H = 85 |
| Giờ hỗ trợ | Lấy từ Other Tasks G | I = 6 |
| Hiệu suất hỗ trợ | 85 / (6 × 15) | J ≈ 0,944444 |
| Công việc BO | 21 dòng BO, 1 giờ xử lý | N = 21; O = 1 |
| Hiệu suất BO | 21 / (1 × 17) | P ≈ 1,235294 |
| KPI tổng | (85 / 15 + 21 / 17) / (6 + 1) | W ≈ 0,9859944, tức 98,60% |

Một giờ việc khác của Phong vẫn được tính vào tổng 8 giờ làm ở Other Tasks, nhưng không đưa vào mẫu số 7 giờ của công thức KPI này.

### “Có công việc nhưng thiếu giờ” nghĩa là gì

Nghĩa là tool đếm được dòng CSKH hoặc case cho một email, nhưng tổng giờ của email đó ở năm mảng G–K trong Other Tasks không đủ để làm mẫu số dương. Không có nghĩa người đó không làm việc, cũng không có nghĩa phải khai đủ 8 giờ.

Nguyên nhân có thể là thiếu khối giờ, ngày không khớp, email khác nhau, giờ lưu thành chữ không đọc được, hoặc chỉ khai giờ việc khác ở M. Cần kiểm tra đúng người, đúng ngày, đúng cột giờ trước khi kết luận.

Việc dùng chung tài khoản chỉ giải quyết được nếu có quy tắc rõ **chuyển phần nào cho ai**. Quy tắc chuyển case không tự chuyển cả chat và giờ làm.

## 9. KPI nhóm và KPI tháng

### KPI nhóm

Tool đọc phạm vi các công thức `SUM` ở dòng tổng 14 của `Daily KPI Result` để xác định ai thuộc nhóm, không mặc định lấy tất cả người xuất hiện trong dữ liệu.

Tool cộng số lượng và giờ của nhóm trước, tính lại G và H của nhóm, rồi tính hiệu suất theo cùng bộ định mức OP. **Không lấy trung bình các W cá nhân**, cũng không mặc nhiên cộng H cá nhân, vì G có chặn tối thiểu bằng 0.

Nếu công thức tổng số lượng và tổng giờ dùng các nhóm dòng khác nhau, hoặc email trong nhóm bị trùng, tool dừng phần xuất thay vì chọn đại một phạm vi.

### Ghi sang KPI tháng

Tool tìm sheet dạng `202609 KPI Result`, `202610 KPI Result`…, tìm cột ngày ở dòng 5 và dòng nhân viên theo email ở cột D. KPI cá nhân được ghi vào đúng giao điểm email/ngày; KPI nhóm được ghi ở dòng 7.

Nhân viên có trong danh sách KPI ngày nhưng nằm ngoài nhóm tổng vẫn có thể được ghi KPI cá nhân. Người không có dòng trong KPI tháng cần được bổ sung vào mẫu; tool không tự chèn người vào vị trí đoán trước.

Các cột ngày trước được giữ theo file nền. Nếu output chuẩn đã sửa tay số của ngày cũ mà file nền chưa sửa, tool không tự lấy các số mới đó từ output để thay lịch sử.

## 10. Cảnh báo và cách lần ngược một số bị lệch

Không phải mọi cảnh báo đều có nghĩa dữ liệu bị sai. Có mục là thao tác tool đã xử lý, có mục cần người lập báo cáo xác nhận.

| Cảnh báo hoặc tình huống | Ý nghĩa | Việc cần làm |
|---|---|---|
| Đã chuyển giờ dạng chữ thành số | Báo cáo đã dùng được số giờ; nguồn chưa bị sửa | Kiểm tra số giờ có đúng không; không cần sửa chỉ vì ô được tô vàng |
| Không đọc được giờ | Nội dung đó chưa được cộng như số | Nhập số giờ rõ ràng, để Excel tính lại và lưu nguồn |
| Có công việc nhưng thiếu giờ | KPI có thể bằng 0 vì thiếu mẫu số | So email, ngày và G–K trong OP |
| Có giờ khai ở ngày khác | Tìm thấy các dòng giờ gần ngày báo cáo nhất của người đang thiếu giờ | Kiểm tra có quên đổi ngày khi copy không; nếu là lịch sử đúng thì giữ nguyên |
| Giờ SBI xác nhận nhập sai | Lỗi nhập liệu đã được xác nhận, chưa sửa trong nguồn đang nạp | Ví dụ Nguyên ngày 12/10: `G2001` vẫn là 5, bản đối chiếu là 5,24; xác nhận và sửa nguồn |
| Chọn “Đánh giá khách hàng” | Chủ đề cần kiểm tra theo yêu cầu SBI | Xem nội dung phiên rồi sửa chủ đề đúng, không tự chọn thay |
| Tên có nhiều email | Tra tên chưa đủ để phân công chính xác | Bổ sung quy tắc theo mã case/người thực hiện |
| Có thể đảo ngày | Có hai cách hiểu cần xác nhận | Kiểm tra ngày thực sự lưu và ngày công việc; chọn nhận hoặc bỏ dòng phù hợp |
| Đã đổi người nhận case | Tool áp dụng luật tên trong cấu hình | Xác nhận phân công còn đúng; giờ không tự chuyển theo |
| Chỉ tiêu chưa có nguồn | Chưa đủ dữ liệu để tự tính | Nhập từ nguồn nghiệp vụ có căn cứ, không dùng số 0 thay cho “chưa biết” |

File tô vàng là bản sao để dò vị trí; có sheet `Loi can sua`, giữ dữ liệu và định dạng ô nguồn. Màu vàng không tự khẳng định ô sai về nghiệp vụ. File gốc không bị sửa.

Cảnh báo Nguyên chỉ áp dụng đúng **ngày 12/10/2026 + email Nguyên + cột G còn là 5**; sửa thành số khác thì cảnh báo đối chiếu này hết. Không coi mọi ô 5 giờ là sai. Một số nhập sai nhưng vẫn có dạng số hợp lệ không thể tự phát hiện nếu chưa có xác nhận hoặc căn cứ đối chiếu.

Ví dụ giờ Đào: tool chỉ các khối `C1881:C1888` và `C2017:C2024` đang ghi 11/10 khi làm báo cáo 12/10. Tài liệu SBI xác nhận khối 2017–2024 quên đổi ngày; cần sửa khối đó trong nguồn. Không vì cùng là giờ Đào mà sửa cả khối lịch sử 1881–1888.

Khi kiểm tra một số cụ thể, đi ngược theo thứ tự:

1. Xác định đúng sheet, cột ngày và email hoặc loại chỉ tiêu đang lệch.
2. Nếu là tổng hợp kênh/topic, lọc lại `1. CSKH` theo các tiêu chí ở mục 7.
3. Nếu là KPI, kiểm tra E–H, số công việc từng mảng và giờ tương ứng trong `Daily KPI Result`.
4. Lần về ba sheet chi tiết: công việc CSKH, case Salesforce, giờ Other Tasks. Đừng chỉ nhìn tổng giờ F.
5. Đối chiếu với ô trong file nguồn; kiểm tra ngày, kiểu dữ liệu, công thức đã lưu và quy tắc đổi người.
6. Sau khi sửa nguồn hoặc quy tắc đã được xác nhận, chạy lại tool. Mở Excel và cho phép tính lại/Refresh All nếu cần kiểm tra Pivot và công thức.

## 11. Những giới hạn đang biết trong DataQuyTac

Với quy tắc v6, **1.555 ô Start Time khác output DataQuyTac cũ** vì nay phải lấy Requested time. Các ô CSKH còn lại, 22 chỉ tiêu tổng hợp và 128 dòng giờ có nguồn vẫn khớp. Không ép Start Time về First response chỉ để đạt 100% với mẫu cũ.

| Phần còn thiếu căn cứ | Ảnh hưởng |
|---|---|
| Output tay chia 22 case Khánh Ly sang email thứ hai | Tool không tự chia; cần xác nhận email được gán, khác biệt kéo theo KPI cá nhân và nhóm |
| OP không có giờ của `ly.ntk1@altius-link.com.vn`, `nguyen.nt@altius-link.com.vn`, `anh.np@altius-link.com.vn` | Không dựng được 24 dòng giờ có trong output; KPI liên quan khác chuẩn |
| Một số chỉ tiêu tổng hợp chưa có nguồn/quy tắc | Vẫn cần người lập báo cáo bổ sung |

Theo lựa chọn đã xác nhận: **giữ nguồn và cảnh báo, không tự bù 7 giờ hỗ trợ + 1 giờ khác, không tự chia case để khớp đáp án**. Khớp số với output không thay thế được căn cứ về người thực hiện và thời gian thực tế.

## 12. Nơi tra cứu trong dự án

| Tài liệu hoặc mã nguồn | Dùng để kiểm tra |
|---|---|
| [README](README.md) | Quy tắc hiện hành, cách sử dụng, kết quả đối chiếu và giới hạn |
| [Kết luận SBI](<../docs/VẤN ĐỀ.docx>) | Quyết định nghiệp vụ được ưu tiên từ ngày 23/10/2026 |
| [Cấu hình](src/config.js) | Nguồn cột, đổi người nhận, định mức, dòng tổng hợp và các tùy chọn |
| [Xử lý dữ liệu và tính toán](src/pipeline.js) | Ghép ba bảng, đếm chỉ tiêu, tính KPI, phát hiện thiếu dữ liệu |
| [Đọc số và ngày](src/parse.js) | Cách đọc ngày, giờ dạng chữ và nội dung ô |
| [Ghi vào báo cáo Excel](src/xlsxpatch.js) | Ghi cột ngày, công thức tổng, KPI tháng, Pivot và giữ lịch sử |
| [Kiểm thử DataQuyTac](tests/dataquytac.cjs) | Kiểm tra bộ quy tắc mặc định mới với dữ liệu chuẩn |
| [Kiểm thử quy tắc SBI](tests/sbi-rules.cjs) | Kiểm tra ngày 12/10, cảnh báo ô nguồn, phân công, độ chính xác giờ và giữ nguyên dữ liệu gốc |
| [Đối chiếu chi tiết DataQuyTac](tests/dataquytac-audit.cjs) | Lần từ file thô qua file đã xử lý đến output |

Các ví dụ số thật trong tài liệu dùng [file thô DataQuyTac](<../DataQuyTac/File thô/>), đối chiếu với [output ngày 13/09/2026](<../DataQuyTac/Output/>). File nền cho phép thử là báo cáo 12/09/2026 trong `Data/test`, không phải lấy chính output ngày 13/09 làm nền.
