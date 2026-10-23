# BÁO CÁO DỰ ÁN TOOL TỰ ĐỘNG HÓA BÁO CÁO EXCEL

**Dự án:** Tool báo cáo ngày CCVN / Altius Link VN  
**Ngày cập nhật:** 14/10/2026  
**Phạm vi:** Tổng hợp công việc đã thực hiện, kết quả kiểm chứng và các vấn đề cần tiếp tục xử lý.

## 1. Tổng quan

Tôi đã xây dựng tool hỗ trợ tự động hóa quy trình lập báo cáo Excel hằng ngày, từ khâu đọc và làm sạch dữ liệu đến tổng hợp, tính KPI và xuất file báo cáo theo mẫu có sẵn.

Ngoài việc phát triển chức năng, tôi đã kiểm tra kết quả trên các bộ dữ liệu thực tế, xử lý những lỗi phát sinh và đối chiếu báo cáo của cả tháng 09 để đánh giá khả năng áp dụng tổng quát. Quá trình này giúp phân biệt lỗi của tool với những chênh lệch do dữ liệu đầu vào thiếu, khác thời điểm chốt hoặc quy tắc nghiệp vụ chưa thống nhất.

**Trạng thái hiện tại:** Đã có tool chạy được, bộ kiểm thử và tài liệu sử dụng. Một số quy tắc nghiệp vụ và dữ liệu nguồn vẫn cần hoàn thiện; chưa thể kết luận tool khớp 100% mọi báo cáo mẫu hoặc thay thế toàn bộ bước kiểm tra của người làm báo cáo.

## 2. Mục tiêu thực hiện

- Giảm thao tác copy/paste giữa nhiều file Excel.
- Tự động lọc, chuẩn hóa và đưa dữ liệu vào đúng sheet, đúng cột.
- Hạn chế mất dữ liệu, sai mã khách hàng, thiếu giờ làm và sai ánh xạ người phụ trách.
- Giữ cấu trúc, định dạng, công thức và PivotTable của file báo cáo tổng.
- Bảo toàn dữ liệu lịch sử khi lập báo cáo cho ngày mới.
- Có công cụ đối chiếu để phát hiện và giải thích các ô không khớp.
- Xử lý file Excel trên máy người dùng, không cần tải workbook lên máy chủ.

## 3. Những công việc đã thực hiện

### 3.1. Phân tích quy trình và dữ liệu đầu vào

Tôi đã đọc, rà soát và cập nhật quy trình trong [flow.md](flow.md), xác định các nhóm dữ liệu cần dùng:

| Nhóm dữ liệu | Vai trò |
|---|---|
| Báo cáo tổng ngày trước | File nền để lập báo cáo ngày mới và giữ lịch sử |
| Session Report | Dữ liệu phiên chatbot từ Live Support |
| Daily Report CCVN | Dữ liệu case xuất từ Salesforce |
| Báo cáo ngày của nhân viên — file OP | Dữ liệu CSKH, hoạt động khác và giờ làm |

Tôi cũng đã đối chiếu dữ liệu gốc, báo cáo mẫu và các ảnh phản ánh sai lệch để xác định vị trí lỗi, thay vì chỉ kiểm tra số tổng cuối cùng.

### 3.2. Xây dựng giao diện và luồng sử dụng

- Xây dựng giao diện web bằng HTML, CSS và JavaScript, có thể mở từ `tool/index.html`.
- Cho phép kéo thả nhiều file, tự nhận diện loại file và thông báo nhóm dữ liệu còn thiếu.
- Hỗ trợ nạp nhiều file OP để tổng hợp dữ liệu.
- Hiển thị ngày báo cáo, số dòng đã xử lý, dữ liệu xem trước và cảnh báo.
- Cho phép xuất báo cáo tổng hoặc xuất riêng các khối dữ liệu để dán thủ công.
- Bổ sung cấu hình ánh xạ, bảng đổi tên người lập case và chức năng nhập/xuất cấu hình JSON.
- Khóa kết quả cũ khi thay file hoặc thay cấu hình cần xử lý lại, tránh tải nhầm báo cáo chưa cập nhật.

### 3.3. Tự động làm sạch và chuẩn hóa dữ liệu

- Lọc các phiên có mô tả thuộc nhóm “lỗi bot” theo cấu hình.
- Chuẩn hóa ngày tháng, xử lý dữ liệu ngày ở dạng serial Excel và dạng chuỗi.
- Phát hiện ngày/tháng bị đảo và cho người dùng lựa chọn có nhận các dòng đó hay không.
- Loại dòng chỉ có ngày nhưng không có nội dung nghiệp vụ.
- Giữ mã khách hàng, sender ID dài và chuỗi thời gian ở kiểu dữ liệu phù hợp, tránh mất chữ số khi Excel chuyển thành số.
- Chuẩn hóa tên người lập case trước khi tra email, xử lý ký tự thừa ở cuối tên.
- Đọc bảng `mail SF` từ file nền, phát hiện tên trùng và tên chưa tra được email.

### 3.4. Xây dựng các khối dữ liệu đầu ra

Tool đã tự động dựng ba khối chính:

| Sheet đích | Nội dung xử lý |
|---|---|
| `1. CSKH` | Gộp phiên chatbot và hoạt động CSKH từ file OP; đánh số và gán ngày báo cáo |
| `2. SF Case info` | Đưa dữ liệu Salesforce vào báo cáo, áp dụng cấu hình tên và tra email |
| `3. Other Tasks & Working time` | Đưa giờ làm và hoạt động khác vào báo cáo, chép đủ vùng C:M |

Khi ghi dữ liệu mới, tool dọn giá trị thừa của lần chạy trước, kể cả trường hợp số dòng giảm hoặc khối dữ liệu mới rỗng.

### 3.5. Tổng hợp báo cáo và tính KPI

- Tính các chỉ tiêu có thể suy ra từ dữ liệu đầu vào cho `Summarize_team VietNam`.
- Tổng hợp nội dung hỗ trợ theo nhóm Topic.
- Xử lý công thức liên kết sang sheet `Summarize`.
- Tính KPI theo từng nhân viên và dòng tổng team, ghi vào cột ngày của sheet KPI tháng.
- Đọc và xử lý các mẫu định mức riêng đã hỗ trợ, gồm trường hợp nhân sự JP trong bộ dữ liệu 04/10.
- Bổ sung kiểm tra người có công việc nhưng thiếu giờ KPI hoặc chưa có dòng trong danh sách nhân sự.
- Giữ riêng các chỉ tiêu chưa đủ nguồn, cần người làm báo cáo bổ sung hoặc xác nhận.

### 3.6. Ghi file Excel và bảo toàn cấu trúc báo cáo

Tôi đã sử dụng cách chỉnh sửa có chọn lọc XML bên trong file `.xlsx`, thay vì dựng lại toàn bộ workbook từ bảng dữ liệu thuần.

Các phần đã xử lý gồm:

- Giữ định dạng ô tại các vùng được cập nhật và duy trì các thành phần workbook liên quan.
- Ghi lại công thức tra email và tổng giờ theo đúng vị trí dòng.
- Cập nhật ngày báo cáo tại các ô tiêu đề.
- Cập nhật vùng dữ liệu phục vụ Pivot và thiết lập tính lại khi mở bằng Excel.
- Bật nhận mục ngày mới trong bộ lọc Pivot để tránh refresh xong không còn dữ liệu.
- Tạo công thức phụ theo tên kênh, hạn chế phụ thuộc vào thứ tự cột của Pivot.
- Giữ dữ liệu các ngày trước trong các vùng lịch sử đã kiểm thử.

### 3.7. Bổ sung hỗ trợ lấy số liệu fanpage

Tôi đã xây dựng extension hỗ trợ lấy số follow của các fanpage SBIR, SMILES và DCOM, đưa số liệu sang tool hoặc sao chép để dán thủ công.

Tool có trạng thái kết nối extension, cảnh báo khi thiếu số liệu và khi Facebook chỉ cung cấp số làm tròn. Phần xử lý workbook vẫn thực hiện trên máy người dùng; chức năng lấy số follow cần truy cập Facebook và phụ thuộc khả năng đọc dữ liệu từ trang tại thời điểm sử dụng.

## 4. Các lỗi đã xử lý

| Vấn đề | Công việc đã thực hiện |
|---|---|
| Thiếu cột giờ “Khác” | Mở rộng ánh xạ Other Tasks từ C:L sang C:M và nâng cấp cấu hình cũ |
| Kênh JP bị lệch khi Pivot đổi cột | Dựng cột phụ theo nhãn kênh thay vì sao chép vị trí cột cố định |
| Các dòng dưới Pivot xuất hiện nhiều số 0 | Giữ công thức dự phòng nhưng trả trống khi không có dòng Pivot tương ứng |
| Mất dữ liệu khi Pivot gặp ngày mới | Bật nhận mục mới trong bộ lọc Pivot |
| 18 case không tra được email trong chuỗi 13 → 14 | Áp dụng quy tắc phân biệt tài khoản `(3M)` trong bảng tra và file xuất; không sửa file nền |
| Bộ so sánh báo khớp sai chuỗi thời gian/ID | Không dùng cách đọc tiền tố số để so chuỗi; kiểm tra đúng nội dung ô |
| Bỏ sót dòng thừa hoặc thiếu cuối bảng | So sánh toàn bộ phạm vi của cả hai bên, không chỉ phần giao nhau |
| Coi hai email khác hậu tố là cùng người | So sánh email đầy đủ, không tự gộp tài khoản |
| Start Time khác nhau giữa các bản chuẩn | Chọn mặc định `First response time`, thêm bộ chọn nguồn và giữ lựa chọn đã lưu |
| Thay cấu hình nhưng vẫn tải kết quả cũ | Khóa đầu ra và yêu cầu xử lý lại |

Việc sửa Start Time đã được điều chỉnh sau khi kiểm tra thêm dữ liệu tháng: không cố định `Requested time` cho mọi ngày chỉ để khớp riêng ngày 14.

## 5. Kiểm thử và kết quả đã ghi nhận

Tôi đã xây dựng các bài kiểm thử tự động, phép đối chiếu dữ liệu mẫu và kiểm tra bằng Microsoft Excel thật để xác nhận kết quả sau khi refresh Pivot, tính lại công thức.

Các số liệu dưới đây là kết quả đã ghi nhận trong quá trình phát triển, không phải tuyên bố đã chạy lại toàn bộ kiểm thử vào ngày lập báo cáo này.

### 5.1. Chuỗi ngày 13 → 14

- Kiểm tra đủ 512 email trong SF Case; sửa được 18 email bị thiếu liên quan đến tài khoản `(3M)`.
- Kiểm chứng 1.929 Start Time theo nguồn được chọn. Riêng bản chuẩn ngày 14 dùng `Requested time`, nên cần chọn nguồn này nếu muốn tái lập cách làm của bản chuẩn.
- Kiểm tra bảo toàn 1.690 ô lịch sử, gồm giá trị và công thức, trên hai sheet tổng hợp.
- Kiểm tra 38 KPI cá nhân và dòng tổng team khớp giữa ngày/tháng sau khi Excel tính lại.
- Vẫn báo rõ chênh lệch trạng thái Salesforce, giờ làm và dòng dữ liệu thiếu so với bản chuẩn.

### 5.2. Chuỗi ngày 14 → 15

- Kiểm tra 2.741 dòng CSKH, 548 dòng SF Case và 176 dòng Other Tasks từ bộ nguồn đang có.
- Khớp 1.997 Start Time với bản chuẩn khi dùng `First response time`.
- Kiểm tra bảo toàn 1.820 ô lịch sử, gồm giá trị và công thức.
- Kiểm tra 38 KPI cá nhân và dòng tổng team khớp giữa ngày/tháng sau refresh bằng Excel.
- Phát hiện nguồn thiếu 16 dòng giờ làm so với 192 dòng trong bản chuẩn; vẫn còn các sai khác về Topic, trạng thái case, phân công và giờ làm.

### 5.3. Bộ dữ liệu ngày 04/10

- Kiểm tra việc chép đủ cột giờ “Khác”, tổng giờ và dữ liệu JP.
- Ghi nhận 90 dòng đảo ngày/tháng trong nguồn; cần bật lựa chọn nhận các dòng này để tái tạo kết quả đối chiếu.
- Kiểm tra KPI JP khoảng 113,08%, chỉ tiêu Tình trạng thẻ là 621 và tổng nội dung hỗ trợ là 2.167 trên bộ nguồn cung cấp.
- Kiểm tra KPI ngày/tháng khớp cho 38 người sau khi Excel tính lại.
- Kiểm tra 336 ô phụ tại các dòng trống dưới Pivot không còn hiện số 0.

### 5.4. Phân biệt kiểm tra công thức với kiểm tra nguồn thực tế

Tôi đã thực hiện phép thử riêng bằng dữ liệu các bảng chi tiết trong bản chuẩn ngày 14 và 15. Với bộ dữ liệu đó, phép tính tái tạo được 39/39 ô KPI đối chiếu.

Kết quả này giúp khoanh vùng nguyên nhân chênh lệch, nhưng **không có nghĩa đầu ra từ bộ nguồn thực tế đã khớp 39/39 ô**. Các file kiểm thử dựng bằng dữ liệu bản chuẩn chỉ dùng để chẩn đoán, không được coi là báo cáo đã hoàn thiện để gửi.

## 6. Đối chiếu toàn bộ báo cáo tháng 09

Tôi đã bổ sung hai bộ kiểm tra để rà soát nguồn Start Time và các quy tắc tính toán trên 30 báo cáo ngày.

Kết quả chính:

- Đã quét 52.985 phiên chatbot trong 30 file báo cáo ngày.
- Ghép và đối chiếu 30.796 phiên chatbot ngày 1–16 với file tổng hợp tháng; file này chưa có chatbot ngày 17–30.
- Xác nhận từ Session gốc: ngày 13 và 15 dùng `First response time`, ngày 14 dùng `Requested time`.
- Phát hiện 28 ô Summary không khớp bảng chi tiết trên 19 ngày.
- Phát hiện 7 ô KPI cá nhân khác giữa ngày và tháng, trong đó 2 ô khác do làm tròn.
- Trong phạm vi 1.122 kết quả KPI cá nhân có định mức được bộ kiểm tra nhận diện là chuẩn, phát hiện 8 kết quả khác phép tính hiện tại của tool; đây không phải kiểm chứng mọi định mức hoặc mọi nhân sự.
- Phát hiện các tên kênh `Webform` và `ĐKM` cần được nhận diện là Đăng ký mới.
- Phát hiện Setsumekai có thời lượng khác nhau và công thức quy đổi chưa thống nhất giữa các báo cáo.
- Phát hiện tên tài khoản gốc và tên được thay thế cùng xuất hiện trong nhiều ngày, cho thấy luật đổi tên toàn bộ chưa đủ để biểu diễn mọi trường hợp phân công.

Qua đối chiếu, tôi đã xác định rằng cần kiểm tra cả tính nhất quán bên trong báo cáo mẫu, không chỉ dùng số liệu của mẫu làm đáp án tuyệt đối.

## 7. Các vấn đề còn tồn tại và hướng xử lý

Các mục dưới đây đã được phát hiện hoặc đề xuất, **chưa được coi là đã triển khai hoàn tất**:

| Nội dung | Tình trạng và hướng xử lý |
|---|---|
| Kênh `Webform` / `ĐKM` | Cần bổ sung vào nhóm Đăng ký mới và đồng bộ giữa JavaScript với công thức Excel |
| Setsumekai | Hiện vẫn có quy đổi theo số dòng và hệ số cấu hình. Cần chuyển sang thời lượng thực tế, đồng thời thống nhất cách chống đếm trùng |
| Phân công người lập case | Hiện có bảng đổi tên theo cấu hình. Cần bổ sung điều kiện theo case, ngày hoặc ca đã xác nhận |
| Thiếu giờ làm | Cần nguồn bổ sung có ngày, email và nguồn gốc rõ ràng; không tự lấy giờ hôm trước hoặc mặc định 8 giờ |
| Trạng thái Salesforce khác bản chuẩn | Cần bộ dữ liệu đúng thời điểm chốt; không suy đoán trạng thái cũ từ export hiện tại |
| Mẫu tự mâu thuẫn | Cần thống nhất nguồn được ưu tiên và quy trình phê duyệt điều chỉnh |
| Chỉ tiêu nhập tay, số follow và dữ liệu ngoài phạm vi | Người làm báo cáo vẫn cần bổ sung, xác nhận trước khi gửi |
| Phạm vi kiểm chứng | Chưa đủ Session gốc của cả tháng để xác nhận mọi quy tắc nguồn; chưa có cơ sở kết luận đúng tuyệt đối cho mọi bộ dữ liệu |

Thứ tự ưu tiên tiếp theo là sửa ánh xạ kênh đã xác định rõ, hoàn thiện quy tắc Setsumekai và phân công case, sau đó bổ sung kiểm tra tính đầy đủ của nguồn trước khi xuất báo cáo.

## 8. Sản phẩm và tài liệu đã xây dựng

| Thành phần | Vị trí |
|---|---|
| Giao diện tool | [tool/index.html](tool/index.html) |
| Mã nguồn xử lý | [tool/src/](tool/src/) |
| Extension lấy số follow | [extension/](extension/) |
| Quy trình nghiệp vụ | [flow.md](flow.md) |
| Hướng dẫn sử dụng | [HUONG-DAN-SU-DUNG.md](HUONG-DAN-SU-DUNG.md) |
| Tài liệu kỹ thuật và ghi chú kiểm thử | [tool/README.md](tool/README.md) |
| Bộ kiểm thử | [tool/tests/](tool/tests/) |
| Kết quả đối chiếu thời gian cả tháng | [month-times-audit.json](tool/tests/artifacts/month-times-audit.json) |
| Kết quả đối chiếu quy tắc cả tháng | [month-rules-audit.json](tool/tests/artifacts/month-rules-audit.json) |
| Bản đóng gói tool | [tool/tool.zip](tool/tool.zip) |

## 9. Kết luận

Tôi đã xây dựng được luồng tự động hóa các bước chính của quy trình lập báo cáo Excel: nạp nguồn, làm sạch, tổng hợp, tính toán, ghi file và kiểm tra kết quả. Đồng thời, tôi đã bổ sung công cụ đối chiếu và kiểm thử thực tế để phát hiện lỗi kỹ thuật cũng như những bất nhất của dữ liệu mẫu.

Kết quả hiện tại là một công cụ hỗ trợ có thể tiếp tục hoàn thiện trên nền tảng đã được kiểm chứng, không chỉ là bản sao thao tác thủ công của một ngày cụ thể. Để vận hành tổng quát và ổn định hơn, phần việc tiếp theo cần tập trung vào các quy tắc còn thiếu và tính đầy đủ, nhất quán của dữ liệu nguồn.

Chưa có số liệu đo thời gian thao tác trước/sau hoặc tỷ lệ giảm lỗi trong vận hành thực tế, vì vậy báo cáo không đưa ra mức tiết kiệm hay hiệu quả định lượng chưa được kiểm chứng.
