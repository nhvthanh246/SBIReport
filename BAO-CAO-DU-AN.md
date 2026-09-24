# Dự án tự động hoá Báo cáo ngày CCVN

*Báo cáo dành cho Ban lãnh đạo và bộ phận Nhân sự — viết cho người không làm kỹ thuật.*

---

## Tóm tắt trong 30 giây

Mỗi ngày, team CCVN phải làm một bản báo cáo tổng gửi khách hàng SBI. Việc đó trước giờ làm tay:
tải file từ web về, mở Excel lọc bỏ dòng rác, rồi copy-paste hơn **2.700 dòng** vào ba bảng khác nhau,
sau đó gõ thêm một loạt con số tổng hợp.

Chúng tôi đã làm một công cụ chạy trên trình duyệt để thay thế phần lớn việc đó. Bây giờ chỉ cần
**kéo 4 file vào, bấm một nút, tải file báo cáo hoàn chỉnh về**.

Quan trọng không kém: trong lúc làm, công cụ phát hiện ra **ba lỗi dữ liệu đã tồn tại từ lâu**
mà không ai biết — trong đó có lỗi làm báo cáo các ngày đầu tháng **thiếu tới 31% số liệu**.

---

## Trước và sau

**Cách làm cũ**

1. Vào 2 trang web tải 2 file báo cáo về
2. Mở từng file, xoá cột thừa, lọc bỏ những phiên chat bị lỗi
3. Mở file báo cáo của nhân viên, lọc ra đúng những dòng của ngày hôm đó
4. Copy-paste vào ba bảng trong file báo cáo tổng, mỗi bảng một chỗ khác nhau
5. Gõ tay khoảng 23 con số tổng hợp
6. Copy tiếp một cột kết quả KPI sang bảng KPI tháng

**Cách làm mới**

1. Kéo 4 file vào công cụ
2. Bấm **Xử lý**
3. Liếc qua phần cảnh báo xem có gì bất thường không
4. Bấm **Tải file báo cáo tổng** — xong
5. Gõ tay 15 con số mà máy không thể tự biết (xem mục *Còn gì phải làm tay*)

Các bước 2, 3, 4, 6 của cách cũ giờ máy làm hết. Bước copy-paste hơn 2.700 dòng — nơi dễ sai nhất —
đã biến mất hoàn toàn.

> Chúng tôi chưa bấm giờ đo cách làm cũ nên chưa dám nêu con số tiết kiệm bao nhiêu phút.
> Nhưng số thao tác tay giảm từ 6 bước xuống 2 bước, và phần dễ sai nhất thì không còn nữa.

---

## Công cụ này làm gì

Nói đơn giản, nó là một **cái phễu**: đổ dữ liệu thô vào một đầu, ra đầu kia là file báo cáo đã điền sẵn.

| Đổ vào | Ra |
|---|---|
| Báo cáo phiên chat (từ hệ thống Live Support) | |
| Báo cáo case (từ Salesforce) | → **File báo cáo tổng** đã điền đủ 3 bảng dữ liệu, |
| Báo cáo ngày của nhân viên | phần tổng hợp, và toàn bộ bảng KPI tháng |
| File báo cáo tổng của hôm trước | |

Nó cũng **tự nhận biết file nào là file nào** — kéo cả 4 file vào một lượt, thứ tự nào cũng được.

Một điểm cần nhấn mạnh: công cụ **không tạo file mới từ đầu**. Nó lấy đúng file báo cáo tổng của
hôm trước, chỉ thay phần dữ liệu, và **giữ nguyên 100% các bảng pivot, công thức, màu sắc, định dạng**
mà team đã dựng suốt nhiều năm. Điều này rất quan trọng: file báo cáo tổng có tuổi đời từ năm 2020,
bên trong có 4 bảng pivot và hàng nghìn công thức. Làm hỏng nó là mất rất nhiều công khôi phục.

---

## Ba lỗi dữ liệu phát hiện được

Đây có lẽ là phần đáng giá nhất của dự án, vì nó ảnh hưởng đến **số liệu đã nộp cho khách hàng**.

### 1. Lỗi đảo ngày và tháng — ảnh hưởng 13,7% dữ liệu

Trong file báo cáo của nhân viên, **1.188 trên 8.677 dòng bị ghi sai ngày**. Nguyên nhân không phải
do ai cẩu thả, mà do Excel: khi gõ `9/1/2026` với ý là *ngày 1 tháng 9*, Excel cài đặt kiểu Mỹ lại
hiểu thành *ngày 9 tháng 1* rồi tự đổi luôn.

Hệ quả: báo cáo các ngày **từ mùng 1 đến 12** bị thiếu số liệu.

| Ngày | Số dòng bị bỏ sót | Tỷ lệ thiếu |
|---|---|---|
| 07/9 | 193 | **31%** |
| 04/9 | 148 | **23%** |
| 10/9 | 161 | 21% |
| 02/9 | 133 | 20% |

Các ngày từ 13 trở đi không bị ảnh hưởng (vì không có "tháng 13" nên Excel chịu, để nguyên).

Công cụ giờ **tự phát hiện và báo đỏ**, kèm một ô tích để tính bù số liệu thiếu vào.
Nhưng cách sửa tận gốc là chỉnh lại định dạng cột ngày trong file nhân viên.

### 2. Lỗi gõ thừa dấu gạch chéo

Có 8 dòng ghi ngày là `13/09//2026` — hai dấu gạch chéo. Máy không đọc được, nên **nguyên khối giờ làm
của một nhân viên bị bỏ qua** trong ngày đó. Công cụ đã được chỉnh để đọc được cả trường hợp này.

### 3. Bảng tra email có tên trùng — 62 case bị gán sai người

Trong file báo cáo tổng có một bảng tra "tên nhân viên → email". Một người bị ghi **hai lần với hai
email khác nhau**. Kết quả là 62 case của người đó bị chia đôi, và chỉ số KPI của hai người đều sai.

Thú vị là chính file gốc cũng đang mâu thuẫn với bản thân nó: 40 dòng ra email này, 22 dòng ra email
kia, dù dùng chung một công thức. Đề nghị dọn lại bảng đó.

---

## Dùng công nghệ gì, và vì sao

Chúng tôi chọn cách **đơn giản nhất có thể chạy được**, vì công cụ này sẽ do team nghiệp vụ dùng
hằng ngày chứ không phải dân IT.

| Thành phần | Là gì | Vì sao chọn |
|---|---|---|
| **Trang web tĩnh** (HTML/CSS/JavaScript) | Một trang web chạy ngay trên máy, mở bằng cách bấm đúp | Không cần cài đặt, không cần xin cấp server, không cần quyền admin |
| **SheetJS** | Thư viện đọc/ghi file Excel | Miễn phí, phổ biến, chạy được trong trình duyệt |
| **fflate** | Thư viện nén/giải nén | File Excel thực chất là một file nén. Thư viện này cho phép mở ra, sửa đúng phần cần sửa, rồi đóng lại — nhờ vậy giữ nguyên được pivot và công thức |

**Không dùng** server, không dùng cơ sở dữ liệu, không dùng dịch vụ đám mây, không cần đăng nhập.
Toàn bộ công cụ là một thư mục vài file, copy đi đâu cũng chạy.

### Vì sao không dùng cách thông thường

Cách thông thường là đọc file Excel vào rồi ghi ra file mới. Nhưng làm vậy sẽ **phá hỏng** 4 bảng
pivot và hàng nghìn công thức trong file báo cáo tổng — mọi thư viện Excel chạy trên trình duyệt đều
vậy.

Nên chúng tôi làm cách khác: mở file ra như mở một cái hộp, **chỉ thay đúng những tờ giấy cần thay**,
rồi đóng hộp lại. Kết quả đo được: trong 56 thành phần bên trong file, **46 thành phần giữ nguyên
không đổi một byte nào**. 10 thành phần còn lại đổi đều là cố ý.

---

## An toàn dữ liệu

Dữ liệu xử lý ở đây có **thông tin cá nhân khách hàng**: họ tên, ngày sinh, mã khách hàng, số điện thoại.

Vì vậy công cụ được thiết kế để **không có đường nào cho dữ liệu rời khỏi máy**:

- Chạy hoàn toàn trong trình duyệt, trên chính máy người dùng
- Không gửi gì lên mạng — không có server để mà gửi
- Không lưu file lại ở đâu cả
- Không cần tài khoản, không cần đăng nhập

Người dùng mở file lên, xử lý, tải kết quả về. Hết. Không khác gì dùng Excel, chỉ là nhanh hơn.

---

## Làm sao biết công cụ chạy đúng

Đây là phần chúng tôi đầu tư nhiều công nhất, vì báo cáo này gửi cho khách hàng nên **sai một con số
là chuyện lớn**.

Cách kiểm chứng: lấy dữ liệu thật của **ngày 13/9** — cả file đầu vào lẫn file báo cáo tổng mà team đã
làm tay hôm đó — rồi cho công cụ chạy lại, và **so từng ô một** xem có ra giống hệt không.

Kết quả: **49.391 ô khớp, không sai ô nào.**

Ngoài ra, file kết quả còn được **mở bằng Excel thật** để kiểm tra:

- Excel mở bình thường, không báo file hỏng
- Vẫn đủ 11 sheet, 4 bảng pivot — đúng như bản gốc
- Các bảng pivot tự cập nhật đúng số liệu mới
- Mã khách hàng dài 17 chữ số vẫn giữ nguyên, không bị làm tròn
- Các ô do người nhập tay **không bị đụng tới**

Công cụ có sẵn một nút **"Tự kiểm chứng"** để team chạy lại bài kiểm tra này bất cứ lúc nào — đặc biệt
hữu ích khi trang web nguồn đổi định dạng.

---

## Còn gì phải làm tay

Nói thẳng để tránh hiểu nhầm: **công cụ chưa tự động hoá được 100%.**

Còn **15 chỉ tiêu** trong bảng tổng hợp phải gõ tay, vì máy không thể tự biết:

- **8 chỉ tiêu tồn đọng Salesforce** — số liệu này nằm ở 4 báo cáo Salesforce khác mà công cụ chưa đọc
- **2 chỉ tiêu nhân sự** (số người yêu cầu / số người thực tế) — do con người quyết định
- **5 chỉ tiêu đăng ký mới** — từ nguồn khác

Công cụ **liệt kê rõ 15 chỉ tiêu này mỗi lần chạy** để không ai quên.

Ngoài ra, mỗi ngày vẫn phải nạp file báo cáo tổng của hôm trước vào, vì file này **cộng dồn số liệu cả
tháng** — không thể tạo mới từ đầu mà không mất lịch sử.

---

## Rủi ro và bảo trì

**Điểm mạnh:** không có server nghĩa là không có gì để hỏng, không có phí duy trì, không phụ thuộc
nhà cung cấp nào.

**Điểm cần lưu ý:**

| Rủi ro | Cách xử lý đã chuẩn bị |
|---|---|
| Trang web nguồn đổi tên cột | Toàn bộ quy tắc nằm trong một màn hình **Cấu hình** sửa được, không cần lập trình viên |
| File báo cáo tổng đổi cấu trúc | Chạy nút **Tự kiểm chứng** là biết ngay chỗ nào lệch |
| Người dùng nhập sai dữ liệu | Công cụ tự phát hiện và cảnh báo (như ba lỗi nêu trên) |
| Người làm nghỉ / chuyển việc | Công cụ có tài liệu kỹ thuật đi kèm; mọi quy tắc nghiệp vụ ghi rõ trong file cấu hình |

**Về phụ thuộc con người:** đây là điểm chúng tôi muốn nhấn mạnh với Nhân sự. Quy trình cũ nằm trong
đầu một vài người — ai làm lâu thì biết mẹo, người mới vào rất dễ sai. Giờ các quy tắc đó đã được viết
ra thành văn bản và đưa vào công cụ, nên **người mới có thể làm được ngay**, và tổ chức không còn phụ
thuộc vào trí nhớ của cá nhân.

---

## Đề xuất bước tiếp theo

**Nên làm sớm — không tốn công, lợi ích rõ:**

1. **Sửa định dạng cột ngày** trong file báo cáo của nhân viên, để chấm dứt lỗi 13,7% dữ liệu sai ngày
2. **Dọn tên trùng** trong bảng tra email
3. **Rà lại báo cáo các ngày 1–12 đã nộp** — số liệu có thể đã thiếu

**Có thể cân nhắc:**

4. Kết nối thêm 4 báo cáo Salesforce còn lại, để tự động nốt 8 chỉ tiêu đang phải gõ tay
5. Cho công cụ ghi nhớ file báo cáo tổng, để mỗi ngày chỉ cần kéo 3 file

---

## Kết luận

Dự án này không chỉ là "làm cho nhanh hơn". Giá trị lớn nhất nằm ở chỗ: khi buộc phải mô tả chính xác
từng quy tắc cho máy hiểu, chúng tôi phát hiện ra những chỗ quy trình cũ đang âm thầm sai — và những
lỗi đó đã ảnh hưởng đến số liệu gửi khách hàng suốt một thời gian.

Công cụ đã chạy được, đã kiểm chứng kỹ trên dữ liệu thật, và sẵn sàng dùng hằng ngày.
