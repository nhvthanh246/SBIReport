# Hướng dẫn làm báo cáo ngày bằng tool

Dành cho nhân viên OP làm báo cáo ngày của team CCVN (khách hàng SBI).

Tool làm thay phần copy/paste và tính toán. Phần cần đầu óc con người — số liệu từ
Salesforce, Head Count, ý kiến khách hàng — vẫn phải nhập tay. Tài liệu này nói rõ
tool làm gì, bạn phải làm gì, và **chỗ nào bắt buộc kiểm tra lại**.

---

## Trước khi bắt đầu

**Mở tool:** bấm đúp vào `tool/index.html`. Nó mở bằng trình duyệt như một trang web
bình thường. Không cần cài gì, không cần mạng.

**Về bảo mật:** file của bạn **không rời khỏi máy**. Tool chạy hoàn toàn trong trình
duyệt, không gửi gì lên mạng, không lưu file lại. Đóng tab là mất hết. Vì dữ liệu có
tên, ngày sinh, mã KH và số điện thoại của khách nên tool được làm như vậy ngay từ đầu.

**Bốn file cần chuẩn bị:**

| File | Lấy ở đâu | Ghi chú |
|---|---|---|
| **File báo cáo tổng** | Bản của **ngày hôm trước** | `AltiuslinkVN_Daily Report_(New)-yyyymmdd.xlsx` |
| **Session Report** | Tải từ FPT.AI Live Support | `Session-Report-xxxxxxxx.xlsx` |
| **Daily Report CCVN** | Export case từ Salesforce | `Daily Report CCVN-xxxxxx.xlsx` |
| **Báo cáo ngày nhân viên** | File OP dùng chung | `SBI DAILY REPORT_MMYY.xlsx` |

---

## Các bước làm

### 1. Nạp file

Kéo **cả 4 file** vào ô lớn giữa màn hình. Tool tự nhận biết file nào là file nào —
không cần thả đúng chỗ, không cần đổi tên.

Nếu có nhiều file OP (ví dụ mỗi team một file), thả hết vào, tool gộp lại.

Nạp đủ 4 file thì nút **Xử lý** sáng lên. Bấm vào.

### 2. Đọc kết quả

Tool hiện:

- **Ngày báo cáo** — tự suy ra từ dữ liệu. Nếu sai, sửa ở ô *Đổi ngày* rồi bấm *Áp dụng*.
- **Số dòng** của từng khối.
- **Danh sách cảnh báo** — phần quan trọng nhất, xem mục [Cảnh báo](#cảnh-báo-nghĩa-là-gì) bên dưới.

> ⚠️ **Đổi file sau khi đã bấm Xử lý thì phải bấm Xử lý lại.** Tool sẽ làm mờ màn hình
> và khoá nút tải để nhắc bạn. Nếu không, bạn sẽ tải ra file dựng từ dữ liệu cũ.

### 2b. Dán số follow fanpage

Trong phần kết quả có ô **"Số follow fanpage"**.

1. Mở extension **Facebook Fanpage Stats** trên thanh công cụ Chrome
2. Bấm **Lấy số liệu hôm nay** — extension tự mở 3 fanpage và đọc số
3. Bấm **Copy**
4. Dán vào ô trong tool

Tool đọc ra 3 số và hiện lại để bạn đối chiếu. Dán thiếu hoặc sai định dạng thì tool
báo ngay chứ không đoán bừa thứ tự.

> **Số like thì sao?** Phía SBI xác nhận chỉ cần ước lượng theo các ngày trước. Đối
> chiếu dữ liệu thật thì 3 số like giữ nguyên suốt cả tháng, nên tool **tự chép lại
> của ngày gần nhất** — bạn không phải làm gì.
>
> Nếu số like vừa đổi so với ngày trước đó, tool sẽ cảnh báo. Nhờ luật này mà phát
> hiện DCOM bị gõ nhầm `3111k` (thừa một số 1) từ ngày 14/9, trong khi cả tháng là `311k`.

Bỏ qua bước này cũng được: 3 dòng follow để trống, phần còn lại vẫn chạy bình thường.

### 3. Tải file báo cáo tổng

Bấm **Tải file báo cáo tổng**. Đây là file hoàn chỉnh, đã điền sẵn mọi thứ tool làm được.

File tải về **giữ nguyên** toàn bộ định dạng, PivotTable, công thức và các sheet khác
của file gốc — tool chỉ sửa đúng những ô cần sửa.

Nếu chỉ cần từng khối để dán tay, dùng các nút *Tải .xlsx* / *Copy* ở phần dưới.

### 4. Mở file bằng Excel

**Bắt buộc mở bằng Excel thật một lần** trước khi gửi đi. Khi mở, Excel sẽ tự tính lại
công thức và làm mới PivotTable — đây là lúc các con số ở `Daily KPI Result` và `Summarize`
hiện ra. Mở bằng trình xem nhanh hoặc Google Sheets sẽ **không** thấy đúng.

### 5. Nhập tay phần còn lại

Xem mục [Những gì phải nhập tay](#những-gì-phải-nhập-tay).

### 6. Soát lại

Xem mục [Danh sách kiểm tra](#danh-sách-kiểm-tra-trước-khi-gửi).

---

## Tool tự làm những gì

| Tab | Tool làm |
|---|---|
| **1. CSKH** | Gộp dữ liệu chat bot + dữ liệu file OP, đánh số thứ tự, điền ngày, lọc bỏ dòng "lỗi bot" |
| **2. SF Case info** | Chép case từ Salesforce, tra email supporter, đổi tên người lập case theo quy tắc |
| **3. Other Tasks** | Chép giờ làm từ file OP, điền ngày |
| **Summarize_team VietNam** | 22 chỉ tiêu + 6 dòng fanpage (xem bảng dưới) |
| **Summarize** | Kéo công thức sang cột ngày mới (sheet này tự rút số từ Summarize_team) |
| **Sheet5** | Kéo công thức cột phụ xuống cho dư chỗ khi có thêm người |
| **Daily KPI Result** | Cập nhật ô ngày |
| **202609 KPI Result** | Tính và điền hiệu suất cho từng người + dòng tổng team |
| **Ô ngày đầu 6 tab** | `Summarize`, `Summarize_team`, `1. CSKH`, `2. SF Case info`, `3. Other Tasks`, `Daily KPI Result` |

**22 chỉ tiêu tool điền ở `Summarize_team VietNam`:**

- Dòng 11–16: ①CUSTOMER, FACEBOOK (CHAT BOT), ZALO, VIBER, LINE, Telephone
- Dòng 18: ② OUTBOUND
- Dòng 28: ⑨CCVN Case check&process
- Dòng 33–45: toàn bộ khối **Nội dung hỗ trợ** (13 dòng, đếm theo Topic Level 1)
- Dòng 57: ② Số lượng comment đã hỗ trợ
- Dòng 61–66: like/follow 3 fanpage — follow lấy từ extension, like chép của ngày gần nhất

---

## Những gì phải nhập tay

Tool không bịa số. Những mục dưới đây không có trong 4 file đầu vào nên **bắt buộc bạn điền**.

### Ở `Summarize_team VietNam`

| Dòng | Mục | Lấy ở đâu |
|---|---|---|
| 7, 8, 9 | ① APP · ② WEBFORM · ③ Kênh khác (CCC) | Báo cáo đăng ký mới |
| 17 | ②Unprocessed/未対応 | Salesforce |
| 20–27 | 8 dòng tồn đọng 『CCVN』team | 4 báo cáo Salesforce khác |
| 29, 30 | Order HC · Actual HC | Số nhân sự, tự khai |
| 46–51 | Ý kiến khách hàng (6 dòng) | Đọc và phân loại tay |
| 53–55 | Chi tiết Outbound (3 dòng) | Nếu ngày đó có Outbound |
| 56 | Others | |
| 58 | ③ Khung giờ hỗ trợ | Giờ chốt ca, tự khai |
| 59 | ④ Số lượng OP (không tính TL, JP, senior) | Đếm tay theo chức danh |
| 60 | ⑤ Số lượng claim/ý kiến | |

> **Không cần điền gì ở sheet `Summarize`.** Mọi ô ở đó là công thức rút thẳng từ
> `Summarize_team VietNam`. Điền xong sheet kia thì sheet này tự có số.

### Ở `202609 KPI Result`

Một số dòng tool **cố tình để trống**, và sẽ nói rõ lý do trong log:

- **Người dùng định mức KPI riêng** — ví dụ NGO QUOC BAO dùng định mức 10/10/10/15 thay
  vì 15/30/17/20 như mọi người. Tool tính theo định mức chuẩn sẽ ra sai (0.74 thay vì 1.86),
  nên thà để trống cho bạn nhập tay.
- **Dòng Leader / JP Staff** — do nguồn khác điền.

---

## Cảnh báo nghĩa là gì

Tool hiện cảnh báo theo 3 mức. **Mức đỏ thì đừng bỏ qua.**

### 🔴 Đỏ — phải xử lý trước khi gửi báo cáo

**`N tên không tra được email trong bảng mail SF`**

Salesforce xuất ra một tên chưa có trong tab ẩn `mail SF` của file báo cáo tổng. Những
case của người đó sẽ **thiếu email supporter**, kéo theo KPI của họ bằng 0.

→ **Cách xử lý:** mở tab `mail SF` trong file báo cáo tổng (bản bạn nạp vào), thêm dòng
cho người đó, lưu lại, rồi chạy lại tool.

> Chuyện này đã xảy ra thật: ngày 14/9 Salesforce bắt đầu xuất tên
> `Nguyen Thi Khanh Ly (3M)` trong khi file ngày 13 chưa có tên đó — 18 case bị thiếu email.

**`Sheet mail SF có N tên bị ghi trùng`**

Hai dòng cùng tên nhưng khác email. Công thức tra email chỉ lấy **dòng đầu tiên**, nên
người ở dòng sau không bao giờ nhận được case nào.

→ **Cách xử lý:** xoá dòng thừa. Nếu đúng là hai người khác nhau thì phải thêm cách phân
biệt vào tên (ví dụ thêm `(3M)` như đã làm với chị Khánh Ly), vì tra theo tên không đủ.

### 🟡 Vàng — nên xem qua

**`Đã đổi tên người lập case theo bảng requestedByAliases`**

Tool đã đổi tên theo quy tắc nghiệp vụ và liệt kê đổi bao nhiêu case cho từng người.
Liếc qua xem có hợp lý không. Nếu quy tắc đã thay đổi, báo người quản lý tool sửa lại.

**`N người có case/phiên chat hôm nay nhưng KHÔNG có khối giờ trong 3. Other Tasks`**

Hiệu suất tính bằng *tổng việc chia tổng giờ*. Không có giờ thì mẫu số bằng 0, buộc phải
ra 0. Tool liệt kê đích danh những người này.

→ **Cách xử lý:** kiểm tra file OP xem có thiếu khối 8 dòng giờ làm của họ không. Nếu họ
thật sự không khai giờ thì số 0 là đúng, bỏ qua.

**`Bỏ qua N dòng chỉ có ô ngày, không có nội dung`**

File OP bị kéo ô ngày xuống hàng chục nghìn dòng trống. Tool tự bỏ. Nên dọn bớt ở file
nguồn cho file nhẹ đi.

**`1 dòng có nội dung nhưng bỏ trống ô ngày — bị bỏ qua`**

Có người quên điền ngày. Tool không đoán được dòng đó thuộc ngày nào nên bỏ.

→ **Cách xử lý:** vào file OP điền ngày cho dòng đó nếu muốn tính.

**`Bỏ N giá trị ở cột luôn-để-trống`**

Cột TEAM bị gõ nhầm ký tự (dấu `` ` `` hoặc số `6`). Cột này trống ở mọi dòng của file
báo cáo tổng nên tool ghi trống. Không ảnh hưởng gì.

**`Có N giá trị CHANNEL / Topic Level 1 chưa gán nhóm`**

Xuất hiện một kênh hoặc một topic mới chưa có trong cấu hình → **không được cộng vào bảng
tổng hợp**. Nếu đây là kênh/topic thật sự mới, báo người quản lý tool thêm vào.

**`Sheet5 cột phụ đang có N kiểu công thức khác nhau`**

Bảng phụ ở Sheet5 bị trộn nhiều phiên bản công thức. Tool kéo theo dòng cuối cùng. Nên
mở Sheet5 kiểm tra lại các dòng dùng kiểu cũ.

### 🔵 Xanh — chỉ để biết

**`Tool điền được 22 chỉ tiêu, còn N chỉ tiêu phải nhập tay`** — kèm danh sách đầy đủ.
Dùng làm checklist khi nhập tay.

---

## Danh sách kiểm tra trước khi gửi

Mở file bằng Excel rồi soát theo thứ tự này:

- [ ] **Ngày ở đầu 6 tab** đúng ngày báo cáo.
- [ ] **Tab `Daily KPI Result`** có số, không phải 0 hết. Nếu 0 hết là pivot chưa
      được làm mới — đóng file, mở lại, chọn *Enable Content* nếu Excel hỏi.
- [ ] **Số tổng màu đỏ ở cuối `Daily KPI Result`** trùng với kết quả trong `Sheet5`.
- [ ] **Tab `Summarize`** đã có số ở cột ngày hôm nay (không trắng).
- [ ] **Đã nhập tay** đủ các dòng ở mục [Những gì phải nhập tay](#những-gì-phải-nhập-tay).
- [ ] **Đã xử lý hết cảnh báo đỏ.**
- [ ] `Summarize_team VietNam` dòng 10 (*Kênh hỗ trợ*) bằng tổng các dòng 12–16.

---

## Đối chứng — khi nào dùng

Trong tool có mục **Đối chứng** (gập lại ở cuối trang). Nạp thêm *file tổng chuẩn* của
đúng ngày đang làm, tool sẽ so từng ô và chỉ ra chỗ lệch.

Dùng khi: bạn nghi tool làm sai, hoặc muốn kiểm tra sau khi ai đó sửa cấu hình.

**Đọc kết quả đối chứng:** nhiều chỗ "lệch" là bình thường, không phải lỗi tool. Tool tự
xếp riêng những chỗ đã truy được nguyên nhân. Các nguyên nhân hay gặp:

| Chỗ lệch | Vì sao |
|---|---|
| `Case Status` lệch cả trăm ô | Export Salesforce chụp muộn hơn lúc chốt báo cáo — case đã chạy sang trạng thái khác. Không tránh được. |
| Giờ làm ở `3. Other Tasks` lệch vài ô | File OP dùng chung bị nhân viên sửa **sau** ngày chốt báo cáo. |
| `202609 KPI Result` lệch vài ô | Hệ quả trực tiếp của dòng trên. |
| `Requested By` + email lệch | Bản chuẩn áp dụng đổi tên không nhất quán (có ngày quên, có ngày sửa nửa chừng). Tool luôn làm đủ 100%. |
| Mục **6b** báo "bản chuẩn sửa lại số của ngày cũ" | Người làm sửa số của ngày trước trong lúc làm báo cáo hôm sau. Tool giữ nguyên theo file nền — đúng. Nếu số mới mới là số đúng, phải sửa trong **file nền** rồi chạy lại. |

---

## Gặp sự cố

**Nút *Xử lý* không sáng** — chưa đủ 4 file. Xem lại các thẻ ở mục *Nạp file*, thẻ nào
chưa có dấu tích là còn thiếu.

**Màn hình bị mờ, nút tải bị khoá** — bạn đã đổi file sau khi bấm *Xử lý*. Bấm *Xử lý* lại.

**Excel báo file hỏng khi mở** — báo ngay người quản lý tool, kèm theo cả 4 file đầu vào.
Đừng tự sửa.

**Ngày báo cáo tool đoán sai** — sửa ở ô *Đổi ngày* rồi bấm *Áp dụng*. Tool tính lại toàn bộ.

**Số ở `Daily KPI Result` bằng 0 hết** — pivot chưa được làm mới. Đóng file rồi mở lại
bằng Excel. Nếu vẫn 0, vào tab `Sheet5`, bấm chuột phải vào bảng pivot → *Refresh*.

**Báo cáo thiếu khoảng 1/3 số dòng** — file OP có cột ngày lẫn hai kiểu (số và chữ).
Tool đọc được cả hai, nhưng nếu thiếu nhiều thì kiểm tra lại cột `Date` trong file OP.

---

## Một số điều nên biết

**Tool đọc bảng `mail SF` từ file báo cáo tổng bạn nạp vào**, không phải từ danh sách
cố định. Nghĩa là thêm người mới thì sửa trong file tổng, không cần đụng vào tool.

**Tool không bao giờ sửa số của ngày cũ.** Nó chỉ ghi vào cột của ngày đang làm. Nếu số
ngày cũ sai, bạn sửa tay trong file — tool sẽ giữ nguyên ở lần chạy sau.

**Bản thân file OP là nguồn sống.** Nhân viên có thể sửa giờ làm sau khi báo cáo đã chốt.
Tool đọc theo bản hiện tại, nên chạy lại báo cáo cũ có thể ra số khác lúc đầu. Đây là
chuyện bình thường.

**Thứ tự dòng trong `1. CSKH`:** toàn bộ dòng chat bot trước, rồi mới đến dòng từ file OP.
