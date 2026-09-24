# Quy trình làm báo cáo ngày — bản đã đối chiếu với dữ liệu thật

> Bản gốc do bên nghiệp vụ cung cấp. Bản này giữ nguyên 19 bước, chỉ **sửa lại những
> chỗ mô tả chưa khớp dữ liệu** và **bổ sung những bước bị thiếu**, sau khi chạy thử
> chuỗi 13/9 → 14/9 → 15/9 trên file thật và mở kết quả bằng Excel để kiểm chứng.
> Chỗ nào có sửa đều ghi rõ bằng chứng. Phần **[TOOL]** cho biết tool đã tự làm chưa.

---

## Các bước

1. Mở file report của ngày hôm trước, định dạng **AltiuslinkVN_Daily Report_(New)-yyyymmdd.xlsx**.
   Đây là file report tổng đã có sẵn dữ liệu các ngày trước; ta thu thập dữ liệu hôm nay và bỏ vào file này.
   *[TOOL] Đây là "File báo cáo tổng" cần nạp vào tool.*

2. Thay đổi thông tin ngày tháng ở **đầu các tab** cho khớp ngày hôm nay.

   > **Bổ sung — đủ 6 ô, không phải "các tab" chung chung.** Đã dò vị trí trên cả 3 file
   > tổng 13, 14, 15/9 và đều trùng nhau:
   >
   > | Tab | Ô |
   > |---|---|
   > | `Summarize` | `C2` |
   > | `Summarize_team VietNam` | `C2` |
   > | `1. CSKH` | `B2` |
   > | `2. SF Case info` | `B3` |
   > | `3. Other Tasks & Working time` | `C2` |
   > | `Daily KPI Result` | `A2` |
   >
   > Tab KPI tháng (`202609 KPI Result`) có ô ngày riêng ở `A2`, tool xử lý cùng phần KPI
   > vì tên tab đổi theo tháng.
   >
   > *[TOOL] Đã tự động, có báo lại đủ 6 ô trong log.*

3. Vào các tab **1. CSKH**, **2. SF Case Info**, **3. Other Tasks & Working Time** → xoá hết
   thông tin trong bảng **trừ dòng thứ nhất** (giữ format của ô, xoá ngay sau khi copy xong).
   *[TOOL] Đã tự động: tool dùng lại đúng thuộc tính style `s` của ô cũ nên format được giữ
   mà không cần chừa dòng mẫu.*

4. Vào file **Session-Report-xxxxxxxx.xlsx**, tab **Session Report-0**, bỏ cột Topic, Sender name,
   Dob, Cpcode.

   > **Sửa: KHÔNG bỏ cột Ma kh.** Bản gốc ghi bỏ cả `Ma kh`, nhưng cột `ma_kh/顧客コード`
   > trong file tổng có dữ liệu thật: 1.455 / 1.821 / 1.862 dòng có giá trị trên 3 ngày
   > 13, 14, 15/9. Nếu bỏ thì mất hẳn mã khách hàng.
   >
   > Lưu ý ghi kiểu chữ: `ma_kh` và `sender_id` là ID 16–17 chữ số, phải để **dạng text**.
   > Ghi kiểu số sẽ mất chính xác (thành `5.85501074122872E+15`).
   >
   > *[TOOL] Đã tự động, và ghi đúng kiểu text.*

5. Trong cột **Description**, lọc các dòng có giá trị là "lỗi bot" và bỏ toàn bộ dòng đó.
   *[TOOL] Đã tự động (khớp không phân biệt hoa thường, chịu được cả "lỗi bot" lẫn "lỗi").
   Số dòng bị bỏ được báo lại mỗi lần chạy.*

6. Copy toàn bộ dữ liệu trong tab **Session Report-0** sang bảng **1. CSKH** theo cột tương ứng.
   Cột **CHANNEL** để toàn bộ là *CHAT BOT*. Đánh số thứ tự cho bảng và ghi ngày hỗ trợ trùng
   ngày viết báo cáo.

   > **Sửa: giá trị đúng là `CHAT BOT` (có dấu cách), không phải `CHATBOT`.** Toàn bộ
   > 1.555 / 1.929 / 1.997 dòng chatbot trong 3 file tổng đều ghi `CHAT BOT`. Viết liền
   > sẽ thành một kênh khác trong pivot và làm sai số ở `Summarize_team VietNam`.
   >
   > *[TOOL] Đã tự động.*

7. Chuyển qua tab **2. SF Case Info**.

8. Mở file **Daily Report CCVN-xxxxxx.xlsx**, tab **Daily Report CCVN**, copy dữ liệu sang tab
   **2. SF Case Info** của file report tổng.

   > **Bổ sung: cắt ký tự thừa ở cuối tên người lập case.** Export Salesforce có
   > `"Dinh Thi Kim Oanh."` (dấu chấm cuối) trong khi bảng `mail SF` ghi `"DINH THI KIM OANH"`.
   > Không cắt thì 35–50 case mỗi ngày không tra được email.
   >
   > **Bổ sung: cột `案件作成日時` có 2 định dạng.** Export đổi giữa chừng sang
   > `10:09 14/09/2026` thay vì `2026/09/14 10:09` — phải chuẩn hoá về một dạng.
   >
   > *[TOOL] Đã tự động cả hai.*

9. Kiểm tra lại các trường hợp cột **Supporter Mail** = N/A. Mở tab ẩn **mail SF**:
   - Chuyển mail `ly.ntk1@altius-link.vn` thành **NGUYEN THI KHANH LY (3M)**.

   > **Làm rõ — đây là lý do gốc, không phải một thao tác lặt vặt.** Ngày 13/9, bảng
   > `mail SF` có **hai dòng cùng tên** `NGUYEN THI KHANH LY` (dòng 34 → `ly.ntk@`,
   > dòng 43 → `ly.ntk1@`). VLOOKUP luôn lấy dòng ĐẦU, nên `ly.ntk1@` không bao giờ
   > nhận được case nào — 22 case bị gán sai người. Từ 14/9 dòng 43 đã đổi thành
   > `NGUYEN THI KHANH LY (3M)` nên hết trùng.
   >
   > **Hệ quả cho người dùng tool:** tool đọc bảng `mail SF` từ **file tổng bạn nạp vào**.
   > Nếu Salesforce xuất ra một tên chưa có trong bảng đó (ví dụ `(3M)` vào ngày 14/9 khi
   > file nền còn là ngày 13), tool sẽ báo **lỗi đỏ**: *"N tên không tra được email"*.
   > Lúc đó phải **thêm dòng vào tab `mail SF` của file tổng rồi chạy lại** — tool không
   > tự đoán được email.
   >
   > *[TOOL] Không tự sửa (đúng ra không được tự sửa), nhưng luôn soát và cảnh báo:
   > báo tên trùng trong `mail SF`, và báo tên không tra được email.*

10. Chuyển cột **Date** toàn bộ thành ngày báo cáo.
    *[TOOL] Đã tự động.*

11. Chuyển qua tab **3. Other Tasks & Working Time**, mở file **SBI DAILY REPORT xxxxxx**,
    copy toàn bộ tab **3. Other Tasks & Working Time** sang file tổng, set cột DATE là ngày báo cáo.

    > **Bổ sung: cột DATE trong file OP lẫn 2 kiểu.** Vừa là số serial Excel (`46278`)
    > vừa là text (`13/9/2026`). Chỉ đọc kiểu serial sẽ mất khoảng 1/3 số dòng.
    >
    > **Bổ sung: đọc giá trị đã tính, không đọc công thức.** File tổng chứa
    > `6.1176470588235299` ở chỗ file OP có `=F1952-I1952-J1952-M1959` — tức là paste-values.
    >
    > *[TOOL] Đã tự động cả hai.*

12. Quay lại file **SBI Daily Report**, sang tab **1. CSKH**, copy toàn bộ vào sheet **1. CSKH**
    của file tổng, **nối tiếp** phần data đã paste từ bước 6.

    > **Bổ sung: bỏ các dòng rác.** File OP có **37.142 dòng chỉ có ô ngày mà không có
    > nội dung** (do kéo ô ngày xuống quá nhiều dòng trống). Chép hết thì báo cáo phình
    > ra hàng chục nghìn dòng. Ngoài ra có **1 dòng có nội dung nhưng bỏ trống ô ngày** —
    > dòng này không xuất hiện trong bất kỳ file tổng nào của 3 ngày, tức là đúng ra
    > phải bỏ; tool bỏ và báo lại để người làm biết mà đi sửa file nguồn.
    >
    > **Bổ sung: cột TEAM luôn để trống.** Cột TEAM trống 100% trong cả 3 file tổng
    > (152+192+192 dòng ở Other Tasks, 2.218+2.577+2.741 dòng ở CSKH) — là cột chết.
    > Nhưng file OP có 16 ô bị gõ nhầm vào đó (9 ô dấu `` ` `` và 7 ô số `6`). Chép
    > nguyên si thì rác lọt vào báo cáo.
    >
    > *[TOOL] Đã tự động, có báo số dòng bỏ và số giá trị bị xoá.*

13. Quay lại file tổng, mở tab ẩn **Sheet5**, kiểm tra các cột của bảng bên phải đã tương ứng
    với bảng bên trái chưa, kéo hàng cho số hàng hai bên bằng nhau.

    > **Làm rõ — đây là bước dễ quên nhất và hậu quả rất lặng lẽ.** Bảng trái là pivot
    > đếm CHANNEL theo email (cột A..J). Bảng phải là cột phụ O..V, công thức kéo tay.
    > Hiện cột phụ phủ **vừa đủ** pivot, **dư 0 dòng** — chỉ cần hôm nào có thêm một
    > supporter là người mới đó không có giá trị HTKH, kéo theo `Daily KPI Result` tính
    > sai cho người đó mà không báo lỗi gì.
    >
    > Bằng chứng: từ 13/9 sang 14/9 pivot nở từ dòng 27 → 33, và có người đã phải kéo
    > tay công thức xuống theo.
    >
    > **Cảnh báo: công thức cột phụ không nhất quán giữa các ngày.** File 15/9 có dòng
    > 5–9 dùng `O = J-…+I*45` còn dòng 10–33 dùng `O = J-…-I`. Kéo nhầm biến thể sẽ ra
    > số sai. Nên kéo từ **dòng cuối cùng** của bảng xuống.
    >
    > **Trần an toàn: dòng 70.** VLOOKUP của `Daily KPI Result` chỉ tra tới
    > `Sheet5!$A$2:$S$73`, và pivot kế tiếp bắt đầu ở `A74`.
    >
    > *[TOOL] Đã tự động: kéo sẵn tới dòng 70, lấy mẫu từ dòng cuối còn công thức, và
    > cảnh báo nếu phát hiện nhiều biến thể công thức trong cùng một bảng.*

14. Kiểm tra lại các con số tổng cộng màu đỏ ở dưới cùng bảng trong tab **Daily KPI Result**
    đã trùng với kết quả trong **Sheet5** chưa.
    *[TOOL] Chưa tự động — vẫn nên soát mắt. Nhưng xem thêm bước 19 bổ sung bên dưới.*

15. Mở file tổng, tab **2. SF Case Info**, từ cột **Requested By** đổi tên:
    - `Nguyen Thi Thu Hoai` → **TRAN LAN HUONG**
    - `LUU NGOC QUANG` → **NGUYEN THAO NGUYEN**
    - `VU NHU QUYNH` → **NGUYEN PHUONG ANH**

    > **Làm rõ tên: export Salesforce ghi `Pham Thi Thu Hoai`, không phải `Nguyen Thi Thu Hoai`.**
    > Bảng `mail SF` chỉ có một người Hoài duy nhất (`hoai.ptt@altius-link.vn`) nên chắc
    > chắn là cùng một người.
    >
    > **Cảnh báo: thực tế đang làm KHÔNG nhất quán.** Đối chiếu 3 ngày:
    >
    > | | 13/9 | 14/9 | 15/9 |
    > |---|---|---|---|
    > | LUU NGOC QUANG | 19/19 đổi | 20/20 đổi | **23/34 đổi** |
    > | VU NHU QUYNH | 19/19 đổi | 13/13 đổi | (không có case) |
    > | PHAM THI THU HOAI | **0/6 đổi** | 18/18 đổi | **0/10 đổi** |
    >
    > Riêng 15/9: 23 case được đổi đều nằm **trên dòng 325**, 11 case giữ nguyên đều từ
    > **dòng 334 trở đi** — cắt theo vị trí dòng chứ không theo dữ liệu, tức là sửa tay
    > rồi bỏ dở giữa chừng.
    >
    > *[TOOL] Đã tự động và làm đủ 100%, có liệt kê rõ đổi bao nhiêu case cho từng tên.
    > Khi đối chứng với file cũ, phần lệch do luật này được xếp riêng vào nhóm
    > "đã truy được nguyên nhân" chứ không tính là lỗi.*

16. Từ đó quay lại file tổng, tab **Daily KPI Result**, fill số liệu cho các nhân viên,
    đảm bảo nhân viên có tên đều có số liệu, đối chiếu với sheet ẩn **Sheet5**.
    *[TOOL] Đã tự động — tool tính lại cả chuỗi KPI trong JS chứ không phụ thuộc công thức Excel.*

17. Copy cột cuối của tab **Daily KPI Result**, paste vào cột ngày làm báo cáo trong tab
    **xxxx KPI Result**.

    > **Bổ sung: chỉ ghi cho người có trong khối OP.** Sheet KPI tháng còn có dòng
    > JP Staff và một số tài khoản khác do nguồn khác điền — 10 dòng mỗi ngày — không
    > được đụng vào.
    >
    > **Bổ sung: dòng tổng team tính từ SUM của cả nhóm, không phải trung bình các W.**
    >
    > *[TOOL] Đã tự động, có liệt kê 10 dòng bị bỏ qua trong log.*

18. Mở tab **Summarize team Vietnam**, đảm bảo fill hết hàng **Kênh hỗ trợ**.

    > **Làm rõ: tool điền được 8/23 chỉ tiêu, 15 chỉ tiêu còn lại phải nhập tay.**
    >
    > Tool tự tính được: ①CUSTOMER, 1. FACEBOOK (CHAT BOT), 2. ZALO, 3. VIBER, 4. LINE,
    > 5. Telephone, ② OUTBOUND, ⑨CCVN Case check&process.
    >
    > Phải nhập tay: Đăng ký mới (Tổng / ①APP / ②WEBFORM / ③Kênh khác), ②Unprocessed,
    > 8 dòng tồn đọng Salesforce (①→⑧), Order HC, Actual HC — vì số tồn đọng lấy từ
    > 4 báo cáo Salesforce khác, còn Head Count là số nhập tay.
    >
    > *[TOOL] Điền 8 dòng, và liệt kê rõ 15 dòng còn thiếu mỗi lần chạy.*

19. **[BƯỚC BỔ SUNG] Sau khi có file, mở bằng Excel và để nó tính lại.**

    > Pivot trong Sheet5 lấy `Date/日付` làm bộ lọc trang, và bộ lọc này là **bộ lọc tay**
    > liệt kê từng ngày một. Ngày báo cáo mới là "mục mới" nên **mặc định không được chọn**
    > — refresh xong pivot bị lọc sạch, cột phụ Sheet5 về 0, và `Daily KPI Result` ra 0 hết
    > trong khi mọi ô dữ liệu vẫn đúng.
    >
    > Nếu làm tay: sau khi refresh pivot, mở bộ lọc `Date` và **tích chọn ngày mới**.
    > Hoặc bật sẵn tuỳ chọn PivotTable *"Include new items in manual filter"*.
    >
    > *[TOOL] Đã tự động bật tuỳ chọn đó cho cả 7 trường pivot. Đã kiểm chứng bằng Excel
    > thật trên cả 3 ngày: pivot giữ nguyên vùng và Grand Total = đúng số dòng của ngày
    > (2218 / 2576 / 2741).*

---

## Những chỗ tool KHÔNG làm thay

| Việc | Vì sao |
|---|---|
| Thêm dòng mới vào tab `mail SF` | Tool không đoán được email của người mới. Có cảnh báo đỏ khi gặp tên lạ. |
| 15 chỉ tiêu ở `Summarize_team VietNam` | Lấy từ 4 báo cáo Salesforce khác + số nhập tay, không suy ra được từ 3 file đầu vào. |
| 6 dòng fanpage (like/follow) | Do tool của bên khác cung cấp. |
| Soát số tổng màu đỏ ở `Daily KPI Result` (bước 14) | Vẫn nên nhìn mắt một lượt. |

## Những chỗ file mẫu đang tự mâu thuẫn — biết để khỏi hoang mang khi đối chứng

- **14/9, `Summarize_team VietNam` dòng VIBER ghi `0`** trong khi chính tab `1. CSKH`
  của file đó có **3 dòng `Viber T`**. Tool ghi `3`.
- **14/9, ①CUSTOMER ghi `1619`** trong khi số khách chatbot khác nhau thật sự là **1618**.
- **15/9, `1. CSKH` có 1 dòng Topic Level 1 = "11. Gặp nhân viên SBI"** nhưng phiên đó
  trong file Session gốc ghi "12. Đánh giá KH" — sửa tay.
- **File OP dùng chung (`SBI DAILY REPORT_0926.xlsx`) bị OP sửa sau ngày chốt báo cáo.**
  Ví dụ giờ làm của `anh.bn1@` ngày 14/9: file tổng ghi `6`, file OP hôm nay ghi `5`.
  Tool đọc theo file OP hiện tại nên sẽ khác file tổng cũ — đây là chuyện bình thường,
  không phải lỗi.
- **Case Status lệch nhiều nhất (118–127 ô/ngày)** vì export Salesforce được chụp muộn
  hơn lúc chốt báo cáo, case đã chạy tiếp sang trạng thái khác.
