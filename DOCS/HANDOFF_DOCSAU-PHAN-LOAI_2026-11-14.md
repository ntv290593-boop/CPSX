# Đọc sâu (10-tender, bước 7c–7e): vấn đề phân loại trước / trong / sau khi đọc

Ngày: 2026-11-14 · Dự án đo: GBIA_VVIP (`_ACTIVE\GBIA_VVIP\10_TENDER`) · Nhờ chuyên gia xem và đề xuất cách sửa.

## 1. Bối cảnh và ràng buộc

- Skill `10-tender` phục vụ phòng tender. Bước đọc sâu mở bản vẽ PDF để ghi kích thước, cấu tạo, mã hoàn thiện, khối lượng theo bản vẽ, so với BOQ của CĐT, rồi chấm độ tin cậy.
- **Ràng buộc của user: vẫn đọc toàn gói.** Không bỏ hạng mục nào và không chia hạng đọc nhẹ / không đọc. Mục tiêu là **gom việc**: các dòng BOQ hoặc công tác tương tự, cùng nội dung, cùng bản vẽ thì đọc một lần cho cả nhóm.
- Giữ quy tắc bằng chứng: mỗi giá trị phải có tờ và số chi tiết. Gemini (agy) kiểm chéo độc lập.

Tài liệu và code liên quan:
- `SKILL\10-tender\SKILL.md` mục 1e, bảng bước (7c, 7d, 7e)
- `SKILL\10-tender\buoc\7d-docsau.md`, `7e-tincay.md`, `6-banve.md`
- `SKILL\10-tender\tool\docsau.py` (`khoa_cum` dòng 103–106, `cmd_docsau` dòng 134–211, `cmd_gemini` dòng 247+)

## 2. Luồng hiện tại

1. **Bước 6 (banve):** chia agent theo **mã phòng/khu**, mỗi agent gán `Bản vẽ hiệu lực` cho dòng BOQ.
2. **Bước 7 (doichieu):** chia agent theo **1–2 sheet BOQ**.
3. **Lệnh `docsau`:**
   - **Cụm** = các dòng trùng **cả 4 thứ**: mô tả BOQ (chuẩn hóa chữ), `Mã VL`, `Mã spec/FF&E`, bộ tờ `Bản vẽ hiệu lực`.
   - **Nhóm** = gom cụm theo **phòng đầu tiên** (hoặc sheet), cắt mỗi 8 cụm một nhóm. Mỗi nhóm giao cho 1 agent.
4. **Bước 7d:** mỗi agent làm **từng cụm** trong nhóm: tìm bộ tờ, đọc, ghi QUAN_SAT, chạy Gemini trên crop của cụm, ghi BOQ_MAP / SCOPE / PAGES / ISSUES vào file `*.docsau-<n>.json`.
5. **Bước 7e (tincay)** và **`neo`:** chấm và kiểm theo cụm.

## 3. Số đo trên GBIA (14/11)

| Chỉ số | Giá trị | Nguồn |
|---|---|---|
| Cụm / nhóm / agent 7d | 147 cụm → 19 nhóm (18×8 + 3) → 19 agent | `_data\_docsau.txt` |
| Đơn vị của cụm | `no` 67 · `m2` 48 · `m` 32 | `_docsau.txt` |
| Số bộ tờ khác nhau | 110 cho 147 cụm | `_docsau.txt` |
| Lượt tờ theo danh sách giao (cụm × tờ) | 518 lượt, **185 trang khác nhau** (gấp ~2,8 lần) | `_docsau.txt` |
| Trang dùng nhiều nhất | FUR-0006 p20, p21: 19 cụm, rải ở 3 nhóm; FUR-0014 p25/26/28: 12 cụm, ở 3 nhóm | `_docsau.txt` |
| Trang thực đọc (ghi PAGES) | 249 lượt, 188 trang khác nhau; **54 trang bị ≥ 2 nhóm đọc lại**, tối đa 3 nhóm | `PAGES.docsau-*.json` |
| Mẫu mô tả (bỏ mã VL) | 105 mẫu cho 147 cụm; 10 mẫu lặp ở ≥ 2 sheet BOQ | `_docsau.txt` |
| Gemini | 135 file `_gemini_so`, 147/147 cụm có dòng `Kiểm chéo Gemini` | `_data\_gemini_so\`, `QUAN_SAT.docsau-*.json` |
| Chi phí mỗi agent 7d | ~118–177k token, ~28 phút (ảnh Agent map nhóm 1–5 và 7c) | màn hình Agent map |

Ghi chú: PAGES ghi theo nhóm nên không đếm được số lần một agent mở lại cùng trang trong nhóm của mình. Con số 249 là mức sàn.

## 4. Vấn đề

### 4.1 Trước khi đọc (gom và chia việc)

1. **Khóa cụm quá chặt.** Hạng mục cùng loại, cùng bộ tờ nhưng khác mã VL vẫn bị tách cụm.
   - Ví dụ nhóm 1, sheet B02: ốp tường WP-04 (`!27`), ốp tường WD-01 (`!28`) và nẹp MT-01 (`!29`) cùng dùng FUR-0003 p3, p5–p8, nhưng thành 3 cụm và đọc lại cùng bộ mặt đứng 3 lần.
2. **Chia nhóm theo phòng, không theo tờ dùng chung.** Trang chi tiết điển hình dùng cho nhiều phòng (FUR-0006 p20/21) rơi vào nhiều nhóm, nên nhiều agent song song cùng đọc nó.
3. **Không có tầng "bộ tờ" hay "trang chi tiết điển hình".** Đơn vị việc chỉ có cụm. Không có bước nào gom theo tờ trước, rồi gán hạng mục vào tờ.
4. **Hạng mục cùng mã hoặc cùng mẫu mô tả ở nhiều sheet/phòng** (10 mẫu trên GBIA) không được nhận ra là cùng cấu tạo. Mỗi phòng đọc lại cấu tạo, đáng lẽ chỉ cần đếm khối lượng theo phòng.
5. **Bước 6 và bước 7 chia việc theo trục khác nhau** (phòng / sheet BOQ), trục của 7d cũng khác (phòng đầu tiên của cụm). Cùng một bộ tờ bị nhiều agent ở nhiều bước mở lại.
6. **Danh sách giao chỉ in 6 tờ đầu** của mỗi cụm (`docsau.py` dòng 201, `[:6]`). Cụm có hơn 6 tờ thì agent không thấy phần còn lại trong prompt, phải tự tìm lại. Cần xác nhận đây có phải nguyên nhân đọc thiếu hoặc tìm lặp hay không.

### 4.2 Trong khi đọc (7d)

1. **Hướng dẫn đọc theo từng cụm** ("Làm, cho từng cụm" trong `7d-docsau.md`). Không bảo agent: mở 1 tờ thì ghi luôn mọi cụm có trên tờ đó.
2. **`PAGES` chỉ dùng cho lần chạy lại** (`7d-docsau.md` mục "Tiết kiệm token" ý 4). Các agent chạy song song không dùng chung được tờ mà agent khác đã đọc.
3. **Gemini chạy theo crop của từng cụm.** Crop của trang dùng chung có thể bị gửi Gemini nhiều lần. Chưa đo được bao nhiêu lần trùng.
4. **48 cụm m2 và 32 cụm m** không được tự đo khối lượng (quy tắc 5 của 7d). Phần đọc của các cụm này chủ yếu là mã hoàn thiện, tờ căn cứ và phạm vi. Nếu gom theo bộ mặt đứng thì một lần đọc phục vụ được nhiều cụm.

### 4.3 Sau khi đọc (gộp, kiểm, chấm)

1. **Kết quả ghi theo cụm, không theo tờ.** Một trang chi tiết điển hình do 2–3 agent đọc độc lập có thể cho giá trị khác nhau, nhưng **không có bước so chéo** cùng tờ giữa các nhóm.
2. **`neo` và `tincay` kiểm theo cụm.** Không có kiểm "cùng tờ, cùng chi tiết thì giá trị phải giống nhau giữa các cụm".
3. **Không có sổ dùng lại cho lần sau.** Lần TCQ tiếp theo chạy `--thaydoi` sẽ đọc lại theo cụm bị ảnh hưởng. Không có lớp "kết quả đọc theo tờ + revision" để tờ không đổi thì khỏi đọc lại.

## 5. Hướng đã nghĩ (chưa làm, chờ chuyên gia)

1. Thêm tầng **đơn vị đọc = bộ tờ**. Gom các cụm theo độ trùng tờ (đồ thị cụm–tờ), mỗi agent nhận một bộ tờ cùng toàn bộ cụm trên đó. Mở mỗi tờ một lần, ghi cho mọi cụm.
2. **Trang chi tiết điển hình** (dùng ≥ N cụm) được đọc riêng một lần thành bản ghi dùng chung. Các cụm tham chiếu bản ghi đó.
3. **Gộp cấu tạo theo mã/mẫu mô tả qua nhiều phòng.** Đọc cấu tạo một lần, mỗi phòng chỉ đếm khối lượng.
4. **Sổ đọc theo tờ + revision** để cả các agent song song lẫn lần TCQ sau dùng lại.
5. Kiểm sau đọc: cùng tờ + chi tiết thì giá trị phải nhất quán giữa các cụm. Lệch thì sinh issue.

Ước lượng lượt mở trang: 518 (theo danh sách) → khoảng 185–240. Số agent chưa ước lượng được, cần chạy thử cách chia mới.

## 6. Câu hỏi cho chuyên gia

1. Nên đặt đơn vị việc là **bộ tờ**, **cụm** hay kết hợp (đọc theo tờ, ghi theo cụm)? Cách nào giữ được quy tắc bằng chứng mà giảm lặp nhiều nhất?
2. Tiêu chí gom các cụm "tương tự": theo độ trùng tờ, theo mẫu mô tả bỏ mã VL, theo mã hạng mục (BF-01B…), hay cả ba? Ngưỡng nào?
3. Trang dùng chung giữa các agent song song: đọc trước thành bản ghi chung (chạy tuần tự trước 7d), hay dùng một sổ chung có khóa?
4. Gemini: gửi theo tờ/chi tiết thay vì theo cụm có làm giảm giá trị kiểm chéo không?
5. Có nên thống nhất trục chia việc của bước 6, 7 và 7d để bộ tờ chỉ bị mở ở một bước?
6. Giới hạn `[:6]` tờ trong danh sách giao: bỏ hẳn hay thay bằng danh sách đầy đủ trong file riêng?

## 7. Trạng thái

- GBIA: 7d đã chạy đủ 19 nhóm theo cách cũ. Kết quả vẫn dùng được, vì sửa cách chia không đổi định dạng ghi theo cụm.
- Chưa sửa code hay skill nào cho vấn đề này.
