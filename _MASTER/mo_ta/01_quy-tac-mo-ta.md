# 01 — Quy tắc mô tả costing (user chốt 25/10/2026)

| # | Quy tắc |
|---|---|
| 1 | **Gộp block ≠ gộp mô tả.** Nhiều dòng BoQ cùng cấu tạo được tính chung 1 block Chiet tinh (E = Σ KL, đơn giá mọi dòng link về H block), nhưng mỗi dòng BoQ vẫn có mô tả RIÊNG theo BoQ của chính nó |
| 2 | **Mã hoàn thiện theo dòng**: ghi mã của dòng BoQ đó (len WD-01 ≠ len WD-03), không chép mã của block nguồn/block dùng chung |
| 3 | **"Không bao gồm" theo vị trí**: ván nền dưới giấy WC-03 khác ván nền dưới gương MR-01 hay kim loại MT-02 → mỗi dòng liệt kê đúng phần của nó (xem `03_khong-bao-gom-theo-ma.md`). Nguồn: QUAN_SAT "Phần ngoài phạm vi" + mã VL/HT của dòng BoQ + dòng "Phạm vi tính giá" ở PROJECT |
| 4 | **Chỉ gộp block khi chi phí như nhau**: khác biệt chỉ nằm ở phần không tính, hoặc user chấp nhận coi như nhau. Hoàn thiện khác giá (veneer WD-01 ↔ WD-04, gỗ ↔ sơn) → tách block |
| 5 | **Mô tả chỉ gồm cái được tính + cái không bao gồm — KHÔNG ghi chú nội bộ.** Không đưa vào AJ: số RFI, "bản vẽ chưa thể hiện…", "chi tiết … không có trong hồ sơ", "tạm tính như …", cách quy đổi/giả định tính (vd "góc bo R225 tính như phẳng"). Những ý đó ghi ở ISSUES/RFI, cột "Giả định tính giá" (PROJECT_DATA) hoặc mục "Mã/hạng mục tạm tính" của email (user sửa 4 dòng SOFITEL 2a, 25/10/2026) |
| 6 | **"Tạm tính" chỉ dùng cho vật tư đang tính giá** khi đơn giá/chủng loại chưa chốt, viết như user: "Mây SP-01 Tạm tính 1tr4/m2", "Bản lề 180 độ tạm tính", "Gối … (tạm tính)". Chi tiết chưa thấy trên bản vẽ → không mô tả như đã chắc; nêu câu hỏi ở RFI, không ở AJ |
| 7 | **Thứ tự dòng**: (Kích thước W×D×H mm: — với hàng Nr) → vật liệu/cấu tạo từng lớp → phụ kiện (mã hãng) → vật tư tạm tính (nếu có) → `Không bao gồm:` + từng dòng `+ …`. Cửa dùng mục đánh số 1., 2., … và `Chưa bao gồm …` + "Kết cấu theo BOHO đề xuất" (xem `02_khung-theo-loai.md`) |
| 8 | **Số lượng phụ thuộc block** (vd da bọc SF/cái) được viết bằng công thức tham chiếu Chiet tinh như L022–L025 để tự cập nhật; khi chép block sang dòng khác phải sửa tham chiếu |
| 9 | **Ghi bằng `L.desc` có tô hồng + nhật ký**, sau khi user duyệt bản nháp `MO_TA_<đợt>.xlsx` (cột cmt). Không ghi đè mô tả có sẵn của dòng khác khi gộp (vd F086, F012 giữ nguyên) |
| 10 | **Kiểm khi gộp**: sau khi gộp phải soát mô tả của TẤT CẢ dòng link về block — không dòng nào trống, không dòng nào mang mã/không-bao-gồm của dòng khác |
| 11 | **Học từ chỗ user sửa**: sau khi user sửa mô tả, so bản đã ghi ↔ bản user (cột AJ, `_EDITLOG` cột F) và cập nhật quy tắc/khung ở thư mục này; bản nháp `MO_TA_<đợt>.xlsx` đổi theo bản cuối của user |
