# 05 — Lệch hồ sơ hay gặp & RFI

| Lệch | Ví dụ | Xử lý |
|---|---|---|
| Spec chung ≠ BoQ/ML | Spec: cánh ≥45, nẹp gỗ cứng, veneer; BoQ: 39 honeycomb, PVC, laminate | Tính BoQ/ML, RFI |
| Ghi chú loại cửa ≠ chi tiết điển hình | 46/38 vs 48±1/41±1 | Tính theo ghi chú loại cửa, RFI |
| Mã HT giữa ML và BV | 5966OG vs LK 4535A | Tính theo ML, RFI |
| Thiếu mã màu | chỉ ghi "laminate finish" | Tạm tính như loại tương tự, RFI |
| Có trên mặt bằng, không có BoQ | loại cửa không có chi tiết | Loại trừ, RFI |
| KL "N/A" | | Không tính, RFI |

**File RFI gửi CĐT:** song ngữ EN/VN; cột: RFI No · Loại cửa · Tham chiếu · Câu hỏi EN · Câu hỏi VN · Giả định EN · Giả định VN · Reply. **Không** để cột nội bộ (Issue ID, Status). RFI phụ kiện bỏ nếu không chào phụ kiện, ghi 1 dòng loại trừ.
