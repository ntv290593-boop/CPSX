# mo_ta — Mô tả costing theo dòng BoQ (dùng chung mọi dự án)

Đọc thư mục này TRƯỚC khi viết hoặc kiểm cột ghi chú costing (SOFITEL: `MATRIX FITFUR-LOOSE`!AJ "GHI CHÚ COSTING"), mô tả hạng mục trong file TONG_HOP gửi khách, hoặc mục "Chưa bao gồm" trong email trả giá (`_MASTER\mail\`).

| File | Nội dung |
|---|---|
| `01_quy-tac-mo-ta.md` | Quy tắc user đã chốt: gộp block ≠ gộp mô tả, viết theo dòng, tạm tính, không bao gồm |
| `02_khung-theo-loai.md` | Khung câu theo loại đồ vật, trích từ mô tả user viết (SOFITEL RE00) |
| `03_khong-bao-gom-theo-ma.md` | Mã vật liệu/hoàn thiện → câu "Không bao gồm" + chỗ lệch văn phong đã thấy |

Bản mô tả của từng đợt nằm ở dự án: `_ACTIVE\<DA>\02_takeoff\MO_TA_<đợt>.xlsx` (cột cmt để user duyệt) + `MO_TA_<đợt>_lenh.txt` (lệnh `L.desc` cho Claude trong Excel). Script mẫu: `_MASTER\takeoff\mo_ta_SOFITEL_2a.py` (soạn nháp). Ghi vào file CPSX: Claude trong Excel `L.desc`, hoặc khi file đang ĐÓNG thì dùng `py -I _MASTER\takeoff\ghi_mo_ta_com.py <CPSX> <MO_TA.xlsx>` (Excel riêng, tắt macro, sao lưu trước, tô hồng + ghi `_EDITLOG` như L.desc, đọc lại từ đĩa để xác nhận). Sau khi user sửa: chỉnh bản nháp theo bản user rồi chạy lại script — ô đã đúng thì không ghi, chỉ cập nhật `_EDITLOG` thành "đã học" (như `L.learned`).

Liên quan: `_MASTER\schemas\README.md` (mục "Gộp block ≠ gộp mô tả"), `_MASTER\modules\SCAN-NHAN-LOAI.md` (trường QUAN_SAT "Phần ngoài phạm vi" là nguồn của "Không bao gồm").
