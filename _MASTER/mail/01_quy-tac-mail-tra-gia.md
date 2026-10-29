# 01 — Quy tắc email trả giá (user chốt 23/10/2026)

| # | Quy tắc |
|---|---|
| 1 | **Reply** đúng email yêu cầu báo giá (giữ To/Cc gốc), **không cc anh Lê Hoàng Bửu Lộc** |
| 2 | Mở đầu: `Dear anh <tên>,` |
| 3 | **Ngắn gọn** — mail chỉ là lời dẫn, chi tiết để file đính kèm. KHÔNG bảng giá, KHÔNG dòng tổng, không đi sâu kỹ thuật. Bố cục (mỗi mục vài gạch đầu dòng 1 dòng): `Em gửi CPSX như sau:` → **Cơ sở tính** → **Mã/hạng mục tạm tính** → **Đơn giá NCC đã dùng (chưa VAT)** → **Tiêu chuẩn đang tính** → **Chưa bao gồm** → **RFI cần làm rõ** (chỉ tiêu đề, xem file) → `Chi tiết anh vui lòng xem theo file đính kèm` → `Best Regards,` + chữ ký Outlook |
| 4 | Đính kèm: file CPSX + file RFI + file tổng hợp `<yymmdd>_<DA>_TONG_HOP.xlsx` (chỉ giá trị, không link/VBA): tổng quan dự án, danh mục cửa & giá, kết cấu, tiêu chuẩn & đơn giá NCC, tạm tính/chưa bao gồm/loại trừ, tóm tắt RFI. **Không** kèm hồ sơ thầu |
| 5 | Mọi file gửi gom vào `_ACTIVE\<DA>\05_mail\` (bản nháp email + bản sao file đính kèm) |
| 6 | Claude chỉ soạn nháp, **không tự gửi** |
| 7 | Định dạng (mail đã gửi NORTON 23/10): HTML, Arial 10, không bảng, gạch đầu dòng bullet. **In đậm**: `Dear anh …,`, 7 tiêu đề mục, đơn giá NCC chính (vd `780.000 đ/tấm`), cảnh báo quan trọng ở Chưa bao gồm (vd chống thấm chân khung bao), `Best Regards,`, tên ký |
| 8 | Mỗi mục 2–4 dòng, chỉ ý chính. KHÔNG đưa vào mail: số tổng/chênh lệch, chi tiết cấu tạo từng lớp, dày tường, liệt kê RFI, giá vật tư phụ — để ở TONG_HOP/RFI. Cơ sở tính: 1 dòng nguồn (BoQ + ML + BV, hoàn thiện, nguồn giá) + 1 dòng phương án đang chào (đổi ở OverView › "Kết cấu cửa") |
| 9 | Trước khi đính kèm: kiểm file đính kèm là bản **mới nhất** trong 05_mail (so dung lượng/giờ sửa) |
| 10 | Mẫu thật: `_ACTIVE\NORTON\05_mail\261023_NORTON_email_SENT.txt` |

Email nội bộ (Tender BOHO) → được gửi file CPSX. Gửi ra ngoài (CĐT/tổng thầu) → chỉ BoQ đã điền giá + RFI, không gửi CPSX.
