# Kết quả chạy tầng A trên toàn bộ 271 DWG của SOFITEL

Ngày: 2026-10-24 · Nối tiếp `KET-QUA-GD3-TANG-A_2026-10-24.md` (5 tờ pilot). Ghi chú người dùng: shop drawing DOT 01/02 **cùng phiên bản** (trả lời câu 6 mục 13 của plan v2; câu 7 về bản chốt trong các bản trùng tên vẫn mở).

Dữ liệu: `_ACTIVE\SOFITEL\01_scan\_dwg_tangA\` (`summary.csv`, `version_check2.csv`, `json\` 271 file). DXF trung gian (hơn 1 GB) nằm ở thư mục tạm, tái tạo được bằng ODA. Script: `_MASTER\scan\dwg_batch.py`, `dwg_version_check.py` (bản git ở `_SCRIPTS\scan\`, chưa commit).

## Số liệu

| Chỉ số | Kết quả |
|---|---|
| ODA chuyển DWG → DXF | 271/271, khoảng 6 phút |
| ezdxf đọc được | 271/271, không lỗi, tối đa 18 giây/tờ, cả lô 4,6 phút (6 tiến trình) |
| Tổng số kích thước (DIMENSION) | 14.013 (trung vị 37/tờ, tối đa 214) |
| Tờ không có kích thước nào | 47 (nhóm CO, GA, LL, SK…: mặt bằng/trần, chỉ có block, text, leader) |
| Text kích thước bị gõ đè | 868 (6,2%): chữ như VARIES/EQ/TYP 521; `<> + ghi chú` như "<> (STRUCTURAL OPENING)" 243; số khác 99; dạng "250/300" 5. 140/271 tờ không có gõ đè nào |
| Gõ đè có giá trị lệch giá trị đo | 252 (1,8% tổng kích thước), là vùng cần người xem |
| Dimstyle | 14 loại (hệ số 1, 2, 5, 10, 15, 20, 25, 50, 100; DIMRND 0 hoặc 5). Hơn nửa số kích thước dùng dimstyle `1-1-PS-25` hoặc `-10` có DIMRND = 5 |
| Tờ có tiêu đề chi tiết đọc được | 271/271 có INSERT `TITLE/Nº`; 1 tờ không đọc được số bản vẽ trong khung tên (`ID-JS-65-709A`) |
| Kích thước chưa gắn được vào chi tiết | 1.847/14.013 (13,2%). Tập trung ở mặt bằng/mặt đứng khu vực chỉ có một view (ví dụ `ID-PS-66-101`: 124/124). Trên các tờ chi tiết điển hình (TY) chỉ 2,8–7% |

## Kiểm phiên bản DWG ↔ PDF (ghép theo tên file)

| Kết quả | Số tờ |
|---|---|
| Khớp (cùng chữ phiên bản, cùng ngày nếu đọc được) | **192** |
| **Lệch** | **4**: `ID-GA-L18-001`, `L19-001`, `L20-001` (DWG rev C, PDF rev D, cùng tháng 01/26) và `ID-CO3-50-302` (DWG C, PDF B) |
| Chưa đọc được khung tên (bộ đọc PDF chưa đọc được tờ khu vực, ví dụ LL, SK, GK) | 50 |
| Không ghép được PDF theo tên | 25 |

Nghĩa là: trong số cặp đọc được, 192/196 khớp. 4 tờ lệch cần xem tay (đặc biệt GA-L18/19/20: PDF mới hơn DWG, nên số lấy từ DWG có thể cũ). 75 tờ còn lại chưa kiểm.

## Phát hiện

1. **Tầng A chạy ổn định ở quy mô thật** (271/271), nên có thể dùng làm nguồn chính cho kích thước trên tờ chi tiết điển hình (TY) ở SOFITEL.
2. **Gõ đè có hai loại cần xử lý khác nhau:** (a) chữ/ghi chú (VARIES, EQ, `<> + ghi chú`, khoảng 760 chỗ): giá trị đo vẫn dùng được, chỉ gắn nhãn ghi chú; (b) số khác giá trị đo (khoảng 100 chỗ + 5 dạng "a/b"): giá trị hiển thị là ý đồ thiết kế (đoạn cắt, kích thước tổng cố định), không tự nhận một trong hai, chuyển người xem.
3. **DIMRND = 5 phổ biến** nên giá trị hiển thị có thể lệch đến 2,5 mm so với giá trị thật; phép cộng chuỗi phải có dung sai (đã ghi trong plan v2 mục 6.1).
4. **Mặt bằng/trần/GA (47 tờ không có DIMENSION)** chứa nhiều block, text, leader: dữ liệu cần lấy ở đó là số đếm (cửa, thiết bị, tag) từ block, không phải kích thước. Đó chính là nguồn "đếm lấy từ mặt bằng" mà plan yêu cầu; chưa thử.
5. **Chưa chạy 37 DWG shop drawing DOT 01/02** (khác thư mục, do mình vẽ; người dùng xác nhận cùng phiên bản).

## Việc tiếp theo (đề xuất)

1. **Trích số đếm từ block/text trên mặt bằng** (47 tờ không kích thước và các tờ khu vực): liệt kê tên block, thuộc tính, vị trí, đếm theo loại.
2. **Hoàn thiện gắn chi tiết** cho tờ khu vực một view (gán toàn bộ kích thước cho view duy nhất) và trích ghi chú/leader/nhãn vật liệu.
3. **Xem tay 4 tờ lệch phiên bản**, và mở rộng bộ đọc khung tên PDF cho 50 tờ chưa đọc được.
4. **Dùng DWG làm đáp án chấm Gemini và Claude** trên PDF (GĐ3b/GĐ2), chỉ nhận kích thước không bị gõ đè và đã khớp phiên bản.
5. **Chạy 37 DWG shop drawing DOT 01/02** cùng pipeline.
6. Áp kết quả pilot vào bản sao v9 (GĐ1) với F023 đã sửa.
