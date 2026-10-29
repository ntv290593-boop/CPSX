# Bước 4 — dwg: sổ bản vẽ + bảng KHU → mã phòng

**Ghi:** `KHU.dwg.json`, `DWG.dwg.json` (`Tên tờ`, `Mã phòng`, `Trạng thái`, `Ghi chú`), `VIEC_MO.dwg.json`.
Tool `dwg` đã lập: mỗi `Số BV` (lấy từ tên file) có ngày issue + đợt hiệu lực, file PDF/DWG, số bản, bản cũ bị thay (bản mới nhất theo ngày issue rồi theo thứ tự đợt). File vai trò Ngoài scope/Mail không tính.

## Làm
1. Đọc README (mục bẫy). `loc <GOC> DWG --cot "Số BV,Khu (thư mục),Ngày issue hiệu lực,Số bản" --max 400`.
2. **KHU** — mỗi thư mục khu của mỗi bộ hồ sơ 1 dòng: `{"Khu": "Area 51 - Guestroom Lift Lobby", "Bộ hồ sơ": "IDD", "Mã phòng": "Sảnh thang L4-17", "Nguồn": [["Drawing Register p2", "F0xxx", 2]]}`.
   - Cũng lập dòng KHU cho tên khu của bộ FF&E / thảm / schedule (vd `"AREA51 thảm"` → `SK-01; ST-01…`) — **bộ khác nhau có thể đánh số Area khác nhau cho cùng khu**; chỉ gắn theo mã phòng.
   - Nguồn mã phòng: Drawing Register, trang bìa/khung tên tờ đầu của khu, FF&E Issue Sheet, Room schedule BOQ. Mở file trong zip: `py -I "<TOOL>" zip <GOC> <FID>` → đường dẫn file đã giải nén; đọc chữ `py -I "D:\COSTING EXCEL\SKILL\01-scan\tool\pd.py" chu "<pdf>" <trang>`.
   - Khu không gắn được → `Mã phòng` = `"KHÔNG THẤY – đã tìm ở …"` + VIEC_MO.
3. **Tên tờ** — từ Drawing Register (1 lần cho cả bộ), ghi `DWG.dwg.json` theo Số BV. Không mở từng tờ để lấy tên.
4. `Mã phòng` của DWG: chỉ ghi khi khác mã phòng của khu (vd tờ chung nhiều loại phòng); còn lại để build tra từ KHU qua `Khu (thư mục)`.
5. `Trạng thái`: Ngoài scope (khu bị loại theo TCQ — lấy từ `loc <GOC> TCQ_LOG "Trạng thái phạm vi=Loại trừ"` nếu bước 3 đã xong, hoặc từ README), Tham khảo, Chờ (khu chưa quyết, vd có bản vẽ nhưng không có BOQ). Mặc định tool = Hiệu lực.
6. Bản vẽ trong Drawing Register nhưng không có file → VIEC_MO; số BV có nhiều bản → kiểm bản hiệu lực đúng chưa (ngày issue trên khung tên nếu tên thư mục không rõ).

## Trước khi trả kết quả
- `check --den dwg`: mọi dòng KHU có `Mã phòng`.
- Trả về: số khu, số tờ theo khu, khu chưa gắn được mã phòng, tờ thiếu file.
