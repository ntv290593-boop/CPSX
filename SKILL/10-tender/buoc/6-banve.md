# Bước 6 — banve: bản vẽ hiệu lực cho từng dòng BOQ, trích gọn theo mã phòng

**Ghi:** `SCOPE.banve-<n>.json` (`Bản vẽ hiệu lực`, `Mã VL`, `Mã spec/FF&E`, `Nguồn`), `02_extract\<mã phòng hoặc khu>.md`, `VIEC_MO.banve-<n>.json`.
**Chỉ làm các mã phòng/khu được giao.** Không bóc khối lượng chi tiết, không tính giá — mục tiêu: mỗi dòng BOQ biết **tờ nào là căn cứ** và **vật liệu nào** theo bản vẽ hiệu lực.

## Làm
1. Đọc README + mục bẫy. Tờ của khu: `loc <GOC> DWG "Khu (thư mục)~<khu>" "Trạng thái=Hiệu lực" --cot "Số BV,Tên tờ,File PDF"`; khu ↔ mã phòng: `loc <GOC> KHU`.
2. Dòng BOQ của mã phòng: `loc <GOC> SCOPE "Mã phòng~<mã>:" --cot "Khóa,Nhóm,Mô tả BOQ,Mã VL,Mã spec/FF&E" --max 300`. Dòng khu chung (hành lang, sảnh thang, cabin) lọc theo nhãn cột tương ứng.
3. **Gán tờ theo tên tờ trước** (mặt bằng sàn / trần / mặt đứng / chi tiết / FF&E) — không mở PDF nếu tên tờ đã đủ cho dòng đó. Chỉ mở PDF khi cần mã VL/FF&E hoặc khi tên tờ không đủ. File PDF trong zip: `py -I "<TOOL>" zip <GOC> <FID>`; đọc chữ `py -I "D:\COSTING EXCEL\SKILL\01-scan\tool\pd.py" chu "<pdf>" 1`; ảnh khi chữ là nét vector: `… pd.py anh "<pdf>" 1 "<scratchpad>\x.png" --dpi 100 [--clip …]` rồi Read PNG. Ảnh rẻ trước (100 dpi, crop), không xuất cả trang dpi cao.
4. Mỗi dòng SCOPE ghi:
   - `Bản vẽ hiệu lực`: list link, **nhãn bắt đầu bằng Số BV** (build dùng nó tra ngày issue/đợt): `[["ID-SK-52-301 mặt đứng A", "F0123", 1], ["ID-SK-52-101 MB sàn", "F0119", 1]]`.
   - `Mã VL` / `Mã spec/FF&E`: theo nhãn trên bản vẽ hiệu lực (ghi đè nháp regex của tool khi khác). Bản vẽ ≠ mô tả BOQ → ghi theo bản vẽ + VIEC_MO nêu lệch (thứ tự ưu tiên trong README quyết định bên nào thắng khi chào giá; bạn không tự quyết).
   - Không thấy tờ cho dòng → `Bản vẽ hiệu lực` = `"KHÔNG THẤY – đã tìm ở <khu/tờ>"`.
   - Gán hàng loạt bằng `_loc` (vd mọi dòng `Nhóm~DEMOLITION` của SK-01 → tờ phá dỡ), rồi dòng cụ thể cho ngoại lệ.
5. `02_extract\<mã>.md` — trích gọn để người khác đọc thay bản vẽ (≤ 80 dòng): bộ tờ hiệu lực (Số BV – tên – ngày), vật liệu theo khu vực (sàn/tường/trần/WC), đồ built-in & FF&E chính (mã), điểm khác BOQ, điểm khác bản cũ/TCQ. Mỗi ý có Số BV hoặc mã TCQ.
6. Hạng mục thấy trên bản vẽ mà không có dòng BOQ, hoặc dòng BOQ không thấy trên bản vẽ → VIEC_MO (Liên quan = mã phòng + Số BV/khóa).

## Trước khi trả kết quả
- Mọi dòng của nhóm có `Bản vẽ hiệu lực` (link hoặc "KHÔNG THẤY – …"); `check --den banve` không còn LỖI ở file của bạn.
- Trả về: số dòng đã gán, số dòng KHÔNG THẤY, file extract đã ghi, lệch BOQ ↔ bản vẽ chính.
