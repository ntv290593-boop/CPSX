# Handoff: CTX bổ sung phần sót sau khi rà độ phủ bản vẽ, 02/11/2026

Viết cho phiên Claude tiếp theo. Đọc thêm `ACTIVE.md`, `_ACTIVE\OneDrive_2026-11-01\10_TENDER\README.md`, mục 0 của `DOCS\HANDOFF_CTX-BOQ-TU-LAP_2026-11-02.md`.

`<GOC>` = `D:\COSTING EXCEL\_ACTIVE\OneDrive_2026-11-01`. Bộ ID = `1. THÔNG TIN\2.261030_HSTK ID\260731_100% Construction Documentation Drawings`. Bộ kiến trúc = `1. THÔNG TIN\1.261009_HSMT\260630 TK BVTC full Kien Truc\01 Ban ve`.

## 1. Việc cần làm: 3 quyết định user chốt 02/11 (chưa làm, BOQ chưa đổi)
User trả lời "1. có. 2. có. 3. có." cho 3 câu dưới. Câu 2 và 3 **ngược đề xuất của tôi**: tôi đề xuất không tính, user chọn tính. Không hỏi lại.

### Câu 1: Sửa 3 chỗ sót trong bộ ID
| Thiếu | Nội dung | Nguyên nhân | Cách sửa dự kiến |
|---|---|---|---|
| Tờ 5300.1–5302.2, Peace sảnh chính | Drawing List: RECEPTION COUNTER (5300), RECEPTION FEATURE WALL (5301), FEATURE PARTITION (5302). Tag trong DWG: WD 05 ×45, MT 01 ×16, RATTAN SP 02 ×7, SP 02 ×1, GL 05 ×5 | File tên `AID2501_CD_PE_L1_ML_5300-5302.dwg` thiếu số, không khớp `rx_ten_dwg` trong `_data\_lap_boq.json` (`AID250121_...`) | Đổi `rx_ten_dwg` nhận cả `AID2501_`. DWG đã đọc sẵn trong `_dwg_tag.json`, không cần đọc lại |
| Tờ 5525, Longevity Gym | JOINERY 6 - WOOD PANEL, WD-01, ngang 2715, cao 4000 (xem hình PDF `5000 Joinery\L6\AID250121_CD_LO_L6_GY_5525.pdf`) | Chỉ có PDF, không có DWG nên không có tag | Thêm cấu hình cho tờ chỉ có PDF trong `_lap_boq.json` (không viết script riêng), dòng ở `Longevity — Gym L6`, số lượng chưa đếm, có cờ. Kiểm K1 của `kiem_boq.py`: tờ này không có tag |
| Cửa D19, Store Door tầng 6 | Cửa gỗ, WD-01, HD-11. Lịch tờ 8004, chi tiết 8112. Ghi chú lịch không ghi bộ HD, HD-11 lấy từ tag | Không có tag trên mặt bằng nào nên `dem` không có mã này | Thêm cấu hình cửa không có tag, sheet `Podium 1-7`. User chốt: ghi **"chưa rõ khu, chưa có số lượng"**, cờ hỏi thiết kế, **không tự gán khu**. SL trống nên không cộng vào `HW.DO` |

### Câu 2: Tính DF30, DF30.1 vào gói 4
Câu hỏi đã duyệt: "DF30, DF30.1 (cửa gỗ chống cháy phòng kỹ thuật) có thuộc gói 4 không?" → **có**.

- Nguồn: `10. MAT BANG DINH VI CUA\CT CUA DIEN HINH\PDF\10-003-CT CUA DI.pdf`.
- DF30: phòng kỹ thuật, phòng trực PCCC, kho thiết bị PCCC, cửa gỗ chống cháy, khuôn 1010×2280, EI 60.
- DF30.1: phòng bơm PCCC tầng 26, cửa gỗ chống cháy, khuôn 1010×2280, EI 70.
- Vật liệu hoàn thiện: **chưa có**, tra `05 VLHT\CTX-SPEC-VLHT-8-260630 (Rv1).xlsx`, không suy.
- Đếm trên `10. MAT BANG DINH VI CUA\<F1, F2, F3, F5, F6, PHAN DE>\PDF`. Mã nằm trong lớp chữ PDF.
- **Bẫy đếm:** tờ tổng và tờ phân khu cùng tầng lặp nhau (`10-110` tầng 1 có 14 DF30, `10-111`…`10-115` là các phần của chính tầng đó), không cộng cả hai. Một tờ có thể áp nhiều tầng (`T7-24`). Số xuất hiện đã thống kê ở lần rà (F1 37, F2 36, F3 33, F5 34, F6 33, PHAN DE 136, PHAN NGAM 125) **chưa phải số lượng**.
- **Hầm:** user đã chốt sheet `Phần hầm` để trống từ trước; câu trả lời "có" không nhắc tới hầm. Mặc định không đưa `PHAN NGAM` vào dòng, ghi vào `VIEC_MO` và nói rõ khi báo.
- `bo_thu_muc` trong `_data\_cauhinh.json` đang loại cả `01 Ban ve`. Muốn có link nguồn phải đưa các PDF cần dùng vào FILE_INDEX, thu hẹp lại, đừng mở cả 2.067 file. Cập nhật README mục "Thư mục bỏ khỏi quét" và dòng ghi bộ kiến trúc là "Tham khảo".
- Thứ tự ưu tiên trong README: "hồ sơ ID thắng hồ sơ kiến trúc". DF30 không có ở ID nên không xung đột, ghi rõ nguồn là kiến trúc trên từng dòng.

### Câu 3: Tính LM-01 vào ốp gỗ
Câu hỏi đã duyệt: "LM-01 có tính vào ốp gỗ không?" → **có**.

- LM-01: Laminate Suede Nordic Acacia, khu Changing Room (Male & Female), `FM_7100` tr.42. Tờ có tag: 3418, 3435, 3436, 3437, 3438, 3445.
- Thêm vào `vat_lieu.ma` trong `_lap_boq.json`: `"LM-01": {"tuong": "WF.WO", "tran": "CF.WO"}`.
- Bề mặt do `be_mat.py` quyết định. Tag chỉ vào đồ gỗ hoặc cửa sẽ không ra dòng ốp, phải kiểm lại sau khi chạy. Tôi chưa xem hình các tờ này.

## 2. Chạy lại
1. Kiểm `~$*.xlsx` trong `10_TENDER\03_boq`: có thì user đang mở file, dừng và báo. Chiều 02/11 đã gặp lock này, mất đi sau vài phút.
2. `lap_boq.py <GOC>` → `td.py kiemke`, `boq`, `dwg`, `mat` → `lap_boq_buoc.py <GOC>` → `td.py phanloai` → `td.py check` → `kiem_boq.py <GOC>` → `td.py dmvt <GOC> --chi-boc` → dịch dòng mới vào `_data\_dmvt_llm.json` (chỉ các khóa chưa có) → `td.py build <GOC>`.
   - Chạy bằng `py -I "D:\COSTING EXCEL\SKILL\10-tender\tool\<file>"`.
   - Script đọc dữ liệu riêng: `_MASTER\tools\.venv\Scripts\python.exe -I`, để ở scratchpad.
   - Lệnh nhiều dòng trong Bash từng lỗi quote: ghi script ra file rồi chạy.
3. Mốc so sánh trước khi sửa (build 15:57): 389 dòng BOQ, 54 dòng cửa, 56 `HW.DO`, 197 dòng DMVT, 193 ảnh, `check` 0 LỖI 1 XEM, `kiem_boq` K2 15, K3 1, K8 CG-4001.
4. Trước khi báo: đọc lại `BOQ_TONGHOP` các dòng mới (5300–5302, 5525, D19, DF30, DF30.1, LM-01), xem ảnh DMVT mã mới, build 2 lần so nội dung. Phân biệt rõ kiểm cấu trúc với kiểm nội dung, nói phần nào chưa xem.
5. Đổi tính năng skill: cập nhật `SKILL\10-tender\SKILL.md`, `DOCS\Flowchart\10-tender_flowchart.md` phần 9, thêm 1 dòng `DOCS\Flowchart\CHANGELOG.md`.

## 3. Đã làm trong phiên 02/11 (chưa commit, user chưa yêu cầu)
- `lap_boq.py`: `vl_cua_tu_tag`, `vl_cua_cua_ma`, `ma_tag_cua`, `ma_cua_tieu_de`. Dòng cửa ghi `Vật liệu: … Phụ kiện: …`. Vật liệu = tag theo ô tờ lịch 8000–8008 + mặt cắt tờ 8100–8117 (bỏ tag mặt bằng chi tiết vì chỉ cả tường hai bên). HD = ghi chú lịch, không có thì tag mặt đứng.
- `td.py ma_cua_dong`: link theo thứ tự mã trong mô tả.
- `_lap_boq.json`: `cua_ma_khung` (D37 bỏ WD-05 khỏi link, user chốt; tag đó chỉ vào khung đầu cánh).
- `_dmvt_llm.json`: dịch HD-12, 13, 14, 16, 22.
- Đã sửa cả `SKILL.md`, flowchart, CHANGELOG, handoff cũ.

## 4. Vẫn chờ user (chưa trả lời, không tự quyết)
- **D21, D22, D23, D25, D26, D27, D43** có còn là cửa gỗ không. Tag chỉ có MT-01 + GL-12 (D43 chỉ MT-05). Lịch ghi TIMBER DOOR, spec HD-12, HD-13 ghi Glass Doors. Nếu là kính thì thuộc `AL.GL`, không được tick.
- **Lệch mã gỗ giữa tờ lịch và mặt cắt** (tag chỉ vào cánh, đang giữ cả 2 mã + cờ): D6a, D6b (WD-05, WD-04, tờ 8100 dùng chung nhiều cửa); D31 (WD-01, WD-03, tờ 8107, WD-03 là gỗ sàn theo spec); D42 (WD-06, WD-04, tờ 8111).
- **Hỏi thiết kế:**
  - D37: khung ghi WD-0x, WD-08, WD-05.
  - D25: HD-14 ở tờ lịch, HD-13 ở tờ 8115 và spec.
  - HD-22: spec ghi Garden Café, tag ở D20.
  - Mã chưa chốt WD-XX (D16, D17, D32), HD-0X (D42).
- **DF40, DF41 trùng D40–D42 bên ID, không thêm dòng:**
  - Căn cứ: F2 T7-25 có 13 DF40 = 1 D40 + 12 D42.
  - Rộng lệch: DF40 1110, DF41 1480; D42 và D41 1480, D40 1550. Cần hỏi thiết kế, ảnh hưởng giá.
  - Số cửa đổi theo tầng (F2: 13 ở T7–25, 9 ở T28–38, 5 ở T40) trong khi BOQ tính cho 1 tầng điển hình. Bộ `10. MAT BANG DINH VI CUA` dùng được khi tính khối lượng.
- **Chưa kiểm:** DW06, DW07, DW09 (tầng 1) có trùng cửa WC sảnh bên ID không. D30: mặt bằng 8107 ghi WD-02, mặt cắt và lịch WD-01, không có cờ vì máy không đọc tag mặt bằng chi tiết. Các tờ cửa chưa xem hình: ngoài 8000, 8007, 8100, 8107, 8111, 8112, 8117.
- **Đề xuất đã nêu, user chưa trả lời riêng:** DF40, DF41 không thêm; DW01, DW41–DW45 ngoài 6 khu của Kế hoạch nên không thêm. Câu 2 chỉ chốt DF30, DF30.1.
- **Từ trước:** SP-18 (cột Ballroom) chưa rõ vật liệu; tờ 5303 có dòng nhưng không tên, không PDF; FF.WF thuộc gói nào; DXF 12 GB chờ duyệt xóa (`HANDOFF_CTX-BOQ-TU-LAP` mục 4 bước 8).

## 5. Kết quả rà độ phủ 02/11 (2.891 file trong `1. THÔNG TIN`)
529 file trong FILE_INDEX; loại 2.067 (`01 Ban ve` kiến trúc, 1.041 PDF + 1.026 DWG), 242 (`CAD` ID, đọc riêng bằng `dwg_tag.py`), 53 (`260830-BVTC de moi thau`, đã so: trùng tên và dung lượng với ID).

- **Đã kiểm, không sót:**
  - `AID250121_CD_LO_L6_BA_3510 - 3514.DWG` đã đọc, chỉ khác đuôi hoa.
  - `PHẠM VI CÔNG VIỆC.xlsx` ở thư mục ID cùng 21 ô tick với bản BOHO, chỉ không có cột gói.
  - 4000 Details không có mã gỗ, gương mới.
  - Kiến trúc: 03, 04, 07, 11, 12, 15 không có từ khóa nào; 14 (thang) chỉ có mã DF30; 19 chỉ có gương cầu lồi, rèm ngăn cháy; 09 trần "vân gỗ" AL-05 là nhôm; 18 có gương, lavabo ở hầm 1 (3 tờ, tờ 18-003 chưa rõ khu).
- **Giới hạn:** kiến trúc mới dò từ khóa trên lớp chữ PDF, chưa xem hình, chưa đọc DWG. `AID250121_CD_PR_7000` (Preamble, 265 trang) chưa đọc, chỉ xem trang đầu. Không được nói "hết sót".
- **Thứ tự phạm vi đã chốt:** ô tick `PHẠM VI CÔNG VIỆC BOHO.xlsx` thắng Kế hoạch. Kế hoạch gói 4 ghi "Cửa gỗ chống cháy (không gồm cửa chống cháy kim loại)", 6 khu: sảnh lobby, sảnh hành lang (không gồm thang bộ), sảnh thang máy, cabin, khu public T1–T6, WC công cộng. Gói 2 là BOH: hầm, phòng kỹ thuật, thang bộ, logia, fitout văn phòng; câu 2 là user chủ động tính DF30 vào gói 4 dù khác Kế hoạch, ghi vào sheet `Ghi chú & ngoài phạm vi` và VIEC_MO.

## 6. Bài học
Cả 2 lỗi đều im lặng: `check` và `kiem_boq` báo sạch trong khi tool bỏ qua file tên lệch và tờ chỉ có PDF. Đề xuất (chưa làm, chưa ai yêu cầu): thêm 1 kiểm vào `kiem_boq.py`: mọi tờ joinery trong Drawing List phải có dòng BOQ hoặc lý do; mọi cửa trong lịch phải có dòng BOQ hoặc lý do; báo DWG không khớp `rx_ten_dwg`.

Ở thư mục gốc `<GOC>` có 2 file `CTX_BOQ_TRACK_261102.xlsx`, `CTX_DMVT_261102.xlsx` ngoài `10_TENDER\03_boq`. Chưa rõ của ai, chưa mở; đừng ghi đè, hỏi user nếu cần.
