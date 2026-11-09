# Dự án đang active

- **Dự án:** Blanca City – Vũng Tàu, nội thất căn hộ DV2.3-DV2.4. Phục vụ **phòng TENDER** (skill `10-tender`, mẫu `cdt_item`, toàn gói).
- **Thư mục:** `_ACTIVE\Blanca_City` (hồ sơ ở `00_input`; làm việc ở `10_TENDER`)
- **Đợt hồ sơ:** `00_input±027_DV2.3-DV2.4` (đợt gốc 27/10, có BOQ CĐT `BOQ NỘI THẤT CĂN HỘ DV2.3-DV2.4 - MT.xlsx`) + `00_input\OneDrive_2026-11-06.zip` (CĐT cập nhật 06/11).
- `01_scan` là bản costing 27/10 bằng 01-scan, không dùng làm số liệu cho tender.
- **Cập nhật:** 2026-11-06

## Dự án tạm dừng: CTX (chưa đóng, không lưu trữ)
Muốn quay lại thì đưa khối này lên đầu file.
- **Dự án:** CTX (bản vẽ ID ghi "Sun Feliza Suites") – tổ hợp 5 tháp tại Cầu Giấy, Hà Nội; phạm vi BOHO các gói 3–6. Phục vụ **phòng TENDER**. Chưa có BOQ CĐT và TCQ, tự lập BOQ theo `Định hướng BOQ.xlsx`.
- **Thư mục:** `_ACTIVE\TCX` (hồ sơ ở `1. THÔNG TIN`; làm việc ở `10_TENDER`, mẫu tool `tu_lap_vn`)
- **Tên tháp:** Wealth = F1 = Tháp 1; Happiness = F2 = Tháp 2; Peace = F3 = Tháp 3; Health = F5 = Tháp 5; Longevity = F6 = Tháp 6. Thuyết minh đánh số khối tháp 1–5 theo trục, khối tháp 4 là Tháp 5.
- **Cập nhật:** 2026-11-01

## Dự án tạm dừng: SOFITEL (chưa đóng, không lưu trữ)
Muốn quay lại thì đưa khối này lên đầu file.
- **Dự án:** Sofitel Saigon Plaza Renovation – gói Main Fit-out (SSP-ID), phục vụ **phòng TENDER** (không phải costing)
- **Thư mục:** `_ACTIVE\SOFITEL`
- **Đợt hồ sơ:** `01. INPUT\new` (HSMT 06/10 + TCQ01–07 + ADD01, đến 26/10/2026). TCQ gộp đầy đủ: `new\10.261026_TCQ07\SSP-ID works-Tender queries TCQ.xlsx`.
- **BOQ tender:** `new\1.261006_HSMT\Editable file\07-08 SSP-ID package-Tender BOQ.xlsx`
- **Thư mục làm việc tender:** `_ACTIVE\SOFITEL\10_TENDER` (đã tạo 28/10, có README). Skill: `10-tender` (`SKILL\10-tender`). Phương án: `DOCS\PLAN_TENDER-SOFITEL_2026-10-28.md`
- **Cập nhật:** 2026-10-28

## Lưu ý
- Thư mục cũ `01. INPUT` (gốc: @DRAW, @SPEC, BOHO r0…) và `01_scan`, `03_costing`, `01.DOT 01`, `02.DOT 02`, file RE00 là của đợt costing trước. Không lấy số liệu từ đó cho tender trừ khi user cho phép.
- Phạm vi tender là toàn gói, KHÔNG áp quy tắc "phạm vi nhà máy 5 xưởng" của skill `01-scan`.

## Dùng chung (không thuộc dự án nào)
`_MASTER`, `_SCRIPTS`, `DOCS`, `01. DOOR` (bài học/ví dụ — chỉ tham khảo quy tắc, không lấy số liệu).

## Đã đóng / lưu trữ (KHÔNG đọc)
- `_ARCHIVE\NORTON` — đóng 27/10/2026.
- `_ARCHIVE\Blanca_City` — bản lưu trữ 28/10/2026 (không đọc). Dự án mở lại ở `_ACTIVE\Blanca_City` từ 06/11.
- `_ARCHIVE\SOFITEL` — bị chặn bằng permissions.deny (bản costing cũ). Thư mục `_TEST\SOFITEL_*` cũng không dùng.
- Memory về SOFITEL costing cũ (RE00, REV03, D15…), NORTON là của dự án đã đóng. Memory Blanca 28/10 là cũ, kiểm lại trước khi dùng.

## Khi đổi dự án
1. Chuyển dự án cũ sang `_ARCHIVE\<tên>`.
2. Sửa file này (dự án, file gốc, bản làm việc).
3. Đánh dấu memory của dự án cũ là đã đóng hoặc xóa.
