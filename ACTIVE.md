# Dự án đang active

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
- `_ARCHIVE\Blanca_City` — chuyển archive 28/10/2026.
- `_ARCHIVE\SOFITEL` — bị chặn bằng permissions.deny (bản costing cũ). Thư mục `_TEST\SOFITEL_*` cũng không dùng.
- Memory về SOFITEL costing cũ (RE00, REV03, D15…), NORTON, Blanca là của dự án đã đóng.

## Khi đổi dự án
1. Chuyển dự án cũ sang `_ARCHIVE\<tên>`.
2. Sửa file này (dự án, file gốc, bản làm việc).
3. Đánh dấu memory của dự án cũ là đã đóng hoặc xóa.
