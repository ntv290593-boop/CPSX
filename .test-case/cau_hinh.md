# Cấu hình kiểm thử — COSTING EXCEL (phạm vi: skill 10-tender)

Khởi tạo: 2026-11-06   ·   Phích: `.test-case/phich.py`

## 1. Lệnh chạy workflow
| | |
|---|---|
| Chạy cả workflow | `/10-tender` (người xác nhận 2026-11-06) — KHÔNG chạy trong bộ test này |
| Chạy một trường hợp lẻ | gọi thẳng hàm trong `SKILL/10-tender/tool/td.py` qua phích |
| Xác nhận | người |

## 2. Dữ liệu test
| | |
|---|---|
| Dữ liệu thật | `<GOC>\10_TENDER\_data\` (dự án ACTIVE) |
| Chế độ | bộ test hiện tại KHÔNG mở dữ liệu thật; chỉ gọi hàm luật thuần |
| Người đã chọn | (A) dựng kho test riêng — CHƯA DỰNG; lệnh có ghi (build, goi, dmvt…) chưa được test |
| Xác nhận | người chọn A; chưa dựng kho nên phần ghi hoãn |

## 3. Đáp án chuẩn neo vào đâu
| | |
|---|---|
| Neo lớp chính | đáp án tay (`gia_dinh`) + `oracle_*` viết bằng thao tác chuỗi, không dùng regex/hàm của td.py |
| Neo dự phòng | `van_tay_ham.py`: vân tay logic từng hàm |
| Xác nhận | người duyệt 83 case đáp án tay (2026-11-06); oracle_* do skill viết |

## 4. Điểm quyết định (lớp luật, 7 điểm, ~0,1 giây)
| Điểm | Mã | Ở đâu | Đo nhánh |
|---|---|---|---|
| vai_tro_doan | TD.DOAN.VAITRO | td.py:509 | mã khai |
| loai_trang | TD.SPEC.LOAI_TRANG | td.py:1856 | mã khai |
| tham_chieu | TD.TCQ.THAM_CHIEU | td.py:1153 | MONG MANH (chưa khai nhanh) |
| date_in | TD.NGAY.DATE_IN | td.py:184 | mã khai |
| dot_of | TD.NGAY.DOT_OF | td.py:199 | mã khai |
| ma_vl | TD.VL.MA_VL | td.py:82 | mã khai |
| chuan_ben | TD.BEN.CHUAN_BEN | td.py:2409 | MONG MANH |

Lớp LLM: chưa khai (docsau, gemini, bước 7c–7e). Chưa có case.

## 5. Chạy lại: lớp luật không cần lặp. Lớp LLM chưa có.

## 7. Việc còn mở
- Dựng kho test để phủ lệnh ghi (build, goi, dmvt, boq).
- Thêm hàm luật khác của td.py (tcq_parse, khop_dong_boq, doi_khoa…).
