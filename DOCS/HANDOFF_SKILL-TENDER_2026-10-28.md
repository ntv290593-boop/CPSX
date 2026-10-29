# HANDOFF – Tạo skill TENDER (session sau) – 2026-10-28

Đọc kèm: `DOCS\PLAN_TENDER-SOFITEL_2026-10-28.md` (phương án chi tiết cho SOFITEL, dữ liệu mẫu đầu tiên).
Trạng thái: **ĐÃ TẠO skill `SKILL\10-tender` (28/10).** Tool `td.py` chạy thử SOFITEL xong (init/kiemke/boq/tcq/dwg/mat/doichieu/check/build ra scratchpad); các bước LLM (subagent) chưa chạy. Mục 7 chốt theo mặc định, chờ user xác nhận: thư mục `10-tender`; build kiểu 01-scan (JSON → tool); viết skill trước rồi chạy SOFITEL; Nhóm thực hiện = Nhà máy/NTP/MEP/Owner/Gói khác/Chung + cột `NTP chi tiết`; cột chính thức đặt sau cột cuối mỗi sheet BOQ.

## 1. Mục tiêu
Skill quét hồ sơ mời thầu cho **phòng tender** (không phải costing):
- Quét **toàn gói**, khoanh vùng bằng nhãn, không bỏ dòng.
- Theo dõi **thay đổi cuối cùng** qua các đợt TCQ/Addendum: vật liệu, bản vẽ, thông tin, phạm vi.
- Ra bản BOQ tender có cột bổ sung + các sổ đăng ký; bản gửi CĐT tách riêng.
- Dữ liệu thử đầu tiên: `_ACTIVE\SOFITEL\01. INPUT\new`.

## 2. Vì sao không dùng `01-scan`
| | 01-scan (costing) | Tender |
|---|---|---|
| Phạm vi | Chỉ 5 xưởng nhà máy (SKILL.md mục 2b) | Toàn gói, gắn nhãn |
| Đầu vào | `00_input\<đợt>` độc lập | Nhiều đợt nối tiếp, đợt sau đè đợt trước |
| Đầu ra | PROJECT_DATA form cố định | BOQ CĐT + cột + sổ TCQ_LOG/MAT/DWG/SCOPE |
| TCQ | Không có | Trọng tâm |

→ **Skill riêng** (user chọn hướng skill mới). Không thêm chế độ tender vào 01-scan để hai bộ quy tắc phạm vi không lẫn.

## 3. Dùng lại từ 01-scan
- Cách điều phối: mỗi bước 1 subagent ngữ cảnh sạch → ghi JSON → tool `check` → sửa tối đa 2 vòng.
- Tool `SKILL\01-scan\tool\pd.py`: `kiemke --ds`, `ungvien`, `tim`, crop khung tên/chi tiết, `boq-tach`. **Chưa thử với BOQ tender** (mẫu cột khác) → phải test trước khi dựa vào.
- Quy ước: PDF ưu tiên, DWG chỉ khi không có PDF; user sửa tay thì user thắng (kiểu `*.user.json`).
- Cấu trúc thư mục skill: `SKILL\<số>-<tên>\` (SKILL.md, JSON.md, buoc\, form\, tool\), junction vào `.claude\skills`.

## 4. Các bước đề xuất của skill
| # | Bước | Làm gì | Ra |
|---|---|---|---|
| 0 | chuẩn bị | GOC từ ACTIVE.md; thư mục input; dừng nếu có `~$` file đang mở | — |
| 1 | kiemke | Liệt kê các đợt theo thứ tự (tiền tố số + yymmdd); **liệt kê nội dung zip, không giải nén hết**; gán vai trò file: chính / bổ sung-thay thế / làm rõ / tham khảo / ngoài scope / chưa có (link chưa tải); bỏ `~$…`, `plot.log` | FILE_INDEX |
| 2 | boq | Bản đồ BOQ: sheet, vùng (cột B), nhóm (cột C), dòng hạng mục (A là số), cột theo mã phòng (hàng 6–8), Base T–W, VE AL–AP, ghi chú AJ/AP. Nên làm bằng tool | BOQ_ITEMS |
| 3 | tcq | Đọc file TCQ **gộp mới nhất** (SOFITEL: TCQ07 có No1–7); mỗi câu 1 dòng; dịch VN (dùng lại takenote nội bộ nếu có); phân loại: loại thay đổi, mã VL, dòng BOQ, "Refer TCQ #n", **bị thay bởi**; câu chưa trả lời / deferred → Việc còn mở; link → "chưa có" | TCQ_LOG |
| 4 | dwg | Số bản vẽ từ tên file (vd `ID-SK-53-602`) + ngày thư mục issue + đợt; bản mới đè bản cũ | DWG register |
| 5 | mat | Mã VL từ Finishes Schedule / FF&E + thay đổi từ TCQ (trạng thái cuối) | MAT register |
| 6 | banve | Theo **mã phòng/vùng**: tìm trang cho từng dòng BOQ, trích gọn ra `02_extract\<mã>.md` | extract |
| 7 | doichieu | Gắn mỗi dòng BOQ: nhóm thực hiện, trạng thái phạm vi, mã TCQ, bản vẽ hiệu lực, việc còn mở | SCOPE + cột B |
| 8 | build | Chép BOQ CĐT → thêm 3 cột chính thức + cột nội bộ + sheet sổ; xuất bản gửi CĐT (chỉ 3 cột chính thức) | BOQ_TRACK, 04_out |

## 5. Đầu ra & cột (theo PLAN)
- **Cột chính thức** (SOFITEL: được chấp thuận ở TCQ04-19): Mô tả NT – Base · Mô tả NT – Đề xuất · KL – Đề xuất.
- **Cột nội bộ:** Mã VL · Bản vẽ hiệu lực · Rev/ngày/đợt · Mã spec/FF&E · Nhóm thực hiện (Nhà máy / NTP / MEP / Owner) · Trạng thái (Base / VE / Loại trừ / Omit / Chờ) · Mã TCQ · Nội dung cuối (VN) · Loại thay đổi (KL / VL / Bản vẽ / Phạm vi) · Việc còn mở · Ngày cập nhật.
- **Sổ trong cùng workbook:** TCQ_LOG, MAT, DWG, SCOPE; cột nội bộ tra từ sổ.
- **Thư mục dự án:**
```
<GOC>\10_TENDER\
  README.md                 ← đợt hiệu lực, quy ước, thứ tự ưu tiên, bẫy
  01_index\FILE_INDEX.xlsx
  02_extract\<mã phòng/vùng>.md
  03_boq\<DA>_BOQ_TRACK_<yymmdd>.xlsx
  04_out\
  _data\                    ← JSON trung gian của các bước
```

## 6. Quy tắc cố định cho skill
- Input chỉ đọc. Mọi số liệu nêu nguồn (file / sheet / ô / số bản vẽ); không có → "chưa có".
- Thứ tự ưu tiên khi mâu thuẫn: theo hồ sơ dự án (SOFITEL: **Spec → Bản vẽ → BOQ**, TCQ04-3) — ghi vào README dự án, không cứng trong skill.
- Gắn theo **mã phòng**, không theo số Area (SOFITEL: số Area IDD ≠ FF&E/thảm).
- Câu TCQ sau đè câu trước phải ghi chuỗi "bị thay bởi" (vd 50-501: TCQ04-22 → TCQ05-24 Omit; kính cửa lùa TCQ04-21 → TCQ05-35).
- Thay đổi ở cấp **mã VL** ghi ở MAT, không nhồi vào từng dòng BOQ.
- Không tự tải link ngoài (OneDrive/Drive); ghi "chưa có".
- Hạng mục có bản vẽ nhưng không có trong BOQ (SOFITEL: Main Lobby) → ghi ở Việc còn mở, không tự lập BOQ khi user chưa quyết.

## 7. Câu hỏi mở cho session sau
1. Tên/số thư mục skill (vd `SKILL\0x-tender`)? Tên gọi kích hoạt ("scan tender", "quét HSMT"…)?
2. Build: theo kiểu 01-scan (JSON → tool build, cấm ghi tay bằng openpyxl) hay cho phép script trực tiếp?
3. Làm thủ công lần đầu với SOFITEL rồi đóng gói, hay viết skill trước rồi chạy thử SOFITEL?
4. Danh mục "Nhóm thực hiện" có cần chia NTP chi tiết (đá, thảm, rèm, kính, inox, chống thấm, xây, phá dỡ)?
5. Vị trí cột chính thức trong BOQ (sau AP?) — kiểm form CĐT có ràng buộc không.

## 8. Việc ghi nhận, chưa xử lý (SOFITEL)
- Main Lobby: có bản vẽ ADD01, không có sheet BOQ → kiểm sau.
- Link TCQ không tải: spec rèm cập nhật (TCQ06-21), as-built Main Lobby (TCQ06-16), thảm hành lang (TCQ04-24), đèn minibar (TCQ05-30).
- Câu còn mở: TCQ04-6 (E1/E2), TCQ06-4 (biển tầng, deferred), TCQ06-6 (ST-07 to be confirmed).
