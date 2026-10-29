---
name: 01-scan
description: Quét hồ sơ thầu (BoQ, bản vẽ PDF, spec, material list, mail) trong 00_input\<đợt> của dự án ghi ở ACTIVE.md, ra 01_scan\PROJECT_DATA_<đợt>.xlsx theo form cố định. Dùng khi user nói "scan", "scan thư mục …", "quét hồ sơ", "làm PROJECT_DATA", "scan lại bước …".
---

# 01-scan — ĐIỀU PHỐI

Phiên này chỉ **điều phối**: chạy script, giao việc, đọc kết quả kiểm. **Không tự đọc BoQ / bản vẽ / spec ở đây** — mỗi bước giao 1 subagent mới (ngữ cảnh sạch, chỉ đọc file hướng dẫn của bước đó). Đó là cách chống ảo giác của skill này.

## 0. Chuẩn bị (không hỏi user nếu đã rõ)
- `GOC` = thư mục dự án trong `D:\COSTING EXCEL\ACTIVE.md`.
- `DOT` = thư mục user nêu trong `GOC\00_input\`. User không nêu → thư mục chưa có `01_scan\PROJECT_DATA_<DOT>.xlsx`; nhiều ứng viên → hỏi 1 câu.
- `TOOL` = `D:\COSTING EXCEL\SKILL\01-scan\tool\pd.py`, chạy `py -I "<TOOL>" <lệnh> …` (`py -I "<TOOL>"` không tham số = xem các lệnh).
- Dừng và báo user nếu: `~$PROJECT_DATA_<DOT>.xlsx` tồn tại (file đang mở); ổ D còn < 2 GB.

## 1. Các bước
| # | Bước | Hướng dẫn | Ai làm | Chạy sau |
|---|---|---|---|---|
| 1 | kiemke | `buoc/1-kiem-ke.md` | 1 subagent làm cả A (phạm vi) → B (lược bỏ) → C (quét); bạn không tự chạy `kiemke` | — |
| 2 | boq | `buoc/2-boq.md` | 1 subagent | 1 |
| 3 | spec | `buoc/3-spec.md` | 1 subagent | 1 (song song với 2) |
| 4 | banve | `buoc/4-ban-ve.md` | mỗi nhóm ≤ 8 mã 1 subagent, chạy song song; tên file `…banve-<nhóm>.json` | 2 |
| 5 | doichieu | `buoc/5-doi-chieu.md` | 1 subagent | 2, 3, 4 |

Prompt giao việc (thay `<…>`; không thêm số liệu hay suy đoán của bạn vào prompt):
```
Làm bước <buoc> của skill 01-scan.
Đọc rồi làm đúng theo: D:\COSTING EXCEL\SKILL\01-scan\buoc\<file>.md và D:\COSTING EXCEL\SKILL\01-scan\JSON.md
GOC = <…>   DOT = <…>   TOOL = D:\COSTING EXCEL\SKILL\01-scan\tool\pd.py
[bước 4: Nhóm <n>, chỉ làm các mã: <mã1, mã2…>. Ghi file *.banve-<n>.json, ID issue dạng V<n>-01]
Trả về ≤ 15 dòng: file đã ghi, số dòng, việc chưa làm được và lý do.
```

## 2. Sau mỗi bước
1. `py -I "<TOOL>" check <GOC> <DOT> --den <buoc>`.
2. Có LỖI → gửi nguyên danh sách lỗi cho subagent vừa làm (SendMessage) để sửa file JSON của nó. Tối đa 2 vòng; còn lỗi → ghi vào báo cáo, làm tiếp bước sau.
   Sau bước 4 (và bước 3 nếu có link SPEC/ML): chạy `py -I "<TOOL>" neo <GOC> <DOT>` → `_data\<DOT>\_neo.json`. Gửi các dòng `NEO LỆCH` của từng nhóm cho subagent nhóm đó (giống vòng sửa LỖI, tối đa 2 vòng, chạy lại `neo` sau mỗi vòng). Còn LỆCH thì để trong CẦN XEM, không chặn build. `check` tự in kết quả `_neo.json`; link đổi sau lần chạy `neo` thì `check` nhắc chạy lại.
3. Bước 2 xong: đọc số mã `Phạm vi nhà máy` TRONG SCOPE / KHÔNG THUỘC SCOPE / CHƯA RÕ. Có CHƯA RÕ → hỏi user 1 câu gộp (danh sách mã + lý do) **trước khi** giao bước 3–4; user trả lời thì ghi vào `QUAN_SAT.user.json` (Trường `Phạm vi nhà máy`). Mã KHÔNG THUỘC SCOPE: không giao bước 3, 4 (PROJECT_DATA vẫn giữ dòng).
4. Rồi mới chia nhóm cho bước 4 (theo cách user làm: **đọc BoQ → biết tính gì → tìm tên tờ → tìm chi tiết trong PDF**):
   - Chạy `py -I "<TOOL>" ungvien <GOC> <DOT>` → `_data\<DOT>\_ungvien.txt`. Mỗi mã có các nhóm: `CHI TIẾT` (xếp điểm), `THỐNG KÊ`, `MẶT BẰNG` (để đếm KL theo khu), `DANH MỤC`, `CỦA MÃ CHA` (phụ kiện/hạng mục con → trang của cửa/tủ cha, chỉ để tìm). Tool đã loại trang trùng, file tổng, file DỰ PHÒNG, tờ MEP/hoàn thiện. Muốn tốt hơn: bước kiemke ghi `_timkiem.kiemke.json`, bước boq ghi `_timkiem.boq.json` (xem `JSON.md`).
   - Lấy danh sách `Mã` trong `_ungvien.txt` (đã bỏ mã ngoài scope), nhóm các mã cùng loại/cùng bộ tờ vào một nhóm. Mã thuộc loại căn **không có PDF** (NOTES kiemke) → không giao; báo user cùng lúc với câu hỏi phạm vi. Trong prompt bước 4 ghi thêm: `Ứng viên: <dán đúng các khối của nhóm đó từ _ungvien.txt>` — subagent mở `CHI TIẾT` trước, `MẶT BẰNG` chỉ để đếm KL theo khu, `THỐNG KÊ` kèm với cửa; mã con cùng nhóm với mã cha. Subagent chỉ mở các trang này (rồi lần theo callout / Drawing List nếu thiếu), không quét cả bộ.

## 2b. Nguyên tắc lọc (giảm token — áp cho mọi bước)
**Chỉ quét hạng mục trong scope nhà máy** (5 xưởng: mộc, sơn đồ gỗ, cửa, cabinet, loose). Ngoài đó (sơn nước, thạch cao, sàn, gạch, đá, đèn…; vách kính và picomat vẫn TRONG SCOPE) → bước 2 ghi `KHÔNG THUỘC SCOPE NHÀ MÁY`, không đọc bản vẽ/spec (`buoc/2-boq.md` 4b). Khác với "Không bao gồm" khi tính giá chi tiết (vải, kim loại…) — bước scan không làm phần đó.
**Phạm vi trước, lược bỏ sau, quét cuối:** `kiemke --ds` (liệt kê gọn) → đọc mail + sheet tổng BoQ để hiểu phạm vi → model tự quyết file/thư mục nào BỎ (ghi `_bo_qua.txt`) → mới `kiemke` quét. Sau đó lọc dần tới chi tiết, lớp rẻ trước: trùng/phiên bản → Drawing List → mã từ BoQ → trang chứa mã (`ungvien`/`tim`) → crop khung tên → crop chi tiết. **PDF luôn ưu tiên; `.dwg` chỉ khi hạng mục không có PDF.** Loose Furniture nằm trong phạm vi. Chi tiết: `buoc/1-kiem-ke.md` (1b) và `buoc/4-ban-ve.md` (mục "Tiết kiệm token").

## 3. Kết thúc
- Bước boq và doichieu: chạy `boq-tach` / `scope` / `doichieu` trước (xem `buoc/2-boq.md`, `buoc/5-doi-chieu.md`) — tool làm phần chép/so máy làm được, LLM chỉ ghi phần mơ hồ. Thiếu dữ liệu gọn thì dùng `loc` thay cho đọc cả file JSON.
- `py -I "<TOOL>" build <GOC> <DOT>` → ghi `GOC\01_scan\PROJECT_DATA_<DOT>.xlsx` (bản cũ tự sao vào `01_scan\_backup`).
- Báo user, ngắn: đường dẫn file · số hạng mục · LỖI còn lại · danh sách **CẦN XEM** (chép từ output build) · câu hỏi ASK_USER. Không kể lại số liệu từng hạng mục — user xem trong file (sheet DASHBOARD).

## 4. Chạy lại
- Bước đã có file `*.<buoc>*.json` → bỏ qua, trừ khi user bảo chạy lại bước đó. Chạy lại = xóa các file JSON của bước đó rồi giao lại.
- **Không bao giờ xóa** `*.user.json`, `_snapshot.json` (chứa sửa tay của user; build tự giữ, user thắng).
- User sửa trực tiếp trong xlsx rồi bảo "build lại" → chỉ chạy `build`.

## 5. Không được
- Viết script build riêng cho dự án, hay ghi/sửa PROJECT_DATA bằng openpyxl/Excel. Mọi dữ liệu đi qua JSON + `pd.py build`.
- Sửa form (`form\PROJECT_DATA_FORM.xlsx`) để qua kiểm. Form chỉ đổi khi user yêu cầu.
- Lấy số liệu từ dự án khác (xem `D:\COSTING EXCEL\CLAUDE.md`).
