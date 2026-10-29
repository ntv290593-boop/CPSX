# Kế hoạch: đọc bản vẽ đáng tin cậy cho bước scan và takeoff

Ngày lập: 2026-10-24 · Trạng thái: **đề xuất, chờ duyệt từng giai đoạn** · Người dùng duyệt plan trước khi làm (HANDOFF §0).

> **Đã được thay bằng `KE-HOACH-TONG-THE_v2.md`** (sửa theo phản biện: SOFITEL có DWG, đọc ≥200 dpi theo R0167, quy tắc quyết định, nguồn đáp án mới). Giữ bản này để đối chiếu.

Tài liệu liên quan trong thư mục này: `PILOT_SOFITEL_2026-10-24.md` (kết quả thử nghiệm).

---

## 1. Vấn đề cần giải quyết

| # | Vấn đề | Hậu quả |
|---|---|---|
| 1 | Bước scan phải quét rất nhiều dữ liệu nên dễ sót | Thiếu mã cửa, thiếu chi tiết, không ai biết là thiếu |
| 2 | Bước takeoff không có cách kiểm tư duy | Một lần điền sai có thể làm sai cả dự án |
| 3 | Model nhìn hình có thể ảo giác hoặc bỏ sót | Số kích thước, số lượng, mã vật liệu sai mà vẫn trông hợp lý |
| 4 | Dùng token vào việc script làm được | Chậm, tốn kém, không lặp lại được |
| 5 | Khả năng đọc bản vẽ là rủi ro chính | Mọi bước sau đều thừa hưởng lỗi đọc |

**Tiêu chí thành công:** mỗi giá trị đi vào CPSX phải (a) có nguồn (file, trang, vị trí, ảnh crop), (b) qua ít nhất một kiểm tra độc lập, (c) nếu không chắc thì dừng ở ASK_USER/RFI, không điền.

## 2. Hiện trạng (đã xác minh)

- Pipeline cũ: `_MASTER\tools\scan_pages.py`, `scan\scan_schedules.py`, `scan\build_project_data_NORTON.py`, `takeoff\door_phieu_NORTON.py`. Chỉ đọc text layer. `door_phieu_NORTON.py` gõ cứng giá trị đã duyệt.
- NORTON: các tờ sheet có text layer chứa số kích thước (ví dụ `41(±1)`, `210`). Đã có sẵn PDF theo sheet (`03_BanVe\_sheets`) và DWG theo sheet (`DWG_GEN-BASEMENT`...).
- SOFITEL: 271 tờ PDF một trang (`01_scan\_sheets`). **Số kích thước là đường vector, không có trong text layer** (đã đếm: 15–40 nghìn nét/tờ, chỉ có số hiệu chi tiết, tỉ lệ, FCL/FFL là text).
- Công cụ đã cài: `agy` 1.4.6 (Gemini 3.8 Flash), pymupdf 2.0, ezdxf 1.4.5 (trong `_MASTER\tools\.venv`). Chưa có ODA File Converter (không có trong winget), nên chưa đọc được DWG.
- Pilot Sofitel (13 dòng THIẾU): xem `PILOT_SOFITEL_2026-10-24.md`. Tóm tắt: Gemini tốt ở ghi chú và tổng, kém ở chuỗi dài và đếm; phép cộng chuỗi kích thước bắt được lỗi của cả Gemini lẫn bản vẽ; session trước có hai kết luận sai (số lượng cột 7 là vòng tròn suy từ BoQ; len kim loại 80 thay vì 40).

## 3. Nguyên tắc thiết kế

1. **Script trích, model chỉ diễn giải.** Mọi thứ máy làm chắc chắn được (đếm trang, đọc text, cộng chuỗi, so mã, tra bảng) thì do script làm. Model chỉ nhận crop nhỏ và một câu hỏi hẹp.
2. **Hai người đọc độc lập, bất đồng thì hỏi người.** Claude và Gemini đọc cùng crop, riêng rẽ. Khớp thì nhận, lệch thì đánh dấu.
3. **Mọi giá trị có nguồn.** Bộ nguồn gồm file, trang, bbox, đường dẫn ảnh crop, người đọc (script/Gemini/Claude), độ tin cậy.
4. **Kiểm bất biến bằng code.** Các quy tắc cứng (mục 6) chạy tự động trước khi ghi vào CPSX.
5. **Không đoán.** Thiếu nguồn thì ASK_USER hoặc RFI. Không suy số lượng từ BoQ rồi gọi là "khớp".
6. **Sổ độ phủ.** Mỗi tờ, mỗi trang có trạng thái; cuối bước scan script phải báo được "đã xử lý X/Y".
7. **Không ghi đè.** Mỗi lần ghi lớn tạo phiên bản mới (v9, v10…). Không ghi trực tiếp .xlsm (HANDOFF §2).

## 4. Kiến trúc: bốn tầng đọc, xếp theo độ tin cậy

| Tầng | Cách | Độ tin cậy | Áp dụng |
|---|---|---|---|
| A | DWG → DXF (ODA File Converter) → `ezdxf`: TEXT, MTEXT, DIMENSION (giá trị đo thật), block attribute, layer | Cao nhất, không ảo giác | NORTON có DWG theo sheet; SOFITEL cần kiểm tra DWG có không |
| B | PDF vector (PyMuPDF `get_drawings`) + tỉ lệ từ khung tên: đo hình học | Cao, kiểm chứng được | Khi cần đo độ dày tường, bề rộng chi tiết |
| C | Text layer + bảng (pdfplumber, pymupdf words) | Cao khi có số | NORTON; SOFITEL chỉ lấy số hiệu, tỉ lệ, nhãn |
| D | Model nhìn ảnh crop (Claude + Gemini), kèm kiểm bất biến | Trung bình, phải kiểm | Mọi thứ còn lại: chi tiết cấu tạo, tag trong ảnh, kích thước vector |

Quy tắc chọn tầng: cố dùng A→B→C trước; chỉ dùng D cho phần chưa có. Kết quả D luôn đi qua kiểm bất biến và đọc kép.

## 5. Quy trình đề xuất (theo bước dự án)

### 5.1 Scan

1. **Kiểm kê (inventory):** liệt kê toàn bộ file, số trang, loại (bản vẽ, spec, BoQ), tờ đã tách. Ghi vào `LEDGER` (mục 7).
2. **Phân loại trang bằng script:** đọc khung tên (đã có trong `scan_pages.py`), gán số BV, tên, phiên bản; trang không đọc được khung tên bị gắn cảnh báo.
3. **Chọn tờ liên quan:** từ BoQ (mã cửa, mã hoàn thiện, hạng mục) tìm tờ ứng viên bằng text layer, DETAILS, `_LINKS`.
4. **Trích tầng A/B/C** tự động; ghi giá trị kèm nguồn.
5. **Phần còn lại sang tầng D:** cắt crop theo chi tiết (dựa vào vị trí tiêu đề chi tiết trong text layer), render 110–200 dpi; Claude và Gemini đọc riêng, theo schema JSON cố định.
6. **Đối soát:** so hai lần đọc, chạy kiểm bất biến (mục 6), tính độ tin cậy, đánh dấu lệch.
7. **Báo cáo độ phủ:** mã cửa trong thống kê vs BoQ vs mặt bằng; tờ chưa xử lý; dòng BoQ chưa có nguồn.
8. **Ghi PROJECT_DATA phiên bản mới** theo cấu trúc NORTON v8 và quy tắc link tự trỏ + `_LINKS` (xem memory `project-data-format`).

### 5.2 Takeoff

1. Đầu vào là dữ liệu đã trích có nguồn (không đọc lại bản vẽ bằng tay).
2. **Phiếu bóc sinh từ dữ liệu**, không gõ cứng như `door_phieu_NORTON.py`.
3. **Kiểm bất biến takeoff** (mục 6.2) chạy tự động; vi phạm thì không xuất AI_INPUT.
4. Mọi chỗ lệch hồ sơ (Spec ≠ BoQ ≠ BV) sinh Issue → RFI theo `bai-hoc\05_lech-ho-so-rfi.md`.
5. Ghi `02_takeoff\AI_INPUT.xlsx` (Power Query nạp vào CPSX), không ghi .xlsm trực tiếp.

## 6. Kiểm bất biến (chạy bằng code)

### 6.1 Ở bước scan

| Kiểm tra | Cách làm |
|---|---|
| Chuỗi kích thước phải cộng ra tổng | Tổng các đoạn = số tổng ghi trên bản vẽ. Lệch thì báo "đọc sai" hoặc "bản vẽ tự mâu thuẫn" |
| Đếm phải có nguồn đếm | Số lượng đơn vị chỉ nhận từ mặt bằng/mặt đứng/bảng thống kê. Không suy từ BoQ chia kích thước |
| Mã trên BV phải có trong Finishes Schedule | Mã lạ → cảnh báo |
| Mã cửa BoQ ↔ thống kê ↔ mặt bằng | Khác tập hợp → báo thiếu/thừa |
| Hai người đọc khớp nhau | Lệch → chuyển người quyết |
| Giá trị có ảnh crop và vị trí | Thiếu nguồn → không ghi |

### 6.2 Ở bước takeoff (chuyển từ rule R0166–R0179 và memory)

| Kiểm tra | Nguồn rule |
|---|---|
| Mỗi mã cửa có độ dày tường kèm tờ mặt bằng nguồn | R0166, memory `door-wall-thickness-check` |
| Cửa EI: mọi cấu phần (cánh, khung bao, rebate) có ≥1 lớp ERON, trừ "chỉ cửa" | R0175, R0178 |
| Tổng các lớp = độ dày ghi trên BV; BV tự mâu thuẫn thì báo, không tự đổi | R0173 |
| Rebate đứng: cửa thường 35, cửa EI 80 | R0177 |
| Lớp mặt rebate theo vật liệu cửa, không gõ cứng 6mdf | R0171 |
| Σ SL các dòng cùng mã = tổng BoQ | R0048 |
| Mã hoàn thiện, mã vật tư có trong Material | R0059, R0143 |
| Ô nhập liệu không có nguồn → ASK_USER | quy tắc mới |

Phần này đưa vào `RULES_MASTER.xlsx` (Chủ đề riêng) và thành một module kiểm tra (mục 8).

## 7. Sổ độ phủ và định dạng dữ liệu trích

**LEDGER** (một dòng cho mỗi tờ hoặc trang): `file · trang · số BV · tầng đã dùng · số giá trị trích · số giá trị có cảnh báo · trạng thái (chưa/đã trích/đã kiểm/bỏ qua + lý do)`.

**Giá trị trích** (JSON/bảng): `mã/hạng mục · thuộc tính · giá trị · đơn vị · nguồn {file, trang, bbox, crop} · người đọc {script, claude, gemini} · khớp {có/không} · kiểm bất biến {đạt/không + lý do} · độ tin cậy`.

Độ tin cậy gợi ý: tầng A/B 0,95; tầng C 0,9; tầng D hai bên khớp và qua kiểm chuỗi 0,8; tầng D một bên 0,5; có lệch 0,3 trở xuống.

## 8. Công cụ và vị trí code

| Thành phần | Việc | Vị trí đề xuất |
|---|---|---|
| `extract_text.py` | Trích words, nhãn, số hiệu chi tiết (có bbox), lọc khung tên | `_MASTER\scan\` (bản git trong `_SCRIPTS\scan\`) |
| `crop_details.py` | Cắt crop theo chi tiết từ vị trí tiêu đề trong text layer, render 110–200 dpi | như trên |
| `read_gemini.py` | Gọi `agy` headless, ép schema JSON, chạy song song, lưu raw | như trên |
| `read_claude.py` | Chuẩn bị gói crop và prompt cho đọc bằng Claude (đọc trong phiên) | như trên |
| `reconcile.py` | So hai lần đọc, chạy kiểm chuỗi, tính độ tin cậy, xuất bảng lệch | như trên |
| `ledger.py` | Sổ độ phủ và báo cáo đối soát tập mã | như trên |
| `takeoff_checks.py` | Kiểm bất biến takeoff (mục 6.2) | `_MASTER\takeoff\` |
| `dwg_extract.py` | DWG → DXF → ezdxf (khi có ODA) | `_MASTER\scan\` |

Dùng venv sẵn có `_MASTER\tools\.venv` (đã có pymupdf, ezdxf, pdfplumber, openpyxl). Chạy bằng `python -I`, bật `sys.stdout.reconfigure(encoding="utf-8")` vì Windows cp1252.

### Cách gọi Gemini (đã kiểm chứng)

```
agy --model gemini-3.8-flash-high --print "<prompt>" --print-timeout 420s
```
Prompt phải có: "Do NOT run shell commands; use your file-viewing tool directly on this absolute path: <png>". Nếu không, Gemini chạy `Get-ChildItem` và chế độ headless tự từ chối. Không dùng `--dangerously-skip-permissions`. Chạy song song bằng `Start-Job`. Plugin telemetry Google đã tắt trong `.gemini\config\plugins\googlecloudtools.datacloud_telemetry\hooks.json` (bản sao `.bak`) vì lỗi quote đường dẫn.

### Mẫu prompt đọc (nguyên tắc)

- Chỉ hỏi chi tiết cụ thể trên crop đó, liệt kê giá trị cần.
- "Nếu không đọc được ghi null, không đoán, không suy từ BoQ."
- Yêu cầu chép nguyên nhãn đã đọc và ghi số chi tiết nguồn cho mỗi giá trị.
- Với chuỗi kích thước: yêu cầu liệt kê từng đoạn **và** tổng, để script cộng kiểm.
- Với đếm: yêu cầu mô tả vị trí từng phần tử đếm.

## 9. Lộ trình theo giai đoạn

| GĐ | Nội dung | Sản phẩm | Tiêu chí hoàn thành |
|---|---|---|---|
| 0 | Pilot Sofitel 5 tờ / 13 dòng | `PILOT_SOFITEL_2026-10-24.md` | **Đã xong**; đã tìm ra 3 lỗi (mục 2) |
| 1 | Áp kết quả pilot vào dữ liệu | v9 trong scratchpad: hạ độ tin cậy F018/F019 (~0,6), ghi F023 lệch 5 mm vào ISSUES/RFI, sửa ghi chú len kim loại 40 | Người dùng duyệt thì mới đưa vào `01_scan` |
| 2 | Benchmark trên NORTON | Bảng đo mỗi phương pháp (A, B, C, Claude, Gemini, kết hợp) so với đáp án đã duyệt (DH114 C210, M38, cánh 46; ID-DH111 M31, cánh 39…) | Biết tỉ lệ đúng/sót từng cách; chọn cấu hình mặc định |
| 3 | Thử tầng A | Quyết định có cài ODA File Converter không; đọc thử vài DWG NORTON/SOFITEL | Nêu rõ `ezdxf` lấy được gì (DIMENSION, TEXT, block) |
| 4 | Viết bộ script chuẩn (mục 8) | `extract → crop → read → reconcile → ledger` chạy trên 1 dự án từ đầu đến cuối | Chạy lại ra cùng kết quả; sổ độ phủ báo đủ |
| 5 | Đưa rule vào code | `takeoff_checks.py` + rule mới trong `RULES_MASTER.xlsx` | Các bẫy NORTON 23/10 (`06_cua-ei-va-rebate.md`) bị bắt tự động |
| 6 | Chạy thật trên Sofitel | Hoàn thiện 23 dòng THIẾU (F022, F085, F095, F098, F099 cần elevation phòng), RFI | Mọi dòng có nguồn hoặc nằm trong RFI |
| 7 | Đóng gói | Cập nhật skill `cpsx-master` (`skill-src\SKILL.template.md` → `build.js`), HANDOFF, hướng dẫn | Người dùng upload skill, chạy thử |

Thứ tự ưu tiên: **2 trước 4**. Cần số đo thật trước khi đầu tư vào pipeline.

## 10. Đo chất lượng

- **Bộ chuẩn NORTON:** các giá trị trong `01. DOOR\vi-du\NORTON\ghi-chu.md` là đáp án đã duyệt. Chấm mỗi phương pháp theo: đúng / sai / sót, và có bị bắt bởi kiểm bất biến hay không.
- **Bộ thử SOFITEL:** 13 dòng đã có đọc của session trước. Chưa phải đáp án (chưa ai duyệt), dùng để tìm chỗ lệch.
- **Chỉ số theo dõi:** tỉ lệ độ phủ, tỉ lệ hai người đọc khớp, số lỗi bị kiểm bất biến bắt, số giá trị chuyển cho người quyết, token tiêu thụ mỗi tờ.
- Mỗi lần người dùng sửa một giá trị, ghi lại loại lỗi để thêm kiểm tra tương ứng.

## 11. Tiết kiệm token

- Script lọc khung tên và ghi chú chung trước (đã thấy: phần lớn text trên tờ là boilerplate; lọc bằng cách bỏ chữ trùng vị trí trên nhiều tờ).
- Chỉ gửi crop vùng cần đọc, không gửi cả tờ. Render thấp (60–70 dpi) để định vị, cao (110–200 dpi) để đọc số.
- Phần đọc hình đẩy sang Gemini qua `agy` khi chỉ cần đọc; Claude chỉ làm người đọc thứ hai ở chỗ nghi ngờ.
- Lưu raw kết quả, không đọc lại cùng một crop.

## 12. Rủi ro và cách giảm

| Rủi ro | Cách giảm |
|---|---|
| Gemini đọc sai chuỗi dài, đếm sai | Cộng chuỗi bằng code, đọc kép, đếm lấy từ mặt bằng |
| Claude và Gemini sai giống nhau | Kiểm bất biến độc lập; mẫu NORTON có đáp án; tầng A/B khi có |
| Bản vẽ tự mâu thuẫn | Báo người dùng, không tự đổi (R0173), đưa vào RFI |
| Hạn mức Gemini (chỉ có Flash) | Cache raw; chạy theo lô; đọc thử một phần trước |
| Ổ D đầy khi tách PDF | Không tách PDF thành từng trang chứa ảnh nhúng; kiểm dung lượng trước khi ghi hàng loạt (đã từng đầy 23/10) |
| Không đọc được DWG | Tầng B/D làm thay; cân nhắc cài ODA File Converter |
| Mất kết quả khi Excel khóa file (~$PROJECT_DATA…v6) | Ghi phiên bản mới, không ghi đè |
| Sửa nhầm cấu hình công cụ của Google | Giữ `.bak`, ghi rõ trong tài liệu (mục 8) |

## 13. Việc cần người dùng quyết

1. Duyệt thứ tự giai đoạn (đề xuất: 1 → 2 → 3 → 4 …).
2. Có tải ODA File Converter không (để thử tầng A).
3. Có API key Gemini/AI Studio không (nếu cần model Pro hoặc hạn mức cao hơn; hiện `agy` chỉ có Gemini 3.8 Flash).
4. Dự án ưu tiên áp dụng đầu tiên: SOFITEL (đang dở) hay NORTON (có đáp án để đo).
5. Mức chấp nhận: có cần mỗi giá trị đọc kép bởi hai model, hay chỉ các giá trị đi vào giá?

## 14. Ghi chú nền

- Thư mục dự án: `D:\COSTING EXCEL\_ACTIVE\NORTON`, `D:\COSTING EXCEL\_ACTIVE\SOFITEL`.
- Hệ thống: `D:\COSTING EXCEL\_MASTER` (RULES_MASTER, macro, skill-src, scan, takeoff, tools); bản git: `D:\COSTING EXCEL\_SCRIPTS`.
- Tri thức cửa: `D:\COSTING EXCEL\01. DOOR\bai-hoc\_index.md`.
- Quy tắc mail: không bảng, không tổng, chi tiết ở file đính kèm (memory `email-tra-gia-style`).
- Ghi nhớ session: memory `agy-gemini-drawing-reading`.
