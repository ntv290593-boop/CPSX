# Kế hoạch v2: đọc bản vẽ đáng tin cậy cho bước scan và takeoff

Ngày lập: 2026-10-24 · Trạng thái: **đề xuất, chờ duyệt từng giai đoạn** · Thay thế `KE-HOACH-TONG-THE.md` (v1, giữ nguyên để đối chiếu).
Sửa theo hai vòng phản biện (một người dùng + một model khác) và theo các số liệu tôi đã kiểm lại ngày 2026-10-24.
Tài liệu liên quan: `PILOT_SOFITEL_2026-10-24.md`.

Thay đổi so với v1: SOFITEL **có** DWG nên tầng A lên trước; thêm nguồn đáp án mới (DWG và shop drawing DOT 01/02); thay thang độ tin cậy bằng quy tắc quyết định; sửa chuẩn độ phân giải theo R0167; đọc kép cho mọi giá trị ảnh hưởng giá; thêm kiểm tra cho hạng mục không phải cửa; sửa số liệu sai.

---

## 1. Vấn đề cần giải quyết

| # | Vấn đề | Hậu quả |
|---|---|---|
| 1 | Bước scan phải quét rất nhiều dữ liệu nên dễ sót | Thiếu mã cửa, thiếu chi tiết, không ai biết là thiếu |
| 2 | Bước takeoff không có cách kiểm tư duy | Một lần điền sai có thể làm sai cả dự án |
| 3 | Model nhìn hình có thể ảo giác hoặc bỏ sót | Số kích thước, số lượng, mã vật liệu sai mà vẫn trông hợp lý |
| 4 | Dùng token vào việc script làm được | Chậm, tốn kém, không lặp lại được |
| 5 | Khả năng đọc bản vẽ là rủi ro chính | Mọi bước sau thừa hưởng lỗi đọc |

**Tiêu chí thành công:** mỗi giá trị đi vào CPSX có (a) nguồn: file, trang, vị trí, ảnh crop; (b) ít nhất một kiểm tra độc lập; (c) nếu không chắc thì dừng ở ASK_USER hoặc RFI, không điền.

## 2. Hiện trạng (đã kiểm lại 2026-10-24)

- **Pipeline cũ** (`_MASTER\tools\scan_pages.py`, `scan\scan_schedules.py`, `scan\build_project_data_NORTON.py`, `takeoff\door_phieu_NORTON.py`): chỉ đọc text layer; `door_phieu_NORTON.py` gõ cứng giá trị đã duyệt.
- **NORTON:** sheet có text layer chứa số kích thước (`41(±1)`, `210`). `03_BanVe\_sheets` có 731 PDF theo sheet (cộng 7 file PDF gộp là 738) và 181 DWG theo sheet (`DWG_*`).
- **SOFITEL:**
  - 271 PDF một trang trong `01_scan\_sheets`. Số kích thước là đường vector (15–40 nghìn nét/tờ), không có trong text layer; chỉ số hiệu chi tiết, tỉ lệ, FCL/FFL là text.
  - **271 DWG** trong `01. INPUT\@DRAW` (PHASE 1: 190, PHASE 2: 81). 242 trong 242 tờ PDF có mã `ID-…` đều có DWG cùng mã; cả 5 tờ pilot có DWG. 29 PDF không có mã trong tên (PLAN…) chưa kiểm. **Chưa biết DWG có cùng phiên bản với PDF đấu thầu không.**
  - `01.DOT 01` và `02.DOT 02` có thêm **37 DWG shop drawing** (25 + 12), kèm 26 + 12 PDF và 6 + 1 xlsm (các file costing cũ): cửa chính DO-01, cửa WC DO-02, tủ minibar CA-02, vanity, khung gương… Có nhiều bản trùng tên (`_OS`, `_recover`, `Rev1/Rev2`, `V0/V1/V2`).
- **Dữ liệu SOFITEL v8:** hiện **11 dòng THIẾU** (F005, F022, F024, F085, F095, F098, F099, L043, D13, D14, D20). Con số 23 là của v7. 13 dòng đã có kết quả đọc từ session trước nhưng chưa ai duyệt.
- **Công cụ:** `agy` 1.4.6 (Gemini 3.8 Flash), pymupdf 2.0, ezdxf 1.4.5 (trong `_MASTER\tools\.venv`). Chưa có ODA File Converter (không có trong winget, tải tay, miễn phí, cần đăng ký).
- **Khổ giấy:** mọi tờ đã kiểm (NORTON và SOFITEL) là 2384×1684 pt, tức A1 đúng 1:1.
- **Pilot SOFITEL:** xem `PILOT_SOFITEL_2026-10-24.md`. Tóm tắt: Gemini tốt ở ghi chú và tổng, kém ở chuỗi dài và đếm; cộng chuỗi kích thước bắt được lỗi của cả Gemini và bản vẽ; hai kết luận sai của session trước (số lượng cột 7 suy từ BoQ; len kim loại 80 thay vì 40).
- **Rule hiện có:** R0166–R0179 đều thuộc dự án NORTON và đều là cửa. R0167 yêu cầu đọc bằng hình, render ≥200 dpi, crop từng chi tiết.

## 3. Nguyên tắc thiết kế

1. **Script trích, model chỉ diễn giải.** Việc máy làm chắc chắn được (đếm trang, đọc text, cộng chuỗi, so mã, tra bảng) do script làm. Model nhận crop nhỏ và câu hỏi hẹp.
2. **Hai lần đọc phải độc lập, không chỉ là hai model.** Pilot cho thấy hai bên cùng đọc "7" và cùng chấp nhận 470 (thực tế 475). Lần đọc thứ hai phải khác: độ phân giải khác, crop khác, hoặc hỏi ngược ("đoạn nào bằng 105?"). Nguồn kiểm độc lập thật là hình học (tầng A/B) và phép cộng chuỗi.
3. **Mọi giá trị có nguồn** gồm file, trang, bbox, ảnh crop, người đọc.
4. **Kiểm bất biến bằng code** (mục 6).
5. **Không đoán.** Thiếu nguồn thì ASK_USER hoặc RFI. Không suy số lượng từ BoQ rồi gọi là khớp.
6. **Đọc kép áp dụng cho mọi giá trị ảnh hưởng đến giá**, không chỉ kích thước: trạng thái OMITTED, số lượng, độ dày tường, mã vật liệu và mã hoàn thiện. Các giá trị khác đọc một lần và qua kiểm bất biến.
7. **Đọc ở độ phân giải ≥200 dpi trên crop từng chi tiết** (R0167). Độ phân giải thấp (60–70 dpi) chỉ dùng để định vị.
8. **Không ghi đè.** Mỗi lần ghi lớn tạo phiên bản mới. Không ghi trực tiếp .xlsm (HANDOFF §2).

## 4. Kiến trúc: bốn tầng đọc theo độ tin cậy

| Tầng | Cách | Giới hạn đã biết | Áp dụng |
|---|---|---|---|
| A | DWG → DXF (ODA File Converter, dự phòng LibreDWG `dwg2dxf`) → `ezdxf`: TEXT, MTEXT, DIMENSION, block attribute, layer | Xem mục 4.1 | SOFITEL (271 DWG), NORTON (181 DWG) |
| B | PDF vector (PyMuPDF `get_drawings`), đo hình học | Hiệu chỉnh bằng một kích thước đã biết trên chính tờ; ưu tiên thấp vì khổ A1 1:1 | Khi cần đo độ dày tường, bề rộng chi tiết |
| C | Text layer + bảng (pdfplumber, pymupdf words) | SOFITEL không có số kích thước trong text | NORTON; SOFITEL chỉ lấy số hiệu, tỉ lệ, nhãn |
| D | Model nhìn crop ≥200 dpi, đọc kép, kèm kiểm bất biến | Sai ở chuỗi dài và đếm | Phần còn lại |

Quy tắc chọn tầng: dùng A→B→C trước, D cho phần chưa có. Kết quả D luôn qua kiểm bất biến và đọc kép.

### 4.1 Rủi ro riêng của tầng A (DWG)

- **Tầng A cho số và nhãn chính xác nhưng không cho nghĩa.** Việc ghép kích thước vào đúng chi tiết và lớp vẫn phải làm bằng vị trí hình học. Tầng A giảm mạnh phần việc của tầng D nhưng không loại bỏ.
- **Người vẽ gõ đè text kích thước:** giá trị đo thật khác số hiển thị. So `get_measurement()` với text hiển thị; chỉ nhận khi khớp. Không tin tầng A ở mức tuyệt đối.
- **Kích thước có thể nằm trong block ẩn (`*D1`), xref hoặc layout (paper space)**, không nằm ở model space.
- **Phiên bản:** phải kiểm DWG cùng phiên bản với PDF đấu thầu (so tên chi tiết, ngày, số revision; so vài số đã biết từ pilot). Nếu chưa kiểm được thì giá trị lấy từ DWG có thể sai có hệ thống.
- **Ghép DWG với tờ PDF:** theo mã `ID-…` (242/242 đã khớp); 29 tờ không mã cần ghép tay.
- **Công cụ:** ODA cần đăng ký; LibreDWG đọc DWG đời mới kém hơn.

## 5. Quy trình

### 5.1 Scan

1. **Kiểm kê:** liệt kê file, số trang, loại; ghép DWG với PDF; ghi vào LEDGER.
2. **Phân loại trang bằng script** (khung tên; `scan_pages.py`). Trang không đọc được khung tên thì cảnh báo.
3. **Chọn tờ liên quan** từ BoQ (mã cửa, mã hoàn thiện, hạng mục) bằng text layer, DETAILS, `_LINKS`; dùng Spec làm nguồn cho mã hoàn thiện.
4. **Trích tầng A/B/C** tự động, ghi nguồn.
5. **Phần còn lại sang tầng D:** cắt crop theo chi tiết (từ vị trí tiêu đề chi tiết), render ≥200 dpi; hai lần đọc độc lập (mục 3.2), schema JSON cố định. Claude chỉ xử lý danh sách lệch và các giá trị nghi ngờ (mục 8).
6. **Đối soát:** so các lần đọc, chạy kiểm bất biến, áp quy tắc quyết định (mục 7).
7. **So sánh bản phát hành lại:** khi có bản mới của một tờ, so khác biệt giữa hai lần phát hành và đánh dấu giá trị bị ảnh hưởng.
8. **Báo cáo độ phủ** (mục 7).
9. **Ghi PROJECT_DATA phiên bản mới** theo cấu trúc NORTON v8, link tự trỏ + `_LINKS` (memory `project-data-format`).

### 5.2 Takeoff

1. Đầu vào là dữ liệu đã trích có nguồn, không đọc lại bản vẽ bằng tay.
2. Phiếu bóc sinh từ dữ liệu, không gõ cứng như `door_phieu_NORTON.py`.
3. Kiểm bất biến takeoff (mục 6.2) chạy tự động; vi phạm thì không xuất AI_INPUT.
4. Lệch hồ sơ (Spec ≠ BoQ ≠ BV) sinh Issue → RFI theo `bai-hoc\05_lech-ho-so-rfi.md`.
5. Ghi `02_takeoff\AI_INPUT.xlsx`, không ghi .xlsm trực tiếp.

## 6. Kiểm bất biến (chạy bằng code)

### 6.1 Ở bước scan

| Kiểm tra | Cách làm |
|---|---|
| Chuỗi kích thước cộng ra tổng (dung sai = số đoạn × DIMRND/2 khi dimstyle có làm tròn, ví dụ 3 khe × 2,5 mm) | Lệch quá dung sai thì báo "đọc sai" hoặc "bản vẽ tự mâu thuẫn" (R0167: tổng dày phải khớp kích thước tổng) |
| Số lượng phải có nguồn đếm | Chỉ nhận từ mặt bằng, mặt đứng, bảng thống kê. Không suy từ BoQ chia kích thước |
| Mã trên BV có trong Finishes Schedule / Spec | Mã lạ → cảnh báo |
| Mã cửa BoQ ↔ thống kê ↔ mặt bằng | Khác tập hợp → báo thiếu/thừa |
| Các lần đọc khớp nhau | Lệch → chuyển người quyết |
| Giá trị đo từ DWG khớp text hiển thị | Lệch → nghi gõ đè, không nhận |
| Có ảnh crop và vị trí | Thiếu nguồn → không ghi |

### 6.2 Ở bước takeoff

**Cửa** (rule hiện có, đều từ dự án NORTON):

| Kiểm tra | Rule |
|---|---|
| Mỗi mã cửa có độ dày tường kèm tờ mặt bằng nguồn | R0166, memory `door-wall-thickness-check` |
| Đọc bằng hình, bảng lớp cấu tạo, tổng dày khớp tổng ghi trên BV | R0167, R0172 |
| Cửa EI: mọi cấu phần có ≥1 lớp ERON, trừ "chỉ cửa" | R0178 (cùng R0174–R0176) |
| Rebate đứng: cửa thường 35, cửa EI 80 | R0177 |
| Lớp mặt rebate theo vật liệu cửa, không gõ cứng 6mdf | R0171 |
| Số tự mâu thuẫn thì báo user, không tự đổi | R0173 (rule viết cho Spec; plan mở rộng cho bản vẽ) |
| Σ SL các dòng cùng mã = tổng BoQ | R0048 |
| Mã hoàn thiện, mã vật tư có trong Material | R0059, R0143 |
| Ô nhập liệu không có nguồn → ASK_USER | quy tắc mới |

**Không phải cửa** (đề xuất, cần duyệt rồi ghi vào `RULES_MASTER.xlsx`; SOFITEL chủ yếu là ốp tường, len, vách, FF&E):

| Hạng mục | Kiểm tra đề xuất |
|---|---|
| Len, ốp | Chiều dài đo trên mặt bằng xấp xỉ khối lượng BoQ (ngưỡng sai lệch do người dùng chọn) |
| FF&E | Số lượng BoQ = số trong bảng thống kê phòng × số phòng |
| Cấu kiện lặp (cột, thanh) | Số đơn vị lấy từ mặt bằng/mặt đứng, không suy từ diện tích BoQ |
| Mọi hạng mục | Kích thước trên BV so với kích thước trong mô tả BoQ; lệch → Issue |

## 7. Quy tắc quyết định, sổ độ phủ và cổng hoàn thành

**Quy tắc quyết định** (thay thang độ tin cậy; mặc định thận trọng, hiệu chỉnh sau benchmark):

| Quyết định | Điều kiện |
|---|---|
| **Nhận** | Tầng A khớp text hiển thị, hoặc hai lần đọc độc lập khớp **và** qua mọi kiểm bất biến áp dụng được |
| **Xem lại** | Chỉ một lần đọc, hoặc hai lần khớp nhưng một kiểm bất biến không áp dụng được, hoặc có cảnh báo nhẹ |
| **ASK_USER / RFI** | Hai lần đọc lệch, vi phạm kiểm bất biến, không có nguồn, hoặc bản vẽ tự mâu thuẫn |

**LEDGER** (một dòng cho mỗi tờ): `file · trang · số BV · tầng đã dùng · số giá trị trích · số cảnh báo · trạng thái (chưa / đã trích / đã kiểm / bỏ qua + lý do)`.

**Giá trị trích:** `mã/hạng mục · thuộc tính · giá trị · đơn vị · nguồn {file, trang, bbox, crop} · người đọc · kết quả đọc kép · kiểm bất biến {đạt/không + lý do} · quyết định (nhận / xem lại / ASK_USER)`.

**Cổng hoàn thành bước scan:** `số dòng BoQ có nguồn ĐÃ QUA KIỂM / tổng số dòng BoQ`. "Có nguồn" không đủ (F018 có nguồn nhưng chỉ chứng minh kích thước, không chứng minh số lượng). Dòng LOẠI TRỪ phải có lý do và người duyệt. Ngoài ra báo số tờ chưa xử lý.

## 8. Công cụ và vai trò của từng người đọc

| Thành phần | Việc | Vị trí đề xuất |
|---|---|---|
| `extract_text.py` | Trích words, nhãn, số hiệu chi tiết (có bbox), lọc khung tên | `_MASTER\scan\` (bản git ở `_SCRIPTS\scan\`) |
| `dwg_extract.py` | DWG → DXF → ezdxf; so `get_measurement()` với text hiển thị; ghép với chi tiết bằng vị trí | `_MASTER\scan\` |
| `crop_details.py` | Cắt crop theo chi tiết, render ≥200 dpi | như trên |
| `read_gemini.py` | Gọi `agy` headless, schema JSON, lưu raw, giới hạn số job song song | như trên |
| `reconcile.py` | So các lần đọc, cộng chuỗi, xuất bảng lệch và quyết định | như trên |
| `ledger.py` | Sổ độ phủ, đối soát tập mã, cổng hoàn thành | như trên |
| `takeoff_checks.py` | Kiểm bất biến takeoff (mục 6.2) | `_MASTER\takeoff\` |

Dùng venv `_MASTER\tools\.venv`. Chạy `python -I`, bật `sys.stdout.reconfigure(encoding="utf-8")`.

**Vai trò Claude (đã chọn một hướng):** Claude chỉ xử lý danh sách lệch và các giá trị nghi ngờ, đọc crop ≥200 dpi trong phiên, không đọc mọi crop. Gọi Claude qua API bằng script (cache được, lặp lại được) cần API key và tính phí riêng, không nằm trong gói Claude Code; quyết cùng câu hỏi về API key ở mục 13. Gemini là người đọc chính cho phần tầng D.

**Cách gọi Gemini (đã kiểm chứng):**
```
agy --model gemini-3.8-flash-high --print "<prompt>" --print-timeout 420s
```
Prompt phải có: "Do NOT run shell commands; use your file-viewing tool directly on this absolute path: <png>". Không dùng `--dangerously-skip-permissions`. Plugin telemetry Google đã tắt trong `.gemini\config\plugins\googlecloudtools.datacloud_telemetry\hooks.json` (bản sao `.bak`) vì lỗi quote đường dẫn. Giới hạn số job chạy cùng lúc (5 job đã chạy ổn; 271 tờ chưa thử, có thể chạm giới hạn lượt gọi).

**Nguyên tắc prompt đọc:** hỏi chi tiết cụ thể trên crop đó; "không đọc được thì null, không đoán, không suy từ BoQ"; chép nguyên nhãn và số chi tiết nguồn; với chuỗi kích thước yêu cầu liệt kê từng đoạn và tổng; với đếm yêu cầu mô tả vị trí từng phần tử.

## 9. Lộ trình

Thứ tự đề xuất: **1 → 3 (thử có giới hạn) → 3b (đáp án) → 2 → 4 → 5 → 6 → 7**.

| GĐ | Nội dung | Sản phẩm | Tiêu chí hoàn thành |
|---|---|---|---|
| 0 | Pilot SOFITEL 5 tờ | `PILOT_SOFITEL_2026-10-24.md` | **Đã xong**; tìm ra 3 lỗi |
| 1 | Áp kết quả pilot vào dữ liệu | v9 trong scratchpad: F018/F019 chỉ kích thước đã kiểm, số lượng chưa; F023 giữ 470 (khe thật 103,33, hiển thị 105 do DIMRND=5, không phải lỗi bản vẽ; xem `KET-QUA-GD3-TANG-A_2026-10-24.md`); sửa ghi chú len kim loại 40 | Người dùng duyệt thì mới đưa vào `01_scan` |
| 3 | **Thử tầng A (có giới hạn thời gian)** | Cài ODA (hoặc LibreDWG); đọc DWG của 5 tờ pilot | Tiêu chí thoát: toàn bộ kích thước 5 tờ khớp với các giá trị đã tự xem lại bằng crop; nếu không đạt thì ghi lý do và quay về tầng D. Kiểm DWG cùng phiên bản với PDF trước khi tin |
| 3b | **Lập bộ đáp án** | Nguồn: (a) DWG thiết kế, chỉ nhận kích thước có giá trị đo = text hiển thị và đã khớp phiên bản; (b) shop drawing DOT 01/02 (đã qua tay người dùng, cần xác nhận phạm vi và bản chốt); (c) `door_phieu_NORTON.py` (~60 dòng có nguồn); người dùng duyệt mẫu 10–15 giá trị | Bộ đáp án có nguồn, người dùng đã duyệt mẫu |
| 2 | Benchmark | Bảng đo mỗi cấu hình (A, B, C, Claude, Gemini, kết hợp) so với đáp án. NORTON dùng để so tầng C với tầng D; SOFITEL cho đúng điều kiện không có text | Chọn cấu hình mặc định theo **ngưỡng đặt trước** (ví dụ: không có sai lọt qua mọi kiểm tra trên bộ đáp án). Lưu ý: 50 giá trị không sai chỉ chứng minh tỉ lệ sai < khoảng 6% (95% tin cậy), nên ngưỡng "<1%" cần bộ lớn hơn nhiều; ghi rõ cỡ mẫu khi báo kết quả |
| 4 | Bộ script chuẩn (mục 8) | `extract → crop → read → reconcile → ledger` chạy trên một dự án đầu đến cuối | Chạy lại ra cùng kết quả; cổng hoàn thành báo đủ |
| 5 | Đưa rule vào code | `takeoff_checks.py`; rule mới (hạng mục không phải cửa) vào `RULES_MASTER.xlsx` sau khi người dùng duyệt | Các bẫy NORTON 23/10 (`06_cua-ei-va-rebate.md`) bị bắt tự động |
| 6 | Chạy thật trên SOFITEL | Hoàn thiện khoảng 11 dòng THIẾU (F022, F085, F095, F098, F099 cần elevation phòng) và RFI | Mọi dòng có nguồn đã qua kiểm hoặc nằm trong RFI |
| 7 | Đóng gói | Cập nhật skill `cpsx-master`, HANDOFF, hướng dẫn | Người dùng upload skill, chạy thử |

## 10. Đo chất lượng

- **Bộ đáp án:** xem GĐ3b. Giá trị trong `01. DOOR\vi-du\NORTON\ghi-chu.md` là cấu hình block đã chọn ở bước takeoff, chưa phải số đọc thẳng từ bản vẽ; cần đáp án ở mức bản vẽ (kích thước, độ dày tường, các lớp, tờ và vị trí).
- **13 dòng SOFITEL** đã có đọc của session trước: chưa phải đáp án, dùng để tìm chỗ lệch.
- **Chỉ số:** tỉ lệ độ phủ đã qua kiểm, tỉ lệ các lần đọc khớp, số lỗi bị kiểm bất biến bắt, số giá trị chuyển cho người quyết, token trên mỗi tờ.
- Mỗi lần người dùng sửa một giá trị thì ghi loại lỗi và thêm kiểm tra tương ứng.

## 11. Tiết kiệm token

- Script lọc khung tên và ghi chú chung trước (bỏ chữ trùng vị trí trên nhiều tờ).
- Định vị bằng ảnh độ phân giải thấp (60–70 dpi); **đọc giá trị ở ≥200 dpi trên crop** (R0167), không gửi cả tờ.
- Đọc hình đẩy sang Gemini qua `agy`; Claude chỉ đọc danh sách lệch.
- Lưu raw kết quả, không đọc lại cùng một crop.

## 12. Rủi ro và cách giảm

| Rủi ro | Cách giảm |
|---|---|
| Gemini đọc sai chuỗi dài, đếm sai | Cộng chuỗi bằng code; đếm lấy từ mặt bằng; lần đọc thứ hai khác độ phân giải/crop |
| Hai model sai giống nhau | Kiểm bất biến và hình học độc lập (tầng A/B); bộ đáp án |
| DWG khác phiên bản với PDF | Kiểm phiên bản trước; nếu chưa kiểm được thì không dùng DWG làm đáp án |
| Text kích thước bị gõ đè trong DWG | So `get_measurement()` với text; chỉ nhận khi khớp |
| Kích thước trong block ẩn, xref, paper space | Duyệt cả ba nơi khi trích |
| Không cài được ODA, LibreDWG kém | Dự phòng tầng B/D; ghi rõ trong kết quả GĐ3 |
| Nguồn "có" nhưng không chứng minh đúng giá trị | Cổng hoàn thành đếm "đã qua kiểm", không chỉ "có nguồn" |
| Bản vẽ tự mâu thuẫn | Báo người dùng, không tự đổi (mở rộng từ R0173), đưa vào RFI |
| Bản vẽ phát hành lại làm đáp án cũ hết đúng | So khác biệt giữa hai lần phát hành (mục 5.1.7) |
| Hạn mức Gemini (chỉ có Flash) | Cache raw; chạy theo lô; giới hạn số job song song |
| Ổ D đầy khi tách PDF | Không tách PDF chứa ảnh nhúng; kiểm dung lượng trước khi ghi hàng loạt (đã từng đầy 23/10) |
| Mất kết quả khi Excel khóa file | Ghi phiên bản mới, không ghi đè |
| Sửa nhầm cấu hình công cụ của Google | Giữ `.bak`, ghi rõ ở mục 8 |

## 13. Việc cần người dùng quyết

Đề xuất đã chốt theo các vòng phản biện (cần người dùng xác nhận):

1. **Thứ tự:** 1 → 3 → 3b → 2 → 4 → 5 → 6 → 7.
2. **ODA File Converter:** nên tải.
3. **API key Gemini / Claude API:** chưa cần; quyết sau GĐ2.
4. **Dự án:** NORTON để đo, SOFITEL để chạy thật.
5. **Đọc kép:** áp dụng cho mọi giá trị ảnh hưởng đến giá (kích thước, số lượng, độ dày tường, mã vật liệu và hoàn thiện, trạng thái OMITTED).

Câu hỏi mới cần người dùng trả lời:

6. **Shop drawing DOT 01/02** (37 DWG, 7 xlsm) có cùng phạm vi và cùng phiên bản thiết kế với đợt đang chào giá không? **Đã trả lời 2026-10-24: cùng phiên bản.** (Câu 7 về bản chốt vẫn mở.)
7. Trong các bản trùng tên (`_OS`, `_recover`, `V0/V1/V2`, `Rev1/Rev2`), **bản nào là bản chốt?**
8. Có duyệt các kiểm tra đề xuất cho hạng mục không phải cửa (mục 6.2) để ghi vào `RULES_MASTER.xlsx` không, và ngưỡng sai lệch chấp nhận được cho len/ốp là bao nhiêu?

## 14. Ghi nền

- Dự án: `D:\COSTING EXCEL\_ACTIVE\NORTON`, `D:\COSTING EXCEL\_ACTIVE\SOFITEL`.
- NORTON bản vẽ: `_ACTIVE\NORTON\00_input\261023_D01\03_BanVe\_sheets` (731 PDF) và `DWG_*` (181 DWG).
- SOFITEL DWG: `_ACTIVE\SOFITEL\01. INPUT\@DRAW\PHASE 1-COMBINE\DWG\` và `PHASE 2-COMBINE\DWG\`; shop drawing: `01.DOT 01`, `02.DOT 02`.
- Hệ thống: `D:\COSTING EXCEL\_MASTER` (RULES_MASTER, macro, skill-src, scan, takeoff, tools); bản git: `D:\COSTING EXCEL\_SCRIPTS`.
- Tri thức cửa: `D:\COSTING EXCEL\01. DOOR\bai-hoc\_index.md`.
- Mail trả giá: không bảng, không tổng, chi tiết ở file đính kèm (memory `email-tra-gia-style`).
- Ghi nhớ: memory `agy-gemini-drawing-reading`.
