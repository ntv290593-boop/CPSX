# Plan: đọc sâu 10-tender — đọc mỗi trang 1 lần + giảm token (bước 6, 7c, 7d)

Ngày: 2026-11-14 · Trạng thái: **ĐANG LÀM — xong A1–A5 + 2 lỗi (15/11, chưa commit); còn A6–A10, chưa đo 7d thật** · Dự án đo: GBIA_VVIP (`_ACTIVE\GBIA_VVIP\10_TENDER`)
Đầu vào: `DOCS\HANDOFF_DOCSAU-PHAN-LOAI_2026-11-14.md` (vấn đề gửi chuyên gia) + phản hồi chuyên gia "Read Once – Extract All – Reuse Everywhere – Resolve Per BOQ" (user dán trong phiên 14/11, không lưu file).

## 1. Kết luận kiểm tra phản hồi chuyên gia

**Đúng hướng, đã đối chiếu code:**
- Chưa có bước gom câu hỏi trước khi đọc. 7d hướng dẫn "Làm, cho từng cụm" (`buoc\7d-docsau.md`).
- Chưa có manifest trang. `docsau.py` `cmd_docsau` gom cụm theo **phòng đầu tiên**, cắt 8 cụm/nhóm (dòng 168–180).
- Agent song song không dùng chung kết quả. `PAGES` chỉ dùng cho lần chạy lại.

**Thiếu / chưa đúng:**
1. Không nói cách chia trang cho agent song song. Trang điển hình (FUR-0006 p20/21: 19 cụm, 3 nhóm) nối hết cụm thành một khối lớn → cần **lượt 0** đọc trang dùng chung trước, rồi mới chia phần còn lại.
2. "Đọc cả trang 1 lần" ≠ 1 ảnh: 100 dpi không đọc được số, vẫn phải crop chi tiết. Hiểu là **1 lần ngồi đọc mỗi trang, nhiều crop**.
3. Ví dụ "BOQ 01 = vật liệu, BOQ 02 = kích thước" sai thực tế: mỗi cụm cần đủ bộ trường cố định (7d mục 3) → **máy tự lập manifest** từ `Bản vẽ hiệu lực`, không cần LLM gom câu hỏi.
4. Bỏ sót bước 6 (banve) và 7c: cũng mở PDF, chia theo trục khác (phòng / sheet).
5. Không nói lần TCQ sau: cần sổ đọc theo **tờ + revision**.

**Lỗi thêm trong skill (handoff chưa nêu):**
- 7d mục Gemini ý 2 ghi "crop dùng chung không tốn thêm lượt vì kết quả lưu cạnh ảnh". Chỉ đúng một phần: `SKILL\_core\gemini_doc.py:63` chỉ cache ở chế độ số và khi **cùng đường dẫn PNG**. `--quansat` luôn gọi agy lại. Agent khác cắt cùng vùng ra file khác cũng gọi lại.
- `docsau.py:201` chỉ in 6 tờ đầu mỗi cụm (`[:6]`).

## 2. Số đo GBIA (14/11)

### Đọc trùng
| Chỉ số | Giá trị | Nguồn |
|---|---|---|
| PAGES | 249 lượt / 188 trang; 54 trang bị ≥ 2 nhóm đọc (tối đa 3) | `_data\PAGES.docsau-*.json` |
| Gemini số | 285 lượt / 214 crop khác nhau; 57 crop dùng ≥ 2 cụm | `_data\_gemini_so\*.json` |
| Danh sách giao | 518 lượt tờ (cụm × tờ) / 185 trang | handoff mục 3 |

### Token (transcript subagent phiên `1ff69063…`, 21 agent 7d)
| Chỉ số | Giá trị |
|---|---|
| cache read / cache write / output | **125M** / 5,2M / 0,84M (quy theo giá: ~54% / ~28% / ~18%) |
| Lượt | 1039 (~49/agent); ngữ cảnh 41k → ~200k mỗi agent |
| Ảnh | 437 ảnh, ~4.200 tok/ảnh; 349/567 PNG > 1,15MP; trang A3 100 dpi = 1654×1170 ≈ 1,9MP; dpi agent tự chọn 45–400 |

Phân rã 125M đọc lại (ước = kích thước × số lượt còn lại, khớp thực tế 127M vs 125M):
| Phần | % |
|---|---|
| Ảnh PNG | ~45% |
| Nền cố định ~41k/lượt (system + tool + CLAUDE.md + MEMORY.md) | ~34% |
| Output lệnh (loc, cot, shell, script tự viết 76 lần) | ~15% |
| File hướng dẫn (7d, JSON.md, _docsau.txt) | ~5% |

**Nguyên nhân:** (1) ảnh to hơn cần; (2) ảnh ở lại ngữ cảnh suốt 8 cụm → chi phí tăng ~bình phương số ảnh/agent; (3) nhiều lượt lặt vặt, mỗi lượt ≥ 41k nền; (4) đọc trùng giữa nhóm (~25–30% ảnh). Gemini (agy) không tốn token Claude, chỉ tốn thời gian.

Script đo (scratchpad, có thể mất): logic = với mỗi tool_result/ảnh, token × (số lượt còn lại của agent); token ảnh = (tăng ngữ cảnh − chữ) / số ảnh. Transcript ở `C:\Users\Admin\.claude\projects\d--COSTING-EXCEL\<phiên>\subagents\agent-*.jsonl`, nhận agent 7d theo chữ "7d"/"docsau" trong prompt đầu.

## 3. Plan gộp (chờ duyệt)

| # | Việc | File | Ước giảm |
|---|---|---|---|
| A1 | `anh` giới hạn pixel: tổng thể ~1000 px cạnh dài, crop ≤ ~1,15MP; bỏ để agent chọn dpi tự do | `tool\td.py` (lệnh anh) / `_core`, 7d mục "Tiết kiệm token" | token ảnh −50–60% (~¼ tổng) |
| A2 | Manifest trang: `docsau` đảo cụm→tờ thành tờ→cụm → `_data\_doc_trang.json` (trang, FID, rev, các cụm, trường cần); bỏ `[:6]`, danh sách tờ đầy đủ | `tool\docsau.py` | nền cho A3–A5 |
| A3 | **Lượt 0** (trước 7d, tuần tự hoặc ít agent): trang dùng ≥ N cụm (N đề xuất 3) đọc 1 lần → `TRANG_QS` (quan sát theo trang + chi tiết, có tờ/số chi tiết). Gemini chạy theo crop của trang | `docsau.py`, JSON.md (sheet mới), 7d | bỏ đọc trùng ~25–30% ảnh |
| A4 | **Lượt 1** (song song): chia agent theo bộ tờ còn lại (thành phần liên thông, cắt theo số trang, ít trang/agent để ảnh không ở lâu). Agent kế thừa `TRANG_QS`, vẫn ghi `QUAN_SAT` theo cụm → 7e, `neo` giữ nguyên | `docsau.py`, 7d, SKILL.md bảng bước + 1e | cắt phần tăng bình phương |
| A5 | Cache Gemini theo **nội dung** (trang + vùng clip, hoặc hash PNG), áp cả `--quansat` | `_core\gemini_doc.py` | Gemini trùng → 0; sửa câu sai ở 7d mục Gemini ý 2 |
| A6 | Lệnh gộp `xem-cum <khóa>` / `xem-trang <FID> <p>`: in 1 lần dòng BOQ, tờ, chữ trang, quan sát đã có | `td.py` / `docsau.py` | ít lượt (mỗi lượt −41k) |
| A7 | Kiểm sau đọc: cùng tờ + chi tiết mà giá trị khác giữa các cụm → issue | `docsau.py` / `check` | chất lượng |
| A8 | Prompt 7d dán sẵn phần JSON.md cần dùng, không để agent tự đọc | SKILL.md mục giao việc | ~5% |
| A9 | Sổ đọc theo tờ + revision cho `--thaydoi` (TCQ sau): tờ không đổi thì dùng lại `TRANG_QS` | `docsau.py`, `thaydoi` | lần TCQ sau |
| A10 | Bước 6 (banve) thống nhất trục với manifest trang (sau cùng, cần bàn) | 6-banve.md | — |

Thứ tự đề xuất: **A1** (ít code, giảm nhiều nhất) → A2 → A5 → A3 + A4 → A6 → A7, A8 → A9, A10.
Đo lại sau khi sửa bằng cùng cách đo mục 2 (chạy 7d thử trên GBIA, vài nhóm).

Theo CLAUDE.md: đổi tính năng skill thì sửa `DOCS\Flowchart\10-tender_flowchart.md` + 1 dòng `CHANGELOG.md` cùng phiên. Sửa ở `SKILL\10-tender` (bản `.claude\skills\10-tender` là junction).

## 3b. Tiến độ 15/11 (kiemthu --mau 86/86, chưa commit)
- **2 lỗi skill:** `docsau` in đủ mọi tờ (bỏ `[:6]`; GBIA 14 cụm > 6 tờ); câu sai 7d mục Gemini ý 2 đã viết lại.
- **A5:** `_core\gemini_doc.py` cache theo sha1(byte PNG) + sha1(model + câu hỏi), mọi chế độ; `gemini --goc` dùng chung `_data\_gemini_cache`. Chỉ trúng khi ảnh giống hệt byte → cần A1 cố định dpi.
- **A1:** `anh` tự hạ dpi (`doc_file.dpi_gioi_han`, bội số 5): cả trang ≤ 1000 px cạnh dài, vùng cắt ≤ 1,15 MP / 1568 px; dưới ngưỡng giữ nguyên (test giống pd.py vẫn đạt).
- **A2:** `_data\_doc_trang.json` (trang → cụm). Bản sao GBIA: 196 trang, 548 lượt tờ; **74 trang ≥ 3 cụm gánh 376 lượt (69%)** → ngưỡng N=3 cho lượt 0 có cơ sở. Cột `rev` trống ở GBIA vì mã tờ `FUR-3901` không khớp `RX_DWG` (xử lý ở A9).
- **A3+A4 (N = 3):** `docsau` chia 2 lượt. Lượt 0 = trang ≥ 3 cụm, ≤ 8 trang/agent → sheet `TRANG_QS` (Khóa trang, Chi tiết, Món, Cụm đọc, Crop…); lượt 1 = cụm theo bộ tờ riêng (thành phần liên thông), ≤ 8 cụm và ≤ 12 trang riêng/agent, đọc `TRANG_QS` thay vì mở lại ảnh; số nhóm lượt 1 nối tiếp lượt 0. `check` LỖI khi trang lượt 0 chưa có TRANG_QS. Bản sao GBIA: 196 trang → lượt 0 10 nhóm/74 trang + lượt 1 23 nhóm/122 trang = mỗi trang giao đúng 1 nhóm (trước 548 lượt tờ). Số agent tăng 19 → 33 nhưng mỗi agent ít trang hơn (lượt 1 tối đa 12 trang riêng).
- Rủi ro chưa kiểm: lượt 1 chép từ TRANG_QS có gán đúng món → cụm không (`Cụm đọc` lượt 0 có thể trống); chỉ đo được khi chạy 7d thật.
- **Đo 7d thật 15/11** (bản sao GBIA ở scratchpad, 5 agent: lượt 0 nhóm 2, 9 → lượt 1 nhóm 13, 14, 20 = 23 cụm; đối chứng = 23 cụm đó trong 21 agent 14/11, chia theo số cụm; lượt 0 tính phần 11/23 và 7/20). Script `tok3.py` (scratchpad, có thể mất; logic như mục 2).

  | 23 cụm | 14/11 | 15/11 | |
  |---|---|---|---|
  | Lượt gọi | 171 | 175 | +2% |
  | Ảnh | 62 | 56 | −8% |
  | cache read | 21,2M | 20,1M | −5% |
  | cache write | 0,93M | 0,56M | −39% |
  | output | 0,11M | 0,16M | +42% |
  | **Quy theo giá** | | | **−9%** |

  - Lượt 1 không mở lại ảnh lượt 0 (nhóm 14: 1 ảnh); peak ngữ cảnh lượt 1 ~145k (trước ~190k). Rủi ro `Cụm đọc` trống: không xảy ra, món nhận đúng cụm; kích thước tương đương bản cũ.
  - Không giảm nhiều vì **số lượt không giảm** (~7,6 lượt/cụm) và cache read/lượt chỉ 112k vs 120k: nền + output lệnh + JSON tự ghi chiếm phần lớn, không phải ảnh. Nhóm 13 (5/8 cụm có tờ riêng) tốn 8,2M = như cũ. Lợi ích bỏ đọc trùng chỉ thấy đủ khi chạy cả 33 nhóm.
  - → Đòn bẩy tiếp theo là **số lượt**: A6 (lệnh gộp), A8 (dán JSON.md cần dùng), bớt `script`/`shell` tự viết (44 lần / 5 agent).
  - Lỗi phát hiện: `anh` FileNotFoundError khi đường dẫn PDF > 260 ký tự (bản sao 310) → đã sửa `doc_file.dai()` (tiền tố `\\?\`), kiemthu 88/88.

## 4. Câu hỏi mở cho user
1. Duyệt thứ tự trên? Làm A1 trước riêng hay làm cả gói?
2. Ngưỡng N trang "dùng chung" cho lượt 0 (đề xuất 3 cụm).
3. Chạy thử lại 7d trên GBIA để đo (tốn token) hay chờ dự án sau?

## 5. Trạng thái liên quan
- Git: còn sửa chưa commit từ trước (`td.py`, `dmvt.py`, `MAU_DU_AN.json`, flowchart, ACTIVE.md) — thuộc việc GBIA build 261114, không phải plan này.
- Kết quả 7d GBIA hiện có vẫn dùng được (định dạng ghi theo cụm không đổi).
