# HANDOFF — SOFITEL: nhận dạng mẫu + bổ sung mục chưa tính vào RE00 (2026-10-24)

Viết cho session Claude tiếp theo (đọc xong là làm tiếp được). Người dùng nói tiếng Việt, thích câu trả lời gọn, mỗi lượt chỉ 1 câu hỏi kèm đề xuất, duyệt theo ngoại lệ (xem memory `review-by-exception`). Đọc trước: `_MASTER\modules\NGUYEN-TAC.md`, `NHAN-DANG-MAU.md` (có cảnh báo ở đầu), `_MASTER\schemas\README.md`.

## 1. Mục tiêu
(1) Dạy model nhận dạng "mẫu" (Template/Door) cho từng hạng mục BoQ từ bản vẽ PDF + dòng BoQ. (2) Bổ sung vào file gốc dự án SOFITEL các hạng mục chưa tính, bằng cách chép block tương tự đã có (ThaoTac `MOI` qua `L.fromAiInput`), không dựng lại.

## 2. File — đừng nhầm
| File | Là gì |
|---|---|
| `_ACTIVE\SOFITEL\261020_SOFITEL_CPSX_RE00.xlsm` | **FILE GỐC** của user. 74 block, `Chiet tinh` đến hàng 5414 (D14). KHÔNG ghi trực tiếp, làm trên bản sao ở `_TEST` |
| `…\03_costing\261024_SOFITEL_CPSX_REV02.xlsm` | RE00 + block D15 (75 block, 5505 hàng) |
| `…\03_costing\261023_SOFITEL_CPSX_REV01.xlsm` | **Bản dẫn xuất**: RE00 + các đợt Claude dựng tối 23/10 + D15 + đợt 1 (114 block, 8342 hàng). Không phải gốc. Có sao lưu `_backup\…truoc-batch1_165202.xlsm` |
| `…\01_scan\PROJECT_DATA_SOFITEL_v8.xlsx` | Scan v8 (đang bị Excel khóa lúc cuối phiên) |
| `…\02_takeoff\BO_SUNG_RE00_v2.xlsx` | Danh sách 84 mục chưa tính — **SAI, user chấm cột `cmt` (8/8 sai mẫu)**; không dùng làm đáp án |
| `…\02_takeoff\AI_INPUT.xlsx` | Hiện là dữ liệu đợt 1 (cho REV01). Bản cũ ở `_MASTER\backup\SOFITEL_AI_INPUT_D15_2026-10-24.xlsx` |
| `_TEST\SOFITEL_BATCH1\` | Bản thử đợt 1 (đã đạt khứ hồi). `_TEST\SOFITEL_D15\` bản thử D15 |
Công cụ (`_MASTER\takeoff\`, chạy `py -I`, đặt `sys.stdout.reconfigure(encoding="utf-8")`): `extract_ai_input.py`, `check_ai_input.py`, `roundtrip_check.py`, `gom_bo_to.py`. Bước ghi `L.fromAiInput` chạy bằng **Claude trong Excel của user** (Office.js) — session này không chạy được; user chạy `RefreshAiInput` (VBA) rồi Claude trong Excel.

## 3. Đã làm (xác minh được)
- Scan v8: thêm cột `Bộ tờ`, `Tờ cùng bộ` (PAGES), `Bộ tờ phải đọc`, `Tờ chưa xác nhận đọc` (BOQ_TONGHOP), sheet `CALLOUT` (7 dòng), dòng `Phạm vi tính giá` ở PROJECT (`gom_bo_to.py`, idempotent, đã đếm hyperlink 2237 trước/sau). Hạn chế: chỉ gom được khung tên có "SHEET n OF m" (NORTON không có); gom theo "TYPE 2/3" đang tách nhầm; callout là nét vector, không có trong text layer.
- Đợt 1 (F026.1, F053, F054, F088) đã áp lên **REV01** hàng 8096–8342, `roundtrip_check` đạt 10/10 ô. Nhưng đích sai (xem §5) và F026.1 trùng: ở RE00 ô AH của F026.1 đã trỏ block 349.
- `check_ai_input.py` thêm phép kiểm GIATRI; `extract_ai_input.py` mặc định loại sheet Kitchen khỏi ứng viên đoán mẫu (`CPSX_ALLOW_KITCHEN=1` để bật).
- Đã đọc bản vẽ: ID-PS-TY-710/711/712 (tủ áo Prestige; 710 là tủ chữ L của F067, 711 là F066 1460×590, 712 mặt cắt chung), 715/716 (minibar type 3 = F074), ID-JS-65-705 (TV panel F053).

## 4. Lỗi/sự cố đã gặp (và bài học)
1. **Script đoán mẫu theo chuỗi nhãn là không đáng tin.** Nhãn module dùng chung nhiều mẫu, user đổi tên. Ví dụ F012 (len chân) bị gán "Vách, trần phẳng" 0,863 vs "Len chân" 0,727; Kitchen thắng Template "Cabinet" với biên 0,01–0,03. Hệ quả: mọi cột "Mẫu" do script trích, bảng §1 của `NHAN-DANG-MAU.md`, thống kê "vanity→TRỐNG, tủ áo→Kitchen#19…" đều không đáng tin. Đừng dùng làm đáp án huấn luyện.
2. **Sheet Kitchen chỉ có hiệu lực khi user cho phép dùng Matrix (R0076).** Không bao giờ gán cho hạng mục thường.
3. **`BO_SUNG_RE00_v2`**: gợi ý theo độ giống TÊN hạng mục với hàng xóm đã làm, không có cổng loại đồ vật (cửa swing→vách, tủ áo→vách, F028→gối), vẫn xuất gợi ý ở điểm 0,47–0,57; nhãn "CAO" chỉ là độ giống tên; gợi ý khi chưa đọc bản vẽ. Kết quả: 0/8 đúng mẫu, 4/8 đúng block nguồn.
4. **openpyxl ghi chuỗi bắt đầu `=` thành công thức thật** → Power Query nạp rỗng → E header không được ghi (4 block đợt 1 lần đầu). Phải `cell.data_type='s'`.
5. **Nhầm file gốc**: tôi tự coi REV01 là gốc. Luôn hỏi/nhìn `Chiet tinh` max hàng: RE00 = 5414.
6. `AI_MODULE` chỉ liệt kê ô nhập; dòng thêm/bớt ở `NGOAI_EDITS` → không kết luận "thiếu" từ AI_MODULE (F074 đã bóc đủ, user chốt giữ nguyên).
7. "Đã làm" thật = ô AH của BoQ (`MATRIX FITFUR-LOOSE`!AH) trỏ về H block. RE00: 80/170 mã có AH. F007, F008, F024, F028, F065 chưa có AH.
8. Ghi `PROJECT_DATA` bị `PermissionError` khi user đang mở file Excel (cả `BO_SUNG_RE00.xlsx`).

## 5. Quy tắc user đã xác nhận
- Phạm vi SOFITEL: chỉ tính gỗ ván + ván nền + sơn trên ván. KHÔNG: toàn bộ kim loại (MT-*, chân inox, khung sắt…), đá ST, giấy WC (phần giấy), OS&E (giá ủi, két, tủ lạnh), đèn, kính GL. **Dự án khác mặc định tính kim loại**, chỉ bỏ khi có yêu cầu. (Bản trong PROJECT_DATA lúc cuối là bản đầu, chưa có dòng "mặc định mọi dự án"/kính GL — cần cập nhật khi file không bị khóa.)
- Len chân (F007, F008, F012, F020…) → Template "Len chân, thanh lam gỗ"; chép block F012/F020. Chỉ gỗ/trim/nan gỗ/thanh nan/vách nan che điều hòa (F023, F024, F026.1) → Template TRỐNG; chép F011/F009. Cửa (kể cả F005) → sheet Door. Tủ áo → Template "Tủ áo". Chưa đọc bản vẽ → không gợi ý mẫu (F028 phải mở bản vẽ).
- Minibar type 3 (F074): 2 hông bo R80 = GHÉP HỘP ×2; thùng 3 ngăn kéo + 3 ngăn kéo; thùng 2 cánh + ngăn kéo thêm; len chân gỗ; chân inox bỏ (SOFITEL). Block F074 hiện có: user chốt GIỮ NGUYÊN.
- Bỏ qua tủ áo F065–F067 (user: "bỏ qua tủ áo đó").
- File gốc là RE00; REV01/REV02 chỉ tham khảo.

## 6. Việc tiếp theo (đề xuất, chờ user duyệt câu 24)
1. Gom 84 mục (56 F, 14 L, 14 D; đã trừ F001–F004, F049 loại trừ) thành ~25 **loại đồ vật**; mỗi loại 1 dòng: số mục, mã, block đã làm cùng loại trong RE00 (để chép). Cột "Mẫu" để trống, user điền 1 lần/loại. Không gợi ý khi không có căn cứ.
2. Đợt 2: chỉ làm loại đã được user xác nhận mẫu; đọc `Bộ tờ phải đọc` trước. Dựng `AI_INPUT` → `check_ai_input.py` → bản sao RE00 trong `_TEST` → user chạy `RefreshAiInput` + `L.fromAiInput` → `roundtrip_check.py` → rồi mới áp RE00 gốc (sao lưu trước).
3. Làm lại đợt 1 trên RE00: F053, F054, F088 (F026.1 đã có AH ở RE00 — bỏ).
4. Sửa `gom_bo_to.py` (tách "TYPE 2/TYPE 3"); dựng lại bảng `NHAN-DANG-MAU.md` §1 từ xác nhận của user, không từ script.

## 7. Câu hỏi mở
- Mẫu thật của vanity (F043…), riser door (D04–D06, D19), bedside (L022–L025) — script nói TRỐNG/Cabinet, chưa xác nhận.
- D20–D26 (cửa có sẵn làm lại mặt: giấy WC, nhãn 3M SP-06, veneer; ~290 cái): sau khi loại giấy/nhãn dán còn gì tính?
- F054: BoQ 1500×465 (trùng F030) nhưng tờ JS-65-705 vẽ 1620 → hỏi CĐT. F053: BoQ sâu 270, mặt bằng ~150.
- F066 chia biến thể PS02/PS02A (hoãn vì user bỏ qua tủ áo).
- Quy ước tra loại→mẫu dùng chung mọi dự án hay theo dự án (chưa kiểm NORTON; user bảo bỏ NORTON, tập trung SOFITEL).
- Mục 1 `_index.md` dòng 35 "Chưa làm…" đã cũ (các file BAN, SOFITEL…, CUA-* đã có) — chưa xử lý.
