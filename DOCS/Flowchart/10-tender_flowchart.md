# Flowchart skill 10-tender

> Cập nhật lần cuối: 2026-11-01 · Theo code: `SKILL/10-tender/` (SKILL.md, `buoc/*.md`, `tool/td.py`)
> Trạng thái: 8 bước + `bo_thu_muc` (GĐ0 plan Code ID/DMVT) + `tim --ma`, vai trò Scope (GĐ1) + `phanloai` Code ID (GĐ2, `nho.json` chờ user duyệt) + droplist, `Code ID`, `Bên thực hiện`, highlight ở sheet BOQ (GĐ3) + `CAU_TAO`, `PHAN_VIEC`, `ben`, `goi`, droplist bên theo `_ben.json` (GĐ3b, danh sách bên đang tạm) + `VL_XUATHIEN`, bảng `AREA`, `Link spec` (GĐ4) đã làm. GĐ5–6 của `DOCS/PLAN_TENDER-CODEID-DMVT_2026-10-31.md` (DMVT, gộp pd.py) chưa làm. Main Lobby (BOQ thứ 2) chưa làm.
> Quy tắc cập nhật: xem [README.md](README.md). Lịch sử thay đổi: [CHANGELOG.md](CHANGELOG.md).

Sơ đồ gồm 7 phần: luồng tổng, kiểm kê, chạy lại khi có đợt mới, gộp dữ liệu và build, phân loại Code ID, bên thực hiện và gửi nhà thầu phụ, sổ xuất hiện vật liệu.

# 0. LUỒNG TỔNG (phiên điều phối)

```
 User: "scan tender" / "có TCQ mới" / "build lại tender"
        │
        ▼
┌─────────────────────────────────────────────────────────────────────┐
│ §0 CHUẨN BỊ                                                         │
│  GOC = thư mục dự án trong ACTIVE.md                                │
│  Chưa có _cauhinh.json → init (hồ sơ, BOQ CĐT, TCQ gộp, bản dịch)   │
│  Thư mục không phải hồ sơ thầu → init --bo-thu-muc                  │
│  DỪNG nếu: ~$<DA>_BOQ_TRACK_*.xlsx hoặc ~$FILE_INDEX.xlsx tồn tại   │
└─────────────────────────────────────────────────────────────────────┘
        │
        ▼   Mỗi bước: tool chạy trước, rồi subagent mới ghi JSON của bước.
            Sau bước: check --den <buoc>; LỖI → subagent sửa, tối đa 2 vòng.

  1 kiemke ──► (xong bước) trang: đọc chữ PDF, chỉ file mới/đổi
     │
     ├──────────────┬──────────────┐
     ▼              ▼              ▼
  2 boq          3 tcq          4 dwg
  2b phanloai    ≤ 3 subagent   (sau 1)
  (mục 5)        (sau 1, 2)
     │              │              │
     └──────┬───────┘              │
            ▼                      │
   HỎI USER 1 câu gộp (sau bước 1–3)
            │                      │
            ▼                      │
         5 mat (sau 3)             │
            │                      │
            ├──────────────────────┘
            ▼
         6 banve  mỗi nhóm mã phòng/khu 1 subagent, song song (sau 4, 5)
            │
            ▼
         doichieu (tool: nháp từ TCQ_LOG)
            │
            ▼
         7 doichieu  mỗi 1–2 sheet BOQ 1 subagent (sau 3, 5, 6)
            │
            ▼
         8 build
            │
            ├─► 03_boq\<DA>_BOQ_TRACK_<yymmdd>.xlsx
            ├─► 01_index\FILE_INDEX.xlsx
            ├─► 04_out\<DA>_BOQ_<yymmdd>_CDT.xlsx (không tự gửi)
            └─► 9 goi "<bên>" khi tender cần ─► 04_out\<DA>_GOI_<bên>_<yymmdd>.xlsx (mục 6, không tự gửi)
```

# 1. KIỂM KÊ (td.py kiemke)

```
 _cauhinh.json: bo_thu_muc
        │
        ▼
 Giá trị hợp lệ? ('', '.', '..', đường dẫn tuyệt đối, không phải danh sách)
        ├─ KHÔNG ─► LỖI, dừng trước khi ghi file
        ▼ CÓ
 Duyệt thư mục hồ sơ
   Thư mục con khớp bo_thu_muc (tính từ gốc hồ sơ, không phân biệt hoa/thường)
        ├─ CÓ ─► không đi vào: không file, không zip, không chữ ký
        ▼ KHÔNG
   Mỗi file (bỏ file rác) → liệt kê, zip thì liệt kê cả file bên trong
        │
        ▼
 Mục bo_thu_muc không khớp thư mục nào → XEM (có thể gõ sai tên)
        │
        ▼
 Gán ID (file cũ giữ ID) · đợt · số BV · đoán vai trò
   (Mail > Scope: `pvcv` trong tên file > Làm rõ > Tham khảo/Ngoài scope)
        │
        ├─► FILE_INDEX.kiemke-tool.json
        ├─► _chuky.json
        └─► _thaydoi.json (từ lần 2: file mới / đổi / mất)
```

Subagent bước 1 gán `Vai trò` cho file còn trống, đọc mail từng đợt, ghi `README.md` của `10_TENDER`.

**Tìm trang theo mã (`td.py tim --ma`, dùng ở bước 5, 6):**

```
 Mã nhập ─► ma_chuan (_core) ─ không phải mã ─► LỖI
                │
                ▼  51 - 500B → 51-500B · ST 01 → ST-01 · AREA53 → AREA 53
           regex_ma: mọi cách viết của mã
             FF&E: bắt buộc có gạch, cho phép cách và xuống dòng quanh gạch
             VL:   chữ hoa; gạch, dấu cách hoặc liền; không qua dòng
             AREA: không phân biệt hoa/thường
                │
                ▼
           Trang PDF trong cache `trang` (vai trò Chính, Bổ sung/Thay thế, Làm rõ, Scope, trống)
           In: FID, trang, tên tờ, chữ khớp thực tế (để loại trùng ngẫu nhiên)

 tim <regex> vẫn giữ nguyên; truyền mã trơn thì nhắc XEM dùng --ma.
```

# 2. CHẠY LẠI KHI CÓ ĐỢT MỚI

```
 Đặt thư mục đợt vào hồ sơ
   → init --tcq (nếu đổi) → kiemke → bước 1 (file mới) → trang → tcq
   → bước 3 (đợt mới) → dwg, mat
   → thaydoi: dòng SCOPE/MAT/TCQ bị ảnh hưởng
   → bước 4, 5 (phần mới) → bước 6 (mã phòng thaydoi báo)
   → doichieu → bước 7 (dòng có Mã TCQ mới hoặc thaydoi báo) → build
```

Bước đã có `*.<buoc>*.json` thì bỏ qua, trừ khi user bảo chạy lại. Không xóa `*.user.json`, `_snapshot.json`.

# 3. GỘP DỮ LIỆU VÀ BUILD

```
 Ưu tiên theo ô:  *.<buoc>-tool.json  <  *.<buoc>.json (LLM)  <  *.user.json
        │
        ▼
 build: kéo chỗ user sửa trong xlsx vào *.user.json
        (hàng BOQ bị chèn/xóa → dừng, không nhận sửa tay)
        Ô được kéo ngược từ sheet BOQ: 3 cột chính thức + Mã cấp 3 + Bên thực hiện + Mã cấu tạo
        Sheet CAU_TAO: thêm, sửa, xóa dòng đều nhận
        │
        ▼
 Có SCOPE.phanloai-tool.json → chạy lại phanloai (mục 5)
   Nguồn/Cờ theo chỗ user vừa sửa; ô sửa vào nho_local.json
        │
        ▼
 Chạy lại area, xuathien (mục 7), rồi tính Link spec từng dòng
        │
        ▼
 Dựng xlsx: sheet BOQ CĐT + 3 cột chính thức + cột nội bộ (không wrap, hàng giữ đúng chiều cao gốc)
            Sheet của tool: hàng cao cố định, chữ dài xem bằng cách click
            Cột nội bộ theo thứ tự BOQ_TONGHOP, không ẩn cột nào, không lặp cột đã có trên hàng (Item, Mô tả, ĐV, KL):
                      Mã cấp 3 (droplist, tô màu nếu có cờ) · Code ID (công thức, tra _CODE)
                      · Bên thực hiện (droplist, tô màu nếu có cờ) · Mã cấu tạo (droplist)
                      · Spec · VL (ghép từ MAT, có link) · các cột còn lại (công thức tra SCOPE)
            Cắt khai báo cột của file gốc lấn vào cột thêm (M:O ẩn ở 4. Provisional sum)
            DASHBOARD (đếm highlight, dòng Base chưa có bên, cờ nhiều bên) · BOQ_TONGHOP, PHAN_VIEC (chỉ xem)
            · TCQ_LOG · MAT · DWG · KHU · AREA · VL_XUATHIEN (chỉ xem) · CAU_TAO · SCOPE · _CODE
            _LINKS cho add-in SRC_LINKS (mở PDF đúng trang, không còn cột mở dự phòng)
            Ẩn sheet: SCOPE, DWG, KHU, AREA, VL_XUATHIEN (không dựng hyperlink), _LINKS, _CODE;
                      CAU_TAO, PHAN_VIEC ẩn khi chưa có mẫu. Sửa ở sheet ẩn vẫn được kéo vào
        │
        ▼
 File trong zip được link → giải nén vào _data\_zip, link tương đối
```

# 4. ĐIỂM THIẾT KẾ ĐÁNG CHÚ Ý

- Quét toàn gói, gắn nhãn chứ không bỏ dòng; không áp quy tắc "phạm vi nhà máy 5 xưởng" của 01-scan.
- Đợt sau đè đợt trước; `FILE_INDEX` giữ ID file cũ khi chạy lại.
- `bo_thu_muc` là đường dẫn đầy đủ từ gốc hồ sơ, nên `HSMT/New folder` không bị loại khi cấu hình chỉ ghi `New folder`.
- Kiểm thử: `DOCS/scan-toi-uu/tender-test/kiemthu.py` (`--mau`, `--sofitel`).
- Code ID chỉ nói hạng mục thuộc nhóm nào, không suy ra ai làm. Bên thực hiện là chiều riêng, nằm ở `SCOPE.ben.json` chứ không ở file doichieu, để chạy lại bước 7 không làm mất.
- `Ngày cập nhật` không tính các cột phân loại, Code ID, Bên thực hiện, Mã cấu tạo, Link spec (đổi danh sách bên không làm dòng BOQ hiện như vừa đổi); cột mới chưa có trong snapshot cũ cũng không tính, nên thêm cột không làm mọi dòng BOQ hiện như vừa đổi.
- Bản `_CDT` không có cột nội bộ, droplist, `_CODE`.

# 5. PHÂN LOẠI CODE ID (td.py phanloai, SKILL/_core/code_id.py)

```
 Dữ liệu: _MASTER\code_id\ bang_ma.json · luat.json · nho.json (chung) · nho_local.json (của máy)
 Dòng xét: có đơn vị, sheet không khớp codeid_bo_sheet (mặc định bỏ prelim, provisional, general req, summary)
        │
        ▼
 Mỗi dòng, dừng ở lớp đầu tiên có kết quả:
   1 User         SCOPE.user.json ──► ghi thêm vào nho_local.json (trừ KHÔNG RÕ)
   2 Luật         luat.json: mục lớn, tiền tố mã (FF&E đọc cả Mục phụ)
   3 Nhớ-mô-tả    mục lớn + mô tả, số thành # (bỏ qua dòng chỉ có kích thước). Local trước, chung sau
   4 Nhớ-đầu-mục  mục lớn + đầu mục + họ mã của dòng. Local trước, chung sau
   5 LLM          _phanloai_llm.json, dùng lại cho dòng cùng đầu mục và mô tả trong dự án
   6 Chờ LLM      subagent làm buoc/2b-phanloai.md rồi chạy lại phanloai
        │
        ▼
 SCOPE.phanloai-tool.json: Mã cấp 3 · Nguồn phân loại · Cờ phân loại
 Cờ (highlight): Nhớ-đầu-mục · LLM · KHÔNG RÕ · Chờ LLM · Dự án khác
                 · Luật chưa duyệt (hết khi chạy phanloai --duyet-luat)

 Nạp bộ nhớ chung (người quản lý, DOCS/scan-toi-uu/tender-test/nap_codeid.py):
   bang ──► bang_ma.json
   deduyet ──► 02_extract\NHO_DUYET.xlsx (dòng nghi ngờ) ──► user điền Mã chốt
   ghi ──► nho.json (dòng nghi ngờ để trống Mã chốt thì không vào)
```

# 6. BÊN THỰC HIỆN VÀ GỬI NHÀ THẦU PHỤ

```
 Nguồn Bên thực hiện (không suy từ Code ID):
   user (sheet BOQ/SCOPE) > SCOPE.ben.json (chữ gốc đã chuyển) > ben-tool (nháp: Mã VL → Loại MAT → bên, cờ Nháp)
        │
        ▼
 build / derive: chuẩn hóa theo _data\_ben.json (gop: chữ cũ → tên trong ds)
        │
        ├─ Dòng có Mã cấu tạo ─► Bên = gộp từ sheet CAU_TAO (sửa mẫu 1 lần, mọi dòng trỏ tới đổi theo)
        │                        Sửa Bên ở dòng ─► không tác dụng, XEM
        │
        └─ Không có, trạng thái Base/VE/Chờ:
             có plywood/MDF, framework, metal frame, consisting of,
             hoặc ≥ 2 mã VL khác bên, hoặc chữ gốc gộp nhiều bên
               ─► Cờ bên = Nhiều bên? (máy không tự tạo mẫu)

 goi <GOC> "<bên>"
   Lấy dòng Base/VE/Chờ có phần việc của đúng bên đó
   Có mẫu: từng thành phần · 1 bên: Cả dòng · nhiều bên không mẫu: Một phần của dòng
   Mô tả NT – Base nếu có, không thì Mô tả BOQ; bỏ mã TCQ
        │
        ▼
 Workbook mới 1 sheet ─► máy đọc lại và soát
   Code ID, mã TCQ, link, đường dẫn nội bộ, 3 cột NTP đã điền ─► LỖI, xóa file
   Mô tả nhắc tên bên khác ─► XEM
```

# 7. SỔ XUẤT HIỆN VẬT LIỆU VÀ BẢNG AREA

```
 area <GOC>
   KHU, chỉ Bộ hồ sơ FF&E + Thảm, đọc số AREA trong tên khu
     ─► AREA: tiền tố mã → tập phòng (nháp; sửa ở sheet AREA, user thắng)
   Bộ IDD không vào bảng: IDD Area 51 = sảnh thang, FF&E AREA 51 = Standard King

 xuathien <GOC>
   Danh sách mã của dự án = MAT + SCOPE (cả Mục phụ, mã viết cách) + TCQ + trang spec
        │
        ▼
   Bỏ file: bản cũ, tờ Ngoài scope/Tham khảo, file tổng, trang trùng, PDF TCQ (TCQ đọc từ TCQ_LOG)
        │
        ▼
   Từng trang PDF, từng mã trong danh sách
     Viết liền (SW10, ST09) ─► bỏ, trừ khi là mã của chính trang spec
     Trùng mã phòng (ST-01) ─► chỉ nhận khi là trang spec của mã, hoặc có chữ vật liệu cạnh đó
     Loại trang:
       Item: / ITEM / DRAWING REFERENCE: ─► Định nghĩa
       ITEM SCHEDULE ─► Danh mục · Image (s) provided ─► Ảnh · còn lại ─► Nhắc tới
        │
        ▼
   Dòng BOQ ─► nguồn BOQ (ô); phòng ngoài tiền tố ─► Cờ Lệch Area
   Bản vẽ hiệu lực của dòng ─► nguồn Bản vẽ, Cờ "Qua dòng BOQ" (tag trên bản vẽ là nét, không có trong chữ)
   Câu TCQ ─► nguồn TCQ; câu đã bị thay có cờ
        │
        ▼
   Hiệu lực theo từng mã: trang Định nghĩa ngày mới nhất = Mới nhất, còn lại Bản cũ
   Khu, Mã phòng lấy theo file (DWG, link Nguồn của KHU), không đọc từ chữ

 Link spec (build, từng dòng BOQ)
   Từng mã của dòng: MAT Nguồn spec nếu có, không thì trang Định nghĩa Mới nhất
```
