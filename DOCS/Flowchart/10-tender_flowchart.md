# Flowchart skill 10-tender

> Cập nhật lần cuối: 2026-11-03 (2a BOQ ra bản mới; 2 cập nhật theo đợt: kiemke cộng dồn, tcq ghi câu mới/đổi, giao lại bước 6, TCQ phạm vi theo phòng/khu; vẽ lại phụ thuộc phần 0) · Theo code: `SKILL/10-tender/` (SKILL.md, `buoc/*.md`, `tool/td.py`, `tool/lap_boq.py` và các tool phần 9, `tool/docsau.py`)
> Trạng thái: 8 bước + `bo_thu_muc` (GĐ0 plan Code ID/DMVT) + `tim --ma`, vai trò Scope (GĐ1) + `phanloai` Code ID (GĐ2, `nho.json` chờ user duyệt) + droplist, `Code ID`, `Bên thực hiện`, highlight ở sheet BOQ (GĐ3) + `CAU_TAO`, `PHAN_VIEC`, `ben`, `goi`, droplist bên theo `_ben.json` (GĐ3b, danh sách bên đang tạm) + `VL_XUATHIEN`, bảng `AREA`, `Link vật liệu` (GĐ4) + `dmvt`, `Link vật tư` (GĐ5, mục nhóm Finishes, Sanitary đang tạm) + `chu`, `anh`, `mail`, `xem` qua `_core/doc_file.py` (GĐ6) đã làm + đọc sâu toàn gói chuyển từ 01-scan (`docsau`, `gemini`, `neo`, `tincay`, `so`, `do`; phần 10) đã làm, kiểm thử bằng dữ liệu mẫu và thử lệnh trên bản sao CTX, các bước LLM 7c–7e chưa chạy trên dự án thật + dự án chưa có BOQ CĐT (mẫu `tu_lap_vn`, `lap_boq.py`; phần 9) đã build thật CTX, dòng cửa lấy vật liệu từ tag tờ cửa, `cua_ma_khung` (02/11, chưa commit) + add-in `SRC_LINKS` mở được khi chuyển file, mở từ SharePoint, máy khác cài bằng xlam rời. Main Lobby SOFITEL (BOQ thứ 2) chưa làm, chờ user quyết.
> Quy tắc cập nhật: xem [README.md](README.md). Lịch sử thay đổi: [CHANGELOG.md](CHANGELOG.md).

Sơ đồ gồm 11 phần (0–10): luồng tổng, kiểm kê, chạy lại khi có đợt mới, gộp dữ liệu và build, điểm thiết kế, phân loại Code ID, bên thực hiện và gửi nhà thầu phụ, sổ xuất hiện vật liệu, danh mục vật tư, dự án chưa có BOQ CĐT, đọc sâu toàn gói.

# 0. LUỒNG TỔNG (phiên điều phối)

```
 User: "scan tender" / "có TCQ mới" / "build lại tender"
        │
        ▼
┌─────────────────────────────────────────────────────────────────────┐
│ §0 CHUẨN BỊ                                                         │
│  GOC = thư mục dự án trong ACTIVE.md                                │
│  Chưa có _cauhinh.json → init (hồ sơ, BOQ CĐT, TCQ gộp, bản dịch)   │
│  init --mau <tên>: mẫu dự án trong form/MAU_DU_AN.json              │
│    cdt_item (mặc định): có BOQ CĐT, TCQ. tu_lap_vn: xem phần 9      │
│    cdt_vn: BOQ CĐT tiếng Việt, 1 sheet = 1 loại căn, mã từ MÃ SP;   │
│      _cauhinh: rx_phong, phong_theo_sheet, dwg_theo_trang (DWG theo  │
│      từng trang PDF), mat cờ KHÁC NGHĨA THEO LOẠI CĂN               │
│  Thư mục không phải hồ sơ thầu → init --bo-thu-muc                  │
│  DỪNG nếu: ~$<DA>_BOQ_TRACK_*.xlsx hoặc ~$FILE_INDEX.xlsx tồn tại   │
└─────────────────────────────────────────────────────────────────────┘
        │
        ▼   Mỗi bước: tool chạy trước, rồi subagent mới ghi JSON của bước.
            Sau bước: check --den <buoc>; LỖI → subagent sửa, tối đa 2 vòng.
            Còn LỖI sau 2 vòng → ghi vào báo cáo, làm tiếp (LỖI "thiếu cột" của bước SAU là bình thường).
            Đợt đầu chưa có TCQ: gần như mọi dòng là Base → báo user số nhóm 7d trước khi giao đọc sâu.

  1 kiemke ──► (xong bước) trang: đọc chữ PDF, chỉ file mới/đổi
     │
     ├─────────────────────────────┐
     ▼                             ▼
  2 boq                         4 dwg (sau 1, song song 2, 3)
  2b phanloai (mục 5)              │  Ngoài scope lấy từ TCQ_LOG/README lúc chạy
     │                             │
     ▼                             │
  3 tcq  ≤ 3 subagent (sau 2: Dòng BOQ lọc từ SCOPE)
     │                             │
     └──────────────┬──────────────┘
                    ▼
   HỎI USER 1 câu gộp (câu hỏi của bước 1–4)
     user trả lời → README "Việc chưa quyết", *.user.json; vá Trạng thái DWG/KHU của bước 4
                    │
                    ▼
                 5 mat (sau 3)
                    │
     ┌──────────────┼───────────────────────────┐
     ▼              ▼                           ▼
  6 banve        8 dmvt  bóc A:I + ảnh,      7b ben  dự án có
  mỗi nhóm mã    LLM dịch, chọn ảnh ca khó   _data\_ben.json: nháp
  phòng/khu 1    (mục 8; chỉ cần 5; Area     bên, tender duyệt (mục 6;
  subagent,      lấy Mã VL dòng BOQ, build   build tự chạy lại ben)
  song song      dựng lại nên 6, 7 sửa sau
  (sau 4, 5)     vẫn đúng)
     │              │                           │
     ▼              │                           │
  doichieu (tool: nháp từ TCQ_LOG theo Dòng BOQ, Mã VL, câu phạm vi theo phòng/khu)
     │              │                           │
     ▼              │                           │
  7 doichieu  mỗi 1–2 sheet BOQ 1 subagent (sau 3, 5, 6)
     │              │                           │
     ▼              │                           │
  docsau (tool: gom cụm, chia nhóm ≤ 8 cụm)     │
     ├─► 7c spec    1 subagent (sau 5, 7)       │
     ├─► 7d docsau  mỗi nhóm 1 subagent, song song (sau 6, 7)
     │      → gemini --cum (người đọc thứ 2) → neo: NEO LỆCH gửi lại nhóm, ≤ 2 vòng
     ▼              │                           │
  7e tincay  tool nháp rồi 1 subagent (sau 7c, 7d; mục 10)
     │              │                           │
     └──────────────┴───────────────┬───────────┘
                                    ▼
         9 build  (luôn dựng DMVT trước để có số hàng; không dựng được thì vẫn ra BOQ_TRACK nhưng thoát lỗi)
            │
            ├─► 03_boq\<DA>_DMVT_<yymmdd>.xlsx
            ├─► 03_boq\<DA>_BOQ_TRACK_<yymmdd>.xlsx (cột Link vật tư trỏ dòng DMVT; bản cùng ngày tự sao _backup)
            │     tra_part_goc: drawing ghép theo TÊN SHEET + chép ảnh gốc (goc_*), ảnh trong ô
            │     (metadata + richData, thuộc tính vm) chép lại — openpyxl đổi tên/bỏ các phần này
            ├─► 01_index\FILE_INDEX.xlsx
            ├─► 04_out\<DA>_BOQ_<yymmdd>_CDT.xlsx (không tự gửi)
            └─► 10 goi "<bên>" khi tender cần ─► 04_out\<DA>_GOI_<bên>_<yymmdd>.xlsx (mục 6, không tự gửi)
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
        ├─► _thaydoi.json (từ lần 2: file mới / đổi / mất, CỘNG DỒN các lần cùng đợt)
        │     --dot-moi (lần đầu của đợt mới): danh sách cũ → _dot\_thaydoi_dot_truoc_<giờ>.json
        └─► _dot\kiemke_<giờ>.json (từng lần, không ghi đè)
              tcq_moi / tcq_doi trong _thaydoi.json (do lệnh tcq ghi) được giữ qua các lần kiemke
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
   → init --tcq (nếu đổi)
   → kiemke --dot-moi (lần đầu của đợt; chạy lại trong đợt thì không thêm cờ, danh sách cộng dồn)
   → bước 1 (file mới) → trang
   → tcq: so với TCQ_LOG.tcq-tool.json cũ → câu mới / đổi câu hỏi, trả lời → _thaydoi.json (tcq_moi, tcq_doi)
   → BOQ ra bản mới? init --boq <file mới> (tu_lap_vn: lap_boq.py tự trỏ) → boq (xem 2a)
   → bước 3 (đợt mới) → dwg, mat
   → thaydoi ─► _thaydoi.json anh_huong:
        SCOPE: link tới file đổi/mất · Số BV có bản mới · dòng BOQ mới / đổi nội dung (_dot\boq_*.json)
               · dòng đã nối với TCQ mới/đổi
        MAT:   nguồn spec đổi/mất · mã được TCQ mới/đổi nhắc (cột Mã VL hoặc chữ câu hỏi/trả lời, regex_ma)
        TCQ:   mới/đổi của đợt + chưa phân loại
   → bước 4 (phần mới), bước 5 (mã MAT thaydoi báo, kể cả khi đã có JSON)
   → bước 6 giao lại (mã phòng thaydoi báo): SCOPE.banve-<n lớn hơn nhóm cũ>.json + "giao_lai": true
        dòng file chạm tới: bỏ kết quả nhóm banve cũ (ô về nháp tool) rồi ghi → tờ bỏ, mã VL đổi không sót
        dòng dùng chung nhiều mã phòng: nhóm mới ghi đủ cả dòng; check báo XEM
   → doichieu: nháp Mã TCQ / Trạng thái
        Dòng BOQ của câu · trùng Mã VL (khác phòng thì bỏ)
        câu phạm vi chỉ có Mã phòng/Khu (khu → mã phòng qua sổ KHU), không Dòng BOQ / Mã VL
          → mọi dòng của phòng đó; Trạng thái nháp chỉ khi mọi phòng của dòng nằm trong câu
          → dòng dùng chung phòng khác: chỉ Mã TCQ, XEM; câu không khớp dòng nào: XEM
        dòng nối với TCQ mới/đổi → thêm vào anh_huong.SCOPE
   → bước 7 (dòng trong anh_huong.SCOPE) → docsau --thaydoi → 7c, 7d (cụm bị ảnh hưởng) → tincay → 7e
   → build
```

Bước đã có `*.<buoc>*.json` thì bỏ qua, trừ khi user bảo chạy lại. Không xóa `*.user.json`, `_snapshot.json`.

## 2a. BOQ ra bản mới: chuyển kết quả theo nội dung dòng

```
 Khóa dòng = '<sheet>!<hàng>' (vị trí). BOQ chèn/xóa hàng → khóa cũ trỏ sang dòng khác.
 boq chạy lại, có SCOPE.boq-tool.json cũ
        │
        ▼
 Khớp dòng cũ ↔ mới (theo từng sheet, giữ thứ tự)
   giống:      Vùng + Nhóm + Mục phụ + Mô tả BOQ + ĐV trùng (không xét Item, KL)
   gần giống:  ở đúng đoạn bị sửa, cùng ĐV, Mô tả giống ≥ 85%
   khác sheet: còn lại, nội dung trùng, cùng số lần ở hai bên
   còn lại:    cũ không còn / dòng mới
        │
        ├─ boq --thu: chỉ in, không ghi
        ▼
 Đổi khóa trong mọi *.json của _data (trừ SCOPE.boq-tool, _boq_map, _snapshot, _chuky,
   _thaydoi, _lap_boq_dong), kể cả _gemini_so\<cụm>.json (đổi cả tên file)
   bản gốc → _data\_dot\boq_<giờ>\
   dòng có Khóa của dòng đã xóa → bỏ, lưu _dot\boq_<giờ>\_dong_bo.json
   khóa đã xóa ở cột khác (TCQ_LOG Dòng BOQ…) → "(đã xóa: <sheet> dòng <n>)", check báo LỖI để sửa
   _loc theo Khóa/Dòng/Cụm đọc → XEM, sửa tay
        │
        ▼
 _dot\boq_<giờ>.json: bản đồ khóa, dòng xóa, dòng mới, dòng đổi nội dung (cột nào)
 _boq_cho_build.json: bản đồ từ khóa của lần build trước (ghép nếu boq chạy nhiều lần trước build)
        │
        ▼
 build: pull_user đổi khóa sửa tay trong BOQ_TRACK cũ theo bản đồ; sửa ở dòng đã xóa → XEM, không nhận
        Ngày cập nhật so với chính dòng đó ở lần build trước; xong → chuyển _boq_cho_build vào _dot
```

# 3. GỘP DỮ LIỆU VÀ BUILD

```
 Ưu tiên theo ô:  *.<buoc>-tool.json  <  *.<buoc>.json (LLM)  <  *.user.json
 Cùng bước nhiều nhóm: xếp theo số nhóm (banve = nhóm 0 < banve-2 < banve-10), file sau thắng;
   ô LINK / cột mã của nhóm khác cùng bước: GỘP (hợp link, hợp mã)
   file có "giao_lai": true: dòng nó chạm tới bỏ kết quả các nhóm trước cùng bước, không gộp
        │
        ▼
 build: kéo chỗ user sửa trong xlsx vào *.user.json
        (hàng sheet BOQ trong BOQ_TRACK bị user chèn/xóa → dừng, không nhận sửa tay;
         BOQ CĐT ra bản mới thì đi đường 2a, không dừng)
        Ô được kéo ngược từ sheet BOQ: 3 cột chính thức + Mã cấp 3 + Bên thực hiện + Mã cấu tạo
        Sheet CAU_TAO: thêm, sửa, xóa dòng đều nhận
        │
        ▼
 Có SCOPE.phanloai-tool.json → chạy lại phanloai (mục 5)
   Nguồn/Cờ theo chỗ user vừa sửa; ô sửa vào nho_local.json
        │
        ▼
 Chạy lại area, xuathien (mục 7), rồi tính Link vật liệu từng dòng
        │
        ▼
 Dựng xlsx: sheet BOQ CĐT + 3 cột chính thức + cột nội bộ (không wrap, hàng giữ đúng chiều cao gốc)
            Sheet của tool: hàng cao cố định, chữ dài xem bằng cách click
            Cột nội bộ theo thứ tự BOQ_TONGHOP, không ẩn cột nào, không lặp cột đã có trên hàng (Item, Mô tả, ĐV, KL):
                      Mã cấp 3 (droplist, tô màu nếu có cờ) · Code ID (công thức, tra _CODE)
                      · Bên thực hiện (droplist, tô màu nếu có cờ) · Mã cấu tạo (droplist)
                      · Bản vẽ hiệu lực (dòng FF&E: thêm trang ảnh sản phẩm, trang tập FF&E chuyển sang Link vật liệu)
                      · Link vật liệu (1 cột thay Mã VL, Spec · VL, Link spec, Mã spec/FF&E; trỏ trang gốc, không qua MAT)
                      · các cột còn lại (công thức tra SCOPE)
            Cắt khai báo cột của file gốc lấn vào cột thêm (M:O ẩn ở 4. Provisional sum)
            DASHBOARD (đếm highlight, dòng Base chưa có bên, cờ nhiều bên) · BOQ_TONGHOP, PHAN_VIEC (chỉ xem)
            · TCQ_LOG · MAT · DWG · KHU · AREA · VL_XUATHIEN (chỉ xem) · CAU_TAO · SCOPE · _CODE
            _LINKS cho add-in SRC_LINKS (mở PDF đúng trang, không còn cột mở dự phòng; H1:H2 BASE_DIR = thư mục file lúc build)
              Add-in tìm đích: file mở từ https → đổi sang thư mục OneDrive đồng bộ trên máy
                (registry SyncEngines; dự phòng %OneDriveCommercial%/%OneDrive% khi registry bị chặn)
                → tương đối theo chỗ đặt file → theo BASE_DIR → bỏ ..\ rồi thử từng thư mục cha
              File mở từ https mà máy chưa đồng bộ → mở thẳng địa chỉ web
              Máy khác: chép SRC_LINKS.xlam vào %APPDATA%\Microsoft\Excel\XLSTART
                (gói _MASTER\vba\GUI_CHO_NGUOI_KHAC: xlam + HUONG DAN CAI.txt; Cortex chặn file .bat tự giải mã)
                Bản cũ (trước 02/11) báo "Khong tim thay file: C:\...\Documents\https:\..." → thay xlam mới
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
- Các bước LLM đọc file qua `td.py chu`, `anh`, `mail`, `xem` (`_core/doc_file.py`); 10-tender không gọi tool của 01-scan.
- Kiểm thử: `DOCS/scan-toi-uu/tender-test/kiemthu.py` (`--mau`, `--sofitel`).
- Code ID chỉ nói hạng mục thuộc nhóm nào, không suy ra ai làm. Bên thực hiện là chiều riêng, nằm ở `SCOPE.ben.json` chứ không ở file doichieu, để chạy lại bước 7 không làm mất.
- `Ngày cập nhật` không tính các cột phân loại, Code ID, Bên thực hiện, Mã cấu tạo, Link vật liệu (đổi danh sách bên không làm dòng BOQ hiện như vừa đổi); cột mới chưa có trong snapshot cũ cũng không tính, nên thêm cột không làm mọi dòng BOQ hiện như vừa đổi.
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

 Link vật liệu (build, từng dòng BOQ)
   Từng mã của dòng: MAT Nguồn spec nếu có, không thì trang Định nghĩa Mới nhất
   Dòng có mã FF&E:
     Cột bản vẽ = trang Ảnh của mã, cùng tập với trang Định nghĩa Mới nhất + tờ có Số BV
     Link vật liệu thêm trang tập FF&E (file không có Số BV) mà bước banve đã ghi
     Bản vẽ hiệu lực gốc không đổi (goi đọc số BV, Ngày cập nhật so theo cột này)
     DASHBOARD đếm dòng FF&E không có link sản phẩm
   Thứ tự mã trong ô = thứ tự mã trong Mô tả BOQ (cửa: vật liệu trước, phụ kiện sau); mã không có trong mô tả xếp cuối
   1 trang chỉ 1 link
```

# 8. DANH MỤC VẬT TƯ (td.py dmvt → dmvt.py, dmvt_xlsx.py)

```
 _dmvt.json (không bắt buộc: tiêu đề, tiền tố FF&E, item_spec, mục I, II…, chon_anh)   MAT, AREA, SCOPE, TCQ_LOG
        │                                                                     │
        ▼                                                                     │
 Quét trang spec, bản mới nhất THEO TỪNG MÃ                                   │
   MATERIAL REFERENCE → từng khối ITEM (khối này thắng trang SPEC của cùng mã)│
   DRAWING REFERENCE  → đọc cột theo đường kẻ bảng                            │
   SPECIFICATION      → mã khớp item_spec, bỏ Fabrication/Installation        │
 + spec dạng bảng xlsx (09/11) cho mã chưa có trang PDF: file có cột mã + mô tả│
   (Material Schedule: loại, mô tả, mã NCC, NCC, ảnh cột Minh họa cùng hàng)  │
   BOQ CĐT xét sau cùng: mã hạng mục (MÃ SP + HÌNH ẢNH) → mục FF&E, có cờ     │
   _dmvt.json "spec_xlsx": [FID] chọn file, false tắt                         │
        │                                                                     │
        ▼                                                                     │
 Gộp dòng: FF&E theo Design + Colour (cùng mã màu, khác vài chữ: gộp, có cờ)  │
           phụ kiện theo nhà cung cấp; mã khác 1 dòng/mã;                     │
           mã có ở MAT, không có spec: vẫn có dòng, có cờ                     │
        │                                                                     │
        ▼                                                                     ▼
 Trạng thái, Area ◄─ Omit: tô xám. Đổi theo TCQ/VE: giữ spec, Remarks + cờ.
                     Chưa có ở MAT: cờ, không suy. Area luôn có cờ.
        │
        ▼
 Ảnh: khối ITEM / vùng Sample Swatch / trang minh họa (ca khó) / ô xlsx cùng hàng
      nhiều ảnh: ảnh chụp trước rồi lớn nhất; dòng có cờ vào _SOI_ANH
        │
        ├─► --chi-boc: _dmvt_boc.json, _dmvt_anh\  ──► bước 8: LLM ghi _dmvt_llm.json
        │
        ▼
 Bản DMVT trước + _dmvt_snapshot.json → gop_sua_tay
   ô A:I / ảnh khác snapshot: giữ; spec đổi: ghi Spec mới khác, không đè
   J:AF, từ AK: giữ. Dòng không mã: đi theo dòng có mã phía trên
   (cột B mọi phần có dạng mã — FR1, LM4A, D-WC — là dòng có mã; trước 09/11 bị coi
    là dòng đuôi nên nhân đôi mỗi lần build)
   tập mã đổi: cờ, không tự chuyển; dòng cũ có công người: mục DÒNG CŨ GIỮ LẠI
        │
        ▼
 Excel COM từ form\DMVT_form.xlsx: ghi file tạm, chiều cao hàng do Excel đo (dư 2 dòng),
   chèn ảnh sau cùng, mở lại kiểm
   mỗi ảnh gọn trong 1 ô cột H, đủ số ảnh ── sai ──► LỖI, file cũ giữ nguyên
        │ đúng
        ▼
 03_boq\<DA>_DMVT_<yymmdd>.xlsx (sheet DMVT, _SOI_ANH, _LINKS ẩn)
 _dmvt_snapshot.json, _dmvt_hang.json (mã → hàng) ──► build: cột Link vật tư
 Không tự gửi. Nhắc xóa: cột từ AG, _LINKS, _SOI_ANH; xem L, AB:AC.
```


# 9. DỰ ÁN CHƯA CÓ BOQ CĐT (mẫu tu_lap_vn, lap_boq.py)

```
init --mau tu_lap_vn (--boq, --tcq không bắt buộc) → kiemke → trang
        │
        ▼
_core/dwg_tag.py <CAD>        DWG → DXF (ODA) → _data/_dwg_tag.json
  tag theo từng tờ: block có thuộc tính, block lồng, MULTILEADER (tag mặt đứng)
  tag Model gán về tờ theo viewport; ngoài mọi viewport: không đếm
  mũi tên, viewport, goc (gốc khung tên, để đổi tọa độ sang ảnh)
spec_boc.py <pdf...>          spec kiểu AID → _data/_spec_boc.json
  mã ghép với khối theo vị trí trên trang schedule; trang chi tiết bổ sung trường
  cột chữ / cột NCC đo theo trang (09/11; trước cố định x 80/255, file SW lệch → mất cả khối)
  khối NCC không nhãn (tên, URL, T/E/M/C) đọc theo vị trí: schedule = cột giữa,
  trang chi tiết = cột phải giữa tiêu đề SPECIFICATION và ô CODE; có nhãn Manufacturer
  mà cột phải ghi tên khác → supplier
        │
        ▼
be_mat.py                     bề mặt từng tag: tuong, tran, san, cua, portal, joinery, khac
  portal (lỗ mở không cánh) → dòng "Ốp portal lỗ mở"
  san, khac → sheet Ghi chú; có khai san_tinh, chi_tiet, ngoai_tick trong _lap_boq.json → có dòng
  block trần → nhãn tag → tên khung hình → còn lại can_soat (không suy từ loại tờ)
chuan_bi_soat.py <GOC>        ảnh khoanh đỏ tag can_soat → _data/_soat/
  người/LLM soát ──► _data/_soat_tag.<n>.json (thắng kết quả máy)
mau_soat.py <GOC> tao | so    soát lại độc lập → _soat2/kq.json, đo tỉ lệ lệch
soi_tag.py                    vẽ 1 tờ để xem riêng
        │
        ▼
lap_boq.py <GOC>              cấu hình riêng dự án: _data/_lap_boq.json; tự trỏ boq trong _cauhinh.json sang file vừa ghi
  phạm vi tick (cột X, gói) + tag theo khu + spec + lịch cửa
  Dòng cửa (vl_cua_tu_tag): mô tả "Rộng x Cao mm. Vật liệu: <mã>. Phụ kiện: <HD>"
    → Mã VL, Link vật liệu, Link vật tư có đủ gỗ, kim loại, kính, phụ kiện
    Vật liệu = tag trên ô cửa của tờ lịch (chữ mã cửa cỡ lớn gần nhất bên trái, cùng hàng)
             + mặt cắt tờ chi tiết (cột khung hình; bỏ tag mặt bằng chi tiết vì chỉ cả tường hai bên)
             DOOR 10 không chữ cái → áp cho D10a, D10b. Sắp WD, SP, WC, UP, MT, GL, ST
    Bộ HD = ghi chú lịch; lịch không ghi → tag mặt đứng
    Nguồn tờ của dòng thêm các tờ cửa có tag
    Cờ [XEM]: mã gỗ, bộ HD lệch giữa 2 tờ; mã chưa chốt (WD-XX, WD-0x); xếp cửa gỗ mà tag không có mã gỗ;
              không có tag vật liệu; bộ HD lấy từ mặt đứng
    Máy không phân biệt tag chỉ cánh hay khung → khai cua_ma_khung (mã cửa → mã → ghi chú) trong _lap_boq.json:
      mã rời khỏi dòng, ghi chú vào cờ
  → 0.BOQ_TU_LAP/<DA>_BOQ_tu_lap_<yymmdd>.xlsx, chép từ form của user
  Mỗi sheet: công tác (A) → khu (A.1) → mã. Cột Ghi chú: gói, nguồn tờ, cờ [XEM]
  Kích thước cột E, F → BOQ_TONGHOP qua cột Kích thước BoQ
  Sheet "Ghi chú & ngoài phạm vi": mã có trên bản vẽ nhưng không được tick
        │
        ▼
init --boq → kiemke → boq (đọc tiêu đề theo mẫu) → tcq (rỗng nếu không có) → dwg → mat
        │
        ▼
lap_boq_buoc.py <GOC>         ghi MAT, DWG, KHU, SCOPE.banve, SCOPE.doichieu, _phanloai_llm.json
phanloai → check (cấu trúc)
        │
        ▼
kiem_boq.py <GOC>             kiểm đúng sai, độc lập với lap_boq.py → _kiem_boq.json, VIEC_MO
  K1 tờ link có mã không · K2 không nguồn · K3 PDF khớp số tờ · K4, K5 sót
  K6 đếm lại · K7 kích thước · K8 mã trùng
  check, kiem_boq không đọc xlsx đã build → trước khi giao đọc lại BOQ_TONGHOP, xem ảnh dòng có cờ của DMVT
        │
        ▼
build (dmvt: _dmvt.json "nguon_spec": "spec_boc" → nội dung và ảnh từ _spec_boc.json,
       tách dòng theo nhãn, mục cửa, mã hoàn thiện, kích thước; bỏ chữ rác layout PDF)
```

Chưa làm: diện tích ốp chưa đo. Mã FF&E trùng giữa 2 file LF: MAT, DMVT chỉ lấy 1 bản.

# 10. ĐỌC SÂU TOÀN GÓI (td.py → tool/docsau.py, _core/gemini_doc.py, _core/neo.py)

Chuyển từ 01-scan. `pd.py` không đổi (đóng băng), nên các hàm này có 2 bản; `kiemthu.py` so hàm tách số, tách mã, câu hỏi Gemini và lệnh `so` với `pd.py`.

```
 SCOPE sau bước 7
   │  bỏ: sheet codeid_bo_sheet, dòng không đơn vị, Trạng thái Omit / Loại trừ / Không áp dụng
   ▼
 docsau  gom dòng cùng (mô tả, Mã VL, Mã spec/FF&E, bộ tờ Bản vẽ hiệu lực) thành CỤM; khóa cụm = Khóa dòng đầu
   │     ghi SCOPE.docsau-tool (Cụm đọc, nhãn Phạm vi nhà máy), BOQ_MAP.docsau-tool (mỗi dòng × khu)
   │     _docsau.json, _docsau.txt (nhóm ≤ 8 cụm); lọc phần giao: --phong --sheet --thaydoi --thieu
   ▼
 7c spec   SPEC, SCOPE (Spec · ML, Thông số riêng), QUAN_SAT nguồn SPEC, ISSUES S…
   ▼
 7d docsau  mỗi nhóm 1 subagent
   │    Claude đọc crop, ghi QUAN_SAT nguồn HÌNH/TEXT TRƯỚC
   │         ▼
   │    gemini --cum  dừng nếu chưa có Kích thước BV của Claude ──► _gemini_so\<cụm>.json
   │    gemini --quansat  so từng trường, lệch: đọc lại; vẫn lệch: 2 dòng, MÂU THUẪN, issue
   │    agy lỗi / không có: KHÔNG THẤY, làm tiếp, cụm tối đa TB
   │         ▼
   │    BOQ_MAP (KL theo BV, tờ đếm), DAY_TUONG (cửa), PAGES, ISSUES V…
   │         ▼
   │    neo  số tờ, mã spec, mã VL, kích thước có trên trang được link không ──► _neo.json (chỉ XEM)
   ▼
 tincay (tool) so số Claude↔Gemini, KL BOQ↔BV, KHÔNG THẤY, MÂU THUẪN
   │    SCOPE/BOQ_MAP Độ tin cậy, ISSUES DT…, QUAN_SAT Kiểm chéo Gemini số, _tincay.json
   ▼
 7e tincay (LLM) Mã hoàn thiện, Giả định tính giá, Hành động/RFI, RFI_LOG, ASK_USER; ô LLM ghi thắng nháp tool
   ▼
 build  merge: dòng cùng cụm nhận kết quả của dòng đại diện, trừ ô có giá trị riêng
        derive: KL theo BV = tổng các khu (còn khu chưa đếm thì để trống), Lệch KL
        cột đọc sâu ở BOQ_TONGHOP + sheet BOQ CĐT; không có ở bản _CDT, goi
        sổ ẩn QUAN_SAT, BOQ_MAP, DAY_TUONG, SPEC, PAGES; sheet hiện ISSUES, RFI_LOG, ASK_USER (không có dữ liệu thì không có sheet)
        DASHBOARD: mục ĐỌC SÂU; CẦN XEM thêm Tin cậy THẤP, Hỏi user
```

Độ tin cậy (tool, lấy mức thấp nhất):

```
 THAP  chưa có quan sát HÌNH/TEXT · trường giá KHÔNG THẤY · trường giá có dòng GEMINI MÂU THUẪN · thiếu Bản vẽ đối chứng
       còn tờ chưa đọc · Gemini lệch ở trường giá · số Claude không có trong số Gemini
 TB    MÂU THUẪN khác · Gemini chưa chạy, lỗi, lệch trường phụ · KL BOQ ≠ BV · còn khu chưa có KL theo BV
 CAO   còn lại
```

`check`: chưa có `_docsau.json` thì chỉ nhắc XEM, không LỖI. Có rồi: `--den docsau` đòi mọi cụm có QUAN_SAT và `Kiểm chéo Gemini`; `--den tincay` đòi `Độ tin cậy`. Luôn LỖI: nguồn HÌNH/TEXT không có tờ, phép cộng `a+b=c` sai, RFI không có trong RFI_LOG. Khóa `BOQ_MAP` mới không có trong nháp tool là LỖI.

Không chuyển, vì 10-tender đã có phần tương ứng: `boq-tach` (`boq`), `kiemke --ds` và `_bo_qua.txt` (`init --bo-thu-muc`), `_timkiem.*` (`ungvien` riêng), sheet `PROJECT`, `NOTES` (`README.md` của `10_TENDER`), `FILES` (`FILE_INDEX`), `BOQ_CDT` (`SCOPE`). `PAGES` chỉ giữ trang đã mở, không lập cho mọi trang như 01-scan, vì hồ sơ tender lớn; khung tên để `neo` kiểm lấy từ cache `trang`.
