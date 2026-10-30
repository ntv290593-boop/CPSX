# Flowchart skill 01-scan

> Cập nhật lần cuối: 2026-10-30 · Theo code: `SKILL/01-scan/` (SKILL.md, `buoc/*.md`, `tool/pd.py`)
> Trạng thái: pd.py N3.1 + N3.2 + N3.3 đã làm (chỉ XEM); N5 (đo thật) chưa làm.
> Quy tắc cập nhật: xem [README.md](README.md). Lịch sử thay đổi: [CHANGELOG.md](CHANGELOG.md).

Sơ đồ gồm 8 phần: luồng tổng, 5 bước, xếp hạng trang ứng viên, gộp dữ liệu / kiểm tra / build.

# 0. LUỒNG TỔNG (phiên điều phối)

```
 User: "scan <thư mục>"
        │
        ▼
┌─────────────────────────────────────────────────────────────────────┐
│ §0 CHUẨN BỊ (điều phối, không hỏi nếu đã rõ)                        │
│  GOC = thư mục dự án trong ACTIVE.md                                │
│  DOT = thư mục user nêu trong GOC\00_input\                         │
│        (không nêu → thư mục chưa có PROJECT_DATA_<DOT>.xlsx;        │
│         nhiều ứng viên → hỏi 1 câu)                                 │
│  DỪNG nếu: tồn tại ~$PROJECT_DATA_<DOT>.xlsx  hoặc  ổ D < 2 GB      │
└─────────────────────────────────────────────────────────────────────┘
        │
        ▼
  ┌──────────────┐  Mỗi bước = 1 SUBAGENT MỚI (ngữ cảnh sạch, chỉ đọc
  │ BƯỚC 1       │  file hướng dẫn của bước) → chống ảo giác.
  │ kiemke       │  Điều phối KHÔNG tự đọc BoQ/bản vẽ/spec.
  └──────┬───────┘
         ▼
   check --den kiemke ──LỖI?──► SendMessage cho subagent sửa (tối đa 2 vòng)
         │ (hết lỗi / hết 2 vòng → ghi báo cáo, đi tiếp)
         ├───────────────────────────────┐
         ▼                               ▼
  ┌──────────────┐               ┌──────────────┐   ← chạy SONG SONG
  │ BƯỚC 2 boq   │               │ BƯỚC 3 spec  │     (cả hai chỉ cần bước 1)
  └──────┬───────┘               └──────┬───────┘
         ▼                               │
   check --den boq                       │
         ▼                               │
  ┌───────────────────────────┐          │
  │ ĐỌC KẾT QUẢ PHẠM VI       │          │
  │ đếm: TRONG / KHÔNG THUỘC  │          │
  │ / CHƯA RÕ                 │          │
  └──────┬────────────────────┘          │
         │ có CHƯA RÕ?                   │
         ├─ CÓ ─► HỎI USER 1 câu gộp ─► ghi QUAN_SAT.user.json
         │        (cũng báo luôn: loại căn không có PDF)
         │ KHÔNG ─┐                      │
         ▼◄───────┘                      │
   Mã KHÔNG THUỘC SCOPE → loại khỏi bước 3,4 (PROJECT_DATA vẫn giữ dòng)
         │                               │
         ▼                               │
   py pd.py ungvien  → _ungvien.txt      │   (xem phần 4: thuật toán xếp hạng)
         │                               │
         ▼                               │
   CHIA NHÓM ≤ 8 mã cùng loại/cùng bộ tờ │
         ▼                               │
  ┌────────────────────────────────┐     │
  │ BƯỚC 4 banve × N nhóm (SONG    │     │
  │ SONG, mỗi nhóm 1 subagent,     │     │
  │ file *.banve-<n>.json)         │     │
  └──────┬─────────────────────────┘     │
         ▼                               ▼
   check --den banve   (và --den spec nếu bước 3 xong)
         ▼
   py pd.py neo  → _neo.json   (xem phần 7)
     ├─ có "NEO LỆCH" → gửi lại đúng nhóm subagent sửa (≤ 2 vòng, chạy lại neo)
     └─ còn LỆCH → để vào CẦN XEM, KHÔNG chặn build
         ▼
  ┌──────────────┐   tool chạy trước: doichieu → nháp độ tin cậy
  │ BƯỚC 5       │   LLM chỉ ghi ô mơ hồ
  │ doichieu     │
  └──────┬───────┘
         ▼
   py pd.py build → GOC\01_scan\PROJECT_DATA_<DOT>.xlsx (bản cũ → _backup)
         ▼
   BÁO USER (ngắn): đường dẫn · số hạng mục · LỖI còn · CẦN XEM · ASK_USER
```

# 1. BƯỚC 1: kiemke (kiểm kê hồ sơ)

```
 A. PHẠM VI (chưa quét gì)
    pd.py kiemke --ds  ─► liệt kê gọn theo thư mục (tên, đuôi, MB, số trang)
    pd.py mail <.msg>  ─► đọc mail
    pd.py xem <BoQ>    ─► chỉ sheet tổng, vài dòng đầu
          │
          ▼  rút ra: công ty tính GÌ / KHÔNG tính gì / hạn / phương án
 B. MODEL TỰ LƯỢC BỎ (suy luận, không luật cứng)
    Với mỗi thư mục/nhóm file, trả lời 4 câu:
      ① phục vụ hạng mục trong phạm vi không?   ② có PDF thay thế không?
      ③ bản cũ/trùng không?                      ④ chỉ để tham khảo không?
          │
          ├─► ĐỌC       → ghi 1 dòng NOTES
          ├─► DỰ PHÒNG  → NOTES + _timkiem.kiemke.json {"du_phong":["F004",…]}
          └─► BỎ        → GOC\01_scan\_bo_qua.txt (1 regex/dòng, # = lý do)
    Mặc định: PDF > CAD · ảnh hiện trạng bỏ · bộ BV loại căn không có trong BoQ bỏ
              bảng vật liệu / chi tiết chung / spec luôn ĐỌC · không chắc → hỏi user
    Bẫy Blanca: bộ FIT-OUT, PDF concept, file trùng, loại căn chỉ có CAD
          │
          ▼
 C. QUÉT  pd.py kiemke (script)
    ┌──────────────────────────────────────────────────────────────┐
    │ duyệt os.walk(00_input\<DOT>)                                │
    │   bỏ: ~$*, đuôi SKIP_EXT, tên SKIP_NAME, thư mục SKIP_DIR,   │
    │       regex trong _bo_qua.txt                          (L0)  │
    │ dup_report: nhóm trùng hash / cùng tên khác ngày-REV   (L1)  │
    │ đánh ID F001… cho từng file; Loại = loai_file(đường dẫn)     │
    │ với mỗi PDF, mỗi trang:                                      │
    │   text = page.get_text()                                     │
    │   Lớp text = CÓ  nếu >200 ký tự                              │
    │              ÍT  nếu 1..200                                  │
    │              KHÔNG nếu 0   (ảnh scan hoặc chữ vẽ bằng nét)   │
    │   title_guess → Số BV, Tên BV (máy đoán từ khung tên)        │
    │   lưu _text\F###.json (cache cho tim/ungvien/neo)            │
    │ build_timkiem:                                               │
    │   trang_trung: md5(text bỏ ngày dd/mm/yy + bỏ khoảng trắng)  │
    │     trùng ≥2 file → giữ bản ở file ÍT TRANG HƠN (hòa: file   │
    │     sau), các bản còn lại vào dup                            │
    │   file_tong: ≥85% trang có chữ là trang trùng file khác      │
    │   loai_to(Số BV + Tên BV) → chi_tiet / mat_dung / mat_bang / │
    │     thong_ke / danh_muc / legend / mep / hoan_thien / khac   │
    │   → ghi _timkiem.json                                        │
    └──────────────────────────────────────────────────────────────┘
          ▼
    Model rà cột `Loại` (máy đoán) → sửa ô sai; ghi PROJECT.kiemke.json
    (Tên DA, CĐT, TV thiết kế, Gói thầu, Đợt, Người YC, Hạn, YC báo giá,
     Phạm vi tính giá; không thấy → "chưa có", KHÔNG suy)
    Trang Lớp text=KHÔNG: để trống Số BV (bước 4 đọc khi cần, "làm lười")
          ▼
    Ghi NOTES (file hỏng, thiếu bộ BV, >200MB…) → check --den kiemke
```

# 2. BƯỚC 2: boq (đọc BoQ của CĐT)

```
 ┌───────────────────────── TOOL LÀM TRƯỚC ─────────────────────────┐
 │ pd.py boq-tach                                                   │
 │   mỗi sheet: sheet_header() tìm dòng tiêu đề (≤30 dòng đầu)      │
 │     ├ KHÔNG THẤY tiêu đề / sheet tổng hợp → in BỎ QUA → LLM tự đọc│
 │     └ thấy → chép MÁY từng ô: Sheet, Dòng, Bill, Item, Mô tả,    │
 │             ĐV, KL, Mô tả kỹ thuật, Link                         │
 │   dòng nhóm có KL nhưng thiếu STT/mã/mô tả/KT → tự bỏ + in      │
 │   KL là chữ (N/A, Rate only) → ISSUES.boq-tool BT01…            │
 │   → BOQ_CDT.boq-tool.json · ISSUES.boq-tool.json · _boq_tach.json│
 └──────────────────────────────────────────────────────────────────┘
          ▼
 LLM ghi BOQ_CDT.boq.json CHỈ gồm Sheet+Dòng + cột tool không làm được:
     Khu (thống nhất tên cho mọi sheet), Mã (BoQ không mã → đặt V01, V02…;
     CÙNG mô tả ở các khu = CÙNG mã). Cột nào LLM ghi thì thắng nháp tool.
          ▼
 BOQ_TONGHOP (mỗi Mã 1 dòng): Hạng mục, ĐV, KT BoQ (R×S×C chỉ khi mô tả rõ),
     Nhóm, Vị trí, Mô tả, Dòng BoQ gốc = link tới MỌI dòng của mã
 BOQ_MAP (mỗi cặp Mã×Khu 1 dòng)
          ▼
 ┌───────────────── PHÂN LOẠI SCOPE: pd.py scope ───────────────────┐
 │ Với mỗi Mã chưa có dòng "Phạm vi nhà máy" của LLM:               │
 │  S = fold(Hạng mục + Mô tả) ; T = fold(Mô tả kỹ thuật)           │
 │                                                                  │
 │  1. S khớp ^(bản lề|tay nắm|tay gạt|…|khóa) và không là khóa     │
 │     điện tử            → TRONG SCOPE – cửa (phụ kiện cửa)        │
 │  2. có "picomat"       → TRONG SCOPE – mộc (picomat CNC)         │
 │  3. "hộp màn" + T có ván ép/MDF → TRONG SCOPE – mộc              │
 │  4. len chân / chỉ gỗ: T có MDF|ván|laminate → TRONG (mộc)       │
 │                        T có nhựa → NGOÀI – Len/chỉ nhựa          │
 │                        còn lại  → None (LLM quyết)               │
 │  5. gạch (trừ nẹp) / sàn + gạch → NGOÀI – Gạch                   │
 │  6. DUYỆT SC_OUT (22 nhóm: khóa điện tử, sơn nước, thạch cao,    │
 │     đá, sàn, nẹp KL, lam, cửa thép/nhôm, compact, lan can, đèn,  │
 │     điện, thiết bị, nệm, chăn ga, rèm, thảm, decor, signage):    │
 │        khớp regex nhóm?                                          │
 │          ├ nhóm "kiểm vật tư" = True?                            │
 │          │    S+T có vật tư xưởng (VT_XUONG: ván ép, MDF, veneer,│
 │          │    melamine, laminate, picomat, nệm, gỗ…)?            │
 │          │      CÓ  → "CHƯA RÕ – <nhóm> nhưng có <vật tư>"       │
 │          │      KHÔNG → "KHÔNG THUỘC SCOPE NHÀ MÁY – <nhóm>"     │
 │          └ nhóm không kiểm → "KHÔNG THUỘC SCOPE NHÀ MÁY – <nhóm>"│
 │  7. DUYỆT SC_IN (mộc/cửa/cabinet/loose) → "TRONG SCOPE – <xưởng>"│
 │  8. không khớp gì → None  → in "KHÔNG CHẮC" cho LLM              │
 └──────────────────────────────────────────────────────────────────┘
   LLM chỉ làm: mã KHÔNG CHẮC + dòng CHƯA RÕ + nhãn xưởng (đoán theo tên)
   → QUAN_SAT.boq.json (LLM thắng nháp tool cùng mã)
          ▼
 _timkiem.boq.json (chỉ khi có ích, phục vụ `ungvien`):
     khu_file {Khu→[F005,F006]} · ma_cha {PK03→[D-01]} ·
     tu_khoa {FR15→"gối nằm"} · ma_bv {V01→["VC1"], V03→["GUONG","MIRROR"]}
          ▼
 Kiểm số: tổng KL từng sheet vs dòng subtotal → lệch = issue B
 ISSUES B…: cùng mã khác mô tả/ĐV giữa các khu · KL chữ · dòng trùng
          ▼
 XONG KHI: đếm dòng có KL mỗi sheet khớp BOQ_CDT · mọi Mã có Phạm vi
           nhà máy · check --den boq không LỖI
```

# 3. BƯỚC 3: spec (chạy song song với bước 2)

```
 Danh sách mã = BOQ_TONGHOP − mã "KHÔNG THUỘC SCOPE…"
          │
   ┌──────┴────────────────────────────┐
   ▼                                   ▼
 [1] SPEC (MasterFormat PDF)          [2] MATERIAL LIST / bảng hoàn thiện
  tim <từ khóa> --file <spec>          mỗi mã: tìm mã VL (WD- LAM- PT- SS- MT-)
  pdftotext -layout → spec.txt         + dòng ML; phụ kiện: Ref ML, hãng, model
  scan_spec.py → mục (ref,title,       ├ có → BOQ_TONGHOP.`Spec · ML` = link
    text, page)                        │       mục spec + dòng ML
  Giữ mục ẢNH HƯỞNG GIÁ/CẤU TẠO:       └ không → "KHÔNG CÓ – đã tìm trong <file/mục>"
   vật liệu, dày, hoàn thiện, phụ      QUAN_SAT (Nguồn=SPEC): Vật liệu gỗ,
   kiện, thử nghiệm, bảo hành, mock-up  Mặt đá, Gương, Kính, Kim loại,
  → SPEC (1 mục/dòng, nguyên văn)       Giấy/vải, Ngoài phạm vi + link trang
          │                                   │
          └────────────┬──────────────────────┘
                       ▼
        Trang là ẢNH (Lớp text KHÔNG/ÍT)?
          CÓ → Claude đọc ảnh TRƯỚC
               → Gemini đọc ĐỘC LẬP cùng ảnh (transcribe JSON)
               → so từng mã:  khớp → ghi
                              lệch → xem lại chỗ đó
                              vẫn lệch → ghi cả 2 (SPEC + GEMINI), MÂU THUẪN
                       ▼
        ISSUES S…: spec↔BoQ, spec↔ML, ML↔BoQ khác (dày, VL, hãng, cấp cháy)
        KHÔNG chọn bên đúng (bước 5 làm) – chỉ ghi nhận
                       ▼
        pd.py neo (xem phần 7): mã/model trong link có trên trang/dòng được link?
                       ▼
        XONG KHI: mỗi mã có Spec·ML (link hoặc "KHÔNG CÓ – …"); check --den spec
```

# 4. THUẬT TOÁN TÌM TRANG ỨNG VIÊN (`pd.py ungvien`)

Cầu nối giữa bước 2 và bước 4. Thay cho việc cho subagent dò cả bộ bản vẽ.

```
 ĐẦU VÀO: _text\F*.json (cache chữ) · PAGES (Số BV/Tên BV) · BOQ_TONGHOP
          · BOQ_MAP (Khu) · _timkiem.kiemke.json · _timkiem.boq.json
                         │
        ┌────────────────▼─────────────────┐
        │ build_index → chi_muc()          │
        │  body[p]  = Counter(từ trong trang)                          │
        │  title[p] = Counter(từ trong Số BV + Tên BV)                 │
        │  codes[p] = Counter(mã chuẩn hóa trong trang)                │
        │  tit_codes[p] = mã trong tên tờ                              │
        │  df[w] = số trang chứa từ w   N = tổng trang                 │
        │  cdf[c] = số trang chứa mã c  avgdl = độ dài TB              │
        │  loai[p] = chi_tiet|mat_dung|mat_bang|thong_ke|danh_muc|     │
        │            legend|mep|hoan_thien|khac                        │
        └────────────────┬─────────────────┘
                         ▼
   Với MỖI Mã trong BOQ_TONGHOP (bỏ mã ngoài scope):
   ┌───────────────────────────────────────────────────────────────────┐
   │ q_ten  = toks(tu_khoa[mã]  hoặc  Hạng mục cắt trước dấu "@")      │
   │ q_mota = toks(Mô tả BoQ[:300])          qs = q_ten ∪ q_mota       │
   │ bc = base_code(mã): bỏ _ST/_1BR/_2BR, -DT/-HM/-TL, chữ a-d cuối   │
   │      ncode: LF-04 = LF04 = lf.4 → "lf4"                           │
   │ bcs = {bc} ∪ ncode(ma_bv[mã])   (chữ không phải mã → thêm vào qs) │
   │ bcs bỏ mã có cdf > max(8, 10%·N)   ← mã dính quá nhiều trang      │
   │                                      (ô lưới F1, cao độ) vô nghĩa │
   │ bm_on = (số từ của qs có trong df) / |qs| ≥ 50%                   │
   │         ← cùng ngôn ngữ BoQ & BV mới dùng BM25                    │
   │ want = khu_key(Khu của mã)                                        │
   └───────────────────────────────────────────────────────────────────┘
        │
        ▼   duyệt TỪNG TRANG p = (fid, trang):
   ┌───────────────────────────────────────────────────────────────────┐
   │ LOẠI NGAY nếu: p ∈ trang trùng │ fid ∈ file tổng │ fid ∈ dự phòng │
   │                │ loai[p] ∈ {mep, hoan_thien}                      │
   │                                                                   │
   │ lech = file khác khu?  (khu_file do LLM ghi ưu tiên; không có     │
   │        thì đoán từ khu theo đường dẫn, sâu nhất trước)            │
   │                                                                   │
   │ if bm_on:                                                         │
   │    sc = BM25(qs, body[p])  +  3.0 × (số từ qs có trong title[p])  │
   │         BM25 = Σ idf·tf·(k1+1)/(tf + k1·(1−b + b·dl/avgdl))       │
   │                idf=ln(1+(N−df+.5)/(df+.5)), k1=1.2, b=0.75        │
   │ else:                                                             │
   │    kh = q_ten ∩ body[p];  sc = |kh| nếu |kh| ≥ min(2,|q_ten|)     │
   │                                     else 0                        │
   │ mh = Σ codes[p][c] với c ∈ bcs           (số lần mã xuất hiện)    │
   │ sc += 8·min(mh, 3)  +  10 nếu bcs ∩ tit_codes[p] ≠ ∅             │
   │ if lech: sc ×= 0.4          ← hạ điểm, KHÔNG loại (file GENERAL   │
   │                                chung vẫn dùng được)               │
   │ giữ nếu sc > 0                                                    │
   └───────────────────────────────────────────────────────────────────┘
        ▼ sắp giảm dần theo điểm
   CHIA NHÓM RA (_ungvien.txt):
     CHI TIẾT  : top 8, loại ∈ {chi_tiet, mat_dung, khac, legend}
     THỐNG KÊ  : top 3, loại thong_ke  VÀ có mã (mh>0)
     MẶT BẰNG  : top 4, loại mat_bang  VÀ có mã   → chỉ để đếm KL theo khu
     DANH MỤC  : top 2, loại danh_muc  VÀ có mã
     CỦA MÃ CHA: 6 trang đầu của mã cha (ma_cha) → CHỈ ĐỂ TÌM, vẫn đọc riêng
     Không trang nào → "KHÔNG CÓ ỨNG VIÊN"
```

# 5. BƯỚC 4: banve (bước dễ sai nhất; mỗi nhóm ≤ 8 mã, chạy song song)

```
 subagent nhận: danh sách mã + khối "Ứng viên" của nhóm từ _ungvien.txt
 NGUYÊN TẮC: ① hình trước, mô tả BoQ sau  ② mỗi giá trị có tờ + số chi tiết
   ③ KHÔNG CÓ (nhìn rõ là không có) ≠ KHÔNG THẤY (chưa đọc được); không bỏ trống
   ④ lệch → giữ CẢ HAI, MÂU THUẪN, không tự chọn
   ⑤ KL đếm trên BV, không lấy KL BoQ chia
  ┌──────────────────────── với TỪNG MÃ ───────────────────────────┐
  │ (1) TÌM BỘ TỜ                                                   │
  │     đọc Ứng viên theo thứ tự rẻ→đắt:                           │
  │       CHI TIẾT → THỐNG KÊ (cửa) → MẶT BẰNG (đếm KL) → DANH MỤC  │
  │     thiếu → tim <mã/tên> → callout (vd 18/ID-PS-TY-712) →       │
  │              Drawing List. PDF trước, .dwg chỉ khi KHÔNG có tờ  │
  │     Trang Lớp text=KHÔNG: crop khung tên 100dpi xác nhận số tờ  │
  │                                                                 │
  │ (2) ĐÚNG TỜ?  số tờ khung tên = tờ được dẫn tới?                │
  │               type/khu khớp dòng BoQ? (type 2 ≠ type 3)         │
  │                  ├ sai → không dùng lẫn, tìm tờ khác            │
  │                  ▼                                              │
  │ (3) ĐỌC (rẻ → đắt)                                              │
  │       Lớp text = CÓ  → chu / so  (chữ, số kích thước + vị trí)  │
  │       Lớp text = KHÔNG/nét → anh --dpi 100 → anh --clip         │
  │                              --dpi 150–220 phóng từng vùng      │
  │     Ghi QUAN_SAT 1 dòng/trường (Nguồn HÌNH|TEXT):               │
  │       Lắp đặt, Kích thước BV, Có cánh, Ngăn kéo, Thanh treo,    │
  │       Kệ/đợt, Hốc TV, Mặt đá, Gương, Kính, Kim loại, Giấy/vải,  │
  │       Mã hoàn thiện trên BV, Ván nền ẩn, Khung xương, Cong/bo,  │
  │       Ngoài phạm vi, Khu, Bộ tờ đã đọc, Tờ chưa đọc, Rev tờ     │
  │                                                                 │
  │ (4) KIỂM CHUỖI KÍCH THƯỚC  a+b+c = tổng ?                       │
  │       khớp → "450+450=900 ✓"      lệch → MÂU THUẪN + issue      │
  │       (check/doichieu có regex CHUOI tự bắt chuỗi cộng sai)     │
  │                                                                 │
  │ (5) GEMINI = NGƯỜI ĐỌC THỨ HAI (bắt buộc cho mọi crop làm bằng chứng)
  │       a. Claude ghi giá trị CỦA MÌNH TRƯỚC                      │
  │       b. gemini --quansat "<tên món>" trên crop mặt đứng/chi tiết│
  │          gemini --claude "900,450,450" trên crop có dãy số      │
  │          (KHÔNG đưa số/nhận xét của Claude vào câu hỏi)         │
  │       c. so từng trường:                                        │
  │            khớp → giữ dòng Claude                               │
  │            lệch → Claude phóng đúng chỗ đọc lại                 │
  │                   vẫn lệch → 2 dòng (HÌNH + GEMINI), MÂU THUẪN  │
  │            Gemini thấy món Claude ghi KHÔNG THẤY → đọc lại theo │
  │                   `where`                                       │
  │       d. KHÔNG nhận số Gemini tự tính (không in trên hình);     │
  │          không dùng Gemini làm nguồn duy nhất                   │
  │       e. ghi QUAN_SAT `Kiểm chéo Gemini`:                       │
  │          "khớp: …; lệch: …; crop: …"   (bước 5 dùng chấm tin cậy)│
  │       f. agy lỗi → "KHÔNG THẤY – <lỗi>" ⇒ tin cậy không được CAO│
  │                                                                 │
  │ (6) Nhóm = Cửa → DAY_TUONG: dày tường tại TỪNG vị trí thật      │
  │ (7) BOQ_MAP: KL theo BV từng khu + Tờ đếm KL                    │
  │       có bảng thống kê → lấy; không → đếm mặt bằng + nhờ Gemini │
  │       đếm cùng tờ; 2 số khác → đếm lại theo danh sách phòng;    │
  │       vẫn khác → issue                                          │
  │ (8) BOQ_TONGHOP: Bản vẽ đối chứng (tờ rõ nhất đứng đầu),       │
  │       Kích thước BV (chiều chưa rõ → "?")                       │
  │ (9) PAGES: mọi trang đã mở (Loại tờ, Mã, Đã đọc bằng, Rev)      │
  │ (10) ISSUES V<nhóm>-nn: thiếu tờ, BV↔BoQ lệch, mã HT BV≠ML, rev │
  └─────────────────────────────────────────────────────────────────┘
  TRƯỚC KHI TRẢ: mỗi mã ≥1 dòng QUAN_SAT nguồn HÌNH · có Bản vẽ đối chứng
     · có Tờ chưa đọc ("không còn" nếu đủ) · có Kiểm chéo Gemini
     → check --den banve → pd.py neo → sửa NEO LỆCH (xem phần 7)
```

# 6. BƯỚC 5: doichieu (thuật toán chấm độ tin cậy)

```
 pd.py doichieu (tool làm trước) – với MỖI Mã trong BOQ_TONGHOP:

   Mã ngoài scope ──► muc = CAO,  Giả định = "Không thuộc scope nhà máy (<lý do>)"
   │                  (không so, không issue, không RFI)
   ▼ còn lại: lv = CAO ; why = []   (worse(a,b) = hạ xuống mức thấp hơn)

   ┌─ điều kiện hạ mức (áp lần lượt, lấy mức THẤP NHẤT) ────────────────────┐
   │ THAP  không có dòng QUAN_SAT nguồn HÌNH/TEXT                            │
   │ THAP  trường quyết định giá chỉ có KHÔNG THẤY, không có CÓ/KHÔNG CÓ     │
   │         (PRICE_FIELDS: Kích thước BV, Có cánh, Ngăn kéo, Kệ/đợt,       │
   │          Mặt đá, Gương, Kính, Kim loại, Giấy/vải)                       │
   │ THAP  trường giá có dòng GEMINI = MÂU THUẪN (Claude↔Gemini lệch)        │
   │ TB    có MÂU THUẪN ở trường khác                                        │
   │ THAP  thiếu Bản vẽ đối chứng (rỗng / bắt đầu "KHÔNG","CHƯA")            │
   │ THAP  còn "Tờ chưa đọc" (khác không|hết|none|-)                         │
   │ TB    Gemini chưa chạy / tất cả KHÔNG THẤY                              │
   │ Gemini lệch: nếu đoạn sau "lệch:" khớp GIA_RX (đợt|cánh|ngăn|kích      │
   │        thước|sâu|rộng|cao|dày|mặt đá|kính|gương|kim loại|vật liệu|     │
   │        giấy|vải) → THAP, ngược lại → TB                                 │
   │ KL từng khu:  KL BoQ (Σ BOQ_CDT) vs KL theo BV (BOQ_MAP)                │
   │      bv rỗng → CHƯA ĐỌC → hạ TB      |Δ|<1e-6 → KHỚP                   │
   │      khác    → LỆCH     → hạ TB                                         │
   └─────────────────────────────────────────────────────────────────────────┘
   BOQ_MAP.Độ tin cậy = worse(muc[mã], {KHỚP:CAO, LỆCH:TB, CHƯA ĐỌC:THAP})
   Issue nháp DT01… chỉ cho mã CHƯA có issue nào:
        lệch KL → "BoQ↔BV" · KHÔNG THẤY → "Thiếu tờ" (CAO nếu thiếu KT)
        Gemini lệch → "Claude↔Gemini"
   In "CHUỖI SAI" nếu a+b=c tổng sai (sai số > 0.51)
   → BOQ_TONGHOP.doichieu-tool · BOQ_MAP.doichieu-tool · ISSUES.doichieu-tool
     · _doichieu_tool.json (lý do từng mã)

 LLM làm tiếp (ô nào LLM ghi thì thắng nháp tool):
   Mã hoàn thiện (BV ↔ ML; mã ghi trên BV tại chỗ ưu tiên hơn block cùng loại)
   Giả định tính giá (đang tạm tính theo gì; KHÔNG tự quyết bên đúng)
   Hành động / RFI / Trạng thái=MỞ cho MỌI issue (B,S,V,D)
   RFI_LOG: gom issue cùng chủ đề → 1 câu hỏi CĐT (Nháp)
   ASK_USER: quyết định nội bộ + "Vì sao cần" (không hỏi lại điều đã có trong .user.json)
   Nghi bước 4 đọc sai → ghi issue D… "Cần đọc lại", KHÔNG sửa số
```

# 7. NEO / CHECK / BUILD (lớp kiểm máy và gộp dữ liệu)

```
 ┌───────────── pd.py neo  (chỉ XEM, không bao giờ là LỖI) ────────────┐
 │ Duyệt mọi link (nhãn, file, trang, ô) trong JSON:                   │
 │  N3.1 SỐ TỜ trong nhãn link ("F005 p43 604.1", ID-CO-TY-703)        │
 │       có nằm trong text trang được link không?                      │
 │       (nhãn chỉ "p30" → KHÔNG KIỂM ĐƯỢC)                            │
 │  N3.2 mã/model ở link SPEC/ML (WD-01, K6121AN, FF&E 51-400) có trên │
 │       trang/dòng Excel được link không? (nhiều link: mã phải có ở   │
 │       ≥1 link). Tên file & kích thước KHÔNG tính là mã.             │
 │  N3.3 Mã VL (WD-, GL1…) và kích thước R×S×C trong QUAN_SAT/BOQ_*    │
 │       có trong text các trang link không?                           │
 │       – chỉ kiểm khi trang có loại đó ở dạng text:                  │
 │         + mã VL: ≥2 mã VL khác nhau dạng text mới kiểm mọi họ;      │
 │           ít hơn chỉ kiểm họ có trên trang                          │
 │         + kích thước: trang cần ≥10 số ≥100 (không tính cao độ,     │
 │           dòng lặp ≥30% trang = khung tên); chỉ lấy chuỗi R×S×C     │
 │           ĐẦU TIÊN của giá trị                                      │
 │       – BV vẽ số bằng nét (GA Design, SOFITEL) → KHÔNG KIỂM ĐƯỢC,   │
 │         không báo LỆCH oan                                          │
 │  Kết quả: KHỚP | NEO LỆCH (+ gợi ý) | KHÔNG KIỂM ĐƯỢC → _neo.json   │
 │  _neo.json gắn hash link: link đổi sau lần chạy → check nhắc chạy lại│
 └─────────────────────────────────────────────────────────────────────┘

 ┌──────────────── pd.py check [--den <bước>] ────────────────────────┐
 │ upto = BUOC[: index(den)+1]   BUOC = kiemke|boq|spec|banve|        │
 │                                      doichieu|user                 │
 │ LỖI (chặn):                                                        │
 │   cột bắt buộc trống (chỉ xét cột thuộc các bước ≤ den)            │
 │   cột SỐ không phải số · cột LIST ngoài danh sách                  │
 │   link: thiếu file / file không tồn tại / trang ngoài file         │
 │   Mã ở BOQ_CDT mà thiếu BOQ_TONGHOP (và ngược lại); cặp Mã/Khu    │
 │   lệch giữa BOQ_CDT và BOQ_MAP                                     │
 │   _timkiem.*.json trỏ mã/file/khu không tồn tại                    │
 │   KL CĐT không phải số                                             │
 │   QUAN_SAT nguồn HÌNH/TEXT mà không có tờ/chi tiết                 │
 │   chuỗi cộng a+b=c sai                                             │
 │   mã chưa có QUAN_SAT / thiếu dòng Kiểm chéo Gemini (khi ≥ banve)  │
 │   ISSUES.RFI không có trong RFI_LOG                                │
 │ XEM (ngoại lệ cho user): KL BoQ≠BV · tin cậy THẤP · BV đối chứng /  │
 │   Spec·ML = KHÔNG · KHÔNG THẤY / MÂU THUẪN · issue CAO đang MỞ ·   │
 │   ASK_USER chưa trả lời · trang Lớp text KHÔNG chưa xem · NEO LỆCH │
 │ exit code 1 nếu có LỖI                                             │
 └─────────────────────────────────────────────────────────────────────┘

 ┌──────────────── pd.py build  (merge → xlsx) ───────────────────────┐
 │ khoá file ~$… tồn tại → DỪNG                                       │
 │ load_form (FORM.xlsx) → pull_user (đọc sửa tay của user từ xlsx    │
 │   hiện có + _snapshot.json) → load_json (sắp theo BUOC; "-tool"    │
 │   trước LLM)                                                       │
 │ MERGE – ai thắng ô nào:                                            │
 │   kiểu KHOA: ghép theo khoá, mỗi ô thuộc 1 bước                    │
 │     • nháp -tool KHÔNG đè ô của LLM/user                           │
 │     • LLM sửa nháp -tool → LLM thắng                               │
 │     • user ghi đè mọi bước (ghi vào "Ghi đè:" trên DASHBOARD)      │
 │     • 2 bước (không phải tool/kiemke) ghi KHÁC nhau cùng ô → LỖI   │
 │     • bước ghi cột không thuộc quyền (meta.buoc) → LỖI             │
 │   kiểu NOI (QUAN_SAT, ISSUES…): nối dòng; riêng "Phạm vi nhà máy"  │
 │     nháp tool nhường dòng LLM/user cùng mã                         │
 │ ghi ô: CT = công thức {cột} → '=…'; LINK → hyperlink + _LINKS;     │
 │   STT = số dòng; FILES.Trạng thái scan tự tính ("đọc x/N trang;    │
 │   n link dẫn tới")                                                 │
 │ run_checks → DASHBOARD → ẩn _SHEETS/_COT/_HUONG_DAN                │
 │ sao bản cũ vào 01_scan\_backup → save → _snapshot.json             │
 │ KHÔNG BAO GIỜ xoá *.user.json, _snapshot.json                      │
 └─────────────────────────────────────────────────────────────────────┘
```

# 8. Ba điểm thiết kế đáng chú ý

1. **Tool trước, LLM sau.** boq-tach, scope, ungvien, doichieu và neo là phần chép, so và xếp hạng bằng máy. LLM chỉ ghi phần mơ hồ. Khi merge, ô LLM ghi thì thắng nháp tool, và ô user ghi thì thắng tất cả.
2. **Lọc theo thứ tự rẻ đến đắt.** Thứ tự: bỏ file → bỏ trang trùng → Drawing List → mã BoQ → BM25 → crop khung tên → crop chi tiết. Trước khi mở ảnh luôn dùng `chu` và `so` (lấy chữ, số từ text).
3. **Không tự chọn bên đúng.** Khi BoQ, bản vẽ, spec hoặc Gemini lệch nhau, hệ thống giữ cả hai giá trị, đặt MÂU THUẪN, tạo issue, và hạ độ tin cậy. Việc chọn thuộc về RFI hoặc user.
