# Bước 2 — boq: đọc BoQ của CĐT

**Ghi:** `BOQ_CDT`, `BOQ_TONGHOP` (cột của boq), `BOQ_MAP` (chỉ `Mã`, `Khu`), `QUAN_SAT` (chỉ trường `Phạm vi nhà máy`, mục 4b), `ISSUES` (ID `B01`…), `ASK_USER`, `NOTES` — file `*.boq.json`.
**Không làm:** mở bản vẽ, spec; không đoán kích thước hay vật liệu ngoài chữ BoQ ghi.

## Làm
0. **Chạy `py -I "<TOOL>" boq-tach <GOC> <DOT>` trước** → nháp `BOQ_CDT.boq-tool.json` (Sheet, Dòng, Bill, Item, Mô tả, ĐV, KL, Mô tả kỹ thuật, Link — chép máy, đúng từng ô) + `ISSUES.boq-tool.json` (`BT01`…: KL ghi chữ) + `_boq_tach.json` (mã BoQ `ma_boq` và R/S/C của từng dòng, để điền `Kích thước BoQ`). Đo trên Blanca: 393/393 dòng trùng bản chép tay. Tool in từng sheet: dòng tiêu đề nó nhận, số dòng; sheet báo `KHÔNG THẤY dòng tiêu đề` hoặc `BỎ QUA` (sheet tổng hợp) thì bạn tự đọc / chạy lại với `--sheet`. Dòng nhóm có KL nhưng không STT/mã/mô tả/kích thước (vd `Phụ kiện cửa` / bộ / 1) tool tự bỏ và in ra — xem có đúng không.
   Bạn **không chép lại** các cột tool đã ghi. Bạn ghi `BOQ_CDT.boq.json` chỉ với khóa `Sheet`, `Dòng` + phần tool không làm được: **`Khu`, `Mã`** (bắt buộc), sửa `Bill`/`Mô tả` nếu tool sai. LLM ghi cột nào thì thắng nháp tool ở cột đó. Dòng tool bỏ sót mà BoQ có KL → thêm đủ cột.
1. Tìm file BoQ trong `FILES.kiemke.json` (`Loại = BoQ`). Xem sheet: `py -I "<TOOL>" xem "<file>"`; xem vùng: `… xem "<file>" "<sheet>" A1:L120`.
2. **BOQ_CDT** — mỗi dòng hạng mục có khối lượng của gói đang báo giá thành 1 dòng:
   - `Sheet`, `Dòng` đúng vị trí trong file; `Link` = `["BoQ <sheet>!<dòng>", "F0xx", "", "'<sheet>'!A<dòng>"]`.
   - `Mô tả`, `Mô tả kỹ thuật`: chép nguyên văn, không dịch, không rút gọn.
   - `KL CĐT`: số. Ô ghi chữ (`N/A`, `Rate only`, `Incl.`) → bỏ cột KL, ghi issue B.
   - `Khu`: tên khu/khối/bill dùng thống nhất (vd `PODIUM BLK01`), cùng tên cho cùng một khu ở mọi sheet.
   - `Mã`: mã trong BoQ (mã cửa, mã hạng mục). BoQ không có mã → đặt mã ngắn theo nhóm (`V01`, `V02`…), **cùng mô tả ở các khu = cùng mã**; ghi cách đặt mã vào NOTES.
3. **BOQ_TONGHOP** — mỗi `Mã` 1 dòng: `Hạng mục` (tên ngắn tiếng Việt), `ĐV`, `Kích thước BoQ` (chỉ khi mô tả ghi rõ, dạng R×S×C mm), `Nhóm`, `Vị trí`, `Mô tả BoQ`, `Mô tả kỹ thuật BoQ` (nguyên văn), `Dòng BoQ gốc` = link tới **mọi** dòng BOQ_CDT của mã. Các cột KL là công thức, không ghi.
4. **BOQ_MAP** — mỗi cặp (Mã, Khu) có trong BOQ_CDT 1 dòng, chỉ ghi `Mã`, `Khu`.
4b. **Phân loại ban đầu theo scope nhà máy (BOHO) — để lọc bản vẽ.** Chạy `py -I "<TOOL>" scope <GOC> <DOT>` sau khi có `BOQ_TONGHOP.boq.json` → nháp `QUAN_SAT.boq-tool.json` (luật từ khóa trên Hạng mục + Mô tả BoQ, ghi `TRONG SCOPE – <xưởng>`, `KHÔNG THUỘC SCOPE NHÀ MÁY – <nhóm>`, hoặc `CHƯA RÕ – <nhóm> nhưng có <vật tư xưởng>`). Tool chỉ gán khi khớp chắc, in danh sách `KHÔNG CHẮC` — **bạn chỉ phân loại các mã đó + rà các dòng CHƯA RÕ và dòng nhãn xưởng** (nhãn xưởng mộc/cabinet/loose là đoán theo tên món). Muốn sửa mã nào thì ghi dòng `Phạm vi nhà máy` của mã đó vào `QUAN_SAT.boq.json`; dòng LLM thay hẳn dòng tool cùng mã. Luật đầy đủ nằm ở `SC_OUT`/`SC_IN`/`scope_phan_loai` trong `pd.py`. Đo trên Blanca (148 mã): 148/148 gán, 0 mã còn lại cho LLM; chưa có đáp án LLM để so, nên user duyệt theo ngoại lệ (xem DOCS plan mục 7).
   Mỗi `Mã` 1 dòng `QUAN_SAT.boq.json`: `Trường = Phạm vi nhà máy`, `Nguồn = BOQ`, `Trạng thái = CÓ`, chỉ dựa vào chữ trong `Mô tả` + `Mô tả kỹ thuật`.
   Scope nhà máy = 5 xưởng: **mộc · sơn (sơn hoàn thiện đồ gỗ) · cửa · cabinet · loose**.
   - `TRONG SCOPE – <xưởng>`: món do một trong 5 xưởng làm. User đã chốt (27/10):
     - tủ, bàn, kệ, khung gương gỗ, cửa gỗ **+ phụ kiện cửa** (bản lề, khóa, tay gạt, chặn cửa, đệm khí)
     - **vách kính** (kể cả khung inox — nhiều khi có ván nền gắn kính)
     - **tấm nhựa picomat** chạy CNC → xưởng mộc
     - **len chân MDF/ván** (kể cả phủ laminate)
     - **hộp màn có lót ván ép** (dù khung thạch cao)
     - đồ loose **kể cả gối trang trí / gối nằm** (chỉ vải + gòn)
     - ốp tường gỗ / vách ốp MDF veneer, laminate; panel tường bọc nệm vải/da; **giấy/vải dán tường**; lam gỗ/MDF ốp tường, trần; trần gỗ / trần ván veneer
     - bục gỗ / sàn nâng gỗ; cầu thang gỗ, tay vịn gỗ
     - cửa gỗ chống cháy; phụ kiện tủ (ray, bản lề, tay nắm, giá kéo); quầy lễ tân / quầy bar; tủ locker; **tủ minibar** (thường làm bằng gỗ)
     - gương thủy dán tường không khung; bàn ghế ngoài trời
   - `KHÔNG THUỘC SCOPE NHÀ MÁY – <nhóm>`: mọi thứ ngoài 5 xưởng. **Bắt buộc ghi nhóm lý do** sau dấu `–`, dùng đúng các nhóm: `Sơn nước` (kể cả sơn lót/bột trét sau lưng tủ) · `Thạch cao` (trần, mặt dựng, hộp đèn, nắp thăm, V/Z lưới, khoét lỗ) · `Sàn` (kể cả sàn gỗ công nghiệp, sàn gỗ tự nhiên) · `Gạch` · `Đá` (kể cả mặt đá tủ bếp/vanity tách dòng riêng) · `Nẹp kim loại` · `Len/chỉ nhựa` · `Lam nhựa/nhôm` (composite, nhôm vân gỗ) · `Cửa thép` (kể cả chống cháy) · `Cửa nhôm kính` (kể cả vách nhôm kính) · `Vách compact` · `Khóa điện tử` (khóa từ, khóa thẻ) · `Kim loại trang trí` (ốp inox/đồng, lam kim loại) · `Lan can` (kính, tay vịn kim loại) · `Đèn` (kể cả LED hắt trong tủ, đèn trang trí rời) · `Điện` (ổ cắm, công tắc, đi dây trong tủ) · `Thiết bị` (chậu, vòi, bếp, hút mùi, lò, TV, két sắt, thiết bị điện rời) · `Nệm` · `Chăn ga gối` · `Rèm` · `Thảm` · `Decor` (tranh, khung tranh, đồ trang trí) · `Signage`. Nhóm mới → thêm vào danh sách này sau khi user chốt. Vd `KHÔNG THUỘC SCOPE NHÀ MÁY – Thạch cao`.
   - **Lẫn vật tư xưởng làm → vẫn báo, không tự loại.** Mã thuộc nhóm KHÔNG ở trên nhưng mô tả có vật tư/công việc của 5 xưởng (gỗ, ván, MDF, veneer, melamine/laminate, picomat, nệm/vải bọc, sơn PU đồ gỗ…) → ghi `CHƯA RÕ – <nhóm> nhưng có <vật tư xưởng>` (vd `CHƯA RÕ – Cửa nhôm kính nhưng có ván nền veneer`; `CHƯA RÕ – Decor nhưng có khung tranh gỗ`) để hỏi user.
   - `CHƯA RÕ – <lý do>`: không chắc thuộc xưởng nào → 1 câu `ASK_USER` gộp mọi mã CHƯA RÕ. Không tự xếp.
   Mã `KHÔNG THUỘC SCOPE NHÀ MÁY` không giao bước spec/banve, không có issue, không vào RFI (tool tự bỏ khỏi `ungvien` và kiểm). PROJECT_DATA vẫn giữ dòng.
   Đây **không phải** phần "Không bao gồm" khi tính giá chi tiết (vải, kim loại, thiết bị trong một món trong scope) — phần đó do bước costing ghi, bước scan không làm.
4c. **`_timkiem.boq.json`** (giúp `ungvien` tìm đúng trang; không vào PROJECT_DATA; `check` kiểm mã/file có thật). Chỉ ghi khi có ích:
   ```json
   {"khu_file": {"<Khu trong BOQ_MAP>": ["F005", "F006"]},
    "ma_cha": {"PK03": ["D-01", "D-02"], "KT2-HM": ["KT2"]},
    "tu_khoa": {"FR15": "gối nằm"},
    "ma_bv": {"V01": ["VC1"], "V03": ["GUONG", "MIRROR"]}}
   ```
   - `khu_file`: Khu nào ứng với file BV nào (LLM nhìn tên file/thư mục; "Hành lang" ≠ "Area 50 - Guestroom Corridor"). Không ghi thì tool đoán theo từ đầu của Khu so với đường dẫn file.
   - `ma_cha`: phụ kiện → cửa, hạng mục con (`-HM`/`-DT`) → tủ. Trang của mã cha chỉ dùng để TÌM; bước 4 vẫn đọc riêng từng mã.
   - `ma_bv`: món mã BoQ tự đặt (V01, F001…) mà BV ghi bằng mã khác (VC1 trong legend) → ghi mã BV; chữ không phải mã (GUONG) thành từ khóa. **Bắt buộc khi Hạng mục tiếng Việt mà BV tiếng Anh** (đo NORTON vanity: không có thì 0/12 tờ phòng tắm, có thì 11/12).
   - `tu_khoa`: tên món bỏ phần vị trí ("Gối nằm @ phòng ngủ" → "gối nằm"; tool đã tự cắt phần sau `@`), thêm tên tiếng Anh nếu BV tiếng Anh.
5. Kiểm số: so tổng KL từng sheet với dòng tổng/subtotal của BoQ (nếu có). Lệch → issue B.
6. **ISSUES (B…)** chỉ cho chuyện nội bộ BoQ: cùng mã mà mô tả/ĐV khác giữa các khu; KL chữ; dòng trùng; tổng lệch subtotal. `Nội dung` nêu cả hai giá trị, `Nguồn` link tới các dòng.

## Xong khi
- Mọi dòng có KL của gói đều nằm trong BOQ_CDT (đếm lại số dòng theo từng sheet và ghi số đó vào kết quả trả về).
- Mọi `Mã` có dòng `Phạm vi nhà máy`; kết quả trả về ghi số mã TRONG SCOPE / KHÔNG THUỘC SCOPE / CHƯA RÕ.
- `py -I "<TOOL>" check <GOC> <DOT> --den boq` không còn LỖI.
