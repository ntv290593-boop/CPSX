# Bước 5 — doichieu: đối chiếu BoQ ↔ bản vẽ ↔ spec, chốt độ tin cậy

**Ghi:** `BOQ_TONGHOP` (`Mã hoàn thiện`, `Issue`, `Độ tin cậy`, `Giả định tính giá`), `BOQ_MAP` (`Issue`, `Trạng thái`, `Độ tin cậy`), `ISSUES` (cột `Hành động`, `RFI`, `Trạng thái` cho mọi issue; issue mới ID `D01`…), `RFI_LOG`, `ASK_USER` — file `*.doichieu.json`.
**Không làm:** đọc lại bản vẽ để đổi số của bước 4. Thấy nghi sai → ghi issue `D…` (`Loại = Cần đọc lại`), không sửa.

## Đầu vào
`py -I "<TOOL>" check <GOC> <DOT> --den banve` (danh sách XEM) + các file JSON của bước 2–4.

## Chạy tool trước (làm xong phần máy làm được)
`py -I "<TOOL>" doichieu <GOC> <DOT>` → nháp `BOQ_TONGHOP.doichieu-tool.json` (`Độ tin cậy`, `Issue`, `Giả định tính giá` cho mã ngoài scope), `BOQ_MAP.doichieu-tool.json` (`Trạng thái` KHỚP/LỆCH/CHƯA ĐỌC từ KL BoQ ↔ KL theo BV, `Độ tin cậy`), `ISSUES.doichieu-tool.json` (`DT01`…: chỉ cho mã chưa có issue nào) và `_doichieu_tool.json` (lý do từng mã). Đo trên Blanca: BOQ_MAP Trạng thái trùng 373/373; Độ tin cậy mã trùng ~91% (tool thận trọng: lệch nhau chủ yếu tool THẤP hơn / TB hơn). Tool chưa đọc được: mã hoàn thiện BV↔ML, issue ghi `Mã` kiểu "Nhiều mã" / "Tất cả", cân nhắc ảnh hưởng giá.
**Việc của bạn:** đọc `_doichieu_tool.json`, ghi `*.doichieu.json` chỉ những ô bạn đổi hoặc bổ sung: `Mã hoàn thiện`, `Giả định tính giá`, `Hành động`/`RFI`/`Trạng thái` cho mọi issue, `Issue` của mã mà issue ghi Mã mơ hồ, nâng/hạ `Độ tin cậy` khi biết thêm (kèm lý do trong `Giả định tính giá`), RFI_LOG, ASK_USER. Ô nào bạn ghi thì thắng nháp tool; ô bạn không ghi thì nháp tool vào PROJECT_DATA. Tool không tự quyết bên nào đúng. Dòng "CHUỖI SAI" tool in ra = chuỗi `a+b=c` ở QUAN_SAT tổng sai → ghi issue `Cần đọc lại`.

## Mã ngoài scope nhà máy
QUAN_SAT `Phạm vi nhà máy` = `KHÔNG THUỘC SCOPE NHÀ MÁY…` → chỉ ghi `Độ tin cậy = CAO`, `Giả định tính giá = Không thuộc scope nhà máy (<lý do>)`. Không so, không issue, không RFI, không ASK_USER.

## Làm, cho từng mã trong scope
1. So 3 nguồn: **kích thước** (BoQ / BV), **KL** (BOQ_MAP từng khu), **vật liệu + mã hoàn thiện** (BoQ / ML / nhãn trên BV), **phạm vi** (phần ngoài phạm vi theo PROJECT `Phạm vi tính giá`). Mỗi chỗ lệch chưa có issue → thêm issue `D…`.
2. `Mã hoàn thiện` = mã sẽ tính + nguồn, vd `WD-02 (BV p13; ML ghi WD-01 → D03)`. Mã ghi trên bản vẽ tại chỗ đó được ưu tiên hơn mã của block cùng loại.
3. `Giả định tính giá`: khi còn lệch, ghi đang tạm tính theo gì. **Không tự quyết bên nào đúng**: lệch BV ↔ BoQ là việc hỏi CĐT (RFI); giả định ghi rõ để RFI nhắc lại.
4. `Độ tin cậy`:
   - **CAO**: KL khớp; đủ kích thước; trường quyết định giá không có `KHÔNG THẤY`; không còn `MÂU THUẪN` chưa có giả định; **Gemini khớp ở các trường quyết định giá** (kích thước, số cánh/ngăn kéo/đợt, vật liệu) theo dòng `Kiểm chéo Gemini`.
   - **TB**: còn lệch nhưng đã có giả định + issue/RFI; hoặc Gemini chưa chạy / lệch ở trường phụ.
   - **THAP**: thiếu tờ; trường quyết định giá `KHÔNG THẤY`; chưa xem hình; tờ có thể lẫn type/khu; hoặc Claude và Gemini vẫn lệch ở trường quyết định giá.
5. BOQ_MAP: `Trạng thái` KHỚP / LỆCH / CHƯA ĐỌC cho từng khu.

## ISSUES, RFI, câu hỏi
- Mọi issue (B, S, V, D): điền `Hành động` (vd `Gửi RFI-02; tạm tính theo BV`), `RFI`, `Trạng thái = MỞ`.
- **RFI_LOG**: gom issue cùng chủ đề thành 1 câu hỏi gửi CĐT (`RFI-01`…), `Trạng thái = Nháp`. Chỉ câu hỏi ngắn; file RFI đầy đủ do skill 04-rfi soạn.
- **ASK_USER**: quyết định nội bộ cần user (tính hay loại trừ, chọn phương án, giả định ảnh hưởng giá lớn). Mỗi câu kèm `Vì sao cần`. Không hỏi điều đã có trả lời trong `ASK_USER.user.json`.

## Xong khi
Mọi mã có `Độ tin cậy`; mọi issue có `Hành động`; `check` (không `--den`) không còn LỖI ở file của bạn. Trả về: số mã CAO / TB / THAP, số RFI, số câu ASK_USER.
