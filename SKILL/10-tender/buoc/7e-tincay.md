# Bước 7e — tincay: đối chiếu BOQ ↔ bản vẽ ↔ spec, chốt độ tin cậy

**Ghi (file `*.tincay.json`):** `SCOPE` (`Mã hoàn thiện`, `Issue`, `Độ tin cậy`, `Giả định tính giá`), `BOQ_MAP` (`Issue`, `Trạng thái`, `Độ tin cậy`), `ISSUES` (`Hành động`, `RFI`, `Trạng thái` cho mọi issue; issue mới ID `D01`…), `RFI_LOG`, `ASK_USER`.
**Không làm:** đọc lại bản vẽ để đổi số của bước 7d. Thấy nghi sai → issue `D…` (`Loại = Cần đọc lại`), không sửa. Không đổi `Trạng thái` của SCOPE (bước 7 quyết định phạm vi), không đụng 3 cột gửi CĐT.

## Chạy tool trước
`py -I "<TOOL>" tincay <GOC>` → `SCOPE.tincay-tool.json` (`Độ tin cậy`, `Issue` từng dòng), `BOQ_MAP.tincay-tool.json` (`Trạng thái` KHỚP/LỆCH/CHƯA ĐỌC, `Độ tin cậy`), `ISSUES.tincay-tool.json` (`DT01`…: chỉ cho cụm chưa có issue nào), `QUAN_SAT.tincay-tool.json` (dòng `Kiểm chéo Gemini số`) và `_tincay.json` (lý do từng cụm). Trên Blanca, bản 01-scan của tool có Độ tin cậy trùng LLM 91% và **lạc quan hơn** LLM ở 8/13 ca lệch, nên phải rà cả cụm tool chấm CAO, không coi nháp là kết quả.
**Việc của bạn:** đọc `_tincay.json`, chỉ ghi những ô bạn đổi hoặc bổ sung. Ô bạn ghi thì thắng nháp tool. Dòng "CHUỖI SAI" tool in ra → issue `Cần đọc lại`.

Mỗi cụm ghi ở dòng đại diện: `Mã hoàn thiện`, `Giả định tính giá` tự lan cho dòng cùng cụm. `Độ tin cậy` tool ghi từng dòng: nâng/hạ thì ghi cho mọi `Khóa` của cụm (`_loc` theo `Cụm đọc`).

## Làm, cho từng cụm
1. So 3 nguồn: **kích thước** (BOQ / BV), **KL** (BOQ_MAP từng khu), **vật liệu + mã hoàn thiện** (BOQ / MAT / nhãn trên BV), **phạm vi** (`Trạng thái`, TCQ còn hiệu lực). Mỗi chỗ lệch chưa có issue → thêm issue `D…`.
2. `Mã hoàn thiện` = mã sẽ tính + nguồn, vd `WD-02 (BV p13; MAT ghi WD-01 → D03)`. Mã ghi trên bản vẽ tại chỗ đó được ưu tiên hơn mã của block cùng loại.
3. `Giả định tính giá`: khi còn lệch, ghi đang tạm tính theo gì. **Không tự quyết bên nào đúng**: lệch BV ↔ BOQ là việc hỏi CĐT (RFI hoặc TCQ); giả định ghi rõ để RFI nhắc lại.
4. `Độ tin cậy`:
   - **CAO**: KL khớp; đủ kích thước; trường quyết định giá không có `KHÔNG THẤY`; không còn `MÂU THUẪN` chưa có giả định; **Gemini khớp ở các trường quyết định giá**: kích thước theo `Kiểm chéo Gemini số`, số cánh/ngăn kéo/đợt, vật liệu theo `Kiểm chéo Gemini`.
   - **TB**: còn lệch nhưng đã có giả định + issue/RFI; hoặc Gemini chưa chạy / chưa so số / lệch ở trường phụ; hoặc KL m2/md chưa đếm trên bản vẽ.
   - **THAP**: thiếu tờ; trường quyết định giá `KHÔNG THẤY`; chưa xem hình; tờ có thể lẫn type/khu; hoặc Claude và Gemini vẫn lệch ở trường quyết định giá. `Kiểm chéo Gemini số` = `chỉ Claude: …` chỉ cho biết Gemini không đọc ra số đó, chưa cho biết bên nào đúng: chỉ nâng mức khi có căn cứ, ghi lý do vào `Giả định tính giá`.
5. Mọi dòng BOQ_MAP có `Trạng thái`.

## ISSUES, RFI, câu hỏi
- Mọi issue (S, V, D, DT): `Hành động` (vd `Gửi RFI-02; tạm tính theo BV`), `RFI`, `Trạng thái = MỞ`.
- **RFI_LOG**: gom issue cùng chủ đề thành 1 câu hỏi gửi CĐT (`RFI-01`…), `Trạng thái = Nháp`. Câu đã có trong TCQ còn hiệu lực thì không hỏi lại, dẫn mã TCQ. **Không tự gửi.**
- **ASK_USER**: quyết định nội bộ cần user (tính hay loại trừ, chọn phương án, giả định ảnh hưởng giá lớn). Mỗi câu kèm `Vì sao cần`. Không hỏi lại điều đã có trong `ASK_USER.user.json`.

## Xong khi
Mọi dòng đọc sâu có `Độ tin cậy`; mọi issue có `Hành động`; `check` (không `--den`) không còn LỖI ở file của bạn. Trả về: số cụm CAO / TB / THAP, số RFI, số câu ASK_USER.
