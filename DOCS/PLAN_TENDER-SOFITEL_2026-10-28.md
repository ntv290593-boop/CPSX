# PLAN – Scan hồ sơ tender SOFITEL (thư mục `new`) – 2026-10-28

Trạng thái: **đã thảo luận, chưa thực thi.** Phạm vi: chỉ `_ACTIVE\SOFITEL\01. INPUT\new`. Phục vụ phòng tender.

## Đã chốt với user (28/10)
1. ACTIVE.md chuyển sang SOFITEL; NORTON, Blanca đã sang archive.
2. Main Lobby **không có sheet trong BOQ** → tạm bỏ qua, chỉ ghi chú, kiểm lại sau.
3. Đồng ý tạo `_ACTIVE\SOFITEL\10_TENDER` và gộp các sổ (TCQ_LOG / MAT / DWG / SCOPE) chung workbook BOQ.
4. Các link OneDrive/Drive trong TCQ **không tải**, chỉ ghi "chưa có" trong FILE_INDEX.

## 1. BOQ và vùng
Nguồn: `new\1.261006_HSMT\Editable file\07-08 SSP-ID package-Tender BOQ.xlsx`
- `2. Finishing-L4-17`, `3. Furniture-L4-17`: cột G–M = SK-01, SK-02, ST-01, GK-01, GT-01, CR-01, CR-01A; N–P hành lang L4/L5/L6-17; Q sảnh thang; R cabin.
- `2.1 Finishing-L18-20`, `3.1 Furniture-suite-L18-20`: G–L = PS-01, PS-02, PS-02A, JS-01, SK-02(L18), SK-02A(L19); hành lang/sảnh thang L18/19/20.
- T–W = Base, AL–AP = VE; ghi chú AJ, AP. `4. Provisional sum` = 5.700.000.000 (H19). `Room schedule` tổng 274 (D25).
- **Bẫy:** số Area của IDD khác bộ FF&E/thảm (IDD Area 51 = Lift Lobby; thảm AREA51 = Standard King & Twin). Mọi thứ gắn theo **mã phòng**.

## 2. Vai trò file
| Vai trò | File |
|---|---|
| Nguồn chính | `HSMT\10. APPENDICES\Drawings\Design Drawing\OneDrive_1_6-10-2026.zip` (Phase 1 IDD 23/06, Phase 2 Suite 12/06, FF&E, Finishes Schedule 23/06) |
| Bổ sung/thay thế | ADD01 Main Lobby; TCQ04 Carpet spec Rev0; TCQ05_UP FF&E Area 51/52 19/10; Facade Drawing |
| Làm rõ | TCQ07 xlsx (gộp No1–7); `new\SSP-ID works-Tender queries TCQ_takenote nội bộ.xlsx` (dịch VN F/G, ghi chú H, đến TCQ06) |
| Tham khảo | As-built 2012, VE list, schedule Rev05/06, nội quy, logistic |
| Ngoài scope | MEP.zip, LIGHTING.zip, Annexes |

## 3. Phạm vi (scan toàn gói, khoanh vùng bằng nhãn, không bỏ dòng)
- Loại trừ: Lobby Bar + L'Olivier (TCQ02-1); deck, bồn cây, vách cây (TCQ02-14/15, TCQ03-32/33); cửa 850×750 (TCQ03-37); ban công (TCQ04-9); giường (TCQ04-14); Artwork + OSE (TCQ05-34); 50-501 omit (TCQ05-24, TCQ06-2); Presidential/Apartment/MUR (Room schedule).
- MEP: điện, đèn, LED kể cả LED trong built-in; cấp + lắp thiết bị vệ sinh (TCQ03-43); khoan lỗ ống vệ sinh (TCQ03-1).
- Owner: thanh lý đồ rời + thiết bị vệ sinh cũ (TCQ03-30/31).
- Bao gồm: mosaic 50-803 cabin (TCQ02-7); biển chỉ dẫn tầng (TCQ03-42); Main Lobby theo bản vẽ, không mezzanine (TCQ06-17).

## 4. Cột bổ sung trong BOQ
- **A – chính thức gửi CĐT** (được chấp thuận ở TCQ04-19), sau AP: Mô tả NT – Base; Mô tả NT – Đề xuất; KL – Đề xuất.
- **B – nội bộ** (nhóm/ẩn, bỏ khi xuất): Mã VL · Bản vẽ hiệu lực · Rev/ngày/đợt · Mã spec/FF&E · Nhóm thực hiện (Nhà máy/NTP/MEP/Owner) · Trạng thái (Base/VE/Loại trừ/Omit/Chờ) · Mã TCQ · Nội dung cuối (VN) · Loại thay đổi (KL/VL/Bản vẽ/Phạm vi) · Việc còn mở · Ngày cập nhật.
- Sổ trong cùng workbook: **TCQ_LOG** (mỗi câu 1 dòng, cột "bị thay bởi"), **MAT** (mỗi mã VL 1 dòng, trạng thái cuối), **DWG** (mỗi số bản vẽ, rev, thay bởi), **SCOPE**. Cột B trong BOQ tra từ các sổ này.
- Ví dụ bị đè: 50-501 (TCQ04-22 → TCQ05-24 Omit); kính cửa lùa (TCQ04-21 GL-02 → TCQ05-35 ribbed như MUR); rèm (TCQ03-14: theo spec động cơ, rèm tay = VE).
- Thứ tự ưu tiên: **Spec → Bản vẽ → BOQ** (TCQ04-3).

## 5. Cấu trúc thư mục
```
SOFITEL\
  01. INPUT\new\            ← chỉ đọc
  10_TENDER\
    README.md
    01_index\FILE_INDEX.xlsx
    02_extract\<mã phòng/vùng>.md
    03_boq\SSP_BOQ_TRACK_<yymmdd>.xlsx
    04_out\
```
Quy tắc LLM: đọc README + FILE_INDEX trước; xử lý từng vùng; theo mã phòng; bỏ `~$…`, `plot.log`; thư mục `01. INPUT` gốc cũ = bị thay bởi `new`.

## Ghi chú – kiểm sau
- **Main Lobby:** có bản vẽ ADD01 (`new\5.261014_ADD01\MAIN LOBBY\`) + TCQ05-9, TCQ06-5/7/8/13–17, nhưng BOQ không có sheet. Chưa quyết: tự lập BOQ hay chờ addendum BOQ.
- **Link chưa tải ("chưa có"):** spec rèm cập nhật (TCQ06-21); as-built Main Lobby (TCQ06-16); thiết kế thảm hành lang (TCQ04-24); bản vẽ đèn minibar (TCQ05-30); bản vẽ 2012 + MEP/lighting (TCQ02-12, TCQ03-6/20 – một phần đã có trong `new`).
- Câu còn mở: TCQ04-6 (E1/E2 chưa trả lời), TCQ04-16 "to be informed later" (đã có ở TCQ05-30), TCQ06-4 biển chỉ dẫn tầng "deferred", TCQ06-6 ST-07 mô tả "to be confirmed".
- Room schedule: TCQ05-13 giải thích SK-02(L18) và SK-02A(L19); L4 cần đối chiếu.

## Bước thực thi (phiên sau)
Không dùng skill `01-scan` (phạm vi nhà máy, form PROJECT_DATA). Session sau **tạo skill tender riêng** — xem `DOCS\HANDOFF_SKILL-TENDER_2026-10-28.md`; các bước dưới đây là dữ liệu chạy thử đầu tiên của skill đó.

1. Tạo `10_TENDER` + README.
2. FILE_INDEX cho toàn bộ `new` (liệt kê cả nội dung zip, không giải nén hết).
3. TCQ_LOG từ TCQ07 + bản dịch takenote nội bộ.
4. Bản sao BOQ + cột A/B + sổ MAT/DWG/SCOPE.
5. Scan từng vùng theo mã phòng → `02_extract`, điền cột B.
