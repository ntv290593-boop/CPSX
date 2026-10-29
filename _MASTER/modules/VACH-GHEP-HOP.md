# VÁCH GHÉP HỘP / VÁCH TIVI (Template#11)

Nguồn: Template "Vách ghép hộp (vách đầu giường, vách tivi)" hàng 790–858. Giá theo **m²** (`N = H/F`). Cơ chế "Hộp" = GHEP-HOP.md với **K = GHÉP HỘP - XƯƠNG SẮT**.

## 1. Ô nhập
Header: `B` rộng 1000, `C` sâu 100, `D` cao 1600, `E` SL. Dòng "CABINET CÓ CÁNH CỬA" (module): B/C/D theo header, E=1.
Dòng "Hộp": `K` = `GHÉP HỘP - XƯƠNG SẮT` (hoặc GỖ), `B`=sâu, `C`=rộng, `D`=cao, `E`.
"Ô trống vách ti vi": `B`, `D` (ô trống, mặc định 800×800), `E`.

## 2. Hộp (K = XƯƠNG SẮT) — khác ghép hộp xương gỗ
| Chi tiết | Công thức | Ghi chú |
|---|---|---|
| Tấm mặt | C+20 × D+20, E, **6MDF**, M=1, N=mã HT (pt-01) | mặt sau `E×0` (tắt) |
| Tấm viền | B+20 × D+20 và B+20 × C+20; `IF(B<51,0,E×2)` | B=100 → 2 hàng ×2 |
| Tấm nền | 12mdf (thay xương gỗ), SL theo mặt trước/viền | xương gỗ = 0 |
Xương: **không có xương gỗ**; khung chịu lực là khung sắt (mục 4).

## 3. Ô trống vách ti vi
Viền đứng / viền ngang (6MDF, N=LA-01, SL = E ô trống ×2), Tấm nền (12mdf, M=1, N=sl, ×2), **Hậu vách tivi** (17mdf, 800×800, N=sl, N để trống theo R0040). **Trừ hoàn thiện** và **Trừ tấm nền** (E = −E ô trống): trừ diện tích tấm mặt/nền ngoài đã tính ở ô trống (E âm là chủ ý).

## 4. Khung sắt và các lớp
- **Kim loại gia công ngoài (MTGC1, 30×30×1,2, hộp):** ngang `ROUNDUP(D'/600+1)×ROUNDUP(D''/600+1)×E`, đứng cùng cách (bước 600). Khối lượng kg: `LET(hộp: (B+C)×2×M×D; Tấm: C×D×M; tròn đặc; Ống) × 0,00000785 × E × F` (xem KIM-LOAI-GIA-CONG). Khung cơ khí cũ `MT.30301.2` ở Vật tư cơ khí được E=0 (tránh tính trùng, R0034).
- **Lưỡi gà treo** `luoigaz`: `IF(E836>0,0,B/6000×2×ROUND(D/1200+1))` (chỉ khi không có khung sắt).
- **Bọc vải:** Ván bọc vải (plw 12), Foam bọc (fd 50), Vải FB.01 + kate (xem BOC-VAI-FOAM).

## 5. Nhân công
PX MỘC = `SUMPRODUCT(C,D,E của chi tiết) × 10⁻⁶ × Dinhmuc!D2`; PX SƠN như VACH-PHANG; còn lại giữ mẫu.

## 6. Số cứng
+20 (cắt gọt), 51 (viền), 600 (bước khung sắt), 1200, 6000, 800 (ô trống).

## 7. Rule
R0038, R0040 (F029/F030 Nr), R0039 (dùng Hậu/Trừ hoàn thiện/Trừ tấm nền/Ô trống có sẵn), R0155, R0034. Sau copy sang hạng mục khác: đặt lại A, B:D, E dòng đầu (R0040).

## 8. Takeoff cần khai
Rộng × cao × sâu vách, có ô trống tivi (kích thước), loại khung (sắt/gỗ), lớp mặt (mã HT, 6MDF), có bọc vải/foam, có đèn/khe đèn.

## 9. Cấu trúc: vách ghép hộp + HỘP ÂM nhét tivi (user chốt 2026-10-24)
Vách tivi = **một hộp ghép lớn (vách) bị khoét một hộp âm** để đặt TV. Template#11 tính theo 3 bước, đúng thứ tự:
1. **Tính vách như hộp đặc:** "Hộp" (K = XƯƠNG SẮT) cho cả khổ B × D: tấm mặt, tấm viền, tấm nền, khung sắt chạy suốt (mục 2, 4).
2. **Khoét hộp âm bằng số ÂM:** diện tích ô trống (B ô × D ô) bị **trừ** khỏi mặt và nền: dòng "Trừ hoàn thiện vách ngoài đã tính" (6MDF, E=−1) và "Trừ tấm nền vách ngoài đã tính" (12mdf, E=−1). E âm là chủ ý, đừng thay bằng 0.
3. **Dựng lại lòng hộp âm** từ 4 thành hốc + đáy hốc: Viền đứng ×2 (C × D ô), Viền ngang ×2 (C × B ô), mỗi thành có tấm nền 12mdf tương ứng (E×2), **Hậu vách tivi** 17mdf (B ô × D ô) làm đáy hốc. `C` của thành hốc = sâu hộp (header C), tức sâu hộp âm.
Công tắc: E ô trống = 0 → vách thường (không hộp âm).

**Quan hệ kích thước** (dòng "Ô trống vách ti vi", gõ B và D; C theo header): rộng ô = B, cao ô = D (mặc định 800×800). Chiều sâu hốc (C) mặc định = sâu hộp 100 (R0040: "thành hốc C=100").

**Ngoài hộp âm (mục tùy chọn theo hạng mục):** phào/khung quanh hốc và len chân chất liệu MT-02 = tấm SSGC (phào C=60, dài chu vi ô trống; R0040) nhưng F029/F030 **không bao gồm MT-02** (R0052a) → "Không bao gồm: + Inox MT-02". Khe đèn dấu: thành khe đèn 17mdf pt-01 (R0040), chưa có dòng riêng trong mẫu.

Bản vẽ user dán (Elevation/Section/Detail 04) khớp cấu trúc này: ô trống cao 845 (210 + 420 + 210), hộp xương dày 100 quanh hốc, hốc sâu khoảng 170–175 ở chỗ TV, hoàn thiện PT 03 cho vách và MT 02 cho viền/len chân.

## 10. Chưa xác nhận
1. Hốc trên bản vẽ sâu ~170–175, lớn hơn sâu hộp 100: khi sâu hốc ≠ sâu hộp, `C` của thành hốc có phải sửa tay (C của "Ô trống vách ti vi" đang `=C header`) không? Cần user.
2. Thành khe đèn 17mdf pt-01 (R0040) chưa có dòng riêng trong Template#11: thêm tay theo hạng mục (copy hàng gần nhất, chèn trước hàng đệm, R0039).
3. Chiều rộng ô trống không đọc được từ hình (chỉ có tổng 1310 và TV 1217×686): lấy từ bản vẽ/BoQ khi bóc.
