# SOFA / BENCH / ARMCHAIR / ĐÔN (Template#15)

Nguồn: Template "Bench, Armchair, Ghế đôn" hàng 1018–1118. Giá theo **md** (`N = H/B×10³`). Khung = các **GHÉP HỘP** (mê ngồi, tay, lưng) + bọc foam/vải chạy sơ đồ.

## 1. Ô nhập (header, công tắc ảnh hưởng nhân công)
`B,C,D` = dài, sâu, cao; `E` SL; `J` = ván nền (PLW); `K` = `Thẳng` | `Cong` (Cong: nhân công ×1,5); `K1019` = `Sofa/Armchair` | `Armchair` | `Đôn` (nhân công sofa ×0,75 armchair, ×0,3 đôn); `L` = `Khó 1` | `Khó 2` (×1,6 | ×2,1, mặc định ×1,1); `B1049` = `Mê fixed` | khác (mê rời).

## 2. Khung gỗ (ghép hộp)
| Module | Kích thước | Ghi chú |
|---|---|---|
| GHÉP HỘP - MÊ NGỒI | `B`=135 (dày), `C = sâu header − dày hộp lưng`, `D = dài − dày tay × E tay`, `E` nhập (1) | tấm mặt 18, tấm viền 12 (**gõ cứng**), K=J header (PLW), xương gỗ 99×50 / 30×30 |
| GHÉP HỘP - TAY | `B`=160 (dày), `C`=sâu header, `D`=545, `E`=2 | tấm mặt 12, viền 12, xương 30×30 |
| GHÉP HỘP - LƯNG | `B`=245, `C = dài hộp mê`, `D`=545, `E`=1 | tương tự |
| VIỀN CHÂN GỖ | `B`=100, `C`=sâu, `D`=dài | Chân gỗ 38ash (`D=(C+D)×2`); Xương giằng thông 30×50 (`ROUND(D/500+1)`); Ke góc 30×120 |
SL xương/tấm của từng hộp: `ROUNDUP(D/400+1)×ROUNDUP(D'/400+1)×E` như GHEP-HOP.

## 3. Foam (Vật tư Sofa)
Cụm foam dùng lại kích thước hộp, **dày gõ cứng**: Đệm ngồi 175, viền 20; mặt mê đế 10 (chỉ khi mê rời: `E=IF(B1049="Mê fixed",0,1)`); tay 20; lưng 50/10/5. K = `fd`. Gối lưng rời (gòn tơi `gontoi`) 200×290, gối trang trí 170×360×360.

## 4. Vải
Mỗi chi tiết foam có dòng vải = kích thước foam **+100** (mỗi chiều); gối ×2 mặt. `Vải chạy sơ đồ cắt` K=`FB.01`, khổ 1370, `D = ΣG vải / AG` (AG = diện tích cắt theo khổ); vải lót `kate` cùng kích thước (xem BOC-VAI-FOAM).

## 5. Nhân công
PX MỘC: `C=5`, `D=B/4+1`; PX SOFA: `C=4`, `D=B/5+1`; mỗi dòng `E = C×D/B × hệ số (Khó) × (Cong) × (loại)`. Đây là **thông số từ header**, không phải số gõ tay.

## 6. Rule
R0091 (mẫu), R0150 (nhân công giữ như mẫu, chỉ đổi qua công tắc header), R0074, R0052(d) (vải: kiểm có tính tiền hay "Không bao gồm").

## 7. Chưa xác nhận
`AG1083` (mẫu số tính chiều dài vải chạy) là ô nào và tính ra sao; Armchair dùng chung block với Bench nhưng mê/tay/lưng có SL (E) riêng phải nhập tay.
