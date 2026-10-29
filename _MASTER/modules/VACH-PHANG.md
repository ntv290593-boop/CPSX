# VÁCH, TRẦN PHẲNG (Template#10)

Nguồn: Template "VÁCH, TẤM PHẲNG" hàng 734–789 (đọc công thức 2026-10-24). Giá theo **m²**: `N = H/F`, `F = B×D×10⁻⁶` (header 1000×1000 = 1 m²), BoQ link `'Chiet tinh'!N` (R0064).

## 1. Ô nhập
Header: `B`, `D` (rộng, cao vách), `C` = sâu/dày (50), `E` = SL (m² quy đổi: `E = AG/F`, R0064).
Dòng module "VÁCH, TRẦN": `B`=18, `E`=1, **`N` = loại vách** (công tắc), `K`="c" (nhãn để SUMIF keo).
`N` nhận 2 giá trị: `VÁCH HOÀN THIỆN GỖ - SƠN-GIẤY DÁN TƯỜNG` | `VÁCH PHẲNG - BỌC VẢI`.

## 2. Chi tiết (công tắc N)
| Chi tiết | Rộng × Dài | SL | Vật liệu |
|---|---|---|---|
| Tấm hoàn thiện | C × D (bằng header) | E | `IF(bọc vải,"9mdf","6mdf")`, M=1, N=mã HT (WD-01…) |
| Tấm nền | = trên | E | 12mdf, M=1, N=sl |
| NAN / Viền tạo hình | 20×20 | gỗ: `ROUND(C/(C nan+20)+1)×E×0` (**tắt `*0`**); vải: viền `(C+D)×2`, SL=E | gỗ: 25ASH; vải: 6MDF |
| Khung xương gỗ (ashc) | 20×50 | `ROUNDUP(D/500+1)×ROUNDUP(D'/500+1)×E` (bước 500) | ashc, 3 hàng, E module = 0 mặc định |

## 3. Vật tư theo công tắc
- **Bọc vải:** Foam (K=`fd`, dày 20, SL = SUMIF "bọc vải"), Vải (K=`FB.01`; khổ 1370: `C = D = MAX(C+2×dày foam+80)`, SL = `MIN(ROUNDUP(D/1370),ROUNDUP(C/1370))×E`), vải lót `kate` (xem BOC-VAI-FOAM).
- **Khung sắt:** Vật tư cơ khí (`MT.20401.2`) nhân `*0` (tắt); khung tính ở **Kim loại gia công ngoài** MTGC1 20×40×1,2 hộp, bước 500 (xem KIM-LOAI-GIA-CONG; R0034).
- Kính ốp (10KCL, G=C×D×E×10⁻⁶), Len inox (SS.01, G=D×E×10⁻³): để sẵn, mở khi dự án có.

## 4. Phụ trợ trong block
`VTPK` G = `0,1×F` (gói/m²); đóng gói `DG.V` G = `((B+D)×2×C + D×B×2)×10⁻⁶×E` (m² thùng); KEOEP G = `SUMIF(K,"c",G)` (m² tấm ghép keo); `TK+VIT` G = `9×F` nếu có khung; KEOCANH theo cột BJ.

## 5. Nhân công (giữ như mẫu, R0150)
PX MỘC = `Σm² tấm nền + tấm HT × Dinhmuc!D7 + Σnan × Dinhmuc!D14`; PX SƠN theo AA (SONPU×D19, sơn khác×D20, lót SL×D21); CƠ KHÍ = `ΣG cơ khí/25`; BỐC XẾP = `ΣBBPL × D37`.

## 6. Số cứng
500 (bước xương), 1370 (khổ vải), 80 (dư vải mỗi bên), 20 (nan), 9×F (vít).

## 7. Rule
R0034, R0035 (F017 dán nguyên bản), R0064, R0065 (nan: tên bắt đầu "Nan ", nhân công SUMIF "Nan*"), R0154 (VNTAP), R0039.

## 8. Takeoff cần khai
Diện tích m², loại vách (hoàn thiện gỗ/sơn/giấy hay bọc vải), mã hoàn thiện + loại tấm hoàn thiện (6 hay 9mdf), có khung sắt/xương gỗ, có nan (kích thước, bước), kính/inox ốp.

## 9. Chưa xác nhận
Chuỗi `"VÁCH HOÀN THIỆN GỖ - SƠN-GIẤY DÁN TƯỜNG"` là tên công tắc dài, nên lấy từ dropdown (Linhtinh?). Chưa đọc dropdown nguồn.
