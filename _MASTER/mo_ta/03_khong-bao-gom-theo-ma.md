# 03 — "Không bao gồm" theo mã vật liệu

Mã trong dòng BoQ / trên bản vẽ thuộc nhóm ngoài phạm vi tính giá (dòng "Phạm vi tính giá" ở PROJECT của dự án) → thêm 1 dòng `+ …` dưới `Không bao gồm:`. Phạm vi theo dự án: SOFITEL bỏ kim loại, đá, giấy WC, gương/kính, OS&E, đèn; dự án khác MẶC ĐỊNH tính kim loại (chỉ bỏ khi có yêu cầu).

| Nhóm mã | Câu (dùng mã cụ thể của dòng) | Ví dụ user viết |
|---|---|---|
| WC-xx giấy/vải dán tường | `+ Giấy dán tường <WC-xx>` (vinyl thì `+ Giấy dán tường vinyl <WC-xx>`) | F086 |
| MR-xx gương | `+ Gương <MR-xx>` | F043, F082 (`+ Gương`) |
| MT-xx kim loại | `+ Kim loại <MT-xx>` / `+ Nẹp kim loại <MT-xx>` / `+ Inox <MT-xx> (ốp và viền kim loại)` | F029, F043 |
| ST-xx đá | `+ Đá` / `+ Mặt đá <tên> (polished)` | F038, L014 |
| GL-xx kính | `+ Kính <mô tả, dày>` | L014, F041 |
| FB-xx vải | `+ Vải <mã>: <md>` (ghi số md) | L006, L010 |
| Da | `+ Da bọc <vị trí> (<mã>): <SF>/cái …` (công thức theo Chiet tinh) | L022 |
| Đèn, LED | `+ Đèn LED` / `Không bao gồm led, kim loại` | F035, L022 |
| MEP, thiết bị, OS&E | `+ MEP`, `+ Thiết bị`, `+ Phụ kiện thoát nhiệt` | F038, F041 |
| Phụ kiện, tay nắm, thanh treo | `+ Tay nắm.`, `+ Thanh treo áo.` | F038 |
| Kết cấu do bên khác làm | `+ Kết cấu bậc trần (khung/thạch cao)`, `+ Tấm thạch cao trần`, `+ Tấm trần` | F083, F087 |
| Cửa | mục `Chưa bao gồm phụ kiện cửa + nhân công lắp đặt: …` | D01–D14 |

## Lệch văn phong đã thấy (chưa sửa — user quyết)
- F021 (len **sơn PT-03**) ghi "Gỗ Ash sơn PU hoàn thiện 18x80mm" — giống len gỗ F012, nghi chép nhầm.
- F034 viết hoa "KO BAO GỒM GƯƠNG" (các dòng khác: "Không bao gồm:" + "+ …").
- F010 "Không bao gồm Gỗ viền" và F048 "Không bao gồm kim loại" viết cùng dòng, không xuống dòng "+".
- Cửa dùng "Chưa bao gồm", đồ khác dùng "Không bao gồm".
- F036 có mô tả nhưng chưa có đơn giá (AH trống) ở RE00.
