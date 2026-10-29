"""Phiếu bóc VANITY (NORTON đợt 261024) v2 -> 02_takeoff\\PHIEU_BOC_VANITY_v2.xlsx.
v2 = v1 + câu trả lời của user 26/10 (CHO_USER cột D v1) + đo bản vẽ (mặt cắt/mặt đứng p12-13) + kiểm chéo Gemini.
Nguồn tham số cho cả phiếu bóc lẫn script dựng COM (s3_build)."""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
sys.stdout.reconfigure(encoding="utf-8")
OUT = r"D:\COSTING EXCEL\_ACTIVE\NORTON\02_takeoff\PHIEU_BOC_VANITY_v2.xlsx"
H = Font(bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="1F4E78")
wb = Workbook(); wb.remove(wb.active)
def sheet(n, cols, w):
    ws = wb.create_sheet(n); ws.append(cols)
    for c in ws[1]: c.font = H; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, x in enumerate(w, 1): ws.column_dimensions[get_column_letter(i)].width = x
    ws.freeze_panes = "A2"; return ws
def add(ws, row):
    ws.append(list(row))
    for c in ws[ws.max_row]:
        if isinstance(c.value, str) and c.value.startswith("="): c.data_type = "s"
        c.alignment = Alignment(wrap_text=True, vertical="top")

ws = sheet("BLOCK", ["BlockId", "Mã BoQ", "Hạng mục", "SL (nr)", "Dòng BoQ (5 bill)", "Mẫu", "W (B)", "D (C)", "H (D)", "Hoàn thiện", "Nguồn", "Độ tin cậy", "Ghi chú"], [9, 7, 22, 8, 28, 40, 7, 7, 7, 40, 24, 9, 50])
BQR = "B5.16!{r} · B6.16!{r} · B8.18!{r} · B9.18!{r} · B10.18!{r}"
add(ws, ["V01-1", "V01", "Tủ lavabo", 513, BQR.format(r=11), "Template#7 Vanity (user chốt)", 900, 550, 440,
         "PA1 LA-02 · PA2 LA-03 (cốt WPB) · PA3 ME-01 (cốt MDF E1 chống ẩm tạm tính)", "BV p12-13 (đo), BoQ, ML 001, user", "TB",
         "Thân tủ 2 cánh 700 + THÙNG FIX 200 (đo mặt đứng p12: panel trái ≈200, 2 cánh ≈350). Header D=440 vì D558=0 (bỏ dải khung 150 của mẫu: phần đó thuộc V02). Không ngăn kéo, không len chân."])
add(ws, ["V02-1", "V02", "Bàn đá (gồm khung)", 513, BQR.format(r=12), "Template#7 Vanity, tắt phần tủ + dòng bổ sung", 900, 550, 150,
         "SS02 Himacs/Krion M553 12mm (đơn giá user điền); cemboard; xương đỡ đá MTGC1 20×40×1.2 hộp", "user 26/10, BV p12-13, ML 001", "TB",
         "Nóc cemboard đã có trong V01 (module tủ); V02 thêm viền cemboard trước 900×100 + 2 hông 550×100 (user: cao 100 giả định); xương đỡ đá theo Template hàng 771-774 (MTGC1). SS02: mặt 900×550 + diềm trước 900×220 (đo p12: dải đá ≈220)."])
add(ws, ["V03-1", "V03", "Gương MCB", 513, BQR.format(r=16), "Template#8 Tủ gương hộp", 900, 180, 1100,
         "Hộp theo phương án tủ lavabo; 2 cánh gương", "user 26/10, BV p12-13, ML 003", "TB",
         "Kích thước user chốt 900×180×1100 (BV rộng/cao, sâu 180 BoQ/ML). Gương 5mm clear (5MR) tạm tính + mài cạnh GL.MAI + dán GL.LD."])

ws = sheet("MODULE", ["BlockId", "Hàng (Off từ header)", "Module (cột A mẫu)", "Cột", "Giá trị", "Loại ô", "Nguồn", "Ghi chú"], [9, 12, 26, 6, 52, 9, 24, 60])
M = [
 ("V01-1", 0, "hdr", "B,C,D", "900, 550, 440", "NHAP", "BoQ/BV", "1 đơn vị BoQ = 1 cái"),
 ("V01-1", 0, "hdr", "E", "=SUM(D11 của 5 sheet bill)", "LINK", "BoQ", "E block = Σ SL BoQ (513)"),
 ("V01-1", 49, "Khung xương (dải 150 cơ khí)", "D558→D49+...", "D (dải khung) = 0", "NHAP", "đề xuất", "đặt hàng 'khung' D=0: tủ cao 440; dải 150 + khung chuyển sang V02"),
 ("V01-1", 2, "TỦ", "B,K,L", "B=700 (2 cánh × 350), K='Tủ treo', L='0 lộ hông'", "NHAP", "BV p12 (đo)", ""),
 ("V01-1", 21, "THÙNG FIX", "B,E", "B=200, E=1 (C,D theo header 550×440)", "NHAP", "BV p12 (đo) + user", "thùng fix cạnh thùng 2 cánh"),
 ("V01-1", 13, "Ngăn kéo", "E", "0", "NHAP", "BV", ""),
 ("V01-1", 30, "LEN CHÂN", "E", "0", "NHAP", "BV", "treo tường"),
 ("V01-1", "*", "Ván K (17MDF/12MDF/25MDF)", "K", "=IF(PA_LAVABO=\"PA3\",\"nnMDF\",\"nnWPB\") (nn = 17/12/25; 25WPB chưa có giá → R0074 dùng 25→ giữ MDF nếu >800)", "LINK", "ML 001 / R0074", "WPB theo ML; PA3 MDF"),
 ("V01-1", "*", "Mã hoàn thiện N (LA-01)", "N", "=MA_LAVABO", "LINK", "mail 24/10", "OverView!C52 chọn PA1/PA2/PA3"),
 ("V01-1", "-", "Tay nắm", "-", "Không bao gồm (ghi vào mô tả)", "-", "user 26/10", ""),
 ("V02-1", 0, "hdr", "B,C,D", "900, 550, 150", "NHAP", "BV", "D chỉ là nhãn (dải khung đỡ đá)"),
 ("V02-1", 2, "TỦ", "E", "0", "NHAP", "đề xuất", "tắt toàn bộ phần tủ (con E=E TỦ)"),
 ("V02-1", "+", "Cemboard viền trước", "C,D,E,K", "900×100, E=1, K=12CEM (hoặc 9CEM)", "THÊM DÒNG", "user 26/10", "dòng bổ sung ở 'Vật tư, phụ kiện khác'"),
 ("V02-1", "+", "Cemboard hông", "C,D,E,K", "550×100, E=2, K=12CEM", "THÊM DÒNG", "user 26/10", ""),
 ("V02-1", "+", "SS02 mặt bàn", "C,D,E,K", "900×550, E=1, K=SS02 (12mm)", "THÊM DÒNG", "ML 001", "đơn giá SS02 = Material!F (ô vàng, user điền)"),
 ("V02-1", "+", "SS02 diềm trước", "C,D,E,K", "900×220, E=1, K=SS02", "THÊM DÒNG", "BV p12 (đo)", "user giả định 100 cho cemboard; SS02 diềm theo hình ≈220"),
 ("V02-1", 66, "Kim loại gia công ngoài (hàng 771-774)", "B,C,D,E,K,M,N", "Khung xương ngang 20×40×1.2 hộp MTGC1 (L=900), đứng (L=550)", "NHAP", "user 26/10", "tham khảo Template hàng 771-774; số thanh: 3 ngang + 4 dọc (giả định)"),
 ("V03-1", 0, "hdr", "B,C,D", "900, 180, 1100", "NHAP", "user 26/10", ""),
 ("V03-1", "*", "Cánh tủ", "E", "2", "NHAP", "user 26/10", "2 cánh gương"),
 ("V03-1", "*", "Ván K/N", "K,N", "như V01 (PA selector)", "LINK", "user 26/10", "hộp gương cùng phương án"),
 ("V03-1", "+", "Gương 5MR", "C,D,E", "900×1100, E=1", "GIỮ MẪU", "ML 003", "mẫu có sẵn dòng 5mr"),
 ("V03-1", "+", "Mài cạnh GL.MAI", "G", "chu vi = 2×(900+1100)/1000 md", "NHAP", "BoQ 'polished edges'", "thay dòng KG.01 trống của mẫu"),
]
for r in M: add(ws, r)

ws = sheet("GIA_DINH", ["#", "Giả định tính giá", "Căn cứ", "RFI"], [4, 90, 40, 8])
for i, x in enumerate([
 ("Mặt bàn = SS02 solid surface 12mm", "User 26/10", "RFI-01"),
 ("Cốt tủ = WPB (PA1/PA2); PA3 = MDF E1 chống ẩm melamine ME-01 tạm tính", "ML 001; mail", "RFI-02/03"),
 ("Vanity 900×550 chung cho mọi loại căn (1BR A1.1a/A1.3a/A2.1a sâu 540); mặt bàn 850 (p221 ghi 870)", "BV 33 MB WC; Gemini kiểm chéo", "—"),
 ("Thân tủ 700 (2 cánh × 350) + thùng fix 200; cao 440 theo BoQ và mặt cắt", "BV p12 (đo tỉ lệ)", "—"),
 ("Gương 900×180×1100, 2 cánh; hoàn thiện hộp = phương án tủ lavabo", "User 26/10", "RFI-04"),
 ("Không bao gồm: tay nắm, chậu LA1, vòi FC1, xi-phông, bệ đá SS02 350×120", "User 26/10", "RFI-05"),
 ("Nhân công giữ như mẫu (R0150); đơn giá SS02, MTGC1, LA-02/LA-03, ME-01 do user nhập", "R0150; user 26/10", "—"),
 ("Cemboard viền cao 100 (user giả định); SS02 diềm trước 220 (đo hình, chưa có số trên BV)", "user 26/10; BV p12", "RFI-04"),
], 1): add(ws, [i] + list(x))

ws = sheet("CHO_USER", ["#", "Điểm", "Quyết định (user 26/10)", "Còn cần xác nhận"], [4, 60, 70, 60])
for i, x in enumerate([
 ("Mẫu tủ lavabo", "Template#7 Vanity", "—"),
 ("Bàn đá + khung", "Nóc cemboard có sẵn ở module tủ; thêm viền trước + 2 hông cemboard cao 100; xương đỡ đá MTGC (hàng 771-774)", "SS02 diềm trước tôi đo ≈220 (không phải 100); khung 3 ngang + 4 dọc là giả định"),
 ("Thùng fix", "Có 1 thùng fix bên cạnh thùng 2 cánh", "Tôi đo: fix ≈200 rộng, 2 cánh 2×350; xác nhận"),
 ("Hộp gương", "Hoàn thiện theo PA tủ lavabo, 2 cánh", "—"),
 ("Kích thước gương", "900×180×1100", "—"),
 ("Tay nắm", "Không bao gồm", "—"),
 ("Bệ đá SS02 350×120", "Không bao gồm", "—"),
], 1): add(ws, [i] + list(x))
wb.move_sheet("CHO_USER", -3)
wb.save(OUT); print(wb.sheetnames)
