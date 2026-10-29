"""Mô tả AJ cho F055, F056, F057 (vách tivi = tủ cao + vách tivi có ô trống + 2 kệ, user dựng 2 block/dòng). Theo quy tắc mô tả: cái được tính + Không bao gồm."""
import sys, openpyxl
sys.stdout.reconfigure(encoding="utf-8")
NL = "\n"
def mo_ta(L, D, tw, td, vw, vd, tu):
    return NL.join([f"Kích thước L{L}xD{D}xH2500 mm :",
        f"Tủ cao W{tw}xD{td}xH2500: thùng tủ MDF dày 17mm (hậu 12mm) hoàn thiện WD-04, {tu}.",
        f"Vách tivi tấm MDF dày 17mm hoàn thiện PT-03, W{vw}xD{vd}.",
        "Kệ ghép hộp dày 40mm hoàn thiện PT-03 (2 kệ trong ô tivi).", "Khung xương gỗ", "Hậu khu vực tivi MDF 17mm melamine màu đen",
        "Không bao gồm:", "+ Giấy dán tường WC-05", "+ Inox MT-02 (ốp và viền kim loại)"])
ND = "1 ngăn kéo MDF (mặt ngoài 17mm, thùng 12mm)"
D = [(64, "F055", mo_ta(2280, 340, 450, 340, 1830, 200, ND), "tủ cao 7017 + vách tivi 7097"),
     (65, "F056", mo_ta(2000, 280, 300, 280, 1700, 200, ND), "tủ cao 7180 + vách tivi 7260"),
     (66, "F057", mo_ta(2525, 290, 805, 290, 1720, 180, "2 cánh tủ MDF dày 17mm"), "tủ cao 6862 + vách tivi 6934")]
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "MO_TA"; ws.append(["Dòng BoQ", "Mã", "Mô tả costing đề xuất (cột AJ)", "Căn cứ", "cmt"])
for r in D: ws.append(list(r) + [None])
wb.save(r"D:\COSTING EXCEL\_ACTIVE\SOFITEL\02_takeoff\MO_TA_tv.xlsx"); print("ok")
