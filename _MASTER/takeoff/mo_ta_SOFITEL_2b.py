"""Mô tả costing (cột AJ) đợt 2b — F053, F054. Quy tắc: _MASTER\mo_ta\01 (chỉ cái được tính + Không bao gồm, KHÔNG ghi chú nội bộ/RFI).
Khung theo F029 (vách TV) + kệ ghép hộp. Ra 02_takeoff\MO_TA_2b.xlsx (cột Căn cứ = ghi chú nội bộ, không vào AJ)."""
import sys, openpyxl
from openpyxl.styles import Alignment
sys.stdout.reconfigure(encoding="utf-8")
NL = "\n"
def mo_ta(L, D, H):
    return (f"Kích thước L{L}xD{D}xH{H} mm :" + NL + "Vách tivi tấm MDF dày 17mm hoàn thiện PT-03." + NL
            + "Kệ ghép hộp dày 40mm hoàn thiện PT-03 (2 kệ trong ô tivi)." + NL + "Khung xương gỗ" + NL
            + "Hậu khu vực tivi MDF 17mm melamine màu đen" + NL + "Không bao gồm:" + NL + "+ Inox MT-02 (ốp và viền kim loại)")
D = [(62, "F053", mo_ta(1620, 270, 2500), "BoQ F053 1620x270x2500; ID-JS-65-705 (2 kệ, ô TV 1300); block 5613"),
     (63, "F054", mo_ta(1500, 465, 2500), "BoQ F054 1500x465x2500; ID-JS-65-705 tham khảo; block riêng 5696 (chép F053, ô TV 1180)")]
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "MO_TA"
ws.append(["Dòng BoQ", "Mã", "Mô tả costing đề xuất (cột AJ)", "Căn cứ", "cmt"])
for r in D: ws.append(list(r) + [None])
ws.column_dimensions["C"].width = 70; ws.column_dimensions["D"].width = 60
for row in ws.iter_rows(min_row=2): row[2].alignment = Alignment(wrap_text=True, vertical="top")
wb.save(r"D:\COSTING EXCEL\_ACTIVE\SOFITEL\02_takeoff\MO_TA_2b.xlsx"); print("ok", len(D))
