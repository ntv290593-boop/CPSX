"""QUAN_SAT đợt 2a SOFITEL (đọc bản vẽ 2026-10-24) -> 02_takeoff\\QUAN_SAT_SOFITEL.xlsx.
Theo thiết kế _MASTER\\modules\\SCAN-NHAN-LOAI.md: 1 dòng = 1 quan sát, có nguồn; chỉ THÊM, không ghi đè.
PROJECT_DATA v8 đang bị khóa nên ghi file riêng; gộp vào PROJECT_DATA khi mở được."""
import sys, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

sys.stdout.reconfigure(encoding="utf-8")
OUT = r"D:\COSTING EXCEL\_ACTIVE\SOFITEL\02_takeoff\QUAN_SAT_SOFITEL.xlsx"
RUN, MODEL, DSL = "2a-20261024", "claude-opus-5-5", "LOAI_DO_VAT_SOFITEL v1 (32 loại)"
HAVE, NO, UNSEEN, CONFLICT = "CÓ", "KHÔNG CÓ", "KHÔNG THẤY", "MÂU THUẪN"

# (Mã, Thành phần, Trường, Giá trị, Trạng thái, Tờ, Chi tiết, Nguồn)
Q = [
    # F006 WD-01 veneer foyer 989.9 m2
    ("F006", 1, "Bộ tờ đã đọc", "GR-TY-702 (thay thế)", UNSEEN, "ID-MK-TY-702", "#01", "BOQ"),
    ("F006", 1, "Ghi chú", "BoQ dẫn ID-MK-TY-702 #01: không có trong bộ hồ sơ (I01/RFI-01). Tạm theo GR-TY-702 #01 timber wall", UNSEEN, "ID-GR-TY-702", "#01", "HÌNH"),
    ("F006", 1, "Cấu tạo vách", "khung xương ~20 + ván ~18 + veneer; tổng ~40 (đo nét, BV không ghi số)", HAVE, "ID-GR-TY-702", "#01", "HÌNH"),
    ("F006", 1, "Vật liệu gỗ", "WD-01 = veneer Walnut K6121AN (Finishes p20)", HAVE, "FINISHES", "p20", "SPEC"),
    # F007 / F008 len chân 80
    ("F007", 1, "Kích thước BV", "len gỗ đặc cao 75 + khe 5 = 80; dày ≈ tấm vách", HAVE, "ID-GR-TY-702", "#01, #02", "HÌNH"),
    ("F007", 1, "Vật liệu gỗ", "WD-01 (BV không ghi mã, ghi 'refer finishes schedule')", HAVE, "ID-GR-TY-702", "#01", "BOQ"),
    ("F007", 1, "Trùng block", "giống block F012 (WD-01, len 20x80, hàng 441) — block F012 làm 80 dù BoQ F012 ghi 100", CONFLICT, "Chiet tinh", "441", "BOQ"),
    ("F008", 1, "Kích thước BV", "len cao 75 + khe 5 = 80", HAVE, "ID-GR-TY-702", "#01, #02", "HÌNH"),
    ("F008", 1, "Vật liệu gỗ", "WD-03 = sàn gỗ kỹ thuật Salvo European Oak (Finishes p22), không phải veneer; RFI-05 chưa trả lời", CONFLICT, "FINISHES", "p22", "SPEC"),
    # F022 WD-01 bedroom entrance PS 12.97 m2
    ("F022", 1, "Cấu tạo vách", "khung xương 20 + ván ~18 + veneer; tổng 40 (đo nét vector 1:5)", HAVE, "ID-PS-TY-702", "#01", "HÌNH"),
    ("F022", 1, "Vị trí", "lối vào phòng ngủ PS: mặt đứng PS không nằm trong bộ tờ đã đọc", UNSEEN, "ID-PS-TY-702", "-", "BOQ"),
    # F023 nan 4 thanh
    ("F023", 1, "Kích thước BV", "4 thanh 40x140, khe 105, tổng rộng 470 (khớp BoQ 4nr 140x40)", HAVE, "ID-PS-TY-702", "#08", "HÌNH"),
    ("F023", 1, "Chiều cao", "2100 theo BoQ; không thấy trên tờ", UNSEEN, "ID-PS-TY-702", "#06", "BOQ"),
    ("F023", 1, "Hoàn thiện", "TIMBER FINISH (#08); #06 FINS DETAIL lại ghi STONE FINISH", CONFLICT, "ID-PS-TY-702", "#06, #08", "HÌNH"),
    # F024 nan che điều hòa
    ("F024", 1, "Chi tiết nan gỗ", "không thấy: #02 chỉ có lưới gió + 'decorative panel (MEP to advise)' cao 290", UNSEEN, "ID-PS-TY-701", "#02", "HÌNH"),
    ("F024", 1, "Trùng block", "mô tả BoQ trùng F009 (Timber slatting to air grille and at top of wardrobe) → block 192", HAVE, "Chiet tinh", "192", "BOQ"),
    # F028 khung trần gỗ
    ("F028", 1, "Loại thật", "KHUNG TRẦN GỖ (Detail - timber ceiling frame), không phải ốp vách", CONFLICT, "ID-LL-TY-701", "#09", "HÌNH"),
    ("F028", 1, "Kích thước BV", "thanh gỗ đặc ~50 rộng x ~70 cao, rãnh 10x10 hai bên; tấm giữa hoàn thiện gỗ (TIMBER FINISH)", HAVE, "ID-LL-TY-701", "#09", "HÌNH"),
    # F088 PVD MT-02 ốp cửa thang 129 m2
    ("F088", 1, "Ván nền ẩn", "MT-02 ốp trên ván + khung (ô gạch chéo = khung)", HAVE, "ID-LL-TY-702", "#03, #08", "HÌNH"),
    ("F088", 1, "Phần ngoài phạm vi", "kim loại MT-02, đá ST-07", HAVE, "ID-LL-TY-702", "#01, #02", "HÌNH"),
    ("F088", 1, "Kích thước BV", "ốp cửa thang cao 2150, hồi xiên 330/440, viền 155/95", HAVE, "ID-LL-TY-702", "#01, #03, #08", "HÌNH"),
    # F089 gương MR-01 sảnh thang 54.9 m2
    ("F089", 1, "Ván nền ẩn", "MR-01 dán trên ván, khung ô 195 x 595, phân cách 40", HAVE, "ID-LL-TY-704", "#06, #10", "HÌNH"),
    ("F089", 1, "Phần ngoài phạm vi", "gương MR-01, kim loại MT-02, sơn chân PT-03", HAVE, "ID-LL-TY-704", "#06", "HÌNH"),
    # F090 giấy WC-08 + MT-02 sảnh thang 65.6 m2
    ("F090", 1, "Ván nền ẩn", "tấm ván dày 20 bọc giấy, trên khung 20; tổng 40", HAVE, "ID-LL-TY-701", "#02, #03", "HÌNH"),
    ("F090", 1, "Bộ tờ", "tấm WC-08 + nẹp MT-02 vẽ ở LL-TY-704 (#05, #07, #13, module 915) — bộ tờ quét chỉ có LL-TY-701", CONFLICT, "ID-LL-TY-704", "#05, #13", "HÌNH"),
    ("F090", 1, "Phần ngoài phạm vi", "giấy WC-08, kim loại MT-02, len đá", HAVE, "ID-LL-TY-701", "#02", "HÌNH"),
    # F091 gương MR-03 cabin thang 5.4 m2
    ("F091", 1, "Ván nền ẩn", "MR-03 trên ván, viền WD-01, nẹp MT-02 5mm", HAVE, "ID-LL-TY-703", "#08", "HÌNH"),
    ("F091", 1, "Gỗ ngoài dòng BoQ", "cabin thang có tấm WD-01 hai bên + len gỗ WD-01 cao 100 (#12): không thấy dòng BoQ nào", CONFLICT, "ID-LL-TY-703", "#04, #12", "HÌNH"),
    # F094 giấy WC-01 JS 45.3 m2
    ("F094", 1, "Ván nền ẩn", "vách giấy: khung + ván, len gỗ 75+5 (giống PS #02)", HAVE, "ID-JS-65-702", "#02", "HÌNH"),
    ("F094", 1, "Bộ tờ", "đã đọc JS-65-702 (BoQ dẫn) + GR-TY-702 #02", HAVE, "ID-JS-65-702", "#02", "HÌNH"),
    # F095 giấy WC-04 + MT-02 đầu giường PS 106 m2
    ("F095", 1, "Kích thước BV", "tấm WC-04 chia EQ, nẹp MT-02 5mm; cao 1290 (PS-01) / 1300 (PS-02); rộng 3825 / 4815", HAVE, "ID-PS-TY-706, -707", "#02, #04", "HÌNH"),
    ("F095", 1, "Cong/bo", "2 góc bo R225 (tấm giấy chạy cong theo góc)", HAVE, "ID-PS-TY-706, -708", "#01, #05", "HÌNH"),
    ("F095", 1, "Ván nền ẩn", "WC-04 trên ván + khung", HAVE, "ID-PS-TY-708", "#06", "HÌNH"),
    # F096 / F097 giấy WC-04 / WC-06 PS
    ("F096", 1, "Ván nền ẩn", "khung 20 + ván ~18 + giấy; tổng 40 (đo nét vector)", HAVE, "ID-PS-TY-702", "#02", "HÌNH"),
    ("F097", 1, "Ván nền ẩn", "khung 20 + ván ~18 + giấy; tổng 40 (đo nét vector)", HAVE, "ID-PS-TY-702", "#02", "HÌNH"),
    ("F097", 1, "Vị trí WC-06", "mặt đứng PS chỉ ra chỗ dùng WC-06 không nằm trong bộ tờ", UNSEEN, "ID-PS-TY-702", "-", "BOQ"),
    # F098 giấy WC-07 đầu giường JS 9.57 m2
    ("F098", 1, "Kích thước BV", "tấm WC-07 chia 640/635, khe 30 PT-03; cao 1260 (BoQ ghi 1300)", CONFLICT, "ID-JS-65-707", "#04, #07", "HÌNH"),
    ("F098", 1, "MT-02", "không thấy nẹp MT-02 (BoQ có ghi)", UNSEEN, "ID-JS-65-706, -707", "#02, #07", "HÌNH"),
    ("F098", 1, "Ván nền ẩn", "WC-07 trên ván + khung", HAVE, "ID-JS-65-707", "#07", "HÌNH"),
    # F099 giấy WC-03 hành lang 316 m2
    ("F099", 1, "Ván nền ẩn", "tấm ván bọc giấy dày 20, khung 40; nẹp gỗ/sơn 20 giữa tấm", HAVE, "ID-CO-TY-701", "#01, #02", "HÌNH"),
]

# Tóm tắt từng mã: loại đề xuất, độ chắc, cách làm, nguồn chép, quyết định cần anh
T = [
    ("F006", "Ốp veneer gỗ lên tường (m2)", "TB", "MOI", "F017 (735), đổi veneer VN.4 → VN.1 (WD-01)", "Tờ MK-TY-702 không có (RFI-01) — tạm tính như vách veneer"),
    ("F007", "Len chân gỗ/sơn", "CAO", "LINK", "link AH → block F012 (441), giống hệt WD-01 len 20x80", "Link chung block hay tách block riêng?"),
    ("F008", "Len chân gỗ/sơn", "TB", "MOI", "F012 (441), đổi gỗ 25ash → ? (WD-03 sồi kỹ thuật)", "Mã gỗ cho WD-03 (RFI-05)?"),
    ("F022", "Ốp veneer gỗ lên tường (m2)", "CAO", "MOI", "F017 (735), đổi veneer VN.4 → VN.1 (WD-01)", ""),
    ("F023", "Chỉ gỗ / trim / thanh nan rời", "CAO", "MOI", "F011 (303): đổi thành 4 thanh 40x140x2100 / Nr", ""),
    ("F024", "Vách nan che điều hòa", "TB", "LINK", "link AH → block F009 (192), mô tả BoQ trùng F009", "Link chung block F009?"),
    ("F028", "KHÔNG RÕ (khung trần gỗ)", "THẤP", "HOÃN", "", "Loại mới 'khung trần gỗ': mẫu nào?"),
    ("F088", "Ván nền ốp gương / kim loại", "CAO", "MOI", "F086 (1983): tấm nền MDF 17 + khung, không kim loại", ""),
    ("F089", "Ván nền ốp gương / kim loại", "CAO", "MOI", "F086 (1983)", ""),
    ("F090", "Ván nền dán giấy", "CAO", "MOI", "F086 (1983)", ""),
    ("F091", "Ván nền ốp gương / kim loại", "CAO", "MOI", "F086 (1983)", "Gỗ WD-01 cabin thang không có dòng BoQ: RFI?"),
    ("F094", "Ván nền dán giấy", "CAO", "MOI", "F086 (1983)", ""),
    ("F095", "Ván nền dán giấy", "TB", "MOI", "F086 (1983)", "Góc bo R225: thêm hệ số cong hay bỏ qua?"),
    ("F096", "Ván nền dán giấy", "CAO", "MOI", "F086 (1983)", ""),
    ("F097", "Ván nền dán giấy", "CAO", "MOI", "F086 (1983)", ""),
    ("F098", "Ván nền dán giấy", "CAO", "MOI", "F086 (1983)", ""),
    ("F099", "Ván nền dán giấy", "CAO", "MOI", "F086 (1983)", ""),
]

if os.path.exists(OUT):
    sys.exit("QUAN_SAT đã có: chỉ được THÊM dòng — viết bước append riêng, không ghi đè")

wb = openpyxl.Workbook()
H, HF = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="305496")
WR = Alignment(wrap_text=True, vertical="top")
FILL = {CONFLICT: "F8CBAD", UNSEEN: "FFE699"}


def sheet(ws, head, widths, rows, stcol=None):
    ws.append(head)
    for j, w in enumerate(widths, 1):
        c = ws.cell(1, j); c.font, c.fill, c.alignment = H, HF, WR
        ws.column_dimensions[c.column_letter].width = w
    for r in rows:
        ws.append(list(r))
        i = ws.max_row
        for j in range(1, len(head) + 1):
            c = ws.cell(i, j); c.alignment = WR
            if j <= len(r) and isinstance(r[j - 1], str) and r[j - 1].startswith("="):
                c.data_type = "s"
        if stcol and r[stcol] in FILL:
            ws.cell(i, stcol + 1).fill = PatternFill("solid", fgColor=FILL[r[stcol]])
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions


ws = wb.active; ws.title = "QUAN_SAT"
sheet(ws, ["Mã", "Thành phần", "Trường", "Giá trị", "Trạng thái", "Tờ", "Chi tiết", "Nguồn", "RunId", "Model", "Bản DS loại"],
      [8, 6, 18, 70, 12, 20, 12, 8, 14, 16, 26], [q + (RUN, MODEL, DSL) for q in Q], stcol=4)
sheet(wb.create_sheet("TOM_TAT"), ["Mã", "Loại đề xuất", "Độ chắc", "Cách làm", "Nguồn chép / thao tác", "Cần anh quyết", "cmt"],
      [8, 30, 8, 8, 55, 45, 30], T)

codes = {q[0] for q in Q}
assert codes == {t[0] for t in T}, codes ^ {t[0] for t in T}
assert all(q[4] in (HAVE, NO, UNSEEN, CONFLICT) for q in Q)
wb.save(OUT)
n = lambda s: sum(q[4] == s for q in Q)
print(f"{len(Q)} quan sát, {len(codes)} mã; MÂU THUẪN {n(CONFLICT)}, KHÔNG THẤY {n(UNSEEN)}")
