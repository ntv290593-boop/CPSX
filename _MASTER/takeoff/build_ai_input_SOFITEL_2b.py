"""AI_INPUT đợt 2b SOFITEL — nhóm vách TV (F053, F054). Nguồn quyết định: đọc ID-JS-65-705 (p012 A65) + user 2026-10-25.
Nền: REV03 (RE00 + 2a). Không làm lại mục đã có AH. Ghi ra 02_takeoff\\AI_INPUT_batch2b.xlsx."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from extract_ai_input import AI_COLS, MOD_COLS, write

sys.stdout.reconfigure(encoding="utf-8")
CPSX = r"D:\COSTING EXCEL\_ACTIVE\SOFITEL\03_costing\261024_SOFITEL_CPSX_REV03.xlsm"
OUT = r"D:\COSTING EXCEL\_ACTIVE\SOFITEL\02_takeoff\\AI_INPUT_batch2b.xlsx"
RUN = "BATCH2B-20261025"
BQ = "'MATRIX FITFUR-LOOSE'"
TV = 'Template "Vách ghép hộp ( Vách đầu giường, vách tivi..)"'
SRC = {"F029": 790, "F030": 863}
TR = "ID-JS-65-705 p12 (plan 01, elevation 02, detail 05)"

wb = openpyxl.load_workbook(CPSX, read_only=True)
boq = {str(r[0]).strip(): (i, r[2], r[1]) for i, r in enumerate(
    wb["MATRIX FITFUR-LOOSE"].iter_rows(min_row=8, max_col=7, values_only=True), start=8) if r[0]}
hdr = {i: r for i, r in enumerate(wb["Chiet tinh"].iter_rows(min_row=1, max_row=1000, max_col=5, values_only=True), start=1) if i in SRC.values()}
wb.close()
v = openpyxl.load_workbook(CPSX, read_only=True, data_only=True)
kl = {str(r[0]).strip(): r[32] for r in v["MATRIX FITFUR-LOOSE"].iter_rows(min_row=8, max_col=33, values_only=True) if r[0]}
done = {c for c, r in ((str(r[0]).strip(), r[33]) for r in v["MATRIX FITFUR-LOOSE"].iter_rows(min_row=8, max_col=34, values_only=True) if r[0]) if r}
v.close()
assert "F053" not in done and "F054" not in done, "đã có AH, không làm lại"

# (BlockId, thao tác, mã, nguồn, W, D, H, sửa module, độ tin, ghi chú)
B = [
 ("F053-1", "MOI", ["F053"], "F029", 1620, 150, 2500,
  [("Ô trống vách ti vi", 1, 14, "B", 1300, "NHAP"), ("Ô trống vách ti vi", 1, 14, "D", 765, "NHAP")], "TB",
  "Chép F029 (hdr 790). Header B:D tự đọc từ BoQ (1620x270x2500); bản vẽ mặt bằng sâu 150 ≠ BoQ 270 → RFI, chưa ghi đè. "
  "Ô trống TV = 1300 rộng x 765 cao (giữa 2 nẹp MT-02), hốc sâu theo C header. MT-02 không tính."),
 ("F030-GOP", "SUA", ["F030", "F054"], "F030", 1500, 465, 2500, [], "TB",
  "Gộp F054 vào block F030 có sẵn (BoQ F054 = F030: 1500x465x2500, cùng mô tả). JS-65-705 chỉ vẽ 1620 → không có BV riêng cho 1500x465, RFI."),
]
ai, mod = [], []
for bid, op, mas, src, w, d, h, edits, tin, ghi in B:
    rows = [boq[m][0] for m in mas]
    hh = hdr[SRC[src]]
    if op == "SUA":
        assert f"!A{rows[0]}" in str(hh[0]).replace("$", "") and f"!AG{rows[0]}" in str(hh[4]).replace("$", ""), (bid, hh)
    r0, ten, loai = boq[mas[0]]
    ai.append({"BlockId": bid, "ThaoTac": op, "Ma": " ".join(mas), "DongBoQ": ",".join(map(str, rows)), "QuanHe": "TACH",
               "Ten": ten, "Loai": loai, "Mau": TV, "MauTen": f"Chiet tinh!#{SRC[src]}", "W": w, "D": d, "H": h,
               "SL": sum(kl[m] for m in mas), "GiaTheo": None, "Nguon": "SCAN", "DoTinCay": tin, "Trang": TR,
               "GhiChu": f"{'sửa' if op == 'SUA' else 'chép'} {src} (hdr {SRC[src]}). " + ghi, "RunId": RUN})
    # B:D header của mẫu là công thức đọc kích thước từ mô tả BoQ → không ghi số chết, chỉ link E
    ed = []
    ed.append(("hdr", 1, 0, "E", "=" + "+".join(f"{BQ}!AG{r}" for r in rows), "LINK"))
    for m, lan, off, cot, val, typ in ed + edits:
        mod.append({"BlockId": bid, "Module": m, "Lan": lan, "Off": off, "Cot": cot, "GiaTri": val, "Loai": typ,
                    "Nguon": "SCAN", "Trang": TR, "DoTinCay": tin, "GhiChu": None})
write(OUT, [("AI_INPUT", "tblAiInput", AI_COLS, ai), ("AI_MODULE", "tblAiModule", MOD_COLS, mod)])
print(len(ai), "block,", len(mod), "ô nhập →", OUT)
for a in ai: print(" ", a["BlockId"], a["ThaoTac"], a["DongBoQ"], "SL=", a["SL"])
