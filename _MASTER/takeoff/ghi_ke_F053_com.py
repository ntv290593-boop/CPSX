"""Đợt 2b: thêm 2 dòng 'Kệ' (2 kệ ghép hộp trong ô trống) vào block F053 (dùng 2 hàng đệm +12 và +23) + trả E block F030 về chỉ F030.
Qua COM (không biểu diễn được trong tblAiModule: thêm dòng). Tô vàng + _EDITLOG. Dùng: py -I ghi_ke_F053_com.py <CPSX.xlsm>"""
import os, sys, shutil, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở trong Excel")
bak = os.path.join(d, "_backup", f"{os.path.splitext(n)[0]}_truoc-ke_{datetime.datetime.now():%H%M%S}.xlsm")
shutil.copy2(cpsx, bak)
H, OT = 5613, 5627   # hdr block F053, dòng 'Ô trống vách ti vi'
BQ = "'MATRIX FITFUR-LOOSE'"
MARK = 0x00FFFF
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh"); lg = wb.Worksheets("_EDITLOG")
    assert str(ws.Cells(OT, 1).Value).startswith("Ô trống"), "sai dòng ô trống"
    assert str(ws.Cells(H + 12, 1).Value or "") == "" and str(ws.Cells(H + 23, 1).Value or "") == "", "hàng đệm không trống"
    assert ws.Cells(863, 5).Formula == f"={BQ}!AG39+{BQ}!AG63", ws.Cells(863, 5).Formula
    last = lg.UsedRange.Row + lg.UsedRange.Rows.Count - 1
    now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M"); log = []
    def put(addr, val, note, formula=False):
        c = ws.Range(addr)
        old = "" if c.Interior.ColorIndex == -4142 else "#%02X%02X%02X" % (int(c.Interior.Color) & 255, (int(c.Interior.Color) >> 8) & 255, (int(c.Interior.Color) >> 16) & 255)
        if formula: c.Formula = val
        else: c.Value = val
        c.Interior.Color = MARK
        log.append(("Chiet tinh", addr, old, now, note, "f:" + str(val)[:200], ""))
    bformula = ws.Cells(H + 22, 2).FormulaR1C1   # công thức B (dày từ mã K) của dòng khe đèn
    for off, ten, cc, e, n_ in ((12, "Kệ ghép hộp trong ô TV: mặt trên + đáy (2 kệ x 2 tấm), chỉ gỗ", 140, 4, "pt-01"),
                                (23, "Kệ ghép hộp trong ô TV: mặt trước (2 kệ), lõi gỗ, bọc MT-02 không tính", 40, 2, "sl")):
        r = H + off
        ws.Cells(r, 2).FormulaR1C1 = bformula
        put(f"A{r}", ten, "2b kệ F053"); put(f"C{r}", cc, "2b kệ F053"); put(f"D{r}", f"=+B{OT}", "2b kệ F053", True)
        put(f"E{r}", e, "2b kệ F053"); put(f"K{r}", "17mdf", "2b kệ F053"); put(f"M{r}", 1, "2b kệ F053"); put(f"N{r}", n_, "2b kệ F053")
        put(f"S{r}", f"=+S{H + 4}", "2b kệ F053", True)
    put("E863", f"={BQ}!AG39", "2b: tách F054 khỏi block F030", True)
    lg.Range(f"A{last + 1}:G{last + len(log)}").Value = tuple(log)
    wb.Saved = False; wb.Save()
    for off in (12, 23):
        r = H + off
        print(r, [ws.Cells(r, c).Text for c in (1, 2, 3, 4, 5, 6, 7, 9, 11)])
    print("E863", ws.Cells(863, 5).Formula, ws.Cells(863, 5).Value, "| F053 hdr E/H/I:", ws.Cells(H, 5).Value, ws.Cells(H, 8).Value)
    wb.Close(False)
finally:
    xl.Quit()
print("sao lưu:", bak)
