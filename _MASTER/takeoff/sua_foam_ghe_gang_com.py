"""Sửa lỗi build_loose_REV04c: 2 hàng foam của L012/L041 bị dán công thức hàng kim loại (tính kg × đơn giá foam → 363 triệu).
Dán lại công thức đúng của hàng foam thường (lấy hàng 'Đệm ngồi' của block L037, cột F..BJ). Dùng: py -I sua_foam_ghe_gang_com.py <CPSX.xlsm>"""
import os, sys, shutil, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở")
shutil.copy2(cpsx, os.path.join(d, "_backup", f"{os.path.splitext(n)[0]}_truoc-suafoam_{datetime.datetime.now():%H%M%S}.xlsm"))
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh")
    ref = next(x for x in range(5900, 6100) if ws.Cells(x, 1).Text.strip() == "Đệm ngồi" and ws.Cells(x, 11).Text.strip() == "fd")  # foam thường của L037
    print("hàng foam chuẩn:", ref)
    for h in (6542, 6594):
        assert str(ws.Cells(h, 1).Formula).find("MATRIX") > 0
        for off in (9, 10):
            r = h + off
            assert str(ws.Cells(r, 1).Value).startswith(("Đệm ngồi foam", "Đệm lưng foam")), (r, ws.Cells(r, 1).Value)
            ws.Range(f"F{ref}:BJ{ref}").Copy(); ws.Range(f"F{r}").PasteSpecial(-4123); xl.CutCopyMode = False
        print(h, "H=", ws.Cells(h, 8).Value, [(ws.Cells(h + o, 1).Text.strip(), ws.Cells(h + o, 7).Text.strip(), ws.Cells(h + o, 9).Text.strip()) for o in (9, 10)])
    wb.Saved = False; wb.Save(); wb.Close(False)
finally:
    xl.Quit()
