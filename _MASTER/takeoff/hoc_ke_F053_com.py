"""Đánh dấu 'đã học': user sửa kệ F053 → _EDITLOG cột F (Claude ghi) = công thức/giá trị hiện tại của user. py -I hoc_ke_F053_com.py <CPSX>"""
import os, sys, shutil, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở")
shutil.copy2(cpsx, os.path.join(d, "_backup", f"{os.path.splitext(n)[0]}_truoc-hoc_{datetime.datetime.now():%H%M%S}.xlsm"))
xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = xl.DisplayAlerts = xl.EnableEvents = False; xl.AutomationSecurity = 1
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh"); lg = wb.Worksheets("_EDITLOG")
    last = lg.UsedRange.Row + lg.UsedRange.Rows.Count - 1; k = 0
    for i in range(2, last + 1):
        if str(lg.Cells(i, 5).Value) == "2b kệ F053 ghép hộp":
            cur = "f:" + str(ws.Range(str(lg.Cells(i, 2).Value)).Formula)[:200]
            if str(lg.Cells(i, 6).Value) != cur:
                print(" học", lg.Cells(i, 2).Value, ":", lg.Cells(i, 6).Value, "→", cur); lg.Cells(i, 6).Value = cur; k += 1
    wb.Saved = False; wb.Save(); wb.Close(False); print("đã học", k, "ô")
finally: xl.Quit()
