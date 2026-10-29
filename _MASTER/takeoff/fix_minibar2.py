import os, sys, win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = r"D:\COSTING EXCEL\_ACTIVE\SOFITEL\03_costing\261025_SOFITEL_CPSX_REV04.xlsm"
if os.path.exists(r"D:\COSTING EXCEL\_ACTIVE\SOFITEL\03_costing\~$261025_SOFITEL_CPSX_REV04.xlsm"): sys.exit("file đang mở")
xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = xl.DisplayAlerts = xl.EnableEvents = False; xl.AutomationSecurity = 1
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh")
    h = 7392
    assert ws.Cells(h, 1).Formula.endswith("!A51")
    print("trước:", ws.Cells(h, 8).Text, "| hông2 C", ws.Cells(7447, 3).Formula, "| cánh C", ws.Cells(7401, 3).Formula)
    assert ws.Cells(7447, 1).Value == "Hông tủ" and ws.Cells(7401, 1).Value == "Cánh tủ"
    for addr, v in (("C7447", 568), ("C7401", 650)):
        ws.Range(addr).Value = v; ws.Range(addr).Interior.Color = 0x00FFFF
    print("sau:", ws.Cells(h, 8).Text, "| hông2", ws.Cells(7447, 3).Text, ws.Cells(7447, 9).Text, "| cánh", ws.Cells(7401, 3).Text, ws.Cells(7401, 9).Text)
    lg = wb.Worksheets("_EDITLOG"); last = lg.UsedRange.Row + lg.UsedRange.Rows.Count - 1
    lg.Range(f"A{last+1}:G{last+2}").Value = (("Chiet tinh", "C7447", "", "2026-10-26T15:00", "REV04 minibar2", "f:568", ""), ("Chiet tinh", "C7401", "", "2026-10-26T15:00", "REV04 minibar2", "f:650", ""))
    wb.Saved = False; wb.Save(); wb.Close(False)
finally: xl.Quit()
