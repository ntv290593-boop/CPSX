"""Bedside F047, F079, F080: bản vẽ (GR-TY-708, JS-65-706) ghi WD 02 cho mặt tủ bedside (WD 04 là ván đầu giường) → tính WD-02 (Material: WD-02 = VN.2),
thay mã hoàn thiện 'wd-01' → 'wd-02' ở các hàng của block và viết lại mô tả AJ. Ô sửa tô vàng, AJ tô hồng, ghi _EDITLOG.
Dùng: py -I doi_wd02_bedside_com.py <CPSX.xlsm>"""
import os, sys, shutil, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở")
shutil.copy2(cpsx, os.path.join(d, "_backup", f"{os.path.splitext(n)[0]}_truoc-wd02b_{datetime.datetime.now():%H%M%S}.xlsm"))
BQ = "MATRIX FITFUR-LOOSE"; YEL, PINK = 0x00FFFF, 0xCCCCFF
ITEMS = [("F044", 53), ("F045", 54), ("F046", 55), ("F048", 57)]
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh"); bq = wb.Worksheets(BQ); lg = wb.Worksheets("_EDITLOG")
    last = lg.UsedRange.Row + lg.UsedRange.Rows.Count - 1; now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M"); log = []
    import re
    for ma, r in ITEMS:
        assert str(bq.Cells(r, 1).Value).strip() == ma
        h = int(re.search(r"H\$?(\d+)", str(bq.Cells(r, 34).Formula)).group(1))
        end = next(x for x in range(h + 1, h + 200) if ws.Cells(x, 1).Value == "END")
        before = ws.Cells(h, 8).Value; cnt = 0
        for x in range(h, end):
            c = ws.Cells(x, 14)
            if str(c.Value or "").strip().lower() == "wd-01":
                c.Value = "wd-02"; c.Interior.Color = YEL; cnt += 1
                log.append(("Chiet tinh", f"N{x}", "", now, f"REV04 bedside cũ {ma} WD-02", "f:wd-02", ""))
        a = bq.Cells(r, 36); old = str(a.Value)
        new = old.replace("WD-01", "WD-02"); a.NumberFormat = "@"; a.Value = new; a.Interior.Color = PINK
        log.append((BQ, f"AJ{r}", "", now, "desc REV04 bedside WD-02", "f:" + new[:2000], ""))
        print(f"{ma}: block {h}-{end}: đổi {cnt} ô wd-01→wd-02 | đơn giá {before:,.0f} → {ws.Cells(h, 8).Value:,.0f} | AJ: {new!r}", flush=True)
    lg.Range(f"A{last + 1}:G{last + len(log)}").Value = tuple(log)
    wb.Saved = False; wb.Save(); wb.Close(False)
finally:
    xl.Quit()
