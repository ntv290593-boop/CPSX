"""REV04: xóa block thừa của L012 (hdr 6543, không dòng BoQ nào trỏ tới), gộp khối lượng L012 vào block chung L041 (E = AG151 + AG122),
ghi lại mô tả riêng cho L012 (mã vải 52-503A/B, md theo block). Qua COM; ô sửa tô vàng + _EDITLOG; cập nhật _AIRUN (Hang dịch lên).
Dùng: py -I xoa_block_thua_L012_com.py <CPSX.xlsm>"""
import os, sys, shutil, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở")
shutil.copy2(cpsx, os.path.join(d, "_backup", f"{os.path.splitext(n)[0]}_truoc-xoa-L012_{datetime.datetime.now():%H%M%S}.xlsm"))
BQ = "MATRIX FITFUR-LOOSE"; YEL, PINK = 0x00FFFF, 0xCCCCFF
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh"); bq = wb.Worksheets(BQ)
    lg = wb.Worksheets("_EDITLOG"); run = next(s for s in wb.Worksheets if s.Name == "_AIRUN")
    h12 = next(x for x in range(6300, 6800) if str(ws.Cells(x, 1).Formula).endswith("!A122"))
    end12 = next(x for x in range(h12 + 1, h12 + 200) if ws.Cells(x, 1).Value == "END")
    assert not any("H" + str(h12) in str(bq.Cells(r, 34).Formula).replace("$", "") for r in range(8, 200)), "còn dòng BoQ trỏ vào block L012"
    assert str(bq.Cells(122, 34).Formula).replace("$", "") == "='Chiet tinh'!H" + str(next(x for x in range(h12, h12 + 200) if str(ws.Cells(x, 1).Formula).endswith("!A151")))
    n_del = end12 - h12 + 1
    print("xóa block L012:", h12, end12, f"({n_del} hàng)")
    ws.Rows(f"{h12}:{end12}").Delete()
    h41 = next(x for x in range(6300, 6800) if str(ws.Cells(x, 1).Formula).endswith("!A151"))
    assert h41 == h12, (h41, h12)   # block L041 dời lên đúng chỗ
    last = lg.UsedRange.Row + lg.UsedRange.Rows.Count - 1; now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M"); log = []
    e = f"=+'{BQ}'!AG151+'{BQ}'!AG122"
    old = ws.Cells(h41, 5).Interior.ColorIndex
    ws.Cells(h41, 5).Formula = e; ws.Cells(h41, 5).Interior.Color = YEL
    log.append(("Chiet tinh", f"E{h41}", "" if old == -4142 else "#?", now, "REV04 L012 gộp vào block L041", "f:" + e, ""))
    ad = ("Kích thước W570xD670xH890 mm." + "\n" + "Đệm ngồi, đệm lưng foam chịu nước (quick dry)." + "\n" + "Khung ghế inox" + "\n"
          + "Không bao gồm:" + "\n" + "+ Vải 52-503A, 52-503B : 1.6md")
    c = bq.Range("AJ122"); oldc = "" if c.Interior.ColorIndex == -4142 else "#?"
    c.NumberFormat = "@"; c.Value = ad; c.WrapText = True; c.Interior.Color = PINK
    log.append((BQ, "AJ122", oldc, now, "desc REV04 L012 riêng theo mã", "f:" + ad, ""))
    lg.Range(f"A{last + 1}:G{last + len(log)}").Value = tuple(log)
    # _AIRUN: bỏ dòng L012-1, dời Hang các block sau
    rl = run.UsedRange.Row + run.UsedRange.Rows.Count - 1
    for i in range(rl, 1, -1):
        bid, hang = str(run.Cells(i, 1).Value), run.Cells(i, 3).Value
        if bid == "L012-1": run.Rows(i).Delete()
        elif isinstance(hang, (int, float)) and hang > end12: run.Cells(i, 3).Value = hang - n_del
    print("E block chung:", ws.Cells(h41, 5).Value, "| đơn giá:", ws.Cells(h41, 8).Value, "| AH122 →", bq.Cells(122, 34).Formula, "| AH151 →", bq.Cells(151, 34).Formula)
    print("AH152 (L042) →", bq.Cells(152, 34).Formula, "| lỗi?", [(r, c_) for r in range(h41, h41 + 110) for c_ in range(1, 63) if str(ws.Cells(r, c_).Text).startswith("#") and set(str(ws.Cells(r, c_).Text)) != {"#"}][:3])
    wb.Saved = False; wb.Save(); wb.Close(False)
finally:
    xl.Quit()
