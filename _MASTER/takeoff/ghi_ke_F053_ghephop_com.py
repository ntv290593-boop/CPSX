"""Đợt 2b (sửa): kệ F053 = module GHÉP HỘP thật. Gỡ 2 dòng MDF rời đã thêm (hàng đệm +12, +23), rồi copy module
'CABINET CÓ CÁNH CỬA'+'Hộp' (10 hàng, offset +2..+11) chèn trước hàng đệm cuối; B module = rộng ô trống, C=sâu kệ 140, D=dày kệ 40, Hộp E=2.
Qua COM; tô vàng + _EDITLOG. Dùng: py -I ghi_ke_F053_ghephop_com.py <CPSX.xlsm>"""
import os, sys, shutil, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở trong Excel")
bak = os.path.join(d, "_backup", f"{os.path.splitext(n)[0]}_truoc-kehop_{datetime.datetime.now():%H%M%S}.xlsm")
shutil.copy2(cpsx, bak)
H, OT = 5613, 5627
MARK = 0x00FFFF
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh"); lg = wb.Worksheets("_EDITLOG")
    assert str(ws.Cells(OT, 1).Value).startswith("Ô trống")
    assert str(ws.Cells(H + 2, 1).Value).startswith("CABINET") and str(ws.Cells(H + 3, 1).Value) == "Hộp"
    assert str(ws.Cells(H + 12, 1).Value).startswith("Kệ") and str(ws.Cells(H + 23, 1).Value).startswith("Kệ"), "2 dòng kệ rời không còn"
    # 1) gỡ 2 dòng kệ rời: xóa ô nhập + trả màu (nhật ký 'màu cũ' đều là trống) 
    for off in (12, 23):
        r = H + off
        for c in "ABCDEKMNS":
            ws.Range(f"{c}{r}").ClearContents(); ws.Range(f"{c}{r}").Interior.ColorIndex = -4142
    last = lg.UsedRange.Row + lg.UsedRange.Rows.Count - 1
    # bỏ các dòng nhật ký '2b kệ F053' của 2 hàng đã gỡ
    for i in range(last, 1, -1):
        if str(lg.Cells(i, 5).Value) == "2b kệ F053": lg.Rows(i).Delete()
    last = lg.UsedRange.Row + lg.UsedRange.Rows.Count - 1
    # 2) chèn 10 hàng trước hàng đệm cuối (H+23), copy module
    at = H + 23
    ws.Rows(f"{at}:{at + 9}").Insert()
    ws.Rows(f"{H + 2}:{H + 11}").Copy(ws.Rows(at))
    xl.CutCopyMode = False
    mod, hop = at, at + 1
    now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M"); log = []
    def put(addr, val, formula=False):
        c = ws.Range(addr)
        old = "" if c.Interior.ColorIndex == -4142 else "#%02X%02X%02X" % (int(c.Interior.Color) & 255, (int(c.Interior.Color) >> 8) & 255, (int(c.Interior.Color) >> 16) & 255)
        if formula: c.Formula = val
        else: c.Value = val
        c.Interior.Color = MARK
        log.append(("Chiet tinh", addr, old, now, "2b kệ F053 ghép hộp", "f:" + str(val)[:200], ""))
    put(f"A{mod}", "KỆ GHÉP HỘP trong ô TV (2 kệ, chỉ gỗ; MT-02 bọc ngoài không tính)")
    put(f"B{mod}", f"=+B{OT}", True); put(f"C{mod}", 140); put(f"D{mod}", 40)
    put(f"E{hop}", 2)
    lg.Range(f"A{last + 1}:G{last + len(log)}").Value = tuple(log)
    wb.Saved = False; wb.Save()
    print("module kệ hàng", mod, "..", mod + 9)
    for r in range(mod, mod + 10):
        print(r, [ws.Cells(r, c).Text.strip()[:22] for c in (1, 2, 3, 4, 5, 9, 11)])
    errs = [(r, c) for r in range(H, H + 95) for c in range(1, 63) if str(ws.Cells(r, c).Text).startswith("#")]
    print("ô lỗi trong block:", errs[:10], "| hdr E,H:", ws.Cells(H, 5).Value, ws.Cells(H, 8).Value)
    print("END ở hàng:", [r for r in range(H + 80, H + 100) if ws.Cells(r, 1).Value == "END"])
    wb.Close(False)
finally:
    xl.Quit()
print("sao lưu:", bak)
