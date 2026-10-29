"""Gỡ màu đánh dấu (vàng ô sửa / hồng mô tả) của các món ĐÃ LÀM, để món mới nổi bật. Trả đúng 'màu cũ' ghi trong _EDITLOG (như L.unmark).
Giữ nguyên dòng nhật ký (cột F 'Claude ghi' còn dùng để học). Bỏ qua các mục nhật ký có ghi chú chứa chuỗi truyền vào (đợt mới, vd 'KL').
Dùng: py -I go_mau_da_lam_com.py <CPSX.xlsm> <chuỗi ghi chú đợt mới, giữ màu>"""
import os, sys, shutil, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); keep_ids = set(sys.argv[2:]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở")
shutil.copy2(cpsx, os.path.join(d, "_backup", f"{os.path.splitext(n)[0]}_truoc-gomau_{datetime.datetime.now():%H%M%S}.xlsm"))
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); lg = wb.Worksheets("_EDITLOG")
    last = lg.UsedRange.Row + lg.UsedRange.Rows.Count - 1
    rows = lg.Range(f"A2:G{last}").Value
    done = kept = 0; skipped = []
    for i, r in enumerate(rows, start=2):
        sh, addr, old, note = r[0], r[1], r[2], str(r[4] or "")
        if not sh or not addr: continue
        if any(k in note for k in keep_ids): kept += 1; continue
        try: c = wb.Worksheets(sh).Range(addr)
        except Exception: skipped.append((sh, addr)); continue
        cur = c.Interior.Color if c.Interior.ColorIndex != -4142 else None
        if cur in (0x00FFFF, 0xCCCCFF):   # chỉ gỡ đúng màu Claude tô (vàng / hồng); user tự tô màu khác giữ nguyên
            if old: c.Interior.Color = int(str(old)[5:7] + str(old)[3:5] + str(old)[1:3], 16)   # '#RRGGBB' → BGR
            else: c.Interior.ColorIndex = -4142
            done += 1
    wb.Saved = False; wb.Save(); wb.Close(False)
    print(f"đã gỡ màu {done} ô; giữ màu (đợt mới) {kept} mục; không mở được {len(skipped)}")
finally:
    xl.Quit()
