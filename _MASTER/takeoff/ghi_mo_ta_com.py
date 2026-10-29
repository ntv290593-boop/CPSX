"""Ghi mô tả costing (MO_TA_<đợt>.xlsx) vào cột ghi chú costing của file CPSX qua COM — làm giống L.desc + mark:
tô #FFCCCC, ghi _EDITLOG (Sheet, Ô, Màu cũ, Thời gian, Ghi chú, Claude ghi 'f:…', Học) để L.unmark trả đúng màu cũ.
Excel riêng (DispatchEx), macro/sự kiện tắt; file CPSX phải đang đóng. Sao lưu trước khi ghi, đọc lại từ đĩa sau khi lưu.
Dùng: py -I ghi_mo_ta_com.py <CPSX.xlsm> <MO_TA.xlsx> [sheet BoQ] [cột mô tả]"""
import os, sys, shutil, datetime
import openpyxl
import win32com.client

sys.stdout.reconfigure(encoding="utf-8")
cpsx, src = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
BQ = sys.argv[3] if len(sys.argv) > 3 else "MATRIX FITFUR-LOOSE"
COL = sys.argv[4] if len(sys.argv) > 4 else "AJ"
MARK, LOG, TAG = 0xCCCCFF, "_EDITLOG", "desc " + os.path.splitext(os.path.basename(src))[0]  # #FFCCCC (BGR)

d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)):
    sys.exit(f"{n} đang mở trong Excel — đóng file rồi chạy lại")
rows = [(int(r[0]), str(r[1]), str(r[2])) for r in openpyxl.load_workbook(src, read_only=True)["MO_TA"].iter_rows(min_row=2, values_only=True)
        if r[0] and r[2]]
bak = os.path.join(d, "_backup", f"{os.path.splitext(n)[0]}_truoc-mota_{datetime.datetime.now():%H%M%S}.xlsm")
os.makedirs(os.path.dirname(bak), exist_ok=True)
shutil.copy2(cpsx, bak)

hexc = lambda c: "#%02X%02X%02X" % (c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF)
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx)
    assert not wb.ReadOnly, "file mở chỉ đọc"
    ws = wb.Worksheets(BQ)
    have = {}
    for r, ma, txt in rows:
        code = str(ws.Range(f"A{r}").Value or "").strip()
        assert code == ma, f"dòng {r}: mã {code!r} ≠ {ma!r} — sai dòng BoQ, dừng"
    targets = {r for r, _, _ in rows}
    ref = next(r for r in range(8, 200) if r not in targets and ws.Range(f"{COL}{r}").Value)  # ô mô tả có sẵn để lấy định dạng
    rc = ws.Range(f"{COL}{ref}")
    log_ws = next((s for s in wb.Worksheets if s.Name == LOG), None)
    if log_ws is None:
        log_ws = wb.Worksheets.Add(); log_ws.Name = LOG; log_ws.Range("A1:G1").Value = (("Sheet", "Ô", "Màu cũ", "Thời gian", "Ghi chú", "Claude ghi", "Học"),)
        log_ws.Visible = 2
    last = log_ws.UsedRange.Row + log_ws.UsedRange.Rows.Count - 1
    idx = {}
    if last >= 2:
        for i, v in enumerate(log_ws.Range(f"A2:B{last}").Value, start=2):
            idx[(v[0], v[1])] = i
    now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M")
    new, kept = [], set()
    for r, ma, txt in rows:
        c = ws.Range(f"{COL}{r}")
        if str(c.Value or "").replace("\r", "") == txt:  # mô tả có sẵn giữ nguyên (vd F086): không ghi, không tô
            i = idx.get((BQ, f"{COL}{r}"))
            if i and str(log_ws.Range(f"F{i}").Value or "") != "f:" + txt[:2000]:  # user đã sửa, bản nháp đã theo user → đã học (như L.learned)
                log_ws.Range(f"F{i}:G{i}").Value = (("f:" + txt[:2000], ""),); print(f"  đã học {COL}{r} ({ma})")
            else:
                print(f"  giữ nguyên {COL}{r} ({ma})")
            kept.add(r); continue
        old = "" if c.Interior.ColorIndex == -4142 else hexc(int(c.Interior.Color))
        c.Font.Name, c.Font.Size = rc.Font.Name, rc.Font.Size
        c.VerticalAlignment, c.HorizontalAlignment = rc.VerticalAlignment, rc.HorizontalAlignment
        c.NumberFormat = "@"
        c.Value = txt
        c.WrapText = True
        c.Interior.Color = MARK
        key = (BQ, f"{COL}{r}")
        if key in idx:  # đã có trong nhật ký: giữ màu cũ gốc, cập nhật giá trị chuẩn
            i = idx[key]; log_ws.Range(f"D{i}:G{i}").Value = ((now, TAG, "f:" + txt[:2000], ""),)
        else:
            new.append((BQ, f"{COL}{r}", old, now, TAG, "f:" + txt[:2000], ""))
    if new:
        log_ws.Range(f"A{last + 1}:G{last + len(new)}").Value = tuple(new)
    wb.Saved = False
    wb.Save()
    wb.Close(False)
finally:
    xl.Quit()

# đọc lại từ đĩa
wv = openpyxl.load_workbook(cpsx, read_only=True)
s, lg = wv[BQ], wv[LOG]
logged = {(r[0], r[1]) for r in lg.iter_rows(min_row=2, max_col=2, values_only=True)}
bad = [r for r, _, txt in rows if str(s[f"{COL}{r}"].value or "") != txt or ((BQ, f"{COL}{r}") not in logged and r not in kept)]
print(f"{len(rows) - len(bad)}/{len(rows)} ô {COL} đúng + có nhật ký; sao lưu: {bak}")
if bad:
    sys.exit(f"LỖI các dòng: {bad}")
