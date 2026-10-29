"""REV04: bedside liền tường (F047, F079, F080) = chép block bedside cùng loại đã làm (F048 cho F047 cùng tờ GR-TY-708; F046 cho F079/F080).
Hdr B:D = kích thước BoQ (nếu ô nguồn là hằng số), E/A link BoQ, AH → H block mới, mô tả AJ theo mô tả của block nguồn (user đã duyệt).
Không tính kim loại (đồ không loose). Không chỉnh nhân công. Ô sửa tô vàng + _EDITLOG + _AIRUN. Dùng: py -I build_bedside_REV04_com.py <CPSX.xlsm>"""
import os, sys, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở")
BQ = "MATRIX FITFUR-LOOSE"; YEL, PINK = 0x00FFFF, 0xCCCCFF
# (mã, dòng BoQ, hdr nguồn, (B, C, D), dòng BoQ nguồn để lấy mô tả)
ITEMS = [("F047", 56, 1707, (1445, 450, 610), 57), ("F079", 88, 1500, (660, 450, 610), 55), ("F080", 89, 1500, (575, 450, 610), 55)]
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh"); bq = wb.Worksheets(BQ)
    lg = wb.Worksheets("_EDITLOG"); run = next(s for s in wb.Worksheets if s.Name == "_AIRUN")
    def last(sh): return sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1
    log = []; now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M"); RUN = "REV04-BEDSIDE-" + now[:10]
    def rgb(c): return "" if c.Interior.ColorIndex == -4142 else "#%02X%02X%02X" % (int(c.Interior.Color) & 255, (int(c.Interior.Color) >> 8) & 255, (int(c.Interior.Color) >> 16) & 255)
    def put(sh, addr, val, color, note, formula=False):
        c = sh.Range(addr); old = rgb(c)
        if formula: c.Formula = val
        else: c.Value = val
        c.Interior.Color = color; log.append((sh.Name, addr, old, now, note, "f:" + str(val)[:200], ""))
    for ma, r, src, (B, C, D), rdesc in ITEMS:
        assert str(bq.Cells(r, 1).Value).strip() == ma and not str(bq.Cells(r, 34).Formula), ma
        assert str(ws.Cells(src, 1).Formula).replace("$", "").endswith(f"!A{rdesc}"), (src, ws.Cells(src, 1).Formula)
        end = next(x for x in range(src + 1, src + 200) if ws.Cells(x, 1).Value == "END")
        new = next(x for x in range(5000, 20000) if ws.Cells(x, 1).Value == "END" and not ws.Cells(x + 1, 1).Value and not ws.Cells(x + 2, 1).Value) + 1
        ws.Rows(f"{src}:{end}").Copy(ws.Rows(new)); xl.CutCopyMode = False
        note = f"REV04 bedside {ma}"
        put(ws, f"A{new}", f"=+'{BQ}'!A{r}", YEL, note, True); put(ws, f"E{new}", f"=+'{BQ}'!AG{r}", YEL, note, True)
        for col, v in zip("BCD", (B, C, D)):
            if not str(ws.Range(f"{col}{new}").Formula).startswith("="): put(ws, f"{col}{new}", v, YEL, note)
        bq.Range("AH126").Copy(); bq.Range(f"AH{r}").PasteSpecial(-4122); xl.CutCopyMode = False
        put(bq, f"AH{r}", f"='Chiet tinh'!H{new}", bq.Range(f"AH{r}").Interior.Color, note, True)
        desc = str(bq.Cells(rdesc, 36).Value)
        c = bq.Range(f"AJ{r}"); old = rgb(c); rc = bq.Range(f"AJ{rdesc}")
        c.Font.Name, c.Font.Size = rc.Font.Name, rc.Font.Size; c.VerticalAlignment, c.HorizontalAlignment = rc.VerticalAlignment, rc.HorizontalAlignment
        c.NumberFormat = "@"; c.Value = desc; c.WrapText = True; c.Interior.Color = PINK
        log.append((BQ, f"AJ{r}", old, now, "desc REV04 bedside", "f:" + desc[:2000], ""))
        lr = last(run) + 1; run.Range(f"A{lr}:E{lr}").Value = ((f"{ma}-1", "MOI", new, RUN, now),)
        print(f"{ma}: block {new}-{new + end - src} (nguồn {src}) | hdr {ws.Cells(new,2).Text.strip()}x{ws.Cells(new,3).Text.strip()}x{ws.Cells(new,4).Text.strip()} | E={ws.Cells(new,5).Text.strip()} | H={ws.Cells(new,8).Text.strip()} | AJ={desc[:60]!r}", flush=True)
    lg.Range(f"A{last(lg) + 1}:G{last(lg) + len(log)}").Value = tuple(log)
    wb.Saved = False; wb.Save()
    errs = [(r_, c_) for r_ in range(6590, last(ws) + 1) for c_ in range(1, 63) if str(ws.Cells(r_, c_).Text).startswith("#") and set(str(ws.Cells(r_, c_).Text)) != {"#"}]
    print("ô lỗi trong block mới:", errs[:5]); wb.Close(False)
finally:
    xl.Quit()
