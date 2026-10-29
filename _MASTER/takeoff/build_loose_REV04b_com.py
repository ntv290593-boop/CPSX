"""REV04: chép block mẫu (L015 bàn, L016/L010 ghế) cho món loose có gỗ, chỉnh theo spec FF&E, link BoQ (A, E, AH), mô tả AJ.
Làm trực tiếp qua COM (user không muốn thao tác L.fromAiInput). Tô vàng ô sửa + _EDITLOG + _AIRUN, mô tả tô hồng.
Dùng: py -I build_loose_REV04_com.py <CPSX REV04.xlsm>"""
import os, sys, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở")
BQ = "MATRIX FITFUR-LOOSE"; YEL, PINK = 0x00FFFF, 0xCCCCFF
NL = "\n"
def mo_ta(*lines, kbg=()):
    return NL.join(lines) + ((NL + "Không bao gồm:" + "".join(NL + "+ " + k for k in kbg)) if kbg else "")
# (mã, dòng BoQ, hdr nguồn, B, C, D, {offset: {cột: giá trị}}, mô tả)
OAK = "Oak sơn PU hoàn thiện (dark stained oak)"
ITEMS = [
 ("L034", 144, ("Chiet tinh", 2435), 400, 400, 450, {11: {"E": 0}, 12: {"E": 0}, 18: {"E": 0}, 28: {"E": 0}},
  mo_ta("Kích thước D400xH450 mm", f"Ghế đôn khung gỗ {OAK}, 4 chân vuốt", "Mê ngồi foam", kbg=("Da PU 53-501A",))),
 # Template#15 (bọc nệm kiểu sofa/armchair): K hdr = Cong (lưng tub cong), K(+1) = Armchair, mê rời (B +31)
 ("L039", 149, ("Template", 1018), 750, 780, 780, {0: {"K": "Cong"}, 1: {"K": "Armchair"}, 31: {"B": "Mê rời"}},
  mo_ta("Kích thước W750xD780xH780 mm", "Ghế bành bọc nệm: khung ván PLW ghép hộp (mê ngồi, tay, lưng cong)", "Foam mê ngồi, tay, lưng",
        kbg=("Vải 54-504A, 54-504B, 54-504C", "Gối lưng", "Đế xoay gỗ Oak (swivel plinth)"))),
]
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh"); bq = wb.Worksheets(BQ)
    lg = wb.Worksheets("_EDITLOG"); run = next(s for s in wb.Worksheets if s.Name == "_AIRUN")
    def last(sh): return sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1
    log = []; now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M"); RUN = "REV04-LOOSE-" + now[:10]
    def rgb(c): return "" if c.Interior.ColorIndex == -4142 else "#%02X%02X%02X" % (int(c.Interior.Color) & 255, (int(c.Interior.Color) >> 8) & 255, (int(c.Interior.Color) >> 16) & 255)
    def put(sh, addr, val, color, note, formula=False):
        c = sh.Range(addr); old = rgb(c)
        if formula: c.Formula = val
        else: c.Value = val
        c.Interior.Color = color; log.append((sh.Name, addr, old, now, note, "f:" + str(val)[:200], ""))
    for ma, r, (ssh, src), B, C, D, edits, desc in ITEMS:
        sws = wb.Worksheets(ssh)
        assert str(bq.Cells(r, 1).Value).strip() == ma, (ma, r)
        assert not str(bq.Cells(r, 34).Formula), f"{ma} đã có AH — không làm lại"
        assert (ssh == "Template" and sws.Cells(src, 1).Value == "Bench, Armchair, Ghế đôn") or str(sws.Cells(src, 1).Formula).startswith("=+'MATRIX FITFUR-LOOSE'!A"), src
        end = next(x for x in range(src + 1, src + 200) if sws.Cells(x, 1).Value == "END")
        new = next(x for x in range(5000, 20000) if ws.Cells(x, 1).Value == "END" and not ws.Cells(x + 1, 1).Value and not ws.Cells(x + 2, 1).Value) + 1
        sws.Rows(f"{src}:{end}").Copy(ws.Rows(new)); xl.CutCopyMode = False
        ne = new + (end - src)
        note = f"REV04 {ma}"
        put(ws, f"A{new}", f"=+'{BQ}'!A{r}", YEL, note, True)
        put(ws, f"E{new}", f"=+'{BQ}'!AG{r}", YEL, note, True)
        for col, v in (("B", B), ("C", C), ("D", D)): put(ws, f"{col}{new}", v, YEL, note)
        for off, cells in edits.items():
            for col, v in cells.items(): put(ws, f"{col}{new + off}", v, YEL, note)
        # AH BoQ → H block mới (định dạng theo ô AH có sẵn)
        bq.Range("AH126").Copy(); bq.Range(f"AH{r}").PasteSpecial(-4122); xl.CutCopyMode = False
        put(bq, f"AH{r}", f"='Chiet tinh'!H{new}", None or bq.Range(f"AH{r}").Interior.Color, note, True)
        c = bq.Range(f"AJ{r}"); old = rgb(c); rc = bq.Range("AJ126")
        c.Font.Name, c.Font.Size = rc.Font.Name, rc.Font.Size; c.VerticalAlignment, c.HorizontalAlignment = rc.VerticalAlignment, rc.HorizontalAlignment
        c.NumberFormat = "@"; c.Value = desc; c.WrapText = True; c.Interior.Color = PINK
        log.append((BQ, f"AJ{r}", old, now, "desc REV04 loose", "f:" + desc[:2000], ""))
        lr = last(run) + 1; run.Range(f"A{lr}:E{lr}").Value = ((f"{ma}-1", "MOI", new, RUN, now),)
        print(f"{ma}: block {new}-{ne} (nguồn {src}) | H={ws.Cells(new, 8).Text} | E={ws.Cells(new, 5).Text} | hdr B:D={ws.Cells(new,2).Text},{ws.Cells(new,3).Text},{ws.Cells(new,4).Text}")
    lg.Range(f"A{last(lg) + 1}:G{last(lg) + len(log)}").Value = tuple(log)
    wb.Saved = False; wb.Save()
    errs = []
    for r in range(6197, last(ws) + 1):
        for c in range(1, 63):
            t = str(ws.Cells(r, c).Text)
            if t.startswith("#") and set(t) != {"#"}: errs.append((r, c, t))
    print("ô lỗi trong các block mới:", errs[:12])
    wb.Close(False)
finally:
    xl.Quit()
