"""REV04 (đợt loose kim loại): 6 món loose chỉ có kim loại/đá/kính (L005, L012, L028, L032, L041, L042). User 26/10: kim loại TÍNH cho đồ loose,
KHÔNG chỉnh nhân công, mục tiêu bóc đủ khối lượng/công tác. Đá, kính/gương, vải/da KHÔNG tính nhưng liệt kê ở 'Không bao gồm'.
Base = block bàn L026 của user (hdr 5779: Mặt + FOAM + 3 hàng 'Chân sắt' gia công ngoài). Tiết diện kim loại ƯỚC theo kích thước tổng (spec không có hình chi tiết)."""
import os, sys, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở")
BQ = "MATRIX FITFUR-LOOSE"; YEL, PINK = 0x00FFFF, 0xCCCCFF; NL = "\n"
def mo_ta(*lines, kbg=()):
    return NL.join(lines) + ((NL + "Không bao gồm:" + "".join(NL + "+ " + k for k in kbg)) if kbg else "")
# hàng kim loại: (tên, K, N, B, C, D, M, E, F)
GANG = "MTGC"; BRZ = "ssgc"
GHE_METAL = [("Hông ghế gang đúc (2)", GANG, "Tấm", 8, 670, 500, 8, 2, 1), ("Lưng ghế gang đúc", GANG, "Tấm", 8, 500, 460, 8, 1, 1),
             ("Khung mặt ngồi gang đúc", GANG, "Tấm", 8, 500, 520, 8, 1, 1), ("Chân đế gang đúc", GANG, "Tấm", 10, 40, 2480, 10, 1, 1)]
GHE_FOAM = [("Đệm ngồi foam", 100, 500, 520), ("Đệm lưng foam", 80, 500, 450)]
ITEMS = [
 dict(ma="L005", r=115, hdr=(600, 600, 650), mat=None, foam=[],
      metal=[("Viền mặt bàn gang đúc", GANG, "Tấm", 8, 30, 1885, 8, 1, 1), ("Thân bàn gang đúc (trụ rỗng)", GANG, "Ống", 100, 100, 550, 10, 1, 1),
             ("Đế bàn gang đúc", GANG, "Tấm", 10, 400, 400, 15, 1, 1)],
      ad=mo_ta("Kích thước D600xH650 mm", "Viền mặt, thân bàn và chân đế gang đúc sơn đen.", kbg=("Mặt đá thiêu kết Neolith Estatuario Silk 12mm",))),
 dict(ma="L028", r=138, hdr=(300, 300, 550), mat=None, foam=[],
      metal=[("Viền mặt bàn gang đúc", GANG, "Tấm", 8, 30, 942, 8, 1, 1), ("Thân bàn gang đúc (trụ rỗng)", GANG, "Ống", 60, 60, 450, 8, 1, 1),
             ("Đế tròn bậc gang đúc", GANG, "Tấm", 12, 280, 280, 12, 1, 1)],
      ad=mo_ta("Kích thước D300xH550 mm", "Viền mặt, thân bàn và đế tròn bậc gang đúc sơn đen.", kbg=("Mặt đá thiêu kết Neolith Iron Frost Satin 12mm",))),
 dict(ma="L032", r=142, hdr=(320, 320, 500), mat=None, foam=[],
      metal=[("Tấm đứng cong (2 tấm)", BRZ, "Tấm", 8, 290, 420, 8, 2, 1.5), ("Khung viền mặt đá", BRZ, "Tấm", 5, 30, 1005, 5, 1, 1),
             ("Khung viền kệ dưới", BRZ, "Tấm", 5, 30, 900, 5, 1, 1), ("Tấm đế phẳng", BRZ, "Tấm", 5, 300, 300, 5, 1, 1)],
      ad=mo_ta("Kích thước D320xH500 mm", "Khung kim loại dark bronze: 2 tấm đứng cong, khung viền mặt, khung viền kệ dưới và tấm đế.",
               kbg=("Mặt đá Statuario Marble honed dày 20mm", "Kệ dưới đá Statuario Marble"))),
 dict(ma="L012", r=122, hdr=(570, 670, 890), mat=None, foam=GHE_FOAM, metal=GHE_METAL,
      ad=mo_ta("Kích thước W570xD670xH890 mm", "Thân ghế và chân đế gang đúc sơn đen.", "Đệm ngồi, đệm lưng foam chịu nước (quick dry).",
               kbg=("Vải 52-503A, 52-503B : 1.1md",))),
 dict(ma="L041", r=151, hdr=(570, 670, 890), mat=None, foam=GHE_FOAM, metal=GHE_METAL,
      ad=mo_ta("Kích thước W570xD670xH890 mm", "Thân ghế và chân đế gang đúc sơn đen.", "Đệm ngồi, đệm lưng foam chịu nước (quick dry).",
               kbg=("Vải 54-506A : 1.1md",))),
 dict(ma="L042", r=152, hdr=(850, 50, 1150), mat=("12mdf", 12, 850, 1150), foam=[],
      metal=[("Viền khung kim loại light bronze", BRZ, "Tấm", 2, 50, 4000, 2, 1, 1), ("Bát treo kim loại light bronze (4)", BRZ, "Tấm", 3, 40, 60, 3, 4, 1)],
      ad=mo_ta("Kích thước W850xD50xH1150 mm", "Tấm nền MDF 12mm sơn lót làm khung đỡ da.", "Viền khung và bát treo kim loại light bronze.",
               kbg=("Gương clear low iron vát cạnh", "Da 54-800A"))),
]
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh"); bq = wb.Worksheets(BQ)
    lg = wb.Worksheets("_EDITLOG"); run = next(s for s in wb.Worksheets if s.Name == "_AIRUN")
    SRC = 5779
    assert ws.Cells(SRC, 1).Formula.endswith("!A136") and str(ws.Cells(SRC + 3, 1).Value) == "Mặt"
    def last(sh): return sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1
    log = []; now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M"); RUN = "REV04-LOOSE-KL-" + now[:10]
    def rgb(c): return "" if c.Interior.ColorIndex == -4142 else "#%02X%02X%02X" % (int(c.Interior.Color) & 255, (int(c.Interior.Color) >> 8) & 255, (int(c.Interior.Color) >> 16) & 255)
    def put(sh, addr, val, color, note, formula=False):
        c = sh.Range(addr); old = rgb(c)
        if formula: c.Formula = val
        else: c.Value = val
        c.Interior.Color = color; log.append((sh.Name, addr, old, now, note, "f:" + str(val)[:200], ""))
    src_end = next(x for x in range(SRC + 1, SRC + 200) if ws.Cells(x, 1).Value == "END")
    assert str(ws.Cells(SRC + 29, 1).Value).startswith("Kim loại gia công"), "lệch cấu trúc"
    for it in ITEMS:
        ma, r = it["ma"], it["r"]
        assert str(bq.Cells(r, 1).Value).strip() == ma, (ma, r)
        assert not str(bq.Cells(r, 34).Formula), f"{ma} đã có AH"
        new = next(x for x in range(5000, 20000) if ws.Cells(x, 1).Value == "END" and not ws.Cells(x + 1, 1).Value and not ws.Cells(x + 2, 1).Value) + 1
        ws.Rows(f"{SRC}:{src_end}").Copy(ws.Rows(new)); xl.CutCopyMode = False
        note = f"REV04 {ma}"
        put(ws, f"A{new}", f"=+'{BQ}'!A{r}", YEL, note, True); put(ws, f"E{new}", f"=+'{BQ}'!AG{r}", YEL, note, True)
        for col, v in zip("BCD", it["hdr"]): put(ws, f"{col}{new}", v, YEL, note)
        m = new + 3
        if it["mat"]:
            k, b, c_, d_ = it["mat"]
            put(ws, f"K{m}", k, YEL, note); put(ws, f"B{m}", b, YEL, note); put(ws, f"C{m}", c_, YEL, note); put(ws, f"D{m}", d_, YEL, note)
            put(ws, f"M{m}", 1, YEL, note); put(ws, f"N{m}", "sl", YEL, note); put(ws, f"A{m}", "Tấm nền khung đỡ da", YEL, note)
        else:
            put(ws, f"E{m}", 0, YEL, note)
        for off in (4, 5): put(ws, f"E{new + off}", 0, YEL, note)
        for i, (nm, t, c_, d_) in enumerate(it["foam"]):
            row = new + 9 + i
            ws.Range(f"A{new+30}:BJ{new+30}").Copy(); ws.Range(f"A{row}").PasteSpecial(-4123); xl.CutCopyMode = False
            for col, v in (("A", nm), ("B", t), ("C", c_), ("D", d_), ("E", 1), ("K", "fd")): put(ws, f"{col}{row}", v, YEL, note)
        slots = [new + 30, new + 31, new + 32, new + 34, new + 35]
        for (nm, k, nn, b, c_, d_, mm, e, f), row in zip(it["metal"], slots):
            if row > new + 32:
                ws.Range(f"A{new+30}:BJ{new+30}").Copy(); ws.Range(f"A{row}").PasteSpecial(-4123); xl.CutCopyMode = False
            for col, v in (("A", nm), ("B", b), ("C", c_), ("D", d_), ("E", e), ("F", f), ("K", k), ("M", mm), ("N", nn)): put(ws, f"{col}{row}", v, YEL, note)
        for row in slots[len(it["metal"]):]:
            if row <= new + 32: put(ws, f"E{row}", 0, YEL, note)
        bq.Range("AH126").Copy(); bq.Range(f"AH{r}").PasteSpecial(-4122); xl.CutCopyMode = False
        put(bq, f"AH{r}", f"='Chiet tinh'!H{new}", bq.Range(f"AH{r}").Interior.Color, note, True)
        c = bq.Range(f"AJ{r}"); old = rgb(c); rc = bq.Range("AJ126")
        c.Font.Name, c.Font.Size = rc.Font.Name, rc.Font.Size; c.VerticalAlignment, c.HorizontalAlignment = rc.VerticalAlignment, rc.HorizontalAlignment
        c.NumberFormat = "@"; c.Value = it["ad"]; c.WrapText = True; c.Interior.Color = PINK
        log.append((BQ, f"AJ{r}", old, now, "desc REV04 loose KL", "f:" + it["ad"][:2000], ""))
        lr = last(run) + 1; run.Range(f"A{lr}:E{lr}").Value = ((f"{ma}-1", "MOI", new, RUN, now),)
        print(f"{ma}: block {new}-{new + src_end - SRC} | H={ws.Cells(new, 8).Text.strip()} | E={ws.Cells(new, 5).Text.strip()} | kim loại: " +
              ", ".join(f"{ws.Cells(s,1).Text.strip()[:18]}={ws.Cells(s,7).Text.strip()}kg/{ws.Cells(s,9).Text.strip()}" for s in slots[:len(it['metal'])]), flush=True)
    lg.Range(f"A{last(lg) + 1}:G{last(lg) + len(log)}").Value = tuple(log)
    wb.Saved = False; wb.Save(); wb.Close(False)
finally:
    xl.Quit()
