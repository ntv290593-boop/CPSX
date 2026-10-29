"""REV04: tủ minibar type 2 (F042, F072, F073: 1300x585x1900, tờ GR-TY-720/721, JS-65-710/711, PS-TY-713/714) = chép block F041 (minibar type 1)
rồi chỉnh theo bản vẽ: tủ dưới 1300x585x550 (ngăn tủ lạnh 1 cánh + vách giữa + 3 ngăn kéo cao 80/125/220), vách kính trên 1300x585x1050,
bỏ len chân (chân kim loại). 3 dòng BoQ dùng chung 1 block (E = Σ AG). Kim loại/kính/đá/gương không tính (đồ không loose) → 'Không bao gồm'.
Mã hoàn thiện: bản vẽ ghi PT-07 (mặt ngoài ngăn kéo), Material chỉ có PT-01=PT.1 → dùng 'pt-01'. Làm trực tiếp trong REV04 (user 26/10).
Dùng: py -I build_minibar2_REV04_com.py <CPSX.xlsm>"""
import os, sys, datetime
import win32com.client
sys.stdout.reconfigure(encoding="utf-8")
cpsx = os.path.abspath(sys.argv[1]); d, n = os.path.split(cpsx)
if os.path.exists(os.path.join(d, "~$" + n)): sys.exit("file đang mở")
BQ = "MATRIX FITFUR-LOOSE"; YEL, PINK = 0x00FFFF, 0xCCCCFF; NL = "\n"
ROWS = [("F042", 51), ("F072", 81), ("F073", 82)]
AJ = NL.join(["Kích thước W1300xD585xH1900 mm :", "Thùng tủ tấm MDF E1 chống ẩm: nóc/đáy 25mm, hông 17mm, hậu 6mm, đợt/vách giữa 17mm.",
              "Ngăn tủ lạnh 1 cánh, 3 ngăn kéo (mặt ngoài MDF hoàn thiện PT-07, thùng ngăn kéo gỗ Ash).", "Vách kính phía trên: thùng tủ MDF E1 chống ẩm.",
              "Khung xương gỗ sơn lót chống ẩm.", "Không bao gồm:", "+ Cánh kính khung inox MT-02, kính GL-05, gương MR-02", "+ Kệ kính GL-03 khung MT-01",
              "+ Chân tủ kim loại MT-02", "+ Mặt đá ST-01", "+ Tay nắm", "+ Tủ lạnh minibar", "+ MEP", "+ Phụ kiện thoát nhiệt"])
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh"); bq = wb.Worksheets(BQ)
    lg = wb.Worksheets("_EDITLOG"); run = next(s for s in wb.Worksheets if s.Name == "_AIRUN")
    SRC = next(x for x in range(2000, 3500) if str(ws.Cells(x, 1).Formula).replace("$", "").endswith("!A50"))
    for ma, r in ROWS: assert str(bq.Cells(r, 1).Value).strip() == ma and not str(bq.Cells(r, 34).Formula), ma
    def last(sh): return sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1
    log = []; now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M"); RUN = "REV04-MINIBAR2-" + now[:10]
    def rgb(c): return "" if c.Interior.ColorIndex == -4142 else "#%02X%02X%02X" % (int(c.Interior.Color) & 255, (int(c.Interior.Color) >> 8) & 255, (int(c.Interior.Color) >> 16) & 255)
    def put(sh, addr, val, color, note, formula=False):
        c = sh.Range(addr); old = rgb(c)
        if formula: c.Formula = val
        else: c.Value = val
        c.Interior.Color = color; log.append((sh.Name, addr, old, now, note, "f:" + str(val)[:200], ""))
    end = next(x for x in range(SRC + 1, SRC + 300) if ws.Cells(x, 1).Value == "END")
    new = next(x for x in range(5000, 20000) if ws.Cells(x, 1).Value == "END" and not ws.Cells(x + 1, 1).Value and not ws.Cells(x + 2, 1).Value) + 1
    ws.Rows(f"{SRC}:{end}").Copy(ws.Rows(new)); xl.CutCopyMode = False
    note = "REV04 minibar2"; nend = new + (end - SRC)
    def find(label, occ=1, a=None, b=None):
        k = 0
        for x in range(a or new, (b or nend) + 1):
            if str(ws.Cells(x, 1).Value or "").strip() == label:
                k += 1
                if k == occ: return x
        raise KeyError((label, occ))
    put(ws, f"A{new}", f"=+'{BQ}'!A51", YEL, note, True)
    put(ws, f"E{new}", "=" + "+".join(f"'{BQ}'!AG{r}" for _, r in ROWS), YEL, note, True)
    for col, v in zip("BCD", (1300, 585, 1900)): put(ws, f"{col}{new}", v, YEL, note)
    t1 = find("TỦ", 1); t2 = find("TỦ", 2)
    for col, v in zip("BCD", (1300, 585, 550)): put(ws, f"{col}{t1}", v, YEL, note)
    for col, v in zip("BCD", (1300, 585, 1050)): put(ws, f"{col}{t2}", v, YEL, note)
    put(ws, f"E{find('Đợt ngang', 1, t1, t2)}", 0, YEL, note)           # ngăn tủ lạnh: không đợt
    put(ws, f"E{find('Cánh tủ', 1, t1, t2)}", 1, YEL, note)             # 1 cánh tủ lạnh
    for i, h in enumerate((80, 125, 220, None), start=1):               # 3 ngăn kéo bên phải, bỏ ngăn thứ 4
        rr = find("Ngăn kéo", i, t1, t2)
        if h is None: put(ws, f"E{rr}", 0, YEL, note)
        else:
            for col, v in (("B", 520), ("C", 450), ("D", h)): put(ws, f"{col}{rr}", v, YEL, note)
    lc = find("LEN CHÂN", 1, t2, nend); put(ws, f"E{lc}", 0, YEL, note)  # chân kim loại
    for ma, r in ROWS:
        bq.Range("AH126").Copy(); bq.Range(f"AH{r}").PasteSpecial(-4122); xl.CutCopyMode = False
        put(bq, f"AH{r}", f"='Chiet tinh'!H{new}", bq.Range(f"AH{r}").Interior.Color, note, True)
        c = bq.Range(f"AJ{r}"); old = rgb(c); rc = bq.Range("AJ50")
        c.Font.Name, c.Font.Size = rc.Font.Name, rc.Font.Size; c.VerticalAlignment, c.HorizontalAlignment = rc.VerticalAlignment, rc.HorizontalAlignment
        c.NumberFormat = "@"; c.Value = AJ; c.WrapText = True; c.Interior.Color = PINK
        log.append((BQ, f"AJ{r}", old, now, "desc REV04 minibar2", "f:" + AJ[:2000], ""))
    lr = last(run) + 1; run.Range(f"A{lr}:E{lr}").Value = (("F042-1", "MOI", new, RUN, now),)
    lg.Range(f"A{last(lg) + 1}:G{last(lg) + len(log)}").Value = tuple(log)
    wb.Saved = False; wb.Save()
    print(f"block {new}-{nend} (nguồn {SRC}) | hdr {ws.Cells(new,2).Text.strip()}x{ws.Cells(new,3).Text.strip()}x{ws.Cells(new,4).Text.strip()} E={ws.Cells(new,5).Text.strip()} | đơn giá H={ws.Cells(new,8).Text.strip()}", flush=True)
    for x in range(new + 1, nend):
        a = str(ws.Cells(x, 1).Text).strip(); e = str(ws.Cells(x, 5).Text).strip()
        if a.startswith(("TỦ", "Nóc", "Đáy", "Hông", "Hậu", "Đợt", "Cánh", "Sắn", "Ngăn kéo", "Mặt ngoài", "LEN", "GHÉP", "Đỡ", "đỡ", "Đáy ốp")) and e not in ("", "-", "0"):
            print(f"  {x} {a[:26]:26} B={ws.Cells(x,2).Text.strip():>5} C={ws.Cells(x,3).Text.strip():>5} D={ws.Cells(x,4).Text.strip():>5} E={e:>3} I={ws.Cells(x,9).Text.strip()}", flush=True)
    print("lỗi:", [(x, c) for x in range(new, nend + 1) for c in range(1, 63) if str(ws.Cells(x, c).Text).startswith("#") and set(str(ws.Cells(x, c).Text)) != {"#"}][:5], flush=True)
    wb.Close(False)
finally:
    xl.Quit()
