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
ROWS = [("F074", 83)]
AJ = NL.join(["Kích thước W1400xD580xH900 mm :", "Thùng tủ tấm MDF E1 chống ẩm hoàn thiện WD-04: nóc/đáy 25mm, hông 17-35mm, hậu 6mm, đợt/sắn giữa 17mm.",
              "Mặt ngoài cánh và ngăn kéo hoàn thiện SP-03", "Ngăn kéo gỗ Ash, thùng ngăn kéo MDF (3 ngăn kéo)", "Tấm nền vách MDF sơn lót chống ẩm", "Khung xương gỗ sơn lót chống ẩm.",
              "Không bao gồm:", "+ Mặt đá ST-01", "+ Chân tủ kim loại MT-02", "+ Tay nắm", "+ Tủ lạnh minibar 40L", "+ MEP", "+ Phụ kiện thoát nhiệt"])
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible, xl.DisplayAlerts, xl.AutomationSecurity, xl.EnableEvents = False, False, 1, False
try:
    wb = xl.Workbooks.Open(cpsx); ws = wb.Worksheets("Chiet tinh"); bq = wb.Worksheets(BQ)
    lg = wb.Worksheets("_EDITLOG"); run = next(s for s in wb.Worksheets if s.Name == "_AIRUN")
    SRC = next(x for x in range(2000, 3500) if str(ws.Cells(x, 1).Formula).replace("$", "").endswith("!A50"))
    for ma, r in ROWS: assert str(bq.Cells(r, 1).Value).strip() == ma and not str(bq.Cells(r, 34).Formula), ma
    def last(sh): return sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1
    log = []; now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M"); RUN = "REV04-MINIBAR3-" + now[:10]
    def rgb(c): return "" if c.Interior.ColorIndex == -4142 else "#%02X%02X%02X" % (int(c.Interior.Color) & 255, (int(c.Interior.Color) >> 8) & 255, (int(c.Interior.Color) >> 16) & 255)
    def put(sh, addr, val, color, note, formula=False):
        c = sh.Range(addr); old = rgb(c)
        if formula: c.Formula = val
        else: c.Value = val
        c.Interior.Color = color; log.append((sh.Name, addr, old, now, note, "f:" + str(val)[:200], ""))
    end = next(x for x in range(SRC + 1, SRC + 300) if ws.Cells(x, 1).Value == "END")
    new = next(x for x in range(5000, 20000) if ws.Cells(x, 1).Value == "END" and not ws.Cells(x + 1, 1).Value and not ws.Cells(x + 2, 1).Value) + 1
    ws.Rows(f"{SRC}:{end}").Copy(ws.Rows(new)); xl.CutCopyMode = False
    note = "REV04 minibar3 F074"; nend = new + (end - SRC)
    def find(label, occ=1, a=None, b=None):
        k = 0
        for x in range(a or new, (b or nend) + 1):
            if str(ws.Cells(x, 1).Value or "").strip() == label:
                k += 1
                if k == occ: return x
        raise KeyError((label, occ))
    put(ws, f"A{new}", f"=+'{BQ}'!A83", YEL, note, True)
    put(ws, f"E{new}", "=" + "+".join(f"'{BQ}'!AG{r}" for _, r in ROWS), YEL, note, True)
    for col, v in zip("BCD", (1400, 580, 900)): put(ws, f"{col}{new}", v, YEL, note)
    cnt = 0
    for x in range(new + 1, nend):
        if str(ws.Cells(x, 14).Value or "").strip().lower() == "wd-01":
            put(ws, f"N{x}", "wd-04", YEL, note); cnt += 1
    print("đổi wd-01→wd-04:", cnt, "ô", flush=True)
    for ma, r in ROWS:
        bq.Range("AH126").Copy(); bq.Range(f"AH{r}").PasteSpecial(-4122); xl.CutCopyMode = False
        put(bq, f"AH{r}", f"='Chiet tinh'!H{new}", bq.Range(f"AH{r}").Interior.Color, note, True)
        c = bq.Range(f"AJ{r}"); old = rgb(c); rc = bq.Range("AJ50")
        c.Font.Name, c.Font.Size = rc.Font.Name, rc.Font.Size; c.VerticalAlignment, c.HorizontalAlignment = rc.VerticalAlignment, rc.HorizontalAlignment
        c.NumberFormat = "@"; c.Value = AJ; c.WrapText = True; c.Interior.Color = PINK
        log.append((BQ, f"AJ{r}", old, now, "desc REV04 minibar3", "f:" + AJ[:2000], ""))
    lr = last(run) + 1; run.Range(f"A{lr}:E{lr}").Value = (("F074-1", "MOI", new, RUN, now),)
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
