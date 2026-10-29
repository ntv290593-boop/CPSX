"""PROJECT_DATA cho gói VANITY (tủ lavabo + bàn đá + gương) của NORTON, theo cấu trúc v8 (NORTON cửa) + QUAN_SAT.
Dùng: py -I build_project_data_VANITY.py  -> _ACTIVE\\NORTON\\01_scan\\PROJECT_DATA_VANITY_v1.xlsx
Link tự trỏ về chính ô (location 'Sheet'!A1); đích thật nằm ở _LINKS (add-in SRC_LINKS.xlam hiện dropdown)."""
import os, re, json, sys
import openpyxl, pymupdf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
sys.stdout.reconfigure(encoding="utf-8")
ROOT = r"D:\COSTING EXCEL\_ACTIVE\NORTON"
INP = ROOT + r"\00_input\261024_VAnity"
OUT = ROOT + r"\01_scan\PROJECT_DATA_VANITY_v1.xlsx"
REL = "..\\00_input\\261024_VAnity\\"
RUN, MODEL = "VAN-SCAN-20261026", "claude-sonnet-5-5"
MASS = "01. DRAWING\\MASS UNITS\\261006 MASS UNITS.pdf"
LIFT = "01. DRAWING\\TYPICAL LIFT LOBBY & CORRIDOR\\261006_ID_TD_TYPICAL CORRIDOR & LIFT LOBBY.pdf"
BOQF = "27_TỦ WC.xlsx"
MAIL = "RE_ NORTON_ Báo giá hạng mục cửa gỗ.msg"
SPEC_T = "02.SPECIFICATION\\02 00 04 - TỦ DÁN NHỰA-PLASTIC LAMINATE CABINETS.pdf"
SPEC_S = "02.SPECIFICATION\\02 00 05 -  MẶT BÀN ĐÁ - STONE COUNTERTOPS.pdf"
SPEC_W = "02.SPECIFICATION\\03 00 01 - THIẾT BỊ VỆ SINH-SANITARY WARE.pdf"
ML1 = "03. MATERIAL LIST\\261005 - Material List - Skybridge Mass Units 001.pdf"
ML3 = "03. MATERIAL LIST\\261005 - Material List - Skybridge Mass Units 003.pdf"
MLL = "03. MATERIAL LIST\\261005 - Material List - Typical Lift Lobby & Coridor.pdf"
H = Font(bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="1F4E78"); LK = Font(color="0563C1", underline="single")
wb = Workbook(); wb.remove(wb.active)
REG = []

def sheet(name, cols, widths=None):
    ws = wb.create_sheet(name); ws.append(cols)
    for c in ws[1]: c.font = H; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A2"
    for i, w in enumerate(widths or [], 1): ws.column_dimensions[get_column_letter(i)].width = w
    return ws

def put(ws, r, c, v):
    cell = ws.cell(r, c, v)
    if isinstance(v, str) and v.startswith("="): cell.data_type = "s"
    return cell

def setlinks(ws, r, c, items, text=None):
    items = [i for i in items if i]
    if not items: return
    cell = ws.cell(r, c)
    cell.value = text or (items[0][0] if len(items) == 1 else f"{items[0][0]}  ▾{len(items)}")
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{ws.title}'!{cell.coordinate}", display=str(cell.value)); cell.font = LK
    for it in items: REG.append((ws.title, cell.coordinate) + it)

def FL(rel, lab=None, p="", sub=""): return (lab or rel.split("\\")[-1], rel, p, sub)
def PG(p, lab=None, f=MASS): return (lab or f"p{p}", f, p, "")
def BQ(sh, r): return (f"BoQ {sh}!{r}", BOQF, "", f"'{sh}'!A{r}")
def SP(f, p, lab): return (lab, f, p, "")

# ---------------- đọc nguồn ----------------
bq = openpyxl.load_workbook(INP + "\\" + BOQF, data_only=False)
ZONE = {"B5.16 SD": ("PODIUM BLK01", "5.16"), "B6.16 SD": ("PODIUM BLK02", "6.16"), "B8.18 SD": ("TOWER BLK01", "8.18"),
        "B9.18 SD": ("TOWER BLK02", "9.18"), "B10.18 SD": ("TOWER BLK03", "10.18")}
ZC = [v[0] for v in ZONE.values()]
ITEMS = [("V01", "Tủ lavabo", 11), ("V02", "Bàn đá (bao gồm khung)", 12), ("V03", "Gương MCB", 16)]
qty = {}  # (mã, sheet) -> SL
for sh in ZONE:
    ws = bq[sh]
    for ma, _, r in ITEMS: qty[(ma, sh)] = ws.cell(r, 4).value
D9 = bq["B5.16 SD"]["B9"].value; D16 = bq["B5.16 SD"]["B16"].value
tot = {ma: sum(qty[(ma, s)] for s in ZONE) for ma, _, _ in ITEMS}

# ---------------- trang bản vẽ ----------------
d = pymupdf.open(INP + "\\" + MASS)
dl = pymupdf.open(INP + "\\" + LIFT)
pat = re.compile(r"([0-9A-Z][0-9A-Za-z.&_\-]*-(?:TD-)?\d{4})")
tailp = re.compile(r"\d\d/\d\d/\d{4} [A-Z] (.*)$")
pages = []
for i, p in enumerate(d):
    t = p.get_text(); T = re.sub(r"\s+", " ", t)
    toks = pat.findall(T); no = toks[-1] if toks else ""
    m = tailp.search(T[-300:]); title = (m.group(1) if m else T[-120:]).strip()
    kind = ""
    if "VC1" in t and "FURNITURE PLAN" in T.upper(): kind = "MB nội thất WC (VC1)"
    elif ("GƯƠNG" in t or "LA1" in t) and "ELEVATION" in T.upper() and "FC1" in t: kind = "Mặt đứng WC (lavabo/gương)"
    elif re.search(r"VANITY COUNTER|MIRROR|GƯƠNG", t): kind = "Có nhắc vanity/gương (legend/phối cảnh)"
    pages.append(dict(p=i + 1, no=no, title=title, kind=kind))
PLAN_SEEN = {10, 22, 58, 70, 74, 90, 150, 154, 170, 198, 202, 214, 218, 230, 234, 246, 250, 266, 279, 283}
ELEV_SEEN = {12, 13, 24, 25, 60, 61, 72, 73, 76, 77, 92, 93, 152, 153, 156, 157, 200, 201, 204, 205, 216, 217, 220, 221, 232, 233,
             236, 237, 248, 249, 252, 253, 268, 269, 281, 282, 285, 286}
for x in pages:
    if x["kind"].startswith("MB nội thất"): x["read"] = "HÌNH" if x["p"] in PLAN_SEEN else "TEXT (trùng tiêu đề với tờ đã xem)"
    elif x["kind"].startswith("Mặt đứng"): x["read"] = "HÌNH" if x["p"] in ELEV_SEEN else "TEXT (trùng tiêu đề với tờ đã xem)"
    else: x["read"] = "TEXT" if x["kind"] else "Không liên quan vanity"
NPT = sum(1 for x in pages if x["kind"].startswith("MB nội thất")); NPS = sum(1 for x in pages if x["kind"].startswith("MB nội thất") and x["read"] == "HÌNH")
NET = sum(1 for x in pages if x["kind"].startswith("Mặt đứng")); NES = sum(1 for x in pages if x["kind"].startswith("Mặt đứng") and x["read"] == "HÌNH")

# ---------------- PROJECT ----------------
ws = sheet("PROJECT", ["Trường", "Giá trị", "Nguồn", "→ OverView"], [28, 78, 40, 14])
ML_VAN = FL(ML1, "ML Mass Units 001 p1 (LAM-02, SS02)", 1)
rows = [
 ("Tên dự án", "TỔ HỢP THƯƠNG MẠI SB – Chung cư khu B (NORTON) / Skybridge Plot B", [FL(MASS, "Bìa MASS UNITS p1", 1)], ""),
 ("CĐT", "Skybridge Co., Ltd (Gamuda Land)", [FL(MAIL, "Email .msg")], ""),
 ("Tư vấn thiết kế nội thất", "DWP Vietnam (ID) / ICU (KTS địa phương) – theo khung tên", [FL(MASS, "Khung tên p10", 10)], ""),
 ("Gói thầu (đợt 2)", "Bổ sung scope VANITY: tủ lavabo + bàn đá (gồm khung) + gương MCB (BoQ Bill x.16 / x.18 SUNDRIES)", [FL(BOQF, "BoQ 27_TỦ WC.xlsx")], ""),
 ("Người yêu cầu", "Phạm Quốc Bảo (Tender) – mail 24/10/2026 08:10", [FL(MAIL, "Email .msg")], ""),
 ("Hạn trả giá", "28/10/2026 (trong ngày)", [FL(MAIL, "Email .msg")], ""),
 ("Đợt hồ sơ", "261024_VAnity: BoQ bổ sung + Material List 4 file + Spec 13 file + MASS UNITS (290 tờ) + Typical Lift Lobby (44 tờ) + CAD", [], ""),
 ("Phương án hoàn thiện tủ lavabo", "PA1 laminate vân da Formica 5342L (Earth) · PA2 laminate vân da Wilsonart D96-L (Shadow) · PA3 melamine (CĐT chưa ghi mã/hãng)", [FL(MAIL, "Email .msg"), FL(ML1, "ML 001 LAM-02", 1)], "C52 PA_LAVABO"),
 ("Mặt bàn lavabo", "SS02 Himacs/Krion M553 Ebony Concrete dày 12mm (ML + legend BV) — user chốt 26/10/2026: tính theo SS02", [ML_VAN, PG(12, "BV p12 legend SS02")], "Material SS02"),
 ("Gương", "MCB Glasso customize, treo tường, khung bo, 850×1150×180 (ML) – BV: 900 rộng × 1100 cao", [FL(ML3, "ML 003 p1 MCB", 1), PG(13, "BV p13 E3")], ""),
 ("Tủ lavabo – Spec", "Spec 02 00 04: HPL trên MDF, không khung (frameless), full overlay, reveal 3mm, bản lề chìm, tay nắm gắn sau, chốt nam châm", [SP(SPEC_T, 4, "02 00 04 p4 §2.02"), SP(SPEC_T, 5, "02 00 04 p5 §2.05")], ""),
 ("Mặt bàn – Spec", "Spec 02 00 05: marble 18mm bóng/honed, Saigon Casa (mâu thuẫn SS02)", [SP(SPEC_S, 3, "02 00 05 p3 §2.02")], ""),
 ("Phạm vi tính giá", "BOHO gia công: thân tủ + cánh + khung thép mạ kẽm + hộp gương. Mặc định tính kim loại. Không tính chậu LA1, vòi FC1, xi-phông (TBVS 03 00 01), bệ đá 350×120 khu tắm đứng.", [], "Phạm vi"),
 ("Vị trí", "Phòng tắm căn hộ BT01/BT02 (Studio, 1BR, 2BR, 3BR) — Mass Units; không có ở Typical Lift Lobby (đã quét 44 tờ: 0 lần nhắc)", [PG(1, "Lift lobby p1", LIFT)], ""),
]
for i, (a, b, lk, o) in enumerate(rows, 2):
    ws.append([a, b, "", o]); setlinks(ws, i, 3, lk)
    ws.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")

# ---------------- FILES ----------------
ws = sheet("FILES", ["ID", "Loại", "Tên file", "Link", "Đường dẫn (trong 261024_VAnity)", "Ext", "MB", "Trạng thái scan"], [7, 14, 60, 8, 70, 7, 8, 40])
files = []
for dp, dn, fn in os.walk(INP):
    for f in fn: files.append(os.path.relpath(os.path.join(dp, f), INP))
files.sort()
def cat(p):
    if p.endswith(".msg"): return "Mail"
    if p.startswith("27_"): return "BoQ"
    if p.startswith("02."): return "Spec"
    if p.startswith("03."): return "Material List"
    if p.lower().endswith((".dwg", ".dwt", ".dst", ".dwl", ".dwl2", ".pcp")): return "CAD (đáp án)"
    return "Bản vẽ"
STAT = {MAIL: "Đã đọc", BOQF: "Đã đọc", MASS: "Đã quét 290 tờ; xem hình các tờ WC (PAGES)", LIFT: "Đã quét text 44 tờ: không có vanity",
        SPEC_T: "Đã đọc", SPEC_S: "Đã đọc", ML1: "Đã đọc", ML3: "Đã đọc", MLL: "Đã đọc"}
for i, p in enumerate(files, 1):
    sz = os.path.getsize(os.path.join(INP, p)) / 1e6; c = cat(p)
    st = STAT.get(p) or ("DWG chỉ làm đáp án, chưa mở" if c == "CAD (đáp án)" else ("Đã trích text, không liên quan vanity" if p.endswith(".pdf") else "Đã liệt kê"))
    ws.append([f"F{i:03d}", c, os.path.basename(p), "", p, os.path.splitext(p)[1], round(sz, 2), st]); setlinks(ws, i + 1, 4, [FL(p, "Mở")])
ws.auto_filter.ref = ws.dimensions

# ---------------- PAGES ----------------
ws = sheet("PAGES", ["File", "Trang", "Số BV", "Tên BV", "Nội dung liên quan", "Đã đọc bằng", "Mở PDF"], [14, 6, 28, 70, 30, 34, 10])
r = 1
for x in pages:
    r += 1; ws.append(["MASS UNITS", x["p"], x["no"], x["title"], x["kind"], x["read"], ""]); setlinks(ws, r, 7, [PG(x["p"], f"p{x['p']}")])
for i, p in enumerate(dl):
    r += 1; t = re.sub(r"\s+", " ", p.get_text()); m = tailp.search(t[-300:]) or re.search(r"(\S+) (TỜ.*|.*)$", t[-120:])
    ws.append(["LIFT LOBBY", i + 1, "", t[-110:].strip(), "", "TEXT: không có vanity", ""]); setlinks(ws, r, 7, [PG(i + 1, f"p{i + 1}", LIFT)])
ws.auto_filter.ref = ws.dimensions

# ---------------- BOQ_CDT ----------------
ws = sheet("BOQ_CDT", ["Sheet", "Dòng", "Bill", "Khu", "Mã", "Hạng mục", "Mô tả", "ĐV", "KL CĐT", "Mô tả kỹ thuật", "Link"], [10, 6, 7, 16, 6, 24, 60, 5, 8, 90, 14])
r = 1
for sh, (zone, bill) in ZONE.items():
    w = bq[sh]
    for ma, name, row in ITEMS:
        r += 1
        desc = w.cell(row, 2).value; un = w.cell(row, 3).value
        spec = D9 if ma in ("V01", "V02") else D16
        ws.append([sh, row, bill, zone, ma, name, desc, un, qty[(ma, sh)], spec]); setlinks(ws, r, 11, [BQ(sh, row)])
        ws.cell(r, 10).alignment = Alignment(wrap_text=True, vertical="top")
ws.auto_filter.ref = ws.dimensions

# ---------------- BOQ_TONGHOP ----------------
cols = ["STT", "Mã", "Hạng mục", "Mô tả (BoQ)", "Mô tả kỹ thuật (BoQ)", "Kích thước BoQ", "Kích thước BV", "Mã VL/HT (Spec·BV·ML)", "ĐV"] + ZC + \
       ["Tổng BoQ", "Dòng BoQ gốc", "Chi tiết BV", "Spec", "Material List", "Issue", "Loại đề xuất", "Mẫu đề xuất (user chốt)", "Giả định tính giá", "Độ chắc"]
ws = sheet("BOQ_TONGHOP", cols, [5, 5, 24, 40, 60, 16, 22, 44, 5] + [10] * len(ZC) + [9, 22, 26, 26, 26, 16, 22, 34, 50, 8])
TH = [
 ("V01", "Tủ lavabo", "Supply and install approved vanity counter top fixed ... (dòng con 'Tủ lavabo')", "900×550×440 H", "900 (R) × 550 (S); mặt bàn cao 850 AFFL; cao thân tủ không ghi",
  "LAM-02 (ML 001): Formica 5342L / Wilsonart D96-L vân da, WPB; PA3 melamine (mail)", "I02,I03,I08,I09,I12", "Tủ lavabo treo tường 2 cánh",
  "Template#7 Vanity (user chốt 26/10) — Giá theo md → dùng cho Nr: header 900×550×440; 2 cánh, không ngăn kéo, bật THÙNG FIX (hộp chậu)", "PA1/PA2 laminate trên WPB; PA3 melamine MDF E1 chống ẩm (tạm tính); không tính tay nắm; nhân công giữ như mẫu", "TB",
  [PG(12, "BV p12 E2/E3"), PG(13, "BV p13 E3/E4"), PG(24, "1BR p24-25"), PG(72, "2BR p72-73"), PG(200, "3BR p200-201")], [SP(SPEC_T, 4, "02 00 04 §2.02"), SP(SPEC_T, 5, "02 00 04 §2.05")], [FL(ML1, "ML 001 LAM-02", 1)]),
 ("V02", "Bàn đá (gồm khung)", "idem (dòng con 'Bàn đá (bao gồm khung)')", "(trong 900×550×440)", "900 (R) × 550 (S); chậu LA1 âm dưới mặt; khung hộp quanh chậu (mặt cắt E3)",
  "SS02 Himacs/Krion M553 Ebony Concrete 12mm (ML 001, legend BV) ↔ BoQ: laminate + khung thép mạ kẽm + bracket ↔ Spec 02 00 05: marble 18mm", "I01,I06,I10", "Mặt bàn solid surface + khung thép",
  "Nóc cemboard có sẵn trong block V01 (Template#7); V02 bổ sung viền trước + 2 hông cemboard cao 100 + xương đỡ đá MTGC (Template hàng 771-774) + SS02 12mm", "Tính SS02 (user chốt); viền cemboard cao 100 giả định; đơn giá SS02/MTGC để user điền; không tính chậu", "TB",
  [PG(13, "BV p13 E3"), PG(12, "BV p12 legend SS02")], [SP(SPEC_S, 3, "02 00 05 §2.02")], [FL(ML1, "ML 001 SS02", 1)]),
 ("V03", "Gương MCB", "Supply and install approved wall-mounted framed mirror with polished edges ... (Type MCB)", "850×1150×180 (user chốt tính 900×180×1100)", "BV: cao 1100, cách mặt bàn 350, rộng ≈ rộng bàn (không ghi), sâu ≈120–150 (không ghi); hộp + ngăn kệ (mặt đứng bên)",
  "MCB Glasso customize, treo tường, khung bo, mài cạnh (ML 003)", "I04,I05,I11", "Tủ gương hộp có ngăn kệ",
  "Template#8 Tủ gương hộp (R0084): header 900×180×1100, 2 cánh gương; gương 5mm + mài cạnh + dán", "User chốt 900×180×1100, 2 cánh; hoàn thiện hộp theo phương án tủ lavabo; gương 5mm tạm tính", "TB",
  [PG(13, "BV p13 E2/E3"), PG(93, "2BR p93"), PG(236, "3BR p236")], [], [FL(ML3, "ML 003 p1 MCB", 1)]),
]
for i, t in enumerate(TH, 1):
    ma, nm, desc, szb, szv, mat, iss, loai, mau, ass, cf, det, spc, mlr = t
    ws.append([i, ma, nm, desc, (D9 if ma != "V03" else D16), szb, szv, mat, "nr"] + [qty[(ma, s)] for s in ZONE] +
              [tot[ma], "", "", "", "", iss, loai, mau, ass, {"TB": 0.7, "THẤP": 0.4, "CAO": 0.9}[cf]])
    rr = i + 1; c0 = 10 + len(ZC)
    setlinks(ws, rr, c0 + 1, [BQ(sh, dict((a, c) for a, _, c in ITEMS)[ma]) for sh in ZONE])
    setlinks(ws, rr, c0 + 2, det); setlinks(ws, rr, c0 + 3, spc); setlinks(ws, rr, c0 + 4, mlr)
    for c in (4, 5, 7, 8, 18, 19): ws.cell(rr, c).alignment = Alignment(wrap_text=True, vertical="top")
ws.auto_filter.ref = ws.dimensions; ws.freeze_panes = "D2"

# ---------------- BOQ_MAP ----------------
ws = sheet("BOQ_MAP", ["Mã", "Khu", "KL BoQ", "KL theo BV", "Lệch", "Kích thước BV", "Tờ MB nội thất", "Tờ mặt đứng", "Spec", "Material List", "Issue", "Trạng thái", "Độ chắc"],
           [6, 16, 8, 14, 22, 22, 26, 26, 24, 24, 16, 22, 8])
r = 1
for ma, nm, row in ITEMS:
    for sh, (zone, bill) in ZONE.items():
        r += 1
        ws.append([ma, zone, qty[(ma, sh)], "KHÔNG THẤY", "n/a", "900×550" if ma != "V03" else "H1100 (R,S không ghi)", "", "", "", "",
                   {"V01": "I02,I03,I08,I09,I12", "V02": "I01,I06,I10", "V03": "I04,I05,I11"}[ma], "KHÔNG ĐỐI CHIẾU ĐƯỢC SL (BV không có thống kê vanity/căn)", 0.5])
        setlinks(ws, r, 7, [PG(10, "p10 Studio"), PG(22, "p22 1BR"), PG(70, "p70 2BR"), PG(198, "p198 3BR")])
        setlinks(ws, r, 8, [PG(12, "p12 Studio"), PG(24, "p24 1BR"), PG(72, "p72 2BR"), PG(200, "p200 3BR")])
        setlinks(ws, r, 9, [SP(SPEC_T, 4, "02 00 04")] if ma == "V01" else ([SP(SPEC_S, 3, "02 00 05")] if ma == "V02" else []))
        setlinks(ws, r, 10, [FL(ML3 if ma == "V03" else ML1, "ML", 1)])
ws.auto_filter.ref = ws.dimensions

# ---------------- ISSUES ----------------
I = [
 ("I01", "LỆCH", "Spec↔BoQ↔ML/BV", "V02", "Mặt bàn lavabo: BoQ ghi laminate + khung thép mạ kẽm; ML + legend BV ghi SS02 Himacs/Krion M553 Ebony Concrete 12mm (solid surface); Spec 02 00 05 ghi marble 18mm Saigon Casa. ML chỉ dùng marble ST02 cho bệ bồn tắm.",
  "User chốt 26/10: tính SS02; gửi RFI-01 xin xác nhận vật liệu", "CAO", "RFI-01", 1.0, [ML_VAN, SP(SPEC_S, 3, "02 00 05 §2.02"), BQ("B5.16 SD", 9), PG(12, "BV p12 legend")]),
 ("I02", "LỆCH", "Spec↔ML", "V01", "Cốt tủ: Spec 02 00 04 ghi MDF (ANSI A208.2 ≥750kg/m3) + HPL; ML LAM-02 ghi tấm WPB hoàn thiện laminate vân da.", "Tính WPB theo ML (PA1/PA2); RFI-02", "TB", "RFI-02", 0.8,
  [SP(SPEC_T, 4, "02 00 04 §2.03"), ML_VAN]),
 ("I03", "THIẾU", "Mail↔ML", "V01", "PA3 'melamine': ML và Spec không có mã/hãng/mã màu melamine cho tủ lavabo; ML chỉ có 2 option laminate (Formica 5342L, Wilsonart D96-L).", "Tính melamine tạm tính (ME-01 file mẫu); RFI-03", "TB", "RFI-03", 1.0,
  [FL(MAIL, "Email .msg"), ML_VAN]),
 ("I04", "LỆCH", "BV↔BoQ/ML", "V03", "Gương: BoQ/ML 850×1150×180; BV cao 1100, cách mặt bàn 350, rộng ≈ rộng bàn 900 (không ghi số), sâu hộp không ghi số (≈120–150 theo hình bên).", "User chốt 26/10: 900×180×1100 (BV rộng/cao + sâu 180 BoQ/ML); RFI-04 xin xác nhận", "TB", "RFI-04", 0.9,
  [FL(ML3, "ML 003 p1", 1), PG(13, "BV p13 E3"), PG(236, "BV p236 E1")]),
 ("I05", "THIẾU", "BV/Spec", "V03", "Hộp gương: mã hoàn thiện khung/hộp không ghi (khung bo); không có Spec mục gương/kính gương (01 00 04 Glazing không nhắc mirror).", "Hoàn thiện hộp theo phương án tủ lavabo (giả định); RFI-04", "TB", "RFI-04", 1.0, [SP(SPEC_W, 1, "Spec 03 00 01"), PG(13, "BV p13")]),
 ("I06", "THIẾU", "BV↔BoQ", "V02", "Bệ đá SS02 350×120mm khu tắm đứng (callout 'BỆ ĐÁ') có trên MB nhiều loại căn nhưng không có trong BoQ vanity.", "Không tính; RFI-05", "TB", "RFI-05", 1.0, [PG(74, "BV p74"), PG(150, "BV p150"), PG(214, "BV p214")]),
 ("I07", "INFO", "BV nội bộ", "V01,V02", "Căn 1BR A1.1a/A1.3a/A2.1a: sâu vanity 540 (BV p22/34/46); loại khác 550. Rộng 900 ở mọi loại.", "Dùng 900×550 chung", "THẤP", "—", 0.8, [PG(22, "BV p22"), PG(70, "BV p70")]),
 ("I08", "THIẾU", "BV↔BoQ", "V01", "BoQ ghi 'overall 900×550×440 high'; BV chỉ ghi mặt bàn cao 850 AFFL, không có số cao thân tủ (440).", "Dùng 440 theo BoQ", "THẤP", "—", 0.7, [BQ("B5.16 SD", 9), PG(13, "BV p13 E3")]),
 ("I09", "INFO", "Phạm vi", "V01", "Cánh 2 cánh vân da, bản lề chìm, tay nắm gắn sau, chốt nam châm theo Spec 02 00 04 §2.05; BoQ không nêu tay nắm.", "Tính bản lề; tay nắm theo mẫu; ghi vào mô tả", "THẤP", "—", 0.8, [SP(SPEC_T, 5, "02 00 04 §2.05")]),
 ("I10", "INFO", "BoQ nội bộ", "V02", "Dòng 9 (mô tả vanity) không có SL/ĐV; SL ở 2 dòng con. Dòng 'Total carried to Collection' F22 = SUBTOTAL(F19:F20) không gồm dòng hàng (lỗi tổng của BoQ CĐT). Bill 7 (Podium Block 03) không có vanity.", "Nhập đơn giá ở dòng con; sửa tổng trang; ghi chú", "THẤP", "—", 1.0, [BQ("B5.16 SD", 22)]),
 ("I11", "INFO", "Đối chiếu KL", "V01-V03", f"Tổng {tot['V01']} bộ = 5 + 10 + 170 + 179 + 149; BV không có thống kê SL căn hộ theo tháp nên không đối chiếu được.", "Dùng KL BoQ", "THẤP", "—", 0.5, [BQ("B8.18 SD", 11)]),
 ("I12", "INFO", "ML nội bộ", "V01", "Mã LAM-02 dùng cho 2 vật liệu khác nhau: Mass Units (laminate vân da tủ lavabo) và Typical Lift Lobby (8164-TX Umber Silk).","Tránh nhầm khi tra mã; ghi rõ vị trí trong mô tả", "THẤP", "—", 1.0, [ML_VAN, FL(MLL, "ML Lift Lobby p1", 1)]),
 ("I13", "INFO", "Thiết bị", "V02", "Chậu LA1 (âm dưới mặt), vòi FC1, xi-phông thuộc TBVS (Spec 03 00 01); mặt bàn phải khoét lỗ chậu.", "Không tính thiết bị; có công khoét lỗ", "THẤP", "—", 0.8, [SP(SPEC_W, 1, "03 00 01"), PG(13, "BV p13 E3")]),
]
ws = sheet("ISSUES", ["ID", "Trạng thái", "Loại", "Mã", "Nội dung", "Hành động", "Mức", "RFI", "Độ tin cậy", "Nguồn (bấm chọn)"], [6, 9, 16, 8, 90, 44, 7, 8, 9, 34])
for i, (a, st, ty, ma, nd, act, lv, rf, cf, lk) in enumerate(I, 2):
    ws.append([a, st, ty, ma, nd, act, lv, rf, cf, ""]); setlinks(ws, i, 10, lk)
    for c in (5, 6): ws.cell(i, c).alignment = Alignment(wrap_text=True, vertical="top")
ws.auto_filter.ref = ws.dimensions

# ---------------- RFI_LOG ----------------
R = [("RFI-01", "Mặt bàn lavabo: BoQ (laminate + khung thép) ↔ ML/BV (SS02 Himacs/Krion M553 12mm) ↔ Spec 02 00 05 (marble 18mm). Xin xác nhận vật liệu và chủng loại.", "I01"),
     ("RFI-02", "Cốt tủ lavabo: WPB (ML) hay MDF (Spec 02 00 04)? Xin xác nhận.", "I02"),
     ("RFI-03", "Phương án 3 'melamine': xin mã/hãng/màu melamine cho tủ lavabo.", "I03"),
     ("RFI-04", "Gương MCB: kích thước 850×1150×180 (BoQ/ML) hay 900×1100 (BV)? Độ sâu hộp, hoàn thiện khung/hộp, ngăn kệ.", "I04,I05"),
     ("RFI-05", "Bệ đá SS02 350×120 khu tắm đứng có thuộc gói vanity không?", "I06")]
ws = sheet("RFI_LOG", ["RFI", "Câu hỏi", "Issue", "Ngày gửi", "Trả lời", "Trạng thái"], [8, 100, 10, 10, 40, 10])
for x in R: ws.append(list(x[:3]) + ["", "", "Nháp"]);
for c in ws["B"][1:]: c.alignment = Alignment(wrap_text=True, vertical="top")

# ---------------- SPEC (02 00 04, 02 00 05) ----------------
ws = sheet("SPEC", ["Mục", "Tiêu đề", "Nội dung", "Spec", "Trang", "Mở"], [8, 26, 100, 12, 6, 18])
ri = 1
for sf, tag in ((SPEC_T, "02 00 04"), (SPEC_S, "02 00 05")):
    doc = pymupdf.open(INP + "\\" + sf)
    for pn, pg in enumerate(doc, 1):
        lines = [l.strip() for l in pg.get_text().split("\n") if l.strip()]
        txt = " ".join(lines); txt = re.sub(r"\s+", " ", txt)
        for m in re.finditer(r"(?:^| )((?:[123]\.\d\d) [A-ZÀ-Ỹ ,/&-]{4,60}?)(?= [A-Z0-9a-zÀ-ỹ])", txt):
            pass
        ri += 1; ws.append([f"p{pn}", "", txt[:1800], tag, pn, ""]); setlinks(ws, ri, 6, [SP(sf, pn, f"{tag} p{pn}")])
        ws.cell(ri, 3).alignment = Alignment(wrap_text=True, vertical="top")

# ---------------- NOTES ----------------
ws = sheet("NOTES", ["#", "Ghi chú", "Loại", "Nguồn", "Độ tin cậy"], [4, 100, 12, 34, 9])
NT = [("Spec 02 00 04: Grade GP28; frameless; full overlay; reveal 3mm; MDF ≥750kg/m3; keo không urea formaldehyde; tay nắm gắn sau; chốt nam châm; bản lề chìm", "spec", [SP(SPEC_T, 4, "02 00 04 §2.02-2.05")], 0.9),
      ("Spec 02 00 05: nhà cung cấp Saigon Casa; mẫu 300×300; mối nối 2mm; không cắt đá tại công trường", "spec", [SP(SPEC_S, 3, "02 00 05 §2.02")], 0.9),
      ("ML Mass Units 001: LAM-02 Option 1 Formica 5342L (Earth), Option 2 Wilsonart D96-L (Shadow), 1200×2400, WPB laminate vân da; SS02 HIMACS/KRION M553 Ebony Concrete 12mm", "material", [ML_VAN], 1.0),
      ("ML 003: gương MCB Glasso customize, mài cạnh, khung bo, 850×1150×180", "material", [FL(ML3, "ML 003 p1", 1)], 1.0),
      ("Mail 24/10: báo giá 3 phương án (PA1 Formica 5342L, PA2 Wilsonart D96-L, PA3 melamine), trả giá trong ngày 28/10/2026", "mail", [FL(MAIL, "Email .msg")], 1.0),
      ("User 26/10: dùng file 260903_DA_CPSX_RE00.xlsm làm nền, không dùng file cửa đã báo; đơn giá 3 phương án và SS02 user xử lý sau", "user", [], 1.0)]
for i, (n, t, lk, cf) in enumerate(NT, 2):
    ws.append([i - 1, n, t, "", cf]); setlinks(ws, i, 4, lk); ws.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")

# ---------------- ASK_USER ----------------
ws = sheet("ASK_USER", ["#", "Câu hỏi nội bộ", "Trả lời"], [4, 100, 40])
AQ = [("Mặt bàn tính SS02 solid surface 12mm?", "SS02 (user 26/10)"), ("3 phương án áp cho phần nào + giá?", "Cứ dựng công tắc chọn phương án trong file; đơn giá user xử lý sau (user 26/10)"),
      ("Gương tính theo mẫu nào?", "Tủ gương hộp (user 26/10)"),
      ("Mẫu cho tủ lavabo: Template#7 Vanity hay TRỐNG (quy ước SOFITEL dùng TRỐNG)?", "Vanity (Template#7) (user 26/10)"),
      ("Bàn đá + khung: TRỐNG hay gộp vào Template#7?", "Nóc cemboard đã có ở module tủ của #7; bổ sung viền mặt trước + 2 hông cemboard cao 100 (giả định); xương đỡ đá tham khảo Template hàng 771-774, mã MTGC (user 26/10). Tủ vanity có 1 thùng fix bên cạnh thùng 2 cánh."),
      ("Hoàn thiện hộp gương theo phương án tủ lavabo?", "Có, hộp gương 2 cánh (user 26/10)"),
      ("Gương 850×1150×180 hay 900×1100?", "900×180×1100 (user 26/10)"), ("Tay nắm?", "Không bao gồm (user 26/10)"),
      ("Bệ đá SS02 350×120 có thuộc gói?", "Không bao gồm (user 26/10)")]
for i, (q, a) in enumerate(AQ, 1): ws.append([i, q, a])

# ---------------- QUAN_SAT ----------------
ws = sheet("QUAN_SAT", ["Mã", "Thành phần", "Trường", "Giá trị", "Trạng thái", "Tờ · Chi tiết", "Nguồn", "RunId", "Model", "Bản DS loại"], [6, 8, 26, 70, 14, 40, 8, 20, 18, 10])
Q = []
def q(ma, tp, f, v, st, ref, src, lk=None): Q.append((ma, tp, f, v, st, ref, src, lk))
E = [PG(12, "p12 E2/E3 ST1.1a"), PG(24, "p24-25 1BR"), PG(72, "p72-73 2BR"), PG(200, "p200-201 3BR")]
# V01
q("V01", "1", "Lắp đặt", "Treo tường (không chân), dưới chậu LA1, mặt trên cao 850 AFFL", "CÓ", "p13 E3; p24 E2", "HÌNH", E)
q("V01", "1", "Kích thước BV", "900 (R) × 550 (S); 1BR A1.1a/A1.3a/A2.1a sâu 540; cao thân tủ KHÔNG GHI", "CÓ", "MB p10,22,58,70,74,90…; E3 p13", "HÌNH", [PG(10, "MB p10"), PG(22, "MB p22")])
q("V01", "1", "Kích thước BoQ", "900×550×440 high", "CÓ", "BoQ B5.16!9", "BOQ", [BQ("B5.16 SD", 9)])
q("V01", "1", "Có cánh", "2 cánh mở, vân da (LAM-02), tay nắm không thấy (gắn sau theo spec)", "CÓ", "p13 E3; p72 E2/E3", "HÌNH", E)
q("V01", "1", "Ngăn kéo", "Mặt đứng chỉ 2 cánh, không đường chia ngăn kéo", "KHÔNG CÓ", "p13 E3; p73 E3", "HÌNH", E)
q("V01", "1", "Kệ/đợt", "Mặt cắt E3 chỉ thấy xi-phông; không thấy đợt giữa", "KHÔNG THẤY", "p13 E3 (mặt cắt)", "HÌNH", [PG(13, "p13")])
q("V01", "1", "Vật liệu gỗ (mã WD/PT)", "LAM-02: Option1 Formica 5342L / Option2 Wilsonart D96-L (vân da), tấm WPB; PA3 melamine KHÔNG có mã", "CÓ", "ML 001 p1; mail 24/10", "SPEC", [ML_VAN])
q("V01", "1", "Vật liệu theo Spec", "HPL trên MDF ≥750kg/m3, frameless, full overlay, reveal 3mm", "MÂU THUẪN", "02 00 04 §2.02-2.03", "SPEC", [SP(SPEC_T, 4, "02 00 04 p4")])
q("V01", "1", "Ván nền ẩn sau lớp hoàn thiện", "Không (tủ hoàn thiện laminate cả cánh và thân; sau tủ là tường gạch TL04)", "KHÔNG CÓ", "p13 E3", "HÌNH", [PG(13, "p13")])
q("V01", "1", "Khung xương", "Không thấy trong mặt đứng/mặt cắt", "KHÔNG THẤY", "p13 E3", "HÌNH", [PG(13, "p13")])
q("V01", "1", "Phần ngoài phạm vi", "Chậu LA1, vòi FC1, xi-phông, mặt bàn (V02), tay nắm?", "CÓ", "p13 E3", "HÌNH", [PG(13, "p13")])
# V02
q("V02", "1", "Mặt đá", "SS02 solid surface Himacs/Krion M553 Ebony Concrete 12mm", "CÓ", "ML 001 p1; legend BV p12", "SPEC", [ML_VAN, PG(12, "legend p12")])
q("V02", "1", "Mặt đá theo BoQ", "Laminate + khung thép mạ kẽm + bracket đỡ", "MÂU THUẪN", "BoQ B5.16!9", "BOQ", [BQ("B5.16 SD", 9)])
q("V02", "1", "Mặt đá theo Spec", "Marble 18mm Saigon Casa, bóng/honed, backsplash/end splash 18mm", "MÂU THUẪN", "02 00 05 §2.02, §2.05", "SPEC", [SP(SPEC_S, 3, "02 00 05 p3")])
q("V02", "1", "Kích thước BV", "900 × 550; chậu LA1 âm dưới mặt; diềm mặt trước có thể thấy ở mặt đứng, dày KHÔNG GHI", "CÓ", "p13 E3; MB p10", "HÌNH", [PG(13, "p13"), PG(10, "p10")])
q("V02", "1", "Kim loại", "Khung hộp quanh chậu (mặt cắt E3); BoQ: khung thép mạ kẽm + bracket", "CÓ", "p13 E3; BoQ", "HÌNH", [PG(13, "p13")])
q("V02", "1", "Khung xương", "Tiết diện thép KHÔNG GHI", "KHÔNG THẤY", "p13 E3", "HÌNH", [PG(13, "p13")])
q("V02", "1", "Phần ngoài phạm vi", "Chậu LA1, vòi FC1, xi-phông; bệ đá SS02 350×120 khu tắm đứng", "CÓ", "MB p74", "HÌNH", [PG(74, "p74")])
# V03
q("V03", "1", "Lắp đặt", "Treo tường trên bàn, đáy cách mặt bàn 350, cao 1100 (đỉnh sát trần thạch cao FCL 2300 trừ 100)", "CÓ", "p13 E3; p12 E2", "HÌNH", [PG(13, "p13"), PG(12, "p12")])
q("V03", "1", "Kích thước BV", "Cao 1100; rộng ≈ 900 (không ghi số); sâu ≈120–150 (không ghi số)", "CÓ", "p13 E3, E2 bên", "HÌNH", [PG(13, "p13")])
q("V03", "1", "Kích thước BoQ/ML", "850×1150×180", "MÂU THUẪN", "ML 003 p1; BoQ!16", "SPEC", [FL(ML3, "ML 003", 1), BQ("B5.16 SD", 16)])
q("V03", "1", "Kệ/đợt", "Hộp có 3 ngăn kệ trong (mặt đứng bên)", "CÓ", "p12 E2 (bên)", "HÌNH", [PG(12, "p12 E2")])
q("V03", "1", "Có cánh", "Mặt đứng 'GƯƠNG' có đường chéo (gương); số cánh KHÔNG GHI", "KHÔNG THẤY", "p13 E3", "HÌNH", [PG(13, "p13")])
q("V03", "1", "Gương", "Glasso customize mirror, mài cạnh, khung bo", "CÓ", "ML 003 p1", "SPEC", [FL(ML3, "ML 003", 1)])
q("V03", "1", "Hoàn thiện hộp/khung", "KHÔNG GHI mã", "KHÔNG THẤY", "ML 003; BV", "SPEC", [FL(ML3, "ML 003", 1)])
q("V03", "1", "Spec gương", "Không có mục Spec gương trong 02.SPECIFICATION", "KHÔNG CÓ", "Mục lục 00 00 01", "SPEC", [FL("02.SPECIFICATION\\00 00 01 - MỤC LỤC-TABLE OF CONTENT.pdf", "Mục lục", 1)])
q("V01", "1", "Thùng fix", "Có: hộp chậu 900×550×~150 nằm trên thân tủ 2 cánh (mặt cắt E3: đáy hộp ~690 AFFL, đo theo tỉ lệ, BV không ghi số); user: 'có 1 thùng fix bên cạnh thùng 2 cánh'", "CÓ", "p13 E3 (mặt cắt); user 26/10", "HÌNH", [PG(13, "p13 E3")])
q("V01", "1", "Cao thân tủ (thùng 2 cánh)", "≈440 theo mặt cắt (đáy ≈200 AFFL → đỉnh ≈655), khớp BoQ 440", "CÓ", "p13 E3 (mặt cắt, đo tỉ lệ)", "HÌNH", [PG(13, "p13 E3")])
q("V02", "1", "Cemboard", "Nóc cemboard đã có sẵn trong module tủ của Template#7; bổ sung viền mặt trước + 2 hông cemboard cao 100 (giả định, BV không ghi); xương đỡ đá = Kim loại gia công ngoài MTGC (Template hàng 771-774)", "CÓ", "user 26/10", "USER", [])
q("V03", "1", "Kích thước chốt", "900 × 180 × 1100 (user chốt: theo BV; sâu 180 theo BoQ/ML)", "CÓ", "user 26/10", "USER", [])
q("V03", "1", "Có cánh (user)", "2 cánh gương", "CÓ", "user 26/10", "USER", [])
q("V01", "1", "Phạm vi (user)", "Không bao gồm: tay nắm; bệ đá SS02 350×120 khu tắm đứng", "CÓ", "user 26/10", "USER", [])
GEM = "Gemini 3.8 flash low đọc 20 crop (13 mặt đứng + 7 mặt bằng): "
q("V01", "1", "Kiểm chéo Gemini", GEM + "vanity rộng 900 (11/11 tờ có số), sâu 550 (540 ở p22), mặt bàn 850 (p221 ghi 870), 2 cánh (6 tờ nêu), 0 ngăn kéo; Gemini báo 'thùng fix cạnh vanity' ở p10/p74 là hộp X khu tắm 900×440 (đã loại), không phải thùng fix của vanity", "CÓ", "tờ p10,12,13,22,25,58,73,74,77,93,150,198,201,217,221,237,249,253,266,269", "GEMINI", [])
q("V03", "1", "Kiểm chéo Gemini", GEM + "gương rộng 900 × cao 1100, đáy cách mặt bàn 350, đỉnh cách trần 100 (12/12 tờ); 2 cánh gương (p12), sâu không ghi", "CÓ", "p12,25,73,77,93,201,217,249", "GEMINI", [])
q("V01", "1", "Biến thể chiều cao mặt bàn", "p221 (3BR C1.1A BT01) ghi 870 thay vì 850", "MÂU THUẪN", "p221 E3", "HÌNH", [PG(221, "p221")])
# chung
for ma in ("V01", "V02", "V03"):
    q(ma, "1", "Khu", "Phòng tắm căn hộ BT01/BT02 (Studio ST1.1a, 1BR A1.1a/A1.3a/A2.1a/A3.1a, 2BR, 3BR); Podium B1/B2 và Tower B1–B3", "CÓ", "MB p10…p283", "HÌNH", [PG(10, "p10")])
    q(ma, "1", "Bộ tờ đã đọc", f"MASS UNITS: {NPS}/{NPT} MB WC + {NES}/{NET} mặt đứng WC xem hình; còn lại đọc text (trùng tiêu đề). Lift lobby 44 tờ: không có vanity", "CÓ", "PAGES", "HÌNH", [])
    q(ma, "1", "Tờ chưa đọc", f"{NPT - NPS} MB + {NET - NES} mặt đứng chỉ đọc text, chưa xem hình (cùng tiêu đề với tờ đã xem); các tờ trần/phối cảnh chưa xem", "KHÔNG THẤY", "PAGES", "HÌNH", [])
    q(ma, "1", "Revision tờ", "Rev A 06/10/2026 (khung tên); Drawing List 1st issue 22/09 – 06/10/2026", "CÓ", "p10 khung tên", "HÌNH", [PG(10, "p10")])
for i, x in enumerate(Q, 2):
    ws.append(list(x[:7]) + [RUN, MODEL, "—"]); setlinks(ws, i, 6, x[7] or [], text=x[5])
    ws.cell(i, 4).alignment = Alignment(wrap_text=True, vertical="top")
ws.auto_filter.ref = ws.dimensions

# ---------------- _LINKS ----------------
ws = sheet("_LINKS", ["Key", "Sheet", "Ô", "Nhãn", "File (tương đối)", "Trang", "Ô đích", "Mở (dự phòng)", "Độ tin cậy"], [22, 14, 7, 40, 70, 6, 30, 10, 9])
for j, (s, c, lab, rel, pg, sub) in enumerate(REG, 2):
    ws.append([f"{s}!{c}", s, c, lab, REL + rel, pg, sub, "", 1.0])
    x = ws.cell(j, 8, "Mở"); x.hyperlink = REL + rel + (("#" + sub) if sub else ""); x.font = LK

# ---------------- DASHBOARD ----------------
ws = wb.create_sheet("DASHBOARD", 0)
ws["A1"] = "NORTON – Skybridge Plot B · VANITY (tủ lavabo + bàn đá + gương)"; ws["A1"].font = Font(bold=True, size=14)
info = [("Hạn trả giá", "28/10/2026"), ("Đợt", "261024_VAnity"), ("File nhập", len(files)), ("Tờ BV MASS UNITS", len(pages)),
        ("Tờ MB WC có VC1", sum(1 for x in pages if x["kind"].startswith("MB nội thất"))), ("Tờ mặt đứng WC", sum(1 for x in pages if x["kind"].startswith("Mặt đứng"))),
        ("Dòng BoQ vanity", len(ZONE) * 3), ("Hạng mục", "3 (V01 tủ lavabo, V02 bàn đá + khung, V03 gương MCB)"), ("Tổng KL", f"{tot['V01']} bộ mỗi hạng mục"),
        ("Issue", f"{len(I)}"), ("RFI nháp", len(R)), ("Dòng QUAN_SAT", len(Q)),
        ("Link", "Cần add-in SRC_LINKS: bấm → mở PDF đúng trang (Edge) / xlsx đúng ô; ô có ▾ → chọn link")]
for i, (a, b) in enumerate(info, 3): ws.cell(i, 1, a).font = Font(bold=True); ws.cell(i, 2, b)
for i, s in enumerate(wb.sheetnames[1:], 3):
    c = ws.cell(i, 4, s); c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{s}'!A1", display=s); c.font = LK
ws.column_dimensions["A"].width = 22; ws.column_dimensions["B"].width = 50; ws.column_dimensions["D"].width = 16
wb.move_sheet("BOQ_TONGHOP", -(wb.sheetnames.index("BOQ_TONGHOP") - 1))
wb.save(OUT)
print(wb.sheetnames, "links:", len(REG), "QUAN_SAT:", len(Q), "pages:", len(pages), "files:", len(files), "tot:", tot)
