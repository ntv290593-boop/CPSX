"""Kiểm AI_INPUT v2 trước khi ghi vào Excel (máy kiểm từng dòng, user chỉ xem NGOAI_LE).

Đầu vào: AI_INPUT.xlsx (tblAiInput, tblAiModule[, tblNgoaiEdits]) + file CPSX của dự án (thư viện mẫu, Material, BoQ).
Ra: KIEM_*.xlsx — TONG_QUAN (đạt/lỗi theo phép kiểm, tổng theo nhóm) + NGOAI_LE (danh sách phải xem).

Phép kiểm (BUOC-TIEP-AI-INPUT-TAKEOFF §3.4):
  MAU      mẫu có trong thư viện của file này, không thuộc nhóm cấm, block mẫu không lỗi (thiếu END)
  OFF      Off nằm trong block mẫu, nhãn A tại Off = Module
  O_NHAP   ô ghi là ô nhập của mẫu (không đè công thức, không thuộc phần nhân công R0150)
  MA_VT    mã vật tư ở cột K (khi mẫu dùng mã Material) có trong Material và có giá (R0074)
  DVT      "Giá theo" (M header) khớp ĐVT BoQ
  SL       Σ SL các block = Σ KL các dòng BoQ chúng trỏ tới (theo nhóm block ↔ dòng BoQ)
  NGUON    mỗi dòng có Nguon + Trang + DoTinCay; THAP / SUY_LUAN → đưa vào danh sách xem
  CUA      block cửa có C header (dày tường) kèm Nguon + Trang (R0166)
  GIATRI   ô GiaTri lưu công thức xlsx thật (Power Query nạp rỗng) → phải lưu text
  NGOAI    ô ngoài vùng ô nhập (đè số, sửa công thức, thêm/bớt dòng) — không ghi được bằng L.fromAiInput

Dùng: py -I check_ai_input.py <AI_INPUT.xlsx> <file CPSX> <KIEM ra.xlsx> [RULES_MASTER.xlsx]
"""
import os, sys, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cpsx_blocks import Book, LIB, CT, norm, is_f, NHAN_CONG
from module_catalog import rule_ids, find_id
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

CAM = {"Door#2": "bản cũ, dùng Door#1 (R0100)", "Template#22": "RAPBOM không dùng chiết tính (R0098)"}
DVT = {"md": {"m", "md", "lm"}, "m2": {"m2", "m²", "sqm"}, None: {"nr", "no", "bộ", "bo", "cái", "cai", "set", "pcs", "ea"}}


def table(wb, name):
    for ws in wb.worksheets:
        if name in ws.tables:
            rows = list(ws[ws.tables[name].ref])
            head = [c.value for c in rows[0]]
            return [dict(zip(head, [c.value for c in r])) for r in rows[1:] if any(c.value is not None for c in r)]
    return []


def main():
    ai_path, cpsx, out = sys.argv[1], sys.argv[2], sys.argv[3]
    ids = rule_ids(sys.argv[4]) if len(sys.argv) > 4 else {}
    wa = openpyxl.load_workbook(ai_path)
    AI, MOD, NG = table(wa, "tblAiInput"), table(wa, "tblAiModule"), table(wa, "tblNgoaiEdits")

    book = Book(cpsx, LIB + ["Material", CT])
    lib = {(b["sheet"], norm(b["name"])): b for s in LIB for b in book.blocks(s)}
    cts = book.blocks(CT)
    mat = {}
    for r in book.V["Material"]:
        if r[0] not in (None, ""):
            mat[str(r[0]).strip().upper()] = r[5]
    loi = []          # (phép kiểm, BlockId, mức, nội dung)

    def bad(chk, bid, msg, muc="LOI"):
        loi.append((chk, bid, muc, msg))

    # GIATRI: ô GiaTri bắt đầu bằng '=' phải lưu dạng TEXT. Nếu xlsx lưu công thức thật (có {boq}, {h+n}) thì Power Query nạp rỗng
    # và L.fromAiInput bỏ qua ô (sự cố SOFITEL batch1: E header không được ghi)
    for ws_ in wa.worksheets:
        if "tblAiModule" in ws_.tables:
            rows_ = list(ws_[ws_.tables["tblAiModule"].ref])
            head_ = [c.value for c in rows_[0]]
            ig, ib, ic = head_.index("GiaTri"), head_.index("BlockId"), head_.index("Cot")
            for r_ in rows_[1:]:
                if r_[ig].data_type == "f":
                    bad("GIATRI", r_[ib].value, f"{r_[ic].value}: GiaTri là công thức xlsx ({str(r_[ig].value)[:40]}) — phải ghi dạng text, nếu không Power Query nạp rỗng")

    # ---- nguồn của từng BlockId: block thư viện, hoặc block Chiet tinh có sẵn (làm tiếp file: MOI chép từ nó / SUA nó)
    src = {}
    for a in AI:
        bid, mt = a["BlockId"], a.get("MauTen") or ""
        op = str(a.get("ThaoTac") or "MOI").upper()
        sn, _, nm = mt.partition("!")
        if sn == CT:
            hit = [b for b in cts if (nm[1:].isdigit() and b["hdr"] == int(nm[1:])) or norm(b["name"]) == norm(nm)]
            if len(hit) != 1:
                bad("MAU", bid, f"'{mt}': {len(hit)} block khớp" + (" → ghi Chiet tinh!#<hàng>" if hit else ""))
                continue
            src[bid] = hit[0]
            if op == "MOI" and not any(m["BlockId"] == bid and int(m["Off"]) == 0 and str(m["Cot"]).upper() == "E" for m in MOD):
                bad("SL", bid, f"Chép từ {mt}: phải khai E header, không kế thừa E nguồn ({book.cell(CT, hit[0]['hdr'], 'E')})")
            continue
        if op == "SUA":
            bad("MAU", bid, "SUA chỉ áp cho block Chiet tinh (MauTen = Chiet tinh!<mã>)")
            continue
        b = lib.get((sn, norm(nm)))
        if not b:
            bad("MAU", bid, f"Mẫu '{mt}' không có trong thư viện của {os.path.basename(cpsx)}")
            continue
        src[bid] = b
        mid = find_id(ids, sn, nm)
        mid = mid["id"] if mid else None
        if mid in CAM:
            bad("MAU", bid, f"{mid} {CAM[mid]}")
        if mid == "Door#8":
            bad("MAU", bid, "Door#8 chỉ dùng cho cửa gỗ tự nhiên nguyên khối (R0106)", "XEM")
        n_gg = sum(1 for r in range(b["hdr"], b["end"] + 1) if norm(book.cell(sn, r, "A", "V")) == "giá gốc nhà máy")
        if n_gg > 1:
            bad("MAU", bid, f"Block mẫu {sn}!{b['hdr']} thiếu END (gộp 2 mẫu) — chép sẽ dài gấp đôi")

    # ---- từng ô nhập
    hdr_c = collections.defaultdict(dict)
    for m in MOD:
        bid, b = m["BlockId"], src.get(m["BlockId"])
        if not b:
            continue
        off, col = int(m["Off"]), str(m["Cot"]).upper()
        if off > b["end"] - b["hdr"]:
            bad("OFF", bid, f"Off {off} vượt block mẫu ({b['end'] - b['hdr'] + 1} dòng)")
            continue
        r = b["hdr"] + off
        lab = book.cell(b["sheet"], r, "A", "V")
        if off == 0:
            hdr_c[bid][col] = m
        elif norm(lab) != norm(m["Module"]):
            bad("OFF", bid, f"Off {off}: nhãn mẫu '{lab}' ≠ Module '{m['Module']}'")
        f0 = book.cell(b["sheet"], r, col)
        f0 = getattr(f0, "text", f0)
        labor = any(norm(book.cell(b["sheet"], x, "A", "V")).startswith(NHAN_CONG) for x in range(b["hdr"], r + 1))
        v = m["GiaTri"]
        if is_f(f0) and not is_f(v):
            bad("O_NHAP", bid, f"{m['Module'] or 'hdr'}!{col} (Off {off}) là công thức của nguồn, ghi số = dán số chết", "LOI")
        elif is_f(f0) and off != 0:
            bad("O_NHAP", bid, f"{m['Module']}!{col} (Off {off}): thay công thức của nguồn bằng {v}", "XEM")
        if labor:
            bad("O_NHAP", bid, f"{m['Module']}!{col}: phần nhân công giữ như mẫu (R0150)")
        k0 = book.cell(b["sheet"], r, "K")
        if col == "K" and isinstance(k0, str) and k0.strip().upper() in mat and isinstance(v, str) and not is_f(v):
            p = mat.get(v.strip().upper())
            where = f"{m['Module'] or '(không nhãn)'} Off {off}"
            if p is None:
                bad("MA_VT", bid, f"{where}: mã '{v}' không có trong Material")
            elif not p:
                ven = v[:2].isdigit() or v[:1].isdigit()
                bad("MA_VT", bid, f"{where}: mã '{v}' đơn giá = 0 trong Material"
                    + (" → đổi dày gần nhất lớn hơn có giá (R0074)" if ven else " (R0073)"))
        if not (m.get("Nguon") and m.get("Trang") and m.get("DoTinCay")):
            bad("NGUON", bid, f"{m['Module']}!{col}: thiếu Nguon/Trang/DoTinCay")
        elif str(m["DoTinCay"]).upper() == "THAP" or str(m["Nguon"]).upper() == "SUY_LUAN":
            bad("NGUON", bid, f"{m['Module']}!{col} = {v} ({m['Nguon']}, {m['DoTinCay']}, {m['Trang']})", "XEM")

    # ---- BoQ: ĐVT, SL (tên sheet BoQ lấy từ link cột A của Chiet tinh)
    boq_name = None
    wbq = openpyxl.load_workbook(cpsx, read_only=True, data_only=True)
    wfq = openpyxl.load_workbook(cpsx, read_only=True)
    ct = wfq["Chiet tinh"] if "Chiet tinh" in wfq.sheetnames else None
    if ct is not None:
        for r in ct.iter_rows(min_row=9, max_row=400, max_col=1, values_only=True):
            if is_f(r[0]) and "!" in r[0]:
                boq_name = r[0].split("!")[0].lstrip("=+").strip("'")
                break
    BQ = {}
    if boq_name:
        for i, r in enumerate(wbq[boq_name].iter_rows(min_col=1, max_col=37, values_only=True), start=1):
            if r[0] not in (None, ""):
                BQ[i] = {"ma": r[0], "nhom": r[1], "ten": r[2], "dvt": r[6], "kl": r[32], "gia": r[33], "tien": r[34]}
    # nhóm block ↔ dòng BoQ (union-find)
    parent = {}

    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x
    rows_of = {}
    for a in AI:
        rows = [int(x) for x in str(a.get("DongBoQ") or "").replace(";", ",").split(",") if x.strip().isdigit()]
        rows_of[a["BlockId"]] = rows
        if not rows:
            bad("SL", a["BlockId"], "Không có DongBoQ")
        for r in rows:
            parent[find(("B", a["BlockId"]))] = find(("R", r))
        b = src.get(a["BlockId"])
        # GiaTheo trống = đơn giá BoQ link H, E = KL BoQ (cách L.newItem làm) → hợp mọi ĐVT;
        # md/m2 = đơn giá BoQ link N, E quy đổi → ĐVT BoQ phải là m / m2. Ô M header của mẫu chỉ là nhãn.
        gt = a.get("GiaTheo")
        a["_gt"] = gt = gt if gt in ("md", "m2") else None
        if gt and rows and BQ:
            for r in rows:
                u = norm(BQ.get(r, {}).get("dvt")).replace(".", "")
                if u and u not in DVT[gt]:
                    bad("DVT", a["BlockId"], f"Giá theo '{gt}' ≠ ĐVT BoQ dòng {r} '{BQ[r]['dvt']}'")
        if b and str(book.cell(b["sheet"], b["hdr"], "G", "V") or "") in ("DR", "DRF"):   # block cửa (thư viện hoặc Chiet tinh)
            c = hdr_c[a["BlockId"]].get("C")
            if not c or not (c.get("Nguon") and c.get("Trang")):
                bad("CUA", a["BlockId"], "Thiếu C header (dày tường theo mặt bằng) kèm nguồn + trang (R0166)")
    grp = collections.defaultdict(lambda: {"blocks": [], "rows": set()})
    for a in AI:
        g = find(("B", a["BlockId"]))
        grp[g]["blocks"].append(a)
        grp[g]["rows"].update(rows_of[a["BlockId"]])
    def qty(a):
        """SL block quy về ĐVT BoQ (R0075): md = B·E/1000, m2 = B·D·E/10^6, trống = E."""
        e, w, h = (float(a.get(k) or 0) for k in ("SL", "W", "H"))
        return w * e / 1000 if a["_gt"] == "md" else w * h * e / 1e6 if a["_gt"] == "m2" else e
    for g in grp.values():
        if not g["rows"] or not BQ:
            continue
        kl = sum(float(BQ.get(r, {}).get("kl") or 0) for r in g["rows"])
        tp = [a for a in g["blocks"] if a.get("QuanHe") == "THANH_PHAN"]
        tach = [a for a in g["blocks"] if a.get("QuanHe") != "THANH_PHAN"]
        # thành phần: mỗi block mang đủ KL các dòng của chính nó
        checks = [(a["BlockId"], qty(a), rows_of[a["BlockId"]]) for a in tp]
        if tach:
            rr = set(r for a in tach for r in rows_of[a["BlockId"]])
            checks.append((", ".join(a["BlockId"] for a in tach), sum(qty(a) for a in tach), rr))
        for ids_, sl, rr in checks:
            kl = sum(float(BQ.get(r, {}).get("kl") or 0) for r in rr)
            if abs(sl - kl) > 1e-3 * max(1, kl):
                bad("SL", ids_, f"SL block (theo ĐVT BoQ) = {sl:g} ≠ Σ KL BoQ dòng {sorted(rr)} = {kl:g}")

    # ---- ô ngoài vùng ô nhập (từ trích ngược / kiểm khứ hồi)
    for bid, c in collections.Counter((n["BlockId"], n["Loai"]) for n in NG).items():
        bad("NGOAI", bid[0], f"{c} ô/dòng {bid[1]}", "XEM")

    write(out, AI, BQ, loi, rows_of)


def write(path, AI, BQ, loi, rows_of):
    sys.stdout.reconfigure(encoding="utf-8")
    CHK = ["MAU", "OFF", "O_NHAP", "MA_VT", "DVT", "SL", "NGUON", "CUA", "GIATRI", "NGOAI"]
    wb = Workbook()
    ws = wb.active
    ws.title = "TONG_QUAN"
    ws.append(["Phép kiểm", "Block lỗi", "Dòng lỗi", "Block cần xem", "Dòng cần xem"])
    print(f"{'Phép kiểm':8} {'block lỗi':>9} {'block xem':>9}")
    for c in CHK:
        e = [x for x in loi if x[0] == c and x[2] == "LOI"]
        w = [x for x in loi if x[0] == c and x[2] == "XEM"]
        row = [c, len({x[1] for x in e}), len(e), len({x[1] for x in w}), len(w)]
        ws.append(row)
        print(f"{c:8} {row[1]:>9} {row[3]:>9}")
    ws.append([])
    ws.append(["Nhóm BoQ", "Số block", "Σ thành tiền BoQ", "Block có lỗi"])
    by = collections.defaultdict(lambda: [0, 0.0, 0])
    errb = {x[1] for x in loi if x[2] == "LOI"}
    seen = set()
    for a in AI:
        rows = rows_of[a["BlockId"]]
        g = (BQ.get(rows[0], {}).get("nhom") if rows else None) or "?"
        by[g][0] += 1
        by[g][2] += a["BlockId"] in errb
        for r in rows:
            if r not in seen:
                seen.add(r)
                by[g][1] += float(BQ.get(r, {}).get("tien") or 0)
    for g, v in by.items():
        ws.append([g, v[0], round(v[1]), v[2]])
    for c in ws[1] + ws[ws.max_row - len(by)]:
        c.font = Font(bold=True)
    w2 = wb.create_sheet("NGOAI_LE")
    w2.append(["Phép kiểm", "BlockId", "Mức", "Nội dung"])
    red, yel = PatternFill("solid", fgColor="FFC7CE"), PatternFill("solid", fgColor="FFEB9C")
    for x in sorted(loi, key=lambda x: (x[2] != "LOI", CHK.index(x[0]), str(x[1]))):
        w2.append(list(x))
        for c in w2[w2.max_row]:
            c.fill = red if x[2] == "LOI" else yel
    w2.freeze_panes = "A2"
    for col, wd in zip("ABCD", (10, 18, 6, 110)):
        w2.column_dimensions[col].width = wd
    wb.save(path)
    print("→", path)


if __name__ == "__main__":
    main()
