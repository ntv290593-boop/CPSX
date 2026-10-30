# -*- coding: utf-8 -*-
"""Đo N3.2 (neo spec/ML) của `pd.py neo` trên link spec/ML thật — gọi thẳng hàm pd.py, không chép logic.
Dữ liệu: SOFITEL\\01_scan\\PROJECT_DATA_SOFITEL_v8.xlsx (đợt costing cũ; user cho phép 30/10 dùng để ĐO, không lấy số liệu sang tender).
v8 là form cũ, nên dựng merged giả theo form v9:
 A. FINISHES!E (Finishes Schedule PDF)      -> SPEC.File · Trang, Mục = 'Mã + Vật liệu/mô tả'       (1 link/ô)
 B. BOQ_TONGHOP!S (spec FF&E PDF, 4 link/ô)  -> BOQ_TONGHOP.Spec · ML, Thông số riêng = Mô tả (BoQ)
 B1. như B nhưng tách 1 link/ô (để kiểm cả mã trong giá trị, không chỉ nhãn)
 C. BOQ_TONGHOP!O (dòng BoQ Excel)           -> BOQ_TONGHOP.Spec · ML, Thông số riêng = Mã VL/HT (BoQ) (đường Excel)
Mỗi bộ: gốc · dời trang/dòng ±1 · tráo giá trị với dòng khác. Ghi lại toàn bộ LỆCH gốc để xem tay.
Chạy: py -I -X utf8 neo01_n32.py"""
import sys, os, re, json, copy, random, shutil, tempfile, collections
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "..", "SKILL", "01-scan", "tool"))
import pd
import openpyxl, pymupdf
random.seed(5)

ROOT = r"D:\COSTING EXCEL\_ACTIVE\SOFITEL"; SCAN = ROOT + r"\01_scan"
wb = openpyxl.load_workbook(SCAN + r"\PROJECT_DATA_SOFITEL_v8.xlsx", read_only=True)
L = collections.defaultdict(list)
for key, sh, cell, lab, f, pg, o, *_ in wb["_LINKS"].iter_rows(min_row=2, values_only=True):
    if key and f: L[key].append((lab, os.path.normpath(os.path.join(SCAN, f)), pg, o))
TH = {i: r for i, r in enumerate(wb["BOQ_TONGHOP"].iter_rows(min_row=2, values_only=True), 2)}
FI = {i: r for i, r in enumerate(wb["FINISHES"].iter_rows(min_row=2, values_only=True), 2)}

tmp = tempfile.mkdtemp(); P = pd.paths(tmp, "X"); os.makedirs(os.path.join(P["data"], "_text")); P["inp"] = ROOT
fid, FILES, npg = {}, [], {}


def F(f):
    """PDF -> F-id + cache text (giống kiemke); Excel giữ đường dẫn tuyệt đối."""
    if not f.lower().endswith(".pdf"): return f
    if f not in fid:
        k = fid[f] = f"F{len(fid) + 1:03d}"; doc = pymupdf.open(f); npg[k] = doc.page_count
        json.dump([p.get_text() for p in doc], open(os.path.join(P["data"], "_text", k + ".json"), "w", encoding="utf-8"), ensure_ascii=False)
        FILES.append({"ID": k, "Đường dẫn": os.path.relpath(f, ROOT)})
    return fid[f]


def lk(lab, f, pg, o): return [lab, F(f), pg, o] if not f.lower().endswith(".pdf") else [lab, F(f), int(pg)]


A = [{"Mục": f"{r[0]} {r[1] or ''}".strip(), "File · Trang": [lk(*l) for l in L[f"FINISHES!E{i}"]]}
     for i, r in FI.items() if L.get(f"FINISHES!E{i}")]
B = [{"Mã": TH[i][2], "Thông số riêng": TH[i][4] or "", "Spec · ML": [lk(*l) for l in L[f"BOQ_TONGHOP!S{i}"]]}
     for i in TH if L.get(f"BOQ_TONGHOP!S{i}")]
B1 = [{"Mã": r["Mã"], "Thông số riêng": r["Thông số riêng"], "Spec · ML": [l]} for r in B for l in r["Spec · ML"]]
C = [{"Mã": TH[i][2], "Thông số riêng": TH[i][6] or "", "Spec · ML": [lk(*l) for l in L[f"BOQ_TONGHOP!O{i}"]]}
     for i in TH if L.get(f"BOQ_TONGHOP!O{i}")]
BO = {"A": ("SPEC", "File · Trang", "Mục", A), "B": ("BOQ_TONGHOP", "Spec · ML", "Thông số riêng", B),
      "B1": ("BOQ_TONGHOP", "Spec · ML", "Thông số riêng", B1), "C": ("BOQ_TONGHOP", "Spec · ML", "Thông số riêng", C)}


def run(sh, rows):
    return pd.neo_run(P, {"FILES": FILES, "PAGES": [], sh: rows})


def doi(rows, cot):
    """Dời mỗi link ±1 trang (PDF) hoặc ±1 dòng (Excel)."""
    m = copy.deepcopy(rows)
    for r in m:
        for l in r[cot]:
            if len(l) == 3:
                q = l[2] + random.choice((-1, 1))
                if not 1 <= q <= npg[l[1]]: q = l[2] - (q - l[2])
                l[2] = q
            else:
                mm = re.match(r"^(.*!\$?[A-Z]{1,3}\$?)(\d+)$", str(l[3]))
                if mm: l[3] = mm.group(1) + str(int(mm.group(2)) + random.choice((-1, 1)))
    return m


def trao(rows, val):
    """Giá trị của dòng khác (khác giá trị) — giả lập LLM ghi nhầm dòng."""
    m = copy.deepcopy(rows); vs = sorted(set(r[val] for r in rows))
    for r in m:
        o = [v for v in vs if v != r[val]]
        if o: r[val] = random.choice(o)
    return m


for ten, (sh, cot, val, rows) in BO.items():
    r0 = run(sh, rows)
    print(f"\n=== {ten}: {sh}.{cot} ({len(rows)} dòng, {sum(len(r[cot]) for r in rows)} link) ===")
    print("  gốc     :", pd.neo_tong(r0))
    for l in r0["lech"][:12]: print("     ", pd.neo_dong(l))
    if len(r0["lech"]) > 12: print(f"      … còn {len(r0['lech']) - 12} LỆCH")
    co0 = r0["tong"].get("CÓ NEO", 0)
    for tn, m in (("dời ±1", doi(rows, cot)), ("tráo giá trị", trao(rows, val))):
        r1 = run(sh, m)
        print(f"  {tn:9}:", pd.neo_tong(r1), f"→ LỆCH tăng {r1['tong'].get('LỆCH', 0) - r0['tong'].get('LỆCH', 0)} / {co0} CÓ NEO gốc")
        if ten == "B1" and tn == "dời ±1":
            # trang mới vẫn thuộc bộ trang của cùng hạng mục (spec FF&E: 1 hạng mục = trang schedule + 3 trang item sheet liền nhau)
            goc = collections.defaultdict(set)
            for r in rows: goc[r["Mã"]] |= {(l[1], l[2]) for l in r[cot]}
            trong = sum((l[1], l[2]) in goc[r["Mã"]] for r in m for l in r[cot])
            print(f"             trang sau khi dời vẫn là trang của chính hạng mục đó: {trong}/{len(m)} (CÓ NEO đúng, không phải lọt)")
    # mã mà neo_ma tách được (để xem có bỏ sót dạng mã nào)
    ma = collections.Counter(re.sub(r"\d", "9", m) for x in pd.neo_items({sh: rows}, {r["ID"]: os.path.join(ROOT, r["Đường dẫn"]) for r in FILES}, ROOT) for m in x["ma"])
    print("  dạng mã tách được:", ma.most_common(10))
shutil.rmtree(tmp, ignore_errors=True)
