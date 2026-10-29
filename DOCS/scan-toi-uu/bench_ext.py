# -*- coding: utf-8 -*-
"""Đo pd.py ungvien trên SOFITEL / NORTON (dự án đã đóng — chỉ làm bộ đo). Text lấy từ _cache\\raw (do sofitel_cache.py / norton_cache.py).
Dựng project giả trong thư mục tạm (_text/F###.json + FILES/PAGES/BOQ_TONGHOP) rồi gọi thẳng pd.build_index / pd.ungvien_all.
Đáp án: link trong PROJECT_DATA cũ (xem norton_rank.py, sofitel_rank.py).
py -X utf8 bench_ext.py [sofitel|d01|vanity ...]"""
import sys, os, re, json, shutil, tempfile, collections, openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "SKILL", "01-scan", "tool"))
import pd

RAW = os.path.join(HERE, "_cache", "raw")
SCAN_S = r"D:\COSTING EXCEL\_ACTIVE\SOFITEL\01_scan"
SCAN_N = r"D:\COSTING EXCEL\_ACTIVE\NORTON\01_scan"


def load_raw(name):
    d = os.path.join(RAW, name)
    meta = json.load(open(os.path.join(d, "_files.json"), encoding="utf-8"))
    fid = {}
    for i, m in enumerate(meta, 1): fid[m["ID"]] = f"F{i:03d}"
    texts = {fid[m["ID"]]: json.load(open(os.path.join(d, m["ID"] + ".json"), encoding="utf-8")) for m in meta}
    return meta, fid, texts


def make_project(meta, fid, texts, boq_rows, khu_rows=()):
    """-> (P, sheets) với data trong thư mục tạm."""
    tmp = tempfile.mkdtemp(); data = os.path.join(tmp, "01_scan", "_data", "X"); os.makedirs(os.path.join(data, "_text"))
    FILES, PAGES = [], []
    for m in meta:
        f = fid[m["ID"]]
        FILES.append(("kiemke", {"ID": f, "Đường dẫn": m["rel"], "Tên file": os.path.basename(m["rel"])}))
        json.dump(texts[f]["texts"], open(os.path.join(data, "_text", f + ".json"), "w", encoding="utf-8"), ensure_ascii=False)
        for i, t in enumerate(texts[f]["titles"], 1): PAGES.append(("kiemke", {"File ID": f, "Trang": i, "Tên BV": t}))
    sheets = {"FILES": FILES, "PAGES": PAGES,
              "BOQ_TONGHOP": [("boq", r) for r in boq_rows], "BOQ_MAP": [("boq", r) for r in khu_rows]}
    return pd.paths(tmp, "X"), sheets, tmp


def texts_fix(texts):  # raw: {fid: {"texts": [...], "titles": [...]}}
    return texts


def run(title, P, sheets, g, tmp, keep=lambda p: p, bq=None):
    if bq: json.dump(bq, open(os.path.join(P["data"], "_timkiem.boq.json"), "w", encoding="utf-8"), ensure_ascii=False)
    ix = pd.build_index(P, sheets, write=False)
    allm, _ = pd.ungvien_all(ix, sheets)
    g = {m: {keep(ix["dup"].get(p, p)) for p in s} for m, s in g.items() if m in allm}
    flat = {m: list(dict.fromkeys(x[1] for x in allm[m]["res"])) for m in g}
    tot = sum(len(v) for v in g.values())
    def hit(k): return sum(1 for m in g for p in g[m] if p in set(flat[m][:k]))
    a12 = sum(1 for m in g if g[m] & set(flat[m][:12]))
    def shown(m):
        d, out = allm[m], []
        for loai, _, n in pd.NHOM_RA:
            if loai == "chi_tiet": hh = [x for x in d["res"] if ix["loai"][x[1]] in ("chi_tiet", "mat_dung", "khac", "legend")]; n = 8
            else: hh = [x for x in d["res"] if ix["loai"][x[1]] == loai and x[2]]
            out += [x[1] for x in hh[:n]]
        return out
    sh = sum(1 for m in g for p in g[m] if p in set(shown(m)))
    print(f"{title:22s} {ix['N']:5d} trang, trùng {len(ix['dup'])}, file tổng {sorted(ix['tong'])} | đáp án {len(g)} mã/{tot} trang | "
          f"@5 {hit(5)/tot:4.0%} @12 {hit(12)/tot:4.0%} @40 {hit(40)/tot:4.0%} hết {hit(10**6)/tot:4.0%} | đầu ra {sh/tot:4.0%} | mã ≥1 đúng @12 {a12}/{len(g)}")
    shutil.rmtree(tmp, ignore_errors=True)
    return ix, allm, g, flat


def sofitel():
    meta, fid, texts = load_raw("sofitel"); byrel = {m["rel"].lower(): fid[m["ID"]] for m in meta}
    wb = openpyxl.load_workbook(os.path.join(SCAN_S, "PROJECT_DATA_SOFITEL_v8.xlsx"), read_only=True)
    rows = list(wb["BOQ_TONGHOP"].iter_rows(min_row=2, values_only=True))
    # Mã SOFITEL (F001…) có thể trùng giữa các khu / là ô lưới: dùng Mã nguyên, nếu trùng thì thêm #dòng
    boq, key, seen = [], {}, collections.Counter()
    for i, r in enumerate(rows):
        if not r[2]: continue
        seen[str(r[2])] += 1; k = str(r[2]) if seen[str(r[2])] == 1 else f"{r[2]}#{i + 2}"
        key[i + 2] = k; mo = (r[4] or "")
        boq.append({"Mã": k, "Hạng mục": mo[:120], "Mô tả BoQ": mo})
    g = collections.defaultdict(set)
    for r in wb["_LINKS"].iter_rows(min_row=2, values_only=True):
        if r[1] != "BOQ_TONGHOP" or not re.fullmatch(r"P\d+", str(r[2])) or not r[5]: continue
        rel = re.sub(r"^\.\.\\01\. INPUT\\", "", str(r[4])).lower()
        row = int(r[2][1:])
        if rel in byrel and row in key: g[key[row]].add((byrel[rel], int(r[5])))
    P, sheets, tmp = make_project(meta, fid, texts, boq)
    print(f"SOFITEL: {len(boq)} dòng BoQ ({sum(1 for k in key.values() if '#' in k)} mã trùng)")
    return run("SOFITEL", P, sheets, g, tmp)


def d01():
    meta, fid, texts = load_raw("norton_d01"); C = {m["ID"]: m["rel"] for m in meta}; byname = {os.path.basename(v).lower(): fid[k] for k, v in C.items()}
    sheetdir = {"BLK01": "BLK 01", "BLK02": "BLK 02", "BLK03": "BLK 03", "GEN-BM": "GENERAL, BASEMENT", "TMDV01": "TMDV 01", "TMDVB": "TMDV B"}
    def page(rel, trang):
        b = os.path.basename(rel).lower()
        if b in byname: return (byname[b], int(trang))
        m = re.search(r"_sheets\\([^\\]+)\\p(\d+)_", rel)
        if m: return (byname[f"sb_tender_plot b_{sheetdir[m.group(1)]}_2026 10 08.pdf".lower()], int(m.group(2)))
    wb = openpyxl.load_workbook(os.path.join(SCAN_N, "PROJECT_DATA_v8.xlsx"), read_only=True)
    rows = list(wb["BOQ_TONGHOP"].iter_rows(min_row=2, values_only=True))
    boq = {i + 2: {"Mã": str(r[2]), "Nhóm": r[1], "Hạng mục": f"{r[1]} {r[3] or ''}", "Mô tả BoQ": f"{r[4] or ''} {r[5] or ''}"} for i, r in enumerate(rows) if r[2]}
    khu = [{"Mã": r[0], "Khu": r[1]} for r in wb["BOQ_MAP"].iter_rows(min_row=2, values_only=True) if r[0]]
    g = collections.defaultdict(set)
    for r in wb["_LINKS"].iter_rows(min_row=2, values_only=True):
        if r[1] != "BOQ_TONGHOP" or not re.fullmatch(r"[XY]\d+", str(r[2])): continue
        row = int(r[2][1:])
        try: p = page(str(r[4]), r[5] or 0)
        except KeyError: p = None
        if p and p[1] and boq.get(row, {}).get("Nhóm") == "Cửa": g[boq[row]["Mã"]].add(p)
    cua = [r for r in boq.values() if r["Nhóm"] == "Cửa"]
    P, sheets, tmp = make_project(meta, fid, texts, cua, khu)
    print(f"NORTON D01: {len(cua)} mã cửa")
    return run("NORTON D01 cửa", P, sheets, g, tmp)


def vanity():
    meta, fid, texts = load_raw("norton_vanity"); fb = {os.path.basename(m["rel"]).lower(): fid[m["ID"]] for m in meta}
    wb = openpyxl.load_workbook(os.path.join(SCAN_N, "PROJECT_DATA_VANITY_v1.xlsx"), read_only=True)
    rows = list(wb["BOQ_TONGHOP"].iter_rows(min_row=2, values_only=True))
    boq = {i + 2: {"Mã": str(r[1]), "Hạng mục": r[2] or "", "Mô tả BoQ": r[4] or ""} for i, r in enumerate(rows) if r[1]}
    maprow = {i + 2: r[0] for i, r in enumerate(wb["BOQ_MAP"].iter_rows(min_row=2, values_only=True))}
    g = collections.defaultdict(set)
    for r in wb["_LINKS"].iter_rows(min_row=2, values_only=True):
        if not ((r[1] == "BOQ_TONGHOP" and str(r[2])[0] == "Q") or (r[1] == "BOQ_MAP" and str(r[2])[0] in "GH")) or not r[5]: continue
        f = fb.get(os.path.basename(str(r[4])).lower())
        ma = boq[int(r[2][1:])]["Mã"] if r[1] == "BOQ_TONGHOP" else maprow.get(int(r[2][1:]))
        if f and ma: g[ma].add((f, int(r[5])))
    P, sheets, tmp = make_project(meta, fid, texts, list(boq.values()))
    print(f"NORTON VANITY: {len(boq)} mã")
    # giả định bước boq có gán mã BV (như prototype H/I): V01 VC1, V02 VC1+SS02, V03 gương
    bq = {"ma_bv": {"V01": ["VC1"], "V02": ["VC1", "SS02"], "V03": ["GUONG", "MIRROR"]}} if "--ma-bv" in sys.argv else None
    out = run("NORTON vanity" + (" +ma_bv" if bq else ""), P, sheets, g, tmp, bq=bq)
    ix, allm = out[0], out[1]  # đáp án cũ chỉ có 1 tờ đại diện/loại căn -> đo thêm: top-12 có bao nhiêu tờ 'chi tiết phòng tắm' (như norton_rank.py)
    bt = lambda p: "phong tam" in pd._fold(ix["titles"].get(p, ""))
    print("   tờ chi tiết phòng tắm trong top-12: " + ", ".join(f"{m} {sum(map(bt, [x[1] for x in allm[m]['res'][:12]]))}/12" for m in allm))
    return out


if __name__ == "__main__":
    for a in ([x for x in sys.argv[1:] if not x.startswith("--")] or ["d01", "vanity", "sofitel"]): {"sofitel": sofitel, "d01": d01, "vanity": vanity}[a]()
