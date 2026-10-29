# -*- coding: utf-8 -*-
"""Test nhanh xếp hạng ứng viên trên SOFITEL. Đáp án = link cột P 'Bản vẽ chi tiết' của BOQ_TONGHOP trong PROJECT_DATA_SOFITEL_v8.
py -X utf8 sofitel_rank.py <thư mục cache của sofitel_cache.py>"""
import sys, os, re, json, math, collections
sys.path.insert(0, r"D:\COSTING EXCEL\SKILL\01-scan\tool")
import pd, openpyxl

CACHE = sys.argv[1]
V8 = r"D:\COSTING EXCEL\_ACTIVE\SOFITEL\01_scan\PROJECT_DATA_SOFITEL_v8.xlsx"
meta = json.load(open(os.path.join(CACHE, "_files.json"), encoding="utf-8"))
byrel = {m["rel"].lower(): m["ID"] for m in meta}
texts, titles = {}, {}
for m in meta:
    d = json.load(open(os.path.join(CACHE, m["ID"] + ".json"), encoding="utf-8"))
    texts[m["ID"]], titles[m["ID"]] = d["texts"], d["titles"]

wb = openpyxl.load_workbook(V8, read_only=True)
rows = list(wb["BOQ_TONGHOP"].iter_rows(min_row=2, values_only=True))
boq = {i + 2: {"Mã": r[2], "Khu": r[3] or "", "Mô tả": r[4] or ""} for i, r in enumerate(rows) if r[2]}
g = collections.defaultdict(set); khong_map = 0
for r in wb["_LINKS"].iter_rows(min_row=2, values_only=True):
    if r[1] != "BOQ_TONGHOP" or not re.fullmatch(r"P\d+", str(r[2])) or not r[5]: continue
    rel = re.sub(r"^\.\.\\01\. INPUT\\", "", str(r[4])).lower()
    if rel not in byrel: khong_map += 1; continue
    row = int(r[2][1:])
    if row in boq: g[row].add((byrel[rel], int(r[5])))

PAGES = [(f, i) for f in texts for i in range(1, len(texts[f]) + 1)]
TOK = re.compile(r"[a-z0-9]+")
STOP = pd.STOP | set("to at nr mm mmh high wide thk refer drawing detail no supply install all as per".split())
toks = lambda s: [w for w in TOK.findall(pd._fold(s)) if len(w) >= 2 and w not in STOP]
body = {p: collections.Counter(toks(texts[p[0]][p[1] - 1])) for p in PAGES}
ttl = {p: collections.Counter(toks(titles[p[0]][p[1] - 1])) for p in PAGES}
N = len(PAGES); avgdl = sum(sum(c.values()) for c in body.values()) / N
df = collections.Counter(w for c in body.values() for w in c)


def bm25(q, c, k1=1.2, b=0.75):
    dl = sum(c.values()); s = 0.0
    for w in q:
        if w in c:
            idf = math.log(1 + (N - df[w] + 0.5) / (df[w] + 0.5))
            s += idf * c[w] * (k1 + 1) / (c[w] + k1 * (1 - b + b * dl / avgdl))
    return s


def ncode(s):
    m = re.fullmatch(r"([a-z]{1,4})[\s._-]*0*(\d{1,3})", pd._fold(s).strip())
    return f"{m.group(1)}{m.group(2)}" if m else None


CODE = re.compile(r"(?<![a-z0-9])([a-z]{1,4})[\s._-]?0*(\d{1,3})(?![0-9])")
codes = {p: collections.Counter(a + b for a, b in CODE.findall(pd._fold(texts[p[0]][p[1] - 1]))) for p in PAGES}
SHEET = re.compile(r"\b[A-Z]{2}-[A-Z]{2}-[A-Z]{2}-\d{3}\b")  # số tờ kiểu ID-LL-TY-701


def current(r):
    """Logic cmd_ungvien: mã regex + ≥2 từ khóa của Hạng mục (ở đây = Mô tả BoQ)."""
    ma, hm = str(r["Mã"]), r["Mô tả"]
    kws = sorted({w for w in re.findall(r"[a-z0-9]{3,}", pd._fold(hm)) if w not in pd.STOP})
    mrx = re.compile(r"(?<![a-z0-9])" + re.escape(pd._fold(ma)) + r"(?![a-z0-9])")
    res = []
    for p in PAGES:
        t = pd._fold(texts[p[0]][p[1] - 1]); mh = len(mrx.findall(t)); kh = [w for w in kws if w in t]
        if mh or len(kh) >= min(2, len(kws)) and kh: res.append((mh * 5 + len(kh), p))
    return [p for s, p in sorted(res, key=lambda x: -x[0])]


def new(r, code_ma=False, code_desc=True, use_bm25=True, title_w=3.0, sheet_ref=True):
    q = set(toks(r["Mô tả"]))
    cs = set()
    if code_ma and ncode(str(r["Mã"])): cs.add(ncode(str(r["Mã"])))
    if code_desc: cs |= {ncode(x) for x in re.findall(r"\b[A-Za-z]{2,3}-\d{2}\b", r["Mô tả"])} - {None}
    refs = set(SHEET.findall(r["Mô tả"])) if sheet_ref else set()
    res = []
    for p in PAGES:
        if use_bm25: sc = bm25(q, body[p]) + title_w * sum(1 for w in q if w in ttl[p])
        else:
            kh = [w for w in q if w in body[p]]; sc = len(kh) if len(kh) >= min(2, len(q)) else 0
        sc += sum(4 * min(codes[p].get(c, 0), 3) for c in cs)
        if refs and any(x in titles[p[0]][p[1] - 1] for x in refs): sc += 20
        if sc > 0: res.append((sc, p))
    return [p for s, p in sorted(res, key=lambda x: -x[0])]


def report(name, fn):
    tot = sum(len(v) for v in g.values()); hits = collections.Counter(); inc = 0; a12 = 0; sizes = []
    for row, gs in g.items():
        rk = fn(boq[row]); sizes.append(len(rk)); s = set(rk)
        inc += len(gs & s)
        for k in (5, 12, 20, 40): hits[k] += len(gs & set(rk[:k]))
        a12 += bool(gs & set(rk[:12]))
    sizes.sort()
    print(f"{name:34s} trong ứng viên {inc/tot:.0%}  @5 {hits[5]/tot:.0%}  @12 {hits[12]/tot:.0%}  @20 {hits[20]/tot:.0%}  @40 {hits[40]/tot:.0%}"
          f" | mã có ≥1 đúng @12: {a12}/{len(g)} | trung vị ứng viên {sizes[len(sizes)//2]}")


if __name__ == "__main__":
    print(f"PDF {len(meta)} file / {N} trang; BoQ {len(boq)} dòng; đáp án {len(g)} mã / {sum(map(len, g.values()))} trang (link không khớp file: {khong_map})")
    print("mã (F0xx) xuất hiện trên BV:", sum(1 for r in boq.values() if any(codes[p].get(ncode(str(r['Mã'])) or '#') for p in PAGES)), "/", len(boq), "(đụng ô lưới/cao độ F1, F2…)")
    report("A hiện tại", current)
    report("C mã chuẩn hóa (Mã F0xx)", lambda r: new(r, code_ma=True, code_desc=False, use_bm25=False, sheet_ref=False))
    report("C' mã trong mô tả (WD-01…)", lambda r: new(r, use_bm25=False, sheet_ref=False))
    report("E BM25 (không tên tờ)", lambda r: new(r, title_w=0, sheet_ref=False))
    report("F BM25 + tên tờ ×3", lambda r: new(r, sheet_ref=False))
    report("K F + số tờ ghi trong mô tả", lambda r: new(r))
    report("L K nhưng không mã trong mô tả", lambda r: new(r, code_desc=False))
