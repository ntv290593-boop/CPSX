# -*- coding: utf-8 -*-
"""Test xếp hạng ứng viên trên NORTON (dự án đã đóng, chỉ làm bộ đo).
Đáp án = link trong PROJECT_DATA cũ:
  d01    : PROJECT_DATA_v8 — BOQ_TONGHOP cột X 'Thống kê BV' + Y 'Chi tiết cửa' của 5 dòng Cửa (mã cửa thật trên BV)
  vanity : PROJECT_DATA_VANITY_v1 — BOQ_TONGHOP Q 'Chi tiết BV' + BOQ_MAP G 'Tờ MB nội thất' / H 'Tờ mặt đứng' (mã V01–V03 tự đặt)
Trang trùng nguyên văn (cùng lớp text) coi là một: đáp án chỉ chọn 1 bản đại diện.
py -X utf8 norton_rank.py <cache d01> <cache vanity>"""
import sys, os, re, json, math, hashlib, collections
sys.path.insert(0, r"D:\COSTING EXCEL\SKILL\01-scan\tool")
import pd, openpyxl

SCAN = r"D:\COSTING EXCEL\_ACTIVE\NORTON\01_scan"
TOK = re.compile(r"[a-z0-9]+")
STOP = pd.STOP | set("to at nr mm mmh high wide thk refer drawing detail no supply install all as per approved overall size with and the of for".split())
toks = lambda s: [w for w in TOK.findall(pd._fold(s)) if len(w) >= 2 and w not in STOP]
CODE = re.compile(r"(?<![a-z0-9])([a-z]{1,4})[\s._-]?0*(\d{1,3})(?![0-9])")


def ncode(s):
    """Mã chuẩn hóa = cụm chữ+số cuối: ID-DH111 / DH-111 / dh111 -> dh111; V01 -> v1."""
    m = re.search(r"([a-z]{1,4})[\s._-]*0*(\d{1,3})$", pd._fold(str(s)).strip())
    return f"{m.group(1)}{m.group(2)}" if m else None


class Corpus:
    def __init__(self, cache):
        self.meta = json.load(open(os.path.join(cache, "_files.json"), encoding="utf-8"))
        self.texts, self.titles = {}, {}
        for m in self.meta:
            d = json.load(open(os.path.join(cache, m["ID"] + ".json"), encoding="utf-8"))
            self.texts[m["ID"]], self.titles[m["ID"]] = d["texts"], d["titles"]
        self.rel = {m["ID"]: m["rel"] for m in self.meta}
        self.PAGES = [(f, i) for f in self.texts for i in range(1, len(self.texts[f]) + 1)]
        T = lambda p: self.texts[p[0]][p[1] - 1]
        # lớp trùng: cùng lớp text (bỏ khoảng trắng); trang rỗng giữ riêng
        self.cls, first = {}, {}
        for p in self.PAGES:
            t = re.sub(r"\s+", " ", T(p)).strip()
            k = hashlib.md5(t.encode()).hexdigest() if len(t) > 30 else p
            self.cls[p] = first.setdefault(k, p)
        self.body = {p: collections.Counter(toks(T(p))) for p in self.PAGES}
        self.ttl = {p: collections.Counter(toks(self.titles[p[0]][p[1] - 1])) for p in self.PAGES}
        self.codes = {p: collections.Counter(a + b for a, b in CODE.findall(pd._fold(T(p)))) for p in self.PAGES}
        self.fold = {p: pd._fold(T(p)) for p in self.PAGES}
        self.N = len(self.PAGES); self.avgdl = sum(sum(c.values()) for c in self.body.values()) / self.N
        self.df = collections.Counter(w for c in self.body.values() for w in c)

    def bm25(self, q, c, k1=1.2, b=0.75):
        dl = sum(c.values()); s = 0.0
        for w in q:
            if w in c:
                idf = math.log(1 + (self.N - self.df[w] + 0.5) / (self.df[w] + 0.5))
                s += idf * c[w] * (k1 + 1) / (c[w] + k1 * (1 - b + b * dl / self.avgdl))
        return s

    def current(self, r):
        """Logic cmd_ungvien: mã regex (bỏ dấu, giữ . -) + ≥2 từ khóa của Hạng mục."""
        kws = sorted({w for w in re.findall(r"[a-z0-9]{3,}", pd._fold(r["Hạng mục"])) if w not in pd.STOP})
        mrx = re.compile(r"(?<![a-z0-9])" + re.escape(pd._fold(r["Mã"])) + r"(?![a-z0-9])")
        res = []
        for p in self.PAGES:
            t = self.fold[p]; mh = len(mrx.findall(t)); kh = [w for w in kws if w in t]
            if mh or len(kh) >= min(2, len(kws)) and kh: res.append((mh * 5 + len(kh), p))
        return [p for s, p in sorted(res, key=lambda x: -x[0])]

    def new(self, r, code=True, bm=True, desc=True, title_w=3.0, area=None, alias=(), qmax=None, skip=None):
        dt = toks(r["Mô tả"])[:qmax] if desc else []
        q = set(toks(r["Hạng mục"])) | set(dt)
        cs = ({ncode(r["Mã"])} - {None} if code else set()) | {ncode(a) for a in alias if ncode(a)}
        ws = [re.compile(r"(?<![a-z0-9])" + pd._fold(a) + r"(?![a-z0-9])") for a in alias if not ncode(a)]
        res = []
        for p in self.PAGES:
            if area and not area(p): continue
            if skip and skip(p): continue
            if bm: sc = self.bm25(q, self.body[p]) + title_w * sum(1 for w in q if w in self.ttl[p])
            else:
                kh = [w for w in q if w in self.body[p]]; sc = len(kh) if len(kh) >= min(2, len(q)) else 0
            sc += sum(8 * min(self.codes[p].get(c, 0), 3) for c in cs)
            sc += sum(8 * min(len(w.findall(self.fold[p])), 3) for w in ws)
            if sc > 0: res.append((sc, p))
        return [p for s, p in sorted(res, key=lambda x: -x[0])]


def report(C, name, g, fn, dedup=False, ks=(5, 12, 20, 40)):
    """Chấm theo lớp trùng (trúng bất kỳ bản trùng nào = trúng). dedup=True: bỏ bản trùng khỏi danh sách trước khi cắt top-k."""
    tot = sum(len(v) for v in g.values()); hits = collections.Counter(); inc = a12 = 0; sizes = []
    for m, gs in g.items():
        rk = [C.cls[p] for p in fn(m)]
        if dedup: rk = list(dict.fromkeys(rk))
        sizes.append(len(rk)); gc = {C.cls[p] for p in gs}
        inc += len(gc & set(rk))
        for k in ks: hits[k] += len(gc & set(rk[:k]))
        a12 += bool(gc & set(rk[:12]))
    sizes.sort()
    print(f"  {name:38s} trong UV {inc/tot:4.0%}" + "".join(f"  @{k} {hits[k]/tot:4.0%}" for k in ks)
          + f" | mã ≥1 đúng @12: {a12}/{len(g)} | trung vị UV {sizes[len(sizes)//2]}")


def links(wb, pred):
    for r in wb["_LINKS"].iter_rows(min_row=2, values_only=True):
        if pred(r[1], str(r[2])): yield r


# ---------------- D01 cửa
def d01(cache):
    C = Corpus(cache); byname = {os.path.basename(v).lower(): k for k, v in C.rel.items()}
    sheetdir = {"BLK01": "BLK 01", "BLK02": "BLK 02", "BLK03": "BLK 03", "GEN-BM": "GENERAL, BASEMENT", "TMDV01": "TMDV 01", "TMDVB": "TMDV B"}
    def page(rel, trang):
        b = os.path.basename(rel).lower()
        if b in byname: return (byname[b], int(trang))
        m = re.search(r"_sheets\\([^\\]+)\\p(\d+)_", rel)
        if m:
            fn = f"sb_tender_plot b_{sheetdir[m.group(1)]}_2026 10 08.pdf".lower()
            return (byname[fn], int(m.group(2)))
    wb = openpyxl.load_workbook(os.path.join(SCAN, "PROJECT_DATA_v8.xlsx"), read_only=True)
    rows = list(wb["BOQ_TONGHOP"].iter_rows(min_row=2, values_only=True))
    boq = {i + 2: {"Mã": r[2], "Nhóm": r[1], "Hạng mục": f"{r[1]} {r[3] or ''}", "Mô tả": f"{r[4] or ''} {r[5] or ''}"} for i, r in enumerate(rows) if r[2]}
    khoi = collections.defaultdict(set)
    for r in wb["BOQ_MAP"].iter_rows(min_row=2, values_only=True):
        if r[0]: khoi[r[0]].add(pd._fold(r[1]).replace(" ", ""))
    g = collections.defaultdict(set)
    for r in links(wb, lambda s, c: s == "BOQ_TONGHOP" and re.fullmatch(r"[XY]\d+", c)):
        row = int(r[2][1:]); p = page(str(r[4]), r[5] or 0)  # link sang _sheets (PDF 1 trang) không ghi Trang
        if p and p[1] == 0: p = None
        if p and boq.get(row, {}).get("Nhóm") == "Cửa": g[boq[row]["Mã"]].add(p)
    B = {r["Mã"]: r for r in boq.values() if r["Nhóm"] == "Cửa"}

    def farea(fid):
        n = pd._fold(C.rel[fid]).replace(" ", "")
        if "general" in n or not re.search(r"blk0\d|basement|tmdv", n): return None  # dùng chung
        return set(re.findall(r"blk0\d|basement|tmdv0?\w?", n))
    FA = {f: farea(f) for f in C.texts}
    area = lambda m: (lambda p: FA[p[0]] is None or bool(FA[p[0]] & khoi[m]))

    sup = [f for f in C.texts if FA[f] is None and "general" not in C.rel[f].lower()]
    ncls = len(set(C.cls.values()))
    print(f"\n=== NORTON D01 cửa: {len(C.meta)} PDF / {C.N} trang ({ncls} lớp sau gộp trùng); file tổng hợp không ghi khối: {[C.rel[f] for f in sup]}")
    for f in C.meta: print(f"    {f['ID']} {f['pages']:4d} tr  {f['rel']}")
    dup_sup = sum(1 for p in C.PAGES if p[0] in sup and C.cls[p] != p) + sum(1 for p in C.PAGES if p[0] not in sup and C.cls[p][0] in sup)
    print(f"  trang của file tổng hợp trùng nguyên văn với file khác: {dup_sup}")
    print(f"  đáp án {len(g)} mã / {sum(map(len, g.values()))} trang: " + "; ".join(f"{m} {sorted((C.rel[f][:22], p) for f, p in s)}" for m, s in g.items()))
    for m in B:
        print(f"    mã {m}: chuẩn hóa '{ncode(m)}' có trên {sum(1 for p in C.PAGES if C.codes[p].get(ncode(m)))} trang;"
              f" nguyên văn '{pd._fold(m)}' trên {sum(1 for p in C.PAGES if re.search(r'(?<![a-z0-9])' + re.escape(pd._fold(m)) + r'(?![a-z0-9])', C.fold[p]))} trang")
    R = lambda name, fn, **kw: report(C, name, g, lambda m: fn(B[m]), **kw)
    R("A hiện tại", C.current)
    R("A + bỏ trang trùng", C.current, dedup=True)
    R("C mã chuẩn hóa (không BM25)", lambda r: C.new(r, bm=False, desc=False), dedup=True)
    R("D C + lọc khối", lambda r: C.new(r, bm=False, desc=False, area=area(r["Mã"])), dedup=True)
    R("E BM25 + tên tờ, KHÔNG mã", lambda r: C.new(r, code=False), dedup=True)
    R("F BM25 + tên tờ + mã", lambda r: C.new(r), dedup=True)
    R("G F + lọc khối", lambda r: C.new(r, area=area(r["Mã"])), dedup=True)
    # L1 ở mức trang: file tổng trùng hết với file con -> bỏ file tổng, để lọc khối có tác dụng
    nosup = lambda m: (lambda p: p[0] not in sup and area(m)(p))
    R("G2 G + bỏ file tổng (trùng file con)", lambda r: C.new(r, area=nosup(r["Mã"])), dedup=True)
    R("D2 D + bỏ file tổng", lambda r: C.new(r, bm=False, desc=False, area=nosup(r["Mã"])), dedup=True)
    # tờ thống kê có chứa mã -> cộng điểm (bảng loại tờ do bước boq/kiemke ghi)
    sch = lambda p: bool(re.search(r"schedule|thong ke", pd._fold(C.titles[p[0]][p[1] - 1])))
    def withsch(r, **kw):
        base = C.new(r, **kw); c = ncode(r["Mã"])
        top = [p for p in base if sch(p) and C.codes[p].get(c)]
        return top + [p for p in base if p not in top]
    R("J G2 + ghim tờ thống kê có mã", lambda r: withsch(r, area=nosup(r["Mã"])), dedup=True)
    R("J' D2 + ghim tờ thống kê có mã", lambda r: withsch(r, bm=False, desc=False, area=nosup(r["Mã"])), dedup=True)
    return C, g, B, area


# ---------------- Vanity
def vanity(cache):
    C = Corpus(cache)
    wb = openpyxl.load_workbook(os.path.join(SCAN, "PROJECT_DATA_VANITY_v1.xlsx"), read_only=True)
    rows = list(wb["BOQ_TONGHOP"].iter_rows(min_row=2, values_only=True))
    boq = {i + 2: {"Mã": r[1], "Hạng mục": r[2] or "", "Mô tả": r[4] or ""} for i, r in enumerate(rows) if r[1]}
    maprow = {i + 2: r[0] for i, r in enumerate(wb["BOQ_MAP"].iter_rows(min_row=2, values_only=True))}
    fid = {os.path.basename(v).lower(): k for k, v in C.rel.items()}
    g = collections.defaultdict(set)
    for r in links(wb, lambda s, c: (s == "BOQ_TONGHOP" and c[0] == "Q") or (s == "BOQ_MAP" and c[0] in "GH")):
        if not r[5]: continue
        f = fid.get(os.path.basename(str(r[4])).lower())
        ma = boq[int(r[2][1:])]["Mã"] if r[1] == "BOQ_TONGHOP" else maprow[int(r[2][1:])]
        if f: g[ma].add((f, int(r[5])))
    B = {r["Mã"]: r for r in boq.values()}
    ncls = len(set(C.cls.values()))
    print(f"\n=== NORTON VANITY: {len(C.meta)} PDF / {C.N} trang ({ncls} lớp sau gộp trùng)")
    print(f"  đáp án {len(g)} mã / {sum(map(len, g.values()))} trang: " + "; ".join(f"{m} {sorted(p for f, p in s)}" for m, s in g.items()))
    for m in B: print(f"    mã {m}: '{ncode(m)}' trên {sum(1 for p in C.PAGES if C.codes[p].get(ncode(m)))} trang; Hạng mục '{B[m]['Hạng mục']}'")
    R = lambda name, fn, **kw: report(C, name, g, lambda m: fn(B[m]), **kw)
    R("A hiện tại (Hạng mục tiếng Việt)", C.current)
    R("A + bỏ trang trùng", C.current, dedup=True)
    R("E BM25 Hạng mục+Mô tả + tên tờ", lambda r: C.new(r, code=False), dedup=True)
    R("E' BM25 không tên tờ", lambda r: C.new(r, code=False, title_w=0), dedup=True)
    R("E2 BM25 truy vấn ngắn (Hạng mục + 8 từ đầu Mô tả)", lambda r: C.new(r, code=False, qmax=8), dedup=True)
    lst = lambda p: bool(re.search(r"danh (muc|sach) ban ve|drawing list|legend", pd._fold(C.titles[p[0]][p[1] - 1])))
    R("E3 E2 + bỏ tờ danh mục BV / legend", lambda r: C.new(r, code=False, qmax=8, skip=lst), dedup=True)
    al = {"V01": ["VC1"], "V02": ["VC1", "SS02"], "V03": ["GUONG", "MIRROR"]}
    R("H E + mã BV do bước boq gán (VC1…)", lambda r: C.new(r, code=False, alias=al[r["Mã"]]), dedup=True)
    R("I H + truy vấn ngắn + bỏ danh mục", lambda r: C.new(r, code=False, alias=al[r["Mã"]], qmax=8, skip=lst), dedup=True)
    # đáp án chỉ lấy 1 tờ đại diện/loại căn -> đo thêm: top-12 có bao nhiêu tờ 'chi tiết phòng tắm'
    bt = lambda p: "phong tam" in pd._fold(C.titles[p[0]][p[1] - 1])
    for name, fn in [("A", C.current), ("E", lambda r: C.new(r, code=False)), ("E3", lambda r: C.new(r, code=False, qmax=8, skip=lst)),
                     ("H", lambda r: C.new(r, code=False, alias=al[r["Mã"]])), ("I", lambda r: C.new(r, code=False, alias=al[r["Mã"]], qmax=8, skip=lst))]:
        print(f"    {name}: tờ chi tiết phòng tắm trong top-12 = " + ", ".join(
            f"{m} {sum(map(bt, list(dict.fromkeys(C.cls[p] for p in fn(B[m])))[:12]))}/12" for m in B))
    return C, g, B


def misses(C, g, B, fn, k=12):
    for m, gs in g.items():
        rk = list(dict.fromkeys(C.cls[p] for p in fn(B[m])))
        for p in sorted(gs):
            if C.cls[p] not in rk[:k]:
                pos = rk.index(C.cls[p]) + 1 if C.cls[p] in rk else None
                print(f"    trượt {m} {C.rel[p[0]][:30]} p{p[1]} (hạng {pos}) | {C.titles[p[0]][p[1]-1][:90]}")


if __name__ == "__main__":
    C, g, B, area = d01(sys.argv[1]); print("  -- trượt @12 của G:"); misses(C, g, B, lambda r: C.new(r, area=area(r["Mã"])))
    C, g, B = vanity(sys.argv[2]); print("  -- trượt @12 của E:"); misses(C, g, B, lambda r: C.new(r, code=False))
