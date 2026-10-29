# -*- coding: utf-8 -*-
"""Thử các cách xếp hạng trang ứng viên, so trên đáp án Blanca (Bản vẽ đối chứng của bước banve)."""
import math, collections
from eval_ungvien import *

boq, g = merged_boq(), gold()
files = {r["ID"]: r for b, r in sheets["FILES"]}
ptitle = {(r["File ID"], int(r["Trang"])): f'{r.get("Số BV") or ""} {r.get("Tên BV") or ""}'
          for b, r in sheets["PAGES"] if b == "kiemke"}
khu = collections.defaultdict(set)
for b, r in sheets["BOQ_MAP"]:
    if b == "boq" and r.get("Khu"): khu[str(r["Mã"])].add(pd._fold(r["Khu"]).split()[0])  # studio / 1br / 2br
AREAS = {a for v in khu.values() for a in v}
# khu của file theo đường dẫn (thư mục CĂN STUDIO / CĂN 1BR …); file không ghi khu -> dùng chung
def _area(rel):
    """Khu ghi ở phần sâu nhất của đường dẫn (tên file, rồi thư mục gần nhất) — thư mục cha kiểu 'STUDIO & 1BR' ghi nhiều khu thì bỏ qua."""
    for part in reversed(re.split(r"[\\/]", rel)):
        hit = {a for a in AREAS if re.search(r"(?<![a-z0-9])" + a + r"(?![a-z0-9])", pd._fold(part))}
        if len(hit) == 1: return hit
        if len(hit) > 1: return set()
    return set()


farea = {fid: _area(files[fid]["Đường dẫn"]) for fid in texts}
ALIAS = {"F023": "F024"}  # L1: hai file trùng 42 trang; bước kiemke đã ghi NOTES dùng bản 250930

TOK = re.compile(r"[a-z0-9]+")
STOP = pd.STOP | set("cung cap lap dat hoan thien ban ve hoan cong phong".split())


def toks(s): return [w for w in TOK.findall(pd._fold(s)) if len(w) >= 2 and w not in STOP]


def ncode(s):
    """Chuẩn hóa mã: bỏ . - _ khoảng trắng, bỏ số 0 đầu: LF-04 / LF04 / lf.4 -> lf4."""
    m = re.fullmatch(r"([a-z]{1,4})[\s._-]*0*(\d{1,3})", pd._fold(s).strip())
    return f"{m.group(1)}{m.group(2)}" if m else None


def base_code(ma):
    """Mã BoQ gốc: bỏ hậu tố tự đặt _ST/_1BR/_2BR, -DT/-HM/-TL (hạng mục con), chữ a/b cuối."""
    s = re.sub(r"_(st|1br|2br)$", "", pd._fold(ma))
    s = re.sub(r"-(dt|hm|tl)$", "", s)
    s = re.sub(r"(\d)[a-d]$", r"\1", s)
    return ncode(s)


CODE_IN_TEXT = re.compile(r"(?<![a-z0-9])([a-z]{1,4})[\s._-]?0*(\d{1,3})(?![0-9])")


def page_codes(t): return collections.Counter(a + b for a, b in CODE_IN_TEXT.findall(pd._fold(t)))


# ---- chỉ mục trang
PAGES = [(fid, i) for fid in texts for i in range(1, len(texts[fid]) + 1)]
body = {p: collections.Counter(toks(texts[p[0]][p[1] - 1])) for p in PAGES}
title = {p: collections.Counter(toks(ptitle.get(p, ""))) for p in PAGES}
codes = {p: page_codes(texts[p[0]][p[1] - 1]) for p in PAGES}
N = len(PAGES); avgdl = sum(sum(c.values()) for c in body.values()) / N
df = collections.Counter(w for c in body.values() for w in c)


def bm25(q, c, k1=1.2, b=0.75):
    dl = sum(c.values()); s = 0.0
    for w in q:
        if w in c:
            idf = math.log(1 + (N - df[w] + 0.5) / (df[w] + 0.5))
            s += idf * c[w] * (k1 + 1) / (c[w] + k1 * (1 - b + b * dl / avgdl))
    return s


def rank(ma, r, use_code=True, use_area=True, use_bm25=True, title_w=3.0, dedup=True, desc=True):
    q = set(toks(r.get("Hạng mục") or ""))
    if desc: q |= set(toks((r.get("Mô tả BoQ") or "")[:300]))
    bc = base_code(ma) if use_code else None
    want = khu.get(ma, set())
    res = []
    for p in PAGES:
        if dedup and p[0] in ALIAS: continue
        if use_area and want and farea[p[0]] and not (farea[p[0]] & want): continue
        if use_bm25:
            sc = bm25(q, body[p]) + title_w * sum(1 for w in q if w in title[p])
        else:
            kh = [w for w in q if w in body[p]]; sc = len(kh) if len(kh) >= min(2, len(q)) else 0
        mh = codes[p].get(bc, 0) if bc else 0
        sc += 8 * min(mh, 3) + (10 if bc and bc in {ncode(x) for x in re.findall(r"[A-Za-z]{1,4}[-._]?\d{1,3}", ptitle.get(p, ""))} else 0)
        if sc > 0: res.append((sc, p[0], p[1], mh))
    res.sort(key=lambda x: -x[0])
    return res


def run(name, **kw):
    out = {m: rank(m, boq[m], **kw) for m in g}
    al = lambda p: (ALIAS.get(p[0], p[0]), p[1])
    g2 = {m: {al(p) for p in s} for m, s in g.items()}
    out2 = {m: [(x[0], *al((x[1], x[2])), x[3]) for x in v] for m, v in out.items()}
    report(name, out2, g2)
    return out2


if __name__ == "__main__":
    al = lambda p: (ALIAS.get(p[0], p[0]), p[1])
    cur = ungvien_current(boq)
    report("A hiện tại (gộp F023)", {m: [(x[0], *al((x[1], x[2])), x[3]) for x in v] for m, v in cur.items()},
           {m: {al(p) for p in s} for m, s in g.items()})
    run("B +bỏ file trùng", use_code=False, use_area=False, use_bm25=False, desc=False)
    run("C +mã chuẩn hóa", use_area=False, use_bm25=False, desc=False)
    run("D +lọc khu", use_bm25=False, desc=False)
    run("E +BM25 (chỉ Hạng mục)", desc=False, title_w=0)
    run("F +tên tờ ×3", desc=False)
    run("G +Mô tả BoQ", desc=True)
    run("H BM25+tên tờ, KHÔNG lọc khu", use_area=False)
    run("I BM25+tên tờ, KHÔNG mã", use_code=False)
