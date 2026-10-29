# -*- coding: utf-8 -*-
"""Bài test thử ý tưởng `neo` (đối chiếu giá trị đã ghi với text trang PDF) trên MAT.mat.json của SOFITEL tender.
Chỉ đọc. Không ghi vào dự án."""
import json, re, sys, zipfile, random, unicodedata, collections, os
import pymupdf

ROOT = r"D:\COSTING EXCEL\_ACTIVE\SOFITEL"
DATA = ROOT + r"\10_TENDER\_data"
INP = ROOT + r"\01. INPUT\new"
random.seed(7)

fi = {x["ID"]: x for x in json.load(open(DATA + r"\FILE_INDEX.kiemke-tool.json", encoding="utf-8"))["rows"]}
mat = json.load(open(DATA + r"\MAT.mat.json", encoding="utf-8"))["rows"]

# ---- chuẩn hoá (theo anchor.py của drawing-analyzer, rút gọn)
FOLD = {"–": "-", "—": "-", "−": "-", "‐": "-", "‘": "'", "’": "'", "“": '"', "”": '"', "×": "x", "Ø": "o", "ø": "o",
        "\u00ad": "", "\u200b": ""}
def norm(s):
    s = re.sub(r"([½¼¾])", r" \1", str(s))
    s = unicodedata.normalize("NFKC", s)
    s = "".join(FOLD.get(c, c) for c in s)
    s = re.sub(r"(?<=\w)-(?=\w)", " ", s)  # gạch nối giữa 2 chữ -> cách (PT-04 == PT 04)
    return re.sub(r"\s+", " ", s.lower()).strip()
def toks(s): return norm(s).split()
def has_digit(t): return any(c.isdigit() for c in t)

# ---- mở PDF trong zip
_pdf = {}
def open_fid(fid):
    if fid in _pdf: return _pdf[fid]
    r = fi[fid]
    z = zipfile.ZipFile(os.path.join(INP, r["Đường dẫn"]))
    data = z.read(r["Trong zip"])
    d = pymupdf.open(stream=data, filetype="pdf")
    _pdf[fid] = d
    return d

_pg = {}
def page_tokens(fid, pg):
    k = (fid, pg)
    if k in _pg: return _pg[k]
    d = open_fid(fid)
    t = [w[4] for w in d[pg - 1].get_text("words")]
    st = toks(" ".join(t))
    _pg[k] = (st, len(" ".join(t).strip()))
    return _pg[k]

# ---- bài A: độ phủ text
print("=== A. Độ phủ lớp text của 3 PDF nguồn ===")
for fid in ("F0042", "F0043", "F0044"):
    d = open_fid(fid)
    n0 = sum(1 for p in d if len(p.get_text().strip()) < 20)
    print(f"{fid}: {d.page_count} trang, {n0} trang gần như không có text; {os.path.basename(fi[fid]['Trong zip'])}")

# ---- chuẩn bị dữ liệu thử: mỗi dòng MAT có link hợp lệ
rows = []
for r in mat:
    l = r.get("Nguồn spec")
    if isinstance(l, list) and len(l) >= 3 and isinstance(l[2], int) and l[1] in fi:
        rows.append(r)
print(f"\nDòng MAT có link F###/trang: {len(rows)}/{len(mat)}")

def code_tokens(*vals):
    out = []
    for v in vals:
        for t in toks(v or ""):
            if has_digit(t): out.append(t)
    return out

def check(vals, fid, pg, mode="all"):
    """Trả về (trạng thái, thiếu). vals = các chuỗi cần có trên trang."""
    st, nchar = page_tokens(fid, pg)
    if nchar < 20: return "KHONG_KIEM_DUOC", []
    S = collections.Counter(st)
    need = code_tokens(*vals)
    if not need: return "KHONG_CO_MA", []
    miss = [t for t in need if S[t] == 0]
    return ("KHOP" if not miss else "LECH"), miss

# ---- bài B: kiểm dương tính
print("\n=== B. Kiểm dương tính (giá trị thật, link thật) ===")
res = collections.Counter(); lech = []
for r in rows:
    _, fid, pg = r["Nguồn spec"][:3]
    s, miss = check([r["Mã VL"], r.get("NCC/Mã SP"), r.get("Mô tả")], fid, pg)
    res[s] += 1
    if s == "LECH": lech.append((r["Mã VL"], fid, pg, miss, (r.get("NCC/Mã SP") or "")[:50]))
print(dict(res))
for x in lech[:25]: print("  LECH", x)

# phân tích từng trường riêng
print("\n  theo từng trường (số dòng có thiếu mã token):")
for fld in ("Mã VL", "NCC/Mã SP", "Mô tả"):
    c = collections.Counter()
    for r in rows:
        _, fid, pg = r["Nguồn spec"][:3]
        s, _m = check([r.get(fld)], fid, pg); c[s] += 1
    print("   ", fld, dict(c))

# ---- bài C: đột biến — sửa 1 chữ số trong token mã, xem có bắt được không
print("\n=== C. Đột biến: đổi 1 chữ số trong mã (giả lập LLM ghi sai số) ===")
def mutate(t):
    idx = [i for i, c in enumerate(t) if c.isdigit()]
    i = random.choice(idx)
    c = t[i]; n = random.choice([d for d in "0123456789" if d != c])
    return t[:i] + n + t[i + 1:]
caught = total = 0; fp_chance = 0; ex = []
for r in rows:
    _, fid, pg = r["Nguồn spec"][:3]
    base_vals = [r["Mã VL"], r.get("NCC/Mã SP")]
    ts = code_tokens(*base_vals)
    if not ts: continue
    for _ in range(5):
        # thay 1 token mã bằng bản đột biến
        orig = random.choice(ts); m = mutate(orig)
        total += 1
        s, miss = check([m], fid, pg)
        if s == "LECH": caught += 1
        else:
            fp_chance += 1
            if len(ex) < 8: ex.append((r["Mã VL"], orig, m, "vẫn có trên trang"))
print(f"bắt được {caught}/{total} = {100*caught/max(total,1):.1f}%; lọt {fp_chance}")
for e in ex: print("  lọt:", e)

# ---- bài D: trang sai — dịch link sang trang khác, có bắt được và gợi ý đúng trang không
print("\n=== D. Link sai trang (giả lập) ===")
caught = total = sugg_ok = 0
for r in rows:
    _, fid, pg = r["Nguồn spec"][:3]
    d = open_fid(fid)
    other = [p for p in range(1, d.page_count + 1) if p != pg]
    if not other: continue
    wrong = random.choice(other)
    total += 1
    s, miss = check([r["Mã VL"], r.get("NCC/Mã SP")], fid, wrong)
    if s == "LECH":
        caught += 1
        # gợi ý: trang nào chứa đủ mã
        best = [p for p in range(1, d.page_count + 1) if check([r["Mã VL"], r.get("NCC/Mã SP")], fid, p)[0] == "KHOP"]
        if pg in best: sugg_ok += 1
print(f"trang sai bị bắt {caught}/{total}; gợi ý có chứa trang đúng {sugg_ok}/{caught}")

# ---- bài E: kiểm yếu "chỉ số" vs có phrase: mã đơn lẻ có hay trùng trên trang không?
print("\n=== E. Độ đặc hiệu: số đứng riêng có xuất hiện sẵn trên trang ngẫu nhiên không? ===")
hits = tot = 0
for r in rows:
    _, fid, pg = r["Nguồn spec"][:3]
    d = open_fid(fid)
    ts = code_tokens(r["Mã VL"], r.get("NCC/Mã SP"))
    if not ts: continue
    other = [p for p in range(1, d.page_count + 1) if p != pg]
    if not other: continue
    w = random.choice(other); st, _n = page_tokens(fid, w); S = set(st)
    tot += 1
    if all(t in S for t in ts): hits += 1
print(f"mã đúng của dòng X lại có đủ trên 1 trang ngẫu nhiên khác: {hits}/{tot}")
