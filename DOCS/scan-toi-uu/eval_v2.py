# -*- coding: utf-8 -*-
"""Đo lại trên đáp án Blanca đã sửa theo KIEM_MAU_DAP_AN_Blanca_v2 (user kiểm 28/10): thêm trang ĐÚNG, bỏ trang SAI.
py -X utf8 eval_v2.py"""
import collections, openpyxl
import proto_rank as P

wb = openpyxl.load_workbook("KIEM_MAU_DAP_AN_Blanca_v2.xlsx", read_only=True)
rows = [r for r in wb["KIEM_MAU"].iter_rows(min_row=2, values_only=True) if r[0]]
al = lambda p: (P.ALIAS.get(p[0], p[0]), p[1])
# thống kê kết quả kiểm
t = collections.Counter((("đáp án" if r[6] == "CÓ" else "") + ("+" if r[6] == "CÓ" and r[7] else "") + ("máy" if r[7] else ""), r[9]) for r in rows)
print("Kết quả kiểm (nguồn → đánh giá):")
for k, v in sorted(t.items()): print(f"   {k[0]:12s} {k[1]:11s} {v}")
add, rm = collections.defaultdict(set), collections.defaultdict(set)
for r in rows:
    p = (r[3], int(r[4]))
    if r[9] == "ĐÚNG": add[r[0]].add(p)
    elif r[9] == "SAI": rm[r[0]].add(p)
codes = sorted({r[0] for r in rows})
g0 = {m: set(s) for m, s in P.g.items()}
for m in codes:
    P.g[m] = (P.g.get(m, set()) | add[m]) - rm[m]
print(f"Đáp án: {sum(map(len, g0.values()))} → {sum(map(len, P.g.values()))} trang (thêm {sum(len(add[m] - g0.get(m, set())) for m in codes)}, bỏ {sum(len(rm[m] & g0.get(m, set())) for m in codes)})")


def suite(title, keep):
    full = {m: s for m, s in P.g.items() if keep(m)}
    saved = dict(P.g); P.g.clear(); P.g.update(full)
    print(f"\n== {title}")
    cur = P.ungvien_current(P.boq)
    P.report("A hiện tại (gộp F023)", {m: [(x[0], *al((x[1], x[2])), x[3]) for x in v] for m, v in cur.items()},
             {m: {al(p) for p in s} for m, s in P.g.items()})
    P.run("C +mã chuẩn hóa", use_area=False, use_bm25=False, desc=False)
    P.run("D +lọc khu", use_bm25=False, desc=False)
    P.run("G BM25+tên tờ+Mô tả+lọc khu", desc=True)
    P.run("H như G, KHÔNG lọc khu", use_area=False)
    P.g.clear(); P.g.update(saved)


suite("10 mã đã kiểm (đáp án đầy đủ hơn)", lambda m: m in codes)
suite("toàn bộ 69 mã (59 mã chưa kiểm giữ nguyên)", lambda m: True)
