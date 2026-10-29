# -*- coding: utf-8 -*-
"""Bộ đo cố định cho phần tìm trang (pd.py ungvien). Gọi thẳng hàm xếp hạng trong pd.py, không chép logic.
py -X utf8 bench.py          đọc _cache\\Blanca\\_data (bản chụp 28/10, không đụng _data của dự án)
Đáp án: 'Bản vẽ đối chứng' của bước banve + sửa theo KIEM_MAU_DAP_AN_Blanca_v2.xlsx (user kiểm 28/10).
Cổng qua: không số nào tụt so với MOC."""
import sys, os, re, json, shutil, tempfile, collections, openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "..", "SKILL", "01-scan", "tool"))
import pd

CACHE = os.path.join(HERE, "_cache", "Blanca", "_data")
# mốc trước khi sửa (28/10, proto_rank): @12 = 73% (10 mã kiểm) / 62% (69 mã); prototype tốt nhất 98% / 89%
MOC = {"kiem@12": 0.73, "all@12": 0.62}


def setup():
    """Bản làm việc tạm của cache để pd.build_timkiem ghi _timkiem.json vào đó, không vào cache."""
    tmp = tempfile.mkdtemp(); d = os.path.join(tmp, "01_scan", "_data", "X")
    shutil.copytree(CACHE, d)
    # quyết định LLM bước kiemke B mô phỏng: F004 = concept/layout dự án -> DỰ PHÒNG (NOTES.kiemke của Blanca ghi "concept/đề xuất")
    json.dump({"du_phong": ["F004"]}, open(os.path.join(d, "_timkiem.kiemke.json"), "w"))
    return pd.paths(tmp, "X")


def gold(sheets):
    g = collections.defaultdict(set)
    for b, r in sheets["BOQ_TONGHOP"]:
        if not b.startswith("banve"): continue
        v = r.get("Bản vẽ đối chứng") or []
        if isinstance(v, dict): v = v.get("links", [])
        if isinstance(v, str): continue
        for l in v:
            if isinstance(l, dict): l = [l.get("nhan"), l.get("file"), l.get("trang")]
            if isinstance(l, list) and len(l) >= 3 and l[2] and re.fullmatch(r"F\d+", str(l[1])):
                g[str(r["Mã"])].add((l[1], int(l[2])))
    wb = openpyxl.load_workbook(os.path.join(HERE, "KIEM_MAU_DAP_AN_Blanca_v2.xlsx"), read_only=True)
    rows = [r for r in wb["KIEM_MAU"].iter_rows(min_row=2, values_only=True) if r[0]]
    add, rm, sai = collections.defaultdict(set), collections.defaultdict(set), collections.defaultdict(set)
    for r in rows:
        p = (r[3], int(r[4]))
        if r[9] == "ĐÚNG": add[r[0]].add(p)
        elif r[9] == "SAI": rm[r[0]].add(p); sai[r[0]].add(p)
    kiem = sorted({r[0] for r in rows})
    for m in kiem: g[m] = (g.get(m, set()) | add[m]) - rm[m]
    return dict(g), kiem, sai


def main():
    P = setup(); sheets = pd.load_json(P["data"])[0]
    ix = pd.build_index(P, sheets)
    allm, _ = pd.ungvien_all(ix, sheets)
    g, kiem, sai = gold(sheets)
    keep = lambda p: ix["dup"].get(p, p)  # trang trùng -> bản giữ lại
    g = {m: {keep(p) for p in s} for m, s in g.items() if m in allm}
    flat = {m: [x[1] for x in allm[m]["res"]] for m in g}
    # nhóm hiển thị (cùng quy tắc cmd_ungvien) để đo "trang có trong đầu ra"
    def shown(m, top=8):
        d, out = allm[m], []
        for loai, _, n in pd.NHOM_RA:
            if loai == "chi_tiet":
                hit = [x for x in d["res"] if ix["loai"][x[1]] in ("chi_tiet", "mat_dung", "khac", "legend")]; n = top
            else:
                hit = [x for x in d["res"] if ix["loai"][x[1]] == loai and x[2]]
            out += [x[1] for x in hit[:n]]
        return out + [x[1] for x in d["res_cha"][:6]]
    def line(title, codes):
        tot = sum(len(g[m]) for m in codes)
        r = {}
        for k in (5, 12, 40, 10**6):
            r[k] = sum(1 for m in codes for p in g[m] if p in set(flat[m][:k]))
        sh = sum(1 for m in codes for p in g[m] if p in set(shown(m)))
        bad5 = sum(1 for m in codes for p in sai.get(m, ()) if keep(p) in set(flat[m][:5]))
        nbad = sum(len(sai.get(m, ())) for m in codes)
        print(f"{title:28s} mã={len(codes):3d} trang={tot:3d}  @5 {r[5]/tot:4.0%}  @12 {r[12]/tot:4.0%}  @40 {r[40]/tot:4.0%}  hết {r[10**6]/tot:4.0%}"
              f" | trong đầu ra {sh/tot:4.0%} | SAI đã biết lọt top-5: {bad5}/{nbad}")
        return r[12] / tot
    print(f"Chỉ mục: {ix['N']} trang, trùng {len(ix['dup'])}, file tổng {sorted(ix['tong'])}")
    a = line("10 mã đã kiểm", [m for m in kiem if m in g])
    b = line("69 mã (59 chưa kiểm)", list(g))
    ok = a >= MOC["kiem@12"] and b >= MOC["all@12"]
    print("CỔNG:", "QUA" if ok else "TỤT so với mốc", MOC)
    shutil.rmtree(os.path.dirname(os.path.dirname(os.path.dirname(P["data"]))), ignore_errors=True)


if __name__ == "__main__":
    main()
