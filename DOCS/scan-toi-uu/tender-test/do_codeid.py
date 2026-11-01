"""Đo phân loại Code ID (plan DOCS\\PLAN_TENDER-CODEID-DMVT_2026-10-31.md, mục 2, GĐ2).

Chạy:  py -I do_codeid.py            đo học chéo, in bảng; thoát 1 nếu không đạt mốc
       py -I do_codeid.py --chitiet  in thêm từng dòng sai theo lớp
       py -I do_codeid.py --165      phân định lệch 165/205 dòng của lớp Luật

Đáp án: SSF-BOQ-CODE ID.xlsx cột AT, bỏ đuôi .xxx. Chỉ để đo, script này không ghi gì.
Luật và cách nhớ lấy từ SKILL\\_core\\code_id.py + _MASTER\\code_id\\luat.json: đo đúng code chạy thật.
Bộ nhớ dựng lại từ phần học của mỗi chiều, không đọc nho.json (nạp rồi mới đo là lộ đáp án).
Mốc plan GĐ2: gộp 2 chiều đúng ≥ 80%, sai ≤ 6%. Không nới mốc sau khi thấy kết quả; đổi mốc do user quyết định.
"""
import difflib
import json
import os
import re
import sys
from collections import Counter, defaultdict

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "..", ".."))
sys.path.insert(0, os.path.join(ROOT, "SKILL", "_core"))
import code_id as ci  # noqa: E402

GOC = os.path.join(ROOT, "_ACTIVE", "SOFITEL")
THU_MUC = os.path.join(GOC, "01. INPUT", "new", "New folder")
SSF = os.path.join(THU_MUC, "SSF-BOQ-CODE ID.xlsx")
SCOPE = os.path.join(GOC, "10_TENDER", "_data", "SCOPE.boq-tool.json")
SHEET = ["2. Finishing-L4-17", "2.1 Finishing-L18-20", "3. Furniture-L4-17", "3.1 Furniture-suite-L18-20"]
COT_AT = 46
MOC_DUNG, MOC_SAI = 0.80, 0.06
RX_DUOI = re.compile(r"\.\d{1,4}$")


def chuan(t):
    return re.sub(r"[^a-z0-9]+", " ", str(t or "").lower()).strip()


def bo_duoi(ma):
    return re.sub(r"\s+", "", RX_DUOI.sub("", str(ma).strip()))


def doc_ssf():
    """{sheet: [{dong, item, mo_ta, ma}]} theo thứ tự. Tiêu đề cột 'CODE-ID' không tính là nhãn."""
    wb = openpyxl.load_workbook(SSF, data_only=True)
    ra = {}
    for sh in SHEET:
        dong = []
        for i, r in enumerate(wb[sh].iter_rows(min_row=9, values_only=True), 9):
            mo_ta = re.sub(r"\s+", " ", " ".join(str(x) for x in r[1:5] if x not in (None, "")))
            ma = r[COT_AT - 1]
            ma = bo_duoi(ma) if ma and str(ma).strip().upper() != "CODE-ID" else ""
            if mo_ta or ma:
                dong.append({"dong": i, "item": str(r[0] or "").strip(), "mo_ta": mo_ta, "ma": ma})
        ra[sh] = dong
    return ra


def noi(scope, ssf):
    """Gán nhãn SSF vào dòng SCOPE theo nội dung: căn chuỗi trong từng sheet, phần lệch chữ nối theo độ giống ≥ 0,8.

    Trả danh sách nhãn không nối được; dòng nối theo độ giống có cờ '_lech' + '_mo_ta_ssf'.
    """
    le = []
    for sh in SHEET:
        a, b = [r for r in scope if r["Sheet"] == sh], ssf[sh]
        khoa_a = [chuan(r["Mô tả BOQ"])[:150] for r in a]
        khoa_b = [chuan(r["mo_ta"])[:150] for r in b]
        dung = set()
        for m in difflib.SequenceMatcher(None, khoa_a, khoa_b, autojunk=False).get_matching_blocks():
            for k in range(m.size):
                dung.add(m.b + k)
                if b[m.b + k]["ma"]:
                    a[m.a + k]["_ma"] = b[m.b + k]["ma"]
                    a[m.a + k]["_dong_ssf"] = b[m.b + k]["dong"]
        con_a = [i for i, r in enumerate(a) if "_ma" not in r]
        for bi in (i for i in range(len(b)) if i not in dung and b[i]["ma"]):
            tot, diem = None, 0.0
            for ai in con_a:
                d = difflib.SequenceMatcher(None, khoa_a[ai], khoa_b[bi]).ratio()
                if d > diem:
                    tot, diem = ai, d
            if tot is not None and diem >= 0.8:
                a[tot].update(_ma=b[bi]["ma"], _lech=True, _mo_ta_ssf=b[bi]["mo_ta"], _dong_ssf=b[bi]["dong"])
                con_a.remove(tot)
            else:
                le.append({"sheet": sh, "dong": b[bi]["dong"], "ma": b[bi]["ma"], "mo_ta": b[bi]["mo_ta"]})
    return le


def nap_du_lieu():
    with open(SCOPE, encoding="utf-8") as fh:
        scope = [r for r in json.load(fh)["rows"] if r["Sheet"] in SHEET]
    le = noi(scope, doc_ssf())
    return scope, le


def do(dong_hoc, dong_doan, luat):
    nho = ci.Nho(ci.hoc(dong_hoc, "X"))
    rong = ci.Nho()
    lop, sai = defaultdict(lambda: [0, 0]), []
    for r in dong_doan:
        ma, nguon, _ = ci.phan_loai(r, luat, rong, nho, du_an="X")
        if ma is None:
            lop["LLM"][1] += 1
        elif ci.rut_ma(ma) == ci.rut_ma(r["_ma"]):
            lop[nguon][0] += 1
        else:
            lop[nguon][1] += 1
            sai.append((nguon, r["Khóa"], ma, r["_ma"], r["Mô tả BOQ"][:60]))
    dung = sum(v[0] for k, v in lop.items() if k != "LLM")
    so_sai = sum(v[1] for k, v in lop.items() if k != "LLM")
    return {"n": len(dong_doan), "dung": dung, "sai": so_sai, "con": lop["LLM"][1], "lop": dict(lop), "ds_sai": sai}


def in_kq(ten, kq):
    n = kq["n"] or 1
    print(f"{ten:<26} n={kq['n']:>3}  đúng {kq['dung']:>3} ({kq['dung'] / n:6.1%})  "
          f"sai {kq['sai']:>3} ({kq['sai'] / n:5.1%})  xuống LLM {kq['con']:>3} ({kq['con'] / n:5.1%})")


def phan_dinh_165(co):
    """Lớp Luật đếm theo từng nguồn đọc tiền tố FF&E. Luật nhóm + IR/SW giữ nguyên trong mọi cách."""
    rx = {"2": "LO.RUG", "3": "DE.CU", "7": "DE.LI"}
    luat_nen = ci.BangLuat([x for x in ci._doc("luat.json")["rows"] if not x["ten"].startswith("FF&E")])

    def chay(ten, ham):
        dung, sai, cach = 0, 0, Counter()
        for r in co:
            ma, _ = luat_nen.xet(r)
            cach_nay = "nhóm/IR/SW"
            if not ma:
                loai = ham(r)
                ma, cach_nay = rx.get(loai), f"{loai}xx"
            if not ma:
                continue
            if ma == r["_ma"]:
                dung += 1
                cach[cach_nay] += 1
            else:
                sai += 1
                cach[cach_nay + " sai"] += 1
        print(f"  {ten:<52} đúng {dung:>3}  sai {sai:>2}  {dict(cach)}")

    def tien_to(t, chi_dau):
        m = re.match(r"\s*5\d\s*-\s*(\d)\d{2}", t or "") if chi_dau else re.search(r"5\d\s*-\s*(\d)\d{2}", t or "")
        return m.group(1) if m else None

    print("\n--- Phân định lớp Luật: đổi nguồn đọc tiền tố FF&E")
    chay("không có luật FF&E", lambda r: None)
    chay("cột Mã spec/FF&E (khớp 165 ở plan mục 2)", lambda r: tien_to(r.get("Mã spec/FF&E"), False))
    chay("đầu Mô tả BOQ", lambda r: tien_to(r["Mô tả BOQ"], True))
    chay("đầu Mô tả BOQ hoặc đầu Mục phụ (cách trong luat.json)", lambda r: tien_to(r["Mô tả BOQ"], True) or tien_to(r.get("Mục phụ"), True))
    chay("mọi vị trí trong Mô tả BOQ hoặc Mục phụ", lambda r: tien_to(r["Mô tả BOQ"], False) or tien_to(r.get("Mục phụ"), False))


def main():
    args = set(sys.argv[1:])
    scope, le = nap_du_lieu()
    co = [r for r in scope if "_ma" in r]
    khong = [r for r in scope if "_ma" not in r]
    print(f"SCOPE 4 sheet: {len(scope)} dòng; nhãn nối được {len(co)} (theo độ giống {sum(1 for r in co if r.get('_lech'))}); "
          f"nhãn SSF không có ở CĐT {len(le)}; dòng SCOPE không có nhãn {len(khong)}")
    bang = ci.nap_bang()
    if bang:
        la = sorted({r["_ma"] for r in co if not bang.co(r["_ma"])})
        print("Nhãn không có trong bảng mã:", la or "không")
    luat = ci.nap_luat()
    f4 = [r for r in co if r["Sheet"] in (SHEET[0], SHEET[2])]
    f18 = [r for r in co if r["Sheet"] in (SHEET[1], SHEET[3])]
    fin = [r for r in co if "Finishing" in r["Sheet"]]
    fur = [r for r in co if "Furniture" in r["Sheet"]]

    print("\n--- Học chéo tầng (mốc GĐ2)")
    a, b = do(f4, f18, luat), do(f18, f4, luat)
    in_kq("L4-17 → L18-20", a)
    in_kq("L18-20 → L4-17", b)
    gop = {k: a[k] + b[k] for k in ("n", "dung", "sai", "con")}
    in_kq("GỘP 2 chiều", gop)
    tong = defaultdict(lambda: [0, 0])
    for x in (a, b):
        for k, v in x["lop"].items():
            tong[k][0] += v[0]
            tong[k][1] += v[1]
    for k in ("Luật", "Nhớ-mô-tả", "Nhớ-đầu-mục"):
        d, s = tong[k]
        print(f"   {k:<12} đúng {d:>3}  sai {s:>3}  tỷ lệ sai {s / (d + s or 1):.1%}")

    print("\n--- Độ phủ Finishing ↔ Furniture (gần đúng cho BOQ khác, không có mốc)")
    in_kq("Finishing → Furniture", do(fin, fur, luat))
    in_kq("Furniture → Finishing", do(fur, fin, luat))

    print(f"\n--- {len(khong)} dòng không có nhãn (không chấm được), theo mục lớn")
    print("  ", Counter(r.get("Nhóm", "") for r in khong).most_common())
    print(f"--- Phân loại thử {len(khong)} dòng này với bộ nhớ học từ cả {len(co)} dòng có nhãn (chỉ để biết nguồn)")
    nho, rong = ci.Nho(ci.hoc(co, "X")), ci.Nho()
    print("  ", Counter(ci.phan_loai(r, luat, rong, nho, du_an="X")[1] or "xuống LLM" for r in khong).most_common())

    if "--165" in args:
        phan_dinh_165(co)
    if "--chitiet" in args:
        for x in a["ds_sai"] + b["ds_sai"]:
            print("  SAI", x)

    dung, sai = gop["dung"] / gop["n"], gop["sai"] / gop["n"]
    dat = dung >= MOC_DUNG and sai <= MOC_SAI
    print(f"\nMốc GĐ2: đúng ≥ {MOC_DUNG:.0%}, sai ≤ {MOC_SAI:.0%} → {dung:.1%} / {sai:.1%} → {'ĐẠT' if dat else 'KHÔNG ĐẠT'}")
    sys.exit(0 if dat else 1)


if __name__ == "__main__":
    main()
