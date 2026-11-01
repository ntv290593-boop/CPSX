"""Nạp dữ liệu Code ID vào _MASTER\\code_id (plan DOCS\\PLAN_TENDER-CODEID-DMVT_2026-10-31.md, mục 3).

Chạy:  py -I nap_codeid.py bang     Datawork.xlsx → _MASTER\\code_id\\bang_ma.json
       py -I nap_codeid.py deduyet  SSF → <_ACTIVE>\\SOFITEL\\10_TENDER\\02_extract\\NHO_DUYET.xlsx (danh sách dòng nghi ngờ)
       py -I nap_codeid.py ghi      đọc NHO_DUYET.xlsx đã duyệt → _MASTER\\code_id\\nho.json

Chỉ người quản lý chạy 'ghi', sau khi user duyệt: nho.json do git theo dõi và dùng chung mọi dự án.
Đo trước khi nạp: do_codeid.py dựng bộ nhớ riêng từ phần học nên không phụ thuộc nho.json.
Quy tắc duyệt: cột 'Mã chốt' có mã thì dòng vào bộ nhớ với mã đó; để trống thì dòng không vào bộ nhớ.
"""
import json
import os
import re
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.realpath(__file__)))
import do_codeid as dc  # noqa: E402

ci = dc.ci
DU_AN = "SSP"
DATAWORK = os.path.join(dc.THU_MUC, "Datawork.xlsx")
MASTER = os.path.join(dc.ROOT, "_MASTER", "code_id")
DUYET = os.path.join(dc.GOC, "10_TENDER", "02_extract", "NHO_DUYET.xlsx")
COT = ["Loại nghi ngờ", "Sheet", "Dòng SSF", "Khóa SCOPE", "Mục lớn", "Đầu mục", "Mô tả BOQ CĐT", "Mô tả SSF", "Mã SSF", "Mã chốt"]


def lenh_bang():
    """Cột D, G, I của Datawork. Giữ đúng giá trị cột G, kể cả dấu cách trong 'LS. HA' (quyết định 3)."""
    ws = openpyxl.load_workbook(DATAWORK, data_only=True).worksheets[0]
    rows, cap2, canh_bao = [], "", []
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if r[2]:
            cap2 = str(r[2]).strip()
        d, g = str(r[3] or "").strip(), str(r[6] or "").strip()
        if not g:
            if d or r[4]:
                print(f"Dòng {i}: nhóm '{str(r[4] or '').strip()}' chưa có mã cột G, bỏ qua")
            continue
        if re.search(r"\s", g):
            canh_bao.append(f"Dòng {i}: mã có dấu cách '{g}', giữ nguyên theo quyết định 3")
        if len(d) <= 4:
            canh_bao.append(f"Dòng {i}: tên cấp 3 chỉ ghi '{d}' ({g}, nhóm {cap2}), droplist dễ nhầm")
        rows.append({"cap3": d, "ma": g, "cap2": cap2, "mo_ta": str(r[8] or "").replace("\n", " / ").strip(), "dong": i})
    if len({x["cap3"] for x in rows}) != len(rows) or len({x["ma"] for x in rows}) != len(rows):
        sys.exit("LỖI: cột D hoặc G bị trùng, droplist sẽ tra sai mã")
    os.makedirs(MASTER, exist_ok=True)
    path = os.path.join(MASTER, "bang_ma.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump({"_ghi_chu": "Chép từ Datawork.xlsx cột D (cap3), G (ma), I (mo_ta); chạy lại: nap_codeid.py bang. "
                               "Không sửa tay.", "rows": rows}, fh, ensure_ascii=False, indent=1)
    print(f"Đã ghi {path}: {len(rows)} mã")
    for x in canh_bao:
        print("XEM", x)


def tim_nghi_ngo(co, le):
    """Danh sách (loại, dòng SCOPE hoặc None, dict nhãn lạ). Dòng có mặt ở đây không tự vào bộ nhớ."""
    ra = [("Khác chữ với BOQ CĐT", r, None) for r in co if r.get("_lech")]
    ra += [("Không có ở BOQ CĐT", None, x) for x in le]
    da = {id(r) for _, r, _ in ra if r is not None}

    nhom_mt, nhom_dm = defaultdict(list), defaultdict(list)
    for r in co:
        if not ci.chi_kich_thuoc(r["Mô tả BOQ"]):
            nhom_mt[ci.khoa_mo_ta(r)].append(r)
    for k, ds in nhom_mt.items():
        if len({r["_ma"] for r in ds}) > 1:
            for r in ds:
                if id(r) not in da:
                    ra.append(("Cùng mô tả nhưng khác mã", r, None))
                    da.add(id(r))
    for r in co:
        if r.get("Mục phụ"):
            nhom_dm[ci.khoa_dau_muc(r)].append(r)
    for k, ds in nhom_dm.items():
        if len({r["_ma"] for r in ds}) > 1:
            for r in ds:
                if id(r) not in da:
                    ra.append(("Cùng đầu mục, họ mã nhưng khác mã", r, None))
                    da.add(id(r))

    # Mã lệch hẳn so với cả họ FF&E (54-404 Lamp table = DE.LI trong khi họ 4xx là LO.TB).
    ho = defaultdict(list)
    for r in co:
        tt = ci.tien_to_ma(r)
        if re.fullmatch(r"F\d", tt):
            ho[tt].append(r)
    for tt, ds in ho.items():
        dem = Counter(r["_ma"] for r in ds)
        top, n = dem.most_common(1)[0]
        if n / len(ds) >= 0.8:
            for r in ds:
                if r["_ma"] != top and id(r) not in da:
                    ra.append((f"Lệch mã so với họ {tt}", r, None))
                    da.add(id(r))
    return ra


def lenh_deduyet():
    scope, le = dc.nap_du_lieu()
    co = [r for r in scope if "_ma" in r]
    ds = tim_nghi_ngo(co, le)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DUYET"
    for j, c in enumerate(COT, 1):
        o = ws.cell(row=1, column=j, value=c)
        o.font, o.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78")
    for i, (loai, r, x) in enumerate(ds, 2):
        if r is not None:
            hang = [loai, r["Sheet"], r.get("_dong_ssf"), r["Khóa"], r.get("Nhóm"), r.get("Mục phụ"), r["Mô tả BOQ"],
                    r.get("_mo_ta_ssf", r["Mô tả BOQ"]), r["_ma"], None]
        else:
            hang = [loai, x["sheet"], x["dong"], None, None, None, None, x["mo_ta"], x["ma"], None]
        for j, v in enumerate(hang, 1):
            o = ws.cell(row=i, column=j, value=v)
            o.alignment = Alignment(wrap_text=True, vertical="top")
    for c, w in zip("ABCDEFGHIJ", (28, 22, 8, 28, 20, 26, 55, 55, 10, 10)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{len(ds) + 1}"
    os.makedirs(os.path.dirname(DUYET), exist_ok=True)
    wb.save(DUYET)
    print(f"Đã ghi {DUYET}: {len(ds)} dòng cần duyệt")
    print(Counter(x[0] for x in ds).most_common())
    bo = {r["Khóa"] for _, r, _ in ds if r is not None}
    nho = ci.hoc([r for r in co if r["Khóa"] not in bo], DU_AN)
    print(f"Không tính dòng nghi ngờ, bộ nhớ có {len(nho['mo_ta'])} mẫu mô tả, {len(nho['dau_muc'])} mẫu đầu mục")


def lenh_ghi():
    bang = ci.nap_bang()
    if not bang:
        sys.exit("LỖI: chưa có bang_ma.json, chạy lệnh bang trước")
    if not os.path.exists(DUYET):
        sys.exit(f"LỖI: chưa có {DUYET}, chạy lệnh deduyet trước")
    ws = openpyxl.load_workbook(DUYET, data_only=True)["DUYET"]
    hang = list(ws.iter_rows(values_only=True))
    idx = {c: hang[0].index(c) for c in COT}
    chot, loi = {}, []
    tat_ca = set()
    for r in hang[1:]:
        k = r[idx["Khóa SCOPE"]]
        if k:
            tat_ca.add(k)
        ma = str(r[idx["Mã chốt"]] or "").strip()
        if not ma:
            continue
        if not k:
            loi.append(f"'{r[idx['Mô tả SSF']]}': dòng không có ở BOQ CĐT, không chốt được vào bộ nhớ")
        elif not bang.co(ma):
            loi.append(f"{k}: Mã chốt '{ma}' không có trong bảng mã")
        else:
            chot[k] = ma
    if loi:
        sys.exit("LỖI:\n  " + "\n  ".join(loi))
    scope, _ = dc.nap_du_lieu()
    dung = []
    for r in scope:
        if "_ma" not in r:
            continue
        if r["Khóa"] in chot:
            r["_ma"] = chot[r["Khóa"]]
        elif r["Khóa"] in tat_ca:
            continue
        dung.append(r)
    nho = ci.hoc(dung, DU_AN)
    for ds in nho.values():
        for x in ds:
            x["ma"] = next(b["ma"] for b in bang.rows if ci.rut_ma(b["ma"]) == ci.rut_ma(x["ma"]))
    so_lo = [x for ds in nho.values() for x in ds for c in ("mo_ta", "dau_muc") if re.search(r"\d", x.get(c, ""))]
    if so_lo:
        sys.exit(f"LỖI: bộ nhớ còn chữ số, có thể lộ số liệu dự án: {so_lo[:3]}")
    path = os.path.join(MASTER, "nho.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(dict({"_ghi_chu": "Mẫu chữ → mã dùng chung; chỉ người quản lý sửa, ghi bằng nap_codeid.py ghi."}, **nho),
                  fh, ensure_ascii=False, indent=1)
    print(f"Đã ghi {path}: {len(nho['mo_ta'])} mẫu mô tả, {len(nho['dau_muc'])} mẫu đầu mục; {len(chot)} dòng theo Mã chốt")


if __name__ == "__main__":
    fn = {"bang": lenh_bang, "deduyet": lenh_deduyet, "ghi": lenh_ghi}.get(sys.argv[1] if len(sys.argv) > 1 else "")
    if not fn:
        sys.exit(__doc__)
    fn()
