"""Kiểm nội dung BOQ tự lập, độc lập với lap_boq.py: đọc lại file BOQ đã ghi và so với tag gốc, spec, PDF.

`td.py check` chỉ kiểm cấu trúc (cột, kiểu, link hợp lệ). File này kiểm đúng sai:
 K1 Tờ link của dòng có thật chứa mã của dòng không (dữ liệu tag DWG).
 K2 Dòng không có nguồn tờ.
 K3 PDF của tờ link có mở được và khung tên có đúng số tờ không.
 K4 Mã tag thuộc họ trong phạm vi có mặt trên tờ mặt đứng/trần của khu mà BOQ không có dòng (sót).
 K5 Mã đồ rời trong lịch FF&E (họ được tick) mà BOQ không có dòng (sót).
 K6 Đếm lại số lượng đồ rời và cửa từ tag gốc, so với cột Số lượng.
 K7 Kích thước đồ rời so lại với chữ spec.
Ghi _data\\_kiem_boq.json + VIEC_MO.doichieu-2.json. Lập lại BOQ thì chạy lại.
  py -I kiem_boq.py <GOC>
"""
import collections
import json
import os
import re
import sys

import openpyxl
import pymupdf as fitz

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # chạy với py -I
from lap_boq import Du, ma_cua, ma_tag, so_to, kich_thuoc_do_roi


def main(goc):
    du = Du(goc)
    canh = json.load(open(os.path.join(du.d, "_lap_boq_dong.json"), encoding="utf-8"))
    tag_to = collections.defaultdict(collections.Counter)  # số tờ -> Counter mã
    for rel, v in du.tag.items():
        for ten, l in v["to"].items():
            so = so_to(ten)
            for t in l["tag"]:
                m = ma_tag(t) or ma_cua(t)
                if m:
                    tag_to[so][m] += 1
    p = json.load(open(os.path.join(du.d, "_dwg_tag_cabin.json"), encoding="utf-8"))
    for v in p.values():
        for ten, l in v["to"].items():
            for t in l["tag"]:
                m = ma_tag(t)
                if m:
                    tag_to[so_to(ten)][m] += 1
    loi = collections.defaultdict(list)
    co_ma = collections.defaultdict(set)
    for k, r in canh.items():
        ma, to = r["ma"], r["to"]
        co_ma[r["khu"]].add(ma)
        la_ma = bool(re.fullmatch(r"[A-Z]{2}-\d{2,4}|D\d+[a-z]?", ma or ""))
        if not to:
            loi["K2 Dòng không có nguồn tờ"].append(f"{k} {ma} ({r['ct']})")
        if la_ma and r["ct"] not in ("HW.DO",):
            khong = [x for x in to if not tag_to[x].get(ma) and not re.fullmatch(r"8\d{3}", x)]
            if to and len(khong) == len(to):
                loi["K1 Mọi tờ link đều không có mã của dòng"].append(f"{k} {ma}: {', '.join(to[:4])}")
            elif khong:
                loi["K1b Một số tờ link không có mã"].append(f"{k} {ma}: {', '.join(khong[:4])}")
    # K3
    da = set()
    for r in canh.values():
        da.update(r["to"])
    for so in sorted(da):
        fs = du.pdf_to([so]) or du.pdf_to([so + ".1"]) or du.pdf_to([so.split(".")[0]])
        if not fs:
            loi["K3 Tờ không có PDF lẻ"].append(so)
            continue
        t = fitz.open(fs[0])[0].get_text()
        if not re.search(rf"_{re.escape(so.split('.')[0])}", t) and so.split(".")[0] not in t:
            loi["K3 Khung tên PDF không có số tờ"].append(f"{so}: {os.path.basename(fs[0])}")
    # K4 sót vật liệu: chỉ xét tờ mặt đứng, WC, trần của từng khu
    luat = du.cfg["vat_lieu"]
    for (t, kh, lv), tos in du.khu.items():
        k_ten = f"{du.cfg['thap'][t]['ten']} — {du.cfg['khu'].get(kh, kh)}"
        for so, o in tos.items():
            if so.startswith("model:") or so[0] not in "123":
                continue
            for (c, bm) in o["bm"]:
                if (bm not in ("tuong", "tran", "portal", "can_soat", "khong_ro") and not (bm == "san" and c in (du.cfg.get("san_tinh") or {})) and not (bm == "khac" and c in (du.cfg.get("chi_tiet") or {}))) or "XX" in c or not (luat["ma"].get(c) or luat["ho"].get(c.split("-")[0])):
                    continue
                if c not in co_ma.get(k_ten, ()) and not (bm == "tran" and "CF." not in " ".join(du.cong_tac)):
                    loi["K4 Tag có trên tờ, BOQ không có dòng"].append(f"{k_ten}: {c} tờ {so}")
        # K6 đồ rời: đếm lại trên tờ FF&E
        dem = collections.Counter()
        for so, o in tos.items():
            if o["loai"] == "ffe":
                for c, n in o["tuong"].items():
                    if re.match(r"(BX|CG|LX|SX)-\d{4}", c):
                        dem[c] = max(dem[c], n)
        for k, r in canh.items():
            if r["khu"] == k_ten and r["ma"] in dem and r["sl"] != dem[r["ma"]]:
                loi["K6 Số lượng đồ rời khác số đếm lại trên tờ FF&E"].append(f"{k} {r['ma']}: BOQ {r['sl']}, tờ FF&E {dem[r['ma']]}")
    # K5 sót đồ rời
    bo = du.cfg["loose"]["bo"]
    dang_co = {r["ma"] for r in canh.values()}
    for (k, m), x in du.lf.items():
        if m[:2] not in bo and m not in dang_co:
            loi["K5 Có trong lịch FF&E, BOQ không có dòng"].append(f"{k} {m} {x.get('item desc', '')[:40]}")
    # K7 kích thước đọc lại từ file BOQ
    ra = max((os.path.join(dp, f) for dp, _, fs in os.walk(os.path.join(du.inp, "0.BOQ_TU_LAP")) for f in fs if f.endswith(".xlsx")), key=os.path.getmtime)
    wb = openpyxl.load_workbook(ra, data_only=True)
    thieu = collections.Counter()
    for k, r in canh.items():
        sh, dong = k.rsplit("!", 1)
        o = wb[sh]
        dai, rong = o.cell(row=int(dong), column=5).value, o.cell(row=int(dong), column=6).value
        if r["ct"].startswith(("LO.", "DE.LI")) or r["ct"] == "DO.WD" or r["ct"].startswith("FUR.OT"):
            if not dai:
                thieu[r["ct"]] += 1
        if r["ct"].startswith(("LO.", "DE.LI")):
            cac = [v for (kk, m), v in du.lf.items() if m == r["ma"]]
            if len(cac) > 1:
                loi["K8 Một mã FF&E có ở nhiều file LF"].append(f"{r['ma']}: " + ", ".join(v["file"][16:22] for v in cac))
            kts = [kich_thuoc_do_roi(v.get("dimensions"), v.get("item desc")) for v in cac]
            if kts and any(q[0] for q in kts) and not any(q[0] == dai and q[1] == rong for q in kts):
                loi["K7 Kích thước đồ rời khác spec"].append(f"{k} {r['ma']}: BOQ {dai}x{rong}, spec {kts[0][0]}x{kts[0][1]}")
    json.dump({"loi": loi, "thieu_kich_thuoc": thieu, "so_dong": len(canh)}, open(os.path.join(du.d, "_kiem_boq.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    print(f"Kiểm {len(canh)} dòng. Thiếu kích thước theo công tác: {dict(thieu)}")
    for k in sorted(loi):
        print(f"{k}: {len(loi[k])}")
        for x in loi[k][:12]:
            print("   ", x)
    if not loi:
        print("Không có lỗi nội dung ở các phép kiểm trên.")
    rows = [{"Bước": "doichieu", "Liên quan": k, "Nội dung": f"Kiểm máy (kiem_boq.py) thấy {len(v)} chỗ, ví dụ: " + "; ".join(v[:5]), "Cần ai trả lời": "Nội bộ"} for k, v in sorted(loi.items())]
    json.dump({"rows": rows}, open(os.path.join(du.d, "VIEC_MO.doichieu-2.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1])
