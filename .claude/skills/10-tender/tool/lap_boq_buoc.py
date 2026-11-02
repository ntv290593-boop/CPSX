"""Sau lap_boq.py và các lệnh tool boq, dwg, mat: ghi JSON các bước mat, dwg, banve, doichieu, phanloai từ chính dữ liệu đã bóc.

Dùng cho dự án tự lập BOQ (mẫu tu_lap_vn): BOQ do máy lập nên nguồn và cờ của từng dòng đã có sẵn ở cột Ghi chú.
Thuộc kiểu dữ liệu của lap_boq.py; không suy thêm số liệu.
  py -I lap_boq_buoc.py <GOC>
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # chạy với py -I
from lap_boq import Du

LOAI = {"HD": "Khác", "WD": "Gỗ/Veneer", "ST": "Đá", "CT": "Gạch/Mosaic", "PT": "Sơn", "UP": "Vải/Da", "CP": "Thảm", "WC": "Giấy/Vải dán tường",
        "GL": "Kính/Gương", "MT": "Kim loại", "LM": "Laminate", "SP": "Khác", "AC": "Khác", "VF": "Khác"}


def nap(d, f):
    return json.load(open(os.path.join(d, f), encoding="utf-8"))["rows"]


def ghi(d, f, rows):
    json.dump({"rows": rows}, open(os.path.join(d, f), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    print(f"{f}: {len(rows)} dòng")


def main(goc):
    du = Du(goc)
    d = du.d
    scope = nap(d, "SCOPE.boq-tool.json")
    dwg = nap(d, "DWG.dwg-tool.json")
    fidx = nap(d, "FILE_INDEX.kiemke-tool.json")
    fid_ten = {os.path.basename(r["Đường dẫn"]): r["ID"] for r in fidx}

    # Tên tờ, sổ bản vẽ
    rows = []
    for r in dwg:
        v = du.dl.get(r["Số BV"]) or du.dl.get(r["Số BV"] + ".1")
        x = {"Số BV": r["Số BV"]}
        if v:
            x["Tên tờ"] = " · ".join(a for a in v[5:8] if a not in ("AS SHOWN", "N.A.", "0"))[:150]
        else:
            x["Ghi chú"] = "Không thấy trong Drawing List"
        if v and "RENDER" in " ".join(v[5:8]).upper():
            x["Trạng thái"] = "Tham khảo"
        rows.append(x)
    ghi(d, "DWG.dwg.json", rows)

    # Khu (thư mục): dự án không có mã phòng, ghi tên khu
    khu = sorted({r.get("Khu (thư mục)") for r in dwg if r.get("Khu (thư mục)")})
    ghi(d, "KHU.dwg.json", [{"Khu": k, "Bộ hồ sơ": "Schedule" if "Specification" in k or "Door" in k else "IDD",
                             "Mã phòng": "Toàn dự án" if not re.match(r"L1|L6|Typical", k) else k,
                             "Ghi chú": "Thư mục bản vẽ ID; khu cụ thể lấy từ tên file (tháp_tầng_khu)"} for k in khu])

    # MAT: vật liệu tra FM, bộ phụ kiện cửa tra IM (lấy khối đầu của mỗi mã, là trang schedule)
    im = {}
    for f in du.spec:
        if "_IM_" in f["file"]:
            for k in f["ma"]:
                if k.get("item desc"):
                    im.setdefault(k["ma"], k)
    rows = []
    for r in nap(d, "MAT.mat-tool.json"):
        m = r["Mã VL"]
        x = du.fm.get(m) or im.get(m)
        o = {"Mã VL": m}
        if m.split("-")[0] in LOAI:
            o["Loại"] = LOAI[m.split("-")[0]]
        if x:
            o["Mô tả"] = (x.get("item desc") or m)[:200]
            ncc = " · ".join(a for a in (x.get("manufacturer"), x.get("prod. code")) if a)
            if ncc:
                o["NCC/Mã SP"] = ncc[:150]
            o["Nguồn spec"] = [[m, fid_ten[x["file"]], x.get("trang_ct") or x["trang"]]]
            chu = (x.get("area", "") + " " + x.get("remarks", "")).upper()
            o["Trạng thái cuối"] = "Omit" if "OMIT" in chu else "Theo spec"
            if x.get("area"):
                o["Nội dung cuối (VN)"] = ("Spec ghi TO BE OMITTED. " if "OMIT" in chu else "") + "Khu dùng theo spec: " + x["area"][:150]
        else:
            o["Mô tả"] = "Không thấy mã trong FM_7100, IM_7400"
            o["Trạng thái cuối"] = "Không thấy spec"
        rows.append(o)
    # Đồ rời đưa vào MAT để có Link vật liệu và dòng DMVT
    dung_ma = {re.sub(r"^(\S+).*", r"\1", r["Mô tả BOQ"]) for r in scope}
    for (k, m), x in sorted(du.lf.items(), key=lambda a: a[0][1]):
        if m not in dung_ma:
            continue
        rows.append({"Mã VL": m, "Loại": "Khác", "Mô tả": (x.get("item desc") or m)[:200],
                     "NCC/Mã SP": " · ".join(a for a in (x.get("manufacturer"), x.get("prod. code")) if a)[:150],
                     "Nguồn spec": [[m, fid_ten[x["file"]], x.get("trang_ct") or x["trang"]]], "Trạng thái cuối": "Theo spec",
                     "Nội dung cuối (VN)": ("Khu dùng theo spec: " + x["area"][:150]) if x.get("area") else ""})
    rows = [{k: v for k, v in r.items() if v != ""} for r in rows]
    ghi(d, "MAT.mat.json", rows)
    omit = {r["Mã VL"] for r in rows if r["Trạng thái cuối"] == "Omit"}

    # Bản vẽ hiệu lực của từng dòng, từ số tờ đã ghi ở cột Ghi chú
    bv = {r["Số BV"]: r["File PDF"][0][1] for r in dwg if r.get("File PDF")}
    banve, doichieu, llm = [], [], []
    canh = json.load(open(os.path.join(d, "_lap_boq_dong.json"), encoding="utf-8"))
    bm = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "..", "_MASTER", "code_id", "bang_ma.json"), encoding="utf-8"))
    cap3 = {x["ma"]: x["cap3"] for x in (bm.get("rows", bm) if isinstance(bm, dict) else bm)}  # tên trong bảng mã chung, tra theo mã công tác
    for r in scope:
        gc = r.get("Ghi chú BOQ", "")
        mt = r["Mô tả BOQ"]
        ma = mt.split(" ", 1)[0]
        lk = []
        for x in canh.get(r["Khóa"], {}).get("to", [])[:30]:  # số tờ do lap_boq.py ghi lại, không đọc lại từ chữ ghi chú
            f = bv.get(x) or bv.get(x + ".1") or bv.get(x.split(".")[0])
            if f:
                lk.append([x, f, 1])
        b = {"Khóa": r["Khóa"]}
        if lk:
            b["Bản vẽ hiệu lực"] = lk
        if re.fullmatch(r"(AX|BX|CG|CX|DX|LX|SX)-\d{4}", ma):
            b["Mã spec/FF&E"] = ma
        if len(b) > 1:
            banve.append(b)
        o = {"Khóa": r["Khóa"], "Trạng thái": "Base"}
        xem = re.findall(r"\[XEM\] ([^\[]*)", gc)
        if ma in omit:
            o["Trạng thái"] = "Omit"
        elif r.get("Nhóm", "").startswith("Wealth — Hành lang"):
            o["Trạng thái"] = "Chờ"
            xem.insert(0, "Hành lang Tháp 1: Kế hoạch ô G8 ghi fitout văn phòng tháp 1 thuộc SOL gói 2, chờ xác nhận")
        if xem:
            o["Việc còn mở"] = " ".join(x.strip() for x in xem)[:400]
        kt = [canh.get(r["Khóa"], {}).get(c) for c in ("dai", "rong")]  # BOQ tự lập có sẵn kích thước ở cột E, F
        if any(kt):
            o["Kích thước BoQ"] = " x ".join(str(x) for x in kt if x)
        doichieu.append(o)
        llm.append({"Khóa": r["Khóa"], "Mã cấp 3": cap3[r["Vùng"].split(" ", 1)[0]]})
    ghi(d, "SCOPE.banve.json", banve)
    ghi(d, "SCOPE.doichieu.json", doichieu)
    json.dump({"rows": llm}, open(os.path.join(d, "_phanloai_llm.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    print("_phanloai_llm.json:", len(llm), "dòng (Mã cấp 3 = công tác của mục lớn trong BOQ)")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1])
