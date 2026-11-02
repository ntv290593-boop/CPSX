"""Bóc spec dạng khối 'Item Desc / Area / Prod. Code' (mẫu spec AID: FF&E, Finishes, Sanitary, Ironmongery) thành danh sách mã.

Ghép mã với khối theo vị trí trên trang schedule: mã ở cột trái, thuộc dòng 'Item Desc' gần nhất phía trên.
Trang chi tiết (mỗi trang 1 mã) bổ sung Dimensions, Finishes, Content, Manufacturer.
Dùng cho dự án có spec cùng kiểu; dự án khác kiểu thì không gọi file này.
  py -I spec_boc.py <file json ra> <pdf> [<pdf> ...]
"""
import json
import os
import re
import sys
import warnings

warnings.filterwarnings("ignore")
import pymupdf as fitz

RX_MA = re.compile(r"^[A-Z]{2}-?\s?\d{2,4}(?:\.\d{1,2})?[A-Z]?$")
TRUONG = ["Item Description", "Item Desc", "Brand", "Material", "Area Used", "Area", "Prod. Code", "Remarks", "Custom finishes", "Finishes", "Content", "Dimensions",
          "Manufacturer", "Supplier", "Contact", "Address", "General Note", "Indicative Lead Time", "Indicative lead time"]
RX_T = re.compile(r"(?:^|\n)\s*(" + "|".join(re.escape(x) for x in TRUONG) + r")\s*:", re.I)
CHAN = re.compile(r"\n\s*(?:https?://|www\.|[TMCE]:\s|Page \d)", re.I)


def gon(t):
    return re.sub(r"\s+", " ", t or "").strip()


def tach_truong(t):
    out = {}
    ms = list(RX_T.finditer(t))
    for i, m in enumerate(ms):
        k = m.group(1).lower()
        k = {"item description": "item desc", "area used": "area", "custom finishes": "finishes", "indicative lead time": "lead"}.get(k, k)
        v = t[m.end():ms[i + 1].start() if i + 1 < len(ms) else len(t)]
        c = CHAN.search(v)
        out.setdefault(k, gon(v[:c.start()] if c else v))
    return out


def ma_cot_trai(pg, x_max):
    """Mã ở cột trái: có thể xuống dòng 'SX-' / '2000.2'. Trả [(mã, y)]."""
    ws = sorted((w for w in pg.get_text("words") if w[0] < x_max), key=lambda w: (w[1], w[0]))
    out, i = [], 0
    while i < len(ws):
        t = ws[i][4]
        if re.fullmatch(r"[A-Z]{2}-", t) and i + 1 < len(ws) and re.fullmatch(r"\d{2,4}(\.\d{1,2})?[A-Z]?", ws[i + 1][4]):
            out.append((t + ws[i + 1][4], ws[i][1]))
            i += 2
            continue
        if re.fullmatch(r"[A-Z]{2}-\d{2,4}(\.\d{1,2})?[A-Z]?", t):
            out.append((t, ws[i][1]))
        i += 1
    return out


def boc_schedule(pg):
    """Các khối của 1 trang schedule: [{ma, item desc, area, ...}]. Khối không thấy mã thì ma=None."""
    t = pg.get_text("blocks")
    bien = sorted(b[1] for b in t if "Item Desc" in b[4])
    if not bien:
        return []
    ma = ma_cot_trai(pg, 80)
    khoi = []
    for j, y in enumerate(bien):
        y2 = bien[j + 1] if j + 1 < len(bien) else 1e9
        chu = "\n".join(b[4] for b in sorted(t, key=lambda b: (b[1], b[0])) if y - 0.5 <= b[1] < y2 and 80 <= b[0] < 255)  # từ x 255 là cột nhà cung cấp
        d = tach_truong(chu)
        ds = [m for m, ym in ma if y - 0.5 <= ym < y2]
        d["ma"] = ds[0] if ds else None
        d["y_ma"] = ds and next(ym for m, ym in ma if m == ds[0])
        if len(ds) > 1:
            d["ma_khac"] = ds[1:]
        khoi.append(d)
    # dải của từng mã trên trang: chia giữa vị trí 2 mã liền nhau, để cắt ảnh đúng mã
    co = sorted((k for k in khoi if k.get("y_ma")), key=lambda k: k["y_ma"])
    for i, k in enumerate(co):
        y = k["y_ma"]
        tren = (co[i - 1]["y_ma"] + y) / 2 if i else None
        duoi = (co[i + 1]["y_ma"] + y) / 2 if i + 1 < len(co) else None
        buoc = (duoi - y) if duoi else ((y - tren) if tren else None)
        k["y0"] = round(tren if tren else (y - buoc if buoc else 0), 1)
        k["y1"] = round(duoi if duoi else (y + buoc if buoc else pg.rect.height), 1)
    return khoi


def boc_file(path):
    d = fitz.open(path)
    ten = os.path.basename(path)
    nhom = None
    ds, khong_ma = [], []
    tieu_de = gon(d[0].get_text())[:200]
    for i, pg in enumerate(d):
        t = pg.get_text()
        if "SCHEDULE" in t[:400] and "Item Desc" in t:
            m = re.search(r"([A-Z&, /-]{4,})\s*\(([A-Z]{2})\)", t)
            for k in boc_schedule(pg):
                k.update(file=ten, trang=i + 1, nhom=(m.group(2) if m else None))
                (ds if k["ma"] else khong_ma).append(k)
    bang = {k["ma"]: k for k in ds}
    ho_schedule = {m[:2] for m in bang}
    # trang chi tiết: 1 mã/trang, bổ sung trường và gán mã cho khối thiếu
    for i, pg in enumerate(d):
        t = pg.get_text()
        if "SPECIFICATION" not in t[:400] or "Item Desc" not in t:
            continue
        ma = [w[4] for w in pg.get_text("words") if re.fullmatch(r"[A-Z]{2}-\d{2,4}(\.\d{1,2})?[A-Z]?", w[4]) and w[0] > 400]  # trang chi tiết: mã ở cột phải
        ct = tach_truong(t)
        for m in ma[:1]:
            k = bang.get(m)
            if k is None:
                # tìm khối thiếu mã cùng Item Desc
                for x in khong_ma:
                    if gon(x.get("item desc"))[:30] == gon(ct.get("item desc"))[:30]:
                        k = x
                        k["ma"] = m
                        khong_ma.remove(x)
                        ds.append(k)
                        bang[m] = k
                        break
            if k is None:
                if m[:2] not in ho_schedule:  # mã vật liệu nhắc trong trang đồ rời (WD-05, MT-02...), không phải mã của file này
                    continue
                k = dict(ct, ma=m, file=ten, trang=i + 1, nhom=m[:2], tu_trang_ct=True)
                ds.append(k)
                bang[m] = k
            for f, v in ct.items():
                if v and not k.get(f):
                    k[f] = v
            k["trang_ct"] = i + 1
    return {"file": ten, "so_trang": d.page_count, "ma": ds, "khong_ma": khong_ma, "bia": tieu_de}


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    kq = [boc_file(p) for p in sys.argv[2:]]
    json.dump(kq, open(sys.argv[1], "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    for k in kq:
        print(f"{k['file']}: {len(k['ma'])} mã, {len(k['khong_ma'])} khối thiếu mã, {k['so_trang']} trang")
