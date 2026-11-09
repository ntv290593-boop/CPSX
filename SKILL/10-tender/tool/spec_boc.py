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
TRUONG = ["Item Description", "Item Desc", "Brand", "Material", "Area Used", "Area", "Prod. Code", "Product Code", "Location", "Colour", "Remarks", "Custom finishes", "Finishes", "Content", "Dimensions",
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
        k = {"item description": "item desc", "area used": "area", "custom finishes": "finishes", "indicative lead time": "lead", "product code": "prod. code"}.get(k, k)
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
    # vị trí cột đo theo trang, không cố định: file FF&E chữ ở x 80, NCC từ x 255; file SW chữ ở x 73, NCC từ x 197
    x_chu = min(b[0] for b in t if "Item Desc" in b[4])
    W = pg.rect.width
    xs_ncc = [b[0] for b in t if x_chu + 90 <= b[0] < W * 0.62 and b[1] >= bien[0] - 0.5 and b[4].strip() not in ("-", "")]
    x_ncc = min(xs_ncc) if xs_ncc else 255
    ma = ma_cot_trai(pg, x_chu - 2)
    khoi = []
    for j, y in enumerate(bien):
        y2 = bien[j + 1] if j + 1 < len(bien) else min((b[1] for b in t if "PLEASE REFER" in b[4] or re.match(r"Page \d", b[4])), default=1e9)  # khối cuối: dừng ở chân trang
        trong = [b for b in sorted(t, key=lambda b: (b[1], b[0])) if y - 0.5 <= b[1] < y2]
        chu = "\n".join(b[4] for b in trong if x_chu - 2 <= b[0] < x_ncc)
        d = tach_truong(chu)
        nc = doc_ncc([x for b in trong if x_ncc <= b[0] < W * 0.62 for x in b[4].split("\n")])
        if nc and not (d.get("manufacturer") or d.get("supplier") or d.get("brand")):
            d["manufacturer"] = nc["manufacturer"]
            if nc["contact"] or nc["web"]:
                d.setdefault("contact", "\n".join(x for x in (nc["contact"], nc["web"]) if x))
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


RX_LH = re.compile(r"^\s*(?:[TMCEF]|Tel|Mob|Email|Fax)\s*:", re.I)
RX_WEB = re.compile(r"https?://|www\.", re.I)


def ncc_cot_phai(pg):
    """Khối NCC ở cột phải trang chi tiết thường KHÔNG có nhãn: tên NCC, web, T:/E:/M:/C:, 'Indicative lead time'.
    Đọc theo vị trí (x > 55% bề ngang, dưới tiêu đề 'SPECIFICATION', trên ô CODE). Trả {manufacturer, web, contact} hoặc {}."""
    W = pg.rect.width
    bl = pg.get_text("blocks")
    y_tieu = min((b[3] for b in bl if "SPECIFICATION" in b[4]), default=0)
    y_code = min((b[1] for b in bl if b[4].strip() == "CODE" and b[0] > W * 0.6), default=pg.rect.height * 0.75)
    ds = sorted((b for b in bl if b[0] > W * 0.55 and y_tieu <= b[1] < y_code), key=lambda b: (b[1], b[0]))
    return doc_ncc([x for b in ds for x in b[4].split("\n")])


def doc_ncc(dong_s):
    """Các dòng của khối NCC không nhãn -> {manufacturer: tên, web, contact}. Gặp nhãn trường khác (lead time…) là hết khối."""
    ten, web, lh = [], [], []
    for dong in (x.strip() for x in dong_s):
        if not dong:
            continue
        if RX_T.match(dong) or re.match(r"(?i)indicative|lead\s*time|page \d", dong):
            if re.match(r"(?i)(manufacturer|supplier|brand)\s*:", dong):  # có nhãn: để tach_truong lo
                return {}
            break
        m = re.search(r"(?i)\burl\s*:", dong)
        if m and m.start() > 0 and not web and not lh:  # 'RIGEL or Equivalent URL: ...' cùng dòng
            ten.append(dong[:m.start()].strip())
            dong = dong[m.start():]
        if RX_WEB.search(dong) or re.match(r"(?i)url\s*:", dong):
            web.append(re.sub(r"(?i)^url\s*:\s*", "", dong))
        elif RX_LH.match(dong) or "@" in dong or (ten and re.match(r"\d{1,5}[A-Z]?\s+\w", dong)):  # địa chỉ '7030 Ang Mo Kio' là liên hệ
            lh.append(dong)
        elif re.match(r"(?i)technical drawings|all dimensions", dong):
            break  # chữ của khung bản vẽ kỹ thuật, không phải NCC
        elif not web and not lh:
            dong = dong.lstrip(": ").strip()
            if dong and dong not in ten:  # 'Contractor to Bid' lặp ở 2 khối cùng dải
                ten.append(dong)
        elif web and not lh and not dong[:1].isupper():
            web[-1] += dong  # URL xuống dòng
    if not ten:
        return {}
    return {"manufacturer": gon(" ".join(ten[:2])), "web": " ".join(web), "contact": "\n".join(lh)}


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
        nc = ncc_cot_phai(pg)
        if nc and not ct.get("supplier") and gon(nc["manufacturer"]) != gon(ct.get("manufacturer")):
            ten_ncc = nc["manufacturer"]
            for f in list(ct):  # tên NCC không nhãn bị tach_truong dán vào cuối trường đứng trước (General Note…St Vincents)
                if ct[f].endswith(ten_ncc):
                    ct[f] = ct[f][: -len(ten_ncc)].strip()
            ct["supplier" if ct.get("manufacturer") else "manufacturer"] = ten_ncc  # có nhãn Manufacturer: khối cột phải là bên cung cấp
            if nc["contact"] or nc["web"]:
                ct.setdefault("contact", "\n".join(x for x in (nc["contact"], nc["web"]) if x))
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
