"""Chuẩn bị soát bề mặt tag bằng hình: mỗi tờ 1 ảnh, tag cần soát khoanh đỏ, đánh số. Chia nhóm cho người/LLM soát.
  py -I chuan_bi_soat.py <GOC> [--nhom 25] [--dpi 100]
Ghi _data\\_soat\\<số tờ>.png, _data\\_soat\\ds.json. Ai soát ghi _data\\_soat_tag.<n>.json: {"rows": [{"khoa": "<tờ>#<id>", "be_mat": "...", "ghi_chu": "..."}]}.
"""
import collections
import json
import os
import re
import sys
import warnings

warnings.filterwarnings("ignore")
import pymupdf as fitz
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # chạy với py -I
from be_mat import doan, nap_soat
from lap_boq import Du, ma_tag, so_to

MM = 2.8346


def ma_can(du, m):
    v = du.cfg["vat_lieu"]
    return "XX" not in m and bool(v["ma"].get(m) or v["ho"].get(m.split("-")[0]))


def main(a):
    du = Du(a[0])
    nhom = int(a[a.index("--nhom") + 1]) if "--nhom" in a else 25
    dpi = int(a[a.index("--dpi") + 1]) if "--dpi" in a else 100
    d = os.path.join(du.d, "_soat")
    os.makedirs(d, exist_ok=True)
    rx = re.compile(du.cfg["rx_ten_dwg"])
    da = nap_soat(du.d)
    ds, thong_ke = {}, collections.Counter()
    for rel, v in du.tag.items():
        if not rx.search(os.path.basename(rel)):
            continue
        for ten, l in v["to"].items():
            so = so_to(ten)
            if so[0] not in "123":
                continue
            for i, t in enumerate(l["tag"]):
                m = ma_tag(t)
                if not m or not ma_can(du, m):
                    continue
                bm, ly_do = doan(t, l)
                thong_ke[bm] += 1
                if bm == "can_soat" and f"{so}#{i}" not in da:
                    ds.setdefault(so, []).append({"id": i, "ma": m, "xy": t.get("xy"), "mui": t.get("mui") or [], "ly_do": ly_do, "goc": l.get("goc") or [0, 0, 1]})
    print("Phân loại máy (số tag):", dict(thong_ke))
    print(f"Cần soát: {sum(len(x) for x in ds.values())} tag trên {len(ds)} tờ")
    try:
        font = ImageFont.truetype("arialbd.ttf", 22)
    except OSError:
        font = None
    khong = []
    for so, tags in ds.items():
        fs = du.pdf_to([so])
        if not fs:
            khong.append(so)
            continue
        pm = fitz.open(fs[0])[0].get_pixmap(dpi=dpi)
        im = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
        dr = ImageDraw.Draw(im)
        k, h = dpi / 72 * MM, pm.height
        for t in tags:
            if not t["xy"]:
                continue
            gx, gy = t["goc"][0], t["goc"][1]  # gốc giấy của tờ có thể lệch
            x, y = (t["xy"][0] - gx) * k, h - (t["xy"][1] - gy) * k
            dr.ellipse([x - 26, y - 26, x + 26, y + 26], outline=(255, 0, 0), width=3)
            for u, v in t["mui"]:
                a2, b2 = (u - gx) * k, h - (v - gy) * k
                dr.line([x, y, a2, b2], fill=(255, 0, 0), width=2)
                dr.ellipse([a2 - 7, b2 - 7, a2 + 7, b2 + 7], fill=(255, 0, 0))
            dr.text((x + 28, y - 30), f"#{t['id']} {t['ma']}", fill=(255, 0, 0), font=font)
        im.save(os.path.join(d, f"{so}.png"))
    cac_to = sorted(s for s in ds if s not in khong)
    chia = [cac_to[i:i + nhom] for i in range(0, len(cac_to), nhom)]
    json.dump({"to": {s: ds[s] for s in cac_to}, "nhom": chia, "khong_pdf": khong}, open(os.path.join(d, "ds.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    print("Số nhóm:", len(chia), "| tờ không có PDF:", khong)
    print({i + 1: (c[0], c[-1], sum(len(ds[s]) for s in c)) for i, c in enumerate(chia)})


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1:])
