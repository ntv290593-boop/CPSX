"""Lấy mẫu ngẫu nhiên tag đã soát để người/LLM khác soát lại độc lập, đo độ tin cậy của lần soát trước.
  py -I mau_soat.py <GOC> tao [--so 80] [--dpi 100]   ảnh _data\\_soat2\\<tờ>.png, ds.json (không ghi kết quả lần trước)
  py -I mau_soat.py <GOC> so                          so _soat2\\kq.json với _soat_tag.*.json
"""
import json
import os
import random
import sys
import warnings

warnings.filterwarnings("ignore")
import pymupdf as fitz
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from be_mat import nap_soat
from lap_boq import Du, ma_tag, so_to

MM = 2.8346


def tao(du, so_mau, dpi):
    d = os.path.join(du.d, "_soat2")
    os.makedirs(d, exist_ok=True)
    soat = nap_soat(du.d)
    # lấy đều: mọi tag không phải tường, cộng mẫu ngẫu nhiên tag tường
    khac = [k for k, r in soat.items() if r["be_mat"] != "tuong"]
    tuong = [k for k, r in soat.items() if r["be_mat"] == "tuong"]
    random.seed(20261101)
    chon = set(khac) | set(random.sample(tuong, min(so_mau, len(tuong))))
    theo_to = {}
    for k in chon:
        so, i = k.split("#")
        theo_to.setdefault(so, []).append(int(i))
    try:
        font = ImageFont.truetype("arialbd.ttf", 22)
    except OSError:
        font = None
    ds = {}
    for rel, v in du.tag.items():
        for ten, l in v["to"].items():
            so = so_to(ten)
            if so not in theo_to:
                continue
            fs = du.pdf_to([so])
            if not fs:
                continue
            pm = fitz.open(fs[0])[0].get_pixmap(dpi=dpi)
            im = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
            dr = ImageDraw.Draw(im)
            k, h = dpi / 72 * MM, pm.height
            g = l.get("goc") or [0, 0, 1]
            for i in theo_to[so]:
                t = l["tag"][i]
                if not t.get("xy"):
                    continue
                x, y = (t["xy"][0] - g[0]) * k, h - (t["xy"][1] - g[1]) * k
                dr.ellipse([x - 26, y - 26, x + 26, y + 26], outline=(255, 0, 0), width=3)
                for a, b in t.get("mui") or []:
                    a2, b2 = (a - g[0]) * k, h - (b - g[1]) * k
                    dr.line([x, y, a2, b2], fill=(255, 0, 0), width=2)
                    dr.ellipse([a2 - 7, b2 - 7, a2 + 7, b2 + 7], fill=(255, 0, 0))
                dr.text((x + 28, y - 30), f"#{i} {ma_tag(t)}", fill=(255, 0, 0), font=font)
                ds.setdefault(so, []).append({"id": i, "ma": ma_tag(t)})
            im.save(os.path.join(d, f"{so}.png"))
    json.dump(ds, open(os.path.join(d, "ds.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    print(f"Mẫu: {sum(len(v) for v in ds.values())} tag, {len(ds)} tờ ({len(khac)} tag không phải tường + {so_mau} tường ngẫu nhiên)")


def so(du):
    soat = nap_soat(du.d)
    kq = json.load(open(os.path.join(du.d, "_soat2", "kq.json"), encoding="utf-8"))["rows"]
    khac = [(r["khoa"], r.get("ma"), soat[r["khoa"]]["be_mat"], r["be_mat"], r.get("ghi_chu", "")) for r in kq if soat[r["khoa"]]["be_mat"] != r["be_mat"]]
    print(f"So {len(kq)} tag: khác {len(khac)} ({100 * len(khac) / max(len(kq), 1):.1f}%)")
    for x in khac:
        print(" ", x)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    a = sys.argv[1:]
    du = Du(a[0])
    if a[1] == "tao":
        tao(du, int(a[a.index("--so") + 1]) if "--so" in a else 80, int(a[a.index("--dpi") + 1]) if "--dpi" in a else 100)
    else:
        so(du)
