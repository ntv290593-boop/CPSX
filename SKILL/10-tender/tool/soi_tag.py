"""Vẽ vòng đỏ quanh tag cần soát trên ảnh PDF của tờ, để LLM hoặc người xác nhận tag chỉ vào bề mặt nào.

Tọa độ giấy DWG (mm) đổi sang ảnh theo tỉ lệ 2.8346 pt/mm, gốc dưới trái, khổ giấy lấy từ trang PDF khi xem.
  py -I soi_tag.py <GOC> <số tờ> <file png> [mã ...] [--dpi 110]
"""
import json
import os
import re
import sys
import warnings

warnings.filterwarnings("ignore")
import pymupdf as fitz
from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # chạy với py -I
from lap_boq import Du, ma_tag, so_to

MM = 2.8346


def tag_cua_to(du, so):
    for rel, v in du.tag.items():
        for ten, l in v["to"].items():
            if so_to(ten) == so:
                yield rel, l


def ve(du, so, ra, ma_loc=None, dpi=110):
    fs = du.pdf_to([so]) or du.pdf_to([so + ".1"]) or du.pdf_to([so.split(".")[0]])
    if not fs:
        return None
    pg = fitz.open(fs[0])[0]
    pm = pg.get_pixmap(dpi=dpi)
    im = Image.frombytes("RGB", (pm.width, pm.height), pm.samples)
    dr = ImageDraw.Draw(im)
    k = dpi / 72 * MM
    h = pm.height
    dem = {}
    for rel, l in tag_cua_to(du, so):
        for t in l["tag"]:
            m = ma_tag(t)
            if not m or (ma_loc and m not in ma_loc) or not t.get("xy"):
                continue
            g = l.get("goc") or [0, 0, 1]  # gốc giấy của tờ có thể lệch, khung tên nằm ở góc trang
            x, y = (t["xy"][0] - g[0]) * k, h - (t["xy"][1] - g[1]) * k
            dr.ellipse([x - 22, y - 22, x + 22, y + 22], outline=(255, 0, 0), width=3)
            for u, v in t.get("mui", []):
                a, b = (u - g[0]) * k, h - (v - g[1]) * k
                dr.line([x, y, a, b], fill=(255, 0, 0), width=2)
                dr.ellipse([a - 7, b - 7, a + 7, b + 7], fill=(255, 0, 0))
            dr.text((x + 24, y - 24), m, fill=(255, 0, 0))
            dem[m] = dem.get(m, 0) + 1
    im.save(ra)
    return dem


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    a = sys.argv[1:]
    dpi = 110
    if "--dpi" in a:
        i = a.index("--dpi")
        dpi = int(a[i + 1])
        del a[i:i + 2]
    print(ve(Du(a[0]), a[1], a[2], set(a[3:]) or None, dpi))
