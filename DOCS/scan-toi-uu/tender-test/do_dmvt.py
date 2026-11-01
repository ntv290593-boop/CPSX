# -*- coding: utf-8 -*-
r"""do_dmvt.py — chấm DMVT máy bóc so với form user lập (plan mục 8, GĐ5). Đáp án chỉ dùng để chấm, máy không học từ đây.

Chạy: py -I do_dmvt.py [GOC] [--ra <thư mục>]   cần đã chạy `td.py dmvt <GOC> --chi-boc` hoặc build. --ra: ghi ảnh ghép dòng lệch để xem bằng mắt.
Đáp án: `01. INPUT\new\New folder\20261031-DMVT SOFITEL BASE (1).xlsx`, sheet DMVT.
Mốc ảnh (không nới sau khi thấy kết quả): hash 16x16, lệch <= 20 bit, >= 47/57 ảnh; ảnh không trúng phải có cờ.
Mốc 47 lập 31/10 là ảnh form có trong ứng viên máy cắt được (đo lại 01/11 bằng average hash ra đúng 47), nên in cả 2 số: có trong ứng viên của đúng dòng, và đúng ảnh máy chọn.
A:I chỉ báo số; cột C không tính là điều kiện qua. So chữ theo tập từ đã chuẩn hóa vì form có sửa tay cách viết.
"""
import io
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", ".."))
sys.path.insert(0, os.path.join(ROOT, "SKILL", "10-tender", "tool"))
import td
import dmvt
import openpyxl
from PIL import Image

NS = {"x": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def anh_form(f):
    """Hàng -> bytes ảnh, đọc thẳng drawing trong zip."""
    out = {}
    with zipfile.ZipFile(f) as z:
        rel = ET.fromstring(z.read("xl/drawings/_rels/drawing1.xml.rels"))
        rid = {r.get("Id"): r.get("Target") for r in rel}
        for a in ET.fromstring(z.read("xl/drawings/drawing1.xml")):
            hang = a.find("x:from/x:row", NS)
            blip = a.find(".//a:blip", NS)
            if hang is None or blip is None:
                continue
            t = rid[blip.get("{%s}embed" % NS["r"])].replace("../", "xl/")
            out[int(hang.text) + 1] = z.read(t)
    return out


def giong_diem_anh(a, b):
    """Khác trung bình <= 3% ở 64x64: cùng 1 ảnh, chỉ khác độ phân giải."""
    import numpy as np
    x, y = [np.asarray(i.convert("RGB").resize((64, 64), Image.LANCZOS), dtype=float) for i in (a, b)]
    return float(np.abs(x - y).mean()) <= 255 * 0.03


def tu(t):
    return set(re.findall(r"[a-z0-9%]+", str(t or "").lower()))


def phu(may, form):
    """Phần trăm từ của form có trong máy."""
    f = tu(form)
    return len(f & tu(may)) / len(f) if f else 1.0


def main():
    a = sys.argv[1:]
    ra = td.opt(a, "--ra")
    goc = a[0] if a else os.path.join(ROOT, "_ACTIVE", "SOFITEL")
    p = td.P(goc).need()
    rows = td.jload(p.f("_dmvt_boc.json"))["rows"]
    f = os.path.join(p.inp, "New folder", "20261031-DMVT SOFITEL BASE (1).xlsx")
    ws = openpyxl.load_workbook(f, data_only=True)["DMVT"]
    anh = anh_form(f)
    may_ma = {}
    for r in rows:
        for m in r["Mã"]:
            may_ma[m] = r
    n = tap_khop = 0
    khong_co, lech_tap, diem = [], [], {"D": 0, "E": 0, "F": 0, "G": 0}
    n_anh = trung = co_ko_trung = co_ung = giong = 0
    ko_trung = []
    n_c = c_khop = 0
    for h in range(8, ws.max_row + 1):
        b = ws.cell(h, 2).value
        if not b or not ws.cell(h, 4).value:
            continue
        ma = list(dict.fromkeys(x[0] for x in td.tach_ma(str(b)) if x[3] in ("ffe", "gach")))
        if not ma:
            continue
        n += 1
        cac = {id(may_ma[m]): may_ma[m] for m in ma if m in may_ma}
        if not cac:
            khong_co.append((h, b))
            continue
        r = max(cac.values(), key=lambda x: len(set(x["Mã"]) & set(ma)))
        if set(r["Mã"]) == set(ma):
            tap_khop += 1
        else:
            lech_tap.append((h, " / ".join(ma), " | ".join(x["Khóa"] for x in cac.values())))
        for c, cot, k in ((4, "D", "Material"), (5, "E", "Location"), (6, "F", "Description"), (7, "G", "NCC")):
            diem[c and cot] += phu(r.get(k), ws.cell(h, c).value)
        if ws.cell(h, 3).value:
            n_c += 1
        if h in anh:
            n_anh += 1
            im_f = Image.open(io.BytesIO(anh[h])).convert("RGB")
            hf = dmvt.phash(im_f)
            duyet = r.get("Ảnh") and os.path.join(p.f("_dmvt_anh"), r["Ảnh"])
            d = dmvt.lech(hf, int(r["Hash ảnh"], 16)) if r.get("Hash ảnh") else 999
            if min([dmvt.lech(hf, int(x, 16)) for x in r.get("Hash ứng viên", [])] or [999]) <= dmvt.HASH_LECH:
                co_ung += 1
            if d <= dmvt.HASH_LECH:
                trung += 1
            else:
                co = bool(r.get("Cờ ảnh"))
                co_ko_trung += co
                # hash của ảnh trơn chủ yếu là nhiễu: so thêm điểm ảnh, chỉ để phân loại lý do, không cộng vào số trúng
                if duyet and giong_diem_anh(im_f, Image.open(duyet)):
                    giong += 1
                ko_trung.append((h, str(b).replace("\n", " ")[:30], d, r.get("Số ứng viên"), r.get("Cờ ảnh")))
                if ra and duyet:
                    os.makedirs(ra, exist_ok=True)
                    tam = Image.new("RGB", (620, 300), "white")
                    x, y = im_f.copy(), Image.open(duyet).convert("RGB")
                    x.thumbnail((300, 300))
                    y.thumbnail((300, 300))
                    tam.paste(x, (0, 0))
                    tam.paste(y, (320, 0))
                    tam.save(os.path.join(ra, f"lech_{h}.png"))
    print(f"Dòng form có mã: {n}. Có dòng máy: {n - len(khong_co)}. Khớp đúng tập mã: {tap_khop}.")
    print("Phủ chữ trung bình (từ của form có trong máy):", {k: f"{v / max(1, n - len(khong_co)):.0%}" for k, v in diem.items()})
    print(f"Ảnh: ảnh form có trong ứng viên của đúng dòng {co_ung}/{n_anh} (mốc >= 47/57, lệch <= {dmvt.HASH_LECH} bit). Đúng ảnh máy chọn {trung}/{n_anh}. Không trúng {len(ko_trung)}, có cờ {co_ko_trung}, trong đó giống từng điểm ảnh {giong} (hash nhiễu do ảnh trơn).")
    for x in ko_trung:
        print("  ảnh lệch hàng", *x)
    print("Không có dòng máy:", khong_co)
    print("Lệch tập mã:")
    for x in lech_tap:
        print("  hàng", *x)
    form_ma = set()
    for h in range(8, ws.max_row + 1):
        form_ma |= {x[0] for x in td.tach_ma(str(ws.cell(h, 2).value or "")) if x[3] in ("ffe", "gach")}
    thua = [r["Khóa"] for r in rows if not (set(r["Mã"]) & form_ma) and r["Mục"] not in ("VII", "VIII")]
    print(f"Dòng máy thêm ngoài form (trừ Finishes, Sanitary): {len(thua)}: {thua}")


if __name__ == "__main__":
    main()
