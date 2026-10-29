# -*- coding: utf-8 -*-
"""Plan NEO 3b-N2 D4, chỉ đọc: Mã VL trong SCOPE.banve-*.json (10-tender SOFITEL) có trên lớp text trang bản vẽ link không.
Đo kèm luật KHÔNG KIỂM ĐƯỢC mới: trang không có mã họ XX-NN nào / không có số 2–5 chữ số ngoài khung tên.
Chạy: py -I -X utf8 neo01_d4.py"""
import os, re, json, glob, zipfile, collections, unicodedata
import pymupdf

ROOT = r"D:\COSTING EXCEL\_ACTIVE\SOFITEL"
DATA = ROOT + r"\10_TENDER\_data"
INP = ROOT + r"\01. INPUT\new"
fi = {}
for f in glob.glob(DATA + r"\FILE_INDEX.*.json"):
    for x in json.load(open(f, encoding="utf-8"))["rows"]:
        if x.get("ID"): fi.setdefault(x["ID"], {}).update({k: v for k, v in x.items() if v not in (None, "")})

def fold(s):
    s = unicodedata.normalize("NFKC", str(s or ""))
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").upper()
CODE = re.compile(r"(?<![A-Z0-9])([A-Z]{2,3})\s*[-._]\s*0*(\d{1,3})(?![0-9])")  # mã VL: 2–3 chữ + gạch + số
def codes(s): return {a + b for a, b in CODE.findall(fold(s))}

_z, _d = {}, {}
def page(fid, pg):
    k = (fid, pg)
    if k in _d: return _d[k]
    r = fi.get(fid) or {}
    try:
        if r.get("Trong zip"):
            zp = os.path.join(INP, r["Đường dẫn"])
            if zp not in _z: _z[zp] = zipfile.ZipFile(zp)
            doc = pymupdf.open(stream=_z[zp].read(r["Trong zip"]), filetype="pdf")
        else:
            doc = pymupdf.open(os.path.join(INP, r["Đường dẫn"]))
        p = doc[pg - 1]; W, H = p.rect.width, p.rect.height
        words = p.get_text("words")
        body = " ".join(w[4] for w in words if w[0] < 0.80 * W)  # bỏ dải khung tên bên phải (GA Design: cột phải)
        _d[k] = (" ".join(w[4] for w in words), body)
    except Exception as e:
        _d[k] = None
    return _d[k]

rows = [r for f in glob.glob(DATA + r"\SCOPE.banve-*.json") for r in json.load(open(f, encoding="utf-8"))["rows"] if r.get("Mã VL")]
st = collections.Counter(); fam_kkd = collections.Counter(); ex = []
for r in rows:
    want = codes(r["Mã VL"])
    for lab, fid, pg in [l for l in (r.get("Nguồn") or []) + (r.get("Bản vẽ hiệu lực") or []) if isinstance(l, list) and len(l) >= 3]:
        t = page(fid, int(pg))
        if t is None: st["không mở được"] += 1; continue
        full, body = t
        pc = codes(full)
        st["trang"] += 1
        if not pc: st["trang KHÔNG có mã XX-NN nào → KKĐ"] += 1; continue
        lab_c = codes(lab) & want or want
        hit = lab_c & pc
        st["mã CÓ NEO"] += len(hit); st["mã LỆCH"] += len(lab_c - hit)
        if lab_c - hit and len(ex) < 8: ex.append((r["Khóa"], lab, sorted(lab_c - hit)))
nb = collections.Counter()
for k, v in _d.items():
    if not v: continue
    nb["trang"] += 1
    nb["có số 2–5 chữ số ngoài khung tên"] += bool(re.search(r"(?<![\d.])\d{2,5}(?![\d])", v[1]))
    nb["có mã XX-NN"] += bool(codes(v[0]))
print("D4 SCOPE Mã VL ↔ trang link:", dict(st))
for e in ex: print("  LỆCH", e)
print("Các trang đã mở:", dict(nb))
