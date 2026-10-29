# -*- coding: utf-8 -*-
"""Test nhanh SOFITEL: trích lớp text + đoán số/tên tờ của các PDF trong @DRAW vào thư mục tạm (không ghi vào dự án).
py -X utf8 sofitel_cache.py <thư mục ra>"""
import sys, os, re, json, glob
sys.path.insert(0, r"D:\COSTING EXCEL\SKILL\01-scan\tool")
import pd, pymupdf

INP = r"D:\COSTING EXCEL\_ACTIVE\SOFITEL\01. INPUT"
OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
tok = re.compile(r"\b[A-Z0-9]{1,6}(?:[.\-_][A-Z0-9]{1,6}){2,}\b")  # như cmd_kiemke
pdfs = sorted(glob.glob(os.path.join(INP, "@DRAW", "**", "*.pdf"), recursive=True), key=str.lower)
meta = []
for i, f in enumerate(pdfs, 1):
    fid, rel = f"S{i:02d}", os.path.relpath(f, INP)
    d = pymupdf.open(f); texts, titles = [], []
    for p in d:
        texts.append(p.get_text())
        try: titles.append(" ".join(pd.title_guess(p, tok)))
        except Exception as e: titles.append("")
    json.dump({"texts": texts, "titles": titles}, open(os.path.join(OUT, fid + ".json"), "w", encoding="utf-8"), ensure_ascii=False)
    meta.append({"ID": fid, "rel": rel, "pages": d.page_count, "size": os.path.getsize(f)})
    print(fid, d.page_count, rel, flush=True)
json.dump(meta, open(os.path.join(OUT, "_files.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
