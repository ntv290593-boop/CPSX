# -*- coding: utf-8 -*-
"""Test NORTON: trích lớp text + đoán số/tên tờ của các PDF bản vẽ vào thư mục tạm (không ghi vào dự án).
py -X utf8 norton_cache.py <thư mục ra> d01|vanity"""
import sys, os, re, json, glob
sys.path.insert(0, r"D:\COSTING EXCEL\SKILL\01-scan\tool")
import pd, pymupdf

INP = {"d01": r"D:\COSTING EXCEL\_ACTIVE\NORTON\00_input\261023_D01\03_BanVe",
       "vanity": r"D:\COSTING EXCEL\_ACTIVE\NORTON\00_input\261024_VAnity\01. DRAWING"}[sys.argv[2]]
OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
tok = re.compile(r"\b[A-Z0-9]{1,6}(?:[.\-_][A-Z0-9]{1,6}){2,}\b")  # như cmd_kiemke
pdfs = sorted((f for f in glob.glob(os.path.join(INP, "**", "*.pdf"), recursive=True) if "_sheets" not in f), key=str.lower)
meta = []
for i, f in enumerate(pdfs, 1):
    fid, rel = f"N{i:02d}", os.path.relpath(f, INP)
    d = pymupdf.open(f); texts, titles = [], []
    for p in d:
        texts.append(p.get_text())
        try: titles.append(" ".join(pd.title_guess(p, tok)))
        except Exception: titles.append("")
    json.dump({"texts": texts, "titles": titles}, open(os.path.join(OUT, fid + ".json"), "w", encoding="utf-8"), ensure_ascii=False)
    meta.append({"ID": fid, "rel": rel, "pages": d.page_count, "size": os.path.getsize(f)})
    print(fid, d.page_count, rel, flush=True)
json.dump(meta, open(os.path.join(OUT, "_files.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
