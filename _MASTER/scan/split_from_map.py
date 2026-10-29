"""Tach PDF ban ve theo bang ten co san (sheetmap.json) - chay tren may Windows.

python split_from_map.py <thu muc 03_BanVe> <sheetmap.json>
sheetmap: {ten file goc khong .pdf: {so trang: "_sheets\\KHOI\\pNNN_xxx.pdf"}}. Bo qua file da co.
"""
import json, os, sys
from pypdf import PdfReader, PdfWriter

def run(banve, mapfile):
    smap = json.load(open(mapfile, encoding="utf-8"))
    done = skip = 0
    for src, pages in smap.items():
        todo = {int(k): v for k, v in pages.items() if not os.path.exists(os.path.join(banve, v))}
        skip += len(pages) - len(todo)
        if not todo:
            continue
        reader = PdfReader(os.path.join(banve, src + ".pdf"))
        for pn, rel in sorted(todo.items()):
            out = os.path.join(banve, rel)
            os.makedirs(os.path.dirname(out), exist_ok=True)
            w = PdfWriter()
            w.add_page(reader.pages[pn - 1])
            w.compress_identical_objects(remove_duplicates=True, remove_unreferenced=True)
            with open(out, "wb") as f:
                w.write(f)
            done += 1
    print(f"tach moi {done}, da co {skip}")

if __name__ == "__main__":
    run(sys.argv[1], sys.argv[2])
