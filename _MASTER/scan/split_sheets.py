"""Tach PDF ban ve gop thanh tung trang (moi trang 1 file nho) de mo nhanh.

Ten file: p081_CD.66.00.pdf (so trang + 3 doan cuoi so ban ve neu co).
"""
import os, re, sys
from pypdf import PdfReader, PdfWriter

def sheet_name(page, drawing_no=""):
    tail = ".".join(re.sub(r"\s+", "", drawing_no).split(".")[-3:]) if drawing_no else ""
    tail = re.sub(r"[^A-Za-z0-9.\-]", "", tail)
    return f"p{page:03d}" + (f"_{tail}" if tail else "") + ".pdf"

def split_pdf(src, out_dir, numbers=None):
    """numbers: {so trang: so ban ve}. Tra ve {so trang: ten file}."""
    os.makedirs(out_dir, exist_ok=True)
    reader, res = PdfReader(src), {}
    for i, page in enumerate(reader.pages, 1):
        w = PdfWriter()
        w.add_page(page)
        w.compress_identical_objects(remove_duplicates=True, remove_unreferenced=True)
        name = sheet_name(i, (numbers or {}).get(i, ""))
        with open(os.path.join(out_dir, name), "wb") as f:
            w.write(f)
        res[i] = name
    return res

if __name__ == "__main__":
    print(split_pdf(sys.argv[1], sys.argv[2]))
