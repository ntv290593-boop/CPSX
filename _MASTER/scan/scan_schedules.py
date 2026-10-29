"""Doc bang thong ke (cua, ...) tu PDF ban ve CAD bang pdfplumber.

Tim trang theo tu khoa tieu de, nhan dien bang theo tieu de cot VI/EN,
tra ve tung dong kem file/trang/bang/dong va do tin cay.
"""
import json, re, sys, unicodedata
import pdfplumber

PAGE_KEYS = [r"TH[ỐO]NG K[ÊE] C[ỬU]A", r"DOOR SCHEDULE"]
# cot chuan -> tu khoa tieu de (khong dau, chu thuong)
COLS = [  # thu tu uu tien: cot cu the truoc cot chung
    ("CW",    ["chieu rong lot", "clear width"]),
    ("CH",    ["chieu cao lot", "clear height"]),
    ("type",  ["loai cua", "ma cua", "ki hieu", "ky hieu", "type mark", "symbol", "mark", "door no", "type"]),
    ("EI",    ["ghcl", "chong chay", "fire"]),
    ("count", ["so luong", "count", "qty", "quantity"]),
    ("W",     ["chieu rong", "width"]),
    ("H",     ["chieu cao phu", "chieu cao/", "overall height", "height"]),
]
TYPE_RE = re.compile(r"^(?:ID-)?D[A-Z]{0,2}\d{3}[A-Z]?$")

def norm(s):
    s = unicodedata.normalize("NFD", (s or "").replace("Đ", "D").replace("đ", "d"))
    return re.sub(r"\s+", " ", "".join(c for c in s if unicodedata.category(c) != "Mn")).lower().strip()

def map_header(row):
    m = {}
    for i, cell in enumerate(row):
        h = norm(cell)
        for k, keys in COLS:
            if k not in m and any(x in h for x in keys):
                m[k] = i
                break
    return m

def num(v):
    v = (v or "").replace(" ", "")
    return int(v) if re.fullmatch(r"\d+", v) else None

def scan_page(page):
    out, last = [], None  # last = (header map, so cot) cua bang truoc: tieu de co the tach rieng
    for ti, tb in enumerate(page.extract_tables()):
        hdr, hi = None, None
        for ri, row in enumerate(tb[:6]):
            m = map_header(row)
            if "type" in m and len(m) >= 3:
                hdr, hi = m, ri
                break
        if hdr:
            last = (hdr, len(tb[0]))
        elif last and tb and len(tb[0]) == last[1]:
            hdr, hi = last[0], -1
        else:
            continue
        for ri, row in enumerate(tb[hi + 1:], hi + 1):
            t = (row[hdr["type"]] or "").replace("\n", "").strip()
            if not TYPE_RE.match(t):
                continue
            rec = {"type": t, "table": ti + 1, "row": ri + 1}
            ok = 1
            for k in ("W", "H", "CW", "CH", "count"):
                if k in hdr:
                    rec[k] = num(row[hdr[k]])
                    ok += rec[k] is not None
            if "EI" in hdr:
                rec["EI"] = re.sub(r"\s+", "", row[hdr["EI"]] or "") or None
            need = 1 + sum(k in hdr for k in ("W", "H", "CW", "CH", "count"))
            rec["conf"] = round(ok / need, 2) if "count" in hdr else round(0.5 * ok / need, 2)
            out.append(rec)
    return out

def scan_pdf(path, pages=None):
    """pages: list so trang (1-based) can doc; None = tu tim theo tu khoa."""
    res = []
    with pdfplumber.open(path) as pdf:
        idx = pages or [i + 1 for i, p in enumerate(pdf.pages)
                        if any(re.search(k, p.extract_text() or "", re.I) for k in PAGE_KEYS)]
        for pn in idx:
            for r in scan_page(pdf.pages[pn - 1]):
                r["page"] = pn
                res.append(r)
    return res

if __name__ == "__main__":
    pg = [int(x) for x in sys.argv[2].split(",")] if len(sys.argv) > 2 else None
    print(json.dumps(scan_pdf(sys.argv[1], pg), ensure_ascii=False, indent=1))
