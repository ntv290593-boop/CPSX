# -*- coding: utf-8 -*-
r"""scan_core.py — lõi cào + tìm trang dùng chung cho 01-scan (pd.py) và 10-tender (td.py).
Chỉ hàm thuần (không biết đường dẫn dự án, form, bước). Sửa ở đây = đổi cho cả 2 skill → chạy
DOCS\scan-toi-uu\bench.py (01-scan) và `td.py ungvien --do` (10-tender) trước/sau khi sửa.

  title_guess(page, tok)      số tờ + tên tờ trong khung tên (góc phải-dưới)
  loai_to(tên tờ)             danh_muc / thong_ke / mep / mat_bang / hoan_thien / chi_tiet / mat_dung / legend / khac
  toks, ncode, _fold          tách từ (bỏ dấu, stopword), chuẩn hóa mã (LF-04 = lf4)
  doc_pdf(path)               -> [{text, so, ten, lop}] mỗi trang
  trang_trung(texts)          trang trùng giữa các file (bỏ ngày), file tổng
  chi_muc(texts, titles)      túi từ / mã mỗi trang + df, cdf, avgdl + danh sách đảo
  _bm25(ix, q, c), bm25_dao(ix, q)   điểm BM25 1 trang / mọi trang có từ truy vấn
  chu_ky_file(path)           (size, mtime, md5) để biết file đổi nội dung
"""
import re, os, json, hashlib, collections

RAC = re.compile(r"PERMISSION|COPYRIGHT|WWW\.|@|TƯ VẤN|CONSULTANT|ARCHITECT|DRAWN|CHECKED|APPROVED|SCALE|TỶ LỆ|TỈ LỆ|"
                 r"HẠNG MỤC|ITEM:|BY:|DIRECTOR|DESIGNER|NGÀY|DATE|REV|CHỦ ĐẦU TƯ|CLIENT|PROJECT|DỰ ÁN|SỐ BẢN VẼ|DRAWING NO|TÊN BẢN VẼ|DRAWING TITLE|SHEET|ISSUE", re.I)


def title_guess(p, tok):
    """Đoán số tờ + tên tờ trong khung tên (góc phải-dưới theo chiều nhìn): số = mã gần nhãn DRAWING NO (+ mảnh tiền tố);
    tên = khối chữ to nhất gần nhãn DRAWING NAME. Chỉ là đoán — bước bản vẽ sửa khi sai."""
    import pymupdf
    W, H, M = p.rect.width, p.rect.height, p.rotation_matrix
    bl = []
    for b in p.get_text("dict")["blocks"]:
        if b.get("type") != 0: continue
        r = pymupdf.Rect(b["bbox"]) * M  # toạ độ theo chiều nhìn (trang xoay 90/270)
        if r.x0 < W * 0.6 or r.y0 < H * 0.55: continue
        sp = [s for l in b["lines"] for s in l["spans"]]
        t = "\n".join("".join(s["text"] for s in l["spans"]) for l in b["lines"])
        bl.append((r.x0, r.y0, r.x1, r.y1, t, max((s["size"] for s in sp), default=0)))
    tx = lambda b: " ".join(b[4].split())
    nos = lambda b: b[4].replace(" ", "").replace("\n", "").upper()
    ctr = lambda b: ((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
    dist = lambda a, b: ((ctr(a)[0] - ctr(b)[0]) ** 2 + (ctr(a)[1] - ctr(b)[1]) ** 2) ** 0.5
    frag = re.compile(r"^[A-Z0-9][A-Z0-9.\-_]*[-.]$")
    lab_no = [b for b in bl if re.search(r"KÝHIỆUBẢNVẼ|DRAWINGNO|SỐBẢNVẼ|SHEETNO", nos(b))]
    lab_ten = [b for b in bl if re.search(r"TÊNBẢNVẼ|DRAWINGNAME|DRAWINGTITLE", nos(b))]
    num = [(b, m) for b in bl for m in tok.findall(b[4]) if not RAC.search(tx(b)) and not frag.match(tx(b))]
    if not num: return "", ""
    nb, so = min(num, key=lambda x: dist(x[0], lab_no[0])) if lab_no else num[-1]
    fr = [b for b in bl if frag.match(tx(b))]
    if fr:
        f = min(fr, key=lambda b: dist(b, nb))
        if dist(f, nb) < W * 0.08: so = tx(f) + so
    anchor = lab_ten[0] if lab_ten else nb
    ok = [b for b in bl if b is not nb and b is not anchor and 6 <= len(tx(b)) <= 120 and not RAC.search(tx(b))
          and not tok.fullmatch(tx(b)) and not frag.match(tx(b)) and re.search(r"[A-ZÀ-Ỹ]{3}", tx(b))
          and not re.search(r"CO\.,? ?LTD|STREET|WARD|TEL[: ]", tx(b), re.I)]
    ok = [b for b in ok if dist(b, anchor) < W * 0.15] or ok
    ok.sort(key=lambda b: (-round(b[5]), dist(b, anchor)))  # tên tờ = chữ to nhất gần nhãn
    return so, " / ".join(tx(b) for b in ok[:2])[:120]


def _fold(s):
    import unicodedata
    s = unicodedata.normalize("NFD", str(s or "").lower().replace("đ", "d"))
    return re.sub(r"\s+", " ", "".join(c for c in s if unicodedata.category(c) != "Mn"))


STOP = set("va voi cho cua the loai dang kieu bang theo tren duoi trong ngoai khong co mot cac cai bo and the with for of in on".split())


LOAI_TO = [  # (loại, regex trên tên tờ đã _fold) — theo thứ tự; "chi tiết" thắng mep/hoan_thien
    ("danh_muc", r"danh muc|drawing list|muc luc|\bindex\b|cover"),
    ("thong_ke", r"thong ke|schedule|bang ke|statistic"),
    ("mep", r"\bdien\b|den tran|\bden\b|cong tac|o cam|cap thoat|\bpccc\b|\bmep\b|lighting|electrical|plumbing|ceiling plan|\bm&e\b|power plan"),
    ("mat_bang", r"mat bang|layout|floor plan|furniture plan|ff&e plan|setting out plan"),
    ("hoan_thien", r"hoan thien|\bsan\b|\btran\b|\btuong\b|finish|flooring|ceiling"),
    ("chi_tiet", r"chi tiet|detail|\bdt\b|section|mat cat"),
    ("mat_dung", r"mat dung|elevation"),
    ("legend", r"legend|ky hieu|ghi chu chung"),
]


TOK = re.compile(r"[a-z0-9]+")
STOP2 = STOP | set("cung cap lap dat hoan thien ban ve hoan cong phong".split())
CODE_IN_TEXT = re.compile(r"(?<![a-z0-9])([a-z]{1,4})[\s._-]?0*(\d{1,3})(?![0-9])")


def loai_to(title):
    s = _fold(title)
    has_ct = bool(re.search(r"chi tiet|detail", s))
    for loai, rx in LOAI_TO:
        if loai in ("mep", "hoan_thien") and has_ct: continue
        if re.search(rx, s): return loai
    return "khac"


def toks(s): return [w for w in TOK.findall(_fold(s)) if len(w) >= 2 and w not in STOP2]


def ncode(s):
    """Chuẩn hóa mã: bỏ . - _ khoảng trắng, bỏ số 0 đầu: LF-04 / LF04 / lf.4 -> lf4. Tiền tố 'ID-' bỏ."""
    s = re.sub(r"^id[\s._-]+", "", _fold(s).strip())
    m = re.fullmatch(r"([a-z]{1,4})[\s._-]*0*(\d{1,3})", s)
    return f"{m.group(1)}{m.group(2)}" if m else None


def _bm25(ix, q, c, k1=1.2, b=0.75):
    import math
    dl = sum(c.values()); s = 0.0
    for w in q:
        if w in c:
            idf = math.log(1 + (ix["N"] - ix["df"][w] + 0.5) / (ix["df"][w] + 0.5))
            s += idf * c[w] * (k1 + 1) / (c[w] + k1 * (1 - b + b * dl / ix["avgdl"]))
    return s


TOK_SO = re.compile(r"\b[A-Z0-9]{1,6}(?:[.\-_][A-Z0-9]{1,6}){2,}\b")  # ứng viên số tờ cho title_guess


def doc_pdf(path=None, data=None):
    """-> [{"text", "so", "ten", "lop"}] mỗi trang. lop: CÓ (>200 ký tự) / ÍT / KHÔNG (ảnh scan hoặc chữ vẽ bằng nét).
    data = bytes PDF (đọc thẳng từ zip, không giải nén ra đĩa)."""
    import pymupdf
    out = []
    with (pymupdf.open(stream=data, filetype="pdf") if data is not None else pymupdf.open(path)) as d:
        for p in d:
            t = p.get_text()
            n = len(t.strip())
            try:
                so, ten = title_guess(p, TOK_SO)
            except Exception:  # khung tên lạ không được làm hỏng cả file
                so, ten = "", ""
            out.append({"text": t, "so": so, "ten": ten, "lop": "CÓ" if n > 200 else ("ÍT" if n else "KHÔNG")})
    return out


def trang_trung(texts):
    """texts {fid: [chữ trang]} -> (dup {(fid, tr): (fid giữ, tr)}, [fid file tổng]).
    Trang trùng = cùng md5 sau khi bỏ ngày dd/mm/yy và khoảng trắng; giữ bản trong file ít trang hơn (cụ thể hơn), hòa -> file sau.
    File tổng = ≥ 85% trang có chữ là trang trùng file khác."""
    h = {}
    for fid, pages in texts.items():
        for i, t in enumerate(pages, 1):
            k = re.sub(r"\s+", "", re.sub(r"\d{1,2}[/.]\d{1,2}[/.]\d{2,4}", "", t))  # bỏ ngày: bản 29/09 vs 30/09 chỉ khác ngày
            if len(k) > 50: h.setdefault(hashlib.md5(k.encode("utf-8")).hexdigest(), []).append((fid, i))
    npg = {fid: len(p) for fid, p in texts.items()}
    dup = {}
    for v in h.values():
        if len({fid for fid, _ in v}) < 2: continue
        keep = min(v, key=lambda x: (npg[x[0]], -int(re.sub(r"\D", "", x[0]) or 0), x[1]))  # file ít trang hơn (cụ thể hơn); hòa -> file sau
        for p in v:
            if p != keep: dup[p] = keep
    tong = []
    for fid, pages in texts.items():
        real = [(fid, i) for i, t in enumerate(pages, 1) if len(re.sub(r"\s+", "", t)) > 50]
        if real and sum(p in dup for p in real) >= 0.85 * len(real): tong.append(fid)
    return dup, tong


def chi_muc(texts, titles):
    """Chỉ mục trang. texts {fid: [chữ]}, titles {(fid, tr): "số + tên tờ"} -> dict:
    pages, body/title (Counter từ), codes (Counter mã chuẩn hóa), tit_codes, N, df, cdf, avgdl, dao {từ: {trang: tf}}."""
    ix = {}
    pages = [(fid, i) for fid in texts for i in range(1, len(texts[fid]) + 1)]
    ix["pages"] = pages
    ix["body"] = {p: collections.Counter(toks(texts[p[0]][p[1] - 1])) for p in pages}
    ix["title"] = {p: collections.Counter(toks(titles.get(p, ""))) for p in pages}
    ix["codes"] = {p: collections.Counter(a + b for a, b in CODE_IN_TEXT.findall(_fold(texts[p[0]][p[1] - 1]))) for p in pages}
    ix["N"] = len(pages)
    ix["cdf"] = collections.Counter(c for cn in ix["codes"].values() for c in cn)
    ix["avgdl"] = (sum(sum(c.values()) for c in ix["body"].values()) / len(pages)) if pages else 1
    ix["df"] = collections.Counter(w for c in ix["body"].values() for w in c)
    ix["tit_codes"] = {p: {ncode(x) for x in re.findall(r"[A-Za-z]{1,4}[-._]?\d{1,3}", titles.get(p, ""))} for p in pages}
    dao = collections.defaultdict(dict)
    for p, c in ix["body"].items():
        for w, n in c.items(): dao[w][p] = n
    ix["dao"] = dao
    ix["dl"] = {p: sum(c.values()) for p, c in ix["body"].items()}
    return ix


def bm25_dao(ix, q, k1=1.2, b=0.75):
    """BM25 cho mọi trang có ít nhất 1 từ truy vấn (qua danh sách đảo) -> {trang: điểm}. Cùng công thức với _bm25."""
    import math
    out = collections.defaultdict(float)
    for w in set(q):
        post = ix["dao"].get(w)
        if not post: continue
        idf = math.log(1 + (ix["N"] - ix["df"][w] + 0.5) / (ix["df"][w] + 0.5))
        for p, tf in post.items():
            out[p] += idf * tf * (k1 + 1) / (tf + k1 * (1 - b + b * ix["dl"][p] / ix["avgdl"]))
    return out


def chu_ky_file(path, cu=None):
    """-> {"size", "mtime", "md5"}. cu = chữ ký lần trước: size + mtime không đổi thì dùng lại md5 (khỏi đọc file lớn)."""
    st = os.stat(path)
    ck = {"size": st.st_size, "mtime": int(st.st_mtime)}
    if cu and cu.get("size") == ck["size"] and cu.get("mtime") == ck["mtime"] and cu.get("md5"):
        ck["md5"] = cu["md5"]
        return ck
    h = hashlib.md5()
    with open(path, "rb") as fh:
        for b in iter(lambda: fh.read(1 << 20), b""): h.update(b)
    ck["md5"] = h.hexdigest()
    return ck
