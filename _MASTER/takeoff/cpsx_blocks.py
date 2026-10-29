"""Lõi dùng chung: đọc block (Template/Door/Kitchen/Chiet tinh), so block với mẫu gốc.

- Block = dòng A có chữ và B hoặc E là số (như libBlocks/ctBlocks trong cpsx_lib.js) đến dòng A="END".
- So block với mẫu: ghép dòng theo chuỗi nhãn cột A (difflib), dịch công thức mẫu sang vị trí đích
  (openpyxl Translator) rồi so; khác = ô đã sửa (edit). Dòng không ghép được = dòng thêm/bớt.
"""
import re, difflib
import openpyxl


LIB = ["Template", "Door", "Kitchen"]
CT = "Chiet tinh"
COLS = "BCDEFGHIJKLMNOPQR"       # vùng ô nhập (A là nhãn/khóa SUMIF)
NHAN_CONG = "chi phí nhân công"


def norm(v):
    return re.sub(r"\s+", " ", "" if v is None else str(v)).strip().lower()


def is_f(x):
    return isinstance(x, str) and x.startswith("=")


def txt(x):
    """ArrayFormula → chuỗi công thức."""
    return getattr(x, "text", x)


REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])")


def shape(f):
    """Bỏ số hàng trong tham chiếu (dịch hàng do chèn/xoá dòng không tính là sửa) và '+' đầu."""
    s = REF.sub(lambda m: m[1] + m[2] + m[3] + "#", str(f)).replace(" ", "")
    return s.replace("=+", "=")


class Book:
    """Đọc 1 lần cả công thức và giá trị các sheet cần (read_only)."""

    def __init__(self, path, sheets):
        wf = openpyxl.load_workbook(path, read_only=True)
        wv = openpyxl.load_workbook(path, read_only=True, data_only=True)
        self.F, self.V = {}, {}
        for sn in sheets:
            if sn in wf.sheetnames:
                self.F[sn] = [list(r) for r in wf[sn].iter_rows(min_col=1, max_col=18, values_only=True)]
                self.V[sn] = [list(r) for r in wv[sn].iter_rows(min_col=1, max_col=18, values_only=True)]

    def cell(self, sn, row, col, kind="F"):
        t = (self.F if kind == "F" else self.V)[sn]
        j = col if isinstance(col, int) else ord(col) - 64
        return t[row - 1][j - 1] if row - 1 < len(t) and j - 1 < len(t[row - 1]) else None

    def blocks(self, sn):
        V, F = self.V[sn], self.F[sn]
        out, cur = [], None
        for i, v in enumerate(V):
            lab = norm(v[0])
            if lab == "end":
                if cur:
                    cur["end"] = i + 1
                    out.append(cur)
                    cur = None
                continue
            if not cur and lab and (isinstance(v[1], (int, float)) or isinstance(v[4], (int, float))):
                m = re.search(r"'?([^'!]+)'?!\$?A\$?(\d+)", str(F[i][0])) if is_f(F[i][0]) else None
                cur = {"sheet": sn, "name": str(v[0]).strip(), "hdr": i + 1,
                       "boq": (m[1], int(m[2])) if m else None}
        return out

    def labels(self, b):
        return [norm(self.V[b["sheet"]][r - 1][0]) for r in range(b["hdr"], b["end"] + 1)]


def same(a, b):
    if a in (None, "") and b in (None, ""):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) < 1e-9
    return str(a).replace(" ", "") == str(b).replace(" ", "")


def match_source(book, blk, cands):
    """Mẫu gốc của 1 block Chiet tinh = block thư viện có chuỗi nhãn A giống nhất."""
    la = book.labels(blk)
    best = (0, None)
    for c in cands:
        r = difflib.SequenceMatcher(None, la, book.labels(c), autojunk=False).ratio()
        if r > best[0]:
            best = (r, c)
    return best


def kind(fd, fs):
    """Loại khác biệt giữa ô đích fd và ô mẫu fs (None = coi như giống)."""
    fd, fs = txt(fd), txt(fs)
    if is_f(fs) and is_f(fd):
        return None if shape(fd) == shape(fs) else "SUA_CONG_THUC"
    if is_f(fs):
        return "DE_SO" if fd not in (None, "") else "XOA_CONG_THUC"
    if is_f(fd):
        return "LINK"
    return None if same(fd, fs) else "NHAP"


def diff_block(book, blk, src):
    """Ô khác mẫu trong block đích. Trả (edits, dòng thêm, dòng bớt).
    edit = dict: off (dòng tính từ header của MẪU), off_dich, Module (nhãn A của mẫu), Lan (lần xuất hiện
    nhãn khớp chính xác trong mẫu), Cot, GiaTri (đích), Mau_GiaTri, Loai (NHAP/LINK/DE_SO/SUA_CONG_THUC/XOA_CONG_THUC)."""
    la, lb = book.labels(blk), book.labels(src)
    sm = difflib.SequenceMatcher(None, la, lb, autojunk=False)
    pairs, added, removed = [], [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal" or (tag == "replace" and i2 - i1 == j2 - j1):
            pairs += list(zip(range(i1, i2), range(j1, j2)))
        else:
            added += list(range(i1, i2))
            removed += list(range(j1, j2))
    edits = []
    for i, j in pairs:
        rd, rs = blk["hdr"] + i, src["hdr"] + j
        lan = sum(1 for x in lb[:j + 1] if x == lb[j])
        lab = book.cell(src["sheet"], rs, "A", "V")
        for c in COLS:
            fd = txt(book.cell(blk["sheet"], rd, c))
            fs = txt(book.cell(src["sheet"], rs, c))
            k = kind(fd, fs)
            if k:
                edits.append({"off": j, "off_dich": i, "Module": "hdr" if j == 0 else ("" if lab is None else str(lab).strip()),
                              "Lan": lan, "Cot": c, "GiaTri": fd, "Mau_GiaTri": fs, "Loai": k})
    return edits, [blk["hdr"] + i for i in added], [src["hdr"] + j for j in removed]
