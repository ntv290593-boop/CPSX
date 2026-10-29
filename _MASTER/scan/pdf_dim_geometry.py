"""pdf_dim_geometry.py - kiem cheo so kich thuoc doc duoc voi chieu dai hinh hoc cua duong kich thuoc trong PDF.
Gia tri = chieu dai (pt) / (72/25.4) * ti le chi tiet. Dau duong kich thuoc la gach cheo ngan (net xien ~5.7 pt)."""
import re, math, collections
A = 72 / 25.4
STD = (1, 2, 5, 10, 15, 20, 25, 50, 100)

def find_ticks(page):
    ticks = []
    for d in page.get_drawings():
        if d.get("type") == "s" and len(d["items"]) == 1 and d["items"][0][0] == "l":
            r = d["rect"]; w = d.get("width") or 0
            if 0.3 <= w <= 0.8 and 4.5 <= r.width <= 6.5 and 4.5 <= r.height <= 6.5 and abs(r.width - r.height) < 1.0:
                c = ((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
                if not any(abs(c[0] - q[0]) < 0.3 and abs(c[1] - q[1]) < 0.3 for q in ticks[-40:]) and c not in ticks:
                    ticks.append(c)
    # loai trung (nhieu net chong nhau)
    u = []
    for t in ticks:
        if not any(math.hypot(t[0] - q[0], t[1] - q[1]) < 0.3 for q in u): u.append(t)
    return u

def scale_labels(page):
    """Cac nhan 'Scale 1: N' trong text layer: [(x, y, N)]. So N la doi tuong text rieng nen ghep theo vi tri."""
    W = [(w[0], w[1], w[4]) for w in page.get_text("words")]
    out = []
    for x, y, t in W:
        if t.lower() != "scale": continue
        line = sorted([(wx, wt) for wx, wy, wt in W if abs(wy - y) < 3.5 and wx > x and wx - x < 80], key=lambda a: a[0])
        s = "".join(wt for _, wt in line)
        m = re.match(r"1:(\d+)", s)
        if m: out.append((x, y, int(m.group(1))))
    return out

def page_scales(page):
    return {n for _, _, n in scale_labels(page)}
class TickIndex:
    def __init__(self, ticks):
        self.h = collections.defaultdict(list); self.v = collections.defaultdict(list)
        for x, y in ticks:
            self.h[round(y * 2)].append(x); self.v[round(x * 2)].append(y)
        for k in self.h: self.h[k].sort()
        for k in self.v: self.v[k].sort()

    def lines(self, orient, c, reach=16.0, lo=0.0):
        """Cac duong gach cheo (ngang neu chu ngang, doc neu chu doc) trong pham vi 'reach' quanh tam chu."""
        src, pos = (self.h, c[1]) if orient == "h" else (self.v, c[0])
        out = []
        for k, vals in src.items():
            if lo <= abs(k / 2 - pos) <= reach and len(vals) >= 2:
                out.append((k / 2, vals))
        return out

def candidates(idx, orient, cx, cy, reach=16.0, near=60.0, lo=0.0, allpairs=False):
    """Cac doan (len pt, trong?) giua hai gach cheo ke nhau gan tam chu."""
    along = cx if orient == "h" else cy
    out = []
    for lpos, vals in idx.lines(orient, (cx, cy), reach, lo):
        pairs = [(vals[i], vals[j]) for i in range(len(vals)) for j in range(i + 1, len(vals))] if allpairs else list(zip(vals[:-1], vals[1:]))
        for a, b in pairs:
            if b - a < 8: continue
            enc = a <= along <= b
            dist = 0 if enc else min(abs(along - a), abs(along - b))
            if enc or dist <= near:
                out.append({"len": b - a, "enclosing": enc, "dist": dist, "line": lpos})
    return out

def check(n, cands, scales, tol=3.0):
    """n: so (int). Tra ve (status, doan khop). status: OK_ENCLOSING / OK_NEARBY / MISMATCH / NO_SEGMENT."""
    if not cands: return "NO_SEGMENT", None
    for enc in (True, False):
        for c in sorted([c for c in cands if c["enclosing"] == enc], key=lambda c: c["dist"]):
            mm = c["len"] / A
            for s in scales:
                if abs(mm * s - n) <= tol:
                    return ("OK_ENCLOSING" if enc else "OK_NEARBY"), {"scale": s, "mm": round(mm, 2), "len": round(c["len"], 1)}
    return "MISMATCH", None


def scale_for(cx, cy, labels, w=0.3):
    """Ti le cua chi tiet chua diem (cx,cy): nhan 'Scale 1:N' o phia duoi gan nhat (uu tien cung cot)."""
    best = None
    for x, y, n in labels:
        if y < cy - 8: continue
        cost = (y - cy) + w * abs(x - cx)
        if best is None or cost < best[0]: best = (cost, n)
    return best[1] if best else None


