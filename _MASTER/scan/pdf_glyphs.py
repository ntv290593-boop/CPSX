"""pdf_glyphs.py - lay net chu so (stroke) tu trang PDF vector, gom thanh cum so, tao anh dac trung."""
import pymupdf, numpy as np
from PIL import Image, ImageDraw

TEXT_H = 5.6   # cao chu so (pt) o dimstyle 1-1-PS (dimtxt 2.0 mm = 5.67 pt)

def _bez(p0, p1, p2, p3, n=6):
    out = []
    for i in range(n + 1):
        t = i / n; u = 1 - t
        out.append((u**3*p0[0]+3*u*u*t*p1[0]+3*u*t*t*p2[0]+t**3*p3[0], u**3*p0[1]+3*u*u*t*p1[1]+3*u*t*t*p2[1]+t**3*p3[1]))
    return out

def glyph_paths(page, lo=4.6, hi=7.2):
    out = []
    for d in page.get_drawings():
        if d.get("type") != "s" or not (2 <= len(d["items"]) <= 60):
            continue
        if abs((d.get("width") or 0) - 0.48) > 0.05:
            continue
        r = d["rect"]; w, h = r.width, r.height
        mx, mn = max(w, h), min(w, h)
        if not (lo <= mx <= hi and 0.8 <= mn <= 5.6):
            continue
        horiz = h >= w
        segs = []
        for it in d["items"]:
            if it[0] == "l":
                segs.append(((it[1].x, it[1].y), (it[2].x, it[2].y)))
            elif it[0] == "c":
                pts = _bez(*[(q.x, q.y) for q in it[1:5]])
                segs += list(zip(pts[:-1], pts[1:]))
        out.append({"rect": (r.x0, r.y0, r.x1, r.y1), "n": len(d["items"]), "segs": segs, "horiz": horiz, "vert": not horiz})
    return out
def cluster(glyphs, gap=4.6, base=2.3, narrow_prev=2.0, narrow_next=1.4, narrow_w=2.0, adaptive=True):
    """Gom net thanh cum so. adaptive: khoang cach cho phep = base + (narrow_prev neu net truoc la '1' hep) + (narrow_next neu net sau hep)."""
    res = []
    for orient in ("h", "v"):
        gs = [g for g in glyphs if (g["horiz"] if orient == "h" else g["vert"])]
        if orient == "h":
            gs.sort(key=lambda g: (round((g["rect"][1] + g["rect"][3]) / 2 / 1.5), g["rect"][0]))
        else:
            gs.sort(key=lambda g: (round((g["rect"][0] + g["rect"][2]) / 2 / 1.5), -g["rect"][3]))
        cur = []
        def flush():
            if cur:
                res.append({"orient": orient, "glyphs": list(cur),
                            "cx": (min(g["rect"][0] for g in cur) + max(g["rect"][2] for g in cur)) / 2,
                            "cy": (min(g["rect"][1] for g in cur) + max(g["rect"][3] for g in cur)) / 2})
        for g in gs:
            if not cur:
                cur = [g]; continue
            p = cur[-1]
            if orient == "h":
                same = abs((g["rect"][1] + g["rect"][3]) / 2 - (p["rect"][1] + p["rect"][3]) / 2) < 1.2
                gp = g["rect"][0] - p["rect"][2]; ok_order = g["rect"][0] - p["rect"][0] > 0.5
                wp = p["rect"][2] - p["rect"][0]; wg = g["rect"][2] - g["rect"][0]
            else:
                same = abs((g["rect"][0] + g["rect"][2]) / 2 - (p["rect"][0] + p["rect"][2]) / 2) < 1.2
                gp = p["rect"][1] - g["rect"][3]; ok_order = p["rect"][1] - g["rect"][1] > 0.5
                wp = p["rect"][3] - p["rect"][1]; wg = g["rect"][3] - g["rect"][1]
            lim = (base + (narrow_prev if wp <= narrow_w else 0) + (narrow_next if wg <= narrow_w else 0)) if adaptive else gap
            if same and ok_order and gp < lim:
                cur.append(g)
            else:
                flush(); cur = [g]
        flush()
    return res
def raster(g, size=40, scale=5.0):
    """Anh 20x20 (uint8) cua mot net chu so, tam bbox o giua, ti le co dinh (khong chuan hoa theo bbox)."""
    x0, y0, x1, y1 = g["rect"]; cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    im = Image.new("L", (size, size), 0); dr = ImageDraw.Draw(im)
    for (a, b) in g["segs"]:
        dr.line([((a[0]-cx)*scale+size/2, (a[1]-cy)*scale+size/2), ((b[0]-cx)*scale+size/2, (b[1]-cy)*scale+size/2)], fill=255, width=2)
    arr = np.asarray(im, dtype=np.float32).reshape(size // 2, 2, size // 2, 2).mean(axis=(1, 3))
    return arr.astype(np.uint8)




