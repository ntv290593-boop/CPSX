"""pdf_dim_decode.py - doc so kich thuoc (chu so ve bang net vector) tu mot trang PDF, KHONG can DWG.
Dung: python -I pdf_dim_decode.py <file.pdf> [templates.npz] [--page N]
Moi cum so: text, vi tri (pt), huong (h/v), khoang cach lon nhat toi mau, accepted (khoang cach <= TH, moi net co mau cung so doan, khong bat dau bang 0 - loai ranh CNC/hinh bau duc giong chu 0)."""
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.abspath(sys.argv[0])))
import numpy as np, pymupdf
from pdf_glyphs import glyph_paths, cluster, raster
TH = 3000.0

class Decoder:
    def __init__(self, npz):
        z = np.load(npz, allow_pickle=False)
        self.X = z["X"].astype(np.float32); self.L = z["L"]; self.O = z["O"]; self.N = z["N"]
    def glyph(self, img, orient, n):
        m = (self.O == orient) & (self.N == n)
        if not m.any(): return "?", 1e9
        d = ((self.X[m] - img.reshape(-1).astype(np.float32)) ** 2).sum(1); i = int(d.argmin())
        return str(self.L[m][i]), float(d[i])

def decode_page(page, dec):
    out = []
    for c in cluster(glyph_paths(page)):
        res = [dec.glyph(raster(g), c["orient"], g["n"]) for g in c["glyphs"]]
        mx = max(d for _, d in res)
        out.append({"text": "".join(ch for ch, _ in res), "cx": round(c["cx"], 1), "cy": round(c["cy"], 1), "orient": c["orient"],
                    "dist": round(mx), "accepted": mx <= TH and "?" not in [ch for ch, _ in res] and not "".join(ch for ch, _ in res).startswith("0")})
    return out

if __name__ == "__main__":
    pdf = sys.argv[1]
    npz = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith("--") else os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "glyph_templates.npz")
    pg = int(sys.argv[sys.argv.index("--page") + 1]) - 1 if "--page" in sys.argv else 0
    sys.stdout.reconfigure(encoding="utf-8")
    r = decode_page(pymupdf.open(pdf)[pg], Decoder(npz))
    print(json.dumps(r, ensure_ascii=False))

