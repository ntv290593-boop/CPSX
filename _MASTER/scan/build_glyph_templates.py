"""build_glyph_templates.py - tao thu vien mau chu so tu dataset (DWG nhan) -> glyph_templates.npz
Dung: python -I build_glyph_templates.py dataset.pkl out.npz"""
import pickle, sys, collections, numpy as np
R = [r for r in pickle.load(open(sys.argv[1], "rb")) if not r["err"]]
raw = []
for r in R:
    for d in r["dims"]:
        if "imgs" in d and d["n_glyph"] == len(d["S"]):
            for ch, im, n in zip(d["S"], d["imgs"], d["items"]):
                raw.append((ch, d["orient"], n, im.reshape(-1).astype(np.uint8)))
cnt = collections.defaultdict(collections.Counter)
for ch, o, n, _ in raw: cnt[ch][n] += 1
modal = {ch: cnt[ch].most_common(1)[0][0] for ch in cnt}
clean = [x for x in raw if x[2] == modal[x[0]]]
# bo trung lap de nho gon
seen = {}; 
for ch, o, n, im in clean: seen[(ch, o, n, im.tobytes())] = im
keys = list(seen)
X = np.array([seen[k] for k in keys], dtype=np.uint8)
np.savez_compressed(sys.argv[2], X=X, L=np.array([k[0] for k in keys]), O=np.array([k[1] for k in keys]), N=np.array([k[2] for k in keys]))
print("raw", len(raw), "clean", len(clean), "unique", len(keys), "modal items", modal)
