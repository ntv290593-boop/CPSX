"""glyph_dataset.py - ghep kich thuoc DWG voi cum chu so PDF, tao mau chu so co nhan."""
import sys, os, re, glob, csv, math, pickle, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(sys.argv[0])))
import pymupdf, ezdxf, numpy as np
from ezdxf.entities.dimstyleoverride import DimStyleOverride
from multiprocessing import Pool
from pdf_glyphs import glyph_paths, cluster, raster
sys.stdout.reconfigure(encoding="utf-8")
A = 72 / 25.4
CODE = re.compile(r"(ID-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{3}[A-Z]?)")

def work(job):
    code, dxf, pdf = job
    res = {"code": code, "dims": [], "n_dims": 0, "n_clusters": 0, "err": None}
    try:
        page = pymupdf.open(pdf)[0]; H = page.rect.height
        C = cluster(glyph_paths(page)); res["n_clusters"] = len(C)
        d = ezdxf.readfile(dxf); ps = [l for l in d.layouts if l.name != "Model"]
        for e in (x for l in ps for x in l.query("DIMENSION")):
            tm = e.dxf.get("text_midpoint")
            if tm is None: continue
            ov = DimStyleOverride(e)
            lf = ov.get("dimlfac", 1.0) or 1.0; rnd = ov.get("dimrnd", 0) or 0; dec = ov.get("dimdec", 0) or 0
            post = (ov.get("dimpost", "") or "").strip()
            txt = e.dxf.get("text", "")
            res["n_dims"] += 1
            if txt not in ("", "<>") or dec != 0 or post not in ("", "<>"):
                res["dims"].append({"skip": "override/dec/post"}); continue
            v = e.get_measurement() * lf
            if rnd: v = math.floor(v / rnd + 0.5) * rnd
            S = str(int(math.floor(v + 0.5)))
            px, py = A * tm[0], H - A * tm[1]
            best = min(C, key=lambda c: (c["cx"]-px)**2 + (c["cy"]-py)**2) if C else None
            if best is None or math.hypot(best["cx"]-px, best["cy"]-py) > 9:
                res["dims"].append({"S": S, "skip": "no cluster"}); continue
            rot = round(float(e.dxf.get("text_rotation", 0) or 0)) % 360
            res["dims"].append({"S": S, "orient": best["orient"], "rot": rot, "n_glyph": len(best["glyphs"]),
                                "imgs": [raster(g) for g in best["glyphs"]], "items": [g["n"] for g in best["glyphs"]]})
    except Exception as ex:
        res["err"] = f"{type(ex).__name__}: {ex}"[:200]
    return res

if __name__ == "__main__":
    dxfroot, pdfroot, vercsv, out = sys.argv[1:5]
    only = sys.argv[5].split(",") if len(sys.argv) > 5 else None
    ok = {r["ma"] for r in csv.DictReader(open(vercsv, encoding="utf-8-sig")) if r["ket_qua"] == "KHOP"}
    dx = collections.defaultdict(list)
    for p in glob.glob(dxfroot + "/**/*.dxf", recursive=True):
        c = (CODE.search(os.path.basename(p).upper()) or [None, None])[1]
        if c: dx[c].append(p)
    pd = collections.defaultdict(list)
    for p in glob.glob(pdfroot + "/**/*.pdf", recursive=True):
        c = (CODE.search(os.path.basename(p).upper()) or [None, None])[1]
        if c: pd[c].append(p)
    jobs = [(c, dx[c][0], pd[c][0]) for c in sorted(ok) if len(dx.get(c, [])) == 1 and len(pd.get(c, [])) >= 1 and (only is None or c in only)]
    print("jobs", len(jobs), flush=True)
    results = []
    with Pool(6) as pool:
        for i, r in enumerate(pool.imap_unordered(work, jobs), 1):
            results.append(r)
            if i % 20 == 0: print(i, "/", len(jobs), flush=True)
    pickle.dump(results, open(out, "wb"))
    tot = sum(r["n_dims"] for r in results)
    usable = sum(1 for r in results for d in r["dims"] if "imgs" in d)
    eq = sum(1 for r in results for d in r["dims"] if "imgs" in d and d["n_glyph"] == len(d["S"]))
    print("sheets", len(results), "errors", sum(1 for r in results if r["err"]), "| dims", tot, "| matched cluster", usable, "| glyph count == digits", eq)


