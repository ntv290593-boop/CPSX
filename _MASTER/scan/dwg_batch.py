"""dwg_batch.py - chay tang A tren thu muc DXF (ODA): kich thuoc that, go de, chi tiet, so phien ban voi PDF.
Dung: python -I dwg_batch.py <dxf_root> <out_dir> [<pdf_sheets_root>] [--workers N]
Moi tờ: out_dir/json/<ma>.json; tong hop: out_dir/summary.csv, out_dir/version_check.csv"""
import ezdxf, sys, json, re, os, csv, glob, collections, time
from multiprocessing import Pool
sys.stdout.reconfigure(encoding="utf-8")
ID_RE = re.compile(r"^ID-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{3}$")
strip = lambda s: re.sub(r"[{}]|\\P", " ", re.sub(r"\\[A-Za-z][^;]*;", "", s or "")).strip()

def process(path):
    t0 = time.time()
    res = {"file": path, "ok": False}
    try:
        d = ezdxf.readfile(path)
        lays = [l for l in d.layouts if l.name != "Model"]
        res["layouts"] = [l.name for l in lays]
        dims, vps, titles, tb = [], [], [], {}
        styles = collections.Counter()
        for lay in lays:
            lvps = []
            for v in lay.query("VIEWPORT"):
                if v.dxf.id == 1 or not v.dxf.view_height:
                    continue
                c, w, h = v.dxf.center, v.dxf.width, v.dxf.height
                lvps.append({"id": v.dxf.id, "x0": c[0]-w/2, "x1": c[0]+w/2, "y0": c[1]-h/2, "y1": c[1]+h/2, "scale": round(h/v.dxf.view_height, 4)})
            vps += [{"layout": lay.name, **v} for v in lvps]
            for e in lay.query("INSERT"):
                at = {a.dxf.tag: a.dxf.text for a in e.attribs if a.dxf.text}
                if not at:
                    continue
                if "TITLE" in at and "N\u00ba" in at:
                    titles.append({"layout": lay.name, "no": at["N\u00ba"], "title": at["TITLE"], "scale": at.get("SCALE"), "x": e.dxf.insert[0], "y": e.dxf.insert[1]})
                if any(ID_RE.match(str(x)) for x in at.values()):
                    tb = at
            for e in lay.query("DIMENSION"):
                ds = d.dimstyles.get(e.dxf.dimstyle)
                lf = ds.dxf.get("dimlfac", 1.0) if ds else 1.0
                rnd = ds.dxf.get("dimrnd", 0) if ds else 0
                try:
                    m = e.get_measurement()
                except Exception:
                    m = None
                if not isinstance(m, (int, float)):
                    continue
                styles[(e.dxf.dimstyle, lf, rnd or 0)] += 1
                txt = e.dxf.get("text", "")
                ov = txt not in ("", "<>")
                val = m * lf
                mism = None
                if ov:
                    try:
                        mism = abs(float(re.sub(r"[^\d.]", "", txt)) - val) > 0.51
                    except ValueError:
                        mism = None
                else:
                    mism = False
                p = e.dxf.defpoint
                dims.append({"layout": lay.name, "val": round(val, 2), "rnd": rnd or 0, "override": txt if ov else None, "mism": mism, "x": p[0], "y": p[1]})
        # gan chi tiet
        def vp_at(x, y, layout):
            for v in vps:
                if v["layout"] == layout and v["x0"] <= x <= v["x1"] and v["y0"] <= y <= v["y1"]:
                    return v
            cand = [v for v in vps if v["layout"] == layout]
            return min(cand, key=lambda a: ((a["x0"]+a["x1"])/2-x)**2+((a["y0"]+a["y1"])/2-y)**2) if cand else None
        def title_of(v):
            best = None
            for t in titles:
                if t["layout"] == v["layout"] and t["y"] <= v["y0"]+5 and v["x0"]-40 <= t["x"] <= v["x1"]:
                    g = v["y0"]-t["y"]
                    if best is None or g < best[0]:
                        best = (g, t)
            return best[1] if best else None
        groups = collections.defaultdict(list)
        unassigned = 0
        for x in dims:
            v = vp_at(x["x"], x["y"], x["layout"])
            t = title_of(v) if v else None
            if t is None:
                unassigned += 1
                key = ("?", "")
            else:
                key = (t["no"], t["title"])
            groups[key].append(x)
        details = [{"no": k[0], "title": k[1], "n": len(g), "vals": sorted(round(x["val"], 1) for x in g),
                    "overrides": [{"text": x["override"], "val": x["val"]} for x in g if x["override"]]} for k, g in sorted(groups.items())]
        sid = next((str(x) for x in tb.values() if ID_RE.match(str(x))), None)
        rev = tb.get("-") or tb.get("REV")
        res.update(ok=True, id=sid, date=tb.get("DATE"), rev=rev, tb=tb, n_dims=len(dims), n_override=sum(1 for x in dims if x["override"]),
                   n_mism=sum(1 for x in dims if x["mism"]), n_vp=len(vps), n_titles=len(titles), unassigned=unassigned,
                   styles=[{"style": k[0], "lfac": k[1], "rnd": k[2], "n": n} for k, n in styles.items()], details=details)
    except Exception as ex:
        res["error"] = f"{type(ex).__name__}: {ex}"[:300]
    res["sec"] = round(time.time()-t0, 1)
    return res

def pdf_title(pdf):
    import pymupdf
    try:
        p = pymupdf.open(pdf)[0]
        W = [(w[4], w[0], w[1]) for w in p.get_text("words")]
        box = [w for w in W if w[1] > 2090 and 1520 < w[2] < 1612]
        sid = next((w[0] for w in box if ID_RE.match(w[0])), None)
        rev = next((w[0] for w in sorted(box, key=lambda a: -a[1]) if w[1] > 2290), None)
        toks = [w[0] for w in sorted(box, key=lambda a: (round(a[2] / 6), a[1]))]
        date = ""
        for i, t in enumerate(toks):
            if re.match(r"^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)'?$", t, re.I):
                date = t + (toks[i + 1] if i + 1 < len(toks) and re.match(r"^'?\d{2}$", toks[i + 1]) else "")
                break
        return {"id": sid, "rev": rev, "date": date}
    except Exception as ex:
        return {"error": str(ex)[:100]}

def main():
    root, out = sys.argv[1], sys.argv[2]
    pdfroot = sys.argv[3] if len(sys.argv) > 3 and not sys.argv[3].startswith("--") else None
    workers = int(sys.argv[sys.argv.index("--workers")+1]) if "--workers" in sys.argv else 4
    os.makedirs(out + "/json", exist_ok=True)
    files = sorted(glob.glob(root + "/**/*.dxf", recursive=True))
    print("DXF files:", len(files), flush=True)
    results = []
    with Pool(workers) as pool:
        for i, r in enumerate(pool.imap_unordered(process, files), 1):
            results.append(r)
            if i % 20 == 0:
                print(i, "/", len(files), flush=True)
    with open(out + "/summary.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["file", "ok", "id", "date", "rev", "dims", "override", "mismatch", "viewports", "titles", "unassigned_dims", "styles", "sec", "error"])
        for r in sorted(results, key=lambda a: a["file"]):
            w.writerow([os.path.relpath(r["file"], root), r["ok"], r.get("id"), r.get("date"), r.get("rev"), r.get("n_dims"), r.get("n_override"), r.get("n_mism"),
                        r.get("n_vp"), r.get("n_titles"), r.get("unassigned"), ";".join(f"{s['style']}x{s['lfac']}r{s['rnd']}:{s['n']}" for s in r.get("styles", [])), r["sec"], r.get("error", "")])
            if r["ok"]:
                rel = os.path.splitext(os.path.relpath(r["file"], root))[0].replace("\\", "__").replace("/", "__")
                json.dump(r, open(f"{out}/json/{rel}.json", "w", encoding="utf-8"), ensure_ascii=False)
    if pdfroot:
        pdfs = {}
        for p in glob.glob(pdfroot + "/**/*.pdf", recursive=True):
            m = ID_RE.search(re.sub(r"^p\d+_", "", os.path.splitext(os.path.basename(p))[0]))
            if m: pdfs[m.group(0)] = p
        with open(out + "/version_check.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f); w.writerow(["id", "dwg_rev", "pdf_rev", "dwg_date", "pdf_date", "result"])
            for r in sorted(results, key=lambda a: a.get("id") or ""):
                if not r["ok"] or not r.get("id"): continue
                p = pdfs.get(r["id"])
                if not p: w.writerow([r["id"], r.get("rev"), "", r.get("date"), "", "KHONG CO PDF"]); continue
                pt = pdf_title(p)
                dd = re.sub(r"[\s']", "", str(r.get("date") or "")).upper(); pd = re.sub(r"[\s']", "", pt.get("date", "")).upper()
                same = (str(r.get("rev")) == str(pt.get("rev"))) and (dd == pd or not dd)
                w.writerow([r["id"], r.get("rev"), pt.get("rev"), r.get("date"), pt.get("date"), "KHOP" if same else "LECH"])
    ok = [r for r in results if r["ok"]]
    print("done: ok", len(ok), "/", len(results), "| dims", sum(r["n_dims"] for r in ok), "| overrides", sum(r["n_override"] for r in ok), "| mismatch", sum(r["n_mism"] for r in ok))

if __name__ == "__main__":
    main()
