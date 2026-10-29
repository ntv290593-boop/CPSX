"""dwg_extract.py - doc kich thuoc va nhan tu DXF (ODA) bang ezdxf.
Gia tri that = get_measurement() * DIMLFAC. Kich thuoc nam o paper space; gan vao viewport theo vi tri.
Dung:  python -I dwg_extract.py <file.dxf> [--json out.json]"""
import ezdxf, sys, json, re, collections
sys.stdout.reconfigure(encoding="utf-8")
P = lambda v: (round(v[0]), round(v[1]))

def strip_mtext(s):
    s = re.sub(r"\\[A-Za-z][^;]*;", "", s or "")
    return re.sub(r"[{}]|\\P", " ", s).strip()

def extract(path):
    d = ezdxf.readfile(path)
    out = {"file": path, "layouts": [], "viewports": [], "dims": [], "texts": [], "warn": []}
    for lay in d.layouts:
        if lay.name == "Model":
            continue
        out["layouts"].append(lay.name)
        vps = []
        for v in lay.query("VIEWPORT"):
            if v.dxf.id == 1 or not v.dxf.view_height:
                continue
            c = v.dxf.center; w, h = v.dxf.width, v.dxf.height
            vps.append({"id": v.dxf.id, "x0": c[0]-w/2, "x1": c[0]+w/2, "y0": c[1]-h/2, "y1": c[1]+h/2,
                        "scale": round(h / v.dxf.view_height, 4)})
        out["viewports"] += [{"layout": lay.name, **v} for v in vps]
        def vp_of(p):
            for v in vps:
                if v["x0"] <= p[0] <= v["x1"] and v["y0"] <= p[1] <= v["y1"]:
                    return v["id"]
            return None
        for e in lay.query("DIMENSION"):
            ds = d.dimstyles.get(e.dxf.dimstyle)
            lf = ds.dxf.get("dimlfac", 1.0) if ds else 1.0
            dec = ds.dxf.get("dimdec", 0) if ds else 0
            try:
                m = e.get_measurement()
            except Exception:
                m = None
            if not isinstance(m, (int, float)):
                out["warn"].append(f"dim khong do duoc {e.dxf.handle}")
                continue
            val = m * lf
            txt = e.dxf.get("text", "")
            ov = txt not in ("", "<>")
            disp = txt if ov else None
            if ov:
                try:
                    num = float(re.sub(r"[^\d.]", "", txt) or "nan")
                    mism = abs(num - val) > 0.51
                except ValueError:
                    mism = None
            else:
                mism = False
            out["dims"].append({"layout": lay.name, "vp": vp_of(e.dxf.defpoint), "val": round(val, 2), "meas": round(m, 3),
                                "lfac": lf, "override": disp, "override_mismatch": mism,
                                "x": round(e.dxf.defpoint[0]), "y": round(e.dxf.defpoint[1]), "handle": e.dxf.handle,
                                "angle": round(float(getattr(e.dxf, "angle", 0) or 0), 1)})
        for e in list(lay.query("MTEXT")) + list(lay.query("TEXT")):
            t = strip_mtext(e.text if e.dxftype() == "MTEXT" else e.dxf.text)
            if t:
                ins = e.dxf.insert
                out["texts"].append({"layout": lay.name, "vp": vp_of(ins), "text": t, "x": round(ins[0]), "y": round(ins[1])})
    return out

if __name__ == "__main__":
    r = extract(sys.argv[1])
    if "--json" in sys.argv:
        json.dump(r, open(sys.argv[sys.argv.index("--json") + 1], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(r["file"].split("\\")[-1].split("/")[-1], "layouts", r["layouts"], "viewports", len(r["viewports"]),
          "dims", len(r["dims"]), "overrides", sum(1 for x in r["dims"] if x["override"]),
          "mismatch", sum(1 for x in r["dims"] if x["override_mismatch"]), "warn", len(r["warn"]))
