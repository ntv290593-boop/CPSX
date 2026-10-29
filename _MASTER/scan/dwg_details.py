import ezdxf,json,sys,collections
sys.stdout.reconfigure(encoding="utf-8")
D=sys.argv[1]; k=sys.argv[2]
r=json.load(open(f"{D}/{k}.json",encoding="utf-8"))
d=ezdxf.readfile(f"{D}/out/{k}.dxf"); ps=d.layout("A1 PLAN")
titles=[]
for e in ps.query("INSERT"):
    at={a.dxf.tag:a.dxf.text for a in e.attribs}
    if "TITLE" in at and "N\u00ba" in at: titles.append({"no":at["N\u00ba"],"title":at["TITLE"],"scale":at.get("SCALE"),"x":e.dxf.insert[0],"y":e.dxf.insert[1]})
vps=r["viewports"]
def detail_of_vp(v):
    cx=(v["x0"]+v["x1"])/2
    best=None
    for t in titles:
        if t["y"]<=v["y0"]+5 and v["x0"]-40<=t["x"]<=v["x1"]:
            gap=v["y0"]-t["y"]
            if best is None or gap<best[0]: best=(gap,t)
    return best[1] if best else None
vp2d={v["id"]:detail_of_vp(v) for v in vps}
groups=collections.defaultdict(list)
for x in r["dims"]:
    v=x["vp"]
    if v is None:  # fallback: nearest viewport centre
        v=min(vps,key=lambda a:(( (a["x0"]+a["x1"])/2-x["x"])**2+((a["y0"]+a["y1"])/2-x["y"])**2))["id"]; src="gan-gan-nhat"
    else: src="trong-viewport"
    t=vp2d.get(v)
    groups[(t["no"],t["title"]) if t else ("?",f"vp{v}")].append((x["val"],src[0]))
for key in sorted(groups):
    vals=sorted(groups[key],key=lambda a:a[0])
    print(key,"n=",len(vals),[ (round(a,1) if a%1 else int(a))*(1) for a,_ in vals][:30])
print("titles found",len(titles),"| viewports",len(vps))
