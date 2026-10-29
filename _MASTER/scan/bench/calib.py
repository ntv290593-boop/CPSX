import sys,json,math,collections,pymupdf,ezdxf
sys.path.insert(0,sys.argv[1])
sys.stdout.reconfigure(encoding="utf-8")
from pdf_glyphs import glyph_paths,cluster
D=sys.argv[2]
pdf=pymupdf.open("D:/COSTING EXCEL/_ACTIVE/SOFITEL/01_scan/_sheets/TY-P2/p031_ID-PS-TY-736.pdf")[0]
H=pdf.rect.height
G=glyph_paths(pdf); C=cluster(G)
print("glyph paths",len(G),"clusters",len(C),"multi-glyph",sum(1 for c in C if len(c["glyphs"])>1))
d=ezdxf.readfile(f"{D}/out/ID-PS-TY-736.dxf"); ps=d.layout("A1 PLAN")
M=[]
for e in ps.query("DIMENSION"):
    tm=e.dxf.get("text_midpoint")
    if tm is None: continue
    ds=d.dimstyles.get(e.dxf.dimstyle); lf=ds.dxf.get("dimlfac",1.0); rnd=ds.dxf.get("dimrnd",0) or 0
    v=e.get_measurement()*lf
    if rnd: v=round(v/rnd)*rnd
    M.append({"S":str(int(round(v))),"x":tm[0],"y":tm[1],"ov":e.dxf.get("text","") not in("","<>")})
print("dwg dims with text midpoint",len(M))
a=72/25.4
bins=collections.Counter()
for m in M:
    for c in C:
        ox=c["cx"]-a*m["x"]; oy=(H-c["cy"])-a*m["y"]
        bins[(round(ox/3),round(oy/3))]+=1
top=bins.most_common(4); print("hough top",top)
(bx,by),n=top[0]
ox=bx*3; oy=by*3
# refine + match
def pos(m): return (a*m["x"]+ox, H-(a*m["y"]+oy))
ok=0;samples=[]
for m in M:
    px,py=pos(m)
    best=min(C,key=lambda c:(c["cx"]-px)**2+(c["cy"]-py)**2)
    dist=math.hypot(best["cx"]-px,best["cy"]-py)
    if dist<9:
        ok+=1
        if len(samples)<12: samples.append((m["S"],len(best["glyphs"]),best["orient"],round(dist,1)))
print("matched within 9pt:",ok,"/",len(M)); print(samples)
