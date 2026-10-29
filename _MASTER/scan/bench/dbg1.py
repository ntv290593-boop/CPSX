import sys,glob,os,math,pymupdf,ezdxf,collections
sys.path.insert(0,sys.argv[1]); sys.stdout.reconfigure(encoding="utf-8")
from pdf_glyphs import glyph_paths,cluster
code="ID-GR-TY-701"
pdf=[p for p in glob.glob("D:/COSTING EXCEL/_ACTIVE/SOFITEL/01_scan/_sheets/**/*.pdf",recursive=True) if code in os.path.basename(p)][0]
dxf=[p for p in glob.glob(sys.argv[2]+"/all/**/*.dxf",recursive=True) if code in os.path.basename(p)][0]
page=pymupdf.open(pdf)[0]; H=page.rect.height; A=72/25.4
G=glyph_paths(page); C=cluster(G)
d=ezdxf.readfile(dxf); ps=d.layout("A1 PLAN")
n=0
for e in ps.query("DIMENSION"):
    ds=d.dimstyles.get(e.dxf.dimstyle); lf=ds.dxf.get("dimlfac",1.0); rnd=ds.dxf.get("dimrnd",0) or 0
    v=e.get_measurement()*lf
    if rnd: v=round(v/rnd)*rnd
    if round(v)!=110 or e.dxf.get("text","") not in("","<>"): continue
    tm=e.dxf.text_midpoint; px,py=A*tm[0],H-A*tm[1]
    c=min(C,key=lambda c:(c["cx"]-px)**2+(c["cy"]-py)**2)
    print("dim 110 style",e.dxf.dimstyle,"rnd",rnd,"meas*lf",round(e.get_measurement()*lf,2),"text_mid px,py",round(px),round(py),"| cluster",c["orient"],round(c["cx"]),round(c["cy"]),[ (g["n"],tuple(round(x,1) for x in g["rect"])) for g in c["glyphs"]])
    # all strokes within 14pt regardless of filters
    near=[(dd.get("type"),round(dd.get("width") or 0,2),len(dd["items"]),tuple(round(x,1) for x in (dd["rect"].x0,dd["rect"].y0,dd["rect"].x1,dd["rect"].y1))) for dd in page.get_drawings() if dd["rect"].intersects(pymupdf.Rect(px-14,py-10,px+14,py+10)) and dd["rect"].width<12 and dd["rect"].height<12]
    for x in near[:12]: print("    ",x)
    n+=1
    if n>=2: break
