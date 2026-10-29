import sys,os,re,glob,csv,math,pickle,random,collections,json
sys.path.insert(0,sys.argv[1]); sys.stdout.reconfigure(encoding="utf-8")
import numpy as np,pymupdf,ezdxf
from ezdxf.entities.dimstyleoverride import DimStyleOverride
from pdf_dim_decode import Decoder,decode_page
A=72/25.4
VEC,DWG,OUT=sys.argv[1],sys.argv[2],sys.argv[3]
R=[r for r in pickle.load(open(VEC+"/dataset3.pkl","rb")) if not r["err"]]
random.seed(1); codes=sorted(r["code"] for r in R); random.shuffle(codes); test=set(codes[:len(codes)//3]); train=set(codes[len(codes)//3:])
# templates tu cac to train
raw=[]
for r in R:
    if r["code"] in train:
        for d in r["dims"]:
            if "imgs" in d and d["n_glyph"]==len(d["S"]):
                for ch,im,n in zip(d["S"],d["imgs"],d["items"]): raw.append((ch,d["orient"],n,im.reshape(-1).astype(np.uint8)))
cnt=collections.defaultdict(collections.Counter)
for ch,o,n,_ in raw: cnt[ch][n]+=1
modal={ch:cnt[ch].most_common(1)[0][0] for ch in cnt}
seen={}
for ch,o,n,im in raw:
    if n==modal[ch]: seen[(ch,o,n,im.tobytes())]=im
keys=list(seen)
np.savez_compressed(OUT+"/templates_train.npz",X=np.array([seen[k] for k in keys],dtype=np.uint8),L=np.array([k[0] for k in keys]),O=np.array([k[1] for k in keys]),N=np.array([k[2] for k in keys]))
DEC=Decoder(OUT+"/templates_train.npz")
CODE=re.compile(r"(ID-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{3}[A-Z]?)")
dx={};pd={}
for p in glob.glob(DWG+"/all/**/*.dxf",recursive=True):
    c=(CODE.search(os.path.basename(p).upper()) or [None,None])[1]; dx.setdefault(c,[]).append(p)
for p in glob.glob("D:/COSTING EXCEL/_ACTIVE/SOFITEL/01_scan/_sheets/**/*.pdf",recursive=True):
    c=(CODE.search(os.path.basename(p).upper()) or [None,None])[1]; pd.setdefault(c,[]).append(p)
cand=[c for c in sorted(test) if len(dx.get(c,[]))==1 and c in pd]
random.seed(11); random.shuffle(cand)
crops=[]; used=0
for code in cand:
    if used>=14 or len(crops)>=36: break
    page=pymupdf.open(pd[code][0])[0]; H=page.rect.height
    d=ezdxf.readfile(dx[code][0]); lays=[l for l in d.layouts if l.name!="Model"]
    dims=[]
    for l in lays:
        for e in l.query("DIMENSION"):
            tm=e.dxf.get("text_midpoint")
            if tm is None: continue
            ov=DimStyleOverride(e); lf=ov.get("dimlfac",1) or 1; rnd=ov.get("dimrnd",0) or 0; dec=ov.get("dimdec",0) or 0
            txt=e.dxf.get("text","") or ""
            if txt not in("","<>"):
                s=txt.strip(); exp=s if s.isdigit() else None
            else:
                v=e.get_measurement()*lf
                if rnd: v=math.floor(v/rnd+0.5)*rnd
                exp=str(int(math.floor(v+0.5))) if dec==0 else "DEC"
            dims.append((A*tm[0],H-A*tm[1],exp))
    vps=[]
    for l in lays:
        for v in l.query("VIEWPORT"):
            if v.dxf.id==1 or not v.dxf.view_height: continue
            c,w,h=v.dxf.center,v.dxf.width,v.dxf.height
            vps.append((v.dxf.id,A*(c[0]-w/2),H-A*(c[1]+h/2),A*(c[0]+w/2),H-A*(c[1]-h/2)))
    res=None; picked=0
    random.shuffle(vps)
    for vid,x0,y0,x1,y1 in vps:
        m=30
        rect=(max(0,x0-m),max(0,y0-m),min(page.rect.width,x1+m),min(H,y1+m))
        if rect[2]-rect[0]<120 or rect[3]-rect[1]<120 or rect[2]-rect[0]>900 or rect[3]-rect[1]>900: continue
        inside=[(x,y,e) for x,y,e in dims if rect[0]<=x<=rect[2] and rect[1]<=y<=rect[3]]
        if len(inside)<8 or len(inside)>45 or any(e=="DEC" for _,_,e in inside): continue
        exp=sorted(e for _,_,e in inside if e)
        if res is None: res=decode_page(page,DEC)
        dec_out=sorted(c_["text"] for c_ in res if c_["accepted"] and rect[0]<=c_["cx"]<=rect[2] and rect[1]<=c_["cy"]<=rect[3])
        name=f"{code}_vp{vid}"
        page.get_pixmap(dpi=200,clip=pymupdf.Rect(*rect)).save(f"{OUT}/crops/{name}.png")
        crops.append({"name":name,"code":code,"rect":[round(q,1) for q in rect],"expected":exp,"decoder":dec_out,"n_nonnumeric_skipped":sum(1 for _,_,e in inside if not e)})
        picked+=1
        if picked>=3: break
    if picked: used+=1
json.dump(crops,open(OUT+"/bench.json","w",encoding="utf-8"),ensure_ascii=False,indent=0)
print("crops",len(crops),"sheets",used,"| expected values",sum(len(c["expected"]) for c in crops),"| decoder outputs",sum(len(c["decoder"]) for c in crops))
print([ (c["name"],len(c["expected"])) for c in crops][:8])
