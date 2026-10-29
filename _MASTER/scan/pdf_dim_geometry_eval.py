import sys,os,re,glob,csv,math,random,collections
sys.path.insert(0,sys.argv[1]); sys.stdout.reconfigure(encoding="utf-8")
import pymupdf,ezdxf
from ezdxf.entities.dimstyleoverride import DimStyleOverride
from multiprocessing import Pool
from pdf_dim_decode import Decoder,decode_page
from pdf_dim_geometry import find_ticks,page_scales,scale_labels,scale_for,TickIndex,candidates,check,STD
A=72/25.4
TOL=float(sys.argv[4]) if len(sys.argv)>4 else 3.0
LO=float(sys.argv[5]) if len(sys.argv)>5 else 0.0
REACH=float(sys.argv[6]) if len(sys.argv)>6 else 16.0
NEAR=float(sys.argv[7]) if len(sys.argv)>7 else 60.0
CODE=re.compile(r"(ID-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{3}[A-Z]?)")
def init(npz):
    global DEC; DEC=Decoder(npz)
def corrupt(s,rng):
    i=rng.randrange(len(s)); ds="0123456789".replace(s[i],""); t=s[:i]+rng.choice(ds)+s[i+1:]
    return t if not t.startswith("0") else s+"0"
def work(job):
    code,dxf,pdf=job; rng=random.Random(sum(map(ord,code)))
    page=pymupdf.open(pdf)[0]; H=page.rect.height
    idx=TickIndex(find_ticks(page)); labels=scale_labels(page); scales=sorted({n for _,_,n in labels}) or list(STD)
    res=decode_page(page,DEC)
    d=ezdxf.readfile(dxf); P=[]
    for l in d.layouts:
        if l.name=="Model": continue
        for e in l.query("DIMENSION"):
            tm=e.dxf.get("text_midpoint")
            if tm is None or e.dimtype!=0: continue
            P.append((A*tm[0],H-A*tm[1],float(DimStyleOverride(e).get("dimlfac",1) or 1),e.dxf.dimstyle,e.dxf.get("angle",0)))
    out=[]
    for c in res:
        if not c["accepted"] or not c["text"].isdigit(): continue
        near=[(math.hypot(px-c["cx"],py-c["cy"]),lf,st_,an_) for px,py,lf,st_,an_ in P if math.hypot(px-c["cx"],py-c["cy"])<9]
        if not near: continue
        truelf=min(near)[1]; stname=min(near)[2]; ang=min(near)[3]
        cands=candidates(idx,c["orient"],c["cx"],c["cy"],reach=REACH,near=NEAR,lo=LO); encl=[x for x in cands if x["enclosing"]]
        sc=truelf; one=[sc]
        n=int(c["text"]); wrong=int(corrupt(c["text"],rng))
        wrong=int(corrupt(c["text"],random.Random(n*7+len(cands))))
        out.append((check(n,cands,one,tol=TOL)[0],check(wrong,cands,one,tol=TOL)[0],len(str(n))))
    return code,out
if __name__=="__main__":
    R,npz=sys.argv[2],sys.argv[3]
    ok=[r["ma"] for r in csv.DictReader(open(R+"/all_result/version_check2.csv",encoding="utf-8-sig")) if r["ket_qua"]=="KHOP"]
    dx={};pd={}
    for p in glob.glob(R+"/all/**/*.dxf",recursive=True):
        c=(CODE.search(os.path.basename(p).upper()) or [None,None])[1]; dx.setdefault(c,[]).append(p)
    for p in glob.glob("D:/COSTING EXCEL/_ACTIVE/SOFITEL/01_scan/_sheets/**/*.pdf",recursive=True):
        c=(CODE.search(os.path.basename(p).upper()) or [None,None])[1]; pd.setdefault(c,[]).append(p)
    random.seed(5); sel=random.sample([c for c in ok if len(dx.get(c,[]))==1 and c in pd],36)
    jobs=[(c,dx[c][0],pd[c][0]) for c in sel]
    by=collections.defaultdict(lambda:[0,0,0])
    with Pool(6,initializer=init,initargs=(npz,)) as pool:
        for code,out in pool.imap_unordered(work,jobs):
            for st,stw,ln in out:
                k="%d chu so"%min(ln,4)+("+" if ln>=4 else ""); b=by[k]; b[0]+=1; b[1]+=st.startswith("OK"); b[2]+=stw.startswith("OK")
    for k in sorted(by): n,a,w=by[k]; print(k,"n=",n,"| xac nhan so DUNG: %.1f%%"%(100*a/n),"| so SAI cung duoc xac nhan: %.1f%%"%(100*w/n))

