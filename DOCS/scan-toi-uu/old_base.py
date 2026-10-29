"""Logic cũ của ungvien (backup pd_2026-10-27) áp lên cùng bộ đo, để so công bằng."""
import re, sys, bench_ext as B, pd
def old_flat(ix, allm, g):
    out={}
    for m in g:
        r=allm[m]["row"]; hm=r.get("Hạng mục") or ""
        kws=sorted({w for w in re.findall(r"[a-z0-9]{3,}", pd._fold(hm)) if w not in pd.STOP})
        mrx=re.compile(r"(?<![a-z0-9])"+re.escape(pd._fold(m))+r"(?![a-z0-9])")
        res=[]
        for p in ix["pages"]:
            t=pd._fold(ix["texts"][p[0]][p[1]-1]); mh=len(mrx.findall(t)); kh=[w for w in kws if w in t]
            if mh or len(kh)>=min(2,len(kws)) and kh: res.append((mh*5+len(kh),p))
        out[m]=[p for s,p in sorted(res,key=lambda x:-x[0])]
    return out
orig=B.run
def run(title,P,sheets,g,tmp,keep=lambda p:p,bq=None):
    ix,allm,g2,flat=orig(title,P,sheets,g,tmp,keep,bq)
    of=old_flat(ix,allm,g2); tot=sum(map(len,g2.values()))
    f=lambda k:sum(1 for m in g2 for p in g2[m] if p in set(of[m][:k]))/tot
    print(f"   LOGIC CŨ cùng bộ đo: @5 {f(5):.0%} @12 {f(12):.0%} @40 {f(40):.0%} hết {f(10**6):.0%}")
B.run=run
for a in sys.argv[1:]: {"sofitel":B.sofitel,"d01":B.d01,"vanity":B.vanity}[a]()
