import pickle,sys,collections,random,numpy as np
sys.stdout.reconfigure(encoding="utf-8")
R=[r for r in pickle.load(open(sys.argv[1],"rb")) if not r["err"]]
seed=int(sys.argv[2]); TH=float(sys.argv[3])
random.seed(seed); codes=sorted(r["code"] for r in R); random.shuffle(codes); test=set(codes[:len(codes)//3]); train=set(codes[len(codes)//3:])
raw=[]
for r in R:
    if r["code"] in train:
        for d in r["dims"]:
            if "imgs" in d and d["n_glyph"]==len(d["S"]):
                for ch,im,n in zip(d["S"],d["imgs"],d["items"]): raw.append((ch,d["orient"],n,im.reshape(-1).astype(np.float32)))
modal={}
cnt=collections.defaultdict(collections.Counter)
for ch,o,n,_ in raw: cnt[ch][n]+=1
for ch in cnt: modal[ch]=cnt[ch].most_common(1)[0][0]
clean=[x for x in raw if x[2]==modal[x[0]]]
print("mau train tho",len(raw),"sau loc theo so doan",len(clean),"| modal items",modal)
X=np.array([x[3] for x in clean]); O=np.array([x[1] for x in clean]); L=np.array([x[0] for x in clean]); N=np.array([x[2] for x in clean])
def decode(d):
    out=[]
    for im,n in zip(d["imgs"],d["items"]):
        m=(O==d["orient"])&(N==n)
        if not m.any(): out.append(("?",1e9)); continue
        Xm=X[m]; Lm=L[m]; dist=((Xm-im.reshape(-1).astype(np.float32))**2).sum(1); i=dist.argmin(); out.append((Lm[i],float(dist[i])))
    return out
cat=collections.Counter(); wrong=[]
for r in R:
    if r["code"] not in test: continue
    for d in r["dims"]:
        if "S" not in d: continue
        if "imgs" not in d: cat["D khong tim thay cum chu so"]+=1; continue
        dec=decode(d); S="".join(c for c,_ in dec); mx=max(x for _,x in dec)
        if mx>TH: cat["C tu choi (glyph la)"]+=1
        elif S==d["S"]: cat["A chap nhan, dung"]+=1
        else:
            cat["B chap nhan, SAI"]+=1
            cat["   trong do: cung so chu so (nghi doc sai that)" if len(S)==len(d["S"]) else "   trong do: khac so chu so (cum bi ghep/nhan lech)"]+=1
            if len(wrong)<14: wrong.append((r["code"],d["S"],S,d["items"],d["orient"]))
tot=sum(cat.values()); print("test sheets",len(test),"quan the",tot)
for k,v in sorted(cat.items()): print(f"  {k}: {v} ({v/tot:.1%})")
a=cat["A chap nhan, dung"]; ad=a+cat["B chap nhan, SAI"]
print(f"Chinh xac trong so duoc chap nhan: {a}/{ad} = {a/ad:.2%} | ty le chap nhan: {ad/tot:.1%}")
for w in wrong: print("   sai:",w)
