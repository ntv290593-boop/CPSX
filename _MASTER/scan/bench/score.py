import json,sys,os,re,collections,glob
sys.stdout.reconfigure(encoding="utf-8")
B=sys.argv[1]
crops=json.load(open(B+"/bench.json",encoding="utf-8"))
def parse(path):
    if not os.path.exists(path): return None
    t=open(path,encoding="utf-8-sig",errors="ignore").read()
    m=re.search(r"\[[^\[\]]*\]",t,re.S)
    if not m: return None
    try: arr=json.loads(m.group(0))
    except Exception: return None
    return [re.sub(r"[\s,.]","",str(x)) for x in arr]
def score(exp,got):
    ce,cg=collections.Counter(exp),collections.Counter(got)
    match=sum((ce&cg).values()); return match,sum(ce.values()),sum(cg.values())
models=sorted(d for d in os.listdir(B+"/out")) if os.path.isdir(B+"/out") else []
sources=["decoder"]+models
tot={s:[0,0,0,0,0,0] for s in sources}  # match, exp, got, crops_ok, crops_exact, crops_scored
extra={s:collections.Counter() for s in sources}; miss={s:collections.Counter() for s in sources}
agree={m:[0,0,0] for m in models}   # intersection decoder&model: size, correct, exp
for c in crops:
    exp=[x for x in c["expected"]]
    outs={"decoder":c["decoder"]}
    for m in models:
        g=parse(f"{B}/out/{m}/{c['name']}.txt"); outs[m]=g
    for s in sources:
        g=outs[s]
        if g is None: continue
        g=[x for x in g if x!="?"]
        match,ne,ng=score(exp,g); t=tot[s]; t[0]+=match;t[1]+=ne;t[2]+=ng;t[5]+=1; t[4]+= (collections.Counter(exp)==collections.Counter(g))
        ce,cg=collections.Counter(exp),collections.Counter(g)
        for k,v in (cg-ce).items(): extra[s][k]+=v
        for k,v in (ce-cg).items(): miss[s][k]+=v
    for m in models:
        if outs[m] is None: continue
        g=[x for x in outs[m] if x!="?"]; inter=collections.Counter(c["decoder"])&collections.Counter(g)
        a=agree[m]; a[0]+=sum(inter.values()); a[1]+=sum((inter&collections.Counter(exp)).values()); a[2]+=len(exp)
print("crops:",len(crops))
print(f"{'nguon':28s} {'crop':>5s} {'phu(recall)':>12s} {'chinh xac':>10s} {'crop khop hoan toan':>20s} {'so thua':>8s} {'so thieu':>9s}")
for s in sources:
    mt,ne,ng,_,ex,ns=tot[s]
    if ns==0: print(s,"khong co ket qua"); continue
    print(f"{s:28s} {ns:5d} {mt/ne:11.1%} {mt/ng:10.1%} {ex:12d}/{ns:<7d} {ng-mt:8d} {ne-mt:9d}")
print()
for m in models:
    a=agree[m]
    if a[0]: print(f"giao cua decoder va {m}: {a[0]} so, dung {a[1]/a[0]:.2%} (phu {a[1]/a[2]:.1%} tong so ky vong)")
for s in sources:
    if tot[s][5]: print(s,"so thua pho bien:",extra[s].most_common(6),"| so thieu pho bien:",miss[s].most_common(6))
