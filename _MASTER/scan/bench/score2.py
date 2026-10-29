import json,sys,os,re,collections
sys.stdout.reconfigure(encoding="utf-8")
B=sys.argv[1]; crops=json.load(open(B+"/bench.json",encoding="utf-8"))
def parse(path):
    t=open(path,encoding="utf-8-sig",errors="ignore").read(); m=re.search(r"\[[^\[\]]*\]",t,re.S)
    return [re.sub(r"[\s,.]","",str(x)) for x in json.loads(m.group(0))] if m else []
C=collections.Counter
S={"decoder":lambda c:C(c["decoder"]),"gemini":lambda c:C(x for x in parse(f"{B}/out/gemini-3.8-flash-low/{c['name']}.txt") if x!="?"),"claude":lambda c:C(x for x in parse(f"{B}/out/claude-sonnet-4-6/{c['name']}.txt") if x!="?")}
agg=collections.defaultdict(lambda:[0,0,0]); common_miss=collections.Counter(); percrop=[]
for c in crops:
    E=C(c["expected"]); R={k:f(c) for k,f in S.items()}
    keys=set(E)|set().union(*[set(r) for r in R.values()])
    U=C({k:min(R[s][k] for s in R) for k in keys}); M=C({k:sorted(R[s][k] for s in R)[1] for k in keys}); X=C({k:max(R[s][k] for s in R) for k in keys})
    Eadj=C({k:min(E[k],X[k]) for k in E}); noise=sum(((E-X)).values()); common_miss.update(E-X)
    for name,Rr in list(R.items())+[("unanime (3/3)",U),("da so (>=2/3)",M)]:
        a=agg[name]; a[0]+=sum((E&Rr).values()); a[1]+=sum(Rr.values()); a[2]+=sum(E.values())
    for name,Rr in R.items():
        a=agg[name+"|nhan hieu chinh"]; a[0]+=sum((Eadj&Rr).values()); a[1]+=sum(Rr.values()); a[2]+=sum(Eadj.values())
    for name,Rr in (("unanime (3/3)",U),("da so (>=2/3)",M)):
        a=agg[name+"|nhan hieu chinh"]; a[0]+=sum((Eadj&Rr).values()); a[1]+=sum(Rr.values()); a[2]+=sum(Eadj.values())
    percrop.append((c["name"],sum(E.values()),noise,sum((X-U).values())))
print("So ky vong khong nguon nao thay (nghi nhan DWG co kich thuoc khong in ra):",sum(common_miss.values()),"/",sum(sum(C(c["expected"]).values()) for c in crops))
print("  gia tri hay gap:",common_miss.most_common(8))
print(f"{'nguon':36s} {'phu':>7s} {'chinh xac':>10s}   (n khop / n bao / n ky vong)")
for k,(m,g,e) in agg.items(): print(f"{k:36s} {m/e:7.1%} {m/g:10.1%}   {m}/{g}/{e}")
print("Crop nhieu so khong nguon nao thay:",sorted([(n,x) for n,e,x,_ in percrop if x],key=lambda a:-a[1])[:6])
print("So gia tri 3 nguon bat dong (can xem lai) tren tong ky vong:",sum(d for *_,d in percrop),"/",sum(e for _,e,_,_ in percrop))
