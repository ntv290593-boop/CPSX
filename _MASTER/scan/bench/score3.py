import json,sys,os,re,collections
sys.stdout.reconfigure(encoding="utf-8")
B=sys.argv[1]; crops=json.load(open(B+"/bench.json",encoding="utf-8"))
def parse(path):
    t=open(path,encoding="utf-8-sig",errors="ignore").read(); m=re.search(r"\[[^\[\]]*\]",t,re.S)
    return [re.sub(r"[\s,]","",str(x)) for x in json.loads(m.group(0))] if m else []
C=collections.Counter
raw={"gemini":lambda c:parse(f"{B}/out/gemini-3.8-flash-low/{c['name']}.txt"),"claude":lambda c:parse(f"{B}/out/claude-sonnet-4-6/{c['name']}.txt")}
for name,f in raw.items():
    ex_all=C(); q=0; rd=0; tot=0
    for c in crops:
        g=[x for x in f(c) if x!="?"]; q+=sum(1 for x in f(c) if x=="?"); tot+=len(g)
        E=C(c["expected"]); G=C(re.sub(r"\.","",x) for x in g); ex=G-E
        ex_all.update(ex)
    nonnum={k:v for k,v in ex_all.items() if not k.isdigit()}
    num={k:v for k,v in ex_all.items() if k.isdigit()}
    print(name,"| so bao",tot,"| '?' (khong doc duoc):",q,"| thua khong phai so thuan (R.., c/c..):",sum(nonnum.values()),dict(list(nonnum.items())[:8]),"| thua la so thuan:",sum(num.values()))
