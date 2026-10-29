src=open(r"D:\COSTING EXCEL\DOCS\scan-toi-uu\neo-test\neo_test2.py",encoding='utf-8').read().split('print("=== B2.')[0]
exec(src)
def page_nums(fid,pg):
    t=" ".join(w[4] for w in open_fid(fid)[pg-1].get_text("words"))
    return set(re.findall(r"(?<![\d.])\d+(?:\.\d+)?(?![\d])",t))
DIM=re.compile(r"^\d[\d.]*(?:\s*mm)?(?:\s*[x×–\-]\s*\d[\d.]*(?:\s*mm)?)+$",re.I)
def chk2(vals,fid,pg):
    ps,n=page_sq(fid,pg)
    if n<20: return "KHONG_KIEM_DUOC",[]
    nums=page_nums(fid,pg); need=[]; miss=[]
    for v in vals:
        for c in codes(v):
            need.append(c)
            if DIM.match(c):
                parts=re.findall(r"\d+(?:\.\d+)?",c)
                if any(p not in nums for p in parts): miss.append(c)
            elif sq(c) not in ps: miss.append(c)
    if not need: return "KHONG_CO_MA",[]
    return ("KHOP" if not miss else "LECH"),miss
rows=[r for r in mat if isinstance(r.get("Nguồn spec"),list) and len(r["Nguồn spec"])>=3 and isinstance(r["Nguồn spec"][2],int) and r["Nguồn spec"][1] in fi]
c=collections.Counter()
for r in rows:
    _,fid,pg=r["Nguồn spec"][:3]
    s,m=chk2([r["Mã VL"],r.get("NCC/Mã SP"),r.get("Mô tả")],fid,pg); c[s]+=1
    if s=="LECH": print("LECH",r["Mã VL"],pg,m,"| Mô tả:",r["Mô tả"])
print(dict(c))
r=[x for x in mat if x["Mã VL"]=="SW-24"][0]
t=" ".join(w[4] for w in open_fid("F0043")[25].get_text("words")); print("SW-24 p26:",t[:700])
