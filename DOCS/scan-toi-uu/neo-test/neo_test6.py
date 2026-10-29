src=open(r"D:\COSTING EXCEL\DOCS\scan-toi-uu\neo-test\neo_test5.py",encoding='utf-8').read().split("stat=collections")[0]
exec(src)
out=[]
for r,sz in pick:
    zp=os.path.join(INP,r["Đường dẫn"])
    try: d=pymupdf.open(stream=zs[zp].read(r["Trong zip"]),filetype="pdf")
    except Exception: continue
    n=d.page_count
    for i in sorted(set([0,n//3,2*n//3,n-1])):
        p=d[i]; t=p.get_text(); L=len(t.strip())
        if L>=20 and len(re.findall(r"\b\d{3,4}\b",t))<5:
            dr=len(p.get_drawings())
            out.append((os.path.basename(r["Trong zip"])[:46],i+1,L,dr,t.strip().replace("\n"," | ")[:90]))
print(len(out))
for o in out[:26]: print(o)
