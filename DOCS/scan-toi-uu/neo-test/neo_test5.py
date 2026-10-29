src=open(r"D:\COSTING EXCEL\DOCS\scan-toi-uu\neo-test\neo_test.py",encoding='utf-8').read().split("# ---- bài A")[0]
exec(src)
import collections, os, zipfile, re
zs={}
cand=[]
for r in fi.values():
    if r.get("Loại")=="pdf" and r.get("Trong zip"):
        zp=os.path.join(INP,r["Đường dẫn"])
        if zp not in zs:
            try: zs[zp]=zipfile.ZipFile(zp)
            except Exception: zs[zp]=None
        z=zs[zp]
        if not z: continue
        try: sz=z.getinfo(r["Trong zip"]).file_size
        except KeyError: continue
        if sz<60e6: cand.append((r,sz))
print("PDF trong zip <60MB:",len(cand),"| tổng PDF trong zip:",sum(1 for r in fi.values() if r.get('Loại')=='pdf' and r.get('Trong zip')))
by=collections.defaultdict(list)
for r,sz in cand:
    key=(r.get("Đợt"),re.sub(r"[^/]*$","",r["Trong zip"])[:60])
    by[key].append((r,sz))
random.seed(3); pick=[]
for k,v in by.items(): pick+=random.sample(v,min(3,len(v)))
pick=pick[:70]
print("nhóm thư mục:",len(by),"| file thử:",len(pick))
stat=collections.defaultdict(lambda:[0,0,0,0])  # trang, khongtext, ittext_ngan, co_so_kich_thuoc
rows=[]
for r,sz in pick:
    try: d=pymupdf.open(stream=zs[os.path.join(INP,r["Đường dẫn"])].read(r["Trong zip"]),filetype="pdf")
    except Exception as e: continue
    n=d.page_count; idx=sorted(set([0,n//3,2*n//3,n-1]))
    a=b=c=0
    for i in idx:
        t=d[i].get_text(); L=len(t.strip())
        if L<20: a+=1
        else:
            nd=len(re.findall(r"\b\d{3,4}\b",t))
            if nd>=5: c+=1
            else: b+=1
    k=(r.get("Đợt"),r["Trong zip"].split("/")[0][:40])
    s=stat[k]; s[0]+=len(idx); s[1]+=a; s[2]+=b; s[3]+=c
print(f"\n{'đợt':16}{'thư mục gốc trong zip':44}{'trang':>6}{'0text':>6}{'ít':>5}{'có số':>6}")
T=[0,0,0,0]
for k,s in sorted(stat.items()):
    print(f"{str(k[0])[:15]:16}{k[1]:44}{s[0]:>6}{s[1]:>6}{s[2]:>5}{s[3]:>6}")
    for i in range(4): T[i]+=s[i]
print(f"TỔNG: {T[0]} trang | không text {T[1]} ({100*T[1]/T[0]:.0f}%) | text ít/ không có số kích thước {T[2]} ({100*T[2]/T[0]:.0f}%) | có text + số 3-4 chữ số {T[3]} ({100*T[3]/T[0]:.0f}%)")
