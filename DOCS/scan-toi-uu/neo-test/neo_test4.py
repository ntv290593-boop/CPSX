src=open(r"D:\COSTING EXCEL\DOCS\scan-toi-uu\neo-test\neo_test3.py",encoding='utf-8').read().split("c=collections.Counter()")[0]
exec(src)
print("=== C3. Đột biến với bản kiểm có xử lý kích thước ===")
cau=tot=0; lot=[]
for r in rows:
    _,fid,pg=r["Nguồn spec"][:3]
    for c in codes(r["Mã VL"])+codes(r.get("NCC/Mã SP"))+codes(r.get("Mô tả")):
        if chk2([c],fid,pg)[0]!="KHOP": continue
        d=[i for i,ch in enumerate(c) if ch.isdigit()]
        for _ in range(3):
            i=random.choice(d); n=random.choice([x for x in "0123456789" if x!=c[i]]); m=c[:i]+n+c[i+1:]; tot+=1
            if chk2([m],fid,pg)[0]=="LECH": cau+=1
            else: lot.append((r["Mã VL"],c,m))
print(f"bắt {cau}/{tot} = {100*cau/tot:.1f}%; lọt {len(lot)}")
for x in lot[:12]: print("   lọt",x)
# --- độ phủ text trên bản vẽ thật (file không nằm trong zip, <80MB), lấy mẫu trang
import os,glob
print("\n=== F. Lớp text của bản vẽ thật (mẫu <=6 trang/file, file <80MB, ngoài zip) ===")
pdfs=[]
for r in fi.values():
    if r.get("Loại")=="pdf" and not r.get("Trong zip"):
        p=os.path.join(INP,r["Đường dẫn"])
        if os.path.exists(p) and os.path.getsize(p)<80e6: pdfs.append((r["ID"],p,r.get("Vai trò")))
print("số file thử:",len(pdfs))
tot_p=no_text=low_digit=0; per=[]
for fid,p,vt in pdfs[:40]:
    try: d=pymupdf.open(p)
    except Exception as e: continue
    n=d.page_count; idx=sorted(set([0,n//4,n//2,3*n//4,n-1]))[:6]
    a=b=0
    for i in idx:
        t=d[i].get_text(); tot_p+=1
        if len(t.strip())<20: no_text+=1; a+=1
        elif sum(ch.isdigit() for ch in t)<5 and d[i].rect.width>800: low_digit+=1; b+=1
    per.append((fid,os.path.basename(p)[:50],n,a,b))
print(f"mẫu {tot_p} trang: không text {no_text} ({100*no_text/max(tot_p,1):.0f}%), có text nhưng gần như không có chữ số (nghi số vẽ bằng nét) {low_digit} ({100*low_digit/max(tot_p,1):.0f}%)")
for x in per:
    if x[3] or x[4]: print("  ",x)
