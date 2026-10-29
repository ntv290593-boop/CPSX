import re,sys,random,collections,unicodedata
src=open(r"D:\COSTING EXCEL\DOCS\scan-toi-uu\neo-test\neo_test.py",encoding='utf-8').read().split("# ---- bài A")[0]
exec(src)
random.seed(11)
def sq(s):
    s=unicodedata.normalize("NFKC",str(s)).lower()
    return re.sub(r"[^a-z0-9]","",s)
_sq={}
def page_sq(fid,pg):
    k=(fid,pg)
    if k not in _sq:
        t=" ".join(w[4] for w in open_fid(fid)[pg-1].get_text("words"))
        _sq[k]=(sq(t),len(t.strip()))
    return _sq[k]
def codes(v):
    """token có chữ số, >=3 ký tự sau khi bỏ dấu câu; ghép 'A-12' giữ liền"""
    out=[]
    for t in re.split(r"[\s,;()\[\]]+",str(v or "")):
        t=t.strip(".:")
        if len(sq(t))>=3 and any(c.isdigit() for c in t): out.append(t)
    return out
def chk(vals,fid,pg):
    ps,n=page_sq(fid,pg)
    if n<20: return "KHONG_KIEM_DUOC",[]
    need=[c for v in vals for c in codes(v)]
    if not need: return "KHONG_CO_MA",[]
    miss=[c for c in need if sq(c) not in ps]
    return ("KHOP" if not miss else "LECH"),miss
rows=[r for r in mat if isinstance(r.get("Nguồn spec"),list) and len(r["Nguồn spec"])>=3 and isinstance(r["Nguồn spec"][2],int) and r["Nguồn spec"][1] in fi]
print("=== B2. Dương tính, chuẩn hoá lại ===")
for fld in ("Mã VL","NCC/Mã SP","Mô tả","(cả 3)"):
    c=collections.Counter(); bad=[]
    for r in rows:
        _,fid,pg=r["Nguồn spec"][:3]
        vals=[r["Mã VL"],r.get("NCC/Mã SP"),r.get("Mô tả")] if fld=="(cả 3)" else [r.get(fld)]
        s,m=chk(vals,fid,pg); c[s]+=1
        if s=="LECH": bad.append((r["Mã VL"],pg,m,(r.get(fld) if fld!="(cả 3)" else "")))
    print(f"{fld:10}",dict(c))
    if fld=="(cả 3)":
        for b in bad: print("   LECH",b)
print("\n=== C2. Đột biến cả mã (đổi 1 chữ số trong 1 mã của NCC/Mã SP hoặc Mã VL) ===")
cau=tot=0; lot=[]
for r in rows:
    _,fid,pg=r["Nguồn spec"][:3]
    cs=[c for c in codes(r["Mã VL"])+codes(r.get("NCC/Mã SP")) ]
    # chỉ mã đang KHỚP mới mutate (để biết lỗi sinh ra do đột biến)
    cs=[c for c in cs if sq(c) in page_sq(fid,pg)[0]]
    for c in cs:
        d=[i for i,ch in enumerate(c) if ch.isdigit()]
        for _ in range(3):
            i=random.choice(d); n=random.choice([x for x in "0123456789" if x!=c[i]])
            m=c[:i]+n+c[i+1:]; tot+=1
            if sq(m) not in page_sq(fid,pg)[0]: cau+=1
            else: lot.append((r["Mã VL"],c,m))
print(f"bắt {cau}/{tot} = {100*cau/tot:.1f}%; lọt {len(lot)}"); 
for x in lot[:10]: print("   lọt",x)
print("\n=== D2. Link sai trang ===")
ok=tot=sg=0; amb=0
for r in rows:
    _,fid,pg=r["Nguồn spec"][:3]; d=open_fid(fid)
    oth=[p for p in range(1,d.page_count+1) if p!=pg]; w=random.choice(oth); tot+=1
    s,_m=chk([r["Mã VL"],r.get("NCC/Mã SP")],fid,w)
    if s=="LECH":
        ok+=1
        best=[p for p in range(1,d.page_count+1) if chk([r["Mã VL"],r.get("NCC/Mã SP")],fid,p)[0]=="KHOP"]
        if best==[pg]: sg+=1
        elif pg in best: amb+=1
print(f"sai trang bị bắt {ok}/{tot}; gợi ý đúng duy nhất {sg}/{ok}; đúng nhưng nhiều ứng viên {amb}")
print("\n=== E2. Nếu CHỈ kiểm Mã VL (vd ST-02) — trang ngẫu nhiên có chứa mã đó? ===")
h=t=0
for r in rows:
    _,fid,pg=r["Nguồn spec"][:3]; d=open_fid(fid); oth=[p for p in range(1,d.page_count+1) if p!=pg]
    w=random.choice(oth); t+=1
    if sq(r["Mã VL"]) in page_sq(fid,w)[0]: h+=1
print(f"{h}/{t} dòng: mã VL cũng có trên 1 trang ngẫu nhiên khác (= kiểm lọt)")
