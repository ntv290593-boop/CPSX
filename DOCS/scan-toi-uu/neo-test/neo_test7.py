src=open(r"D:\COSTING EXCEL\DOCS\scan-toi-uu\neo-test\neo_test5.py",encoding='utf-8').read().split("stat=collections")[0]
exec(src)
random.seed(5)
CH=re.compile(r"\d+(?:\.\d+)?\s*(?:mm)?\s*(?:[x×X]|-|–)\s*\d+(?:\.\d+)?(?:\s*mm)?(?:\s*(?:[x×X]|-|–)\s*\d+(?:\.\d+)?(?:\s*mm)?)*")
NUM=re.compile(r"(?<![\d.])\d+(?:\.\d+)?(?![\d])")
def stream(t): return NUM.findall(re.sub(r"(\d),(\d{3})",r"\1\2",t))
def set_ok(parts,nums): return all(p in set(nums) for p in parts)
def ord_ok(parts,nums,slack=3):
    # các số xuất hiện đúng thứ tự, mỗi số cách số trước tối đa `slack` token số
    n=len(nums)
    for i,v in enumerate(nums):
        if v!=parts[0]: continue
        j=i; ok=True
        for p in parts[1:]:
            k=next((q for q in range(j+1,min(n,j+1+slack)) if nums[q]==p),None)
            if k is None: ok=False;break
            j=k
        if ok: return True
    return False
pages=[]
for r,sz in pick:
    zp=os.path.join(INP,r["Đường dẫn"])
    try: d=pymupdf.open(stream=zs[zp].read(r["Trong zip"]),filetype="pdf")
    except Exception: continue
    for i in range(d.page_count):
        t=d[i].get_text("words"); t=" ".join(w[4] for w in sorted(t,key=lambda w:(w[5],w[6],w[7])))  # thứ tự khối/dòng/từ gốc
        if len(CH.findall(t))>=1 and len(stream(t))>=20: pages.append((r["ID"],i+1,t))
        if len(pages)>400: break
    if len(pages)>400: break
print("trang có chuỗi kích thước & >=20 số:",len(pages))
real=[];tot=0
for fid,pg,t in pages:
    nums=stream(t)
    for m in CH.finditer(t):
        parts=NUM.findall(m.group(0))
        if len(parts)<2: continue
        tot+=1; real.append((fid,pg,parts,nums))
print("chuỗi thật:",tot)
# dương tính: chuỗi thật có tự khớp không
sp=sum(set_ok(p,n) for _,_,p,n in real); op=sum(ord_ok(p,n) for _,_,p,n in real)
print(f"dương tính: tập {sp}/{tot}  thứ tự {op}/{tot}")
# đột biến khó: thay 1 số bằng số KHÁC đang có trên trang
c_set=c_ord=n=0; ex=[]
for fid,pg,parts,nums in real:
    pool=[x for x in set(nums) if x not in parts]
    if not pool: continue
    for _ in range(3):
        k=random.randrange(len(parts)); m=list(parts); m[k]=random.choice(pool); n+=1
        if not set_ok(m,nums): c_set+=1
        if not ord_ok(m,nums): c_ord+=1
        elif len(ex)<5: ex.append((fid,pg,parts,m))
print(f"đột biến khó (số thay = số khác có trên trang): kiểm TẬP bắt {c_set}/{n} ({100*c_set/n:.0f}%) | kiểm THỨ TỰ bắt {c_ord}/{n} ({100*c_ord/n:.0f}%)")
for e in ex: print("  lọt thứ tự:",e)
# hoán đổi chiều: RxSxC -> đổi chỗ 2 số (lỗi gán nhầm chiều như Gemini)
sw_set=sw_ord=m2=0
for fid,pg,parts,nums in real:
    if len(parts)<3 or len(set(parts))<3: continue
    m=list(parts); m[0],m[1]=m[1],m[0]; m2+=1
    if not set_ok(m,nums): sw_set+=1
    if not ord_ok(m,nums): sw_ord+=1
print(f"đổi chỗ 2 chiều (R↔S): tập bắt {sw_set}/{m2}; thứ tự bắt {sw_ord}/{m2}")
