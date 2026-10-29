import json,glob,os,re,sys,csv,collections,pymupdf
sys.stdout.reconfigure(encoding="utf-8")
R=sys.argv[1]; PDFROOT=sys.argv[2]
CODE=re.compile(r"(ID-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{3}[A-Z]?)")
code=lambda s:(CODE.search(s.upper()) or [None,None])[1]
MON=re.compile(r"^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)'?$",re.I)
def pdf_info(path):
    p=pymupdf.open(path)[0]; W=[(w[4],w[0],w[1]) for w in p.get_text("words")]
    ids=[w for w in W if re.fullmatch(r"ID-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{3}[A-Z]?",w[0]) and w[1]>1990 and w[2]>1400]
    if not ids: return {}
    i=max(ids,key=lambda a:a[2]); line=[w for w in W if abs(w[2]-i[2])<=6 and w[1]>i[1]+5]
    rev=sorted(line,key=lambda a:-a[1])[0][0] if line else None
    box=sorted([w for w in W if w[1]>2000 and i[2]-120<w[2]<i[2]-20],key=lambda a:(round(a[2]/6),a[1])); toks=[w[0] for w in box]; date=""
    for k,t in enumerate(toks):
        if MON.match(t): date=t+(toks[k+1] if k+1<len(toks) and re.match(r"^'?\d{2}$",toks[k+1]) else ""); break
    return {"id":i[0],"rev":rev,"date":date}
norm=lambda s:re.sub(r"[\s']","",str(s or "")).upper()
pdfs={}
for p in glob.glob(PDFROOT+"/**/*.pdf",recursive=True):
    c=code(os.path.basename(p))
    if c: pdfs.setdefault(c,[]).append(p)
out=[];cnt=collections.Counter()
for f in sorted(glob.glob(R+"/json/*.json")):
    j=json.load(open(f,encoding="utf-8")); c=code(os.path.basename(f))
    if not c or c not in pdfs: cnt["khong ghep duoc PDF theo ten"]+=1; out.append([c or os.path.basename(f),j.get("rev"),"","","","KHONG GHEP PDF"]); continue
    pi=pdf_info(pdfs[c][0])
    if not pi.get("rev") or not j.get("rev"): res="KHONG DOC DUOC"
    else:
        same_rev=norm(j["rev"])==norm(pi["rev"]); dd,pd=norm(j.get("date")),norm(pi.get("date"))
        res="KHOP" if same_rev and (not dd or not pd or dd==pd) else "LECH"
    cnt[res]+=1; out.append([c,j.get("rev"),pi.get("rev"),j.get("date"),pi.get("date"),res])
print(dict(cnt))
for r in out:
    if r[-1]=="LECH": print("  LECH",r)
print("khong doc duoc (vi du):",[r[0] for r in out if r[-1]=="KHONG DOC DUOC"][:10])
with open(R+"/version_check2.csv","w",newline="",encoding="utf-8-sig") as f:
    w=csv.writer(f); w.writerow(["ma","dwg_rev","pdf_rev","dwg_date","pdf_date","ket_qua"]); w.writerows(out)
