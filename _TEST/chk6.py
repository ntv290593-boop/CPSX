import sys,os,re,difflib,win32com.client
sys.stdout.reconfigure(encoding="utf-8")
BQ={"L005":115,"L012":122,"L028":138,"L032":142,"L041":151,"L042":152}
def col(j): return chr(65+j) if j<26 else "A"+chr(65+j-26)
xl=win32com.client.DispatchEx("Excel.Application");xl.Visible=False;xl.DisplayAlerts=False;xl.AutomationSecurity=1;xl.EnableEvents=False
try:
    wb=xl.Workbooks.Open(os.path.abspath("probe4.xlsm"));ws=wb.Worksheets("Chiet tinh");bq=wb.Worksheets("MATRIX FITFUR-LOOSE");lg=wb.Worksheets("_EDITLOG")
    # 1) ô Claude ghi bị user đổi (so nhật ký đợt KL)
    last=lg.UsedRange.Row+lg.UsedRange.Rows.Count-1;ch=[]
    for i in range(2,last+1):
        note=str(lg.Cells(i,5).Value or "")
        if "REV04 L0" in note and any(m in note for m in BQ) or "loose KL" in note:
            sh,addr,val=lg.Cells(i,1).Value,lg.Cells(i,2).Value,str(lg.Cells(i,6).Value)[2:]
            c=(ws if sh=="Chiet tinh" else bq).Range(addr); cur=str(c.Value) if addr.startswith("AJ") else str(c.Formula)
            if cur!=val and cur[:150]!=val[:150]: ch.append((sh,addr,val[:50],cur[:50]))
    print("Claude-ghi bị user đổi:",len(ch))
    for x in ch[:40]: print("  ",x)
    # 2) block-level aligned diff vs nguồn 5779
    NC=14
    def grab(s):
        end=next(x for x in range(s+1,s+300) if ws.Cells(x,1).Value=="END");rng=ws.Range(ws.Cells(s,1),ws.Cells(end,NC));return end,rng.Value,rng.FormulaR1C1
    e1,v1,f1=grab(5779)
    for ma,r in BQ.items():
        h=int(re.search(r"H(\d+)",str(bq.Cells(r,34).Formula)).group(1))
        e0,v0,f0=grab(h);lab=lambda v:[str(x[0] or "").strip()[:30] for x in v]
        a=str(ws.Cells(h,1).Formula);e=str(ws.Cells(h,5).Formula);hv=ws.Cells(h,8).Value
        gg=next(x for x in range(h,e0) if str(ws.Cells(x,1).Value).strip()=="GIÁ GỐC NHÀ MÁY")
        errs=[(x,c) for x in range(h,e0+1) for c in range(1,63) if str(ws.Cells(x,c).Text).startswith("#") and set(str(ws.Cells(x,c).Text))!={"#"}]
        print(f"\n=== {ma} block {h}-{e0}: H={hv:,.0f} | A {'OK' if a.endswith('!A'+str(r)) else 'SAI'} E {'OK' if e.endswith('!AG'+str(r)) else 'SAI'} | giá gốc {'=H' if abs(ws.Cells(gg,9).Value-hv)<1 else '≠H'} | lỗi {errs[:3]} | AJ {'có' if bq.Cells(r,36).Value else 'TRỐNG'}")
        sm=difflib.SequenceMatcher(None,lab(v1),lab(v0),autojunk=False)
        for tag,i1,i2,j1,j2 in sm.get_opcodes():
            if tag=="equal":
                for k in range(i2-i1):
                    for c in range(1,NC):
                        x,y=f1[i1+k][c],f0[j1+k][c]
                        if x!=y and not(str(x).startswith('=') and str(y).startswith('=') and re.sub(r'\d+','#',str(x))==re.sub(r'\d+','#',str(y))):
                            print(f"  sửa {col(c)}{h+j1+k} (+{j1+k}) [{lab(v0)[j1+k]}]: {str(x)[:30]!r} → {str(y)[:30]!r}")
            elif tag=="insert":
                for k in range(j1,j2): print(f"  THÊM +{k}: {[str(z)[:18] for z in v0[k][:6]]}")
            elif tag=="delete":
                for k in range(i1,i2): print(f"  BỎ nguồn +{k}: {lab(v1)[k]}")
    wb.Close(False)
finally: xl.Quit()
