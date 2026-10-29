import sys,os,re,difflib,win32com.client
sys.stdout.reconfigure(encoding="utf-8")
SRC={"L026":4536,"L027":4536,"L035":4586,"L037":2435,"L038":2435,"L040":2435,"L034":2435,"L039":1018}
BQROW={"L026":136,"L027":137,"L035":145,"L037":147,"L038":148,"L040":150,"L034":144,"L039":149}
def col(j): return chr(65+j) if j<26 else "A"+chr(65+j-26)
xl=win32com.client.DispatchEx("Excel.Application");xl.Visible=False;xl.DisplayAlerts=False;xl.AutomationSecurity=1;xl.EnableEvents=False
try:
    wb=xl.Workbooks.Open(os.path.abspath("probe4.xlsm"));ws=wb.Worksheets("Chiet tinh");bq=wb.Worksheets("MATRIX FITFUR-LOOSE");tp=wb.Worksheets("Template")
    NC=14
    def grab(sh,s):
        end=next(x for x in range(s+1,s+300) if sh.Cells(x,1).Value=="END")
        rng=sh.Range(sh.Cells(s,1),sh.Cells(end,NC))
        return end,rng.Value,rng.FormulaR1C1
    for ma,src in [(k,v) for k,v in SRC.items() if k in sys.argv[1:]]:
        ah=str(bq.Cells(BQROW[ma],34).Formula); new=int(re.search(r"H(\d+)",ah).group(1))
        sws=tp if src==1018 else ws
        e1,v1,f1=grab(sws,src); e0,v0,f0=grab(ws,new)
        lab=lambda v:[str(r[0] or "").strip()[:30] for r in v]
        sm=difflib.SequenceMatcher(None,lab(v1),lab(v0),autojunk=False)
        print(f"\n===== {ma}: block {new}-{e0} ({e0-new+1} hàng) | nguồn {src} ({e1-src+1} hàng) | H={ws.Cells(new,8).Text.strip()} E={ws.Cells(new,5).Text.strip()}")
        for tag,i1,i2,j1,j2 in sm.get_opcodes():
            if tag=="equal":
                for k in range(i2-i1):
                    for c in range(1,NC):
                        a,b=f1[i1+k][c],f0[j1+k][c]
                        if a!=b and not (str(a).startswith('=') and str(b).startswith('=') and re.sub(r'\d+','#',str(a))==re.sub(r'\d+','#',str(b))):
                            r=new+j1+k; 
                            print(f"  sửa  {col(c)}{r} (+{j1+k}) [{lab(v0)[j1+k]}]: nguồn={str(a)[:34]!r} → nay={str(b)[:34]!r}")
            elif tag=="insert":
                for k in range(j1,j2): print(f"  THÊM hàng +{k} (r{new+k}): {[str(x)[:20] for x in v0[k][:5]]}")
            elif tag=="delete":
                for k in range(i1,i2): print(f"  BỎ hàng nguồn +{k}: {lab(v1)[k]}")
            else:
                print(f"  ĐỔI nhãn nguồn {lab(v1)[i1:i2]} → {lab(v0)[j1:j2]}")
    wb.Close(False)
finally: xl.Quit()
