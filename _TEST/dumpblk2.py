import sys,os,win32com.client
sys.stdout.reconfigure(encoding="utf-8")
xl=win32com.client.DispatchEx("Excel.Application")
xl.Visible,xl.DisplayAlerts,xl.AutomationSecurity,xl.EnableEvents=False,False,1,False
try:
    wb=xl.Workbooks.Open(os.path.abspath("probe4.xlsm"));ws=wb.Worksheets("Chiet tinh");bq=wb.Worksheets("MATRIX FITFUR-LOOSE")
    for r in map(int,sys.argv[1].split(",")): print("AJ",bq.Cells(r,1).Value,":",str(bq.Cells(r,36).Value).replace("\n"," / "))
    for s in map(int,sys.argv[2:]):
        print("=========== block",s,"hdr",[ws.Cells(s,c).Text.strip() for c in (1,2,3,4,5,8,10,11,12)])
        for r in range(s+1,s+130):
            a=ws.Cells(r,1).Text.strip()
            if a=="END": break
            e=ws.Cells(r,5).Text.strip()
            if a.startswith(("TỔNG CHI PHÍ VẬT TƯ CHÍNH","GIÁ GỐC")): print(r,"+%d"%(r-s),a,ws.Cells(r,9).Text);continue
            if a and e not in ("","-","0") or (a.isupper() and len(a)>6 and not a.startswith(("TỔNG","CHI PHÍ","GIA CÔNG","VẬT TƯ"))):
                print(r,"+%d"%(r-s),a[:36].ljust(36),[ws.Cells(r,c).Text.strip() for c in (2,3,4)],"E=",e,"K=",ws.Cells(r,11).Text.strip()[:18],"I=",ws.Cells(r,9).Text.strip())
    wb.Close(False)
finally: xl.Quit()
