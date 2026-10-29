import sys,os,win32com.client
sys.stdout.reconfigure(encoding="utf-8")
xl=win32com.client.DispatchEx("Excel.Application")
xl.Visible,xl.DisplayAlerts,xl.AutomationSecurity,xl.EnableEvents=False,False,1,False
try:
    wb=xl.Workbooks.Open(os.path.abspath("probe4.xlsm"));ws=wb.Worksheets("Chiet tinh")
    for s in map(int,sys.argv[1:]):
        print("=========== block",s)
        for r in range(s,s+130):
            a=ws.Cells(r,1).Text.strip()
            if r>s and a=="END": print(r,"END");break
            e=ws.Cells(r,5).Text.strip();vals=[ws.Cells(r,c).Text.strip() for c in (2,3,4)]
            k=ws.Cells(r,11).Text.strip(); i=ws.Cells(r,9).Text.strip()
            if a or any(vals) or e:
                print(r,"+%d"%(r-s),a[:38].ljust(38),vals,"E=",e,"K=",k[:22],"I=",i)
    wb.Close(False)
finally: xl.Quit()
