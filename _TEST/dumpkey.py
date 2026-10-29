import sys,os,win32com.client
sys.stdout.reconfigure(encoding="utf-8")
xl=win32com.client.DispatchEx("Excel.Application");xl.Visible=False;xl.DisplayAlerts=False;xl.AutomationSecurity=1;xl.EnableEvents=False
try:
    wb=xl.Workbooks.Open(os.path.abspath("probe4.xlsm"));ws=wb.Worksheets("Chiet tinh");T=lambda r,c:ws.Cells(r,c).Text.strip()
    for h in map(int,sys.argv[1:]):
        end=next(x for x in range(h+1,h+250) if ws.Cells(x,1).Value=="END")
        print(f"\n===== hdr {h}-{end} | {T(h,1)} {T(h,2)}x{T(h,3)}x{T(h,4)} E={T(h,5)} H={T(h,8)}")
        for x in range(h+1,end):
            a=T(x,1);e=T(x,5)
            if a.startswith(("TỔNG","GIÁ GỐC")): print(f" {x} +{x-h:<3}{a[:34]:34} I={T(x,9)}");continue
            if (e not in ("","-","0") and a) or (a.isupper() and len(a)>8 and not a.startswith(("CHI PHÍ","GIA CÔNG","VẬT TƯ"))):
                print(f" {x} +{x-h:<3}{a[:34]:34} B={T(x,2):>6} C={T(x,3):>6} D={T(x,4):>6} E={e:>5} I={T(x,9):>10} K={T(x,11)[:10]}")
    wb.Close(False)
finally: xl.Quit()
