import sys,os,win32com.client
sys.stdout.reconfigure(encoding="utf-8")
xl=win32com.client.DispatchEx("Excel.Application");xl.Visible=False;xl.DisplayAlerts=False;xl.AutomationSecurity=1;xl.EnableEvents=False
try:
    wb=xl.Workbooks.Open(os.path.abspath("probe4.xlsm"));ws=wb.Worksheets("Chiet tinh")
    T=lambda r,c:ws.Cells(r,c).Text.strip()
    for h in (6386,6438,6491,6543,6595,6650):
        end=next(x for x in range(h+1,h+200) if ws.Cells(x,1).Value=="END")
        print(f"\n===== hdr {h}-{end} ({end-h+1} hàng) | {T(h,2)}x{T(h,3)}x{T(h,4)} E={T(h,5)} H={T(h,8)} | E formula={str(ws.Cells(h,5).Formula)[-22:]}")
        for x in range(h+1,end):
            a=T(x,1);e=T(x,5);k=T(x,11)
            if a.startswith(("TỔNG","CHI PHÍ","GIA CÔNG","Vật tư phụ","Vật tư đóng","VẬT TƯ","Vật tư Sofa","VẢI")) : continue
            if e not in ("","-","0") and a not in ("-",) :
                print(f" {x} +{x-h:<3}{a[:34]:34} B={T(x,2):>6} C={T(x,3):>6} D={T(x,4):>6} E={e:>5} F={T(x,6):>5} G={T(x,7):>7} I={T(x,9):>10} K={k[:8]:8} M={T(x,13):>5} N={T(x,14)}")
    wb.Close(False)
finally: xl.Quit()
