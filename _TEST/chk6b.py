import sys,os,win32com.client
sys.stdout.reconfigure(encoding="utf-8")
xl=win32com.client.DispatchEx("Excel.Application");xl.Visible=False;xl.DisplayAlerts=False;xl.AutomationSecurity=1;xl.EnableEvents=False
try:
    wb=xl.Workbooks.Open(os.path.abspath("probe4.xlsm"));ws=wb.Worksheets("Chiet tinh");bq=wb.Worksheets("MATRIX FITFUR-LOOSE")
    for h,ma,r in ((6386,"L005",115),(6438,"L028",138),(6491,"L032",142),(6543,"L012(block thừa)",122),(6595,"L041(+L012)",151),(6650,"L042",152)):
        end=next(x for x in range(h+1,h+200) if ws.Cells(x,1).Value=="END")
        gg=next(x for x in range(h,end) if str(ws.Cells(x,1).Value).strip()=="GIÁ GỐC NHÀ MÁY")
        errs=[(x,c) for x in range(h,end+1) for c in range(1,63) if str(ws.Cells(x,c).Text).startswith("#") and set(str(ws.Cells(x,c).Text))!={"#"}]
        refs=[r2 for r2 in range(8,160) if "H"+str(h) in str(bq.Cells(r2,34).Formula).replace("$","")]
        print(f"{ma}: hdr {h} H={ws.Cells(h,8).Value:,.0f} giá gốc {'=H' if abs(ws.Cells(gg,9).Value-ws.Cells(h,8).Value)<1 else '≠H'} lỗi={errs[:3]} | BoQ trỏ tới block: {refs}")
    wb.Close(False)
finally: xl.Quit()
