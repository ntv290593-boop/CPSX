"""Phiếu hướng dẫn bóc cửa NORTON -> block Door (RE00). Offset = số hàng tính từ dòng tiêu đề block."""
import json,collections
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.worksheet.hyperlink import Hyperlink
D='/tmp/claude-0/norton/'
IN='..\\00_input\\261023_D01\\'
SH=json.load(open(D+'sheetmap.json'))
GEN='SB_TENDER_PLOT B_GENERAL, BASEMENT_2026 10 08'
def pg(p): return IN+'03_BanVe\\'+SH[GEN][str(p)]
BV54,BV55=pg(54),pg(55)
# --- loại cửa: tổng SL BoQ, dòng BoQ, kích thước cửa (overall) theo bảng thống kê
boq=[b for b in json.load(open(D+'boq.json')) if b['kind']=='DOOR' and b['type']]
T=collections.OrderedDict()
for b in boq:
  t=T.setdefault(b['type'],{'W':b['W'],'H':b['H'],'SL':0,'rows':[]})
  t['SL']+=b['qty'] if isinstance(b['qty'],(int,float)) else 0;t['rows'].append(f"{b['sheet']}!{b['row']}")
D1='Door#1 Cửa phẳng 1 cánh - new';D3=D1  # cửa thường: cùng block, K='Cửa thường'
COMMON3=lambda t: [  # Door#3 cửa thường
 (0,'B','Header: rộng cửa','900',T[t]['W'],'Bảng thống kê CD.66.00 / BoQ overall size',BV54),
 (0,'C','Header: sâu khung bao','110',120,'Mặt bằng TL1:20 CD.00.68A: tường 120',BV55),
 (0,'D','Header: cao cửa','2250',T[t]['H'],'Bảng thống kê / BoQ',BV54),
 (0,'E','Header: SL','0',T[t]['SL'],'Σ KL các dòng BoQ cùng mã (R0048)',''),
 (0,'K','Header: loại cửa','EI30','Cửa thường','BV: fire rating không yêu cầu',BV54),
 (0,'L','Header: mã hoàn thiện','LA-01','LAM-01','LAM-01 = 5966OG Acacia Naturale (ML Mass Units 001); BV ghi LK 4535A → RFI-02',BV55)]
SPEC={
 'DH114':(D1,[
  (0,'B','Header: rộng cửa','900',T['DH114']['W'],'Bảng thống kê / BoQ overall size',BV54),
  (0,'C','Header: sâu khung bao','110',210,'Mặt bằng TL1:20 CD.00.68A: tường 210',BV55),
  (0,'D','Header: cao cửa','2250',T['DH114']['H'],'Bảng thống kê / BoQ',BV54),
  (0,'E','Header: SL','0',T['DH114']['SL'],'Σ KL các dòng BoQ cùng mã (R0048)',''),
  (0,'K','Header: loại cửa','EI30','EI30','BV + BoQ: EI30',BV54),
  (0,'L','Header: mã hoàn thiện','LA-01','LAM-01','LAM-01 = 5966OG Acacia Naturale (ML Mass Units 001); BV ghi LK 4535A → RFI-02',BV55),
  (0,'M','Header: dày khung bao','=35',38,'BV CD.00.68 ghi chú DH114: khung bao dày 38mm (chi tiết điển hình 41±1 → ghi chú)',BV54),
  (2,'B','Cánh cửa: dày','=50',46,'BV CD.00.68 DH114: cánh 46mm (chi tiết điển hình 48±1)',BV54),
  (4,'K','Lớp FR cánh','5ERON','6ERON','2 tấm chống cháy 6mm',BV54),
  (5,'K','Lớp FR cánh (2)','5ERON','6ERON','như trên',BV54),
  (11,'K','Lõi','RW','RW','Rockwool 40kg/m³ → dùng RW 80kg/m³ (Material không có 40) – giả định',BV54),
  (15,'K','Hoàn thiện khung bao','3MDF','6MDF','Khung bao: lõi 25 + 2×MDF chống ẩm 6mm',BV54),
  (19,'K','Chỉ cửa (nẹp khung bao)','9MDF','12MDF','Nẹp khung bao MDF chống ẩm 12mm',BV54),
  (23,'K','Rebate (chặn cửa)','15MDF','17MDF','Chặn cửa 17mm (10 gỗ ghép + 6 MDF) → tính 17MDF',BV54),
  (24,'K','FR chặn cửa','5ERON','','BV/BoQ: chặn cửa không có tấm chống cháy → bỏ (E tự =0)',BV54),
  (14,'K','FR khung bao','5ERON','6ERON','BV CD.00.68A: khung bao gỗ ghép + tấm chống cháy + MDF',BV55)]),
 'ID-DH112':(D3,COMMON3('ID-DH112')+[
  (0,'J','Header: ván nền','MDF','MDF','MDF chống ẩm',BV54),
  (0,'M','Header: dày khung bao','=30',29,'Khung bao dày 29±1',BV54),
  (2,'B','Cánh cửa: dày','=40 (cửa thường)',37,'Cánh 37±1 (tổ ong 25 + MDF 6×2)',BV54),
  (19,'K','Chỉ cửa','12MDF','12MDF','Nẹp L MDF chống ẩm 12mm – khớp mẫu',BV54),
  (23,'K','Rebate (chặn cửa)','15MDF','25MDF','Chặn cửa MDF chống ẩm 25mm',BV54),
  (25,'E','Rebate đứng lớp 6mdf','=E23','0','Chặn cửa liền 25mm, không phủ thêm 6mm',BV54),
  (28,'E','Rebate ngang lớp 6mdf','=E26','0','như trên',BV54)]),
}
for t in ('ID-DH111','DH117','DH118'):
  SPEC[t]=(D3,COMMON3(t)+[
  (0,'J','Header: ván nền','MDF','WPB','Mặt cánh/khung WPB chống nước',BV54),
  (0,'M','Header: dày khung bao','=30',31,'Khung bao dày 31±1 (gỗ ghép 17 + WPB 6)',BV54),
  (2,'B','Cánh cửa: dày','=40 (cửa thường)',39,'Cánh 39±1 (tổ ong 25 + WPB 6×2); lõi tổ ong theo công thức = 27mm',BV54),
  (19,'K','Chỉ cửa','12MDF','15WPB','Nẹp L WPB 14mm → 15WPB (14WPB giá 0, R0074)',BV54),
  (23,'K','Rebate (chặn cửa)','15MDF','17WPB','Chặn cửa WPB 17mm',BV54),
  (25,'E','Rebate đứng lớp 6mdf','=E23','0','Chặn cửa liền 17mm',BV54),
  (28,'E','Rebate ngang lớp 6mdf','=E26','0','như trên',BV54)])
OV=[('OverView','C21','Mã test cửa CC','EI601','EI301','DH114 EI30 – tính chi phí test + ra giấy (R0144)'),
    ('OverView','C43','Địa điểm giao','Quận 1+ Bình Thạnh','Huyện Thuận An - Bình Dương','Công trường P. Thuận Giao (title block)'),
    ('OverView','E43','Loại xe','Xe 2,5T','Xe 15T','~3.040 bộ cửa – giả định, chỉnh nếu cần'),
    ('OverView','C47:C49','Ron CC / ron cách âm / foam','Không','Có','DH114: smoke seal + fire rated foam trong mô tả cửa BoQ'),
    ('OverView','L33:Q63','Mã HT mới','—','LAM-01 Laminate An Cường 5966OG Acacia Naturale (mã thiết kế ML Mass Units 001)','Giá tạm = LA-01 (R0059/R0143)'),
    ('Material','cụm LA-01','Chép cụm LA-01 → LA-02','—','LAM-01','R0059')]
H=Font(bold=True,color='FFFFFF');HF=PatternFill('solid',fgColor='1F4E78');LK=Font(color='0563C1',underline='single')
def sheet(wb,name,cols,w):
  ws=wb.create_sheet(name);ws.append(cols)
  for c in ws[1]: c.font=H;c.fill=HF;c.alignment=Alignment(wrap_text=True)
  for i,x in enumerate(w): ws.column_dimensions[chr(65+i)].width=x
  ws.freeze_panes='A2';return ws
def build_phieu(out):
  wb=Workbook();wb.remove(wb.active)
  ws=sheet(wb,'TONG',['Mã cửa','Block mẫu','Rộng','Cao','SL (Σ BoQ)','Dòng BoQ trỏ AH về block','Loại'],[11,32,7,7,10,70,12])
  for t,v in T.items(): ws.append([t,SPEC[t][0],v['W'],v['H'],v['SL'],', '.join(v['rows']),'EI30' if t=='DH114' else 'Cửa thường'])
  ws=sheet(wb,'PHIEU_BOC',['Mã cửa','Block mẫu','Offset hàng','Cột','Mục','Giá trị mẫu','Giá trị NORTON','Lý do / nguồn','Bản vẽ'],[10,30,8,5,24,12,16,70,12])
  for t,(blk,rows) in SPEC.items():
    for off,col,lab,old,new,why,bv in sorted(rows,key=lambda x:(x[0],x[1])):
      ws.append([t,blk,off,col,lab,old,new,why,'Mở' if bv else ''])
      if bv: c=ws.cell(ws.max_row,9);c.hyperlink=bv;c.font=LK
  ws=sheet(wb,'OVERVIEW',['Sheet','Ô','Mục','Mẫu','NORTON','Ghi chú'],[10,10,26,20,40,60])
  for x in OV: ws.append(list(x))
  ws=sheet(wb,'AI_INPUT',['Ma','DongBoQ','Ten','Loai','W','D','H','SL','VatLieu','HoanThien','ThongSo','Nguon','DoTinCay','Trang','GhiChu'],[10,40,30,12,6,6,6,7,10,10,50,40,9,8,40])
  for t,v in T.items():
    blk,rows=SPEC[t];d={(o,c):n for o,c,_,_,n,_,_ in rows}
    ws.append([t,', '.join(v['rows']),blk,'EI30' if t=='DH114' else 'Cửa thường',v['W'],d.get((0,'C')),v['H'],v['SL'],d.get((0,'J'),'MDF'),'LAM-01',
      f"Cánh {d.get((2,'B'))}mm; khung bao {d.get((0,'M'))}mm; chỉ {d.get((19,'K'))}; chặn {d.get((23,'K'))}",'GEN p54 CD.00.68 / p55 CD.00.68A; BoQ',0.8,'54-55','Chỉ cửa, không phụ kiện 9.xx'])
  for s in wb.worksheets:
    for r in s.iter_rows(min_row=2):
      for c in r: c.alignment=Alignment(wrap_text=True,vertical='top')
  wb.save(out)
if __name__=='__main__':
  build_phieu('/mnt/project-files/NORTON/02_takeoff/PHIEU_BOC_NORTON.xlsx');print({t:v['SL'] for t,v in T.items()})
