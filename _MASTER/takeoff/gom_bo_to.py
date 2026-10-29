"""Gom tờ bản vẽ thành BỘ từ khung tên ("... SHEET n OF m") và đánh dấu tờ chưa xác nhận đọc.

Chạy: py -I gom_bo_to.py <PROJECT_DATA.xlsx>
Chỉ THÊM cột/sheet (PAGES: Bộ tờ, Tờ cùng bộ; BOQ_TONGHOP: Bộ tờ phải đọc, Tờ chưa xác nhận đọc;
sheet CALLOUT; PROJECT: Phạm vi). Không đổi/xóa phần có sẵn. Có sao lưu + đếm hyperlink trước/sau.
"""
import re, shutil, sys, datetime
from copy import copy
import openpyxl

SHEET_OF = re.compile(r'^(.*?)\s*SHEET\s*(\d+)\s*OF\s*(\d+)', re.I)
CODE = re.compile(r'ID-[A-Z0-9]+(?:-[A-Z0-9]+)*-\d{3}')


def norm(code):
    code = code.upper().strip()
    return code[3:] if code.startswith('ID-') else code


def codes_in(text):
    """Mọi mã tờ trong text, kể cả dạng rút gọn PS-TY-711."""
    t = str(text or '').upper()
    out = {norm(c) for c in CODE.findall(t)}
    out |= {m for m in re.findall(r'\b[A-Z0-9]{2,4}-[A-Z]{2}-\d{3}\b', t)}
    return out


def add_header(ws, col, text):
    for c in ws[1]:
        if c.value == text:
            return c.column
    ref = ws.cell(1, col - 1)
    c = ws.cell(1, col, text)
    c.font, c.fill, c.alignment, c.border = copy(ref.font), copy(ref.fill), copy(ref.alignment), copy(ref.border)
    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 28
    return col


def count_links(wb):
    return sum(1 for ws in wb for row in ws.iter_rows() for c in row if c.hyperlink)


def main(path):
    sys.stdout.reconfigure(encoding='utf-8')
    stamp = datetime.date.today().isoformat()
    bak = path.replace('\\', '/').rsplit('/', 1)
    shutil.copy2(path, f"D:/COSTING EXCEL/_MASTER/backup/{bak[1][:-5]}_{stamp}.xlsx")
    wb = openpyxl.load_workbook(path)
    links0 = count_links(wb)

    # --- PAGES: gom bộ
    ws = wb['PAGES']
    hdr = [c.value for c in ws[1]]
    iF, iNo, iT = hdr.index('File'), hdr.index('Số BV'), hdr.index('Tên BV')
    series = {}   # (file, prefix) -> {n: so_bv}
    row_key = {}
    for r in range(2, ws.max_row + 1):
        m = SHEET_OF.match(str(ws.cell(r, iT + 1).value or ''))
        if not m:
            continue
        key = (ws.cell(r, iF + 1).value, m.group(1).strip())
        series.setdefault(key, {})[int(m.group(2))] = (str(ws.cell(r, iNo + 1).value), int(m.group(3)))
        row_key[r] = (key, int(m.group(2)), int(m.group(3)))
    if not series:
        print('KHÔNG GOM ĐƯỢC: tên tờ không có "SHEET n OF m" (khung tên khác). Cần quy tắc riêng cho dự án này.')
        return
    cP, cM = ws.max_column + 1, ws.max_column + 2
    cP = add_header(ws, cP, 'Bộ tờ'); cM = add_header(ws, cM, 'Tờ cùng bộ')
    sheet_to_set = {}
    for r, (key, n, m) in row_key.items():
        members = series[key]
        so = sorted(v[0] for v in members.values())
        flag = '' if len(members) >= m else f' ⚠ thiếu {m - len(members)}/{m} tờ trong PAGES'
        ws.cell(r, cP, f'{key[1]} | {n}/{m}{flag}')
        ws.cell(r, cM, '; '.join(so))
        for s in so:
            sheet_to_set[norm(s)] = so
    print('bộ tờ:', len(series), '| tờ có bộ:', len(row_key))

    # --- BOQ_TONGHOP: tờ phải đọc vs đã xác nhận
    wt = wb['BOQ_TONGHOP']
    th = [c.value for c in wt[1]]
    iCite, iRead, iChk = th.index('Bản vẽ chi tiết'), th.index('Chi tiết (số – tên)'), th.index('Kiểm tờ BV')
    c1, c2 = wt.max_column + 1, wt.max_column + 2
    c1 = add_header(wt, c1, 'Bộ tờ phải đọc'); c2 = add_header(wt, c2, 'Tờ chưa xác nhận đọc')
    nmiss = 0
    for r in range(2, wt.max_row + 1):
        cited = codes_in(wt.cell(r, iCite + 1).value)
        need = []
        for c in cited:
            for s in sheet_to_set.get(c, [c]):
                if norm(s) not in [norm(x) for x in need]:
                    need.append(s)
        if not need:
            continue
        ev = ' '.join(re.findall(r'(?i)đã xem hình[^|;]*', str(wt.cell(r, iRead + 1).value or '') + ' ' + str(wt.cell(r, iChk + 1).value or '')))
        seen = codes_in(ev)
        miss = [s for s in need if norm(s) not in seen]
        wt.cell(r, c1, '; '.join(need))
        wt.cell(r, c2, '; '.join(miss) if miss else 'ĐỦ')
        nmiss += bool(miss)
    print('hạng mục còn tờ chưa xác nhận đọc:', nmiss)

    # --- CALLOUT
    if 'CALLOUT' not in wb.sheetnames:
        wc = wb.create_sheet('CALLOUT', index=wb.sheetnames.index('PAGES') + 1)
        wc.append(['Tờ nguồn', 'Chi tiết nguồn', 'Tờ đích', 'Chi tiết đích', 'Cách đọc', 'Ghi chú'])
        for c in wc[1]:
            c.font, c.fill = copy(wt.cell(1, 1).font), copy(wt.cell(1, 1).fill)
        for row in [
            ('ID-PS-TY-711', '07', 'ID-PS-TY-712', '13', 'ảnh', 'dấu mặt cắt trên mặt đứng 07'),
            ('ID-PS-TY-711', '10', 'ID-PS-TY-712', '12', 'ảnh', 'dấu mặt cắt trên mặt đứng 10'),
            ('ID-PS-TY-711', '10', 'ID-PS-TY-712', '14', 'ảnh', 'dấu mặt cắt trên mặt đứng 10'),
            ('ID-PS-TY-711', '07; 10', 'ID-PS-TY-712', '18', 'ảnh', 'bong bóng chi tiết cạnh kích thước 1460'),
            ('ID-PS-TY-712', '12', 'ID-PS-TY-712', '15', 'ảnh', 'chi tiết kệ WD-04'),
            ('ID-PS-TY-712', '13', 'ID-PS-TY-712', '16', 'ảnh', 'chi tiết đế ngăn kéo'),
            ('ID-PS-TY-712', '14', 'ID-PS-TY-712', '17', 'ảnh', 'chi tiết gioăng cao su (kim loại)'),
        ]:
            wc.append(row)
        for col, w in zip('ABCDEF', (16, 16, 16, 14, 10, 44)):
            wc.column_dimensions[col].width = w

    # --- PROJECT: phạm vi
    wp = wb['PROJECT']
    if not any(str(wp.cell(r, 1).value or '').startswith('Phạm vi tính giá') for r in range(1, wp.max_row + 1)):
        r = next(r for r in range(2, wp.max_row + 2) if wp.cell(r, 1).value is None)
        wp.cell(r, 1, 'Phạm vi tính giá (nội thất)')
        wp.cell(r, 2, 'Chỉ tính gỗ ván (kể cả ván nền sau lớp hoàn thiện). KHÔNG tính: kim loại (MT-*), đá (ST-*), '
                      'giấy dán tường WC-* (phần giấy), OS&E (giá ủi, két), đèn.')
        wp.cell(r, 3, 'User 2026-10-24')

    wb.save(path)
    wb2 = openpyxl.load_workbook(path)
    print('hyperlink trước/sau:', links0, count_links(wb2))


if __name__ == '__main__':
    main(sys.argv[1])
