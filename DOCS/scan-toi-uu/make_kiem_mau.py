# -*- coding: utf-8 -*-
"""Tạo KIEM_MAU_DAP_AN_Blanca.xlsx: 10 mã, trang đáp án + top-5 của máy, để user đánh ĐÚNG/SAI.
Link theo chuẩn PROJECT_DATA: ô tự trỏ về chính nó, đích ở sheet _LINKS -> add-in SRC_LINKS mở Edge #page=N."""
import random, openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from proto_rank import *

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "KIEM_MAU_DAP_AN_Blanca_v2.xlsx")
LK = Font(color="0563C1", underline="single")
al = lambda p: (ALIAS.get(p[0], p[0]), p[1])
g2 = {m: {al(p) for p in s} for m, s in g.items()}


def top(m): return [(x[1], x[2]) for x in rank(m, boq[m], desc=False)][:5]


# 10 mã: chỉ lấy mã mà đáp án và máy chọn khác nhau, mỗi nhóm tiền tố một mã
random.seed(7)
lech = [m for m in g if set(top(m)) != g2[m]]
groups = {}
for m in lech: groups.setdefault(re.match(r"[A-Za-z]+", m).group(0).upper(), []).append(m)
pick = [random.choice(groups[k]) for k in sorted(groups, key=lambda k: -len(groups[k]))][:10]

wb = openpyxl.Workbook(); ws = wb.active; ws.title = "KIEM_MAU"
ws.append(["Mã", "Hạng mục", "Khu", "File", "Trang", "Số BV · Tên tờ", "Đáp án (subagent)", "Máy xếp top-5", "Mở", "Bạn đánh giá", "Ghi chú"])
for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
dv = DataValidation(type="list", formula1='"ĐÚNG,SAI,KHÔNG CHẮC"', allow_blank=True); ws.add_data_validation(dv)
vang = PatternFill("solid", fgColor="FFF2CC")
wl = wb.create_sheet("_LINKS")
wl.append(["Key", "Sheet", "Ô", "Nhãn", "File (tương đối)", "Trang", "Ô đích", "Mở (dự phòng)", "Độ tin cậy"])
for m in pick:
    t = top(m)
    for p in sorted(g2[m] | set(t)):
        ws.append([m, boq[m].get("Hạng mục"), ", ".join(sorted(khu.get(m, []))), p[0], p[1], ptitle.get(p, "")[:90],
                   "CÓ" if p in g2[m] else "", f"#{t.index(p) + 1}" if p in t else "", f"{p[0]} p{p[1]}", None, None])
        i = ws.max_row; c = ws.cell(i, 9)
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'KIEM_MAU'!{c.coordinate}", display=c.value); c.font = LK
        rel = os.path.relpath(os.path.join(P["inp"], files[p[0]]["Đường dẫn"]), HERE)
        wl.append([f"KIEM_MAU!{c.coordinate}", "KIEM_MAU", c.coordinate, f"{p[0]} p{p[1]}", rel, p[1], None, "Mở", 1.0])
        wl.cell(wl.max_row, 8).hyperlink = rel; wl.cell(wl.max_row, 8).font = LK
        dv.add(ws.cell(i, 10))
        if (p in g2[m]) != (p in t):
            for x in ws[i]: x.fill = vang
for col, w in zip("ABCDEFGHIJK", [11, 34, 10, 7, 6, 55, 10, 10, 10, 13, 30]): ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
hd = wb.create_sheet("HUONG_DAN")
for l in ["Mỗi dòng = 1 trang của 1 mã. Cột J chọn ĐÚNG/SAI/KHÔNG CHẮC: trang này có cần để đọc mã đó không (kích thước, cấu tạo, số lượng, hoàn thiện).",
          "Tô vàng = đáp án và máy KHÔNG đồng ý (chỉ một bên chọn) — ưu tiên xem các dòng này.",
          "Cột I: bấm để mở bản vẽ đúng trang bằng trình duyệt (add-in SRC_LINKS). F023 đã gộp vào F024 (hai file trùng).",
          "Thấy trang đúng mà cả hai bên đều thiếu: ghi ở cột Ghi chú của mã đó (vd 'thiếu F005 p17')."]:
    hd.append([l])
hd.column_dimensions["A"].width = 140
wb.save(OUT)
# tự kiểm link: 0 link trực tiếp ở KIEM_MAU, mỗi ô Mở có dòng _LINKS, file tồn tại
direct = [c.coordinate for r in ws.iter_rows(min_row=2) for c in r if c.hyperlink and not c.hyperlink.location]
keys = {r[0] for r in wl.iter_rows(min_row=2, values_only=True)}
miss = [f"KIEM_MAU!I{i}" for i in range(2, ws.max_row + 1) if f"KIEM_MAU!I{i}" not in keys]
nofile = [r[4] for r in wl.iter_rows(min_row=2, values_only=True) if not os.path.isfile(os.path.join(HERE, r[4]))]
print(pick, ws.max_row - 1, "dòng | link trực tiếp:", len(direct), "| thiếu _LINKS:", len(miss), "| file không có:", len(nofile))
