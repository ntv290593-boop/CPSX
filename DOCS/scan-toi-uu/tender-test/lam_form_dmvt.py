# -*- coding: utf-8 -*-
r"""lam_form_dmvt.py — dựng SKILL\10-tender\form\DMVT_form.xlsx từ form user lập, bỏ hết dữ liệu dự án. Chạy 1 lần, chạy lại khi user đổi form.

Chạy: py -I lam_form_dmvt.py "<file DMVT của user>"
Giữ: hàng 1–5 (tiêu đề), hàng 6, hàng 7 (mẫu mục), hàng 8 (mẫu dòng), độ rộng cột. Bỏ: dữ liệu, ảnh, comment, tên định nghĩa, thiết lập máy in.
Thêm AG:AJ (máy ghi): Spec mới khác, Cờ, Link spec, Link bản vẽ.
"""
import os
import re
import shutil
import sys
import tempfile
import zipfile

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", ".."))
RA = os.path.join(ROOT, "SKILL", "10-tender", "form", "DMVT_form.xlsx")
COT_THEM = [("AG", "SPEC MỚI KHÁC (máy ghi)", 40), ("AH", "CỜ (máy ghi)", 40), ("AI", "LINK SPEC", 24), ("AJ", "LINK BẢN VẼ", 24)]


def cat_ten_rac(src, ra):
    """Form gốc có hàng nghìn tên định nghĩa rác (940 KB workbook.xml) sai cú pháp nên COM không xóa được: cắt trong zip, chỉ giữ _xlnm (vùng in)."""
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(ra, "w", zipfile.ZIP_DEFLATED) as b:
        for n in a.namelist():
            d = a.read(n)
            if n == "xl/workbook.xml":
                x = d.decode("utf-8")
                giu = re.findall(r"<definedName [^>]*name=\"_xlnm[^>]*>.*?</definedName>", x, re.S)
                x = re.sub(r"<definedNames>.*?</definedNames>", ("<definedNames>" + "".join(giu) + "</definedNames>") if giu else "", x, flags=re.S)
                d = x.encode("utf-8")
            b.writestr(n, d)


def main():
    import win32com.client
    src = sys.argv[1]
    tmp = tempfile.mkdtemp(prefix="dmvt_form_")
    f = os.path.join(tmp, "DMVT_form.xlsx")
    shutil.copy2(src, f)
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    try:
        wb = xl.Workbooks.Open(f, UpdateLinks=0)
        ws = wb.Worksheets("DMVT")
        for i in range(ws.Shapes.Count, 0, -1):
            ws.Shapes.Item(i).Delete()
        ws.Cells.ClearComments()
        cuoi = ws.UsedRange.Rows.Count + ws.UsedRange.Row
        ws.Range(f"9:{cuoi + 5}").Delete()
        ws.Range("A7:AF8").ClearContents()
        ws.Range("B6").ClearContents()
        ws.Range("A2").Value = ""  # ô merge: ClearContents trên 1 phần vùng merge báo lỗi
        ws.Range("A3").Value = ""
        ws.Range("A1").Value = "DANH MỤC VẬT TƯ – LIST OF MATERIAL"
        for c, t, w in COT_THEM:
            ws.Range(f"AF5").Copy(ws.Range(f"{c}5"))
            ws.Range(f"{c}5").Value = t
            ws.Range("I8").Copy(ws.Range(f"{c}8"))
            ws.Range(f"{c}8").ClearContents()
            ws.Columns(c).ColumnWidth = w
        ws.Range("AG4").Value = "NỘI BỘ – xóa trước khi gửi CĐT"
        ws.Range("AG4").Font.Bold = True
        ws.Range("AG4").Font.Color = 0x0000C0
        wb.Saved = False
        wb.Close(SaveChanges=True)
    finally:
        xl.Quit()
    os.makedirs(os.path.dirname(RA), exist_ok=True)
    cat_ten_rac(f, RA)
    # soát không còn dữ liệu dự án
    with zipfile.ZipFile(RA) as z:
        ten = z.namelist()
        chu = z.read("xl/sharedStrings.xml").decode("utf-8") if "xl/sharedStrings.xml" in ten else ""
        media = [n for n in ten if "/media/" in n]
    bay = [x for x in ("SOFITEL", "Sofitel", "KAWAJUN", "HAFELE", "Royal Thai", "51-", "IR-0", "CHECK VE") if x in chu]
    print(f"Đã ghi {RA}: {os.path.getsize(RA)} byte, {len(media)} ảnh, chữ dự án còn lại: {bay or 'không'}")
    print("Chữ:", re.sub(r"<[^>]+>", "|", chu)[:900])
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
