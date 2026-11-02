# -*- coding: utf-8 -*-
r"""doc_file.py — đọc 1 file đầu vào cho người xem: mail Outlook, chữ và ảnh trang PDF, ô Excel.
Gộp từ pd.py (GĐ6 plan Code ID/DMVT), kết quả giữ giống hệt bản cũ. Hàm thuần, không biết đường dẫn dự án, form, bước.
Mỗi hàm trả chữ để người gọi in; không giữ file mở sau khi trả về.

  chu(pdf, trang, clip)            lớp chữ của 1 trang
  so(pdf, trang, clip)             các số trong lớp chữ kèm toạ độ
  anh(pdf, trang, png, dpi, clip)  xuất ảnh trang hoặc vùng -> (rộng, cao)
  mail(msg)                        tiêu đề, người gửi/nhận, đính kèm, nội dung file .msg
  xem(xlsx, sheet, vung)           giá trị ô; không nêu sheet thì liệt kê sheet
clip = "x0,y0,x1,y1" theo tỉ lệ 0..1 của trang.
"""
import os


def _vung(page, clip):
    if not clip:
        return None
    import pymupdf
    x0, y0, x1, y1 = [float(v) for v in clip.split(",")]
    r = page.rect
    return pymupdf.Rect(r.x0 + x0 * r.width, r.y0 + y0 * r.height, r.x0 + x1 * r.width, r.y0 + y1 * r.height)


def chu(pdf, trang, clip=None):
    import pymupdf
    with pymupdf.open(pdf) as d:
        p = d[int(trang) - 1]
        return p.get_text(clip=_vung(p, clip), sort=True)


def so(pdf, trang, clip=None):
    """Các số trong lớp chữ kèm toạ độ tỉ lệ (đọc kích thước, chỉ khi trang có lớp chữ)."""
    import re
    import pymupdf
    with pymupdf.open(pdf) as d:
        p = d[int(trang) - 1]
        W, H = p.rect.width, p.rect.height
        return "\n".join(f"{w[4]}\t x={w[0] / W:.3f} y={w[1] / H:.3f}" for w in p.get_text("words", clip=_vung(p, clip), sort=True)
                         if re.fullmatch(r"[\d.,x×/+\-]+", w[4]) and re.search(r"\d", w[4]))


def anh(pdf, trang, png, dpi=150, clip=None):
    import pymupdf
    with pymupdf.open(pdf) as d:
        p = d[int(trang) - 1]
        pix = p.get_pixmap(dpi=int(dpi), clip=_vung(p, clip))
    os.makedirs(os.path.dirname(os.path.abspath(png)), exist_ok=True)
    pix.save(png)
    return pix.width, pix.height


def mail(msg):
    import olefile
    with olefile.OleFileIO(msg) as o:
        def get(tag, base=""):
            for enc, suf in (("utf-16-le", "001F"), ("cp1252", "001E")):
                p = f"{base}__substg1.0_{tag}{suf}"
                if o.exists(p):
                    return o.openstream(p).read().decode(enc, "replace").strip("\x00")
            return ""
        att = sorted({e[0] for e in o.listdir() if e[0].startswith("__attach")})
        return "\n".join([
            f"Tiêu đề : {get('0037')}",
            f"Từ      : {get('0C1A')} {get('0C1F')}",
            f"Tới     : {get('0E04')}",
            f"Cc      : {get('0E03')}",
            f"Đính kèm: {[get('3707', a + '/') or get('3704', a + '/') for a in att]}",
            "-" * 60,
            get("1000"),
        ])


def xem(xlsx, sheet=None, vung=None):
    import openpyxl
    from openpyxl.utils import get_column_letter as L, range_boundaries
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    try:
        if sheet is None:
            return "\n".join(f"[{ws.title}] {ws.max_row}x{ws.max_column}" for ws in wb.worksheets)
        ws = wb[sheet]
        vung = vung or f"A1:{L(min(ws.max_column or 26, 30))}{min(ws.max_row or 60, 80)}"
        c0, r0, c1, r1 = range_boundaries(vung)
        ra = []
        for i, row in enumerate(ws.iter_rows(min_row=r0, max_row=r1, min_col=c0, max_col=c1, values_only=True), r0):
            v = [f"{L(c0 + j)}={str(x)[:60]}" for j, x in enumerate(row) if x not in (None, "")]
            if v:
                ra.append(f"{i} " + " | ".join(v))
        return "\n".join(ra)
    finally:
        wb.close()
