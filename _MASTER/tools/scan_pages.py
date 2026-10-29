# -*- coding: utf-8 -*-
"""scan_pages.py - doc khung ten ban ve (title block) bang pdfplumber -> PAGES_AUTO + kiem trang trong _LINKS.

Chay:  .venv\\Scripts\\python.exe -I scan_pages.py "<thu muc du an>" [--data PROJECT_DATA_v3.xlsx] [--out PAGES_AUTO.xlsx]
  <thu muc du an> vi du D:\\COSTING EXCEL\\_ACTIVE\\NORTON (co 00_input\\<dot>\\03_BanVe, 02_Spec va 01_scan).
Khong sua file PROJECT_DATA; chi ghi file ket qua moi trong 01_scan.
"""
import argparse, re, sys, time, unicodedata
from pathlib import Path

import openpyxl
import pdfplumber
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

# Vung khung ten (ti le theo kich thuoc trang) - dung chung cho bo ban ve DPA Skybridge.
TB_X0 = 0.860
TB_SPLIT = 0.930            # cot phai cua khung ten (ty le, phien ban)
FIELDS = {                  # ten: (top_min, top_max, x_min, x_max)
    "hang_muc":  (0.859, 0.875, TB_X0, 1.0),
    "ten_vn":    (0.879, 0.900, TB_X0, 1.0),
    "ten_en":    (0.900, 0.912, TB_X0, 1.0),
    "ty_le":     (0.915, 0.924, TB_SPLIT, 1.0),
    "giai_doan": (0.927, 0.935, TB_X0, TB_SPLIT),
    "phien_ban": (0.927, 0.935, TB_SPLIT, 1.0),
    "ngay":      (0.939, 0.946, TB_X0, 1.0),
    "so_bv":     (0.950, 0.957, TB_X0, 1.0),
}
# Loi font trong bo GENERAL: D-gach (Đ) bi doc thanh E-caron (Ě)
FONT_FIX = str.maketrans({"Ě": "Đ", "ě": "đ"})
DWG_CODE = re.compile(r"\b(CD|FA|TD)[.\- ]?(\d{2})[.\- ](\d{2}[A-Z]?)\b")
SPEC_ART = re.compile(r"§\s*(\d\.\d{2})")


def norm(s):
    return unicodedata.normalize("NFC", (s or "").translate(FONT_FIX)).strip()


def join_frags(text):
    """Ghep manh chu bi tach do dau tieng Viet (vd 'PH Ầ N' -> 'PHẦN')."""
    toks = text.split()
    out = []
    for t in toks:
        if out and (len(t) == 1 and not t.isascii() or len(out[-1]) == 1 and not out[-1].isascii()) and t.isalpha() and out[-1].isalpha():
            out[-1] += t
        else:
            out.append(t)
    return " ".join(out)


def region_text(words, W, H, box):
    t0, t1, x0, x1 = box
    sel = [w for w in words if t0 * H <= w["top"] < t1 * H and x0 * W <= w["x0"] < x1 * W]
    lines = []
    for w in sorted(sel, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(w["top"] - lines[-1][0]) < 4:
            lines[-1][1].append(w)
        else:
            lines.append([w["top"], [w]])
    txt = " / ".join(" ".join(x["text"] for x in sorted(ws, key=lambda x: x["x0"])) for _, ws in lines)
    return join_frags(norm(txt))


def dwg_key(s):
    m = DWG_CODE.search(s or "")
    return f"{m.group(1)}.{m.group(2)}.{m.group(3)}" if m else ""


def short_name(pdf_name):
    m = re.match(r"SB_TENDER_PLOT B_(.+?)_\d{4} \d{2} \d{2}\.pdf$", pdf_name)
    return m.group(1) if m else Path(pdf_name).stem


def scan_drawings(bv_dir, door_codes, log):
    rows, page_text = [], {}
    code_re = re.compile(r"\b(" + "|".join(sorted(map(re.escape, door_codes), key=len, reverse=True)) + r")\b") if door_codes else None
    for pdf_path in sorted(bv_dir.glob("SB_TENDER_PLOT B_*.pdf")):
        if re.match(r"SB_TENDER_PLOT B_\d{4} ", pdf_path.name):
            log(f"  bo qua file gop: {pdf_path.name}")
            continue
        t = time.time()
        with pdfplumber.open(pdf_path) as pdf:
            for i, p in enumerate(pdf.pages, 1):
                W, H = p.width, p.height
                words = p.extract_words()
                f = {k: region_text(words, W, H, b) for k, b in FIELDS.items()}
                so_bv = f["so_bv"].replace(" ", "")
                body = [w["text"] for w in words if not (w["x0"] > TB_X0 * W)]
                doors = {}
                if code_re:
                    for w in body:
                        for m in code_re.finditer(w):
                            doors[m.group(1)] = doors.get(m.group(1), 0) + 1
                warn = []
                if not so_bv:
                    warn.append("khong doc duoc so BV")
                if not words:
                    warn.append("trang khong co text (scan/anh?)")
                rows.append({
                    "file": short_name(pdf_path.name), "pdf": pdf_path.name, "trang": i,
                    "so_bv": so_bv, "ma": dwg_key(so_bv), "ten_vn": f["ten_vn"], "ten_en": f["ten_en"],
                    "hang_muc": f["hang_muc"], "ty_le": f["ty_le"], "giai_doan": f["giai_doan"],
                    "phien_ban": f["phien_ban"], "ngay": f["ngay"],
                    "ma_cua": ", ".join(f"{k}({v})" for k, v in sorted(doors.items())),
                    "canh_bao": "; ".join(warn),
                })
            log(f"  {pdf_path.name}: {len(pdf.pages)} trang, {time.time() - t:.0f}s")
    return rows


def scan_specs(spec_dir, log):
    """Moi trang spec -> tap so dieu (vd '3.02') xuat hien dau dong."""
    idx = {}
    for pdf_path in sorted(spec_dir.glob("*.pdf")):
        arts = []
        with pdfplumber.open(pdf_path) as pdf:
            for p in pdf.pages:
                txt = norm(p.extract_text() or "")
                arts.append(set(re.findall(r"(?m)^\s*(\d\.\d{2})\b", txt)))
        idx[pdf_path.name] = arts
        log(f"  {pdf_path.name}: {len(arts)} trang")
    return idx


def check_links(data_xlsx, dwg_rows, spec_idx):
    by_pdf = {}
    for r in dwg_rows:
        by_pdf.setdefault(r["pdf"], {})[r["trang"]] = r
    wb = openpyxl.load_workbook(data_xlsx, read_only=True)
    ws = wb["_LINKS"]
    out = []
    for n, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        key, _, _, label, rel, trang = (list(row) + [None] * 6)[:6]
        if not rel or not str(rel).lower().endswith(".pdf") or trang in (None, ""):
            continue
        name = str(rel).replace("/", "\\").split("\\")[-1]
        label = str(label or "")
        try:
            trang = int(trang)
        except (TypeError, ValueError):
            out.append((n, key, label, name, trang, "", "", "LOI", "Trang khong phai so"))
            continue
        if name in by_pdf:
            pages = by_pdf[name]
            want = dwg_key(label)
            if trang not in pages:
                out.append((n, key, label, name, trang, want, "", "LOI", f"File chi co {len(pages)} trang"))
                continue
            got = pages[trang]["ma"]
            if not want:
                out.append((n, key, label, name, trang, "", got, "-", "Nhan khong ghi ma BV"))
            elif want == got:
                out.append((n, key, label, name, trang, want, got, "OK", ""))
            else:
                sug = [str(k) for k, v in pages.items() if v["ma"] == want]
                out.append((n, key, label, name, trang, want, got, "SAI",
                            ("Goi y trang " + ",".join(sug)) if sug else "Khong tim thay ma trong file"))
        elif name in spec_idx:
            arts = spec_idx[name]
            m = SPEC_ART.search(label)
            if trang > len(arts):
                out.append((n, key, label, name, trang, m.group(1) if m else "", "", "LOI", f"File chi co {len(arts)} trang"))
                continue
            if not m:
                out.append((n, key, label, name, trang, "", "", "-", "Nhan khong ghi so dieu"))
                continue
            art = m.group(1)
            # dieu co the bat dau o trang truoc va keo sang trang nay: dieu cuoi cung xuat hien truoc trang nay = art
            prev = [a for k in range(trang - 1) for a in sorted(arts[k])]
            if art in arts[trang - 1]:
                out.append((n, key, label, name, trang, art, "co tren trang", "OK", ""))
            elif prev and prev[-1] == art:
                out.append((n, key, label, name, trang, art, "dieu bat dau trang truoc", "OK", ""))
            else:
                sug = [str(k + 1) for k, a in enumerate(arts) if art in a]
                out.append((n, key, label, name, trang, art, "", "SAI",
                            ("Dieu bat dau o trang " + ",".join(sug)) if sug else "Khong thay dieu trong file"))
    return out


def door_codes_from(data_xlsx):
    wb = openpyxl.load_workbook(data_xlsx, read_only=True)
    ws = wb["BOQ_TONGHOP"]
    hdr = next(ws.iter_rows(max_row=1, values_only=True))
    col = list(hdr).index("Mã cửa")
    return {str(r[col]).strip() for r in ws.iter_rows(min_row=2, values_only=True) if r[col]}


def write_xlsx(out, dwg_rows, links, meta):
    wb = openpyxl.Workbook()
    bold, head = Font(bold=True), PatternFill("solid", fgColor="DDEBF7")
    s = wb.active; s.title = "PAGES_AUTO"
    cols = [("File", "file"), ("Trang", "trang"), ("Số BV", "so_bv"), ("Mã", "ma"), ("Tên BV (VN)", "ten_vn"),
            ("Tên BV (EN)", "ten_en"), ("Hạng mục", "hang_muc"), ("Tỷ lệ", "ty_le"), ("Giai đoạn", "giai_doan"),
            ("Phiên bản", "phien_ban"), ("Ngày", "ngay"), ("Mã cửa xuất hiện (SL)", "ma_cua"), ("Cảnh báo", "canh_bao")]
    s.append([c for c, _ in cols])
    for r in dwg_rows:
        s.append([r[k] for _, k in cols])
    widths = [16, 7, 24, 11, 40, 34, 28, 12, 10, 10, 9, 40, 24]
    l = wb.create_sheet("LINKS_CHECK")
    l.append(["Dòng _LINKS", "Key", "Nhãn", "File", "Trang", "Mã/điều cần", "Thấy trên trang", "Kết quả", "Ghi chú"])
    fills = {"SAI": PatternFill("solid", fgColor="F8CBAD"), "LOI": PatternFill("solid", fgColor="F8CBAD"),
             "OK": PatternFill("solid", fgColor="E2EFDA")}
    for row in links:
        l.append(list(row))
        if row[7] in fills:
            l.cell(l.max_row, 8).fill = fills[row[7]]
    for ws_, w in ((s, widths), (l, [10, 18, 40, 44, 7, 12, 22, 9, 36])):
        for i, x in enumerate(w, 1):
            ws_.column_dimensions[get_column_letter(i)].width = x
            ws_.cell(1, i).font = bold; ws_.cell(1, i).fill = head
        ws_.freeze_panes = "A2"; ws_.auto_filter.ref = ws_.dimensions
    m = wb.create_sheet("SUMMARY")
    for k, v in meta:
        m.append([k, v])
    m.column_dimensions["A"].width = 34; m.column_dimensions["B"].width = 80
    wb.save(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("project")
    ap.add_argument("--data", default="PROJECT_DATA_v3.xlsx")
    ap.add_argument("--out", default="PAGES_AUTO.xlsx")
    a = ap.parse_args()
    proj = Path(a.project)
    scan = proj / "01_scan"
    data = scan / a.data
    inp = sorted((proj / "00_input").glob("*/03_BanVe"))[-1].parent
    log = lambda s: print(s, flush=True)
    t = time.time()
    log(f"Input: {inp}")
    doors = door_codes_from(data)
    log(f"Ma cua tu BOQ_TONGHOP: {len(doors)}")
    log("Doc ban ve...")
    dwg = scan_drawings(inp / "03_BanVe", doors, log)
    log("Doc spec...")
    spec = scan_specs(inp / "02_Spec", log)
    log("Kiem _LINKS...")
    links = check_links(data, dwg, spec)
    cnt = {}
    for r in links:
        cnt[r[7]] = cnt.get(r[7], 0) + 1
    meta = [("Thời điểm", time.strftime("%Y-%m-%d %H:%M")), ("Input", str(inp)), ("PROJECT_DATA", str(data)),
            ("Số trang bản vẽ", len(dwg)), ("Trang không đọc được số BV", sum(1 for r in dwg if not r["so_bv"])),
            ("Mã cửa dò", len(doors)), ("_LINKS (PDF có trang) đã kiểm", len(links))] + \
           [(f"  Kết quả {k}", v) for k, v in sorted(cnt.items())] + [("Thời gian (s)", round(time.time() - t))]
    out = scan / a.out
    write_xlsx(out, dwg, links, meta)
    for k, v in meta:
        log(f"{k}: {v}")
    log(f"Da ghi: {out}")


if __name__ == "__main__":
    main()
