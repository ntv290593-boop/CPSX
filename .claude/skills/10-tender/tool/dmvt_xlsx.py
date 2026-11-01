"""dmvt_xlsx.py — đọc bản DMVT trước, giữ sửa tay, ghi file bằng Excel COM (plan mục 5).

Không lưu bằng openpyxl: mất ảnh WMF và một số thành phần của form (QUY_TRINH bước 6). openpyxl chỉ dùng để đọc bản cũ.
Nhận ra sửa tay bằng snapshot: _data\\_dmvt_snapshot.json lưu giá trị A:I máy đã ghi và dấu ảnh theo từng dòng; ô hoặc ảnh khác snapshot = người đã sửa.
"""
import glob
import hashlib
import io
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict

import td
from td import s, jload, jdump, tach_ma, link_item, TODAY, ma_chuan, split_list
import dmvt

HERE = os.path.dirname(os.path.abspath(__file__))
FORM_XLSX = os.path.join(HERE, "..", "form", "DMVT_form.xlsx")
SHEET = "DMVT"
HANG_DAU = 8  # hàng dữ liệu đầu tiên của form
MAU_MUC, MAU_DONG = 7, 8  # hàng mẫu định dạng trong form
COT_AI = {"A": 1, "B": 2, "C": 3, "D": 4, "E": 5, "F": 6, "G": 7, "H": 8, "I": 9}
COT_MAY = {"C": "Area", "D": "Material", "E": "Location", "F": "Description", "G": "NCC", "I": "Remarks"}
COT_SPEC_MOI, COT_CO, COT_LINK_SPEC, COT_LINK_BV = 33, 34, 35, 36  # AG, AH, AI, AJ: máy ghi; J:AF và từ AK trở đi là của tender
TU_CU_DAU, TU_CU_CUOI = 10, 32  # J:AF
CAO_MIN, CAO_MAX = 150, 409
SO_KY_TU = {"C": 34, "D": 21, "E": 35, "F": 44, "G": 28, "I": 23}  # ký tự mỗi dòng theo độ rộng cột của form
MAU_XAM = 0xD9D9D9
MAX_LINK_O = 30  # số link tối đa mỗi ô; quá nhiều thì menu của add-in không dùng được, hàng dư ghi vào nhãn


def ten_file(p, stamp=None):
    return f"{p.cfg['du_an']}_DMVT_{stamp or TODAY.strftime('%y%m%d')}.xlsx"


def chuan_o(v):
    """So 2 ô: bỏ khác biệt khoảng trắng, xuống dòng và dấu nháy đầu ô."""
    return re.sub(r"\s+", " ", s(v)).strip()


def tap_ma(chu):
    return frozenset(x[0] for x in tach_ma(s(chu)) if x[3] in ("ffe", "gach"))


# ---------------------------------------------------------------- đọc file cũ
NS = {"x": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def anh_trong_file(f, sheet=SHEET):
    """Số hàng -> (bytes, đuôi) đọc thẳng từ zip, kể cả ảnh openpyxl không nạp được (WMF). Bỏ comment vì comment nằm ở vmlDrawing."""
    out = {}
    with zipfile.ZipFile(f) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rel_wb = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        rid = next((x.get(NS_R + "id") for x in wb.iter() if x.tag.endswith("}sheet") and x.get("name") == sheet), None)
        if not rid:
            return out
        p_sh = rel_wb[rid].lstrip("/")
        p_sh = p_sh if p_sh.startswith("xl/") else "xl/" + p_sh
        p_rel = os.path.dirname(p_sh) + "/_rels/" + os.path.basename(p_sh) + ".rels"
        if p_rel not in z.namelist():
            return out
        dr = next((r.get("Target") for r in ET.fromstring(z.read(p_rel)) if r.get("Type").endswith("/drawing")), None)
        if not dr:
            return out
        p_dr = os.path.normpath(os.path.join(os.path.dirname(p_sh), dr)).replace("\\", "/") if not dr.startswith("/") else dr.lstrip("/")
        rel_dr = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read(os.path.dirname(p_dr) + "/_rels/" + os.path.basename(p_dr) + ".rels"))}
        for a in ET.fromstring(z.read(p_dr)):
            hang, blip = a.find("x:from/x:row", NS), a.find(".//a:blip", NS)
            if hang is None or blip is None:
                continue
            t = rel_dr.get(blip.get(NS_R + "embed"), "")
            t = t.lstrip("/") if t.startswith("/") else os.path.normpath(os.path.join(os.path.dirname(p_dr), t)).replace("\\", "/")
            if t in z.namelist():
                out[int(hang.text) + 1] = (z.read(t), t.rsplit(".", 1)[-1].lower())
    return out


def doc_file_cu(f):
    """Dòng của bản trước. 'ma' = hàng có mã ở cột B; 'duoi' = hàng không mã, đi theo dòng có mã phía trên; hàng tiêu đề mục bỏ qua."""
    import openpyxl
    wb = openpyxl.load_workbook(f)
    ws = wb[SHEET]
    anh = anh_trong_file(f)
    out, cuoi = [], None
    for h in range(HANG_DAU, ws.max_row + 1):
        ai = {c: ws.cell(h, j).value for c, j in COT_AI.items() if c != "H"}
        sau = {j: ws.cell(h, j).value for j in list(range(TU_CU_DAU, TU_CU_CUOI + 1)) + list(range(COT_LINK_BV + 1, ws.max_column + 1)) if ws.cell(h, j).value is not None}
        co_gi = any(s(v) for v in ai.values()) or sau or h in anh
        if not co_gi:
            continue
        ma = tap_ma(ai["B"])
        la_muc = not ma and s(ai["A"]) and not s(ai["A"]).isdigit() and not any(s(ai[c]) for c in "CDEFGI") and not sau
        if la_muc:
            continue
        r = {"hang": h, "ai": ai, "sau": sau, "cao": ws.row_dimensions[h].height, "anh": anh.get(h)}
        if ma:
            r.update(loai="ma", tap=ma)
            r["duoi"] = []
            out.append(r)
            cuoi = r
        elif cuoi:
            r["loai"] = "duoi"
            cuoi["duoi"].append(r)
        else:
            r.update(loai="ma", tap=frozenset())
            r["duoi"] = []
            out.append(r)
    wb.close()
    return out


def md5(b):
    return hashlib.md5(b).hexdigest()


def dau_anh(b):
    """Dấu của ảnh để nhận ra người thay ảnh: 16x16 RGB dạng hex. Không dùng md5 vì Excel nén lại ảnh khi lưu, bytes đổi dù không ai sửa.
    Ảnh PIL không đọc được (WMF) mới dùng md5."""
    if not b:
        return ""
    from PIL import Image
    try:
        with Image.open(io.BytesIO(b)) as im:
            return "px" + im.convert("RGB").resize((16, 16), Image.LANCZOS).tobytes().hex()
    except Exception:
        return md5(b)


def cung_anh(a, b):
    """Khác trung bình <= 1% mỗi kênh màu: cùng ảnh, chỉ khác do nén lại."""
    if a == b:
        return True
    if not (a.startswith("px") and b.startswith("px")) or len(a) != len(b):
        return False
    x, y = bytes.fromhex(a[2:]), bytes.fromhex(b[2:])
    return sum(abs(p - q) for p, q in zip(x, y)) / len(x) <= 255 * 0.01


# ---------------------------------------------------------------- gộp máy + sửa tay
def gop_sua_tay(may, cu, snap):
    """may: dòng máy vừa bóc (đã có chữ cuối từng cột). cu: doc_file_cu. snap: {khóa: {cột: chữ, 'anh': md5}}.
    Trả [dòng ra], mỗi dòng: chữ từng cột, 'sau', 'anh_cu', 'duoi', 'Spec mới khác', 'Cờ'. Dòng cũ không còn khớp thì giữ ở cuối mục nếu có sửa tay."""
    cu_theo_tap = {r["tap"]: r for r in cu if r["tap"]}
    dung = set()
    ra = []
    tk = Counter()

    def co_sua(r, k):
        """Dòng cũ có công người: J trở đi, dòng đuôi, ô A:I khác snapshot, ảnh khác snapshot, dòng ngoài snapshot."""
        sn = snap.get(k)
        if r["sau"] or r["duoi"] or not sn or sn.get("_giu"):
            return True
        if any(chuan_o(r["ai"][c]) != chuan_o(sn.get(c)) for c in COT_MAY):
            return True
        return not cung_anh(dau_anh(r["anh"][0]) if r["anh"] else "", sn.get("anh", ""))

    for m in may:
        t = frozenset(m["Mã"])
        d = dict(m, spec_moi=[], sau={}, duoi=[], anh_cu=None, cao_cu=None, moc=dict(m["o"]), moc_anh=None, moc_hash=m.get("Hash ảnh", ""))
        d["Cờ"] = list(m["Cờ"])
        o = cu_theo_tap.get(t)
        k = m["Khóa"]
        if o:
            dung.add(t)
            k_cu = khoa_cua_tap(o, snap)
            sn = snap.get(k_cu) or {}
            for c, ten in COT_MAY.items():
                cu_tay, cu_may, moi = chuan_o(o["ai"][c]), chuan_o(sn.get(c)), chuan_o(m["o"].get(c))
                if cu_tay != cu_may and k_cu in snap:  # người đã sửa ô này
                    d["o"] = dict(d["o"], **{c: o["ai"][c]})
                    d["moc"][c] = sn.get(c, "")  # mốc vẫn là chữ máy cũ: nếu ghi chữ người vào mốc thì lần build sau tưởng chưa ai sửa và đè mất
                    tk["ô sửa tay giữ"] += 1
                    if moi != cu_may and moi != cu_tay:  # người đã gõ đúng chữ spec mới thì thôi báo
                        d["spec_moi"].append(f"{c}: {m['o'].get(c)}")
                        tk["spec mới khác"] += 1
            d["sau"] = o["sau"]
            d["duoi"] = o["duoi"]
            d["cao_cu"] = o["cao"]
            dau_cu = dau_anh(o["anh"][0]) if o["anh"] else ""
            if k_cu in snap and not cung_anh(dau_cu, sn.get("anh", "")):  # người thay hoặc xóa ảnh
                d["anh_cu"] = o["anh"] or ("xoa",)
                d["moc_anh"], d["moc_hash"] = sn.get("anh", ""), sn.get("hash", "")
                tk["ảnh sửa tay giữ"] += 1
                if m.get("Hash ảnh") and sn.get("hash") and m["Hash ảnh"] != sn["hash"]:
                    d["spec_moi"].append("H: ảnh trong spec đã đổi")
        else:
            lq = [r for r in cu if r["tap"] and r["tap"] & t and r["tap"] != t]
            if lq:
                d["Cờ"].append("Tập mã đổi so với bản trước: " + " | ".join(" / ".join(sorted(r["tap"])) for r in lq) + ". Sửa tay không tự chuyển")
                tk["tách/gộp tập mã"] += 1
        ra.append(d)
    # dòng cũ không còn khớp tập mã: giữ nếu có công người, không thì bỏ
    giu = []
    for r in cu:
        if r["tap"] in dung and r["tap"]:
            continue
        k = khoa_cua_tap(r, snap)
        if not co_sua(r, k):
            tk["dòng cũ bỏ"] += 1
            continue
        lq = any(set(r["tap"]) & set(m["Mã"]) for m in may)
        giu.append({"Khóa": " / ".join(sorted(r["tap"])) or f"(hàng {r['hang']})", "Mã": sorted(r["tap"]), "o": {c: r["ai"][c] for c in COT_MAY}, "ma_goc": r["ai"]["B"],
                    "Cờ": ["Dòng cũ, tập mã đã tách/gộp" if lq else ("Không còn trong spec" if k in snap else "Dòng user thêm")], "Mục": "", "Tên mục": "DÒNG CŨ GIỮ LẠI, CẦN XEM",
                    "Omit": False, "spec_moi": [], "sau": r["sau"], "duoi": r["duoi"], "anh_cu": r["anh"], "cao_cu": r["cao"], "Cờ ảnh": []})
        tk["dòng cũ giữ"] += 1
    return ra + giu, tk


def moc_snapshot(r):
    """Giá trị máy làm mốc để lần sau nhận ra sửa tay. Dòng cũ giữ lại không có mốc: lần sau vẫn tính là công người."""
    moc = r.get("moc")
    if moc is None:
        return {"_ma": r["Mã"], "_giu": True}
    return dict({c: s(moc.get(c)) for c in COT_MAY}, _ma=r["Mã"], anh=r.get("moc_anh"), hash=r.get("moc_hash", ""))


def khoa_cua_tap(r, snap):
    """Khóa snapshot của dòng cũ: so theo tập mã để không phụ thuộc thứ tự mã người gõ ở cột B."""
    for k in snap:
        if frozenset(snap[k].get("_ma", [])) == r["tap"]:
            return k
    return " / ".join(sorted(r["tap"]))


# ---------------------------------------------------------------- chữ từng cột
def chu_cuoi(rows, llm):
    """Ghép chữ gốc của spec với phần LLM. Thiếu bản dịch thì giữ tiếng Anh và có cờ, không để trống."""
    for r in rows:
        x = llm.get(r["Khóa"]) or {}
        co = list(r["Cờ"])
        d = r["Material"]
        if x.get("Material VN"):
            d = f"{d}\n{x['Material VN']}"
        else:
            co.append("Chưa dịch")
        rem = "\n".join(v for v in (r["Remarks"], x.get("Ghi chú")) if v)
        r["o"] = {"C": r["Area"], "D": d, "E": x.get("Location VN") or r["Location"], "F": r["Description"], "G": r["NCC"], "I": rem}
        r["Cờ"] = co
    return rows


def cao_dong(o, co_anh=True):
    n = 1
    for c, w in SO_KY_TU.items():
        tong = sum(max(1, -(-len(l) // w)) for l in s(o.get(c)).split("\n"))
        n = max(n, tong)
    return min(CAO_MAX, max(CAO_MIN if co_anh else 40, n * 15.5 + 8))


def an_toan_o(v):
    """Chữ bắt đầu = + - @ mà ghi thẳng qua COM thì Excel hiểu là công thức."""
    t = s(v)
    return "'" + t if t[:1] in "=+-@" else t


def link_cua_dong(r, res, p):
    """(link spec, link bản vẽ) từ nguồn đã bóc + VL_XUATHIEN. Mỗi trang 1 link."""
    fidx = res["FILE_INDEX"]
    sp, bv = [], []
    for n in r.get("Nguồn", []):
        sp.append([f"{n['ma']} · {dmvt.ten_file(fidx, n['fid'])[:40]} p{n['trang']}", n["fid"], n["trang"]])
    tap = set(r["Mã"])
    for x in res["VL_XUATHIEN"].values():
        if x.get("Mã VL") in tap:
            if x.get("Nguồn") == "Bản vẽ":
                bv += td.links_of(x.get("File"))
    sp, bv = td.moi_trang_mot_link(sp), td.moi_trang_mot_link(bv)
    return [link_item(p, l, fidx) for l in sp], [link_item(p, l, fidx) for l in bv]


# ---------------------------------------------------------------- COM
def mo_excel():
    import win32com.client
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    xl.AutomationSecurity = 3  # không chạy macro
    return xl


def them_link(ws, o, dich, chu):
    """Link trong workbook; đích trỏ chính ô thì add-in SRC_LINKS đọc _LINKS để mở PDF đúng trang. Ghi chữ sau cùng vì Hyperlinks.Add có thể ghi địa chỉ vào ô."""
    ws.Hyperlinks.Add(o, "", dich)
    o.Value = chu


def ghi_o(ws, h, j, v):
    ws.Cells(h, j).Value = an_toan_o(v) if isinstance(v, str) else v


def dat_anh(ws, f, h, le=4):
    """Chèn ảnh vừa ô cột H, căn giữa, Placement=1. Gọi sau khi đã chốt chiều cao dòng và độ rộng cột."""
    from PIL import Image
    c = ws.Cells(h, COT_AI["H"])
    with Image.open(f) as im:
        w0, h0 = im.size
    k = min((c.Width - 2 * le) / w0, (c.Height - 2 * le) / h0)
    w, cao = w0 * k, h0 * k
    sh = ws.Shapes.AddPicture(os.path.abspath(f), False, True, c.Left + (c.Width - w) / 2, c.Top + (c.Height - cao) / 2, w, cao)
    sh.Placement = 1
    sh.AlternativeText = f"Ảnh hàng {h}"


def build(p, res, cfg, args):
    a = list(args)
    ra = td.opt(a, "--ra")
    if not os.path.exists(FORM_XLSX):
        sys.exit(f"LỖI: thiếu form {FORM_XLSX}")
    d_boq = ra or os.path.join(p.t, "03_boq")
    os.makedirs(d_boq, exist_ok=True)
    out = os.path.join(d_boq, ten_file(p))
    if os.path.exists(os.path.join(d_boq, "~$" + os.path.basename(out))):
        sys.exit(f"LỖI: {out} đang mở trong Excel — đóng file rồi chạy lại.")
    rows = dmvt.cmd_boc(p, res, cfg)
    llm = jload(p.f("_dmvt_llm.json")) or {}
    chu_cuoi(rows, llm)
    snap = jload(p.f("_dmvt_snapshot.json")) or {}
    cu_file = snap.get("_file") if snap else None
    if cu_file and not os.path.isabs(cu_file):
        cu_file = os.path.join(p.t, cu_file)
    if not (cu_file and os.path.exists(cu_file)):
        ds = sorted(glob.glob(os.path.join(p.t, "03_boq", f"{p.cfg['du_an']}_DMVT_*.xlsx")))
        cu_file = ds[-1] if ds else None
    cu = doc_file_cu(cu_file) if cu_file else []
    dong, tk = gop_sua_tay(rows, cu, {k: v for k, v in snap.items() if not k.startswith("_")})
    # thứ tự mục theo cấu hình, trong mục theo thứ tự máy; dòng cũ giữ lại nằm cuối
    thu_tu = {n["so"]: i for i, n in enumerate(cfg["nhom"])}
    dong.sort(key=lambda r: (thu_tu.get(r["Mục"], 99)))
    if os.path.exists(out):
        bk = os.path.join(d_boq, "_backup")
        os.makedirs(bk, exist_ok=True)
        shutil.copy2(out, os.path.join(bk, f"{os.path.basename(out)[:-5]}_{time.strftime('%H%M%S')}.xlsx"))
    p.base, p.links = d_boq, []
    tmp = tempfile.mkdtemp(prefix="dmvt_")
    f_tmp = os.path.join(tmp, os.path.basename(out))
    shutil.copy2(FORM_XLSX, f_tmp)
    thu_muc_anh = p.f("_dmvt_anh")
    xl = mo_excel()
    wb = None
    hang_cua, snap_moi, soi = {}, {}, []
    try:
        wb = xl.Workbooks.Open(f_tmp, UpdateLinks=0)
        ws = wb.Worksheets(SHEET)
        for o, v in (cfg.get("tieu_de") or {}).items():
            ws.Range(o).Value = v
        mau = wb.Worksheets.Add(None, ws)
        mau.Name = "_mau"
        ws.Rows(MAU_MUC).Copy(mau.Rows(1))
        ws.Rows(MAU_DONG).Copy(mau.Rows(2))
        cao_muc = ws.Rows(MAU_MUC).RowHeight
        ws.Range(f"{MAU_MUC}:{MAU_DONG}").ClearContents()
        h = MAU_MUC
        muc_hien, stt = None, 0
        o_anh = []
        for r in dong:
            if (r["Mục"], r["Tên mục"]) != muc_hien:
                muc_hien = (r["Mục"], r["Tên mục"])
                mau.Rows(1).Copy(ws.Rows(h))
                ws.Rows(h).RowHeight = cao_muc
                ghi_o(ws, h, 1, r["Mục"])
                ghi_o(ws, h, 2, r["Tên mục"])
                h += 1
            mau.Rows(2).Copy(ws.Rows(h))
            stt += 1
            ghi_o(ws, h, 1, stt)
            ghi_o(ws, h, 2, r.get("ma_goc") or "\n".join(r["Mã"]))
            for c, ten in COT_MAY.items():
                v = r["o"].get(c)
                if s(v):
                    ghi_o(ws, h, COT_AI[c], v)
            co_anh = bool(r.get("Ảnh") or r.get("anh_cu"))
            ws.Rows(h).RowHeight = r.get("cao_cu") or cao_dong(r["o"], True)
            for j, v in r["sau"].items():
                c = ws.Cells(h, j)
                if isinstance(v, str) and v.startswith("="):
                    c.Formula = v
                else:
                    c.Value = v
            if r["spec_moi"]:
                ghi_o(ws, h, COT_SPEC_MOI, "\n".join(r["spec_moi"]))
            co = list(dict.fromkeys(r["Cờ"] + r.get("Cờ ảnh", [])))
            if co:
                ghi_o(ws, h, COT_CO, "; ".join(co))
            if r.get("Omit"):
                ws.Range(f"A{h}:I{h}").Interior.Color = MAU_XAM
            # ảnh: bản người thay thắng ảnh máy cắt
            if r.get("anh_cu") and r["anh_cu"][0] != "xoa":
                f_anh = os.path.join(tmp, f"cu_{h}.{r['anh_cu'][1]}")
                with open(f_anh, "wb") as fh:
                    fh.write(r["anh_cu"][0])
                o_anh.append((h, f_anh))
            elif not r.get("anh_cu") and r.get("Ảnh"):
                o_anh.append((h, os.path.join(thu_muc_anh, r["Ảnh"])))
            if r.get("Nguồn") is not None or r["Mã"]:
                sp, bv = link_cua_dong(r, res, p)
                for j, ds, ten in ((COT_LINK_SPEC, sp, "trang spec"), (COT_LINK_BV, bv, "tờ bản vẽ")):
                    ds = [x for x in ds if x[1]]
                    if not ds:
                        continue
                    o = ws.Cells(h, j)
                    du = max(0, len(ds) - MAX_LINK_O)
                    nhan = f"{min(len(ds), MAX_LINK_O)} {ten}" + (f" (còn {du} ở VL_XUATHIEN của BOQ_TRACK)" if du else "")
                    dc = o.Address.replace("$", "")
                    them_link(ws, o, f"'{SHEET}'!{dc}", nhan)
                    for lab, f, pg, sub in ds[:MAX_LINK_O]:
                        p.links.append((SHEET, dc, lab, f, pg, sub))
            hang_cua[r["Khóa"]] = h
            for m in r["Mã"]:
                hang_cua[f"#{m}"] = h
            snap_moi[r["Khóa"]] = moc_snapshot(r)
            if r.get("Cờ ảnh"):
                soi.append((h, r))
            h += 1
            for dd in r["duoi"]:
                mau.Rows(2).Copy(ws.Rows(h))
                for c, j in COT_AI.items():
                    if c != "H" and s(dd["ai"].get(c)):
                        ghi_o(ws, h, j, dd["ai"][c])
                for j, v in dd["sau"].items():
                    ws.Cells(h, j).Value = v
                if dd["cao"]:
                    ws.Rows(h).RowHeight = dd["cao"]
                if dd.get("anh"):
                    f_anh = os.path.join(tmp, f"cu_{h}.{dd['anh'][1]}")
                    with open(f_anh, "wb") as fh:
                        fh.write(dd["anh"][0])
                    o_anh.append((h, f_anh))
                h += 1
        tong_hang = h - 1
        xl.DisplayAlerts = False
        mau.Delete()
        for hh, f_anh in o_anh:  # chèn ảnh sau cùng: đổi chiều cao dòng sau khi chèn sẽ làm méo ảnh
            dat_anh(ws, f_anh, hh)
        ghi_soi_anh(wb, soi, p, res, thu_muc_anh)
        ghi_links(wb, p)
        ws.Activate()
        ws.Range("A1").Select()
        wb.Saved = False
        wb.Close(SaveChanges=True)
        wb = None
        # kiểm lại bằng cách mở lần nữa, không tin lần lưu
        wb = xl.Workbooks.Open(f_tmp, UpdateLinks=0, ReadOnly=True)
        loi = kiem_anh(wb.Worksheets(SHEET), len(o_anh))
        wb.Close(SaveChanges=False)
        wb = None
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()
    if loi:
        shutil.rmtree(tmp, ignore_errors=True)
        sys.exit("LỖI: kiểm ảnh sau khi lưu:\n  " + "\n  ".join(loi[:20]))
    shutil.copy2(f_tmp, out)
    anh_ra = anh_trong_file(out)
    for k, sn in snap_moi.items():
        h = hang_cua[k]
        if sn.get("anh") is None:  # dòng có ảnh người thay đã giữ sẵn mốc cũ
            sn["anh"] = dau_anh(anh_ra[h][0]) if h in anh_ra else ""
    if not ra:
        snap_moi["_file"] = os.path.relpath(out, p.t)
        jdump(p.f("_dmvt_snapshot.json"), snap_moi)
        jdump(p.f("_dmvt_hang.json"), {"file": os.path.relpath(out, p.t).replace("\\", "/"), "hang": {k[1:]: v for k, v in hang_cua.items() if k.startswith("#")}})
    shutil.rmtree(tmp, ignore_errors=True)
    print(f"Đã ghi {out}: {len(dong)} dòng, {len(o_anh)} ảnh, {tong_hang - HANG_DAU + 1} hàng.")
    print("Giữ qua build lại:", dict(tk) or "chưa có bản trước")
    print("Cờ:", dict(Counter(c.split(":")[0] for r in dong for c in r["Cờ"])))
    print("Cờ ảnh:", dict(Counter(c.split(":")[0] for r in dong for c in r.get("Cờ ảnh", []))))
    if cfg.get("tam"):
        print(f"XEM: {cfg['tam']}")
    print("Trước khi gửi CĐT, tender tự xóa: các cột từ AG (Spec mới khác, Cờ, Link), sheet _LINKS, sheet _SOI_ANH; xem thêm L và AB:AC (giá).")
    return out


def kiem_anh(ws, can):
    """Mỗi ảnh: ô trên-trái = ô dưới-phải, ở cột H; số ảnh đúng. Comment cũng là shape nên chỉ đếm Type = 13 (ảnh)."""
    loi, n = [], 0
    for i in range(1, ws.Shapes.Count + 1):
        sh = ws.Shapes.Item(i)
        if sh.Type != 13:
            continue
        n += 1
        a, b = sh.TopLeftCell, sh.BottomRightCell
        if a.Address != b.Address or a.Column != COT_AI["H"]:
            loi.append(f"ảnh '{sh.Name}' nằm từ {a.Address} tới {b.Address}")
    if n != can:
        loi.append(f"có {n} ảnh, cần {can}")
    return loi


def ghi_soi_anh(wb, soi, p, res, thu_muc):
    """Chỉ các dòng có cờ ảnh: mã, ảnh thu nhỏ, link trang nguồn. Tender duyệt ngoại lệ; xóa sheet này trước khi gửi CĐT."""
    ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))  # tham số tên After của pywin32 không có tác dụng
    ws.Name = "_SOI_ANH"
    for j, (c, w) in enumerate([("Hàng DMVT", 10), ("Mã", 26), ("Cờ ảnh", 42), ("Ảnh máy chọn", 22), ("Số ứng viên", 10), ("Trang nguồn", 34)], 1):
        o = ws.Cells(1, j)
        o.Value = c
        o.Font.Bold = True
        o.Font.Color = 0xFFFFFF
        o.Interior.Color = 0x784E1F
        ws.Columns(j).ColumnWidth = w
    for i, (h, r) in enumerate(soi, 2):
        ws.Rows(i).RowHeight = 90
        them_link(ws, ws.Cells(i, 1), f"'{SHEET}'!A{h}", str(h))
        ws.Cells(i, 2).Value = an_toan_o(r["Khóa"])
        ws.Cells(i, 3).Value = an_toan_o("; ".join(r["Cờ ảnh"]))
        ws.Cells(i, 2).WrapText = ws.Cells(i, 3).WrapText = True
        ws.Cells(i, 5).Value = r.get("Số ứng viên") or 0
        if r.get("Ảnh"):
            f = os.path.join(thu_muc, r["Ảnh"])
            from PIL import Image
            c = ws.Cells(i, 4)
            with Image.open(f) as im:
                w0, h0 = im.size
            k = min((c.Width - 6) / w0, (c.Height - 6) / h0)
            sh = ws.Shapes.AddPicture(os.path.abspath(f), False, True, c.Left + 3, c.Top + 3, w0 * k, h0 * k)
            sh.Placement = 1
        ng = r.get("Nguồn") or []
        if ng:
            o = ws.Cells(i, 6)
            dc = o.Address.replace("$", "")
            t = ng[0]
            pg = r.get("Trang ảnh") or t["trang"]
            them_link(ws, o, f"'_SOI_ANH'!{dc}", f"{t['fid']} p{pg}")
            it = link_item(p, [f"{t['ma']} p{pg}", t["fid"], pg], res["FILE_INDEX"])
            p.links.append(("_SOI_ANH", dc, it[0], it[1], it[2], it[3]))


def ghi_links(wb, p):
    """Sheet _LINKS đúng dạng add-in SRC_LINKS đọc (xem td.write_links). Dòng cùng Key phải liền nhau: p.links ghi theo từng ô nên đã liền."""
    ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "_LINKS"
    hdr = ["Key", "Sheet", "Ô", "Nhãn", "File (tương đối)", "Trang", "Ô đích"]
    n = len(p.links) + 1
    du = [hdr] + [[f"{sh}!{c}", sh, c, lab, f, pg if pg else "", sub or ""] for sh, c, lab, f, pg, sub in p.links]
    ws.Range(f"A1:G{n}").Value = du
    ws.Rows(1).Font.Bold = True
    ws.Visible = 0  # xlSheetHidden
