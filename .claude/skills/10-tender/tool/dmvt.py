"""dmvt.py — danh mục vật tư (DMVT) của skill 10-tender. Gọi qua: td.py dmvt <GOC> ...

Máy bóc A:I từ trang spec có cấu trúc, cắt ảnh theo khối, áp Trạng thái cuối của MAT, dựng file bằng Excel COM.
LLM chỉ dịch, ghi chú, chọn ảnh ca khó (_data\\_dmvt_llm.json). Plan: DOCS\\PLAN_TENDER-CODEID-DMVT_2026-10-31.md mục 5.
"""
import hashlib
import io
import os
import re
import shutil
import zipfile
from collections import Counter, OrderedDict, defaultdict
from difflib import SequenceMatcher

import td
from td import P, merge, jload, jdump, s, nap_trang, file_bo, trang_trung, loai_trang, ma_dinh_nghia, ma_chuan, tach_ma, links_of, unzip, phong_of, tach_cham_phay, ma_cua_dong, TODAY, tcq_chain, split_list

# Không có _data\_dmvt.json thì dùng mặc định này. Cách đánh số 5x-yyy là của từng dự án nên sửa ở file đó, không viết cứng ở đây.
MAC_DINH = {
    "tieu_de": {"A1": "DANH MỤC VẬT TƯ – LIST OF MATERIAL", "A2": "", "A3": ""},
    "tien_to_ffe": [],  # rỗng = mọi tiền tố có trong sheet AREA
    "item_spec": r"-[23]\d\d$",  # mã có dòng DMVT lấy từ trang SPECIFICATION (mã vật liệu còn lại lấy từ MATERIAL REFERENCE)
    "bo_mo_ta": r"Installation|Fabrication",
    "gop_theo_nha_cc": r"Hardware|Curtain Track|Track",  # phụ kiện không có Design + Colour: gộp theo nhà cung cấp
    "nhom": [{"so": "I", "ten": "FF&E", "khop": r"^\d\d-"},
             {"so": "II", "ten": "FINISHES – VẬT LIỆU HOÀN THIỆN", "bo": "FINISHES"},
             {"so": "III", "ten": "SANITARYWARE – THIẾT BỊ VỆ SINH", "bo": "SANITARYWARE"},
             {"so": "IV", "ten": "IRONMONGERY", "bo": "IRONMONGERY"}],
    "tam": "Chưa có _data\\_dmvt.json: đang dùng mục nhóm mặc định, tiêu đề dòng 2 và 3 trống; tạo file đó để sửa",
}
TRAI_ANH = (800, 800)  # cạnh dài tối đa của ảnh đưa vào xlsx
HASH_LECH = 20  # plan mục 8: perceptual hash 16x16, lệch <= 20 bit là cùng ảnh


def nap_cau_hinh(p):
    c = dict(MAC_DINH)
    rieng = jload(p.f("_dmvt.json"))
    if rieng:  # đã có file của dự án: bỏ cảnh báo "chưa có file"
        c["tam"] = ""
        c.update(rieng)
    return c


# ---------------------------------------------------------------- đọc khối chữ
RX_NHAN = re.compile(r"^([A-Z][A-Za-z&/. -]{0,24}?)\s*:\s*(.*)$")
NHAN_LH = ("Tel", "Mobile", "Contact", "E-Mail", "Email", "Web", "Address", "Postcode", "Country", "Fax")
NHAN_BO = ("Date", "Revised", "Revision No", "Reason", "Quantity", "Unit", "Area", "Project No.", "Item", "Description", "Supplier") + NHAN_LH
TRONG = ("", "-", "–", "—", "N/A")


def nhan_gia_tri(dong):
    """Dòng 'Nhãn:' + giá trị cùng dòng hoặc các dòng sau -> [(nhãn, giá trị)] theo đúng thứ tự trang."""
    out = []
    for d in dong:
        d = d.strip()
        if not d:
            continue
        m = RX_NHAN.match(d)
        if m:
            out.append([m.group(1).strip(), m.group(2).strip()])
        elif out:
            out[-1][1] = (out[-1][1] + "\n" + d).strip()
    return out


def chuan_chu(t):
    return re.sub(r"[^a-z0-9%/]+", " ", s(t).lower()).strip()


def lien_he(kv):
    d = {}
    for k, v in kv:
        d.setdefault(k, v)
    ten = d.get("Supplier", "")
    out = [ten] if ten and ten.lower() != "to bid" else ([ten] if ten else [])
    for k, nhan in (("Contact", "Contact"), ("Tel", "Tel"), ("Mobile", "Mobile"), ("E-Mail", "E"), ("Email", "E")):
        v = d.get(k, "")
        if v and v not in TRONG:
            out.append(f"{nhan}: {v}")
    return "\n".join(out)


def thong_so(kv):
    return "\n".join(f"{k}: {v}" for k, v in kv if k not in NHAN_BO and s(v) not in TRONG)


def ma_khoi_mr(t):
    """Trang MATERIAL REFERENCE -> [(mã, [dòng])] theo thứ tự tiêu đề ITEM."""
    dong = t.split("\n")
    vt = [i for i, d in enumerate(dong) if d.strip() == "ITEM" and i + 1 < len(dong) and ma_chuan(dong[i + 1])]
    out = []
    for n, i in enumerate(vt):
        out.append((ma_chuan(dong[i + 1]), dong[i + 2:(vt[n + 1] if n + 1 < len(vt) else len(dong))]))
    return out


def doc_khoi_mr(ma, dong):
    """Tiêu đề khối là dòng không nhãn đầu tiên sau Date/Revised."""
    tieu_de, con = "", []
    for d in dong:
        d = d.strip()
        if not d:
            continue
        if not tieu_de and not con and re.match(r"^(Date|Revised)\b", d):
            continue
        if not tieu_de and not RX_NHAN.match(d):
            tieu_de = d
            continue
        con.append(d)
    return tieu_de, nhan_gia_tri(con)


def doc_spec(t):
    """Trang SPECIFICATION -> dict: mã, mô tả, kv thông số, chữ SPECIAL INSTRUCTIONS."""
    dong = [d.strip() for d in t.split("\n")]
    try:
        i = dong.index("Item:")
    except ValueError:
        return None
    ma = ma_chuan(dong[i + 1]) if i + 1 < len(dong) else None
    if not ma:
        return None
    si = next((j for j, d in enumerate(dong) if d.startswith("SPECIAL INSTRUCTIONS")), len(dong))
    kv = nhan_gia_tri(dong[i + 2:si])
    d = dict((k, v) for k, v in reversed(kv))
    return {"ma": ma, "mo_ta": d.get("Description", ""), "kv": kv, "si": "\n".join(dong[si + 1:]) if si < len(dong) else "", "ngay": d.get("Revised") or d.get("Date", "")}


# ---------------------------------------------------------------- trang DRAWING REFERENCE: theo cột
def cot_bang(trang, y):
    """Các đường dọc cắt qua hàng tiêu đề bảng, từ trái sang phải. Tiêu đề nằm lệch trong cột nên không suy được cột từ vị trí chữ."""
    xs = set()
    for g in trang.get_drawings():
        for it in g["items"]:
            if it[0] == "l" and abs(it[1].x - it[2].x) < 1 and min(it[1].y, it[2].y) <= y <= max(it[1].y, it[2].y):
                xs.add(round(it[1].x))
            elif it[0] == "re" and it[1].y0 <= y <= it[1].y1 and it[1].height > 10:
                xs |= {round(it[1].x0), round(it[1].x1)}
    out = []
    for x in sorted(xs):
        if not out or x - out[-1] > 3:
            out.append(x)
    return out


def doc_dr(trang):
    """Trang schedule 4 cột: chữ đọc theo đường kẻ bảng, lớp chữ PDF đảo thứ tự nên không đọc tuần tự được."""
    tu = trang.get_text("words")
    dau = [w for w in tu if w[4] == "Item" and any(x[4] == "Description" and abs(x[1] - w[1]) < 3 for x in tu)]
    if not dau:
        return None
    y0 = dau[0][3]
    tieu = {w[4]: w for w in tu if abs(w[3] - y0) < 3 and w[4] in ("Item", "Reference", "Finish", "Supplier")}
    if len(tieu) < 4:
        return None
    ke = cot_bang(trang, (dau[0][1] + dau[0][3]) / 2)
    cach = [next((x for x in reversed(ke) if x <= tieu[k][0]), None) for k in ("Reference", "Finish", "Supplier")]
    if None in cach or cach != sorted(set(cach)):
        return None
    y_sw = next((w[1] for w in tu if w[4] == "Sample" and w[1] > y0), None)
    y1 = next((w[1] for w in tu if w[4] == "Location:" and w[1] > y0), y_sw)
    if not y1:
        return None
    cot = {"Item": [], "Reference": [], "Finish": [], "Supplier": []}
    lim = [("Item", cach[0]), ("Reference", cach[1]), ("Finish", cach[2]), ("Supplier", 1e6)]
    for w in tu:
        if y0 + 1 < w[1] and w[3] < y1 + 1:
            cot[next(k for k, x in lim if (w[0] + w[2]) / 2 < x)].append(w)

    def thanh_dong(ws):
        ws = sorted(ws, key=lambda w: (round(w[1] / 4), w[0]))
        dong, cur, yc = [], [], None
        for w in ws:
            if yc is not None and abs(w[1] - yc) > 4:
                dong.append(" ".join(cur))
                cur = []
            cur.append(w[4])
            yc = w[1]
        if cur:
            dong.append(" ".join(cur))
        return [d for d in dong if d.strip()]

    t = trang.get_text()
    m = re.search(r"As shown (?:as|on) Designer.?s Drawings\.?", t)
    note = re.search(r"NOTE:\s*(.+?)(?:\n\s*Page \d|\Z)", t, re.S)
    return {"mo_ta": "\n".join(thanh_dong(cot["Item"])), "ref": thanh_dong(cot["Reference"]), "finish": thanh_dong(cot["Finish"]),
            "ncc": "\n".join(thanh_dong(cot["Supplier"])), "vi_tri": m.group(0).rstrip(".") if m else "",
            "note": re.sub(r"\s+", " ", note.group(1)).strip() if note else "", "y_swatch": y_sw}


# ---------------------------------------------------------------- ảnh
def phash(im):
    """Average hash 16x16, trả int 256 bit. Đúng phép đo mốc 47/57 của plan (DCT và dHash mất ổn định với vải trơn)."""
    import numpy as np
    from PIL import Image
    a = np.asarray(im.convert("L").resize((16, 16), Image.LANCZOS), dtype=float)
    return int("".join("1" if b else "0" for b in (a > a.mean()).flatten()), 2)


def lech(a, b):
    return bin(a ^ b).count("1")


def do_anh(im):
    """(tỉ lệ gần trắng, độ đa dạng màu): tách ảnh chụp với bản vẽ nét."""
    import numpy as np
    x = np.asarray(im.convert("RGB").resize((96, 96)), dtype=int)
    trang = float((x.min(axis=2) > 235).mean())
    da_dang = len({tuple(v) for v in (x // 24).reshape(-1, 3)})
    return trang, da_dang


def la_anh_chup(im):
    trang, da_dang = do_anh(im)
    return trang < 0.75 and da_dang >= 12


def lay_anh(doc, info):
    """Ảnh gốc theo xref; không có xref (ảnh nội tuyến) thì render vùng 150 dpi. Trả PIL.Image hoặc None."""
    import pymupdf
    from PIL import Image
    x = info.get("xref") or 0
    try:
        if x:
            px = pymupdf.Pixmap(doc, x)
            if px.alpha or px.colorspace is None or px.colorspace.n not in (1, 3):
                px = pymupdf.Pixmap(pymupdf.csRGB, px)
            return Image.open(io.BytesIO(px.tobytes("png"))).convert("RGB")
    except Exception:
        pass
    return None


def anh_trong_vung(doc, trang, y0, y1, nho=30):
    """Ảnh có tâm nằm giữa y0..y1, cạnh >= nho pt. Trả [{xref, bbox, img}]. Ảnh trùng xref trong vùng chỉ tính 1."""
    out, thay = [], set()
    for i in trang.get_image_info(xrefs=True):
        b = i["bbox"]
        yc = (b[1] + b[3]) / 2
        if not (y0 <= yc < y1) or min(b[2] - b[0], b[3] - b[1]) < nho or i.get("xref") in thay:
            continue
        thay.add(i.get("xref"))
        im = lay_anh(doc, i)
        if im is None:  # render vùng 150 dpi, không đoán vùng cắt ngoài khung ảnh
            from PIL import Image
            import pymupdf
            px = trang.get_pixmap(dpi=150, clip=pymupdf.Rect(b))
            im = Image.open(io.BytesIO(px.tobytes("png"))).convert("RGB")
        out.append({"xref": i.get("xref"), "bbox": [round(v) for v in b], "img": im})
    return out


def chon_mot(ung, cach="chup"):
    """Nhiều ảnh: cach='chup' ảnh chụp trước rồi lớn nhất (quyết định 12); cach='lon_nhat' lớn nhất. Trả chỉ số trong ung."""
    if not ung:
        return None
    return max(range(len(ung)), key=lambda i: (cach == "chup" and la_anh_chup(ung[i]["img"]), ung[i]["img"].width * ung[i]["img"].height))


def luu_anh(im, thu_muc, ten):
    """Thu nhỏ rồi lưu; ảnh ít màu dùng PNG cho nét, còn lại JPEG. Trả tên file."""
    from PIL import Image
    im = im.copy()
    im.thumbnail(TRAI_ANH, Image.LANCZOS)
    _, da = do_anh(im)
    duoi = "png" if da < 40 else "jpg"
    f = f"{ten}.{duoi}"
    os.makedirs(thu_muc, exist_ok=True)
    im.save(os.path.join(thu_muc, f), quality=85) if duoi == "jpg" else im.save(os.path.join(thu_muc, f))
    return f


# ---------------------------------------------------------------- bóc
def duong_dan_pdf(p, fidx, fid):
    r = fidx[fid]
    return unzip(p, fid, r) if s(r.get("Trong zip")) else os.path.join(p.inp, r["Đường dẫn"])


def ngay_file(fidx, fid):
    r = fidx.get(fid, {})
    return s(r.get("Ngày issue")) or s(r.get("Ngày đợt"))


def ten_file(fidx, fid):
    r = fidx.get(fid, {})
    return os.path.basename(s(r.get("Trong zip")) or s(r.get("Đường dẫn")))


def boc_nguon_spec_boc(p, res):
    """Spec kiểu 'Item Desc / Area' (mẫu AID): lấy từ _data\\_spec_boc.json do spec_boc.py bóc, chỉ các mã có ở MAT. Trả như boc_nguon."""
    import pymupdf
    fidx = res["FILE_INDEX"]
    fid_ten = {os.path.basename(s(r.get("Đường dẫn"))): k for k, r in fidx.items()}
    moi, tk = {}, Counter()
    con = defaultdict(list)
    for f in jload(p.f("_spec_boc.json")) or []:
        bo = "FINISHES" if "_FM_" in f["file"] else "IRONMONGERY" if "_IM_" in f["file"] else "SANITARYWARE" if "_SW_" in f["file"] else "FF&E"
        for x in f["ma"]:
            ma = x["ma"]
            if "." in ma:
                con[ma.split(".")[0]].append(x)
            elif ma in res["MAT"] and (ma not in moi or (not moi[ma].get("mo_ta") and x.get("item desc"))):
                moi[ma] = {"loai": "AID", "fid": fid_ten[f["file"]], "trang": x.get("trang_ct") or x["trang"], "trang_sc": x["trang"], "ngay": ngay_file(fidx, fid_ten[f["file"]]), "bo": bo, "mo_ta": s(x.get("item desc")), "kv": x}
    for ma, r in moi.items():
        r["con"] = con.get(ma, [])
    docs = {}

    def mo(fid):
        if fid not in docs:
            docs[fid] = pymupdf.open(duong_dan_pdf(p, fidx, fid))
        return docs[fid]

    return moi, defaultdict(list), tk, mo, docs


# ---------------------------------------------------------------- spec dạng bảng xlsx (Material Schedule, cột Minh họa)
COT_XLSX = {  # nhận cột theo chữ tiêu đề (đã bỏ dấu, chữ thường); thứ tự thử: khớp đầu tiên
    "ma": r"item code|ma vl\b|ma vat lieu|^ma sp\b|^code\b",
    "loai": r"item type|loai vat lieu",
    "ten": r"hang muc cong viec|ten hang muc|^item name",
    "mo_ta": r"^description|mo ta",
    "vi_tri": r"^location|vi tri",
    "anh": r"illustrat|minh hoa|hinh anh|^image",
    "ma_ncc": r"supplier.?s code|ma cua n",
    "ncc": r"^supplier(?!.?s code)|nha cung cap|^brand",
    "note": r"^note|chu thich|ghi chu",
}


def _bo_dau(t):
    import unicodedata
    t = unicodedata.normalize("NFD", s(t).lower().replace("đ", "d"))
    return re.sub(r"\s+", " ", "".join(c for c in t if unicodedata.category(c) != "Mn")).strip()


def anh_theo_o(f, sheet):
    """{(hàng, cột) 1-based: [(bytes, đuôi)]} của ảnh neo trong sheet, đọc thẳng từ zip (kể cả ảnh openpyxl không nạp)."""
    import xml.etree.ElementTree as ET
    NS = {"x": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    out = defaultdict(list)
    with zipfile.ZipFile(f) as z:
        nm = set(z.namelist())
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rel_wb = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        rid = next((x.get(R + "id") for x in wb.iter() if x.tag.endswith("}sheet") and x.get("name") == sheet), None)
        if not rid:
            return out
        p_sh = rel_wb[rid].lstrip("/")
        p_sh = p_sh if p_sh.startswith("xl/") else "xl/" + p_sh
        p_rel = os.path.dirname(p_sh) + "/_rels/" + os.path.basename(p_sh) + ".rels"
        if p_rel not in nm:
            return out
        dr = next((r.get("Target") for r in ET.fromstring(z.read(p_rel)) if r.get("Type").endswith("/drawing")), None)
        if not dr:
            return out
        p_dr = dr.lstrip("/") if dr.startswith("/") else os.path.normpath(os.path.join(os.path.dirname(p_sh), dr)).replace("\\", "/")
        p_drr = os.path.dirname(p_dr) + "/_rels/" + os.path.basename(p_dr) + ".rels"
        if p_drr not in nm:
            return out
        rel_dr = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read(p_drr))}
        for a in ET.fromstring(z.read(p_dr)):
            hang, cot = a.find("x:from/x:row", NS), a.find("x:from/x:col", NS)
            for blip in a.iter("{%s}blip" % NS["a"]):  # ảnh trong nhóm (grpSp) cũng neo theo ô của nhóm
                t = rel_dr.get(blip.get(R + "embed"), "")
                t = t.lstrip("/") if t.startswith("/") else os.path.normpath(os.path.join(os.path.dirname(p_dr), t)).replace("\\", "/")
                if hang is not None and t in nm:
                    out[(int(hang.text) + 1, int(cot.text) + 1)].append((z.read(t), t.rsplit(".", 1)[-1].lower()))
    return out


def boc_xlsx(path, fid, ngay, la_boq=False):
    """Bảng mã trong 1 file xlsx -> {mã: bản ghi loai XLSX}. Mỗi sheet: tìm hàng tiêu đề có cột mã + cột mô tả trong 15 hàng đầu."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(max_row=min(ws.max_row or 0, 3000), values_only=True)]
        hdr = None
        for h, r in enumerate(rows[:15]):
            cot = {}
            for j, v in enumerate(r):
                t = _bo_dau(v)
                if not t:
                    continue
                for k, rx in COT_XLSX.items():
                    if k not in cot and re.search(rx, t):
                        cot[k] = j
                        break
            if "ma" in cot and "mo_ta" in cot:
                hdr = (h, cot)
                break
        if not hdr:
            continue
        h, cot = hdr
        anh = anh_theo_o(path, ws.title)
        nhom = ""
        for i, r in enumerate(rows[h + 1:], h + 2):
            v = lambda k: s(r[cot[k]]) if k in cot and cot[k] < len(r) and r[cot[k]] is not None else ""
            ma_o = v("ma")
            if not ma_o:
                chu = [s(x) for x in r if s(x)]
                if len(chu) == 1 and len(chu[0]) < 60:
                    nhom = chu[0]  # dòng nhóm: PAINT/ SƠN
                continue
            cac_ma = [m.upper() for m in re.split(r"[\s;,/]+", ma_o) if re.search(r"\d", m)]
            if not cac_ma:
                continue
            ca = [x for (hh, cc), xs in anh.items() if hh == i and cc == cot.get("anh", -1) + 1 for x in xs]
            if not ca:  # ảnh neo lệch cột: lấy ảnh cùng hàng ở cột bất kỳ
                ca = [x for (hh, cc), xs in anh.items() if hh == i for x in xs]
            o = f"'{ws.title}'!{openpyxl.utils.get_column_letter(cot['ma'] + 1)}{i}"
            for m in cac_ma:
                if m in out:
                    continue
                out[m] = {"loai": "XLSX", "fid": fid, "trang": None, "o": o, "ngay": ngay, "bo": "FF&E" if la_boq else "FINISHES",
                          "mo_ta": v("mo_ta").lstrip("'"), "kv": {"loai": v("loai") or v("ten"), "vi_tri": v("vi_tri"), "ma_ncc": v("ma_ncc"), "ncc": v("ncc"),
                                                    "note": v("note"), "nhom": nhom, "sheet": ws.title, "hang": i},
                          "anh_xlsx": ca, "la_boq": la_boq}
    wb.close()
    return out


def them_nguon_xlsx(p, res, cfg, moi):
    """Spec dạng bảng xlsx bổ sung cho mã chưa có trang spec PDF. File: _dmvt.json 'spec_xlsx' (list FID) hoặc tự tìm mọi xlsx
    có cột mã + mô tả. BOQ CĐT xét sau cùng: chỉ cho mã hạng mục (cột MÃ SP + HÌNH ẢNH) chưa có ở file khác."""
    fidx = res["FILE_INDEX"]
    boq = s(p.cfg.get("boq")).replace("\\", "/")
    ds = cfg.get("spec_xlsx") or [k for k, r in fidx.items() if s(r.get("Loại")).lower() in ("xlsx", "xlsm") and not s(r.get("Trong zip"))]
    ds = [k for k in ds if k in fidx]
    ds = sorted(sorted(ds, key=lambda k: s(ngay_file(fidx, k)), reverse=True),  # file mới trước: mã có ở 2 file lấy bản mới
                key=lambda k: s(fidx[k].get("Đường dẫn")).replace("\\", "/") == boq)
    them = Counter()
    for fid in ds:
        r = fidx.get(fid)
        if not r:
            continue
        path = os.path.join(p.inp, s(r.get("Đường dẫn")))
        if not os.path.exists(path):
            continue
        la_boq = s(r.get("Đường dẫn")).replace("\\", "/") == boq
        try:
            bang = boc_xlsx(path, fid, ngay_file(fidx, fid), la_boq)
        except Exception as e:  # file hỏng/khóa: bỏ, không chặn DMVT
            print(f"XEM: không đọc được {fid} làm spec xlsx ({str(e)[:80]})")
            continue
        for m, x in bang.items():
            if m in moi or (m not in res["MAT"]):
                continue
            moi[m] = x
            them[fid] += 1
    if them:
        print("Spec xlsx bổ sung:", ", ".join(f"{k} {v} mã" for k, v in them.items()))


def boc_nguon(p, res, cfg):
    kq = _boc_nguon(p, res, cfg)
    if cfg.get("spec_xlsx") is not False:
        them_nguon_xlsx(p, res, cfg, kq[0])
    return kq


def _boc_nguon(p, res, cfg):
    """Quét mọi trang spec, lấy bản mới nhất theo từng mã (không theo cả file). Trả {mã: bản ghi}."""
    import pymupdf
    if cfg.get("nguon_spec") == "spec_boc":
        return boc_nguon_spec_boc(p, res)
    fidx = res["FILE_INDEX"]
    texts, _, _ = nap_trang(p, res)
    dup, tong = trang_trung(texts)
    bo = file_bo(res, texts) | set(tong)
    tien_to = set(cfg.get("tien_to_ffe") or res["AREA"])
    moi = {}
    cu = defaultdict(list)
    tk = Counter()

    def nhan(ma, rec):
        """Cùng ngày file thì giữ bản đầu tiên; bản bị thay ghi vào cu để biết spec đã đổi qua các đợt."""
        hien = moi.get(ma)
        # mã có khối MATERIAL REFERENCE thì khối đó là nguồn, kể cả khi có trang SPECIFICATION riêng hoặc mới hơn của mã
        if hien and hien["loai"] != rec["loai"] and "MR" in (hien["loai"], rec["loai"]):
            thay = rec["loai"] == "MR"
        else:
            thay = hien is None or rec["ngay"] > hien["ngay"]
        if thay:
            if hien:
                cu[ma].append(f"{hien['fid']} p{hien['trang']}")
            moi[ma] = rec
        else:
            cu[ma].append(f"{rec['fid']} p{rec['trang']}")

    for fid, pages in texts.items():
        if fid in bo or s(fidx[fid].get("Vai trò")) == "Làm rõ":
            continue
        cac_loai = [loai_trang(t) for t in pages]
        if not any(cac_loai):
            continue
        ngay = ngay_file(fidx, fid)
        for i, t in enumerate(pages, 1):
            lt = cac_loai[i - 1]
            if (fid, i) in dup or lt not in ("DR", "MR", "SPEC"):
                continue
            if lt == "DR":
                ms = ma_dinh_nghia(t, lt)
                if len(ms) != 1:
                    tk["DR không đọc được mã"] += 1
                    continue
                bo_ho_so = next((k for k in ("FINISHES", "SANITARYWARE", "IRONMONGERY") if f"{k} SCHEDULE" in t), "KHÁC")
                nhan(next(iter(ms)), {"loai": "DR", "fid": fid, "trang": i, "ngay": ngay, "bo": bo_ho_so})
            elif lt == "MR":
                for ma, dong in ma_khoi_mr(t):
                    if ma[:2] in tien_to:
                        tieu_de, kv = doc_khoi_mr(ma, dong)
                        nhan(ma, {"loai": "MR", "fid": fid, "trang": i, "ngay": ngay, "bo": "FF&E", "tieu_de": tieu_de, "kv": kv})
            else:
                sp = doc_spec(t)
                if sp and sp["ma"][:2] in tien_to:
                    anh = []  # trang minh họa ngay sau trang SPEC của mã, tới mã kế tiếp
                    for j in range(i, len(pages)):
                        if cac_loai[j] == "SPEC" and doc_spec(pages[j]):
                            break
                        if cac_loai[j] == "ANH":
                            anh.append(j + 1)
                    nhan(sp["ma"], dict(sp, loai="SPEC", fid=fid, trang=i, ngay=ngay, bo="FF&E", trang_anh=anh))
    # mở PDF đọc cột và ảnh, chỉ cho bản được chọn
    docs = {}

    def mo(fid):
        if fid not in docs:
            docs[fid] = pymupdf.open(duong_dan_pdf(p, fidx, fid))
        return docs[fid]

    for ma, r in moi.items():
        if r["loai"] == "DR":
            r.update(doc_dr(mo(r["fid"])[r["trang"] - 1]) or {"mo_ta": "", "ref": [], "finish": [], "ncc": "", "vi_tri": "", "note": "", "loi_cot": True})
    return moi, cu, tk, mo, docs


CHU_CHUA_RO = {"refer to supplier", "refer to design intent", "tbc", "to be confirmed", "as per design intent", "by supplier"}


def khoa_cua(cac_ma):
    return " / ".join(cac_ma)


def sap_xep_ma(ma):
    return (ma[:2].isdigit(), ma)


def dong_dmvt(p, res, cfg):
    """Gộp bản ghi nguồn thành dòng DMVT: FF&E theo Design + Colour, phụ kiện theo nhà cung cấp, mã khác 1 dòng/mã."""
    moi, cu, tk, mo, docs = boc_nguon(p, res, cfg)
    fidx = res["FILE_INDEX"]
    rx_item, rx_bo = re.compile(cfg["item_spec"]), re.compile(cfg["bo_mo_ta"], re.I)
    rx_ncc = re.compile(cfg["gop_theo_nha_cc"], re.I)
    co_khoi_mr = {m for m, r in moi.items() if r["loai"] == "MR"}
    nhom = OrderedDict()
    ghi_chu_gop = defaultdict(list)
    for ma in sorted(moi, key=sap_xep_ma):
        r = moi[ma]
        if r["loai"] in ("DR", "AID", "XLSX"):
            nhom[ma] = [ma]
        elif r["loai"] == "MR":
            d = dict(r["kv"])
            ds, cl = chuan_chu(d.get("Design No")), chuan_chu(d.get("Colour"))
            khop = None
            for k in nhom:
                r0 = moi[nhom[k][0]]
                if r0["loai"] != "MR":
                    continue
                d0 = dict(r0["kv"])
                ds0, cl0 = chuan_chu(d0.get("Design No")), chuan_chu(d0.get("Colour"))
                if not cl or cl != cl0 or cl in CHU_CHUA_RO:  # 'Refer to Supplier' không phải mã màu: không gộp được
                    continue
                if ds == ds0:
                    khop = k
                    break
                if SequenceMatcher(None, ds, ds0).ratio() >= 0.8:  # cùng mã màu, tên mẫu khác vài chữ: gộp, có cờ
                    khop = k
                    ghi_chu_gop[k].append(f"Gộp theo mã màu, khác chữ: {d0.get('Design No')} / {d.get('Design No')}")
                    break
            if khop:
                nhom[khop].append(ma)
            else:
                nhom[f"mr:{ds}|{cl}|{ma}"] = [ma]
        elif r["loai"] == "SPEC" and rx_item.search(ma) and not rx_bo.search(s(r.get("mo_ta"))) and ma not in co_khoi_mr:
            d = dict(r["kv"])
            mt = s(r.get("mo_ta"))
            if rx_ncc.search(mt):
                k = "ncc:" + chuan_chu(d.get("Supplier"))
            else:
                k = "spec:" + chuan_chu(d.get("Ref") or d.get("Ref No") or mt) + "|" + chuan_chu(d.get("Content"))
            nhom.setdefault(k, []).append(ma)
    # mã có ở MAT mà không có trang spec: vẫn có dòng, có cờ
    for m in res["MAT"]:
        if m not in moi and not m[:2].isdigit():
            nhom[m] = [m]
    return nhom, moi, cu, tk, mo, docs, ghi_chu_gop


def so_nhom(cfg, cac_ma, moi):
    for n in cfg["nhom"]:
        if n.get("bo"):
            # mã có ở MAT mà không có trang spec xếp vào Finishes, vì Sanitary, IR đều đủ trang
            if any(moi.get(m, {}).get("bo", "FINISHES" if not m[:2].isdigit() else "") == n["bo"] for m in cac_ma):
                return n
            continue
        # mục FF&E: mã dạng 51-500 hoặc mã hạng mục lấy từ BOQ CĐT (spec xlsx, bo FF&E: LF-02, BU-01)
        if n.get("khop") and any(re.search(n["khop"], m) or (moi.get(m, {}).get("loai") == "XLSX" and moi[m].get("bo") == "FF&E") for m in cac_ma):
            return n
    return {"so": "", "ten": "CHƯA XẾP MỤC"}


# Trình bày chữ spec kiểu AID: spec_boc.py gộp hết về 1 dòng. Chỉ thêm xuống dòng và bỏ chữ rác dính từ layout PDF, không đổi nội dung.
_NHAN_HD = ["Product Code", "Colour/ Finish", "Pull Handle-EXT", "I-Shaped Grab Bar-INT", "Grab Bar-INT", "Lever Handle", "Pull Handle", "Lockset", "Wall Stopper", "Concealed Hinge Depth", "Concealed Hinge", "Flush Bolt", "Dimension", "Notes", "Note", "Location", "Handle", "Length"]
_RX_NHAN_HD = re.compile(r"(?<![\w(/-])(?<!Pull )(?<!Lever )(" + "|".join(re.escape(x) for x in sorted(_NHAN_HD, key=len, reverse=True)) + r")\s*:\s*")
_RX_CUA = re.compile(r"\s*(?:;\s*)?(-\s*\(D\d|\(D\d[^)]*\)\s+(?=[A-Z]))")
_RX_NHAN_FFE = re.compile(r"(?<!\()\s+\b(Base|Mattress|Frame material selection|Colour|Size)\s*:\s*")
_RX_KT = re.compile(r"\s+(?=(?:ø\s*\d|[LWHD]\s?\d{2,5}\s?mm\b))")
_RX_MA_VL = re.compile(r"(?<!with)(?<!match)(?<!to)\s+(?=[A-Z]{2}-\d{2}\b)")


def chu_hd(t):
    """Brand: X Product Code: Y Colour/ Finish: ... -> mỗi nhãn, mỗi mục cửa (D..) 1 dòng."""
    t = re.sub(r"\s*\bW(Location\s*:)", r" W\n\1", t)  # chữ W thừa có sẵn trong PDF: giữ, chỉ tách
    t = re.sub(r"\s+-\s*(Location\s*:)", r" \1", t)
    t = _RX_NHAN_HD.sub(lambda m: "\n" + m.group(1) + ": ", t)
    t = _RX_CUA.sub(lambda m: "\n" + m.group(1) + (" " if m.group(1).endswith(")") else ""), t)
    t = re.sub(r"\s+-(?=(?:For|Max)\b)", "\n-", t)
    t = re.sub(r"\s+-\s+(?=L\d\b)", "\n- ", t)
    return re.sub(r"\s*\n\s*", "\n", t).strip()


def chu_material(t):
    """Tên và kích thước tách dòng. Không đoán ranh giới giữa loại và tên."""
    return _RX_NHAN_FFE.sub(lambda m: "\n" + m.group(1) + ": ", _RX_KT.sub("\n", t, count=1))


def bo_ma_rac(v, ma):
    """'NA MT-01', 'WXE5987 for Background only WC-01': mã của chính dòng dính từ cột bên cạnh. Bỏ mã; còn 'NA...' ngắn hoặc rỗng thì coi như không có."""
    v = re.sub(rf"(?:^|\s+){re.escape(ma)}$|^{re.escape(ma)}\s+|\s+{re.escape(ma)}\s+(?=[A-Z]{{3,}})", " ", v).strip()
    return "" if v in ("", "-") or re.fullmatch(r"(?:NA|N/A)(?:\s.{0,40})?", v) else v


def chu_ncc(v):
    """'AGC Flat Glass Supplier': bỏ nhãn cột dính cuối. 'Grace Lee 7030 Ang Mo Kio': tách địa chỉ."""
    v = re.sub(r"\s+Supplier$", "", v.strip())
    return "" if v in ("Supplier", "T:") else re.sub(r"\s+(?=\d{1,5}[A-Z]?\s+[A-Z][a-z])", "\n", v)


def viet_dong(cac_ma, moi, res, cfg, ghi_chu_gop_k):
    """Dòng A:I bằng tiếng gốc của spec (LLM dịch thêm sau) + cờ + trạng thái."""
    r0 = moi.get(cac_ma[0])
    area = res["AREA"]
    d = {"Khóa": khoa_cua(cac_ma), "Mã": cac_ma, "Area": "", "Material": "", "Location": "", "Description": "", "NCC": "", "Remarks": "", "Cờ": [], "Omit": False}
    if not r0:  # có ở MAT, không có trang spec
        x = res["MAT"][cac_ma[0]]
        d.update(Material=s(x.get("Mô tả")), NCC=s(x.get("NCC/Mã SP")), Description="")
        d["Cờ"].append(s(x.get("Trạng thái cuối")) or "Không thấy spec")
    elif r0["loai"] == "AID":
        kv = r0["kv"]
        ma0 = cac_ma[0]
        tt = [(n, bo_ma_rac(s(kv.get(k)), ma0)) for n, k in (("Dimensions", "dimensions"), ("Finishes", "finishes"), ("Material", "material"), ("Content", "content"), ("Prod. Code", "prod. code"), ("Colour", "colour"), ("Brand", "brand"))]
        chinh = {"Dimensions": lambda v: _RX_NHAN_FFE.sub(lambda m: "\n" + m.group(1) + ": ", v), "Finishes": lambda v: _RX_MA_VL.sub("\n", v), "Brand": chu_hd}
        mo_ta = [f"{n}: {chinh.get(n, str)(v)}" for n, v in tt if v and v not in TRONG]
        if r0.get("con"):
            mo_ta.append("Kèm:\n" + "\n".join(dict.fromkeys(f"- {x['ma']} {s(x.get('item desc'))}"[:90] for x in r0["con"])))  # 1 mã con có cả ở trang schedule và trang chi tiết
        ncc = [y for y in (chu_ncc(s(kv.get(k))) for k in ("manufacturer", "supplier", "contact", "address")) if y]
        d.update(Material=chu_material(s(r0.get("mo_ta"))), Location=s(kv.get("area") or kv.get("location")), Description="\n".join(mo_ta), NCC="\n".join(ncc), Remarks=bo_ma_rac(s(kv.get("remarks")), ma0))
    elif r0["loai"] == "XLSX":  # bảng xlsx: Material Schedule (loại + mô tả + mã NCC) hoặc BOQ CĐT (mô tả hạng mục)
        kv = r0["kv"]
        mo_ta = [s(r0.get("mo_ta"))] + ([f"Mã NCC: {kv['ma_ncc']}"] if kv.get("ma_ncc") else [])
        d.update(Material=kv.get("loai") or (s(r0.get("mo_ta")).split(chr(10))[0] if r0.get("la_boq") else kv.get("nhom", "")),
                 Location=kv.get("vi_tri", ""), Description=chr(10).join(x for x in mo_ta if x), NCC=kv.get("ncc", ""), Remarks=kv.get("note", ""))
        if r0.get("la_boq"):
            d["Cờ"].append("Lấy từ BOQ CĐT (không có spec riêng)")
    elif r0["loai"] == "DR":
        ref = list(r0.get("ref") or [])
        fin = list(r0.get("finish") or [])
        if fin and not re.match(r"^\s*Finish", fin[0], re.I):
            fin[0] = "Finish: " + fin[0]
        d.update(Material=s(r0.get("mo_ta")), Location=s(r0.get("vi_tri")), Description="\n".join(ref + fin), NCC=s(r0.get("ncc")), Remarks=s(r0.get("note")))
        if r0.get("loi_cot"):
            d["Cờ"].append("Không đọc được cột trang schedule")
    elif r0["loai"] == "MR":
        kv = dict(r0["kv"])
        ten = kv.get("Design No") if s(kv.get("Design No")) not in TRONG else kv.get("Colour", "")
        d.update(Material=f"{r0.get('tieu_de', '')} – {ten}".strip(" –"), Description=thong_so(r0["kv"]), NCC=lien_he(r0["kv"]),
                 Location="\n".join(f"- {moi[m].get('tieu_de', '')} – {m}" for m in cac_ma))
    else:  # SPEC
        mts = list(dict.fromkeys(s(moi[m].get("mo_ta")) for m in cac_ma))
        d.update(Material=mts[0], Description=thong_so(r0["kv"]), NCC=lien_he(r0["kv"]),
                 Location="\n".join(f"- {moi[m].get('mo_ta', '')} – {m}" for m in cac_ma))
    # Area: tiền tố mã -> tập phòng (luôn cờ); mã không có tiền tố khu -> Mã phòng dòng BOQ (luôn cờ); không có thì trống
    ph = []
    if any(m[:2].isdigit() for m in cac_ma):
        for m in cac_ma:
            g = area.get(m[:2])
            if g:
                ph += tach_cham_phay(g.get("Mã phòng"))
        d["Cờ"].append("Area theo tiền tố")
    else:
        for r in res["SCOPE"].values():
            if set(ma_cua_dong(r, area)) & set(cac_ma):
                ph += sorted(phong_of(r.get("Mã phòng"))) or ([s(r.get("Nhóm"))] if cfg.get("nguon_spec") == "spec_boc" and s(r.get("Nhóm")) else [])
        if ph:
            d["Cờ"].append("Area theo dòng BOQ")
    d["Area"] = "\n".join(dict.fromkeys(ph))
    for g in ghi_chu_gop_k:
        d["Cờ"].append(g)
    # Trạng thái: chỉ đọc MAT; FF&E chưa có ở MAT thì có cờ, không suy. Spec ghi thẳng OMITTED thì giữ và ghi đúng chữ đó.
    tcq = res["TCQ_LOG"]
    ch = tcq_chain(tcq)
    ghi = []
    for m in cac_ma:
        x = res["MAT"].get(m)
        if x:
            st = s(x.get("Trạng thái cuối"))
            if st == "Omit":
                d["Omit"] = True
                ghi.append(f"OMITTED: {s(x.get('Nội dung cuối (VN)'))}".strip())
                d["Cờ"].append("Omit")
            elif st in ("Đổi theo TCQ", "VE", "Chờ", "Không thấy spec"):
                ghi.append(f"[{st}] {s(x.get('Nội dung cuối (VN)'))}".strip())
                d["Cờ"].append(st)
        else:
            d["Cờ"].append("Chưa qua MAT")
            nhac = [k for k, r in tcq.items() if ch.get(k) == [k] and m in [ma_chuan(v) for v in split_list(r.get("Mã VL"))]]
            if nhac:
                ghi.append(f"[TCQ nhắc {m}, chưa kết luận] {'; '.join(nhac[:6])}")
                d["Cờ"].append("TCQ nhắc mã")
        if moi.get(m) and re.search(r"OMITTED", s(moi[m].get("tieu_de")) + s(moi[m].get("mo_ta")), re.I):
            d["Omit"] = True
            ghi.append(f"OMITTED theo spec: {m}")
            d["Cờ"].append("Omit")
    if ghi:
        d["Remarks"] = "\n".join(x for x in [d["Remarks"]] + list(dict.fromkeys(ghi)) if x)
    d["Cờ"] = list(dict.fromkeys(d["Cờ"]))
    return d


def ung_vien_anh(ma, r, mo):
    """Ứng viên ảnh theo khối của mã. Trả (ứng viên, cách lấy)."""
    if not r:
        return [], ""
    if r["loai"] == "XLSX":  # ảnh nhúng ở cột Minh họa cùng hàng; WMF/EMF PIL không mở được thì bỏ
        from PIL import Image
        out = []
        for b, duoi in r.get("anh_xlsx", []):
            try:
                out.append({"xref": None, "img": Image.open(io.BytesIO(b)).convert("RGB"), "trang": None})
            except Exception:
                continue
        return out, "ô xlsx"
    doc = mo(r["fid"])
    if r["loai"] == "AID":  # ảnh trên trang chi tiết của mã; thiếu thì dùng khối của mã trên trang schedule
        out = []
        kv = r["kv"]
        if not kv.get("trang_ct") and kv.get("y0") is not None:  # không có trang chi tiết riêng: chỉ lấy ảnh trong dải của mã trên trang schedule
            trang = doc[r["trang_sc"] - 1]
            # ảnh trên schedule thường là nhiều mảnh nhỏ ghép lại (phụ kiện): cắt nguyên vùng cột ảnh của mã, có mảnh ảnh mới cắt
            r_cot = (trang.rect.width * 0.68, kv["y0"], trang.rect.width * 0.97, kv["y1"])
            co_anh = any(r_cot[0] <= (i["bbox"][0] + i["bbox"][2]) / 2 <= r_cot[2] and r_cot[1] <= (i["bbox"][1] + i["bbox"][3]) / 2 < r_cot[3] for i in trang.get_image_info())
            if co_anh:
                from PIL import Image
                import pymupdf
                px = trang.get_pixmap(dpi=150, clip=pymupdf.Rect(r_cot))
                out.append({"xref": None, "bbox": [round(v) for v in r_cot], "img": Image.open(io.BytesIO(px.tobytes("png"))).convert("RGB"), "trang": r["trang_sc"]})
            return out, "khối"
        for t_so in dict.fromkeys((r["trang"], r["trang_sc"])):
            trang = doc[t_so - 1]
            for u in anh_trong_vung(doc, trang, 0, trang.rect.height, nho=60):
                u["trang"] = t_so
                out.append(u)
            if out and t_so == r["trang"] and r["trang"] != r["trang_sc"]:
                break
        return out, "trang chi tiết"
    if r["loai"] == "MR":
        trang = doc[r["trang"] - 1]
        ms = [m for m, _ in ma_khoi_mr(trang.get_text())]
        ys = [w[1] for w in trang.get_text("words") if w[4] == "ITEM" and w[0] < 100]
        if len(ys) != len(ms) or ma not in ms:
            return [], "khối"
        i = ms.index(ma)
        return anh_trong_vung(doc, trang, ys[i], ys[i + 1] if i + 1 < len(ys) else trang.rect.height), "khối"
    if r["loai"] == "DR":
        trang = doc[r["trang"] - 1]
        y_ht = [w[1] for w in trang.get_text("words") if w[4] == "Revision"]
        return anh_trong_vung(doc, trang, r.get("y_swatch", 0), min(y_ht) if y_ht else trang.rect.height), "swatch"
    out = []
    for t in r.get("trang_anh", [])[:6]:
        trang = doc[t - 1]
        for u in anh_trong_vung(doc, trang, 0, trang.rect.height, nho=60):
            u["trang"] = t
            out.append(u)
    return out, "minh họa"


def cmd_boc(p, res, cfg, nap_lai_anh=True):
    """Bóc dòng + cắt ảnh -> _data\\_dmvt_boc.json, _data\\_dmvt_anh\\."""
    nhom, moi, cu, tk, mo, docs, gop = dong_dmvt(p, res, cfg)
    thu_muc = p.f("_dmvt_anh")
    if os.path.isdir(thu_muc):
        shutil.rmtree(thu_muc)
    os.makedirs(thu_muc)
    llm = jload(p.f("_dmvt_llm.json")) or {}
    rows, hash_cua = [], {}
    for k, cac_ma in nhom.items():
        d = viet_dong(cac_ma, moi, res, cfg, gop.get(k, []))
        n = so_nhom(cfg, cac_ma, moi)
        d["Mục"], d["Tên mục"] = n["so"], n["ten"]
        r0 = moi.get(cac_ma[0])
        d["Nguồn"] = [{"ma": m, "fid": moi[m]["fid"], "trang": moi[m]["trang"], "ngay": moi[m]["ngay"], "loai": moi[m]["loai"], **({"o": moi[m]["o"]} if moi[m].get("o") else {})}
                      for m in cac_ma if m in moi]
        cu_ma = [x for m in cac_ma for x in cu.get(m, [])]
        if cu_ma:
            d["Bản cũ"] = cu_ma[:6]
        ung, cach = ung_vien_anh(cac_ma[0], r0, mo)
        d["Hash ứng viên"] = [f"{phash(u['img']):064x}" for u in ung]
        if not ung and len(cac_ma) > 1:  # mã đầu không có ảnh, thử các mã còn lại cùng Design + Colour
            for m in cac_ma[1:]:
                ung, cach = ung_vien_anh(m, moi.get(m), mo)
                if ung:
                    break
        chon = chon_mot(ung, cfg.get("chon_anh", "chup"))
        ten_llm = (llm.get(d["Khóa"]) or {}).get("Ảnh chọn")
        co_a = []
        if not ung:
            co_a.append("Không có ảnh")
        else:
            if isinstance(ten_llm, int) and 0 <= ten_llm < len(ung):
                chon = ten_llm
                co_a.append("LLM chọn ảnh")
            if cach == "minh họa":
                co_a.append("Ảnh ca khó")
            elif len(ung) > 1:
                co_a.append("Nhiều ảnh, máy chọn")
            im = ung[chon]["img"]
            trang_ti, da = do_anh(im)
            if trang_ti > 0.97:
                co_a.append("Ảnh gần như trắng")
            if min(im.size) < 80:
                co_a.append("Ảnh độ phân giải thấp")
            ten = hashlib.md5(d["Khóa"].encode("utf-8")).hexdigest()[:12]
            d["Ảnh"] = luu_anh(im, thu_muc, ten)
            d["Hash ảnh"] = f"{phash(im):064x}"
            hash_cua[d["Khóa"]] = phash(im)
            d["Số ứng viên"] = len(ung)
            d["Trang ảnh"] = ung[chon].get("trang") or (r0 or {}).get("trang")
        d["Cờ ảnh"] = co_a
        rows.append(d)
    # 2 mã khác Design/Colour cùng ảnh; (cùng Design/Colour đã gộp 1 dòng nên không cần so ở đây)
    ks = list(hash_cua)
    dau_mo_ta = {x["Khóa"]: chuan_chu(s(x["Description"]).split("\n")[0]) for x in rows}
    for i, a in enumerate(ks):
        for b in ks[i + 1:]:
            if hash_cua[a] == hash_cua[b] and dau_mo_ta[a] != dau_mo_ta[b]:
                for x in rows:
                    if x["Khóa"] in (a, b):
                        x["Cờ ảnh"].append("Cùng ảnh với dòng khác: " + (b if x["Khóa"] == a else a))
    for dc in docs.values():
        dc.close()
    jdump(p.f("_dmvt_boc.json"), {"rows": rows, "dem": tk})
    return rows


def main(goc, args):
    p = P(goc).need()
    res = merge(p)[0]
    cfg = nap_cau_hinh(p)
    if "--chi-boc" in args:
        rows = cmd_boc(p, res, cfg)
        print(f"Đã bóc {len(rows)} dòng -> {p.f('_dmvt_boc.json')}")
        print("Theo mục:", dict(Counter(r["Mục"] + " " + r["Tên mục"] for r in rows)))
        print("Cờ ảnh:", dict(Counter(c.split(":")[0] for r in rows for c in r["Cờ ảnh"])))
        return rows
    import dmvt_xlsx
    return dmvt_xlsx.build(p, res, cfg, args)
