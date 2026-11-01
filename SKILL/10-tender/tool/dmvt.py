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
    "tam": True,
}
TRAI_ANH = (800, 800)  # cạnh dài tối đa của ảnh đưa vào xlsx
HASH_LECH = 20  # plan mục 8: perceptual hash 16x16, lệch <= 20 bit là cùng ảnh


def nap_cau_hinh(p):
    c = dict(MAC_DINH)
    c.update(jload(p.f("_dmvt.json")) or {})
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


def boc_nguon(p, res, cfg):
    """Quét mọi trang spec, lấy bản mới nhất theo từng mã (không theo cả file). Trả {mã: bản ghi}."""
    import pymupdf
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
        if r["loai"] == "DR":
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
        if n.get("khop") and any(re.search(n["khop"], m) for m in cac_ma):
            return n
    return {"so": "", "ten": "CHƯA XẾP MỤC"}


def viet_dong(cac_ma, moi, res, cfg, ghi_chu_gop_k):
    """Dòng A:I bằng tiếng gốc của spec (LLM dịch thêm sau) + cờ + trạng thái."""
    r0 = moi.get(cac_ma[0])
    area = res["AREA"]
    d = {"Khóa": khoa_cua(cac_ma), "Mã": cac_ma, "Area": "", "Material": "", "Location": "", "Description": "", "NCC": "", "Remarks": "", "Cờ": [], "Omit": False}
    if not r0:  # có ở MAT, không có trang spec
        x = res["MAT"][cac_ma[0]]
        d.update(Material=s(x.get("Mô tả")), NCC=s(x.get("NCC/Mã SP")), Description="")
        d["Cờ"].append(s(x.get("Trạng thái cuối")) or "Không thấy spec")
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
                ph += sorted(phong_of(r.get("Mã phòng")))
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
    doc = mo(r["fid"])
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
        d["Nguồn"] = [{"ma": m, "fid": moi[m]["fid"], "trang": moi[m]["trang"], "ngay": moi[m]["ngay"], "loai": moi[m]["loai"]} for m in cac_ma if m in moi]
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
