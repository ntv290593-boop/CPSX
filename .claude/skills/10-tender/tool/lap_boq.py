"""Lập BOQ khi dự án chưa có BOQ CĐT (mẫu tu_lap_vn): ghép phạm vi được tick + tag bản vẽ (DWG) + spec + lịch cửa vào form của user.

Chỉ ghi file BOQ mới, không đụng hồ sơ gốc. Phần riêng từng dự án nằm ở <GOC>\\10_TENDER\\_data\\_lap_boq.json.
Đầu vào: _data\\_dwg_tag.json (SKILL\\_core\\dwg_tag.py), _data\\_spec_boc.json (spec_boc.py).
  py -I lap_boq.py <GOC>
Mỗi dòng ghi nguồn tờ và cờ ở cột Ghi chú. Không suy số: thiếu thì để trống và ghi cờ.
"""
import collections
import datetime
import json
import os
import re
import shutil
import sys
import warnings

warnings.filterwarnings("ignore")
import openpyxl
import pymupdf as fitz
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # chạy với py -I
from be_mat import doan as doan_be_mat, nap_soat

PA = ["ITEM", "=", "MTRL-CODE", "CLNG-CODE"]
NA = ["NUM", "--", "MTRL-NMBR", "CLNG-NMBR"]
RX_CHU = re.compile(r"[A-Z]{1,3}")
RX_SO = re.compile(r"[0-9O]{1,4}[A-Z]?")
BLOCK_TRAN = {"CEILING", "ceiling", "CODE-CLNG.1", "CODE-CLNG.2"}


def chuan_so(n):
    return n.replace("O", "0") if re.fullmatch(r"[0-9O]+", n) else n


def ma_tag(t):
    """Mã từ 1 tag. Trả 'WD-04' hoặc None. Multileader có thể đảo thứ tự, có nhãn đi kèm, gõ chữ O thay số 0."""
    if t.get("ml"):
        v = [x.strip("' ") for x in t["ml"]]
        for i in range(len(v) - 1):
            for a, b in ((v[i], v[i + 1]), (v[i + 1], v[i])):
                if RX_CHU.fullmatch(a) and RX_SO.fullmatch(b) and re.search(r"\d|O", b):
                    return f"{a}-{chuan_so(b)}"
        m = re.fullmatch(r"([A-Z]{2,3})[- ]?([0-9O]{2,4}[A-Z]?)", " ".join(v))
        return f"{m[1]}-{chuan_so(m[2])}" if m else None
    a = t["a"]
    for p, n in zip(PA, NA):
        if RX_CHU.fullmatch(a.get(p, "")) and RX_SO.fullmatch(a.get(n, "")):
            return f"{a[p]}-{chuan_so(a[n])}"
    return None


def ma_cua(t):
    a = t["a"]
    v = a.get("DOOROBJECTS:NUMBER") or a.get("DXX")
    if v and re.fullmatch(r"D\d+[A-Za-z]?", v):
        m = re.fullmatch(r"(D\d+)([A-Za-z]?)", v)
        return m[1] + m[2].lower()
    return None


def so_to(ten):
    m = re.search(r"(\d{4}(?:\.\d)?)$", ten)
    return m.group(1) if m else ten


def doc_drawing_list(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for r in wb[sheet].iter_rows(values_only=True):
        v = [str(x).strip() for x in r if x is not None]
        if len(v) >= 8 and v[0].startswith("AID250121"):
            out[v[4]] = v
    return out


def loai_to(so, dl):
    """Loại tờ từ Drawing List; không có thì theo số tờ."""
    v = dl.get(so) or dl.get(so + ".1")
    ten = " ".join(v[5:8]).upper() if v else ""
    if so[0] == "5":
        return "joinery"
    if "RENDER" in ten:
        return "render"
    if "REFLECTED" in ten or (not v and re.fullmatch(r"1\d\d3", so)):
        return "tran"
    if "FURNITURE" in ten or (not v and re.fullmatch(r"1\d\d1", so)):
        return "ffe"
    if "PARTITION" in ten or (not v and re.fullmatch(r"1\d\d2", so)):
        return "tuong_mb"
    if "FLOOR" in ten or (not v and re.fullmatch(r"1\d\d4", so)):
        return "san"
    if "ELECTRICAL" in ten or (not v and re.fullmatch(r"1\d\d5", so)):
        return "dien"
    if so[0] == "2":
        return "mat_dung"
    if so[0] == "3":
        return "wc"
    return "khac"


def dong_chu(pg):
    out = []
    for b in pg.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            x = " ".join(s["text"] for s in l["spans"]).strip()
            if x:
                out.append((x, l["bbox"], (round(l["dir"][0]), round(l["dir"][1])), l["spans"][0]["size"]))
    return out


def huong_ngang(pg):
    """Hướng chữ nằm ngang khi xem trang: trang xoay 90/270 thì hệ tọa độ thô đổi trục."""
    return (0, 1) if pg.rotation in (90, 270) else (1, 0)


def kich_thuoc_cua(pg):
    """{mã cửa: (rộng, cao)} trên 1 tờ lịch cửa: chia ô theo vị trí chữ mã cửa cỡ lớn, lấy số cỡ thường lớn nhất mỗi hướng trong ô."""
    ds = dong_chu(pg)
    ma = [(x, b) for x, b, d, sz in ds if sz > 30 and re.fullmatch(r"D\d+[a-zA-Z]?", x)]
    ngang = huong_ngang(pg)
    out = {}
    for x, b in ma:
        r, c = [], []
        for t2, q, d, sz in ds:
            if not (re.fullmatch(r"\d{3,5}", t2) and sz > 8):
                continue
            if ngang == (0, 1):
                trong = b[0] <= q[0] <= b[0] + 760 and b[1] - 40 <= q[1] <= b[1] + 660
            else:
                trong = b[1] - 760 <= q[1] <= b[1] and b[0] - 40 <= q[0] <= b[0] + 660
            if trong and d == ngang:
                r.append(int(t2))
            elif trong:
                c.append(int(t2))
        if r and c:
            out.setdefault(x, (max(r), max(c)))
    return out


def kich_thuoc_joinery(pdf):
    """(ngang lớn nhất, đứng lớn nhất) đọc từ số kích thước trên tờ joinery; chỉ là kích thước bao, luôn có ghi chú."""
    pg = fitz.open(pdf)[0]
    ngang = huong_ngang(pg)
    r, c = [], []
    for x, b, d, sz in dong_chu(pg):
        if re.fullmatch(r"\d{3,5}", x) and 100 <= int(x) <= 30000 and sz > 6:
            (r if d == ngang else c).append(int(x))
    return (max(r) if r else None, max(c) if c else None)


RX_KT = re.compile(r"([LWDHø])\s*(\d{2,5})\s*mm", re.I)


def kich_thuoc_do_roi(*chu):
    """Từ 'W750mm x D750mm x H780mm' -> (dài, rộng, cao). L/W là dài, D/W là rộng; ø cho cả dài và rộng."""
    gt = {}
    for c in chu:
        for k, v in RX_KT.findall(c or ""):
            gt.setdefault(k.upper() if k != "ø" else "ø", int(v))
        if gt:
            break
    if "Ø" in gt or "ø" in gt:
        v = gt.get("ø", gt.get("Ø"))
        return v, v, gt.get("H")
    if "L" in gt:
        return gt["L"], gt.get("W") or gt.get("D"), gt.get("H")
    return gt.get("W"), gt.get("D"), gt.get("H")


def doc_lich_cua(thu_muc):
    """Lịch cửa từ các tờ PW_TP_DR_8000_8008: mã, vị trí + ghi chú, tờ, kích thước. Một mã có thể xuất hiện 2 lần (giữ cả 2)."""
    out = []
    for f in sorted(os.listdir(thu_muc)):
        if "_DR_8000_8008" not in f or not f.lower().endswith(".pdf"):
            continue
        so = f.rsplit("-", 1)[-1][:-4]
        pg = fitz.open(os.path.join(thu_muc, f))[0]
        t = pg.get_text()
        kt = kich_thuoc_cua(pg)
        for p in re.split(r"REMARKS :\s*\n\s*LOCATION :\s*\n\s*PLAN\s*\n", t)[1:]:
            ls = [x.strip() for x in p.split("\n") if x.strip()]
            c = next((x for x in ls if re.fullmatch(r"D\d+[a-zA-Z]?", x)), None)
            if not c:
                continue
            mo = " ".join(ls[:ls.index(c)]).replace("SECTION ", "", 1)
            out.append({"ma": c, "mo_ta": mo, "to": so, "kt": kt.get(c)})
    return out


def hd_theo_thap(mo, thap_ten):
    """Bộ phụ kiện HD của cửa, có thể ghi riêng theo tháp ('HD-01 & HD-03 FOR WEALTH TOWER')."""
    cau = re.findall(r"((?:HD-\d{2}(?:\s*&\s*)?)+)\s*FOR\s+([A-Z, &]+?TOWERS?)", mo)
    if cau:
        for hd, tp in cau:
            if thap_ten.upper() in tp:
                return re.findall(r"HD-\d{2}", hd)
        return []
    return re.findall(r"HD-\d{2}", mo)


class Du:
    def __init__(self, goc):
        self.goc = os.path.abspath(goc)
        self.d = os.path.join(self.goc, "10_TENDER", "_data")
        self.cfg = json.load(open(os.path.join(self.d, "_lap_boq.json"), encoding="utf-8"))
        ch = json.load(open(os.path.join(self.d, "_cauhinh.json"), encoding="utf-8"))
        self.inp = os.path.join(self.goc, ch["input"])
        self.bv = os.path.join(self.inp, self.cfg["thu_muc_ban_ve"])
        self.tag = json.load(open(os.path.join(self.d, "_dwg_tag.json"), encoding="utf-8"))
        self.spec = json.load(open(os.path.join(self.d, "_spec_boc.json"), encoding="utf-8"))
        self.dl = doc_drawing_list(os.path.join(self.bv, self.cfg["drawing_list"]), self.cfg["sheet_drawing_list"])
        self.fm = {}
        self.lf = {}
        self.lf_con = collections.defaultdict(list)
        for f in self.spec:
            if "_LF_" in f["file"]:
                k = self.cfg["lf_khu"].get(re.search(r"_LF_(\d{4})", f["file"]).group(1))
                for x in f["ma"]:
                    x["khu"] = k
                    if "." in x["ma"]:
                        self.lf_con[(k, x["ma"].split(".")[0])].append(x)
                    else:
                        self.lf[(k, x["ma"])] = x
            elif "_FM_" in f["file"]:
                self.fm.update({x["ma"]: x for x in f["ma"]})
        self.cong_tac = self.doc_pham_vi()
        self.khu = self.gom_khu()

    def pdf_to(self, ds):
        """Đường dẫn PDF tờ lẻ theo số tờ (đuôi tên file)."""
        if not hasattr(self, "_pdf"):
            self._pdf = {}
            rx = re.compile(r"(?<!\d)(\d{4}(?:\.\d)?)(?:_[A-Z])?\.pdf$", re.I)
            for dp, _, fs in os.walk(self.bv):
                if "CAD" in dp.split(os.sep):
                    continue
                for x in fs:
                    m = rx.search(x)
                    if m and x.startswith("AID"):
                        self._pdf.setdefault(m[1], os.path.join(dp, x))
        return [self._pdf[x] for x in ds if x in self._pdf]

    def doc_pham_vi(self):
        ws = openpyxl.load_workbook(os.path.join(self.inp, self.cfg["pham_vi"]), data_only=True).worksheets[0]
        hd = [c.value for c in ws[1]]
        ic, ip, ig, it = hd.index("Cú pháp"), hd.index("PHẠM VI CÔNG VIỆC"), hd.index("GÓI THẦU"), hd.index("CẤP 3")
        out = collections.OrderedDict()
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[ip]:
                out[r[ic].strip()] = {"ten": r[it], "goi": r[ig]}
        out.update(self.cfg.get("ngoai_tick") or {})  # công tác không được tick nhưng user chốt phải tính
        return out

    def gom_khu(self):
        """{(tháp, khu): {to: {số tờ: {'loai', 'vl': Counter tường, 'tran': Counter, 'loose': Counter, 'cua': Counter, 'ref': Counter}}}}"""
        rx = re.compile(self.cfg["rx_ten_dwg"])
        out = collections.defaultdict(dict)
        soat = nap_soat(self.d)
        for rel, v in self.tag.items():
            m = rx.search(os.path.basename(rel))
            if not m:
                continue
            k = (m[1], m[3], m[2])
            for ten, l in list(v["to"].items()) + [("model:" + os.path.basename(rel), v["model"])]:
                so = so_to(ten) if not ten.startswith("model:") else ten
                o = out[k].setdefault(so, {"loai": "model" if ten.startswith("model:") else loai_to(so, self.dl), "tuong": collections.Counter(), "tran": collections.Counter(), "cua": collections.Counter(), "ref": collections.Counter(), "chu": [], "bm": collections.Counter()})
                for i_tag, t in enumerate(l["tag"]):
                    if ten.startswith("model:") and "gan" in t:
                        continue  # đã gán về tờ theo viewport, hoặc nằm ngoài mọi khung nên không hiện trên tờ nào
                    c = ma_cua(t)
                    if c:
                        o["cua"][c] += 1
                        continue
                    if "SHEET" in t["a"]:
                        o["ref"][t["a"]["SHEET"]] += 1
                    c = ma_tag(t)
                    if c:
                        o["tran" if t["b"] in BLOCK_TRAN else "tuong"][c] += 1
                        kq = soat.get(f"{so}#{i_tag}")
                        o["bm"][(c, kq["be_mat"] if kq else doan_be_mat(t, l)[0])] += 1
                o["chu"] += [x for x in l["chu"]] + [" ".join(t["ml"]) for t in l["tag"] if t.get("ml")]
        return out


def ten_vl(du, ma):
    x = du.fm.get(ma)
    if not x:
        return None, None
    return x.get("item desc", ""), x


def to_chu(ds, toi_da=8):
    ds = sorted({re.sub(r"^model:.*?_(\d{4}[^/]*?)\.dwg$", r"\1 (Model)", x) for x in ds})
    return ", ".join(ds[:toi_da]) + (f" và {len(ds) - toi_da} tờ khác" if len(ds) > toi_da else "")


def lap_dong(du):
    """{sheet: {công tác: {khu: [dòng]}}} và danh sách ngoài phạm vi."""
    cfg = du.cfg
    thap, khu_ten = cfg["thap"], cfg["khu"]
    bang = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(list)))
    ngoai = collections.defaultdict(lambda: collections.defaultdict(set))
    co_chung = collections.Counter()
    lich = doc_lich_cua(os.path.join(du.bv, "8000 Door"))
    lich_ma = collections.defaultdict(list)
    for x in lich:
        lich_ma[x["ma"].lower() if len(x["ma"]) > 2 else x["ma"]].append(x)
    lich_ma = {re.sub(r"^(D\d+)([a-z]?)$", lambda m: m[1] + m[2], k.replace("d", "D", 1)): v for k, v in lich_ma.items()}
    luat_vl = cfg["vat_lieu"]

    def luat(ma):
        return luat_vl["ma"].get(ma) or luat_vl["ho"].get(ma.split("-")[0])

    def sheet_cua(t, lv):
        return cfg["sheet_khoi_de"] if lv in cfg["tang_khoi_de"] else thap[t]["sheet"]

    for (t, kh, lv), tos in sorted(du.khu.items()):
        sh = sheet_cua(t, lv)
        k_ten = f"{thap[t]['ten']} — {khu_ten.get(kh, kh)}"
        dien_hinh = kh == "CR"
        # 1. vật liệu tường, trần
        # Bề mặt lấy theo từng tag (be_mat.py: nhãn, khung hình, kết quả soát hình), không suy từ loại tờ
        gom = {k: collections.defaultdict(set) for k in ("tuong", "tran", "portal", "san", "khac")}
        san_tinh, chi_tiet = cfg.get("san_tinh") or {}, cfg.get("chi_tiet") or {}  # tag sàn và chi tiết được user chốt phải có dòng
        chua_ro = collections.defaultdict(set)
        for so, o in tos.items():
            if so.startswith("model:") or so[0] in "123":
                for (c, bm), _ in o["bm"].items():
                    if bm in ("tuong", "can_soat", "khong_ro"):
                        gom["tuong"][c].add(so)
                        if bm != "tuong":
                            chua_ro[c].add(so)
                    elif bm == "tran":
                        gom["tran"][c].add(so)
                    elif bm == "portal" or (bm == "san" and c in san_tinh) or (bm == "khac" and c in chi_tiet):
                        gom[bm][c].add(so)
                    elif bm in ("san", "cua", "joinery", "khac") and luat(c) and c.split("-")[0] != "":
                        ngoai[sh][({"san": "Tag chỉ vào sàn", "cua": "Tag chỉ vào cửa, tính ở dòng cửa", "joinery": "Tag chỉ vào đồ gỗ, tính ở dòng joinery", "khac": "Tag chỉ vào chi tiết khác (con tiện, tấm chắn), CHƯA có dòng riêng"}[bm], "không tính ốp tường, trần")].add(f"{c} tờ {so}" if bm == "khac" else c)
        for vi_tri, g in gom.items():
            for c, s_to in sorted(g.items()):
                s_that = {x for x in s_to if not x.startswith("model:")}
                if "XX" in c:
                    continue
                r = luat(c)
                ho = c.split("-")[0]
                if not r:
                    if ho in du_ho_vl(du) or ho in ("MT", "ST", "PT", "CT", "CP", "SP", "LM", "VF", "AC", "UP"):
                        ngoai[sh][(ho, "trần" if vi_tri == "tran" else "tường")].add(c)
                    continue
                ct = san_tinh[c] if vi_tri == "san" else r["tran" if vi_tri == "tran" else "tuong"]
                if ct not in du.cong_tac:
                    continue
                ten, x = ten_vl(du, c)
                co = []
                if r.get("co"):
                    co.append(r["co"])
                if not s_that:
                    co.append("Tag nằm trong block ở Model space, chưa xác định được tờ")
                if vi_tri == "tuong" and chua_ro.get(c) and chua_ro[c] >= s_that:
                    co.append("Chưa xác định được bề mặt tag chỉ vào, đang xếp tạm vào ốp tường")
                lech = lech_area(cfg, x, t, kh)
                if lech:
                    co.append(lech)
                if ten is None:
                    co.append("Không thấy mã trong FM_7100")
                if x and "OMIT" in (x.get("area", "") + x.get("remarks", "")).upper():
                    co.append("Spec ghi TO BE OMITTED")
                if vi_tri == "portal":
                    co.append("Portal lỗ mở không có cánh: tách khỏi ốp tường, không có dòng cửa nhận")
                if vi_tri == "san":
                    co.append(f"Công tác {ct} không có trong ô tick, thêm theo chốt của user")
                if vi_tri == "khac":
                    co.append("Chi tiết riêng, không phải ốp tường, trần: số lượng chưa đếm")
                ten_ct = chi_tiet[c]["nd"] + " " if vi_tri == "khac" else ""
                nd = {"tuong": "Ốp tường ", "tran": "Ốp trần ", "portal": "Ốp portal lỗ mở ", "san": "Sàn ", "khac": ten_ct}[vi_tri] + (ten or c)
                ct_vl = "; ".join(y for y in [(x or {}).get("prod. code", ""), (x or {}).get("remarks", "")] if y and y != "NA")
                bang[sh][ct][k_ten].append({"ki_hieu": c, "nd": nd, "ct": ct_vl[:120], "dv": chi_tiet[c]["dv"] if vi_tri == "khac" else "m2", "sl": None, "to": sorted(s_that),
                                            "note": f"Tờ {to_chu(s_that or s_to)}. FM_7100 tr.{x['trang_ct'] if x and x.get('trang_ct') else '?'}. {'Số lượng' if vi_tri == 'khac' else 'Diện tích'} chưa đo.", "co": co})
        # 2. cửa gỗ + phụ kiện cửa
        dem = collections.defaultdict(lambda: {"mb": 0, "wc": 0, "to": set()})
        for so, o in tos.items():
            for c, n in o["cua"].items():
                d = dem[c]
                ten_to = so.replace("model:", "")
                if o["loai"] == "wc" or "_33" in ten_to or re.search(r"_3\d{3}", ten_to):
                    d["wc"] += n
                else:
                    d["mb"] = max(d["mb"], n)
                if not ten_to.endswith(".dwg"):
                    d["to"].add(ten_to)
        hd_dem = collections.Counter()
        hd_to = collections.defaultdict(set)
        hd_lich = collections.defaultdict(set)
        for c, d in sorted(dem.items(), key=lambda x: (int(re.search(r"\d+", x[0]).group()), x[0])):
            ds = lich_ma.get(c) or lich_ma.get(c.upper()) or []
            sl = max(d["mb"], d["wc"])
            mo = " / ".join(x["mo_ta"] for x in ds)
            co = []
            if not ds:
                if c == "D36":
                    ngoai[sh][("cửa cubicle WC", "không có trong lịch cửa")].add(f"{c}×{sl}")
                    continue
                co.append("Mã cửa có trên bản vẽ, không có trong lịch cửa")
            if len(ds) > 1:
                co.append(f"Mã {c} có {len(ds)} mục trong lịch cửa (tờ {ds[0]['to']}): {ds[1]['mo_ta'][:40]}")
            if re.search(r"GLASS|FRAMELESS", mo) and "TIMBER" not in mo:
                ngoai[sh][("cửa kính", "AL.GL không được tick")].add(f"{c}×{sl}")
                continue
            if ds and "TIMBER" not in mo:
                co.append("Lịch cửa không ghi rõ cửa gỗ")
            if d["mb"] and d["wc"] and d["mb"] != d["wc"]:
                co.append(f"Đếm trên mặt bằng chung {d['mb']}, trên tờ WC {d['wc']}, lấy số lớn")
            hd = hd_theo_thap(mo, thap[t]["ten"])
            if not hd and ds:
                co.append("Lịch cửa chưa ghi bộ phụ kiện HD")
            for h in hd:
                hd_dem[h] += sl
                hd_to[h].add(c)
                hd_lich[h].update(x["to"] for x in ds)
            ten_cua = re.sub(r"\s+", " ", ds[0]["mo_ta"].split(" SINGLE ")[0].split(" DOUBLE ")[0].split(" TIMBER ")[0].split(" CONCEALED ")[0])[:90] if ds else c
            la = "2 cánh" if re.search(r"DOUBLE LEAF", mo) else ("1 cánh" if re.search(r"SINGLE LEAF", mo) else "")
            ghi = f"Tờ lịch {ds[0]['to'] if ds else '?'}; đếm tag tờ {to_chu(d['to'], 4)}." + (" Số lượng là cho 1 tầng điển hình; KL chờ số tầng áp dụng." if dien_hinh else "")
            kt = ds[0].get("kt") if ds else None
            if ds and not kt:
                co.append("Không đọc được kích thước trên tờ lịch cửa")
            ct_cua = (f"Rộng {kt[0]} x Cao {kt[1]} mm. " if kt else "") + ("Phụ kiện: " + "; ".join(hd) if hd else "")
            bang[sh]["DO.WD"][k_ten].append({"ki_hieu": c, "nd": f"Cửa gỗ {la} — {ten_cua}".replace("  ", " "), "ct": ct_cua.strip(), "dv": "bộ", "sl": sl, "kl": None if dien_hinh else sl, "note": ghi, "co": co,
                                             "to": sorted(d["to"] | {x["to"] for x in ds}), "dai": kt[1] if kt else None, "rong": kt[0] if kt else None})
        for h, n in sorted(hd_dem.items()):
            bang[sh]["HW.DO"][k_ten].append({"ki_hieu": h, "nd": f"Bộ phụ kiện cửa {h}", "ct": ten_hd(du, h), "dv": "bộ", "sl": n, "kl": None if dien_hinh else n, "to": sorted(hd_lich[h]),
                                             "note": f"Cho cửa {', '.join(sorted(hd_to[h]))}. IM_7400." + (" Số lượng cho 1 tầng điển hình." if dien_hinh else ""), "co": []})
        # 3. đồ rời, đèn, rèm
        khoa = f"{t}_{kh}"
        dem_lo = collections.Counter()
        nguon = collections.defaultdict(set)
        theo_loai = collections.defaultdict(collections.Counter)
        for so, o in tos.items():
            for c, n in o["tuong"].items():
                if re.match(r"(AX|CX|BX|CG|DX|LX|SX)-\d{4}", c):
                    theo_loai[o["loai"]][c] += n
                    nguon[c].add(so)
        cac_ma = set(c for (kk, c) in du.lf if kk == khoa) | set(nguon)
        for c in sorted(cac_ma):
            ho = c[:2]
            x = du.lf.get((khoa, c))
            bo = cfg["loose"]["bo"].get(ho)
            n = max([theo_loai[l][c] for l in ("ffe", "tran", "tuong_mb")] + [0])
            if not n:
                n = sum(theo_loai[l][c] for l in ("wc",)) or max(theo_loai[l][c] for l in theo_loai) if theo_loai else 0
            if bo:
                ngoai[sh][(bo, "không được tick")].add(f"{c}×{n}" if n else c)
                continue
            ten = (x or {}).get("item desc", "")
            ct = cfg["loose"]["ma_ho"].get(ho)
            if not ct:
                for rx, m in cfg["loose"]["luat"]:
                    if re.search(rx, ten):
                        ct = m
                        break
                else:
                    ct = cfg["loose"]["mac_dinh"]
            co = []
            if not x:
                co.append("Có tag trên bản vẽ, không thấy trong lịch FF&E của khu")
            if not n:
                co.append("Có trong lịch FF&E, không thấy tag trên bản vẽ")
            if ho == "DX":
                co.append("Rèm: số là số tag, khối lượng md/m2 chưa đo")
            con = du.lf_con.get((khoa, c), [])
            ct_vl = (x or {}).get("finishes", "")
            if con:
                ct_vl = (ct_vl + " | Kèm: " if ct_vl else "Kèm: ") + "; ".join(y.get("item desc", "")[:40] for y in con[:5])
            nhan = re.search(r"LF_\d{4}", x["file"]).group() if x else ""
            ghi = f"Tag tờ {to_chu(nguon[c], 4) or '—'}" + (f". {nhan} tr.{x['trang']}" if x else "") + (" Số lượng cho 1 tầng điển hình." if dien_hinh else "")
            dai, rong, cao = kich_thuoc_do_roi((x or {}).get("dimensions"), ten)
            if not dai and ho != "DX":
                co.append("Spec không ghi kích thước đọc được")
            ct_kt = f"Cao {cao} mm. " if cao else ""
            bang[sh][ct][k_ten].append({"ki_hieu": c, "nd": ten[:110] or c, "ct": (ct_kt + ct_vl)[:220], "dv": "cái" if ho != "DX" else "bộ", "sl": n or None, "kl": (None if dien_hinh else n) or None, "note": ghi, "co": co,
                                        "to": sorted(nguon[c]), "dai": dai, "rong": rong})
        # 4. joinery
        jo = collections.defaultdict(lambda: {"to": set(), "vl": collections.Counter()})
        for so, o in tos.items():
            if o["loai"] == "joinery":
                goc_so = so.split(".")[0]
                jo[goc_so]["to"].add(so)
                jo[goc_so]["vl"].update(o["tuong"])
        ref = collections.Counter()
        for so, o in tos.items():
            if o["loai"] == "tuong_mb":
                for r, n in o["ref"].items():
                    ref[r.split(".")[0]] += n
        for goc_so, j in sorted(jo.items()):
            v = du.dl.get(goc_so + ".1") or du.dl.get(goc_so)
            ten = v[6] if v else ""
            co = []
            if not ten:
                co.append("Tờ joinery không có trong Drawing List, chưa có tên")
            ct = cfg["joinery_mac_dinh"]
            for rx, m in cfg["joinery"]:
                if re.search(rx, ten):
                    ct = m
            sl = ref.get(goc_so) or 0
            if not sl:
                co.append("Không thấy ký hiệu dẫn tới tờ này trên mặt bằng, đặt tạm 1")
                sl = 1
            vl = sorted(c for c in j["vl"] if re.match(r"(WD|ST|MT|GL|UP|WC|LM|SP)-", c) and "XX" not in c)
            ngang, dung = None, None
            for f_pdf in du.pdf_to(sorted(j["to"])[:1]):
                ngang, dung = kich_thuoc_joinery(f_pdf)
            ghi_kt = " Kích thước là số lớn nhất đọc được trên tờ đầu: ngang, đứng; đứng có thể là chiều cao hoặc chiều sâu." if ngang else ""
            if not ngang:
                co.append("Không đọc được kích thước trên tờ joinery")
            bang[sh][ct][k_ten].append({"ki_hieu": goc_so, "nd": ten.title() or f"Joinery {goc_so}", "ct": ", ".join(vl), "dv": "bộ", "sl": sl, "kl": None if dien_hinh else sl, "to": sorted(j["to"]), "dai": ngang, "rong": dung,
                                        "note": f"Tờ {to_chu(j['to'], 5)}. Đá và LED đi kèm chưa tính theo PHẠM VI." + ghi_kt, "co": co})
        if jo and "HW.FU" in du.cong_tac:
            bang[sh]["HW.FU"][k_ten].append({"ki_hieu": "", "nd": "Phụ kiện đi kèm các hạng mục joinery khu này", "ct": "", "dv": "khu", "sl": None, "to": sorted(x for j in jo.values() for x in j["to"]), "note": "IM_7400 chỉ có bộ HD cho cửa, chưa có lịch phụ kiện đồ nội thất.", "co": ["Chưa có lịch phụ kiện đồ nội thất, cần thiết kế cung cấp"]})
        # 5. vanity
        chu_wc = [x for so, o in tos.items() if o["loai"] == "wc" for x in o["chu"]]
        s_wc = sorted(so for so, o in tos.items() if o["loai"] == "wc")
        if "FUR.WC" in du.cong_tac and any(re.search(r"VANITY|BASIN|COUNTERTOP|MIRROR", x) for x in chu_wc):
            n_tg = sorted({m for x in chu_wc for m in re.findall(r"(?:MIRROR\s*)?GL\s*\d{2}", x)})
            bang[sh]["FUR.WC"][k_ten].append({"ki_hieu": "", "nd": "Tủ lavabo, gương trong WC của khu", "ct": ", ".join(n_tg), "dv": "bộ", "sl": None, "to": s_wc,
                                              "note": f"Tờ {to_chu(s_wc, 5)}. Mặt đá lavabo chưa tính theo PHẠM VI.", "co": ["Chưa đếm: nhận biết qua chữ MIRROR/BASIN trên tờ WC"]})
    # cửa có tag ở từ 2 khu trở lên mà lịch cửa không ghi nhiều tháp: có thể là 1 cửa nối 2 khu hoặc tag sai khu
    noi = collections.defaultdict(list)
    for sh, cts in bang.items():
        for kh, ds in cts.get("DO.WD", {}).items():
            for d in ds:
                noi[d["ki_hieu"]].append((kh, d))
    for c, ds in noi.items():
        mo = " ".join(x["mo_ta"] for x in lich_ma.get(c, []))
        if len(ds) > 1 and not re.search(r"ALL TOWERS|WEALTH & HAPPINESS|HAPPINESS,\s*PEACE", mo):
            for kh, d in ds:
                khac = "; ".join(k for k, _ in ds if k != kh)
                d["co"].append(f"Cùng mã cửa còn có tag ở: {khac}. Có thể đếm trùng hoặc tag sai khu")
    return bang, ngoai, lich


def lech_area(cfg, x, t, kh):
    """So 'Area Used' của spec với tháp, khu đang xét. Chỉ báo khi spec có nêu tháp hoặc khu mà không có nơi này."""
    if not x or not x.get("area"):
        return None
    a = " " + x["area"].upper().replace(",", " ").replace("(", " ").replace(")", " ") + " "
    tu_khoa = cfg.get("tu_khoa_area") or {}
    thap_tk, khu_tk = tu_khoa.get("thap", {}), tu_khoa.get("khu", {})
    co_thap = [k for k, v in thap_tk.items() if any(f" {w} " in a for w in v)]
    co_khu = [k for k, v in khu_tk.items() if any(w in a for w in v)]
    if (co_thap and t not in co_thap) or (co_khu and kh not in co_khu):
        return f"Spec ghi dùng ở: {x['area'][:60]}; tag lại có ở khu này"
    return None


def du_ho_vl(du):
    return {m.split("-")[0] for m in du.fm}


def ten_hd(du, h):
    for f in du.spec:
        if "_IM_" in f["file"]:
            for x in f["ma"]:
                if x["ma"] == h and x.get("item desc"):
                    return x["item desc"][:110]
    return ""


def them_cabin(du, bang):
    c = du.cfg.get("cabin")
    p = os.path.join(du.d, c["file_tag"]) if c else None
    if not p or not os.path.exists(p):
        return
    tag = json.load(open(p, encoding="utf-8"))
    gom = {"tuong": collections.defaultdict(set), "tran": collections.defaultdict(set)}
    for v in tag.values():
        for ten, l in v["to"].items():
            for t in l["tag"]:
                m = ma_tag(t)
                if m:
                    gom["tran" if t["b"] in BLOCK_TRAN else "tuong"][m].add(so_to(ten))
    luat_vl = du.cfg["vat_lieu"]
    for t, tv in du.cfg["thap"].items():
        k = f"{tv['ten']} — {c['ten']}"
        for vt, g in gom.items():
            for m, s_to in sorted(g.items()):
                r = luat_vl["ma"].get(m) or luat_vl["ho"].get(m.split("-")[0])
                if not r or "XX" in m or r[vt] not in du.cong_tac:
                    continue
                ten, x = ten_vl(du, m)
                bang[tv["sheet"]][r[vt]][k].append({"ki_hieu": m, "nd": ("Ốp vách cabin " if vt == "tuong" else "Ốp trần cabin ") + (ten or m), "ct": "", "dv": "m2", "sl": None, "to": sorted(s_to),
                                                    "note": f"Tờ {to_chu(s_to)}. Bản vẽ điển hình; số cabin từng tháp chưa có.", "co": ([r["co"]] if r.get("co") else []) + ([] if ten else ["Không thấy mã trong FM_7100"])})


VIEN = Side(style="thin")
KHUNG = Border(left=VIEN, right=VIEN, top=VIEN, bottom=VIEN)
F = "Times New Roman"


def ghi_o(ws, r, c, v, dam=False, nghieng=False, boc=False, nen=None, mau=None):
    o = ws.cell(row=r, column=c, value=v)
    o.font = Font(name=F, size=11, bold=dam, italic=nghieng, color=mau)
    o.border = KHUNG
    o.alignment = Alignment(wrap_text=boc, vertical="top")
    if nen:
        o.fill = PatternFill("solid", fgColor=nen)
    return o


def ghi_file(du, bang, ngoai, lich):
    cfg = du.cfg
    src = os.path.join(du.inp, cfg["form"])
    ra = os.path.join(du.inp, cfg["ra"].format(yymmdd=datetime.date.today().strftime("%y%m%d")))
    os.makedirs(os.path.dirname(ra), exist_ok=True)
    shutil.copy2(src, ra)
    wb = openpyxl.load_workbook(ra)
    tong = collections.Counter()
    canh = {}
    for sh, cts in bang.items():
        ws = next(x for x in wb.worksheets if x.title.strip() == sh)
        for rr in range(5, max(ws.max_row, 30) + 1):
            for cc in range(1, 13):
                o = ws.cell(row=rr, column=cc)
                o.fill = PatternFill(fill_type=None)
                o.value = None
        r = 5
        stt = 0
        for i, ct in enumerate(k for k in du.cong_tac if k in cts):
            chu = chr(65 + i)
            tt = du.cong_tac[ct]
            ghi_o(ws, r, 1, chu, dam=True, nen="FFFF00")
            ghi_o(ws, r, 2, ct, dam=True, nen="FFFF00")
            ghi_o(ws, r, 3, tt["ten"], dam=True, nen="FFFF00")
            for c in range(4, 12):
                ghi_o(ws, r, c, None, nen="FFFF00")
            ghi_o(ws, r, 12, f"Gói {tt['goi']}", dam=True, nen="FFFF00")
            r += 1
            for j, (kh, ds) in enumerate(sorted(cts[ct].items()), 1):
                ghi_o(ws, r, 1, f"{chu}.{j}", dam=True, nghieng=True)
                ghi_o(ws, r, 2, None)
                ghi_o(ws, r, 3, kh, dam=True, nghieng=True, boc=True)
                for c in range(4, 13):
                    ghi_o(ws, r, c, None)
                r += 1
                for d in ds:
                    stt += 1
                    tong[sh] += 1
                    co = " ".join("[XEM] " + x + "." for x in d["co"])
                    ghi_o(ws, r, 1, stt)
                    ghi_o(ws, r, 2, d["ki_hieu"])
                    ghi_o(ws, r, 3, d["nd"], boc=True)
                    ghi_o(ws, r, 4, d["ct"], boc=True)
                    ghi_o(ws, r, 5, d.get("dai"))
                    ghi_o(ws, r, 6, d.get("rong"))
                    canh[f"{ws.title}!{r}"] = {"to": d.get("to", []), "ma": d["ki_hieu"], "ct": ct, "dv": d["dv"], "sl": d.get("sl"), "khu": kh, "dai": d.get("dai"), "rong": d.get("rong")}
                    ghi_o(ws, r, 7, d.get("sl"))
                    ghi_o(ws, r, 8, d["dv"])
                    ghi_o(ws, r, 9, d.get("kl"))
                    ghi_o(ws, r, 10, None)
                    ghi_o(ws, r, 11, None)
                    ghi_o(ws, r, 12, (d["note"] + " " + co).strip(), boc=True, mau="C00000" if d["co"] else None)
                    r += 1
        ws.column_dimensions["C"].width = 42
        ws.column_dimensions["D"].width = 36
        ws.column_dimensions["L"].width = 62
    wh = wb.worksheets[0]  # Phần hầm: để trống, có ghi chú
    wh["C5"] = "Không có hạng mục BOHO: hồ sơ ID không có bản vẽ hầm, khu vực gói 3–6 không nhắc hầm."
    wh["C5"].font = Font(name=F, size=11, italic=True)
    wg = wb.create_sheet("Ghi chú & ngoài phạm vi")
    wg["A1"] = "Ghi chú chung"
    wg["A1"].font = Font(name=F, size=12, bold=True)
    r = 2
    for x in cfg["ghi_chu_chung"]:
        wg.cell(row=r, column=1, value=x).font = Font(name=F, size=11)
        wg.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        wg.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        wg.row_dimensions[r].height = 46
        r += 1
    r += 1
    wg.cell(row=r, column=1, value="Có trên bản vẽ nhưng không thuộc phạm vi BOHO được tick").font = Font(name=F, size=12, bold=True)
    r += 1
    for j, h in enumerate(["Sheet", "Nhóm", "Mã (mã×số đếm tag nếu có)"], 1):
        ghi_o(wg, r, j, h, dam=True, nen="D9D9D9")
    r += 1
    for sh in sorted(ngoai):
        for (a, b), ms in sorted(ngoai[sh].items()):
            ghi_o(wg, r, 1, sh)
            ghi_o(wg, r, 2, f"{a} ({b})", boc=True)
            ghi_o(wg, r, 3, ", ".join(sorted(ms)), boc=True)
            r += 1
    wg.column_dimensions["A"].width = 16
    wg.column_dimensions["B"].width = 44
    wg.column_dimensions["C"].width = 90
    wb.save(ra)
    json.dump(canh, open(os.path.join(du.d, "_lap_boq_dong.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    return ra, tong


def main(goc):
    du = Du(goc)
    bang, ngoai, lich = lap_dong(du)
    them_cabin(du, bang)
    ra, tong = ghi_file(du, bang, ngoai, lich)
    print("Đã ghi", ra)
    # Tên file có ngày: sang ngày mới là file mới, phải trỏ cấu hình sang đó, nếu không bước boq vẫn đọc file cũ
    p_ch = os.path.join(du.d, "_cauhinh.json")
    ch = json.load(open(p_ch, encoding="utf-8"))
    ch["boq"] = os.path.relpath(ra, du.inp).replace("\\", "/")
    json.dump(ch, open(p_ch, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("Số dòng hạng mục:", dict(tong), "tổng", sum(tong.values()))
    dem_ct = collections.Counter()
    dem_co = 0
    for sh, cts in bang.items():
        for ct, kh in cts.items():
            for ds in kh.values():
                dem_ct[ct] += len(ds)
                dem_co += sum(1 for d in ds if d["co"])
    print("Theo công tác:", dict(sorted(dem_ct.items())))
    print("Công tác không có dòng:", [k for k in du.cong_tac if k not in dem_ct])
    print("Dòng có cờ XEM:", dem_co)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1])
