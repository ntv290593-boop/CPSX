# -*- coding: utf-8 -*-
r"""docsau.py — phần ĐỌC SÂU của skill 10-tender (chuyển từ 01-scan), td.py gọi vào.
Áp cho toàn gói: mọi dòng BOQ còn làm (có đơn vị, không thuộc sheet Prelims/Provisional/Summary, Trạng thái không phải Omit/Loại trừ/Không áp dụng).
Khác 01-scan: đơn vị đọc là CỤM (các dòng cùng mô tả + mã VL + mã FF&E + tờ bản vẽ hiệu lực), khóa cụm = `Khóa` của dòng đầu tiên, cột `Cụm đọc`.

  cmd_docsau   gom cụm, hợp đồng kiểm theo profile (form\PROFILE.json), chia nhóm THEO TRANG (mỗi trang 1 nhóm chủ),
               nhóm tìm tờ, nhóm đối chiếu; nháp Phạm vi nhà máy, BOQ_MAP
  cmd_ghep     TRANG_QS → QUAN_SAT / SCOPE / BOQ_MAP / DAY_TUONG (tool ghép, LLM không chép lại), báo độ phủ → _ghep.json
  cmd_gemini   Gemini đọc độc lập, --trang lưu số theo trang (ghep chia về cụm), --cum lưu số theo cụm
  cmd_neo      kiểm giá trị LLM ghi có trên trang được link không (chỉ XEM)
  cmd_tincay   Độ tin cậy, trạng thái KL, issue nháp
  kiem         phần kiểm thêm cho td.py check
"""
import glob
import json
import os
import re
import sys
from collections import Counter, OrderedDict

import td
from td import jdump, jload, links_of, merge, opt, s, write_tool, _txt
from scan_core import _fold
import doc_file
import gemini_doc
import neo as neo_lib

GEM_SO_TRUONG = "Kiểm chéo Gemini số"
PRICE_FIELDS = ["Kích thước BV", "Có cánh", "Ngăn kéo", "Kệ/đợt", "Mặt đá", "Gương", "Kính", "Kim loại", "Giấy/vải"]
MUC = {"CAO": 3, "TB": 2, "THAP": 1}
GIA_RX = re.compile(r"dot|canh|ngan|kich thuoc|\bsau\b|\brong\b|\bcao\b|\bday\b|mat da|\bda\b|kinh|guong|kim loai|vat lieu|giay|vai")
CHUOI = re.compile(r"(?<![+×x*\d.,])(\d+(?:[.,]\d+)?(?:\s*\+\s*\d+(?:[.,]\d+)?)+)\s*=\s*(\d+(?:[.,]\d+)?)")
KHU_CA_DONG = "(cả dòng)"
TT_BO = ("Omit", "Loại trừ", "Không áp dụng")

# Nhãn Phạm vi nhà máy: chép luật từ 01-scan. Ở đây CHỈ GẮN NHÃN, không lọc dòng (tender làm toàn gói).
VT_XUONG = re.compile(r"\b(van ep|van mdf|mdf|veneer|melamine|laminate|picomat|nem\b|van nen|go (tram|ash|tu nhien|dac|tap|oc cho)|son pu)\b")
SC_PK_CUA = r"^(ban le|tay nam|tay gat|than khoa|ruot khoa|chan cua|dem khi|khoa)\b"
SC_OUT = [
    ("Khóa điện tử", r"khoa (dien tu|tu|the)|digital lock", True),
    ("Sơn nước", r"^(thi cong )?son\b", False),
    ("Thạch cao", r"tran (chim|thach cao|phang|giat cap)|hoc den|mat dung|nap tham|thanh [vz]\b|khoet lo|mieng gio|hop den|hop man|shadowline|^thach cao", True),
    ("Đá", r"^lat da\b|ngach cua|^op da\b", False),
    ("Sàn", r"^lat van san|^san\b|san go|^lat san", False),
    ("Nẹp kim loại", r"^(nep|u|v)\b.*(inox|zenta|nhom|dong)|^nep\b", True),
    ("Lam nhựa/nhôm", r"\blam (nhua|nhom|composite)|composite", True),
    ("Cửa thép", r"cua thep", True),
    ("Cửa nhôm kính", r"cua nhom|vach nhom", True),
    ("Vách compact", r"compact", True),
    ("Kim loại trang trí", r"^op (inox|dong)|lam (inox|kim loai)", True),
    ("Lan can", r"lan can", True),
    ("Đèn", r"^(den|led)\b|\bden (tha|ban|tran|trang tri)", True),
    ("Điện", r"o cam|cong tac|di day", True),
    ("Thiết bị", r"^(chau|voi|bep|hut mui|lo nuong|tv|ket sat|may|thiet bi)\b", True),
    ("Nệm", r"^nem\b", True), ("Chăn ga gối", r"^(chan|ga giuong|drap)\b", True),
    ("Rèm", r"^rem\b", True), ("Thảm", r"^tham\b", True),
    ("Decor", r"^(tranh|khung tranh|decor|binh hoa|do trang tri)\b", True), ("Signage", r"signage|bien hieu", True),
]
SC_IN = [
    ("mộc", r"picomat"),
    ("cửa", r"^vach kinh|vach kinh|^cua\b|^cung cap va lap dat vach kinh"),
    ("mộc", r"guong"),
    ("cabinet", r"\btu\b|he tu|^bang|^hop che|^dot tu|^khung che|^he (bang|khung)|bench|vanity|ban console|ban lam viec|locker|minibar|quay (le tan|bar)|\bke bep"),
    ("loose", r"^(goi|sofa|giuong|ghe|ban|ke|bục|buc)\b|goi (trang tri|nam|tay)"),
]


def scope_phan_loai(mt):
    """Mô tả dòng BOQ -> nhãn Phạm vi nhà máy, None = tool không chắc (LLM quyết ở bước docsau)."""
    S = _fold(mt)
    if re.search(SC_PK_CUA, S) and not re.search(r"khoa (dien tu|tu|the)", S):
        return "TRONG SCOPE – cửa (phụ kiện cửa)"
    if re.search(r"picomat", S):
        return "TRONG SCOPE – mộc (picomat CNC)"
    if re.search(r"hop man", S) and re.search(r"van ep|van mdf", S):
        return "TRONG SCOPE – mộc (hộp màn lót ván ép)"
    if re.search(r"^(len chan|chi go|chi nhua)|\blen chan\b", S):
        if re.search(r"mdf|van (ep|nen)|laminate", S):
            return "TRONG SCOPE – mộc (len chân MDF/ván)"
        if re.search(r"nhua", S):
            return "KHÔNG THUỘC SCOPE NHÀ MÁY – Len/chỉ nhựa"
        return None
    if re.search(r"\bgach\b", S) and not re.match(r"^(nep|u|v)\b", S):
        return "KHÔNG THUỘC SCOPE NHÀ MÁY – Gạch"
    for nhom, rx, kiem_vt in SC_OUT:
        if re.search(rx, S):
            if kiem_vt:
                vt = sorted({m.group(0).strip() for m in VT_XUONG.finditer(S)})
                if vt:
                    return f"CHƯA RÕ – {nhom} nhưng có {', '.join(vt[:3])}"
            return f"KHÔNG THUỘC SCOPE NHÀ MÁY – {nhom}"
    for xuong, rx in SC_IN:
        if re.search(rx, S):
            return f"TRONG SCOPE – {xuong}"
    return None


def worse(a, b):
    return a if MUC[a] <= MUC[b] else b


def can_doc(p, r):
    """Dòng thuộc phạm vi đọc sâu: toàn gói trừ dòng không ai làm và sheet không có bản vẽ."""
    return td.can_ma(p, r) and s(r.get("Trạng thái")) not in TT_BO


def khoa_cum(r):
    """Hai dòng cùng cụm khi cùng mô tả, mã VL, mã FF&E và cùng bộ tờ. Khác tờ là khác món, không gộp."""
    to = sorted({(l[1], str(l[2]) if len(l) > 2 else "") for l in links_of(r.get("Bản vẽ hiệu lực"))})
    return (re.sub(r"\W+", " ", _fold(s(r.get("Mô tả BOQ")))).strip(), s(r.get("Mã VL")), s(r.get("Mã spec/FF&E")), tuple(to))


def tach_phong(v):
    """'SK-01:2 | Hành lang L4:1.5' -> [('SK-01', 2.0), ('Hành lang L4', 1.5)]."""
    out = []
    for x in s(v).split(" | "):
        if ":" in x:
            a, b = x.rsplit(":", 1)
            try:
                out.append((a.strip(), float(b)))
            except ValueError:
                pass
    return out


def nap_cum(p):
    return (jload(p.f("_docsau.json")) or {}).get("cum", {})


def co_quan_sat(res):
    q = {}
    for r in res["QUAN_SAT"]:
        q.setdefault(s(r.get("Cụm đọc")), []).append(r)
    return q


# ---------------------------------------------------------------- docsau (đọc theo trang, plan PAGEFIRST 16/11)
# Đơn vị tổ chức = TRANG (mỗi trang đúng 1 nhóm chủ, đọc ảnh 1 lần cho mọi món trên trang).
# Đơn vị kiểm soát = cụm × trường theo hồ sơ trường (form\PROFILE.json). TRANG_QS (LLM ghi) → ghep (tool) → QUAN_SAT/SCOPE/BOQ_MAP/DAY_TUONG.
TRANG_NHOM = 8     # số trang tối đa mỗi nhóm đọc trang
MON_NHOM = 40      # tổng lượt món (cụm × trang) tối đa mỗi nhóm đọc trang
CUM_TIM = 8        # cụm chưa có tờ: số cụm mỗi nhóm tìm tờ
CUM_DOICHIEU = 25  # cụm mỗi nhóm đối chiếu (không mở ảnh)
XEMCUM_MAX = 24000  # xemcum nhiều cụm: dừng trước khi output quá dài (Bash cắt ~30k ký tự, phần thừa phải Read lại)
CHUA_PL = "CHƯA PHÂN LOẠI"
BB, DK = "BB", "DK"
PHIEN_BAN = 3      # _docsau.json bản đọc theo trang
RX_KT = re.compile(r"\d+(?:[.,]\d+)?(?:\s*(?:mm|cm|m)\b)?(?:\s*[x×X*]\s*\d+(?:[.,]\d+)?(?:\s*(?:mm|cm|m)\b)?)+"
                   r"|\b[WDHL](?:\s*[:=]\s*|\s+)\d+(?:[.,]\d+)?\s*(?:mm|cm|m)?"  # 'W 1500', 'H: 3000'; tag 'D01' liền số nên giữ
                   r"|\d+(?:[.,]\d+)?\s*(?:mm|cm|m2|m²|md|m)\b", re.I)


def khoa_trang(l):
    pg = l[2] if len(l) > 2 and isinstance(l[2], int) else None
    return f"{l[1]}|{pg if pg is not None else ''}", pg


def to_cua(scope, c):
    """Các khóa trang <FID>|<trang> của cụm c theo Bản vẽ hiệu lực, giữ thứ tự, không trùng."""
    return list(dict.fromkeys(khoa_trang(l)[0] for l in links_of(scope.get(c, {}).get("Bản vẽ hiệu lực"))))


def nap_profile():
    f = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "form", "PROFILE.json")
    prof = jload(f)
    if not prof:
        sys.exit(f"LỖI: không đọc được {f}")
    return prof


def profile_cum(prof, r):
    """Loại món theo `Mã cấp 3` đã phân loại (không đoán từ mô tả). Không có / KHÔNG RÕ → CHƯA PHÂN LOẠI."""
    c3 = s(r.get("Mã cấp 3"))
    if not c3 or _fold(c3).startswith("khong ro"):
        return CHUA_PL
    return next((n for n, x in prof["profile"].items() if c3 in x.get("cap3", [])), "CHUNG")


def hop_dong(prof, ten, dv):
    """Hợp đồng kiểm của 1 cụm: {trường: {muc, gia, to}}; `KL theo BV` bắt buộc khi đơn vị là đơn vị đếm."""
    tr = {k: dict(v) for k, v in prof["profile"]["CHUNG" if ten == CHUA_PL else ten]["truong"].items()}
    dem = {_fold(x).strip(" .") for x in prof.get("kl_dem", [])}
    tr["KL theo BV"] = {"muc": BB, "gia": True, "to": ["MB", "BTK"]} if _fold(s(dv)).strip(" .") in dem else {"muc": DK, "to": ["BTK"]}
    return tr


def loai_to(prof, *txt):
    """Loại tờ đoán từ nhãn link / PAGES (Loại tờ, Tên BV). [] = chưa rõ → mọi trường của món đều cần xem trên tờ."""
    t, out = _fold(" ".join(s(_txt(x)) for x in txt)), []
    for k, rx in prof["loai_to"].items():
        if re.search(rx, t):
            out.append(k)
            if k == "MBT":
                t = re.sub(rx, " ", t)  # 'mặt bằng trần' không tính thêm là MB
    return out


def truong_tren_to(truong, loai):
    """-> (BB, DK) các trường nên tìm trên tờ loại này. Tờ chưa rõ loại: mọi trường."""
    bb, dk = [], []
    for f, x in truong.items():
        if not loai or not x.get("to") or set(x["to"]) & set(loai):
            (bb if x.get("muc") == BB else dk).append(f)
    return bb, dk


def ten_mon(r, n=110):
    """Tên món cho danh sách đọc trang: bỏ kích thước và mã hoàn thiện của BOQ (đáp án của trường cần đọc), giữ tag/mã FF&E để tìm trên hình."""
    t = s(r.get("Mô tả BOQ"))
    for m in sorted({x.strip() for x in re.split(r"[;,]", s(r.get("Mã VL"))) if x.strip()}, key=len, reverse=True):
        t = re.sub(re.escape(m), "", t, flags=re.I)
    t = re.sub(r"\s+", " ", re.sub(r"\s*[×*]\s*|\s[xX](?=\s)", " ", RX_KT.sub("", t))).strip(" ,;:-")
    return t[:n] + ("…" if len(t) > n else "")


RX_TO = re.compile(r"(?<![A-Z0-9-])[A-Z]{1,5}-\d{3,5}[A-Z]?(?![0-9])")  # số tờ ngắn trong nhãn link: 'FUR-1512'


def tim_bv(rx, t):
    """Số BV theo regex dự án (`dwg_regex` trong _cauhinh.json, mặc định td.RX_DWG)."""
    return {m if isinstance(m, str) else m[0] for m in rx.findall(s(t).upper())}


def so_to_nhan(nhan, dwg):
    """Số tờ trong nhãn link, bỏ số bộ đã có trong sổ DWG: 'FUR-0011 p13 BF-12 FUR-2111' -> ['FUR-2111']."""
    return [x for x in dict.fromkeys(RX_TO.findall(s(nhan).upper())) if x not in dwg]


def lap_manifest(res, scope, cum, chon, prof=None, rx=None):
    """Đảo cụm → tờ thành tờ → cụm theo `Bản vẽ hiệu lực` của dòng đại diện, chỉ cụm giao lần này (chon).
    -> {"tong", "trang": {"<FID>|<trang>": {to, fid, trang, rev, so_to, loai, cum, so_cum}}}; trang = null khi link không có số trang.
    rev = ngày issue + đợt hiệu lực của số BV trong sổ DWG; so_to = số tờ ngắn trong nhãn (để PAGES khỏi cần crop khung tên).
    Loại tờ: nhãn link + PAGES; `Tên tờ` của sổ DWG chỉ dùng khi số BV đó ứng với đúng 1 trang (sổ ghi theo bộ thì tên bộ không nói loại từng tờ)."""
    dwg, chon_set, trang = res.get("DWG", {}), set(chon), OrderedDict()
    rx = rx or re.compile(td.RX_DWG)
    for c in cum:
        if c not in chon_set:
            continue
        for l in links_of(scope[c].get("Bản vẽ hiệu lực")):
            key, pg = khoa_trang(l)
            t = trang.setdefault(key, {"to": l[0], "fid": l[1], "trang": pg, "rev": "", "loai": [], "cum": []})
            if c not in t["cum"]:
                t["cum"].append(c)
    bv_trang = {kt: sorted(b for b in tim_bv(rx, t["to"]) if b in dwg) for kt, t in trang.items()}
    dem = Counter(b for v in bv_trang.values() for b in v)
    for kt, t in trang.items():
        bv = bv_trang[kt]
        t["rev"] = "; ".join(f"{b} {s(dwg[b].get('Ngày issue hiệu lực'))} ({s(dwg[b].get('Đợt hiệu lực'))})" for b in bv)
        t["so_to"] = "; ".join(so_to_nhan(t["to"], dwg))
        t["so_cum"] = len(t["cum"])
        if prof:
            pg = res["PAGES"].get(kt, {})
            ten_dwg = [s(dwg[b].get("Tên tờ")) for b in bv if dem[b] == 1]
            t["loai"] = loai_to(prof, t["to"], pg.get("Loại tờ"), pg.get("Tên BV"), *ten_dwg)
    return {"tong": {"trang": len(trang), "luot": sum(t["so_cum"] for t in trang.values())}, "trang": trang}


def chia_trang(trang, max_trang, max_mon):
    """Mỗi trang đúng 1 nhóm. Xếp theo file rồi số trang (trang liền nhau thường cùng bộ tờ), cắt theo số trang và tổng lượt món."""
    def k(x):
        f, pg = x.split("|", 1)
        return f, int(pg) if pg.isdigit() else 0
    nhom, cur, mon = [], [], 0
    for t in sorted(trang, key=k):
        n = trang[t]["so_cum"]
        if cur and (len(cur) >= max_trang or mon + n > max_mon):
            nhom.append(cur)
            cur, mon = [], 0
        cur.append(t)
        mon += n
    if cur:
        nhom.append(cur)
    return nhom


def _dong_mon(c, scope, info, loai=None, pb=""):
    bb, dk = truong_tren_to(info["truong"], loai)
    r = scope.get(c, {})
    return (f"  - {c} [{info['profile']}] {ten_mon(r)}{f' ({pb})' if pb else ''} | ĐV {s(r.get('ĐV')) or '?'} | cần: {', '.join(bb) or '—'}"
            + (f" | nếu có: {', '.join(dk)}" if dk else ""))


def ds_mon(cs, scope, cmap, loai=None):
    """Danh sách món của 1 trang. ≥ 2 món giống tên (cùng profile, 40 ký tự đầu như nhau sau khi bỏ kích thước/mã hoàn thiện:
    'wall cladding with finish …' ốp gỗ vs ốp giấy) → in thêm Mã VL / mã FF&E để agent phân biệt, thay vì ghi Cụm ứng viên cho cả hai."""
    cs = [c for c in cs if c in cmap]
    ten = {c: (cmap[c].get("profile"), _fold(ten_mon(scope.get(c, {})))[:40]) for c in cs}
    dem = Counter(ten.values())
    ma = {c: " · ".join(x for x in (s(scope.get(c, {}).get("Mã VL")), s(scope.get(c, {}).get("Mã spec/FF&E"))) if x) for c in cs}
    dem_ma = Counter((ten[c], ma[c]) for c in cs)
    out = []
    for c in cs:
        pb = ""
        if dem[ten[c]] > 1:
            pb = f"trùng tên, phân biệt: {ma[c] or 'không có mã'}"
            if dem_ma[(ten[c], ma[c])] > 1:
                pb += " — trùng cả mã (khác dòng/bộ tờ BOQ): cùng 1 đồ vật thì ghi các cụm vào Cụm đọc, cách '; '"
        out.append(_dong_mon(c, scope, cmap[c], loai, pb))
    return out


def cmd_docsau(goc, args):
    p = td.P(goc).need()
    loc_phong, loc_sheet = opt(args, "--phong"), opt(args, "--sheet")
    max_t, max_mon = int(opt(args, "--trang", str(TRANG_NHOM))), int(opt(args, "--mon", str(MON_NHOM)))
    n_tim, n_dc = int(opt(args, "--nhom", str(CUM_TIM))), int(opt(args, "--doichieu", str(CUM_DOICHIEU)))
    prof = nap_profile()
    res, _ = merge(p)
    scope = res["SCOPE"]
    if not scope:
        sys.exit("LỖI: chưa có SCOPE — chạy boq trước.")
    chua_tt = [k for k, r in scope.items() if td.can_ma(p, r) and not s(r.get("Trạng thái"))]
    cum, bo = OrderedDict(), Counter()
    theo_khoa = {}
    for k, r in scope.items():
        if not can_doc(p, r):
            bo[s(r.get("Trạng thái")) if td.can_ma(p, r) else "sheet/dòng không có đơn vị"] += 1
            continue
        kc = khoa_cum(r)
        rep = theo_khoa.setdefault(kc, k)
        cum.setdefault(rep, []).append(k)
    sc_rows, map_rows, nhan = [], [], Counter()
    for rep, ds in cum.items():
        for k in ds:
            r = scope[k]
            row = {"Khóa": k, "Cụm đọc": rep}
            nv = scope_phan_loai(s(r.get("Mô tả BOQ")))
            if nv:
                row["Phạm vi nhà máy"] = nv
            nhan[(nv or "(LLM quyết)").split("–")[0].strip()] += 1
            sc_rows.append(row)
            ph = tach_phong(r.get("Mã phòng")) or [(KHU_CA_DONG, r.get("KL BOQ"))]
            for khu, kl in ph:
                map_rows.append({"Khóa KL": f"{k}|{khu}", "Khóa": k, "Khu": khu, "KL BOQ": kl})
    write_tool(p, "SCOPE", "docsau", sc_rows)
    write_tool(p, "BOQ_MAP", "docsau", map_rows)
    # chọn phần giao: mặc định mọi cụm; lọc khi chạy lại
    chon = list(cum)
    if loc_phong:
        chon = [c for c in chon if any(loc_phong in s(scope[k].get("Mã phòng")) for k in cum[c])]
    if loc_sheet:
        chon = [c for c in chon if any(s(scope[k].get("Sheet")) == loc_sheet for k in cum[c])]
    if "--thaydoi" in args:
        ah = set(((jload(p.f("_thaydoi.json")) or {}).get("anh_huong") or {}).get("SCOPE", {}))
        if not ah:
            print("XEM: _thaydoi.json chưa có dòng bị ảnh hưởng, chạy lệnh thaydoi trước.")
        chon = [c for c in chon if ah & set(cum[c])]
    if "--thieu" in args:
        qs = co_quan_sat(res)
        chon = [c for c in chon if not any(r.get("Nguồn") in ("HÌNH", "TEXT") for r in qs.get(c, []))]
    # hợp đồng kiểm từng cụm (profile theo Mã cấp 3) + manifest trang → món → trường
    hd = OrderedDict()
    for c, ds in cum.items():
        ten = profile_cum(prof, scope[c])
        hd[c] = {"khoa": ds, "profile": ten, "truong": hop_dong(prof, ten, scope[c].get("ĐV"))}
    man = lap_manifest(res, scope, cum, chon, prof, re.compile(p.cfg.get("dwg_regex", td.RX_DWG)))
    co_trang = {k: t for k, t in man["trang"].items() if t["trang"] is not None}
    nhom_t = chia_trang(co_trang, max_t, max_mon)
    tim = [c for c in chon if not any(t in co_trang for t in to_cua(scope, c))]
    nhom_tim = [tim[i:i + n_tim] for i in range(0, len(tim), n_tim)]
    nhom_dc = [chon[i:i + n_dc] for i in range(0, len(chon), n_dc)]
    # số nhóm nối tiếp: trang 1..A, tìm tờ A+1..B, đối chiếu B+1.. → file *.docsau-<n>.json không đè nhau
    so_t = list(range(1, len(nhom_t) + 1))
    so_tim = list(range(len(so_t) + 1, len(so_t) + len(nhom_tim) + 1))
    so_dc = list(range(len(so_t) + len(so_tim) + 1, len(so_t) + len(so_tim) + len(nhom_dc) + 1))
    for n, ds in zip(so_t, nhom_t):
        for t in ds:
            man["trang"][t]["chu"] = n
    man["phien_ban"] = PHIEN_BAN
    jdump(p.f("_doc_trang.json"), man)
    jdump(p.f("_docsau.json"), {"phien_ban": PHIEN_BAN, "profile": f"{prof.get('phien_ban')} ({prof.get('trang_thai')})", "cum": hd,
                                "trang_nhom": dict(zip(map(str, so_t), nhom_t)), "tim_to": dict(zip(map(str, so_tim), nhom_tim)),
                                "so_nhom": dict(zip(map(str, so_dc), nhom_dc))})

    dong = []
    for n, ds in zip(so_t, nhom_t):
        cac = {c for t in ds for c in man["trang"][t]["cum"]}
        dong.append(f"== Trang · Nhóm {n} ({len(ds)} trang, {len(cac)} cụm, {sum(man['trang'][t]['so_cum'] for t in ds)} lượt món) "
                    f"— đọc ảnh mỗi trang 1 lần, ghi TRANG_QS.docsau-{n}.json + PAGES.docsau-{n}.json")
        for t in ds:
            m = man["trang"][t]
            dong.append(f"TRANG {t} | {m['to']} | loại tờ: {'+'.join(m['loai']) or '? (chưa rõ: nhìn ảnh cả trang rồi xemtrang --loai)'}"
                        + (f" | {m['rev']}" if m["rev"] else ""))
            dong += ds_mon(m["cum"], scope, hd, m["loai"])
    for n, ds in zip(so_tim, nhom_tim):
        dong.append(f"== Tìm tờ · Nhóm {n} ({len(ds)} cụm chưa có trang) — tìm tờ bằng ungvien/tim rồi đọc như nhóm trang, ghi TRANG_QS.docsau-{n}.json + PAGES.docsau-{n}.json")
        for c in ds:
            dong.append(_dong_mon(c, scope, hd[c]) + f" | Bản vẽ hiệu lực: {s(_txt(scope[c].get('Bản vẽ hiệu lực'))) or '(trống)'}")
    for n, ds in zip(so_dc, nhom_dc):
        dong.append(f"== Đối chiếu · Nhóm {n} ({len(ds)} cụm) — chạy SAU khi mọi nhóm trang/tìm tờ xong và đã chạy ghep; không mở ảnh trừ chỗ ghep/check báo thiếu")
        for c in ds:
            r = scope[c]
            dong.append(f"{c} [{len(cum[c])} dòng] [{hd[c]['profile']}] {s(r.get('ĐV'))} | {s(r.get('Mô tả BOQ'))[:140]} | VL: {s(r.get('Mã VL'))} {s(r.get('Mã spec/FF&E'))}")
    with open(p.f("_docsau.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(dong) + "\n")
    pl = Counter(hd[c]["profile"] for c in chon)
    print(f"Đọc sâu: {sum(len(v) for v in cum.values())} dòng → {len(cum)} cụm; giao lần này {len(chon)} cụm. Profile: {dict(pl)}")
    print(f"Đọc trang: {len(co_trang)} trang → {len(nhom_t)} nhóm (≤ {max_t} trang, ≤ {max_mon} lượt món/nhóm), mỗi trang đúng 1 nhóm; "
          f"tìm tờ: {len(tim)} cụm → {len(nhom_tim)} nhóm; đối chiếu: {len(nhom_dc)} nhóm (≤ {n_dc} cụm) -> {p.f('_docsau.txt')}")
    print(f"Manifest -> {p.f('_doc_trang.json')}: {man['tong']['luot']} lượt món trên {man['tong']['trang']} trang; "
          f"loại tờ chưa rõ: {sum(1 for t in co_trang.values() if not t['loai'])} trang; "
          f"có Rev từ sổ DWG: {sum(1 for t in co_trang.values() if t['rev'])}, có số tờ trong nhãn: {sum(1 for t in co_trang.values() if t['so_to'])}")
    print(f"Không đọc sâu: {dict(bo) or 'không'}. Nhãn Phạm vi nhà máy: {dict(nhan)}")
    if pl.get(CHUA_PL):
        print(f"XEM: {pl[CHUA_PL]} cụm CHƯA PHÂN LOẠI (thiếu Mã cấp 3) → dùng trường chung; nên phân loại xong bước ben/doichieu trước.")
    if chua_tt:
        print(f"XEM: {len(chua_tt)} dòng chưa có Trạng thái, vẫn đọc sâu; nên chạy xong bước 7 trước để khỏi đọc dòng Omit/Loại trừ.")
    khong_so = [k for k, t in man["trang"].items() if t["trang"] is None]
    if khong_so:
        print(f"XEM: {len(khong_so)} link Bản vẽ hiệu lực không có số trang (cả file) — cụm chỉ có link đó vào nhóm tìm tờ.")


# ---------------------------------------------------------------- ghep: TRANG_QS → QUAN_SAT / SCOPE / BOQ_MAP / DAY_TUONG (tool, không LLM)
def _so_dau(v):
    m = re.search(r"-?\d+(?:[.,]\d+)?", s(v))
    return float(m.group(0).replace(",", ".")) if m else None


def _ds_cum(v, cmap):
    return [x for x in dict.fromkeys(y.strip() for y in re.split(r"[;,]", s(v))) if x in cmap]


def trang_gem_file(p, kt):
    return p.f(os.path.join("_gemini_so", "trang", re.sub(r'[\\/:*?"<>|]', "_", kt) + ".json"))


def thieu_hop_dong(res, cmap):
    """{cụm: [trường BB chưa có kết luận]} — tính sống từ QUAN_SAT + BOQ_MAP, dùng chung cho ghep, check, tincay, xemcum."""
    qs, bm_c, out = co_quan_sat(res), {}, {}
    for x in res["BOQ_MAP"].values():
        bm_c.setdefault(s(x.get("Khóa")), []).append(x)
    for c, info in cmap.items():
        co = {s(r.get("Trường")) for r in qs.get(c, []) if s(r.get("Trạng thái"))}
        bm = [x for k in info.get("khoa", []) for x in bm_c.get(k, [])]
        if bm and all(x.get("KL theo BV") not in (None, "") for x in bm):
            co.add("KL theo BV")
        miss = [f for f, x in (info.get("truong") or {}).items() if x.get("muc") == BB and f not in co]
        if miss:
            out[c] = miss
    return out


def cmd_ghep(goc):
    """Ghép quan sát theo trang về từng cụm. LLM không chép TRANG_QS sang QUAN_SAT nữa."""
    p = td.P(goc).need()
    ds = jload(p.f("_docsau.json")) or {}
    if ds.get("phien_ban") != PHIEN_BAN:
        sys.exit("LỖI: _docsau.json chưa phải bản đọc theo trang — chạy docsau trước.")
    cmap = ds["cum"]
    man = (jload(p.f("_doc_trang.json")) or {}).get("trang", {})
    for f in ("QUAN_SAT", "SCOPE", "BOQ_MAP", "DAY_TUONG"):  # kết quả ghep cũ không được lẫn vào lần ghép mới
        if os.path.exists(p.f(f"{f}.ghep-tool.json")):
            os.remove(p.f(f"{f}.ghep-tool.json"))
    res, _ = merge(p)
    scope, bmap, tq = res["SCOPE"], res["BOQ_MAP"], res["TRANG_QS"]
    da_doc = {s(r.get("Khóa trang")) for r in tq} | set(res["PAGES"])
    theo, phu, chua_ghep, la, ngoai_ds, dc_rong = {}, set(), [], Counter(), 0, []
    for r in tq:
        kt = s(r.get("Khóa trang"))
        cs, uv = _ds_cum(r.get("Cụm đọc"), cmap), _ds_cum(r.get("Cụm ứng viên"), cmap)
        if not cs and s(r.get("Trường")) == "Dẫn chiếu":
            dc_rong.append(r)  # mặt cắt tường... chưa gán cụm: không được rơi mất, xét riêng bên dưới
        for x in re.split(r"[;,]", s(r.get("Cụm đọc"))):
            if x.strip() and x.strip() not in cmap:
                la[x.strip()] += 1
        phu |= {(kt, c) for c in cs + uv}
        if cs:
            for c in cs:
                theo.setdefault(c, []).append(r)
        elif uv:
            chua_ghep.append({"trang": kt, "mon": s(r.get("Món")), "truong": s(r.get("Trường")), "ung_vien": uv})
        elif s(r.get("Món")) not in ("", "(không có)"):
            ngoai_ds += 1
    nhan = {kt: (t["to"], t["fid"], t["trang"]) for kt, t in man.items()}

    def lk(kt):
        if kt in nhan:
            a, f, pg = nhan[kt]
            return [a, f, pg] if pg is not None else [a, f]
        f, _, pg = kt.partition("|")
        pr = res["PAGES"].get(kt, {})
        return [s(pr.get("Số BV")) or kt, f, int(pg)] if pg.isdigit() else [kt, f]

    rx_bv = re.compile(p.cfg.get("dwg_regex", td.RX_DWG))

    def so_bv(t):  # số BV theo regex dự án + mẫu ngắn 'FUR-1512' (số tờ trong bộ)
        t = s(t).upper()
        return {m if isinstance(m, str) else m[0] for m in rx_bv.findall(t)} | set(re.findall(r"(?<![A-Z0-9-])[A-Z]{1,5}-\d{3,5}[A-Z]?(?![0-9])", t))

    qs_rows, sc_rows, bm_rows, dt_rows = [], [], [], []
    kl_mt, kl_lech, dan_chieu, gem_cum = [], [], {}, 0
    tool = {"Nguồn": "TOOL", "Model": "td.py ghep"}
    for c, info in cmap.items():
        rows = theo.get(c, [])
        cua = to_cua(scope, c)
        if not rows and not any(kt in da_doc for kt in cua):
            continue  # cụm chưa ai đọc: không sinh dòng tự động, check vẫn báo "chưa có dòng QUAN_SAT"
        moi, gem = [], []
        for r in rows:
            tr, kt = s(r.get("Trường")), s(r.get("Khóa trang"))
            if tr == "Kiểm chéo Gemini":
                gem.append(r)
                continue
            x = {"Cụm đọc": c, "Trường": tr, "Giá trị": r.get("Giá trị"), "Trạng thái": r.get("Trạng thái"),
                 "Tờ · Chi tiết": r.get("Tờ · Chi tiết"), "Nguồn": r.get("Nguồn"), "Model": f"ghep ← {kt}"}
            if s(r.get("Thành phần")):
                x["Thành phần"] = r.get("Thành phần")
            if s(r.get("Khu")) and tr != "Khu":
                x["Giá trị"] = f"{s(r.get('Giá trị'))} @ {s(r.get('Khu'))}"
            moi.append(x)
        # bộ tờ: trang của cụm đã đọc + trang có quan sát gán cụm; tờ chưa đọc: trang của cụm chưa ai đọc + số BV dẫn chiếu ngoài bộ
        doc = list(dict.fromkeys([kt for kt in cua if kt in da_doc] + [s(r.get("Khóa trang")) for r in rows]))
        chua = [kt for kt in cua if kt not in da_doc]
        biet = {b for kt in cua + doc for b in so_bv(s(lk(kt)[0]) + " " + s(res["PAGES"].get(kt, {}).get("Số BV")))}
        dc = sorted({b for r in rows if s(r.get("Trường")) == "Dẫn chiếu" for b in so_bv(r.get("Giá trị"))} - biet)
        if dc:
            dan_chieu[c] = dc
        moi.append(dict(tool, **{"Cụm đọc": c, "Trường": "Bộ tờ đã đọc", "Trạng thái": "CÓ" if doc else "KHÔNG THẤY",
                                 "Giá trị": "; ".join(s(lk(kt)[0]) for kt in doc) or "KHÔNG THẤY – chưa trang nào của cụm được đọc",
                                 **({"Tờ · Chi tiết": [lk(kt) for kt in doc]} if doc else {})}))
        cd = [s(lk(kt)[0]) for kt in chua] + [f"{b} (dẫn chiếu, ngoài bộ tờ)" for b in dc]
        moi.append(dict(tool, **{"Cụm đọc": c, "Trường": "Tờ chưa đọc", "Trạng thái": "CÓ", "Giá trị": "; ".join(cd) or "không còn"}))
        if gem:
            tt = [s(r.get("Trạng thái")) for r in gem]
            moi.append({"Cụm đọc": c, "Trường": "Kiểm chéo Gemini", "Nguồn": "GEMINI", "Model": "td.py ghep",
                        "Trạng thái": "KHÔNG THẤY" if all(t == "KHÔNG THẤY" for t in tt) else ("MÂU THUẪN" if "MÂU THUẪN" in tt else "CÓ"),
                        "Giá trị": " | ".join(f"{s(r.get('Khóa trang'))}: {s(r.get('Giá trị'))}" for r in gem)})
        qs_rows += moi
        # Gemini số của cụm = số Gemini đọc trên các trang cụm có Kích thước BV (gemini --trang), cộng dồn với --cum cũ
        tr_kt = {s(r.get("Khóa trang")) for r in rows if s(r.get("Trường")) == "Kích thước BV" and r.get("Nguồn") in ("HÌNH", "TEXT")}
        crops = {}
        for kt in sorted(tr_kt):
            crops.update((jload(trang_gem_file(p, kt)) or {}).get("crops", {}))
        if crops:
            f = gem_file(p, c)
            cu = jload(f) or {}
            jdump(f, {"cum": c, "crops": dict(cu.get("crops", {}), **crops), "tu_trang": sorted(tr_kt)})
            gem_cum += 1
        # SCOPE dòng đại diện: Kích thước BV đầu tiên đọc từ hình/chữ + Bản vẽ đối chứng
        if c in scope and doc:
            kt_r = next((r for r in rows if s(r.get("Trường")) == "Kích thước BV" and r.get("Nguồn") in ("HÌNH", "TEXT")
                         and r.get("Trạng thái") in ("CÓ", "MÂU THUẪN")), None)
            dcl, seen = [], set()
            for l in (links_of(kt_r.get("Tờ · Chi tiết")) if kt_r else []) + [lk(kt) for kt in doc]:
                if tuple(map(str, l)) not in seen:
                    seen.add(tuple(map(str, l)))
                    dcl.append(l)
            sc_rows.append({"Khóa": c, "Bản vẽ đối chứng": dcl, **({"Kích thước BV": kt_r.get("Giá trị")} if kt_r else {})})
        # BOQ_MAP: KL đếm theo khu; cùng khu ở nhiều trang → lấy 1 lần nếu bằng nhau, khác nhau = mâu thuẫn (không cộng)
        kl = {}
        for r in rows:
            if s(r.get("Trường")) != "KL theo BV" or s(r.get("Trạng thái")) != "CÓ":
                continue
            v = _so_dau(r.get("Giá trị"))
            if v is None:
                continue
            khu = s(r.get("Khu")) or KHU_CA_DONG
            d = kl.setdefault(_fold(khu), {"khu": khu, "v": {}, "l": []})
            d["v"][s(r.get("Khóa trang"))] = v
            d["l"] += links_of(r.get("Tờ · Chi tiết")) or [lk(s(r.get("Khóa trang")))]
        bm_c = [x for x in bmap.values() if s(x.get("Khóa")) in info["khoa"]]
        theo_khu = {}
        for x in bm_c:
            theo_khu.setdefault(_fold(s(x.get("Khu"))), []).append(x)
        con = []
        for kf, d in kl.items():
            if len(set(d["v"].values())) > 1:
                kl_mt.append({"cum": c, "khu": d["khu"], "trang": d["v"]})
                continue
            v = next(iter(d["v"].values()))
            dst = theo_khu.get(kf, [])
            if len(dst) == 1:
                bm_rows.append({"Khóa KL": dst[0]["Khóa KL"], "KL theo BV": v, "Tờ đếm KL": d["l"]})
            else:
                con.append((d["khu"], v, d["l"], len(dst)))
        ca = theo_khu.get(_fold(KHU_CA_DONG), [])
        if con and len(ca) == 1 and len(bm_c) == 1 and not any(x[3] for x in con):  # BOQ không chia khu: cộng các khu khác nhau
            bm_rows.append({"Khóa KL": ca[0]["Khóa KL"], "KL theo BV": sum(x[1] for x in con), "Tờ đếm KL": [l for x in con for l in x[2]]})
        else:
            kl_lech += [{"cum": c, "khu": x[0], "kl": x[1], "ly_do": "khu trùng nhiều dòng BOQ" if x[3] else "khu không có trong BOQ_MAP của cụm"} for x in con]
        for r in rows:
            if s(r.get("Trường")) == "Dày tường":
                vt = s(r.get("Khu")) or s(r.get("Chi tiết")) or s(r.get("Khóa trang"))
                dt_rows.append({"Khóa DT": f"{c}|{vt}", "Cụm đọc": c, "Vị trí": vt, "Dày hoàn thiện (mm)": s(r.get("Giá trị")),
                                "Tờ mặt bằng": links_of(r.get("Tờ · Chi tiết")) or [lk(s(r.get("Khóa trang")))], "Ghi chú": f"ghep ← {s(r.get('Khóa trang'))}"})
    # dẫn chiếu chưa gán cụm: số BV đó thuộc bộ tờ của cụm nào, đã có trang nào đọc chưa
    bv_cum, bv_doc = {}, set()
    for c in cmap:
        for kt in to_cua(scope, c):
            for b in so_bv(s(lk(kt)[0]) + " " + s(res["PAGES"].get(kt, {}).get("Số BV"))):
                bv_cum.setdefault(b, [])
                if c not in bv_cum[b]:
                    bv_cum[b].append(c)
    for kt in da_doc:
        bv_doc |= so_bv(s(lk(kt)[0]) + " " + s(res["PAGES"].get(kt, {}).get("Số BV")))
    dc_khong_cum = []
    for r in dc_rong:
        bs = sorted(so_bv(r.get("Giá trị")))
        dc_khong_cum.append({"trang": s(r.get("Khóa trang")), "mon": s(r.get("Món")), "gia_tri": s(r.get("Giá trị")),
                             "ung_vien": _ds_cum(r.get("Cụm ứng viên"), cmap),
                             "so_bv": {b: {"cum": bv_cum.get(b, [])[:6], "da_doc": b in bv_doc} for b in bs}})
    dc_chua_doc = sorted({b for x in dc_khong_cum for b, v in x["so_bv"].items() if not v["da_doc"]})
    # trang đã đọc mà món được giao trên trang không có dòng nào (kể cả KHÔNG THẤY) = bỏ sót
    thieu_trang = {}
    for kt, t in man.items():
        if kt in da_doc:
            sot = [c for c in t["cum"] if (kt, c) not in phu]
            if sot:
                thieu_trang[kt] = sot
    write_tool(p, "QUAN_SAT", "ghep", qs_rows)
    write_tool(p, "SCOPE", "ghep", sc_rows)
    write_tool(p, "BOQ_MAP", "ghep", bm_rows)
    write_tool(p, "DAY_TUONG", "ghep", dt_rows)
    res2, _ = merge(p)
    thieu = thieu_hop_dong(res2, cmap)
    import datetime
    jdump(p.f("_ghep.json"), {"luc": datetime.datetime.now().isoformat(timespec="seconds"), "thieu_truong": thieu, "thieu_trang": thieu_trang,
                              "chua_ghep": chua_ghep, "kl_mau_thuan": kl_mt, "kl_khong_khop": kl_lech, "dan_chieu_ngoai": dan_chieu,
                              "dan_chieu_khong_cum": dc_khong_cum, "dan_chieu_khong_cum_chua_doc": dc_chua_doc,
                              "cum_la": dict(la), "ngoai_danh_sach": ngoai_ds})
    print(f"Ghép: {len(tq)} dòng TRANG_QS → {len(qs_rows)} dòng QUAN_SAT ({len(theo)} cụm có quan sát / {len(cmap)}); "
          f"SCOPE {len(sc_rows)}, BOQ_MAP {len(bm_rows)} khu có KL, DAY_TUONG {len(dt_rows)}; Gemini số {gem_cum} cụm")
    print(f"Thiếu trường bắt buộc: {len(thieu)} cụm; món được giao mà trang không ghi gì: {sum(map(len, thieu_trang.values()))} ở {len(thieu_trang)} trang; "
          f"chưa gán cụm (chỉ có Cụm ứng viên): {len(chua_ghep)}; KL mâu thuẫn giữa trang: {len(kl_mt)}; KL không khớp khu: {len(kl_lech)}; "
          f"số BV dẫn chiếu ngoài bộ tờ: {sum(map(len, dan_chieu.values()))}; dẫn chiếu chưa gán cụm: {len(dc_khong_cum)} dòng "
          f"({len(dc_chua_doc)} số BV chưa trang nào đọc); quan sát ngoài danh sách món: {ngoai_ds} -> {p.f('_ghep.json')}")
    if la:
        print(f"XEM: Cụm đọc không có trong _docsau.json: {', '.join(list(la)[:8])}")


# ---------------------------------------------------------------- gemini
def gem_file(p, cum):
    return p.f(os.path.join("_gemini_so", re.sub(r'[\\/:*?"<>|]', "_", cum) + ".json"))


def claude_kt(rows):
    """Số Claude đã ghi của 1 cụm: chuỗi R×S×C đầu tiên mỗi dòng 'Kích thước BV' nguồn HÌNH/TEXT."""
    return sorted({x for r in rows if r.get("Trường") == "Kích thước BV" and r.get("Nguồn") in ("HÌNH", "TEXT") and r.get("Trạng thái") in ("CÓ", "MÂU THUẪN")
                   for x in neo_lib.neo_kt_gt(r.get("Giá trị"))}, key=float)


def gemini_so(p, cum, rows):
    """So số Claude với số Gemini trên các crop của cụm. Một chiều: số Claude phải có trong số Gemini; Gemini thừa là số khác trên crop.
    Chỉ so theo tập số: không biết số thuộc chiều nào. -> {tt: KHÔNG CÓ SỐ | CHƯA CHẠY | LỖI | KHỚP | LỆCH, khop, thieu, crops}"""
    mine = claude_kt(rows)
    if not mine:
        return dict(tt="KHÔNG CÓ SỐ", khop=[], thieu=[], crops=[])
    crops = (jload(gem_file(p, cum)) or {}).get("crops", {})
    if not crops:
        return dict(tt="CHƯA CHẠY", khop=[], thieu=mine, crops=[])
    g = {x for c in crops.values() for x in (c.get("so") or [])}
    thieu, loi = [x for x in mine if x not in g], any(c.get("so") is None for c in crops.values())
    return dict(tt=("LỖI" if loi else "LỆCH") if thieu else "KHỚP", khop=[x for x in mine if x in g], thieu=thieu, crops=sorted(crops))


def gemini_so_dong(cum, kq):
    crop = "; crop: " + ", ".join(kq["crops"]) if kq["crops"] else ""
    gt = {"CHƯA CHẠY": "chưa chạy gemini --cum", "LỖI": f"agy lỗi ở một số crop, chưa so được: {', '.join(kq['thieu'])}",
          "KHỚP": f"khớp: {', '.join(kq['khop'])} (so theo tập số, chưa kiểm số thuộc chiều nào)",
          "LỆCH": f"chỉ Claude: {', '.join(kq['thieu'])}; khớp: {', '.join(kq['khop']) or 'không'}"}[kq["tt"]]
    return {"Cụm đọc": cum, "Trường": GEM_SO_TRUONG, "Giá trị": gt + crop, "Trạng thái": {"KHỚP": "CÓ", "LỆCH": "MÂU THUẪN"}.get(kq["tt"], "KHÔNG THẤY"), "Nguồn": "GEMINI", "Model": "td.py gemini"}


def cmd_gemini(args):
    pngs = [a for a in args if a.lower().endswith(".png")]
    if not pngs:
        sys.exit("LỖI: gemini cần ít nhất 1 file .png")
    hoi, mine, par = opt(args, "--hoi"), opt(args, "--claude"), int(opt(args, "--par", "4"))
    vat, cum, goc = opt(args, "--quansat"), opt(args, "--cum"), opt(args, "--goc")  # opt lấy cờ ra khỏi args: chỉ gọi 1 lần mỗi cờ
    trang = opt(args, "--trang")
    p = qs_cum = None
    if (cum or trang) and (hoi or vat or not goc):
        sys.exit("--cum / --trang chỉ dùng ở chế độ số và cần --goc <GOC>")
    if cum and trang:
        sys.exit("--cum và --trang: chọn 1 (đọc theo trang dùng --trang <FID>|<trang>)")
    if goc:  # có --goc: mọi agent dùng chung cache _data\_gemini_cache (theo nội dung ảnh + câu hỏi)
        p = td.P(goc).need()
    if cum:
        qs_cum = [r for r in co_quan_sat(merge(p)[0]).get(cum, []) if r.get("Trường") != GEM_SO_TRUONG]
        if not claude_kt(qs_cum):
            sys.exit(f"{cum}: QUAN_SAT chưa có 'Kích thước BV' dạng R×S×C (nguồn HÌNH/TEXT). Ghi số của Claude vào QUAN_SAT.docsau-<nhóm>.json trước rồi mới chạy Gemini.")
    if trang:  # độc lập: Claude ghi số của trang vào TRANG_QS trước, Gemini đọc sau
        qs_cum = [r for r in merge(p)[0]["TRANG_QS"] if s(r.get("Khóa trang")) == trang]
        if not claude_kt(qs_cum):
            sys.exit(f"{trang}: TRANG_QS chưa có 'Kích thước BV' dạng R×S×C (nguồn HÌNH/TEXT) của trang này. Ghi số của Claude vào TRANG_QS.docsau-<nhóm>.json trước rồi mới chạy Gemini.")
    kq = gemini_doc.doc_anh(pngs, hoi, vat, par, p.f("_gemini_cache") if p else None)
    tat_ca = []
    for png, txt in kq:
        print(f"== {os.path.basename(png)}\n{gem_gon(txt, hoi, vat)}")
        if not hoi and not vat:
            doc = gemini_doc.gem_doc(txt)
            if doc is None:
                print("  (không đọc được JSON của Gemini)")
            else:
                tat_ca += doc
    if cum:
        f = gem_file(p, cum)
        crops = (jload(f) or {}).get("crops", {})
        for png, txt in kq:
            doc = gemini_doc.gem_doc(txt)
            crops[os.path.basename(png)] = {"so": None if doc is None else [v for x in doc if x != "?" for v in neo_lib.neo_tach_so(x)], "file": png}
        jdump(f, {"cum": cum, "crops": crops})
        print("-- " + gemini_so_dong(cum, gemini_so(p, cum, qs_cum))["Giá trị"])
    elif trang:  # lưu theo trang; ghep chia về các cụm có Kích thước BV trên trang
        f = trang_gem_file(p, trang)
        crops = (jload(f) or {}).get("crops", {})
        for png, txt in kq:
            doc = gemini_doc.gem_doc(txt)
            crops[os.path.basename(png)] = {"so": None if doc is None else [v for x in doc if x != "?" for v in neo_lib.neo_tach_so(x)], "file": png}
        jdump(f, {"trang": trang, "crops": crops})
        g = {x for c in crops.values() for x in (c.get("so") or [])}
        mine = claude_kt(qs_cum)
        print(f"-- {trang}: khớp {', '.join(x for x in mine if x in g) or 'không'}; chỉ Claude: {', '.join(x for x in mine if x not in g) or 'không'} "
              f"(so theo tập số cả trang; ghep chia về từng cụm)")
    elif mine and not hoi and not vat:
        a, g = Counter(x.strip() for x in mine.split(",") if x.strip()), Counter(x for x in tat_ca if x != "?")
        print(f"-- SO SÁNH: khớp {sum((a & g).values())} · chỉ Claude {sorted((a - g).elements())} · chỉ Gemini {sorted((g - a).elements())}")


def gem_gon(txt, hoi=None, vat=None):
    """In gọn trả lời Gemini (agent khỏi tự viết script tóm tắt log): số → 1 dòng mảng; quan sát → 1 dòng mỗi trường; lỗi / không JSON → nguyên văn."""
    if txt.startswith("LỖI") or hoi:
        return txt
    if not vat:
        doc = gemini_doc.gem_doc(txt)
        return txt if doc is None else "so: " + ", ".join(doc)
    m = re.search(r"\{.*\}", txt, re.S)
    try:
        d = json.loads(m.group(0)) if m else None
    except ValueError:
        d = None
    if not isinstance(d, dict):
        return txt
    dong = []
    for k, v in d.items():
        if isinstance(v, dict):
            gt, tt, o = s(v.get("value")), s(v.get("status")), s(v.get("where"))
            dong.append(f"  {k}: {gt} [{tt}]" + (f" @ {o}" if o else ""))
        else:
            dong.append(f"  {k}: {s(v)}")
    return "\n".join(dong)


# ---------------------------------------------------------------- xem gộp (plan Read Once A6)
# 1 lệnh in đủ thứ agent 7d cần cho 1 cụm / 1 trang, thay ~5 lượt loc / grep / find / đọc _docsau.txt (mỗi lượt đọc lại ~40k nền).
def _gt(v, n=0):
    t = re.sub(r"\s+", " ", s(_txt(v)))
    return t[:n] + "…" if n and len(t) > n else t


def _dong_gon(r, cot):
    return " | ".join(f"{c}: {_gt(r.get(c), 160)}" for c in cot if _gt(r.get(c)))


def _pdf(p, res, fid, nho):
    if fid not in nho:
        try:
            nho[fid] = td.link_item(p, ["", fid], res["FILE_INDEX"])[1] or "(không mở được)"
        except Exception as e:  # zip hỏng, file mất: in lý do, không dừng cả lệnh
            nho[fid] = f"(không mở được: {e})"
    return nho[fid]


def _in_tq(rows, cum=None):
    """In dòng TRANG_QS gọn; * = Cụm đọc có cụm này, ? = cụm này chỉ là ứng viên."""
    for r in rows:
        dau = " "
        if cum:
            dau = "*" if cum in _ds_cum(r.get("Cụm đọc"), {cum: 1}) else ("?" if cum in _ds_cum(r.get("Cụm ứng viên"), {cum: 1}) else " ")
        print(f"   {dau} [{_gt(r.get('Khóa trang'))}] CT {_gt(r.get('Chi tiết')) or '?'} · {_gt(r.get('Món'), 50)} · {_gt(r.get('Trường'))} = {_gt(r.get('Giá trị'), 140)}"
              + (f" @ {_gt(r.get('Khu'))}" if _gt(r.get("Khu")) else "")
              + f" [{_gt(r.get('Trạng thái'))}/{_gt(r.get('Nguồn'))}]" + (f" crop: {_gt(r.get('Crop'))}" if _gt(r.get("Crop")) else ""))


def _in_tq_gon(rows, cum):
    """TRANG_QS gom theo (trang, chi tiết, món, khu, crop): đầu khối in 1 lần, mỗi trường 1 dòng ngắn; nguồn HÌNH không in."""
    khoi = OrderedDict()
    for r in rows:
        dau = "*" if cum in _ds_cum(r.get("Cụm đọc"), {cum: 1}) else "?"
        k = (dau, _gt(r.get("Khóa trang")), _gt(r.get("Chi tiết")) or "?", _gt(r.get("Món"), 50), _gt(r.get("Khu")), _gt(r.get("Crop")))
        khoi.setdefault(k, []).append(r)
    for (dau, kt, ct, mon, khu, crop), rs in khoi.items():
        print(f"   {dau} [{kt}] CT {ct} · {mon}" + (f" @ {khu}" if khu else "") + (f" · crop: {crop}" if crop else ""))
        for r in rs:
            ng = _gt(r.get("Nguồn"))
            tt = _gt(r.get("Trạng thái"))
            print(f"       {_gt(r.get('Trường'))} = {_gt(r.get('Giá trị'), 140)}"
                  + ("" if tt == "CÓ" and ng == "HÌNH" else f" [{tt}{'' if ng == 'HÌNH' else '/' + ng}]"))


# Số trên trang trùng kích thước BOQ của cụm khác (vd len 150 dưới WP-04 3727 ở E3 ↔ BF-02A L3727×H150): máy báo, không trông LLM để ý.
_MA_SO = re.compile(r"\b[A-Za-z]{1,6}\d*[-_.]?\d*[A-Za-z]?[-_]\d+[A-Za-z]?\b|\b(?![LWHDRCØ]\d)[A-Z]{1,4}\d+\b|\b\d+(?:[.,]\d+)?\s*%")  # giữ L3727 / H150 / W600 (kích thước)


def _so_kt(v):
    """Số 3–5 chữ số ≥ 100 (mm) trong chữ, bỏ mã (FUR-1523, WD-01, SW3, F0004) và số thập phân / %."""
    t = _MA_SO.sub(" ", _gt(v))
    return {int(x) for x in re.findall(r"(?<![\d.,])\d{3,5}(?![\d]|[.,]\d)", t) if int(x) >= 100}


def _kt_boq_cum(scope, c):
    r = scope.get(c, {})
    return _so_kt(r.get("Kích thước BoQ")) or _so_kt(r.get("Mô tả BOQ"))


def trung_kt_cum_khac(res, cmap, rep, tq):
    """[(cụm khác, [(trang, số trùng, giá trị)])]: dòng TRANG_QS đã gán cụm rep có số trùng kích thước BOQ của cụm khác cùng sheet BOQ
    hoặc cùng trang. Chỉ báo khi phần trùng (bỏ số có trong kích thước BOQ của chính rep) có số lẻ (không chia hết 50) — 600/3000 quá phổ biến."""
    scope = res["SCOPE"]
    rieng = _kt_boq_cum(scope, rep)
    sheet = rep.split("!")[0]
    ra = OrderedDict()
    for x in tq:
        if rep not in _ds_cum(x.get("Cụm đọc"), cmap):
            continue
        so = _so_kt(x.get("Giá trị"))
        if not so:
            continue
        kt = _gt(x.get("Khóa trang"))
        ung = {c for c in cmap if c != rep and (c.split("!")[0] == sheet or kt in to_cua(scope, c))}
        ung |= {c for y in res["TRANG_QS"] if _gt(y.get("Khóa trang")) == kt for c in _ds_cum(y.get("Cụm đọc"), cmap) if c != rep}
        for o in sorted(ung):
            chung = so & _kt_boq_cum(scope, o)
            if any(v % 50 for v in chung - rieng):
                ra.setdefault(o, []).append((kt, sorted(chung), _gt(x.get("Giá trị"), 110)))
    return list(ra.items())


def _la_ghep(x):
    """Dòng QUAN_SAT do `ghep` sinh (chép từ TRANG_QS / tóm tắt bộ tờ, Gemini) — --gon không in lại."""
    return s(x.get("Model")).startswith(("ghep ←", "td.py ghep"))


def cmd_xemcum(goc, args):
    """xemcum <GOC> <Cụm đọc | Khóa dòng> […] [--gon] | xemcum <GOC> --nhom <N> [--gon]: dùng ở nhóm ĐỐI CHIẾU — mô tả BOQ, dòng + KL, MAT,
    tờ (đã đọc chưa), TRANG_QS của cụm, QUAN_SAT đã ghép, hợp đồng kiểm (trường bắt buộc còn thiếu), SCOPE, BOQ_MAP, ISSUES, phần ghep báo.
    --gon: TRANG_QS gom theo trang/món, bỏ QUAN_SAT do ghep chép lại, mô tả BOQ cắt ngắn. Nhiều cụm / --nhom = in liền 1 lần."""
    args = list(args)
    gon = "--gon" in args
    args = [a for a in args if a != "--gon"]
    n = opt(args, "--nhom")
    if not args and not n:
        sys.exit("xemcum <GOC> <Cụm đọc hoặc Khóa dòng> [… thêm cụm] [--gon]  |  xemcum <GOC> --nhom <N> [--gon]")
    p = td.P(goc).need()
    ds = jload(p.f("_docsau.json")) or {}
    cmap = ds.get("cum") or {}
    if n:
        if n not in (ds.get("so_nhom") or {}):
            sys.exit(f"LỖI: không có nhóm đối chiếu {n} trong _docsau.json (có: {', '.join(ds.get('so_nhom') or {}) or 'không'})")
        args += ds["so_nhom"][n]
    reps = []
    for k in args:
        rep = k if k in cmap else next((c for c, v in cmap.items() if k in v.get("khoa", [])), None)
        if not rep:
            sys.exit(f"LỖI: '{k}' không phải cụm / dòng trong _docsau.json — chạy docsau trước, khóa dạng '<sheet>!<dòng>'")
        reps.append(rep)
    import contextlib
    import io
    res, _ = merge(p)
    nho, da_in = {}, 0
    reps = list(dict.fromkeys(reps))
    for i, rep in enumerate(reps):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            _xemcum_1(p, ds, cmap, res, rep, gon, nho)
        t = buf.getvalue()
        if i and da_in + len(t) > XEMCUM_MAX:  # output Bash bị cắt ~30k ký tự → dừng, in lệnh gọi tiếp
            con = reps[i:]
            print(f"\n== DỪNG trước {len(con)} cụm (output sắp quá {XEMCUM_MAX} ký tự). Gọi tiếp:\n"
                  f"xemcum <GOC> {' '.join(chr(34) + c + chr(34) for c in con)}{' --gon' if gon else ''}")
            return
        print(("\n" if i else "") + t, end="")
        da_in += len(t)


def _xemcum_1(p, ds, cmap, res, rep, gon, nho):
    scope, info = res["SCOPE"], cmap[rep]
    khoa = info["khoa"]
    r = scope.get(rep, {})
    nhom = next((n for n, cs in (ds.get("so_nhom") or {}).items() if rep in cs), "?")
    print(f"== CỤM {rep} · {len(khoa)} dòng · profile {info.get('profile', '?')} · nhóm đối chiếu {nhom}")
    print(f"Mô tả BOQ: {_gt(r.get('Mô tả BOQ'), 300 if gon else 0)}")
    for c in ("ĐV", "Mã cấp 3", "Mã VL", "Mã spec/FF&E", "Ghi chú BOQ", "Trạng thái", "Nội dung cuối (VN)", "Phạm vi nhà máy"):
        if _gt(r.get(c)):
            print(f"{c}: {_gt(r.get(c), 300)}")
    print("-- dòng (Khóa | KL BOQ | Mã phòng):")
    for kk in khoa:
        x = scope.get(kk, {})
        print(f"  {kk} | {_gt(x.get('KL BOQ'))} {_gt(x.get('ĐV'))} | {_gt(x.get('Mã phòng'), 120)}")
    ma = [m.strip() for m in re.split(r"[;,]", s(r.get("Mã VL"))) if m.strip()]
    if ma:
        print("-- MAT:")
        for m in ma:
            x = res["MAT"].get(m)
            print(f"  {m}: " + (_dong_gon(x, ("Loại", "Mô tả", "NCC/Mã SP", "Trạng thái cuối")) if x else "(không có trong MAT)"))
    da = {s(x.get("Khóa trang")) for x in res["TRANG_QS"]} | set(res["PAGES"])
    to = links_of(r.get("Bản vẽ hiệu lực"))
    print(f"-- tờ ({len(to)}; [đã đọc] = có TRANG_QS/PAGES, đọc ở nhóm trang — không mở lại ảnh nếu không cần):")
    for l in to:
        kt = khoa_trang(l)[0]
        print(f"  {'[đã đọc]' if kt in da else '[CHƯA]  '} {l[0]} | {kt}")
    for fid in dict.fromkeys(l[1] for l in to if len(l) > 1):
        print(f"  PDF {fid}: (đã in ở cụm trên)" if fid in nho else f"  PDF {fid}: {_pdf(p, res, fid, nho)}")
    tq = [x for x in res["TRANG_QS"] if rep in _ds_cum(x.get("Cụm đọc"), {rep: 1}) or rep in _ds_cum(x.get("Cụm ứng viên"), {rep: 1})]
    print(f"-- TRANG_QS của cụm: {len(tq)} dòng (* = đã gán, ? = ứng viên chưa gán)")
    (_in_tq_gon if gon else _in_tq)(tq, rep)
    trung = trung_kt_cum_khac(res, cmap, rep, tq)
    if trung:
        print("-- TRÙNG SỐ BOQ CỤM KHÁC (máy dò số trên trang = kích thước BOQ cụm khác): xét có phải cùng 1 món không — tính trùng KL? gán nhầm cụm?")
        for o, ds_ in trung[:6]:
            ro = res["SCOPE"].get(o, {})
            print(f"  {o} [{_gt(ro.get('Kích thước BoQ')) or _gt(ro.get('Mô tả BOQ'), 70)}]:")
            for kt, chung, gt in ds_[:2]:
                print(f"     [{kt}] số {', '.join(map(str, chung))} ← {gt}")
            if len(ds_) > 2:
                print(f"     … +{len(ds_) - 2} dòng ({', '.join(dict.fromkeys(k for k, _, _ in ds_[2:]))})")
        if len(trung) > 6:
            print(f"  … +{len(trung) - 6} cụm: {', '.join(o for o, _ in trung[6:])}")
    qs = co_quan_sat(res).get(rep, [])
    qs_in = [x for x in qs if not _la_ghep(x)] if gon else qs
    print(f"-- QUAN_SAT (ghep + đối chiếu): {len(qs)} dòng"
          + (f" — --gon ẩn {len(qs) - len(qs_in)} dòng ghep (chép từ TRANG_QS ở trên), in {len(qs_in)} dòng đối chiếu/gán" if gon else ""))
    for x in qs_in:
        print("  " + _dong_gon(x, ("Thành phần", "Trường", "Giá trị", "Trạng thái", "Nguồn", "Tờ · Chi tiết")))
    if info.get("truong"):
        thieu = thieu_hop_dong(res, {rep: info}).get(rep, [])
        print(f"-- Hợp đồng kiểm [{info.get('profile')}]: bắt buộc {', '.join(f for f, x in info['truong'].items() if x.get('muc') == BB)}"
              f" | CÒN THIẾU: {', '.join(thieu) or 'không'}")
    sc = _dong_gon(r, ("Kích thước BoQ", "Kích thước BV", "Bản vẽ đối chứng"))
    if sc:
        print(f"-- SCOPE: {sc}")
    print("-- BOQ_MAP (Khóa KL | KL BOQ | KL theo BV | Tờ đếm KL):")
    for x in res["BOQ_MAP"].values():
        if s(x.get("Khóa")) in khoa:
            print(f"  {x.get('Khóa KL')} | {_gt(x.get('KL BOQ'))} | {_gt(x.get('KL theo BV')) or '-'} | {_gt(x.get('Tờ đếm KL'), 80) or '-'}")
    g = jload(p.f("_ghep.json")) or {}
    bao = [f"KL mâu thuẫn giữa trang: {x['khu']} {x['trang']}" for x in g.get("kl_mau_thuan", []) if x["cum"] == rep]
    bao += [f"KL không gán được: {x['khu']} = {x['kl']} ({x['ly_do']})" for x in g.get("kl_khong_khop", []) if x["cum"] == rep]
    bao += [f"dẫn chiếu ngoài bộ tờ: {', '.join(g.get('dan_chieu_ngoai', {}).get(rep, []))}"] if g.get("dan_chieu_ngoai", {}).get(rep) else []
    bao += [f"dẫn chiếu chưa gán cụm (cụm này là ứng viên): [{x['trang']}] {x['mon']} → {x['gia_tri']}"
            + "".join(f" | {b}: {'đã đọc' if v['da_doc'] else 'CHƯA ĐỌC'}" for b, v in x["so_bv"].items())
            for x in g.get("dan_chieu_khong_cum", []) if rep in x.get("ung_vien", [])]
    if bao:
        print("-- ghep báo:\n  " + "\n  ".join(bao))
    iss = [x for x in res["ISSUES"].values() if s(x.get("Cụm đọc")) == rep]
    if iss:
        print("-- ISSUES:")
        for x in iss:
            print(f"  {x.get('ID')}: {_gt(x.get('Nội dung'), 200)}")
    gs = gemini_so(p, rep, [x for x in qs if x.get("Trường") != GEM_SO_TRUONG])
    if gs["tt"] != "KHÔNG CÓ SỐ":
        print(f"-- Gemini số: {gemini_so_dong(rep, gs)['Giá trị']}")


def cmd_xemtrang(goc, args):
    """xemtrang <GOC> <FID> <trang> [--chu]: dùng ở nhóm TRANG — đường dẫn PDF, khung tên, lớp chữ, DANH SÁCH CẦN KIỂM (món + trường theo loại tờ,
    không có giá trị BOQ), TRANG_QS, PAGES."""
    chu = "--chu" in args
    args = [a for a in args if a != "--chu"]
    loai_ep = opt(args, "--loai")
    if len(args) < 2:
        sys.exit("xemtrang <GOC> <FID> <trang> [--chu] [--loai MB|MBT|MĐ|MC|CT|BTK (nối bằng +)]")
    p = td.P(goc).need()
    fid, tr = args[0], int(args[1])
    if loai_ep is not None:
        loai_ep = [x.strip() for x in loai_ep.split("+") if x.strip()]
        sai = [x for x in loai_ep if x not in nap_profile()["loai_to"]]
        if sai or not loai_ep:
            sys.exit(f"LỖI: --loai {'+'.join(sai) or '(trống)'} không có trong PROFILE.json, dùng: {', '.join(nap_profile()['loai_to'])}")
    kt = f"{fid}|{tr}"
    res, _ = merge(p)
    ng = Nguon(p, res)
    t = (ng._pages(fid) or [None] * tr)[tr - 1] if 1 <= tr <= ng.so_trang(fid) else None
    ds = jload(p.f("_docsau.json")) or {}
    print(f"== TRANG {kt} · {_pdf(p, res, fid, {})}")
    if t:
        print(f"Khung tên (máy đoán): {s(t.get('so')) or '?'} {s(t.get('ten'))} · lớp chữ: {s(t.get('lop'))} ({len(s(t.get('text')))} ký tự)")
    else:
        print(f"(chưa có chữ trang trong _text — chạy trang; file có {ng.so_trang(fid)} trang)")
    pg = res["PAGES"].get(kt)
    if pg:
        print("PAGES: " + _dong_gon(pg, ("Số BV", "Tên BV", "Rev", "Loại tờ", "Đã đọc bằng")))
    m = ((jload(p.f("_doc_trang.json")) or {}).get("trang") or {}).get(kt)
    if m:
        if m.get("so_to") and m.get("rev"):
            print(f"Số BV / Rev (từ nhãn + sổ DWG — chép vào PAGES, KHÔNG crop khung tên): {m['so_to']} · {m['rev']}")
        else:
            print(f"Số BV / Rev: sổ thiếu ({'không có số tờ trong nhãn' if not m.get('so_to') else 'số BV không có trong sổ DWG'}) → đọc khung tên trên ảnh cả trang, chỉ crop nếu không đọc được")
        cmap = ds.get("cum") or {}
        loai = loai_ep or m.get("loai") or []
        nguon = "agent đặt --loai" if loai_ep else "máy đoán"
        print(f"-- CẦN KIỂM: {m['so_cum']} món · nhóm chủ {m.get('chu', '?')} · loại tờ {'+'.join(loai) + f' ({nguon})' if loai else '? (chưa rõ — nhìn ảnh cả trang rồi chạy lại với --loai, vd --loai MĐ+MC+CT)'}"
              f" (cần = phải kết luận nếu tờ thể hiện; nếu có = ghi khi thấy):")
        for x in ds_mon(m["cum"], res["SCOPE"], cmap, loai):
            print(x)
    else:
        print("-- trang không có trong danh sách giao (_doc_trang.json): ghi món thấy được, Cụm ứng viên nếu đoán được")
    print("-- TRANG_QS:")
    _in_tq([r for r in res["TRANG_QS"] if s(r.get("Khóa trang")) == kt])
    if t and s(t.get("text")):
        x = re.sub(r"\s+", " ", s(t.get("text")))
        print(f"-- chữ trang{'' if chu or len(x) <= 600 else ' (600 ký tự đầu; --chu = cả trang)'}:\n{x if chu else x[:600]}")


LUU_SHEET = ("PAGES", "TRANG_QS", "ISSUES", "ASK_USER")


def _kt_dong(r):
    return s((r.get("_chung") or {}).get("Khóa trang")) if "_chung" in r else s(r.get("Khóa trang"))


def cmd_luu(goc, args):
    """luu <GOC> <nhóm> <FID>|<trang> [--them] < JSON (stdin): nhóm trang ghi kết quả 1 trang — thay Write/Edit file JSON (đo 16/11: Edit viết lại
    cả khối, ~25% output). stdin = {"PAGES": {…}, "TRANG_QS": [dòng | {"_chung": {…}, "dong": [{…}]}], "ISSUES": [...], "ASK_USER": [...]} (phần nào
    không có thì bỏ). Tool tự điền `Khóa trang` mọi dòng, PAGES `File ID`/`Trang`/`Số BV`/`Rev`/`Lớp text` (manifest + _text) nếu agent không ghi.
    Kiểm cột / giá trị như check; có lỗi → in lỗi, KHÔNG ghi gì. Mặc định THAY mọi dòng TRANG_QS cũ của trang (sửa = gửi lại cả trang);
    --them = nối thêm dòng (bổ sung sau Gemini). PAGES / ISSUES / ASK_USER: cùng khóa thì thay."""
    them = "--them" in args
    args = [a for a in args if a != "--them"]
    if len(args) < 2 or "|" not in args[1]:
        sys.exit("luu <GOC> <nhóm> <FID>|<trang> [--them] < JSON")
    p = td.P(goc).need()
    n, kt = args[0], args[1]
    fid, tr = kt.split("|", 1)
    try:
        tr = int(tr)
        raw = sys.stdin.buffer.read().decode("utf-8-sig")
        vao = json.loads(raw)
    except ValueError as e:
        sys.exit(f"LỖI: trang phải là số / JSON stdin hỏng: {e}")
    if not isinstance(vao, dict) or set(vao) - set(LUU_SHEET):
        sys.exit(f"LỖI: stdin phải là object với các khóa {', '.join(LUU_SHEET)} (có: {', '.join(vao) if isinstance(vao, dict) else type(vao).__name__})")
    res, _ = merge(p)
    m = ((jload(p.f("_doc_trang.json")) or {}).get("trang") or {}).get(kt) or {}
    ng = Nguon(p, res)
    t = (ng._pages(fid) or [None] * tr)[tr - 1] if 1 <= tr <= ng.so_trang(fid) else None
    pg = vao.get("PAGES") or {}
    if isinstance(pg, list):
        pg = pg[0] if len(pg) == 1 else {"_loi": "PAGES chỉ 1 dòng (trang này)"}
    cu_pg = next((r for r in (jload(p.f(f"PAGES.docsau-{n}.json")) or {}).get("rows") or [] if s(r.get("Khóa trang")) == kt), None)
    if "PAGES" in vao or not cu_pg:  # lần ghi đầu của trang luôn có PAGES; gửi lại không kèm PAGES thì giữ dòng cũ
        tu = {"Khóa trang": kt, "File ID": fid, "Trang": tr, "Số BV": m.get("so_to"), "Rev": m.get("rev"), "Lớp text": s((t or {}).get("lop")).upper()}
        pg = {**{c: v for c, v in tu.items() if v not in (None, "")}, **(cu_pg or {}), **pg}
        vao["PAGES"] = [pg]
    loi, moi = [], {}
    for sh in LUU_SHEET:
        ds = vao.get(sh) or []
        if isinstance(ds, dict):
            ds = [ds]
        d = td.SHEETS[sh]
        for r in ds:
            if not isinstance(r, dict):
                loi.append(f"{sh}: dòng phải là object: {r!r:.60}")
                continue
            if sh == "TRANG_QS":
                if "_chung" in r:
                    r["_chung"].setdefault("Khóa trang", kt)
                else:
                    r.setdefault("Khóa trang", kt)
        hang = td.bung_chung(ds) if sh == "TRANG_QS" else ds
        for i, r in enumerate(hang):
            tag = f"{sh} dòng {i + 1}" + (f" ({r.get('Món')} · {r.get('Trường')})" if sh == "TRANG_QS" else "")
            for c, v in r.items():
                x = d["cot"].get(c)
                if not x:
                    loi.append(f"{tag}: cột lạ '{c}'")
                elif "docsau" not in x["buoc"] and "*" not in x["buoc"] and not (d["loai"] == "KHOA" and c == d["khoa"]):
                    loi.append(f"{tag}: bước docsau không được ghi cột '{c}'")
                elif v != "(xóa)":
                    loi += [f"{tag}: {e}" for e in td.check_val(c, x, v, res["FILE_INDEX"])]
            if sh == "TRANG_QS":
                if s(r.get("Khóa trang")) != kt:
                    loi.append(f"{tag}: Khóa trang '{r.get('Khóa trang')}' ≠ trang đang lưu {kt}")
                loi += [f"{tag}: thiếu '{c}'" for c in ("Món", "Trường", "Trạng thái") if not r.get(c)]
                if r.get("Nguồn") in ("HÌNH", "TEXT") and not r.get("Tờ · Chi tiết"):  # cùng luật check: lỗi bắt ngay lúc ghi, khỏi chạy check
                    loi.append(f"{tag}: nguồn {r.get('Nguồn')} mà không ghi Tờ · Chi tiết")
                if s(r.get("Trường")) == "KL theo BV" and s(r.get("Trạng thái")) == "CÓ" and _so_dau(r.get("Giá trị")) is None:
                    loi.append(f"{tag}: 'KL theo BV' trạng thái CÓ mà Giá trị không có số")
                loi += [f"{tag}: chuỗi cộng sai '{ch}' (tổng thật {tong:g})" for _, _, ch, tong in kt_chuoi_sai([r])]
            elif d["loai"] == "KHOA" and not s(r.get(d["khoa"])):
                loi.append(f"{tag}: thiếu khóa '{d['khoa']}'")
        moi[sh] = ds
    if loi:
        print(f"LỖI — chưa ghi gì ({len(loi)} lỗi), sửa rồi gửi lại cả lệnh:")
        print("\n".join(loi[:30]))
        sys.exit(1)
    tom = []
    for sh, ds in moi.items():
        if not ds:
            continue
        path = p.f(f"{sh}.docsau-{n}.json")
        cu = (jload(path) or {}).get("rows") or []
        if sh == "TRANG_QS":
            if not them:  # bỏ dòng cũ của trang (nhóm _chung lẫn trang thì lọc từng dòng)
                giu = []
                for r in cu:
                    if "_chung" in r and isinstance(r.get("dong"), list):
                        kc = s(r["_chung"].get("Khóa trang"))
                        dong = [x for x in r["dong"] if s(x.get("Khóa trang")) != kt and not (kc == kt and "Khóa trang" not in x)]
                        if dong:
                            giu.append({**r, "dong": dong})
                    elif _kt_dong(r) != kt:
                        giu.append(r)
                cu = giu
            rows = cu + ds
        else:
            k = td.SHEETS[sh]["khoa"]
            thay = {s(r.get(k)) for r in ds}
            rows = [r for r in cu if s(r.get(k)) not in thay] + ds
        jdump(path, {"rows": rows})
        tom.append(f"{sh} {len(td.bung_chung(ds)) if sh == 'TRANG_QS' else len(ds)}")
    # độ phủ: món trong danh sách cần kiểm của trang chưa có dòng nào
    tq = [r for r in td.bung_chung((jload(p.f(f"TRANG_QS.docsau-{n}.json")) or {}).get("rows") or []) if s(r.get("Khóa trang")) == kt]
    cmap = (jload(p.f("_docsau.json")) or {}).get("cum") or {}
    co = {c for r in tq for c in _ds_cum(r.get("Cụm đọc"), cmap) + _ds_cum(r.get("Cụm ứng viên"), cmap)}
    thieu = [c for c in m.get("cum", []) if c not in co]
    print(f"Đã lưu {kt} → nhóm {n}: {', '.join(tom)} ({'nối thêm' if them else 'thay dòng cũ của trang'}); trang có {len(tq)} dòng TRANG_QS"
          + (f"\nXEM: {len(thieu)} món trong danh sách chưa có dòng nào (Cụm đọc/Cụm ứng viên): {'; '.join(thieu[:12])}" if thieu else ""))


LOAI_KHOI = {"Trang": "trang", "Tìm tờ": "tim", "Đối chiếu": "doichieu"}


def khoi_giao(p, n):
    """Khối `Trang · Nhóm n` / `Tìm tờ · Nhóm n` / `Đối chiếu · Nhóm n` trong _docsau.txt -> (loại, các dòng)."""
    try:
        with open(p.f("_docsau.txt"), encoding="utf-8") as fh:
            dong = fh.read().splitlines()
    except OSError:
        sys.exit("LỖI: chưa có _docsau.txt — chạy docsau trước")
    for i, x in enumerate(dong):
        m = re.match(rf"== ({'|'.join(LOAI_KHOI)}) · Nhóm {n} \(", x)
        if m:
            j = next((j for j in range(i + 1, len(dong)) if dong[j].startswith("== ")), len(dong))
            return LOAI_KHOI[m.group(1)], dong[i:j]
    sys.exit(f"LỖI: không có Nhóm {n} trong _docsau.txt (chạy lại docsau nếu _docsau.txt là bản cũ)")


def cmd_giao(goc, args):
    """giao <GOC> <nhóm>: in prompt giao việc bước 7d cho 1 nhóm — khối việc, cột được ghi, thư mục ảnh.
    Phiên điều phối dán nguyên văn vào subagent: agent khỏi đọc JSON.md, khỏi grep _docsau.txt, khỏi tìm chỗ để ảnh."""
    if not args:
        sys.exit("giao <GOC> <nhóm>")
    import contextlib
    import io
    p = td.P(goc).need()
    n = args[0]
    loai, khoi = khoi_giao(p, n)
    png = os.path.join(p.t, "_png", f"n{n}")
    os.makedirs(png, exist_ok=True)
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        td.cmd_cot(["--buoc", "docsau"])
    if loai == "doichieu":
        can = ("[SCOPE]", "[QUAN_SAT]", "[BOQ_MAP]", "[DAY_TUONG]", "[ISSUES]", "[ASK_USER]")
        ghi = f"SCOPE / QUAN_SAT / BOQ_MAP / DAY_TUONG / ISSUES / ASK_USER .docsau-{n}.json (chỉ phần cần sửa/bổ sung; KHÔNG chép lại quan sát đã ghép)"
        ten = "ĐỐI CHIẾU (không mở ảnh, trừ chỗ thiếu)"
    else:
        can = ("[TRANG_QS]", "[PAGES]", "[ISSUES]", "[ASK_USER]")
        ghi = f"TRANG_QS.docsau-{n}.json, PAGES.docsau-{n}.json, ISSUES.docsau-{n}.json (nếu có)"
        ten = "ĐỌC TRANG" if loai == "trang" else "TÌM TỜ + ĐỌC TRANG"
    cot = "\n\n".join("\n".join(x for x in b.strip().splitlines() if not re.match(r"\s+(Bước|Model) \[", x))  # tool tự điền
                      for b in buf.getvalue().strip().split("\n\n") if b.strip().startswith(can))
    here = os.path.dirname(os.path.abspath(__file__))
    if loai != "doichieu":  # agent doc-trang: hướng dẫn nằm sẵn trong system prompt (SKILL\_agents\doc-trang.md), ghi bằng luu
        try:
            with open(os.path.join(p.t, "README.md"), encoding="utf-8") as fh:
                rd = fh.read()
        except OSError:
            rd = ""
        bay = "\n".join(re.findall(r"(?ms)^## (?:Thứ tự ưu tiên|Bẫy đã biết)[^\n]*\n.*?(?=^## |\Z)", rd)).strip()
        print(f"""Bước 7d — {ten}, nhóm {n} (agent doc-trang: làm theo system prompt, KHÔNG đọc 7d-docsau.md / README).
GOC = {p.goc}
TOOL = {os.path.join(here, 'td.py')}
Ảnh tạm (PNG crop): {png}
Ghi: chỉ bằng `luu <GOC> {n} "<FID>|<trang>"` (file TRANG_QS/PAGES/ISSUES .docsau-{n}.json do tool ghi).

Cột được ghi:
{cot}
""" + (f"\nTừ README dự án (thứ tự ưu tiên, bẫy đã biết):\n{bay}\n" if bay else "") + """
Phần giao:
""" + "\n".join(khoi))
        return
    print(f"""Làm bước 7d (docsau) của skill 10-tender — {ten}, nhóm {n}. Chỉ làm phần giao dưới đây.
Đọc rồi làm đúng theo: {os.path.join(os.path.dirname(here), 'buoc', '7d-docsau.md')} (đã gồm cách ghi JSON: KHÔNG đọc JSON.md, không đọc mã tool) và {os.path.join(p.t, 'README.md')}.
GOC = {p.goc}   TOOL = {os.path.join(here, 'td.py')}
Ảnh tạm (PNG crop) để ở: {png}
Ghi: {os.path.join(p.t, '_data')}\\{ghi}
Không tự viết script / sed / grep / find / cat file _data: dùng xemtrang, xemcum, loc, chu, so, anh, gemini; ghi JSON bằng Write, sửa bằng Edit.
Xem cụm: `xemcum <GOC> <cụm 1> <cụm 2> … --gon` (nhiều cụm 1 lần; tool tự dừng trước khi output quá dài và in lệnh gọi tiếp). Mục `TRÙNG SỐ BOQ CỤM KHÁC`: xét từng cụm nêu — cùng 1 món (tính trùng KL / gán nhầm cụm) → ISSUES; chỉ trùng số ngẫu nhiên → bỏ qua. Kiểm phần của mình: `check <GOC> --den docsau --nhom {n}`.

Cột được ghi:
{cot}

Phần giao:
""" + "\n".join(khoi) + """

Trả về ≤ 15 dòng: file đã ghi, số trang/cụm đã làm, món còn KHÔNG THẤY ở trường quyết định giá, câu cần hỏi user.""")


# ---------------------------------------------------------------- neo
class Nguon:
    """Chữ trang từ cache lệnh trang (_data\\_text), không mở lại PDF."""

    def __init__(self, p, res):
        self.p, self.fidx, self.c = p, res["FILE_INDEX"], {}

    def _pages(self, fid):
        if fid not in self.c:
            self.c[fid] = (jload(self.p.f(os.path.join("_text", fid + ".json"))) or {}).get("pages")
        return self.c[fid]

    def text(self, fid, pg):
        t = self._pages(fid)
        return t[pg - 1].get("text") if t and 1 <= pg <= len(t) else None

    def so_trang(self, fid):
        return len(self._pages(fid) or [])

    def _ten(self, fid):
        r = self.fidx.get(fid, {})
        return (s(r.get("Trong zip")) or s(r.get("Đường dẫn"))).lower()

    def la_pdf(self, fid):
        return self._ten(fid).endswith(".pdf")

    def la_xlsx(self, fid):
        return self._ten(fid).endswith((".xlsx", ".xlsm"))

    def duong_dan(self, fid):
        return td.link_item(self.p, ["", fid], self.fidx)[1]


def _l(l):
    return (l[1], l[2] if len(l) > 2 else None, l[3] if len(l) > 3 else None)


NEO_TO = [("SCOPE", "Bản vẽ hiệu lực"), ("SCOPE", "Bản vẽ đối chứng"), ("BOQ_MAP", "Tờ đếm KL"), ("BOQ_MAP", "Tờ chi tiết"), ("DAY_TUONG", "Tờ mặt bằng")]
NEO_SPEC = {("SCOPE", "Spec · ML"): "Thông số riêng", ("SPEC", "File · Trang"): "Mục", ("MAT", "Nguồn spec"): "Mã VL"}
NEO_GT = [("SCOPE", "Kích thước BV", "Bản vẽ đối chứng", "kt"), ("BOQ_MAP", "Kích thước BV", "Tờ chi tiết", "kt")]
NEO_GT_QS = {"Mã hoàn thiện trên BV": "ma", "Kích thước BV": "kt"}


def _dong(res, sh):
    v = res[sh]
    return list(v.items()) if isinstance(v, dict) else [(f"{sh} dòng {i}", r) for i, r in enumerate(v, 2)]


def _chi_doc_sau(res, sh, k, r):
    """SCOPE: chỉ dòng đại diện của cụm, tránh kiểm lặp cùng link."""
    return sh != "SCOPE" or s(r.get("Cụm đọc")) in ("", k)


def neo_items(res, ng):
    out = []
    for sh, cot in NEO_TO:
        for i, (k, r) in enumerate(_dong(res, sh), 2):
            if not _chi_doc_sau(res, sh, k, r):
                continue
            for l in links_of(r.get(cot)):
                f, pg, o = _l(l)
                out.append(dict(sh=sh, dong=i, cot=cot, ten=f"{k} · {cot}", kieu="to", nhan=l[0], file=f, trang=pg, o=o, ma=[]))
    qs_spec = [(("QUAN_SAT", "Tờ · Chi tiết"), "Giá trị")]
    for (sh, cot), cv in list(NEO_SPEC.items()) + qs_spec:
        for i, (k, r) in enumerate(_dong(res, sh), 2):
            if not _chi_doc_sau(res, sh, k, r):
                continue
            ls = links_of(r.get(cot))
            if not ls:
                continue
            nguon = r.get("Nguồn")
            kieu = "spec" if sh != "QUAN_SAT" or nguon == "SPEC" else ("to" if nguon in ("HÌNH", "TEXT") else None)
            if not kieu:
                continue
            val = r.get(cv) or (k if sh == "MAT" else "")
            ten = f"{r.get('Cụm đọc') or k} · {r.get('Trường') or cot}"
            for l in ls:
                f, pg, o = _l(l)
                ma = neo_lib.neo_ma(l[0], val if len(ls) == 1 else "", file=ng._ten(f)) if kieu == "spec" else []
                out.append(dict(sh=sh, dong=i, cot=cot, ten=ten, kieu=kieu, nhan=l[0], file=f, trang=pg, o=o, ma=ma))
            if kieu == "spec" and len(ls) > 1 and neo_lib.neo_ma(val):  # mã trong giá trị phải có trên ít nhất 1 trang/dòng được link của ô
                f, pg, o = _l(ls[0])
                out.append(dict(sh=sh, dong=i, cot=cot, ten=ten, kieu="spec", nhan=f"giá trị ({len(ls)} link)", file=f, trang=pg, o=o, ma=neo_lib.neo_ma(val), nhom=[_l(l) for l in ls]))
    vals = [r.get("Giá trị") for r in res["QUAN_SAT"] if "mã" in s(r.get("Trường")).lower()]
    vals += [r.get("Mã hoàn thiện") for r in res["SCOPE"].values()] + [r.get("Mã liên quan") for r in res["SPEC"]] + list(res["MAT"])
    ho = neo_lib.neo_ho(vals)
    for sh, cv, cl, kieu in NEO_GT + [("QUAN_SAT", "Giá trị", "Tờ · Chi tiết", None)]:
        for i, (k, r) in enumerate(_dong(res, sh), 2):
            if not _chi_doc_sau(res, sh, k, r):
                continue
            kk = kieu or (NEO_GT_QS.get(r.get("Trường")) if r.get("Nguồn") in ("HÌNH", "TEXT") else None)
            val = r.get(cv)
            if not kk or not val or s(val).upper().startswith("KHÔNG") or r.get("Trạng thái") in ("KHÔNG THẤY", "KHÔNG CÓ"):
                continue
            ls = [_l(l) for l in links_of(r.get(cl)) if ng.la_pdf(l[1]) and len(l) > 2 and str(l[2]).isdigit()]
            if not ls:
                continue
            ma = list(neo_lib.neo_mavl(val, ho).values()) if kk == "ma" else neo_lib.neo_kt_gt(val)
            out.append(dict(sh=sh, dong=i, cot=r.get("Trường") or cv, ten=f"{r.get('Cụm đọc') or k} · {r.get('Trường') or cv}", kieu=kk, nhan=s(val)[:60],
                            file=ls[0][0], trang=ls[0][1], o=ls[0][2], ma=ma, nhom=ls, ho=sorted(ho)))
    return out


def trang_bv(p, res):
    """(fid, trang, số BV): khung tên đọc được ở cache chữ, tờ 1 trang lấy Số BV của FILE_INDEX, PAGES do LLM sửa thắng."""
    out = {}
    for fid, r in res["FILE_INDEX"].items():
        pages = (jload(p.f(os.path.join("_text", fid + ".json"))) or {}).get("pages")
        for i, x in enumerate(pages or [], 1):
            out[(fid, i)] = s(r.get("Số BV")) if s(r.get("Số BV")) and len(pages) == 1 else s(x.get("so"))
    for r in res["PAGES"].values():
        if s(r.get("Số BV")) and s(r.get("File ID")) and r.get("Trang"):
            out[(s(r["File ID"]), int(r["Trang"]))] = s(r["Số BV"])
    return [(f, t, v) for (f, t), v in out.items()]


def cmd_neo(goc):
    p = td.P(goc).need()
    if not os.path.isdir(p.f("_text")):
        sys.exit("Chưa có cache chữ (_data\\_text): chạy trang trước.")
    res, _ = merge(p)
    ng = Nguon(p, res)
    kq = neo_lib.run(neo_items(res, ng), ng, trang_bv(p, res))
    jdump(p.f("_neo.json"), kq)
    for l in kq["lech"]:
        print("  XEM", neo_lib.neo_dong(l))
    print(neo_lib.neo_tong(kq))
    print(f"-> {p.f('_neo.json')}  (LỆCH: mở lại đúng trang, sửa link/giá trị hoặc ghi lý do; không phải LỖI)")


def neo_xem(p, res):
    kq = jload(p.f("_neo.json"))
    if not kq:
        return []
    if kq.get("hash") != neo_lib.neo_hash(neo_items(res, Nguon(p, res))):
        return [f"NEO: link đã đổi sau lần chạy neo {kq.get('ngay', '')} — chạy lại lệnh neo để kiểm"]
    return [neo_lib.neo_dong(l) for l in kq["lech"]] + [neo_lib.neo_tong(kq)]


# ---------------------------------------------------------------- tincay
def kt_chuoi_sai(qs_rows):
    """Chuỗi 'a+b=c' trong Giá trị mà tổng sai -> [(cụm, trường, chuỗi, tổng thật)]."""
    bad = []
    for r in qs_rows:
        for m in CHUOI.finditer(s(r.get("Giá trị"))):
            try:
                parts = [float(x.replace(",", ".")) for x in re.split(r"\s*\+\s*", m.group(1))]
                tot = float(m.group(2).replace(",", "."))
            except ValueError:
                continue
            if abs(sum(parts) - tot) > 0.51:
                bad.append((s(r.get("Cụm đọc")), s(r.get("Trường")), m.group(0), sum(parts)))
    return bad


def gemini_lech(gem):
    """Dòng 'Kiểm chéo Gemini' -> đoạn sau 'lệch:' nếu khác 'không', else None."""
    lech = [x.strip() for r in gem for x in re.findall(r"lệch:\s*([^;|]*)", s(r.get("Giá trị")))]  # ghep nối nhiều trang bằng ' | '
    lech = [x for x in lech if x and not _fold(x).startswith("khong")]
    return "; ".join(lech) if lech else None


def cmd_tincay(goc):
    p = td.P(goc).need()
    cum = nap_cum(p)
    if not cum:
        sys.exit("LỖI: chưa có _docsau.json — chạy docsau trước.")
    res, _ = merge(p)
    scope, bmap = res["SCOPE"], res["BOQ_MAP"]
    qs = co_quan_sat(res)
    thieu_hd = thieu_hop_dong(res, cum)
    iss_of = {}
    for k, r in res["ISSUES"].items():
        if k.startswith("DT") or r.get("Trạng thái") in ("ĐÃ CHỐT", "ĐÓNG"):  # DT = nháp lần trước, file này ghi lại
            continue
        for c in td.split_list(r.get("Cụm đọc")):
            iss_of.setdefault(c, []).append(k)
    map_state = {}
    for k, r in bmap.items():
        bv = td.num(r.get("KL theo BV")) if r.get("KL theo BV") not in (None, "") else None
        kl = td.num(r.get("KL BOQ"))
        map_state[k] = "CHƯA ĐỌC" if bv is None else ("KHỚP" if kl is not None and abs(kl - bv) < 1e-6 else "LỆCH")
    muc, ly_do, moi, gs_rows = {}, {}, [], []
    new_id = lambda: f"DT{len(moi) + 1:02d}"
    t_rows, m_rows = [], []
    for c, info in cum.items():
        rows = [r for r in qs.get(c, []) if r.get("Trường") != GEM_SO_TRUONG]
        why, lv = [], "CAO"
        by = {}
        for r in rows:
            by.setdefault(r.get("Trường"), []).append(r)
        if not [r for r in rows if r.get("Nguồn") in ("HÌNH", "TEXT")]:
            lv = worse(lv, "THAP")
            why.append("chưa có quan sát từ hình/text")
        tr_hd = info.get("truong") or {}
        gia = [f for f, x in tr_hd.items() if x.get("gia")] or PRICE_FIELDS  # theo profile; _docsau.json cũ: danh sách cố định
        kt = [f for f in gia if f in by and not any(x.get("Trạng thái") in ("CÓ", "KHÔNG CÓ") for x in by[f]) and any(x.get("Trạng thái") == "KHÔNG THẤY" for x in by[f])]
        if kt:
            lv = worse(lv, "THAP")
            why.append("KHÔNG THẤY: " + ", ".join(kt))
        miss = thieu_hd.get(c, [])  # trường bắt buộc không có dòng nào = chưa đọc, không phải "tốt"
        miss_gia = [f for f in miss if tr_hd.get(f, {}).get("gia")]
        if miss_gia:
            lv = worse(lv, "THAP")
            why.append("chưa kết luận (bắt buộc, quyết định giá): " + ", ".join(miss_gia))
            kt = kt + [f for f in miss_gia if f not in kt]
        elif miss:
            lv = worse(lv, "TB")
            why.append("chưa kết luận (bắt buộc): " + ", ".join(miss))
        mt = sorted({s(r.get("Trường")) for r in rows if r.get("Trạng thái") == "MÂU THUẪN"})
        mtg = [f for f in mt if f in PRICE_FIELDS and any(x.get("Nguồn") == "GEMINI" and x.get("Trạng thái") == "MÂU THUẪN" for x in by[f])]
        if mtg:
            lv = worse(lv, "THAP")
            why.append("Claude↔Gemini lệch ở: " + ", ".join(mtg))
        if mt:
            lv = worse(lv, "TB")
            why.append("MÂU THUẪN: " + ", ".join(mt))
        rep = scope.get(c, {})
        dc = rep.get("Bản vẽ đối chứng")
        if not links_of(dc):
            lv = worse(lv, "THAP")
            why.append("thiếu Bản vẽ đối chứng")
        cd = [s(x.get("Giá trị")) for x in by.get("Tờ chưa đọc", [])]
        cd = [x for x in cd if x and not re.match(r"^(khong|het|none|-)", _fold(x).strip())]
        if cd:
            lv = worse(lv, "THAP")
            why.append("còn tờ chưa đọc")
        gem = by.get("Kiểm chéo Gemini", [])
        gl = gemini_lech(gem)
        if not gem or all(x.get("Trạng thái") == "KHÔNG THẤY" for x in gem):
            lv = worse(lv, "TB")
            why.append("Gemini chưa chạy")
        elif gl:
            lv = worse(lv, "THAP" if GIA_RX.search(_fold(gl)) else "TB")
            why.append("Gemini lệch: " + gl[:80])
        gs = gemini_so(p, c, rows)  # phần số do tool so, không dựa vào câu LLM viết
        if gs["tt"] != "KHÔNG CÓ SỐ":
            gs_rows.append(gemini_so_dong(c, gs))
        if gs["tt"] == "LỆCH":
            lv = worse(lv, "THAP")
            why.append("Gemini không đọc ra số: " + ", ".join(gs["thieu"]))
        elif gs["tt"] in ("CHƯA CHẠY", "LỖI"):
            lv = worse(lv, "TB")
            why.append(f"Gemini chưa so số kích thước ({gs['tt']})")
        ds = set(info["khoa"])
        st = {k: v for k, v in map_state.items() if s(bmap[k].get("Khóa")) in ds}
        lech = [k for k, v in st.items() if v == "LỆCH"]
        chua = [k for k, v in st.items() if v == "CHƯA ĐỌC"]
        if lech:
            lv = worse(lv, "TB")
            why.append("KL BOQ ≠ KL BV ở " + ", ".join(lech[:6]))
        if chua:
            lv = worse(lv, "TB")
            why.append(f"chưa có KL theo BV ở {len(chua)} khu")
        muc[c], ly_do[c] = lv, why
        if c not in iss_of:  # issue nháp chỉ cho cụm chưa có issue nào
            lk = links_of(rep.get("Bản vẽ đối chứng")) or links_of(rep.get("Bản vẽ hiệu lực")) or ""
            def them(loai, muc_, nd):
                moi.append({"ID": new_id(), "Loại": loai, "Cụm đọc": c, "Mức": muc_, "Nội dung": nd, **({"Nguồn": lk} if lk else {}), "Trạng thái": "MỞ"})
            if lech:
                them("BOQ↔BV", "TB", "KL lệch — " + "; ".join(f"{bmap[k].get('Khu')} ({bmap[k].get('Khóa')}): BOQ {bmap[k].get('KL BOQ')} ≠ BV {bmap[k].get('KL theo BV')}" for k in lech[:10]))
            if kt:
                them("Thiếu tờ / KHÔNG THẤY", "CAO" if "Kích thước BV" in kt else "TB", "Trường quyết định giá chưa đọc được: " + ", ".join(kt))
            if gl:
                them("Claude↔Gemini", "TB", "Gemini lệch: " + gl[:200])
            if gs["tt"] == "LỆCH":
                them("Claude↔Gemini", "TB", "Gemini không đọc ra số Claude ghi ở Kích thước BV: " + ", ".join(gs["thieu"]))
            iss_of[c] = [x["ID"] for x in moi if x["Cụm đọc"] == c]
        for k in info["khoa"]:
            t_rows.append({"Khóa": k, "Độ tin cậy": lv, "Issue": ", ".join(iss_of.get(c, []))})
        for k in st:
            v = st[k]
            m_rows.append({"Khóa KL": k, "Trạng thái": v, "Độ tin cậy": worse(lv, {"KHỚP": "CAO", "LỆCH": "TB", "CHƯA ĐỌC": "THAP"}[v]),
                           **({"Issue": ", ".join(iss_of[c])} if v == "LỆCH" and iss_of.get(c) else {})})
    write_tool(p, "SCOPE", "tincay", t_rows)
    write_tool(p, "BOQ_MAP", "tincay", m_rows)
    write_tool(p, "ISSUES", "tincay", moi)
    write_tool(p, "QUAN_SAT", "tincay", gs_rows)
    jdump(p.f("_tincay.json"), {"muc": muc, "ly_do": ly_do})
    print(f"Độ tin cậy nháp: {dict(Counter(muc.values()))} cụm; KL từng khu: {dict(Counter(map_state.values()))}; issue nháp {len(moi)}; Gemini so số: {dict(Counter(r['Trạng thái'] for r in gs_rows))}")
    for b in kt_chuoi_sai(res["QUAN_SAT"]):
        print("  CHUỖI SAI", b)
    print("LLM còn phải làm: Mã hoàn thiện (BV↔MAT), Giả định tính giá, Hành động/RFI cho issue, RFI_LOG, ASK_USER; lý do từng cụm ở _tincay.json")


# ---------------------------------------------------------------- check
def kiem(p, res, buoc_da):
    """-> (LỖI, XEM). buoc_da = các bước đã tới theo --den. Chưa chạy docsau thì chỉ nhắc, dự án cũ không bị báo lỗi."""
    errs, xem = [], []
    cum = nap_cum(p)
    if not cum:
        if {"docsau", "tincay"} & set(buoc_da):
            xem.append("Chưa đọc sâu: chưa chạy docsau, không có kích thước, cấu tạo, KL theo BV, Độ tin cậy")
        return errs, xem
    qs = co_quan_sat(res)
    for c in qs:
        if c not in cum:
            xem.append(f"QUAN_SAT: Cụm đọc '{c}' không còn trong _docsau.json (cụm đã đổi do Bản vẽ hiệu lực đổi?) → đọc lại")
    for r in res["QUAN_SAT"]:
        if r.get("Nguồn") in ("HÌNH", "TEXT") and not r.get("Tờ · Chi tiết"):
            errs.append(f"QUAN_SAT {r.get('Cụm đọc')} · {r.get('Trường')}: nguồn {r.get('Nguồn')} mà không ghi tờ/chi tiết")
    for a, b, ch, tong in kt_chuoi_sai(res["QUAN_SAT"]):
        errs.append(f"QUAN_SAT {a} · {b}: chuỗi cộng sai '{ch}' (tổng thật {tong:g})")
    # lượt 0: quan sát theo trang dùng chung
    tq = res.get("TRANG_QS", [])
    for r in tq:
        if r.get("Nguồn") in ("HÌNH", "TEXT") and not r.get("Tờ · Chi tiết"):
            errs.append(f"TRANG_QS {r.get('Khóa trang')} · {r.get('Món')} · {r.get('Trường')}: nguồn {r.get('Nguồn')} mà không ghi tờ/chi tiết")
    for a, b, ch, tong in kt_chuoi_sai([dict(r, **{"Cụm đọc": f"{s(r.get('Khóa trang'))} · {s(r.get('Món'))}"}) for r in tq]):
        errs.append(f"TRANG_QS {a} · {b}: chuỗi cộng sai '{ch}' (tổng thật {tong:g})")
    ds = jload(p.f("_docsau.json")) or {}
    theo_trang = ds.get("phien_ban") == PHIEN_BAN
    if theo_trang:
        for r in tq:
            if s(r.get("Trường")) == "KL theo BV" and s(r.get("Trạng thái")) == "CÓ" and _so_dau(r.get("Giá trị")) is None:
                errs.append(f"TRANG_QS {r.get('Khóa trang')} · {r.get('Món')}: 'KL theo BV' trạng thái CÓ mà Giá trị không có số")
    if theo_trang and "docsau" in buoc_da:
        da = {s(r.get("Khóa trang")) for r in tq} | set(res["PAGES"])
        for n, v in sorted((ds.get("trang_nhom") or {}).items(), key=lambda x: td.so_tu_nhien(x[0])):  # báo theo nhóm để check --nhom lọc được
            chua = sorted(set(v) - da, key=td.so_tu_nhien)
            if chua:
                errs.append(f"Đọc trang · nhóm {n}: {len(chua)}/{len(v)} trang được giao chưa có TRANG_QS/PAGES (nhóm chưa làm xong) vd: {', '.join(chua[:8])}")
        moi = max((os.path.getmtime(f) for f in glob.glob(p.f("TRANG_QS.docsau-*.json")) + glob.glob(p.f("PAGES.docsau-*.json"))), default=0)
        g, fg = jload(p.f("_ghep.json")), p.f("_ghep.json")
        if tq and (not g or os.path.getmtime(fg) < moi):
            errs.append("ghep: chưa chạy hoặc TRANG_QS/PAGES mới hơn lần ghép cuối — chạy `ghep <GOC>` rồi mới giao nhóm đối chiếu")
        elif g:
            for kt, cs in g.get("thieu_trang", {}).items():
                errs.append(f"Đọc trang {kt}: món được giao không có dòng nào trên trang (ghi giá trị hoặc KHÔNG THẤY + lý do): {', '.join(cs[:8])}")
            if g.get("chua_ghep"):
                xem.append(f"TRANG_QS: {len(g['chua_ghep'])} quan sát chỉ có Cụm ứng viên (chưa gán cụm) — nhóm đối chiếu quyết, ghi QUAN_SAT cho đúng cụm"
                           f" vd: {'; '.join(x['trang'] + ' ' + x['mon'] for x in g['chua_ghep'][:4])}")
            if g.get("kl_mau_thuan"):
                xem.append(f"KL theo BV mâu thuẫn giữa các trang (không cộng): {len(g['kl_mau_thuan'])} khu, vd: {g['kl_mau_thuan'][0]['cum']} {g['kl_mau_thuan'][0]['khu']}")
            if g.get("kl_khong_khop"):
                xem.append(f"KL theo BV không gán được vào BOQ_MAP: {len(g['kl_khong_khop'])} khu (tên khu khác BOQ hoặc khu trùng nhiều dòng)")
            if g.get("dan_chieu_ngoai"):
                xem.append(f"Dẫn chiếu tới số BV ngoài bộ tờ: {sum(map(len, g['dan_chieu_ngoai'].values()))} ở {len(g['dan_chieu_ngoai'])} cụm (có ở Tờ chưa đọc) — giao tìm tờ/đọc bổ sung")
            if g.get("dan_chieu_khong_cum_chua_doc"):
                xem.append(f"Dẫn chiếu chưa gán cụm tới tờ chưa trang nào đọc: {', '.join(g['dan_chieu_khong_cum_chua_doc'][:8])}"
                           f" — nhóm đối chiếu gán cụm (Cụm ứng viên) hoặc giao tìm tờ/đọc bổ sung (_ghep.json: dan_chieu_khong_cum)")
        pl = [c for c, v in cum.items() if v.get("profile") == CHUA_PL]
        if pl:
            xem.append(f"{len(pl)} cụm CHƯA PHÂN LOẠI (thiếu Mã cấp 3): hợp đồng kiểm dùng trường chung, vd: {', '.join(pl[:6])}")
    elif tq and not theo_trang and "docsau" in buoc_da:
        xem.append("TRANG_QS có dữ liệu nhưng _docsau.json là bản cũ (2 lượt) — chạy lại docsau để chia theo trang")
    for r in res["ISSUES"].values():
        for x in re.split(r"[,;\s]+", s(r.get("RFI"))):
            if x and x not in res["RFI_LOG"]:
                errs.append(f"ISSUES {r.get('ID')}: RFI '{x}' không có trong RFI_LOG")
    if "docsau" in buoc_da:
        gem_chua = []
        for c in cum:
            rows = [r for r in qs.get(c, []) if r.get("Trường") != GEM_SO_TRUONG]
            if not rows:
                errs.append(f"{c}: chưa có dòng QUAN_SAT nào (bước docsau chưa làm cụm này)")
                continue
            if not any(x.get("Trường") == "Kiểm chéo Gemini" for x in rows):
                errs.append(f"{c}: thiếu dòng QUAN_SAT 'Kiểm chéo Gemini' (phải cho Gemini đọc lại các crop)")
            gs = gemini_so(p, c, rows)
            if gs["tt"] in ("CHƯA CHẠY", "LỖI"):
                gem_chua.append(c)
            elif gs["tt"] == "LỆCH":
                xem.append(f"{c}: Gemini không đọc ra số Claude ghi: {', '.join(gs['thieu'])}")
            kt = [s(x.get("Trường")) for x in rows if x.get("Trạng thái") == "KHÔNG THẤY"]
            mt = [s(x.get("Trường")) for x in rows if x.get("Trạng thái") == "MÂU THUẪN"]
            if kt:
                xem.append(f"{c}: KHÔNG THẤY {len(kt)} trường ({', '.join(kt[:5])})")
            if mt:
                xem.append(f"{c}: MÂU THUẪN {len(mt)} trường ({', '.join(mt[:5])})")
        if gem_chua:
            xem.append(f"{len(gem_chua)} cụm có kích thước nhưng chưa so số với Gemini (gemini --trang / --cum), không được CAO: {', '.join(gem_chua[:15])}")
        if theo_trang:  # hợp đồng kiểm: trường bắt buộc của profile phải có kết luận (kể cả KHÔNG THẤY + lý do)
            for c, miss in thieu_hop_dong(res, {c: v for c, v in cum.items() if qs.get(c)}).items():
                errs.append(f"{c} [{cum[c].get('profile')}]: trường bắt buộc chưa có kết luận: {', '.join(miss)}")
        xem.extend(neo_xem(p, res))
    if "tincay" in buoc_da:
        thieu = [k for c in cum.values() for k in c["khoa"] if not s(res["SCOPE"].get(k, {}).get("Độ tin cậy"))]
        if thieu:
            errs.append(f"SCOPE: {len(thieu)} dòng đọc sâu thiếu 'Độ tin cậy' (bước tincay) vd: {', '.join(thieu[:6])}")
        for k, r in res["ISSUES"].items():
            if r.get("Trạng thái") in (None, "", "MỞ") and r.get("Mức") == "CAO":
                xem.append(f"Issue {k} ({r.get('Cụm đọc')}) mức CAO đang mở: {s(r.get('Nội dung'))[:70]}")
        for k, r in res["ASK_USER"].items():
            if not r.get("Trả lời"):
                xem.append(f"Hỏi user {k}: {s(r.get('Câu hỏi'))[:90]}")
        nt = [k for k, r in res["PAGES"].items() if r.get("Lớp text") == "KHÔNG" and not r.get("Đã đọc bằng")]
        if nt:
            xem.append(f"{len(nt)} trang PDF không có lớp text chưa được xem bằng hình (PAGES)")
    loai = Counter(s(r.get("Vai trò")) for r in res["FILE_INDEX"].values() if s(r.get("Vai trò")) in ("Ngoài scope", "Bản cũ/Trùng"))
    if loai:
        xem.append(f"ĐÃ LOẠI khỏi đọc: {dict(loai)} file (FILE_INDEX, cột Vai trò); thư mục bỏ: {p.cfg.get('bo_thu_muc') or 'không'}")
    return errs, xem


def loc_nhom(p, n, errs, xem):
    """check --nhom N: chỉ giữ LỖI/XEM nhắc tới trang / cụm / file của nhóm N (trang_nhom, tim_to, so_nhom trong _docsau.json).
    -> (errs, xem, số dòng ẩn). Dòng chung cả dự án (ghep chưa chạy, RFI…) bị ẩn: phiên điều phối chạy check không --nhom."""
    ds = jload(p.f("_docsau.json")) or {}
    tok = set()
    for k in ("trang_nhom", "tim_to", "so_nhom"):
        tok |= set((ds.get(k) or {}).get(str(n), []))
    if not tok:
        sys.exit(f"LỖI: không có nhóm {n} trong _docsau.json")
    if (ds.get("tim_to") or {}).get(str(n)) or (ds.get("so_nhom") or {}).get(str(n)):  # nhóm theo cụm: thêm trang của các cụm
        scope = merge(p)[0]["SCOPE"]
        tok |= {t for c in list(tok) for t in to_cua(scope, c)}
    rx = re.compile(r"(?<![\w|!])(?:" + "|".join(re.escape(t) for t in sorted(tok, key=len, reverse=True)) + r")(?![\w])"
                    + rf"|\bnhóm {re.escape(str(n))}\b|\.docsau-{re.escape(str(n))}\.json")
    e2, x2 = [x for x in errs if rx.search(x)], [x for x in xem if rx.search(x)]
    return e2, x2, len(errs) + len(xem) - len(e2) - len(x2)


def do(goc):
    """Tỉ lệ ô do tool ghi trên tổng ô, theo sheet KHÓA."""
    p = td.P(goc).need()
    res, src = merge(p)
    for sh, chu in src.items():
        ca = [b for o in chu.values() for b in o.values()]
        if ca:
            n = sum(1 for b in ca if b.endswith("-tool"))
            print(f"{sh}: tool {n}/{len(ca)} ô ({n * 100 // len(ca)}%)")


def so(pdf, trang, clip=None):
    print(doc_file.so(pdf, trang, clip))
