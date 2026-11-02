# -*- coding: utf-8 -*-
r"""docsau.py — phần ĐỌC SÂU của skill 10-tender (chuyển từ 01-scan), td.py gọi vào.
Áp cho toàn gói: mọi dòng BOQ còn làm (có đơn vị, không thuộc sheet Prelims/Provisional/Summary, Trạng thái không phải Omit/Loại trừ/Không áp dụng).
Khác 01-scan: đơn vị đọc là CỤM (các dòng cùng mô tả + mã VL + mã FF&E + tờ bản vẽ hiệu lực), khóa cụm = `Khóa` của dòng đầu tiên, cột `Cụm đọc`.

  cmd_docsau   gom cụm, chia nhóm, nháp Phạm vi nhà máy, BOQ_MAP
  cmd_gemini   Gemini đọc độc lập, --cum lưu số theo cụm
  cmd_neo      kiểm giá trị LLM ghi có trên trang được link không (chỉ XEM)
  cmd_tincay   Độ tin cậy, trạng thái KL, issue nháp
  kiem         phần kiểm thêm cho td.py check
"""
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


# ---------------------------------------------------------------- docsau
def cmd_docsau(goc, args):
    p = td.P(goc).need()
    n_nhom, loc_phong, loc_sheet = int(opt(args, "--nhom", "8")), opt(args, "--phong"), opt(args, "--sheet")
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
    # nhóm: cùng phòng đầu tiên (hoặc cùng sheet) đi với nhau, tối đa n_nhom cụm
    theo = OrderedDict()
    for c in chon:
        r = scope[c]
        ph = tach_phong(r.get("Mã phòng"))
        theo.setdefault(ph[0][0] if ph else s(r.get("Sheet")), []).append(c)
    nhom, cur = [], []
    for ds in theo.values():
        for c in ds:
            cur.append(c)
            if len(cur) >= n_nhom:
                nhom.append(cur)
                cur = []
    if cur:
        nhom.append(cur)
    jdump(p.f("_docsau.json"), {"cum": {c: {"khoa": ds} for c, ds in cum.items()}, "nhom": nhom})
    dong = []
    for i, ds in enumerate(nhom, 1):
        dong.append(f"== Nhóm {i} ({len(ds)} cụm)")
        for c in ds:
            r = scope[c]
            lk = "; ".join(l[0] for l in links_of(r.get("Bản vẽ hiệu lực"))[:6]) or s(_txt(r.get("Bản vẽ hiệu lực"))) or "(chưa có tờ)"
            dong.append(f"{c} [{len(cum[c])} dòng] {s(r.get('ĐV'))} | {s(r.get('Mô tả BOQ'))[:140]} | VL: {s(r.get('Mã VL'))} {s(r.get('Mã spec/FF&E'))} | Tờ: {lk}")
    with open(p.f("_docsau.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(dong) + "\n")
    print(f"Đọc sâu: {sum(len(v) for v in cum.values())} dòng → {len(cum)} cụm; giao lần này {len(chon)} cụm → {len(nhom)} nhóm (≤ {n_nhom} cụm/nhóm) -> {p.f('_docsau.txt')}")
    print(f"Không đọc sâu: {dict(bo) or 'không'}. Nhãn Phạm vi nhà máy: {dict(nhan)}")
    if chua_tt:
        print(f"XEM: {len(chua_tt)} dòng chưa có Trạng thái, vẫn đọc sâu; nên chạy xong bước 7 trước để khỏi đọc dòng Omit/Loại trừ.")
    khong_to = [c for c in cum if not links_of(scope[c].get("Bản vẽ hiệu lực"))]
    if khong_to:
        print(f"XEM: {len(khong_to)} cụm chưa có link Bản vẽ hiệu lực (bước 6), subagent phải tự tìm tờ hoặc ghi KHÔNG THẤY.")


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
    p = qs_cum = None
    if cum:
        if hoi or vat or not goc:
            sys.exit("--cum chỉ dùng ở chế độ số và cần --goc <GOC>")
        p = td.P(goc).need()
        qs_cum = [r for r in co_quan_sat(merge(p)[0]).get(cum, []) if r.get("Trường") != GEM_SO_TRUONG]
        if not claude_kt(qs_cum):
            sys.exit(f"{cum}: QUAN_SAT chưa có 'Kích thước BV' dạng R×S×C (nguồn HÌNH/TEXT). Ghi số của Claude vào QUAN_SAT.docsau-<nhóm>.json trước rồi mới chạy Gemini.")
    kq = gemini_doc.doc_anh(pngs, hoi, vat, par)
    tat_ca = []
    for png, txt in kq:
        print(f"== {os.path.basename(png)}\n{txt}")
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
    elif mine and not hoi and not vat:
        a, g = Counter(x.strip() for x in mine.split(",") if x.strip()), Counter(x for x in tat_ca if x != "?")
        print(f"-- SO SÁNH: khớp {sum((a & g).values())} · chỉ Claude {sorted((a - g).elements())} · chỉ Gemini {sorted((g - a).elements())}")


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
    m = re.search(r"lệch:\s*([^;]*)", s(gem[0].get("Giá trị"))) if gem else None
    return m.group(1).strip() if m and not _fold(m.group(1)).strip().startswith("khong") else None


def cmd_tincay(goc):
    p = td.P(goc).need()
    cum = nap_cum(p)
    if not cum:
        sys.exit("LỖI: chưa có _docsau.json — chạy docsau trước.")
    res, _ = merge(p)
    scope, bmap = res["SCOPE"], res["BOQ_MAP"]
    qs = co_quan_sat(res)
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
        kt = [f for f in PRICE_FIELDS if f in by and not any(x.get("Trạng thái") in ("CÓ", "KHÔNG CÓ") for x in by[f]) and any(x.get("Trạng thái") == "KHÔNG THẤY" for x in by[f])]
        if kt:
            lv = worse(lv, "THAP")
            why.append("KHÔNG THẤY: " + ", ".join(kt))
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
            xem.append(f"{len(gem_chua)} cụm có kích thước nhưng chưa so số với Gemini (gemini --cum), không được CAO: {', '.join(gem_chua[:15])}")
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
