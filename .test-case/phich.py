"""PHÍCH — dự án COSTING EXCEL, phạm vi: skill 10-tender (SKILL/10-tender/tool/td.py).

Chỉ phủ các hàm LUẬT THUẦN của td.py (không đọc/ghi file, chạy phần trăm giây):
  vai_tro_doan · loai_trang · tham_chieu · date_in · dot_of · ma_vl · chuan_ben
Không đụng 218 chỗ ghi của dự án (chọn phương án A: không test lệnh ghi trên dữ liệu thật).
Lớp LLM (docsau, gemini, bước 7c–7e): chưa khai, chưa có case.

Oracle viết lại bằng thao tác chuỗi thủ công (không dùng regex của td, không gọi td).
"""
import os
import sys
import datetime

DU_AN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOL = os.path.join(DU_AN, "SKILL", "10-tender", "tool")
sys.path.insert(0, TOOL)

import td  # noqa: E402  (nạp form + bảng mã ở chế độ đọc; không ghi)

tu_choi_them = []


# ─── Chạy hàm thật ─────────────────────────────────────────────────────────────

def _chuoi(x):
    if x is None:
        return "NONE"
    if isinstance(x, (list, tuple, set)):
        x = list(x)
        return "|".join(str(i) for i in x) if x else "(rong)"
    if isinstance(x, set):
        return "|".join(sorted(x))
    return str(x)


def chay_vai_tro(vao):
    return _chuoi(td.vai_tro_doan(vao["rel"], vao.get("inner"), vao["ext"]))


def chay_loai_trang(vao):
    return _chuoi(td.loai_trang(vao["t"]))


def chay_tham_chieu(vao):
    return _chuoi(td.tham_chieu(vao["text"]))


def chay_date_in(vao):
    return _chuoi(td.date_in(vao["parts"]))


def chay_dot_of(vao):
    return _chuoi(td.dot_of(vao["rel"]))


def chay_ma_vl(vao):
    return _chuoi(td.ma_vl(vao["text"], set(vao.get("rc", [])), set(vao.get("tien_to", []))))


def chay_ma_bang(vao):
    return _chuoi(td.ma_vl_theo_bang(vao["text"], set(vao["biet"])))


def chay_tap_ma(vao):
    import dmvt_xlsx
    return _chuoi(sorted(dmvt_xlsx.tap_ma(vao["o"])))


def chay_doc_ncc(vao):
    import spec_boc
    d = spec_boc.doc_ncc(vao["dong"])
    if not d:
        return "{}"
    return "%s|%s|%s" % (d["manufacturer"], d["web"], " / ".join(x for x in d["contact"].split(chr(10)) if x))


def chay_chuan_ben(vao):
    return _chuoi(td.chuan_ben(vao["raw"], vao.get("gop", {})))


# ─── ORACLE: viết độc lập, thao tác chuỗi thủ công ─────────────────────────────

def _gach_xuoi(s):
    """Đổi dấu gạch ngược thành / bằng vòng ký tự (kiem_oracle coi str.replace là hàm của hệ thống đang test)."""
    return "".join("/" if ord(ch) == 92 else ch for ch in s)


def _tu(c):
    return c.isalnum() or c == "_"


def _tim_tu(chuoi, tu):
    """tu đứng riêng (hai bên không phải ký tự chữ/số/_)."""
    i = chuoi.find(tu)
    while i != -1:
        truoc = i == 0 or not _tu(chuoi[i - 1])
        sau = i + len(tu) >= len(chuoi) or not _tu(chuoi[i + len(tu)])
        if truoc and sau:
            return True
        i = chuoi.find(tu, i + 1)
    return False


def _tim_bat_dau_tu(chuoi, tu):
    """tu đứng sau ranh giới (chỉ kiểm bên trái)."""
    i = chuoi.find(tu)
    while i != -1:
        if i == 0 or not _tu(chuoi[i - 1]):
            return True
        i = chuoi.find(tu, i + 1)
    return False


def oracle_vai_tro(vao):
    rel, inner, ext = vao["rel"], vao.get("inner"), vao["ext"]
    t = (rel + "/" + (inner or "")).lower()
    ten = _gach_xuoi(inner or rel).split("/")[-1].lower()  # basename kiểu Windows/POSIX gộp
    if ext == ".msg":
        return "Mail"
    scope_ext = (".xlsx", ".xlsm", ".xls", ".pdf", ".doc", ".docx")
    if ext in scope_ext and _pvcv(ten):
        return "Scope"
    if ("tender quer" in ten or _tim_bat_dau_tu(ten, "tcq")) and ext in (".xlsx", ".pdf"):
        return "Làm rõ"
    if "asbuilt" in t or "as-built" in t:
        return "Tham khảo"
    if _ngoai_scope(t):
        return "Ngoài scope"
    for k in ("project schedule", "rules and regulations", "logistic", "annex"):
        if k in t:
            return "Tham khảo"
    return "NONE"


def _pvcv(ten):
    i = ten.find("pvcv")
    while i != -1:
        a = i == 0 or not ten[i - 1].isascii() or not ten[i - 1].isalnum()
        b = i + 4 >= len(ten) or not ten[i + 4].isascii() or not ten[i + 4].isalnum()
        if a and b:
            return True
        i = ten.find("pvcv", i + 1)
    return False


def _ngoai_scope(t):
    for ten in ("mep.zip", "lighting.zip"):
        i = t.find(ten)
        while i != -1:
            if i == 0 or t[i - 1] == "/":
                return True
            i = t.find(ten, i + 1)
    for th in ("elek", "elv", "fire", "hvac", "plum"):
        if "/" + th + "/" in t:
            return True
    return False


def oracle_loai_trang(vao):
    t = vao["t"]
    if "DRAWING REFERENCE:" in t:
        return "DR"
    dau = t.lstrip()
    for k, v in (("SPECIFICATION", "SPEC"), ("MATERIAL REFERENCE", "MR"),
                 ("ITEM SCHEDULE", "IS"), ("Image (s) provided", "ANH")):
        if dau[:len(k)] == k:
            return v
    return "NONE"


def oracle_date_in(vao):
    def hop_le(y, m, d):
        try:
            return datetime.date(int(y), int(m), int(d)).isoformat()
        except ValueError:
            return None

    for part in reversed(vao["parts"]):
        run, runs = "", []
        for ch in part + " ":
            if ch.isdigit():
                run += ch
            else:
                if run:
                    runs.append(run)
                run = ""
        for r in runs:
            if len(r) == 8 and r[:2] == "20":
                v = hop_le(r[:4], r[4:6], r[6:8])
                if v:
                    return v
        if len(part) >= 7 and part[:6].isdigit() and part[6] in "_ -":
            v = hop_le("20" + part[:2], part[2:4], part[4:6])
            if v:
                return v
    return "NONE"


def oracle_dot_of(vao):
    rel = _gach_xuoi(vao["rel"])
    if "/" not in rel:
        return "(gốc)|None|0"
    first = rel.split("/")[0]
    i = 0
    while i < len(first) and first[i].isdigit():
        i += 1
    if i > 0 and first[i:i + 1] == "." and len(first) >= i + 7 and first[i + 1:i + 7].isdigit():
        yy, mm, dd = first[i + 1:i + 3], first[i + 3:i + 5], first[i + 5:i + 7]
        try:
            ngay = datetime.date(2000 + int(yy), int(mm), int(dd)).isoformat()
        except ValueError:
            ngay = None
        return "%s|%s|%d" % (first, ngay, int(first[:i]))
    return "%s|None|999" % first


_TU_PHONG = ("rooms", "room", "units", "unit", "suites", "suite", "room type", "căn", "phòng", "loại phòng")


def _gan_phong(cua_so):
    w = cua_so.lower()
    for k in _TU_PHONG:
        i = w.find(k)
        while i != -1:
            truoc = i == 0 or not _tu(w[i - 1])
            sau = i + len(k) >= len(w) or not _tu(w[i + len(k)])
            if truoc and sau:
                return True
            i = w.find(k, i + 1)
    return False


def oracle_ma_vl(vao):
    text, rc, tien_to = vao["text"], set(vao.get("rc", [])), set(vao.get("tien_to", []))
    n = len(text)
    ra = set()
    for i in range(n):
        if i > 0 and (_tu(text[i - 1]) or text[i - 1] == "-"):
            continue
        j = i
        while j < n and "A" <= text[j] <= "Z":
            j += 1
        if not (2 <= j - i <= 3) or j >= n or text[j] != "-":
            continue
        k = j + 1
        if k + 2 > n or not (text[k:k + 2].isascii() and text[k:k + 2].isdigit()):
            continue
        k += 2
        if k < n and "A" <= text[k] <= "Z":
            k += 1
        if k < n and (_tu(text[k]) or text[k] == "-"):
            continue
        ma = text[i:k]
        if ma in rc and (ma.split("-")[0] not in tien_to or _gan_phong(text[max(0, i - 25):k + 25])):
            continue
        ra.add(ma)
    return "|".join(sorted(ra)) if ra else "(rong)"


def _la_ma_bang(tok):
    """1-4 chữ cái, '-' tùy chọn, 1-4 chữ số, 1 chữ cái tùy chọn (đếm tay, không regex)."""
    i, n = 0, len(tok)
    while i < n and tok[i].isascii() and tok[i].isalpha():
        i += 1
    if not 1 <= i <= 4:
        return False
    if i < n and tok[i] == "-":
        i += 1
    j = i
    while j < n and tok[j].isascii() and tok[j].isdigit():
        j += 1
    if not 1 <= j - i <= 4:
        return False
    if j < n and tok[j].isascii() and tok[j].isalpha():
        j += 1
    return j == n


def _la_ffe(m):
    """51-500 / 51-500B hoặc AB-1234(B): thuộc cột Mã spec/FF&E, không vào Mã VL."""
    a = m.split("-")
    if len(a) != 2:
        return False
    x, y = a
    if len(x) == 2 and x.isdigit() and len(y) >= 3 and y[:3].isdigit() and (len(y) == 3 or (len(y) == 4 and y[3].isalpha())):
        return True
    if len(x) == 2 and x.isalpha() and len(y) >= 4 and y[:4].isdigit():
        return len(y) == 4 or (len(y) == 5 and y[4].isalpha())
    return False


def oracle_ma_bang(vao):
    text, biet = vao["text"], set(vao["biet"])
    ra, run = [], ""
    for ch in text + " ":
        if _tu(ch) or ch == "-":
            run += ch
            continue
        if run and _la_ma_bang(run) and run.upper() in biet and not _la_ffe(run.upper()) and run.upper() not in ra:
            ra.append(run.upper())
        run = ""
    return "|".join(ra) if ra else "(rong)"


def oracle_chuan_ben(vao):
    raw = "" if vao["raw"] is None else str(vao["raw"]).strip()
    gop = vao.get("gop", {})
    if raw in gop:
        return gop[raw]
    ra = []
    for x in raw.split(";"):
        x = x.strip()
        if x:
            y = gop.get(x, x)
            if y not in ra:
                ra.append(y)
    return "; ".join(ra)


# ─── Nhánh: suy từ KẾT QUẢ CỦA HÀM THẬT (không từ oracle), mã ổn định ──────────────
# Case của điểm khác (thiếu khóa) trả "KHAC" để không nổ khi đo độ phủ chéo.

def _can(vao, *khoa):
    return all(k in vao for k in khoa)


def nhanh_vai_tro(vao):
    if not _can(vao, "rel", "ext"):
        return "KHAC"
    return "VT." + chay_vai_tro(vao).replace(" ", "_")


def nhanh_loai_trang(vao):
    return "KHAC" if not _can(vao, "t") else "LT." + chay_loai_trang(vao)


def nhanh_date_in(vao):
    if not _can(vao, "parts"):
        return "KHAC"
    return "DI.NONE" if chay_date_in(vao) == "NONE" else "DI.OK"


def nhanh_dot_of(vao):
    if not _can(vao, "rel") or "ext" in vao:
        return "KHAC"
    r = chay_dot_of(vao)
    return "DOT.GOC" if r.startswith("(gốc)") else ("DOT.KHONG_MA" if r.endswith("|999") else "DOT.MA")


def nhanh_ma_vl(vao):
    if not _can(vao, "text") or "rel" in vao:
        return "KHAC"
    return "VL.RONG" if chay_ma_vl(vao) == "(rong)" else "VL.CO"


def nhanh_tham_chieu(vao):
    if not _can(vao, "text") or "rc" in vao:
        return "KHAC"
    return "TC.RONG" if chay_tham_chieu(vao) == "(rong)" else "TC.CO"


def nhanh_chuan_ben(vao):
    if not _can(vao, "raw"):
        return "KHAC"
    return "BEN.RONG" if chay_chuan_ben(vao) == "" else "BEN.CO"


def nhanh_ma_bang(vao):
    if not _can(vao, "text", "biet"):
        return "KHAC"
    return "MB.RONG" if chay_ma_bang(vao) == "(rong)" else "MB.CO"


def nhanh_tap_ma(vao):
    if not _can(vao, "o"):
        return "KHAC"
    return "TM.RONG" if chay_tap_ma(vao) == "(rong)" else "TM.CO"


def nhanh_doc_ncc(vao):
    if not _can(vao, "dong"):
        return "KHAC"
    return "NCC.RONG" if chay_doc_ncc(vao) == "{}" else "NCC.CO"


def _duong(ma, chay, nhanh):
    def f(vao, so_ghi):
        kq = chay(vao) if nhanh(vao) != "KHAC" else None
        n = nhanh(vao)
        if n != "KHAC":
            so_ghi.ghi(ma, n)
        return kq
    return f


def nguon_quyet_dinh():
    """Các hàm đang test là luật thuần trong mã; không đọc bảng dữ liệu nào.
    Thay đổi luật được bắt bằng van_tay_ham.py (vân tay logic hàm), không phải ở đây."""
    return {}


def du_lieu_can_cach_ly():
    return []


# ─── Điểm quyết định ───────────────────────────────────────────────────────────

def diem():
    return [
        {"ten": "vai_tro_doan", "ma": "TD.DOAN.VAITRO", "lop": "luat", "duong_chay": _duong("TD.DOAN.VAITRO", chay_vai_tro, nhanh_vai_tro), "o_dau": "SKILL/10-tender/tool/td.py:528",
         "chay": chay_vai_tro, "nhanh": nhanh_vai_tro,
         "nhan": ["VT.Mail", "VT.Scope", "VT.Làm_rõ", "VT.Tham_khảo", "VT.Ngoài_scope", "VT.NONE"],
         "moc": [], "cham": "khop_dung", "oracle": "oracle_vai_tro",
         "sai_thi_sao": "file bị xếp nhầm vai trò (Scope/Làm rõ/Ngoài scope) → bỏ sót hoặc quét thừa; không có gì chặn"},
        {"ten": "loai_trang", "ma": "TD.SPEC.LOAI_TRANG", "lop": "luat", "duong_chay": _duong("TD.SPEC.LOAI_TRANG", chay_loai_trang, nhanh_loai_trang), "o_dau": "SKILL/10-tender/tool/td.py:1880",
         "chay": chay_loai_trang, "nhanh": nhanh_loai_trang,
         "nhan": ["LT.DR", "LT.SPEC", "LT.MR", "LT.IS", "LT.ANH", "LT.NONE"],
         "moc": [], "cham": "khop_dung", "oracle": "oracle_loai_trang",
         "sai_thi_sao": "trang spec bị bỏ hoặc nhận sai → mã VL thiếu link; không có gì chặn"},
        {"ten": "tham_chieu", "ma": "TD.TCQ.THAM_CHIEU", "lop": "luat", "duong_chay": _duong("TD.TCQ.THAM_CHIEU", chay_tham_chieu, nhanh_tham_chieu), "o_dau": "SKILL/10-tender/tool/td.py:1175",
         "chay": chay_tham_chieu, "nhanh": nhanh_tham_chieu, "nhan": ["TC.CO", "TC.RONG"], "moc": [], "cham": "khop_dung",
         "sai_thi_sao": "chuỗi TCQ bị nối sai → trả lời TCQ cuối sai; không có gì chặn"},
        {"ten": "date_in", "ma": "TD.NGAY.DATE_IN", "lop": "luat", "duong_chay": _duong("TD.NGAY.DATE_IN", chay_date_in, nhanh_date_in), "o_dau": "SKILL/10-tender/tool/td.py:203",
         "chay": chay_date_in, "nhanh": nhanh_date_in, "nhan": ["DI.OK", "DI.NONE"], "moc": [],
         "cham": "khop_dung", "oracle": "oracle_date_in",
         "sai_thi_sao": "ngày issue sai → bản 'mới nhất' chọn nhầm; không có gì chặn"},
        {"ten": "dot_of", "ma": "TD.NGAY.DOT_OF", "lop": "luat", "duong_chay": _duong("TD.NGAY.DOT_OF", chay_dot_of, nhanh_dot_of), "o_dau": "SKILL/10-tender/tool/td.py:218",
         "chay": chay_dot_of, "nhanh": nhanh_dot_of, "nhan": ["DOT.GOC", "DOT.MA", "DOT.KHONG_MA"], "moc": [],
         "cham": "khop_dung", "oracle": "oracle_dot_of",
         "sai_thi_sao": "đợt/thứ tự đợt sai → TCQ/Addendum cũ đè mới; không có gì chặn"},
        {"ten": "ma_vl", "ma": "TD.VL.MA_VL", "lop": "luat", "duong_chay": _duong("TD.VL.MA_VL", chay_ma_vl, nhanh_ma_vl), "o_dau": "SKILL/10-tender/tool/td.py:84",
         "chay": chay_ma_vl, "nhanh": nhanh_ma_vl, "nhan": ["VL.CO", "VL.RONG"], "moc": [],
         "cham": "khop_dung", "oracle": "oracle_ma_vl",
         "sai_thi_sao": "mã phòng lẫn vào mã vật liệu (ST-01) hoặc ngược lại; không có gì chặn"},
        {"ten": "chuan_ben", "ma": "TD.BEN.CHUAN_BEN", "lop": "luat", "duong_chay": _duong("TD.BEN.CHUAN_BEN", chay_chuan_ben, nhanh_chuan_ben), "o_dau": "SKILL/10-tender/tool/td.py:2433",
         "chay": chay_chuan_ben, "nhanh": nhanh_chuan_ben, "nhan": ["BEN.CO", "BEN.RONG"], "moc": [], "cham": "khop_dung", "oracle": "oracle_chuan_ben",
         "sai_thi_sao": "tên bên thực hiện không chuẩn → gói gửi nhà thầu phụ thiếu dòng; soát 'goi' chặn một phần"},
        {"ten": "ma_vl_theo_bang", "ma": "TD.VL.MA_BANG", "lop": "luat", "duong_chay": _duong("TD.VL.MA_BANG", chay_ma_bang, nhanh_ma_bang), "o_dau": "SKILL/10-tender/tool/td.py:108",
         "chay": chay_ma_bang, "nhanh": nhanh_ma_bang, "nhan": ["MB.CO", "MB.RONG"], "moc": [], "cham": "khop_dung", "oracle": "oracle_ma_bang",
         "sai_thi_sao": "hạng mục nhiều mã VL viết liền (WO1, FR1) chỉ hiện 1 mã, hoặc mã FF&E lẫn vào Mã VL (lỗi 09/11); không có gì chặn"},
        {"ten": "tap_ma", "ma": "DMVT.TAP_MA", "lop": "luat", "duong_chay": _duong("DMVT.TAP_MA", chay_tap_ma, nhanh_tap_ma), "o_dau": "SKILL/10-tender/tool/dmvt_xlsx.py:49",
         "chay": chay_tap_ma, "nhanh": nhanh_tap_ma, "nhan": ["TM.CO", "TM.RONG"], "moc": [], "cham": "khop_dung",
         "sai_thi_sao": "dòng DMVT có mã viết liền bị coi là dòng đuôi -> nhân đôi mỗi lần build (lỗi 09/11); không có gì chặn"},
        {"ten": "doc_ncc", "ma": "SPEC.DOC_NCC", "lop": "luat", "duong_chay": _duong("SPEC.DOC_NCC", chay_doc_ncc, nhanh_doc_ncc), "o_dau": "SKILL/10-tender/tool/spec_boc.py:114",
         "chay": chay_doc_ncc, "nhanh": nhanh_doc_ncc, "nhan": ["NCC.CO", "NCC.RONG"], "moc": [], "cham": "khop_dung",
         "sai_thi_sao": "NCC không nhãn bị bỏ hoặc dán nhầm -> DMVT trống NCC (lỗi CTX 09/11); không có gì chặn"},
    ]


# ─── Case ──────────────────────────────────────────────────────────────────────

def _c(id_, diem_, vao, cho=None, bien=None, nguon="nguoi"):  # người duyệt 83 case ngày 2026-11-06
    c = {"id": id_, "diem": diem_, "vao": vao, "so": None, "nguon": nguon, "vai": "user"}
    if cho is not None:
        c["cho"] = cho
    if bien:
        c["bien"] = bien
    return c


def _case_vai_tro():
    d = "vai_tro_doan"
    L = [  # (id, rel, inner, ext, mong đợi tay)
        ("VT01", "01.260101/mail.msg", None, ".msg", "Mail"),
        ("VT02", "x/PVCV_Room.xlsx", None, ".xlsx", "Scope"),
        ("VT03", "x/a pvcv.pdf", None, ".pdf", "Scope"),
        ("VT04", "x/pvcvs.xlsx", None, ".xlsx", "NONE"),          # pvcv dính chữ sau
        ("VT05", "x/pvcv.dwg", None, ".dwg", "NONE"),             # đuôi không thuộc scope
        ("VT06", "x/TCQ01.xlsx", None, ".xlsx", "Làm rõ"),
        ("VT07", "x/Tender Queries.xlsx", None, ".xlsx", "Làm rõ"),
        ("VT08", "x/TCQ01.docx", None, ".docx", "NONE"),          # TCQ nhưng đuôi docx
        ("VT09", "x/atcq.xlsx", None, ".xlsx", "NONE"),           # 'atcq' không có ranh giới trái
        ("VT10", "x/AS-BUILT/a.pdf", None, ".pdf", "Tham khảo"),
        ("VT11", "x/asbuilt/a.pdf", None, ".pdf", "Tham khảo"),
        ("VT12", "x/mep.zip", None, ".zip", "Ngoài scope"),
        ("VT13", "mep.zip", None, ".zip", "Ngoài scope"),
        ("VT14", "x/amep.zip", None, ".zip", "NONE"),
        ("VT15", "x/hvac/a.pdf", None, ".pdf", "Ngoài scope"),
        ("VT16", "x/Project Schedule.pdf", None, ".pdf", "Tham khảo"),
        ("VT17", "x/Annex 3.pdf", None, ".pdf", "Tham khảo"),
        ("VT18", "x/random.pdf", None, ".pdf", "NONE"),
        ("VT19", "x/a.zip", "elek/inside.pdf", ".pdf", "Ngoài scope"),
        ("VT20", "x/a.zip", "sub/elek/inside.pdf", ".pdf", "Ngoài scope"),
        ("VT21", "x/a.zip", "PVCV.xlsx", ".xlsx", "Scope"),       # tên lấy từ inner
        ("VT22", "x/PVCV.xlsx", "other.xlsx", ".xlsx", "NONE"),   # inner đè tên
        ("VT24", "x/TCQ01.pdf", None, ".pdf", "Làm rõ"),
        ("VT23", "x/Rules and Regulations.pdf", None, ".pdf", "Tham khảo"),
    ]
    return [_c(i, d, {"rel": r, "inner": n, "ext": e}, m) for i, r, n, e, m in L]


def _case_loai_trang():
    d = "loai_trang"
    L = [
        ("LT01", "DRAWING REFERENCE: A", "DR"), ("LT02", "  SPECIFICATION foo", "SPEC"),
        ("LT03", "MATERIAL REFERENCE x", "MR"), ("LT04", "\n\nITEM SCHEDULE", "IS"),
        ("LT05", "Image (s) provided here", "ANH"), ("LT06", "header\nSPECIFICATION", "NONE"),
        ("LT07", "SPECIFICATION\nDRAWING REFERENCE: B", "DR"), ("LT08", "", "NONE"),
        ("LT09", "specification", "NONE"), ("LT10", "   ", "NONE"),
    ]
    return [_c(i, d, {"t": t}, m) for i, t, m in L]


def _case_tham_chieu():
    d = "tham_chieu"
    L = [
        ("TC01", "Refer TCQ 03-12 for details", "TCQ03-12"),
        ("TC02", "see TC#5", "TCQ05"),
        ("TC03", "as per TCQ.07", "TCQ07"),
        ("TC04", "item 4 of TCQ 02", "TCQ02-4"),
        ("TC05", "Refer TCQ 03-12 and see TCQ 03-12", "TCQ03-12"),
        ("TC06", "no reference here", "(rong)"),
        ("TC07", "response to TCQ 11", "TCQ11"),
        ("TC08", "refer TCQ 1 then see TCQ 2/3", "TCQ01|TCQ02-3"),
    ]
    return [_c(i, d, {"text": t}, m) for i, t, m in L]


def _case_date_in():
    d = "date_in"
    L = [
        ("DI01", ["a", "20261105_x"], "2026-11-05"),
        ("DI02", ["20250101", "b"], "2025-01-01"),
        ("DI03", ["261105_issue"], "2026-11-05"),
        ("DI04", ["261105-issue"], "2026-11-05"),
        ("DI05", ["20261305"], "NONE"),            # tháng 13 không hợp lệ
        ("DI06", ["201234567"], "NONE"),           # 9 chữ số
        ("DI07", ["x", "y"], "NONE"),
        ("DI08", [], "NONE"),
        ("DI09", ["20260101", "20261231"], "2026-12-31"),  # ưu tiên phần trong cùng
        ("DI10", ["261301_x"], "NONE"),
        ("DI11", ["19991231"], "NONE"),           # không bắt đầu 20
    ]
    return [_c(i, d, {"parts": p}, m) for i, p, m in L]


def _case_dot_of():
    d = "dot_of"
    L = [
        ("DO01", "file.pdf", "(gốc)|None|0"),
        ("DO02", "10.261026_TCQ07/a.pdf", "10.261026_TCQ07|2026-10-26|10"),
        ("DO03", "7.261026/a.pdf", "7.261026|2026-10-26|7"),
        ("DO04", "10.261326/a.pdf", "10.261326|None|10"),
        ("DO05", "docs/a.pdf", "docs|None|999"),
        ("DO06", "10.2610/a.pdf", "10.2610|None|999"),
        ("DO07", "01.260101\\sub\\a.pdf", "01.260101|2026-01-01|1"),
        ("DO08", "a/", "a|None|999"),
    ]
    return [_c(i, d, {"rel": r}, m) for i, r, m in L]


def _case_ma_vl():
    d = "ma_vl"
    L = [
        ("VL01", {"text": "Use ST-02 finish"}, "ST-02"),
        ("VL02", {"text": "Standard room ST-01", "rc": ["ST-01"], "tien_to": ["ST"]}, "(rong)"),   # có chữ room gần mã -> là phòng
        ("VL03", {"text": "ST-01 black forest stone", "rc": ["ST-01"], "tien_to": ["ST"]}, "ST-01"),
        ("VL04", {"text": "ST-01 black forest stone", "rc": ["ST-01"], "tien_to": []}, "(rong)"),
        ("VL05", {"text": "XST-01"}, "XST-01"),
        ("VL06", {"text": "ABCD-01"}, "(rong)"),
        ("VL07", {"text": "ST-011"}, "(rong)"),
        ("VL08", {"text": "ST-01A"}, "ST-01A"),
        ("VL09", {"text": "ST-01AB"}, "(rong)"),
        ("VL10", {"text": "-ST-01"}, "(rong)"),
        ("VL11", {"text": "(ST-01), PT-02"}, "PT-02|ST-01"),
        ("VL12", {"text": ""}, "(rong)"),
        ("VL14", {"text": "room 0123456789012345 ST-01", "rc": ["ST-01"], "tien_to": ["ST"]}, "(rong)"),  # từ phòng cách mã ~18 ký tự
        ("VL13", {"text": "A-01"}, "(rong)"),
    ]
    return [_c(i, d, v, m) for i, v, m in L]


def _case_chuan_ben():
    d = "chuan_ben"
    g = {"Joinery": "Mộc", "joinery A": "Mộc", "a;b": "AB"}
    L = [
        ("BN01", "Joinery", "Mộc"),
        ("BN02", "Joinery; Joinery A", "Mộc; Joinery A"),  # 'Joinery A' khác 'joinery A' (phân biệt hoa/thường)
        ("BN03", "joinery A; Joinery", "Mộc"),
        ("BN04", "a;b", "AB"),                 # cả ô khớp trước
        ("BN05", "x; ; y", "x; y"),
        ("BN06", "", ""),
        ("BN07", None, ""),
        ("BN08", "  Joinery  ", "Mộc"),
    ]
    return [_c(i, d, {"raw": r, "gop": g}, m) for i, r, m in L]


def bo_test():
    return (_case_vai_tro() + _case_loai_trang() + _case_tham_chieu() + _case_date_in()
            + _case_dot_of() + _case_ma_vl() + _case_chuan_ben()
            + _case_ma_bang() + _case_tap_ma() + _case_doc_ncc())


def _case_ma_bang():
    d = "ma_vl_theo_bang"
    B = ["WO1", "FR1", "SP1", "LM4A", "LM2", "LF-02", "CG-2102", "PT1A", "G1"]
    L = [  # đáp án tay, giả định theo mô tả BOQ Blanca STUDIO!66 và các biên
        ("MB01", "Bàn sofa LF-02 veneer WO1 nẹp WO1 chân sơn SP1", "LF-02|WO1|SP1"),
        ("MB02", "theo mẫu(LM4a). - khung", "LM4A"),
        ("MB03", "Standing table CG-2102 MT-02", "(rong)"),
        ("MB04", "vải HN-UP-WO1 dệt", "(rong)"),
        ("MB05", "WO1x và WO12", "(rong)"),
        ("MB06", "ván FR1_A", "(rong)"),
        ("MB07", "mã WO2 chưa có trong bảng", "(rong)"),
        ("MB08", "", "(rong)"),
        ("MB09", "sơn PT1a, trần G1; tường LM2", "PT1A|G1|LM2"),
        ("MB10", "chống mài mòn AC4 tỷ trọng", "(rong)"),
        ("MB11", "LM2-ẩm", "(rong)"),
    ]
    return [_c(i, d, {"text": t, "biet": B}, m, nguon="gia_dinh") for i, t, m in L]


def _case_tap_ma():
    d = "tap_ma"
    L = [  # ô cột B của DMVT; đáp án tay, giả định
        ("TM01", "FR1", "FR1"),
        ("TM02", "LM4A", "LM4A"),
        ("TM03", "CU-01" + chr(10) + "CU-02" + chr(10) + "CU-03", "CU-01|CU-02|CU-03"),
        ("TM04", "D-WC", "D-WC"),
        ("TM05", "401PL", "401PL"),
        ("TM06", "SX-2200", "SX-2200"),
        ("TM07", "FF&E", "(rong)"),
        ("TM08", "FINISHES – VẬT LIỆU HOÀN THIỆN", "(rong)"),
        ("TM09", "I", "(rong)"),
        ("TM10", "22", "(rong)"),
        ("TM11", "", "(rong)"),
        ("TM12", "Gối trang trí 3", "(rong)"),
    ]
    return [_c(i, d, {"o": o}, m, nguon="gia_dinh") for i, o, m in L]


def _case_doc_ncc():
    d = "doc_ncc"
    L = [  # dòng khối NCC theo bố cục thật trang CTX (LF_7221_A p14, SW_7300_C p2); đáp án tay
        ("NC01", ["St Vincents", "https://stvincents.co/product/anna-karlin-", "counter-height-chess-stools/", "T: +32 (0) 492314159",
                  "E: hello@stvincents.com", "Indicative lead time: To be verified"],
         "St Vincents|https://stvincents.co/product/anna-karlin-counter-height-chess-stools/|T: +32 (0) 492314159 / E: hello@stvincents.com"),
        ("NC02", ["LIXIL INTERNATIONAL", "or Equivalent", "URL: www.lixil.com", "C: Theresa Zhou", "theresa.zhou@lixil.com"],
         "LIXIL INTERNATIONAL or Equivalent|www.lixil.com|C: Theresa Zhou / theresa.zhou@lixil.com"),
        ("NC03", ["RIGEL or Equivalent URL:", "Waster Closet Sense Flush"], "RIGEL or Equivalent||"),
        ("NC04", ["Manufacturer: Poliform", "T: 123"], "{}"),
        ("NC05", ["Contractor to Bid", "Contractor to Bid"], "Contractor to Bid||"),
        ("NC06", ["Hafary", "105 Eunos Avenue 3"], "Hafary||105 Eunos Avenue 3"),
        ("NC07", ["TECHNICAL DRAWINGS", "All dimensions in mm"], "{}"),
        ("NC08", [], "{}"),
        ("NC09", ["Indicative lead time: 12 weeks", "Some Brand"], "{}"),
        ("NC10", [": U1 Floor Pte Ltd"], "U1 Floor Pte Ltd||"),
    ]
    return [_c(i, d, {"dong": x}, m, nguon="gia_dinh") for i, x, m in L]
