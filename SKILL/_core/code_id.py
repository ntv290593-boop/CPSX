"""Phân loại Code ID theo lớp (plan DOCS\\PLAN_TENDER-CODEID-DMVT_2026-10-31.md, mục 3).

Thứ tự, dừng ở lớp đầu tiên có kết quả:
  User > Luật > Nhớ-mô-tả > Nhớ-đầu-mục > LLM (câu đã trả lời trong cùng dự án) > chờ LLM.
Trong mỗi lớp nhớ: nho_local.json đọc trước nho.json.

Dữ liệu ở _MASTER\\code_id: bang_ma.json, luat.json, nho.json (chỉ người quản lý sửa), nho_local.json (của từng máy).
Bộ nhớ chỉ chứa mẫu chữ → mã: số đã thay bằng #, không có khối lượng, giá, mã phòng.
"""
import json
import os
import re

KHONG_RO = "KHÔNG RÕ"
NGUON = ("User", "Luật", "Nhớ-mô-tả", "Nhớ-đầu-mục", "LLM")
HERE = os.path.dirname(os.path.realpath(__file__))  # realpath: chạy qua junction .claude\skills vẫn ra đúng thư mục SKILL


def thu_muc():
    """TD_CODE_ID_DIR chỉ dành cho kiểm thử, để không ghi vào _MASTER thật."""
    return os.environ.get("TD_CODE_ID_DIR") or os.path.normpath(os.path.join(HERE, "..", "..", "_MASTER", "code_id"))


def _doc(ten, mac_dinh=None):
    path = os.path.join(thu_muc(), ten)
    if not os.path.exists(path):
        return mac_dinh
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


# ---------------------------------------------------------------- bảng mã
def rut_ma(ma):
    """So khớp mã, bỏ khoảng trắng (Datawork G57 ghi 'LS. HA'). Giá trị xuất ra vẫn đúng như bảng mã."""
    return re.sub(r"\s+", "", str(ma or ""))


class BangMa:
    def __init__(self, rows):
        self.rows = rows
        self.ten = {rut_ma(r["ma"]): r["cap3"] for r in rows}
        self.ma = {r["cap3"]: r["ma"] for r in rows}

    def ten_cua(self, ma):
        return self.ten.get(rut_ma(ma))

    def co(self, ma):
        return rut_ma(ma) in self.ten


def nap_bang():
    d = _doc("bang_ma.json")
    return BangMa(d["rows"]) if d else None


# ---------------------------------------------------------------- khóa nhớ
def _gon(t):
    return re.sub(r"\s+", " ", str(t or "").lower()).strip()


def so_hoa(t):
    """Số thành #, gồm cả số trong mã (ST-04 → st-#), để bộ nhớ không giữ số liệu dự án."""
    return re.sub(r"\d+(?:[.,]\d+)*", "#", _gon(t))


_RX_REFER = re.compile(r"[;,]?\s*(?:please\s+)?refer(?:\s+to)?\b.*$", re.I)
_RX_KT = re.compile(r"\(?\b[dwlh]?\d[\d.,+x×*()\s]*(?:mm|cm|m)?h?\b\)?", re.I)
_TU_KT = {"overall", "size", "sizes", "dimension", "dimensions", "type", "mm", "mmh", "h", "x", "nr", "approx", "thick",
          "thickness", "wide", "width", "high", "height", "depth", "deep", "long", "length", "diameter", "dia", "d", "w", "l"}


def chi_kich_thuoc(mo_ta):
    """Dòng con chỉ nêu kích thước, phải có ít nhất 1 chữ số (không nói vật liệu hay vị trí)."""
    t = _RX_REFER.sub("", str(mo_ta or ""))
    if not re.search(r"\d", t):
        return False
    t = _RX_KT.sub(" ", t)
    tu = re.findall(r"[^\W\d_]+", t.lower())
    return all(x in _TU_KT for x in tu)


def khoa_mo_ta(r):
    return _gon(r.get("Nhóm")), so_hoa(r.get("Mô tả BOQ"))


_RX_FFE_LOAI = re.compile(r"\s*\d{2}\s*-\s*(\d)\d{2}")


def tien_to_ma(r):
    """Họ mã của chính dòng: 'ST' từ ST-04, 'F4' từ 51-400, '' nếu dòng không có mã. Dòng nhiều mã lấy mã đầu."""
    vl = str(r.get("Mã VL") or "").split(";")[0].strip()
    if vl:
        return vl.split("-")[0].upper()
    m = _RX_FFE_LOAI.match(str(r.get("Mô tả BOQ") or ""))
    return "F" + m.group(1) if m else ""


def khoa_dau_muc(r):
    """Cùng đầu mục nhưng khác họ mã thường khác công tác (đá ốp tường, gỗ ốp tường), nên khóa có cả họ mã."""
    return _gon(r.get("Nhóm")), so_hoa(r.get("Mục phụ")), tien_to_ma(r)


def khoa_llm(r):
    """Câu LLM dùng lại trong dự án. Có đầu mục để 'overall size #' ở 2 món khác nhau không dùng chung câu trả lời."""
    return _gon(r.get("Nhóm")), so_hoa(r.get("Mục phụ")), so_hoa(r.get("Mô tả BOQ"))


# ---------------------------------------------------------------- luật
class BangLuat:
    """luat.json: danh sách theo thứ tự ưu tiên, {"ten", "ma", "dk": [{"cot": [...], "rx": "..."}]}. Mọi dk phải khớp."""

    def __init__(self, rows):
        self.rows = [dict(x, dk=[dict(d, _rx=re.compile(d["rx"], re.I)) for d in x["dk"]]) for x in rows]

    def xet(self, r):
        for x in self.rows:
            if all(any(d["_rx"].search(str(r.get(c) or "")) for c in d["cot"]) for d in x["dk"]):
                return x["ma"], x["ten"]
        return None, None


def nap_luat():
    d = _doc("luat.json")
    return BangLuat(d["rows"] if d else [])


# ---------------------------------------------------------------- bộ nhớ
class Nho:
    """{khoa: {"ma", "du_an"}} cho 2 loại mô tả / đầu mục."""

    def __init__(self, d=None):
        d = d or {}
        self.mo_ta = {(x["nhom"], x["mo_ta"]): x for x in d.get("mo_ta", [])}
        self.dau_muc = {(x["nhom"], x["dau_muc"], x.get("tien_to", "")): x for x in d.get("dau_muc", [])}


def nap_nho():
    """Trả (local, chung)."""
    return Nho(_doc("nho_local.json")), Nho(_doc("nho.json"))


def hoc(dong, du_an=""):
    """Dòng có nhãn ['_ma'] → dict dạng nho.json. Khóa có nhiều nhãn thì lấy nhãn nhiều nhất, bằng nhau thì bỏ.

    Dòng chỉ có kích thước không vào nhớ mô tả ('overall size #x#' không nói được là món gì).
    """
    from collections import Counter, defaultdict
    a, b = defaultdict(Counter), defaultdict(Counter)
    for r in dong:
        if not chi_kich_thuoc(r.get("Mô tả BOQ")):
            a[khoa_mo_ta(r)][r["_ma"]] += 1
        if r.get("Mục phụ"):
            b[khoa_dau_muc(r)][r["_ma"]] += 1

    def chon(c):
        top = c.most_common(2)
        return None if len(top) == 2 and top[0][1] == top[1][1] else top[0][0]
    return {"mo_ta": [{"nhom": k[0], "mo_ta": k[1], "ma": chon(c), "du_an": du_an} for k, c in a.items() if chon(c)],
            "dau_muc": [{"nhom": k[0], "dau_muc": k[1], "tien_to": k[2], "ma": chon(c), "du_an": du_an} for k, c in b.items() if chon(c)]}


def ghi_nho_local(r, ma, du_an):
    """Ô user sửa droplist → nho_local.json. Trả True nếu có đổi. Dòng chỉ kích thước nhớ theo đầu mục."""
    path = os.path.join(thu_muc(), "nho_local.json")
    d = _doc("nho_local.json") or {"mo_ta": [], "dau_muc": []}
    if chi_kich_thuoc(r.get("Mô tả BOQ")):
        if not r.get("Mục phụ"):
            return False
        k = khoa_dau_muc(r)
        ds, khoa = d["dau_muc"], {"nhom": k[0], "dau_muc": k[1], "tien_to": k[2]}
    else:
        ds, khoa = d["mo_ta"], {"nhom": khoa_mo_ta(r)[0], "mo_ta": khoa_mo_ta(r)[1]}
    cu = next((x for x in ds if all(x.get(k) == v for k, v in khoa.items())), None)
    if cu and cu["ma"] == ma and cu.get("du_an") == du_an:
        return False
    if cu:
        cu.update(ma=ma, du_an=du_an)
    else:
        ds.append(dict(khoa, ma=ma, du_an=du_an))
    os.makedirs(thu_muc(), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(d, fh, ensure_ascii=False, indent=1)
    return True


# ---------------------------------------------------------------- phân loại 1 dòng
def phan_loai(r, luat, local, chung, llm_du_an=None, du_an="", user=None):
    """→ (mã hoặc KHONG_RO hoặc None, nguồn hoặc None, từ dự án khác?).

    user: mã user đã chọn cho đúng dòng này. llm_du_an: {khoa_llm: mã hoặc KHONG_RO} LLM đã trả lời trong dự án này.
    Dòng chỉ có kích thước bỏ qua nhớ mô tả, đi thẳng xuống nhớ đầu mục.
    """
    if user:
        return user, "User", False
    ma, _ = luat.xet(r)
    if ma:
        return ma, "Luật", False
    if not chi_kich_thuoc(r.get("Mô tả BOQ")):
        k = khoa_mo_ta(r)
        for nho in (local, chung):
            if k in nho.mo_ta:
                x = nho.mo_ta[k]
                return x["ma"], "Nhớ-mô-tả", x.get("du_an") != du_an
    if r.get("Mục phụ"):
        k = khoa_dau_muc(r)
        for nho in (local, chung):
            if k in nho.dau_muc:
                x = nho.dau_muc[k]
                return x["ma"], "Nhớ-đầu-mục", x.get("du_an") != du_an
    if llm_du_an and khoa_llm(r) in llm_du_an:
        return llm_du_an[khoa_llm(r)], "LLM", False
    return None, None, False


def co_highlight(nguon, ma, tu_du_an_khac, luat_da_duyet):
    """Plan mục 3: chỗ máy không chắc. User không bao giờ highlight."""
    if nguon == "User":
        return ""
    if ma is None:
        return "Chờ LLM"
    if ma == KHONG_RO:
        return KHONG_RO
    if nguon == "Luật":
        return "" if luat_da_duyet else "Luật chưa duyệt"
    if nguon == "Nhớ-mô-tả":
        return "Dự án khác" if tu_du_an_khac else ""
    return nguon  # Nhớ-đầu-mục, LLM
