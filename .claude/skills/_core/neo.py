# -*- coding: utf-8 -*-
r"""neo.py — kiểm giá trị LLM ghi có thật trên trang được link không (chỉ XEM, không chặn build).
Chép từ 01-scan pd.py (đóng băng nên giữ bản riêng); các hàm tách số, tách mã giữ nguyên, kiemthu.py so với pd.py.
Không biết form, bước. Người gọi dựng danh sách item và đưa nguồn chữ:

  item = dict(sh, dong, cot, ten, kieu 'to'|'spec'|'ma'|'kt', nhan, file (FID), trang, o, ma, [nhom = [(FID, trang, ô)], ho])
  nguon: text(fid, trang) -> chữ | None; so_trang(fid) -> n; la_pdf(fid) -> bool; duong_dan(fid) -> file xlsx
  run(items, nguon, trang_bv)  trang_bv = [(fid, trang, số BV)]
"""
import datetime
import hashlib
import json
import os
import re
import unicodedata
from collections import Counter

from scan_core import _fold

NEO_BO = re.compile(r"^(f\d{3,4}|p\d+|r\d+|\d+(?:[.,]\d+)?(?:mm|cm|m|m2|m3|md|kg|ly|li|nr|no|pcs|bo|cai|set)?"  # file/trang/dòng, số + đơn vị
                    r"|[a-z]?\d+(?:[.,]\d+)?(?:x[a-z]?\d+(?:[.,]\d+)?)+(?:mm|cm|m)?[whdl]?)$")  # kích thước 900x650x725mmH: để N3.3
NEO_ITEM = re.compile(r"^\d{1,3}-\d{3,4}[A-Za-z]{0,2}$")  # mã hạng mục FF&E không có chữ: 51-400, 53-701A (30-31 không nhận)
NEO_STOPW = {"DETAIL", "PLAN", "SECTION", "ELEVATION", "TYPICAL", "AND", "THE", "WITH"}
NEO_KK = "KHÔNG KIỂM ĐƯỢC"
# mã VL dạng WD-01, GL1, LM2a; không dính '-' hai đầu nên đoạn trong số tờ (ID-SK-52-101) không tính
NEO_MAVL = re.compile(r"(?<![A-Za-z0-9-])([A-Z]{1,3})-?(\d{1,3})([A-Za-z]?)(?![A-Za-z0-9]|[-./]\d)")
NEO_HO_BO = {"F", "P", "R", "B", "S", "V", "D", "Q", "BT", "DT"}  # file/trang/dòng, ID issue các bước — không phải họ mã VL
NEO_MA_CONG = 2  # trang có ≥ 2 mã VL khác nhau dạng text: kiểm được mọi họ; ít hơn: chỉ họ có trên trang
NEO_KT_MIN, NEO_KT_CONG, NEO_LAP = 100, 10, 0.3  # số ≥ 100 · trang cần ≥ 10 số như vậy · dòng chữ lặp ở ≥ 30% trang = khung tên
NEO_LEV = re.compile(r"^\s*(?:[+±-]\s*\d+[.,]\d{3}\s*$|.*\b(?:LEV|FFL|FCL|SSL|EL|CL|TOS|AFF)\b)", re.I)  # cao độ: không tính vào cổng
NEO_FOLD = {"×": "x", "X": "x", "–": "-", "—": "-", "−": "-", "＋": "+", " ": " "}
NEO_JUNK = [
    re.compile(r"\b\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}\b"),                    # ngày
    re.compile(r"\b\d+\s*[:/]\s*\d+\b"),                                     # tỉ lệ, phân số, giờ
    re.compile(r"\+\s*\(?\d{1,3}\)?[\s.\-]?\(?\d{1,4}\)?(?:[\s.\-]?\d{2,4}){2,}"),  # điện thoại có +
    re.compile(r"\(\d{2,4}\)\s*\d{3,4}[\s.\-]\d{3,4}"),                      # điện thoại (028) 3824 1525
    re.compile(r"#\s*\d[\d\-]*"),                                            # #1-02
    re.compile(r"\b[A-Za-z][\w]*[\-.]\w*\d[\w\-.]*"),                        # mã dính chữ: IN-PS2-501, EN2-4
    re.compile(r"\b\d+(?:\.\d+){2,}\b"),                                     # 931.84.233
    re.compile(r"(?i)\brev\.?\s*\d+"),
]
_NEO_DINH = r"0-9A-Za-wyz_À-ỹ"  # ký tự dính số làm số thành mã; trừ x để 1200x600 vẫn tách
_NEO_N0 = r"(?<![" + _NEO_DINH + r".,#/:])(?:[WDHLRS]|Ø)?" + r"(\d+(?:\.\d+)?)(?:\s?(?:mm|cm))?(?![" + _NEO_DINH + r"#/:]|\.\d)"
NEO_NUM = re.compile(_NEO_N0)
NEO_CHAIN = re.compile(_NEO_N0.replace("(\\d", "(?:\\d") + r"(?:\s*[x*+=]\s*" + _NEO_N0.replace("(\\d", "(?:\\d") + r")+")  # 1200x600, 450+450=900
_RX_O = r"^'?(.+?)'?!\$?[A-Z]{1,3}\$?(\d+)(?::\$?[A-Z]{1,3}\$?(\d+))?$"


def neo_chuan(s):
    s = "".join(NEO_FOLD.get(c, c) for c in unicodedata.normalize("NFKC", str(s)))
    s = re.sub(r"(?<=\d),(?=\d{3}(?!\d))", "", s)                       # 1,200 -> 1200
    s = re.sub(r"(?<![\d.])(\d{1,3})\.(\d{3})(?![.\d])", r"\1\2", s)    # 1.200 -> 1200
    for j in NEO_JUNK:
        s = j.sub(" ¦ ", s)
    return s


def neo_so1(t):
    if len(t) > 1 and t[0] == "0" and not t.startswith("0."):
        return None  # 0 đứng đầu: rác
    return t[:-2] if t.endswith(".0") else t


def neo_tach_so(s):
    """Dòng số sạch theo thứ tự đọc."""
    return [v for cell in neo_chuan(s).split("¦") for v in (neo_so1(m.group(1)) for m in NEO_NUM.finditer(cell)) if v]


def neo_kt_gt(v):
    """Giá trị 'Kích thước BV' -> các số ≥ 100 của chuỗi R×S×C ĐẦU TIÊN. Phần diễn giải sau đó không kiểm
    (đo 30/10: kiểm mọi số trong câu ra 20/46 LỆCH oan, chỉ chuỗi đầu ra 4/44)."""
    v = re.sub(r"\s*\([^()\d]{1,12}\)", "", str(v or "")).replace("±", "").replace("~", "")  # '1540 (R) × 2375 (C)', '±130'
    v = re.sub(r"[x×X*]\s*\?\s*(?=[x×X*])", "", v)                                        # '900×?×440': bỏ chiều chưa rõ
    for cell in neo_chuan(v).split("¦"):
        for m in NEO_CHAIN.finditer(cell):
            p = [neo_so1(x) for x in NEO_NUM.findall(m.group(0))]
            if len(p) >= 2 and all(p):
                return [x for x in p if float(x) >= NEO_KT_MIN]
    return []


def neo_mavl(s, ho):
    """Mã VL thuộc họ mã dự án trong chuỗi -> {khóa chuẩn: chữ gốc}. 'WD-01' và 'WD1' cùng khóa WD1."""
    return {f"{a}{int(b)}{c.upper()}": m.group(0) for m in NEO_MAVL.finditer(str(s or "")) for a, b, c in [m.groups()] if a in ho}


def neo_ho(vals):
    """Họ mã VL của dự án (WD, PT, GL…) từ các giá trị có mã; họ phải gặp ≥ 2 lần, bỏ tiền tố file/trang/issue."""
    c = Counter()
    for v in vals:
        for m in NEO_MAVL.finditer(str(v or "")):
            c[m.group(1)] += 1
    return {h for h, n in c.items() if n >= 2 and h not in NEO_HO_BO}


def _an(s):
    return re.sub(r"[^a-z0-9]", "", _fold(s))


def neo_ma(*vals, file=""):
    """Token dạng mã/model (có cả chữ và số, ≥ 3 ký tự chữ-số): WO1, PT-04, MK630, 5040-Y90B; mã FF&E 51-400.
    Bỏ F024/p11/r16/18mm/2nr, kích thước 900x650mmH, và token thuộc tên file link."""
    out, ten = [], _an(os.path.splitext(os.path.basename(file))[0]) if file else ""
    for v in vals:
        for t in re.split(r"[\s,;()\[\]/:–—]+", str(v or "")):
            t = t.strip(".-")
            k = _an(t)
            if len(k) < 3 or not re.search(r"\d", k) or (ten and k in ten):
                continue
            if NEO_ITEM.match(t) or (re.search(r"[a-z]", k) and not NEO_BO.match(k)):
                out.append(t)
    return list(dict.fromkeys(out))


def neo_hash(items):
    return hashlib.md5(json.dumps([[x.get(k) for k in ("sh", "dong", "cot", "nhan", "file", "trang", "o", "ma")] for x in items],
                                  ensure_ascii=False, default=str).encode("utf-8")).hexdigest()


def neo_to_index(trang_bv):
    """[(fid, trang, số BV)] -> (số BV của trang, {đuôi chữ-số của số BV: {số BV}}, {số BV: [trang]}). Đuôi tách theo -/_/khoảng trắng:
    'HC-VTTEV-1BR-504.4' có đuôi 5044, 1br5044, … nên nhãn ghi gọn '504.4' hay đủ đều nhận ra."""
    so, duoi, tc = {}, {}, {}
    for fid, pg, sbv in trang_bv:
        sbv = str(sbv or "").strip()
        if not sbv:
            continue
        p = (fid, int(pg))
        so[p] = sbv
        tc.setdefault(sbv, []).append(p)
        seg = [x for x in re.split(r"[-_\s]+", sbv) if x]
        for j in range(len(seg)):
            k = _an("".join(seg[j:]))
            if len(k) >= 3 and re.search(r"\d", k):
                duoi.setdefault(k, set()).add(sbv)
    return so, duoi, tc


def neo_so_to(nhan, duoi):
    """Số tờ nhắc trong nhãn link -> ({số BV}, {đuôi đã khớp}). Mỗi token lấy đuôi dài nhất có trong danh sách; bỏ F024 (file), p17 (trang)."""
    want, keys = set(), set()
    for t in re.split(r"[\s,;·/()]+", str(nhan)):
        seg = [x for x in re.split(r"[-_]+", t) if x]
        for j in range(len(seg)):
            k = _an("".join(seg[j:]))
            if re.fullmatch(r"f\d{3,4}|p\d+", k):
                break
            if k in duoi:
                want |= duoi[k]
                keys.add(k)
                break
    return want, keys


_NEO_RX = {}


def neo_rx(x):
    """Số tờ trên text trang, cho phép ký tự ngăn giữa các ký tự nhưng không dính đuôi: 709 ≠ 709A."""
    if x not in _NEO_RX:
        _NEO_RX[x] = re.compile(r"(?<![a-z0-9])" + r"[\W_]*".join(_an(x)) + r"(?![a-z0-9])")
    return _NEO_RX[x]


def run(items, ng, trang_bv):
    """Kiểm từng item -> dict(hash, tong, theo_kieu, khong_kiem, lech)."""
    so, duoi, tc = neo_to_index(trang_bv)
    sq, xl, lap = {}, {}, {}
    tong, kk, lech, tk = Counter(), Counter(), [], {}

    def sq_of(fid, pg):
        if (fid, pg) not in sq:
            sq[(fid, pg)] = _an(ng.text(fid, pg) or "")
        return sq[(fid, pg)]

    def xl_rows(fid, sheet):
        """Dòng của 1 sheet Excel -> {số dòng: chữ-số cả dòng}."""
        if (fid, sheet) not in xl:
            import openpyxl
            try:
                wb = openpyxl.load_workbook(ng.duong_dan(fid), read_only=True, data_only=True)
                xl[(fid, sheet)] = {i: _an(" ".join(str(c) for c in row if c is not None)) for i, row in enumerate(wb[sheet].iter_rows(values_only=True), 1)}
                wb.close()
            except Exception:
                xl[(fid, sheet)] = None
        return xl[(fid, sheet)]

    def dich(fid, pg, o):
        """1 link -> (chữ-số của trang PDF / các dòng Excel được link, None) hoặc (None, lý do KHÔNG KIỂM ĐƯỢC)."""
        if ng.la_pdf(fid) and str(pg).isdigit():
            t = ng.text(fid, int(pg))
            return (sq_of(fid, int(pg)), None) if t is not None and len(t.strip()) >= 20 else (None, "trang không có lớp text")
        m_ = re.match(_RX_O, str(o))
        if not m_ or not ng.la_xlsx(fid):
            return None, "link không chỉ tới trang PDF / ô Excel"
        rows = xl_rows(fid, m_.group(1))
        if not rows:
            return None, "không mở được sheet Excel"
        return "".join(rows.get(i, "") for i in range(int(m_.group(2)), int(m_.group(3) or m_.group(2)) + 1)), None

    def ket(x, st, why="", goi_y=()):
        tong[st] += 1
        tk.setdefault(x["kieu"], Counter())[st] += 1
        if st == NEO_KK:
            kk[why] += 1
        if st == "LỆCH":
            lech.append(dict(sh=x["sh"], dong=x["dong"], cot=x["cot"], ten=x["ten"], nhan=x["nhan"], link=x["link"], ly_do=why, goi_y=list(goi_y)[:3]))

    def lap_of(fid):
        """Dòng có chữ lặp ở ≥ 30% trang của file (khung tên; file ≥ 5 trang). Dòng chỉ có số là kích thước, giữ."""
        if fid not in lap:
            n = ng.so_trang(fid)
            c = Counter()
            for q in range(1, n + 1):
                c.update({l.strip() for l in (ng.text(fid, q) or "").splitlines() if l.strip()})
            lap[fid] = {l for l, m in c.items() if n >= 5 and m >= NEO_LAP * n and re.search(r"[A-Za-z]", l)}
        return lap[fid]

    def neo3(fid, pg, kieu, ho):
        """Trang -> (tập neo, cổng). 'ma': khóa mã VL thuộc họ dự án trừ mã trong số tờ; cổng = các họ kiểm được.
        'kt': số ≥ 100 vùng vẽ (bỏ dòng khung tên lặp); cổng = có ≥ 10 số không kể cao độ."""
        t = ng.text(fid, pg)
        if kieu == "ma":
            x = set(neo_mavl(t, ho)) - set(neo_mavl(so.get((fid, pg), ""), ho))
            return x, (ho if len(x) >= NEO_MA_CONG else {re.match(r"[A-Z]+", k).group(0) for k in x})
        ls = [l for l in t.splitlines() if l.strip() and l.strip() not in lap_of(fid)]
        x = {v for v in neo_tach_so(" ¦ ".join(ls)) if float(v) >= NEO_KT_MIN}
        cong = sum(float(v) >= NEO_KT_MIN for v in neo_tach_so(" ¦ ".join(l for l in ls if not NEO_LEV.match(l))))
        return x, cong >= NEO_KT_CONG

    for x in items:
        fid, pg = x["file"], x["trang"]
        pdf = ng.la_pdf(fid) and str(pg).isdigit()
        x["link"] = f"{fid} p{pg}" if pdf else f"{fid} {x.get('o') or ''}".strip()
        if x["kieu"] in ("ma", "kt"):
            ma_ = x["kieu"] == "ma"
            ho = set(x["ho"])
            if len(x["nhom"]) > 1:
                x["link"] += f" +{len(x['nhom']) - 1} trang"
            if not x["ma"]:
                ket(x, NEO_KK, "giá trị không có mã thuộc họ mã dự án" if ma_ else "giá trị không có chuỗi R×S×C (≥ 2 số ≥ 100)")
                continue
            seen, cong, kk_, trs = set(), [], [], []
            for f_, pg_, _o in x["nhom"]:
                t = ng.text(f_, int(pg_))
                if t is None or len(t.strip()) < 20:
                    kk_.append("trang không có lớp text")
                    continue
                v, c_ = neo3(f_, int(pg_), x["kieu"], ho)
                seen |= v
                trs.append((f_, int(pg_)))
                cong.append(c_)
            key = (lambda m: next(iter(neo_mavl(m, ho)))) if ma_ else (lambda m: m)
            miss = list(dict.fromkeys(m for m in x["ma"] if key(m) not in seen))
            if not miss:
                ket(x, "CÓ NEO")
                continue
            if kk_:
                ket(x, NEO_KK, f"{kk_[0]} (có thể nằm ở tờ đó)")  # còn trang không có lớp text: không kết luận LỆCH
                continue
            if ma_:  # mã chỉ kiểm khi MỌI trang được link có mã cùng họ dạng text; thiếu thì mã đó có thể vẽ bằng nét
                hm = {m: re.match(r"[A-Z]+", key(m)).group(0) for m in miss}
                ko = [m for m in miss if any(hm[m] not in c_ for c_ in cong)]
                if ko == miss:
                    ket(x, NEO_KK, f"trang không có mã họ {'/'.join(sorted({hm[m] for m in ko}))} dạng text (mã vẽ bằng nét)")
                    continue
                miss = [m for m in miss if m not in ko]
            elif not any(cong):  # hợp các trang có text, có ≥ 1 trang qua cổng là kết luận được
                ket(x, NEO_KK, f"trang không có kích thước dạng text (< {NEO_KT_CONG} số ≥ {NEO_KT_MIN})")
                continue
            if ma_:
                gy = [f"{a} p{q}" for a in dict.fromkeys(a for a, _ in trs) for q in range(1, ng.so_trang(a) + 1)
                      if (a, q) not in trs and all(key(m) in neo3(a, q, "ma", ho)[0] for m in x["ma"])]
                ket(x, "LỆCH", f"không thấy mã {', '.join(miss[:4])} trên trang (trang có: {', '.join(sorted(seen)[:6])})", gy)
            else:
                ket(x, "LỆCH", f"không thấy số {', '.join(miss[:4])} trong kích thước dạng text của trang")
            continue
        if x["kieu"] == "to":
            if not pdf:
                continue  # link tới dwg/ảnh/Excel: không kiểm số tờ
            pg = int(pg)
            want, keys = neo_so_to(x["nhan"], duoi)
            if not want:
                ket(x, NEO_KK, "nhãn không có số tờ (khớp khung tên đã đọc)")
                continue
            t = ng.text(fid, pg)
            if so.get((fid, pg)) not in want:
                if t is None or len(t.strip()) < 20:
                    ket(x, NEO_KK, "trang không có lớp text")
                    continue
                words = {_an(w) for w in re.split(r"[\s/()]+", t)}  # nhãn gọn 'b/502.1' thường là callout in trên chính trang đó
                if not (keys & words) and not any(len(_an(v)) >= 6 and neo_rx(v).search(_fold(t)) for v in want):
                    ket(x, "LỆCH", f"trang là tờ '{so.get((fid, pg), '?')}', không có số tờ {'/'.join(sorted(want)[:3])}",
                        [f"{a} p{b} ({v})" for v in sorted(want) for a, b in tc.get(v, []) if (a, b) != (fid, pg)])
                    continue
            mt = re.search(r"#(\d{1,2})\s*([A-Za-z][^;#]*)$", x["nhan"].strip())  # '… #12 DETAIL - CABINET …': tên chi tiết là text
            words = {w for w in re.findall(r"[A-Z]{3,}", _fold(mt.group(2)).upper()) if w not in NEO_STOPW} if mt else set()
            if words and t:
                tw = set(re.findall(r"[A-Z]{3,}", _fold(t).upper()))
                if len(words & tw) < 0.8 * len(words):
                    ket(x, "LỆCH", f"đúng tờ nhưng tên chi tiết #{mt.group(1)} thiếu trên trang: {', '.join(sorted(words - tw)[:4])}")
                    continue
            ket(x, "CÓ NEO")
        elif x.get("nhom"):
            ss = [dich(*l) for l in x["nhom"]]
            v = "".join(a for a, _ in ss if a is not None)
            if not v:
                ket(x, NEO_KK, ss[0][1])
                continue
            miss = [m for m in x["ma"] if _an(m) not in v]
            kk_ = [w for a, w in ss if a is None]
            if miss and kk_:
                ket(x, NEO_KK, f"{kk_[0]} (mã có thể nằm ở link đó)")
            elif miss:
                ket(x, "LỆCH", "không thấy mã " + ", ".join(miss[:4]) + f" trên {len(ss)} trang/dòng được link")
            else:
                ket(x, "CÓ NEO")
        else:
            if not x["ma"]:
                ket(x, NEO_KK, "không có mã/model để kiểm")
                continue
            if pdf:
                pg = int(pg)
                t = ng.text(fid, pg)
                if t is None or len(t.strip()) < 20:
                    ket(x, NEO_KK, "trang không có lớp text")
                    continue
                miss = [m for m in x["ma"] if _an(m) not in sq_of(fid, pg)]
                gy = [f"{fid} p{q}" for q in range(1, ng.so_trang(fid) + 1) if q != pg and all(_an(m) in sq_of(fid, q) for m in x["ma"])] if miss else []
            else:
                m_ = re.match(_RX_O, str(x.get("o")))
                if not m_ or not ng.la_xlsx(fid):
                    ket(x, NEO_KK, "link không chỉ tới trang PDF / ô Excel")
                    continue
                rows = xl_rows(fid, m_.group(1))
                if not rows:
                    ket(x, NEO_KK, "không mở được sheet Excel")
                    continue
                r1 = int(m_.group(2))
                r2 = int(m_.group(3) or r1)
                v = "".join(rows.get(i, "") for i in range(r1, r2 + 1))
                miss = [m for m in x["ma"] if _an(m) not in v]
                gy = [f"'{m_.group(1)}'!A{i}" for i, b in rows.items() if not r1 <= i <= r2 and all(_an(m) in b for m in x["ma"])] if miss else []
            if miss:
                ket(x, "LỆCH", "không thấy mã " + ", ".join(miss[:4]), gy)
            else:
                ket(x, "CÓ NEO")
    return dict(ngay=f"{datetime.datetime.now():%Y-%m-%d %H:%M}", hash=neo_hash(items), tong=dict(tong),
                theo_kieu={k: dict(v) for k, v in tk.items()}, khong_kiem=dict(kk), lech=lech)


def neo_dong(l):
    return f"NEO LỆCH {l['ten']}: '{str(l['nhan'])[:50]}' → {l['link']}: {l['ly_do']}" + (f"; gợi ý {', '.join(l['goi_y'])}" if l["goi_y"] else "")


def neo_tong(res):
    t, tk = res["tong"], res.get("theo_kieu", {})
    ten = {"to": "số tờ", "spec": "spec·ML", "ma": "mã VL", "kt": "kích thước"}
    return (f"NEO: {t.get('CÓ NEO', 0)} CÓ NEO · {t.get('LỆCH', 0)} LỆCH · {t.get(NEO_KK, 0)} {NEO_KK}"
            + (" [" + " · ".join(f"{ten.get(k, k)} {v.get('CÓ NEO', 0)}/{v.get('LỆCH', 0)}/{v.get(NEO_KK, 0)}" for k, v in tk.items()) + "]" if tk else "")
            + (" (" + "; ".join(f"{k}: {v}" for k, v in sorted(res["khong_kiem"].items(), key=lambda z: -z[1])) + ")" if res["khong_kiem"] else ""))
