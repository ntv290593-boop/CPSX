"""Xác định bề mặt mà 1 tag vật liệu chỉ vào: tường, trần, sàn, cửa, joinery.

Lớp 1 (máy, chắc): chữ nhãn đi kèm tag, block trần, tên khung hình chứa tag.
Lớp 2 (nhìn hình): tag trên mặt đứng còn lại là 'can_soat', vì có thể chỉ vào cửa hoặc đồ gỗ.
Kết quả soát nằm ở _data\\_soat_tag.*.json, khóa <số tờ>#<số thứ tự tag trên tờ>, luôn thắng lớp 1.
"""
import glob
import json
import os
import re

BLOCK_TRAN = {"CEILING", "ceiling", "CODE-CLNG.1", "CODE-CLNG.2"}
RX_NHAN = [(re.compile(r"\bPORTAL\b", re.I), "portal"), (re.compile(r"CEILING|\bCLG\b", re.I), "tran"), (re.compile(r"\bDOOR\b", re.I), "cua"), (re.compile(r"FLOOR|SKIRTING|STEP\b|DECK", re.I), "san"),
           (re.compile(r"COUNTER|SHELF|SHELV|CABINET|VANITY|LOCKER|BENCH", re.I), "joinery")]
RX_KHUNG = [(re.compile(r"REFLECTED|\bRCP\b|CEILING", re.I), "tran"), (re.compile(r"FLOOR", re.I), "san"),
            (re.compile(r"FURNITURE|PARTITION|SETTING OUT|ELECTRICAL|\bEMP\b", re.I), "mb_khac"), (re.compile(r"ELEV|SECTION", re.I), "mat_dung")]


def nhan_tag(t):
    if t.get("ml"):
        return " ".join(x for x in t["ml"] if len(x) > 3)
    return " ".join(v for k, v in t["a"].items() if len(v) > 3 and k not in ("note1",))


def khung_cua_tag(t, l):
    """Tên khung hình chứa tag: viewport chứa đầu mũi tên (hoặc tag), lấy tiêu đề gần nhất phía dưới khung. Tờ 1 khung thì lấy tiêu đề duy nhất."""
    td = []
    for x in l["tag"]:
        a = x.get("a") or {}
        ten = a.get("%%AREA") or a.get("%%UROOM")
        if ten and x.get("xy"):
            td.append((ten, x["xy"]))
    if not td:
        return ""
    if len(td) == 1:
        return td[0][0]
    p = (t.get("mui") or [t.get("xy")])[0]
    if not p:
        return ""
    vps = l.get("vp") or []
    vp = next((v for v in vps if v[0] <= p[0] <= v[2] and v[1] <= p[1] <= v[3]), None)
    if vp is None and vps:
        vp = min(vps, key=lambda v: abs((v[0] + v[2]) / 2 - p[0]) + abs((v[1] + v[3]) / 2 - p[1]))
    if vp is not None:
        cx = (vp[0] + vp[2]) / 2
        return min(td, key=lambda q: abs(q[1][1] - vp[1]) + 0.3 * abs(q[1][0] - cx))[0]
    return min(td, key=lambda q: abs(q[1][0] - p[0]) + abs(q[1][1] - p[1]))[0]


def doan(t, l):
    """Trả (bề mặt, lý do). Bề mặt: tuong, tran, san, cua, portal (lỗ mở, không có cánh nên không có dòng cửa), joinery, can_soat, mb_khac. Kết quả soát có thêm khac, khong_ro."""
    if t["b"] in BLOCK_TRAN:
        return "tran", "block trần"
    nh = nhan_tag(t)
    for rx, kq in RX_NHAN:
        if rx.search(nh):
            return kq, f"nhãn: {nh[:40]}"
    k = khung_cua_tag(t, l)
    for rx, kq in RX_KHUNG:
        if rx.search(k):
            return ("can_soat" if kq == "mat_dung" else kq), f"khung: {k[:50]}"
    return "can_soat", f"khung không rõ: {k[:40]}"


def nap_soat(d):
    out = {}
    for f in sorted(glob.glob(os.path.join(d, "_soat_tag.*.json"))):
        for r in json.load(open(f, encoding="utf-8"))["rows"]:
            out[r["khoa"]] = r
    return out
