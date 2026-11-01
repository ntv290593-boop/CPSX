# -*- coding: utf-8 -*-
r"""do_area.py — đo cột Area của DMVT (plan Code ID/DMVT mục 4, GĐ4): bảng tiền tố dựng từ KHU so với form DMVT user lập.

Chạy: py -I do_area.py [GOC]     (mặc định _ACTIVE\SOFITEL; cần đã chạy `td.py area` hoặc `build`)
Đáp án: `01. INPUT\new\New folder\20261031-DMVT SOFITEL BASE (1).xlsx` sheet DMVT, B8:C51 (chỉ dùng để chấm, bảng không học từ đây).
Quy ước so: `STANDARD KING 02 / 02A` trong đáp án tính là SK-02, vì KHU của bộ FF&E chỉ có SK-02; mọi hành lang gộp thành CORRIDOR, mọi sảnh thang thành LIFT LOBBY.
Chỉ báo số, không có mốc (plan mục 8a).
"""
import os, re, sys
sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", ".."))
sys.path.insert(0, os.path.join(ROOT, "SKILL", "10-tender", "tool"))
import td
import openpyxl

TEN = {"STANDARD KING": "SK-01", "STANDARD KING 02 / 02A": "SK-02", "STANDARD TWIN 01": "ST-01", "CONNECTING 01A": "CR-01A", "CONNECTING 01": "CR-01",
       "GARDEN KING 01": "GK-01", "GARDEN TWIN 01": "GT-01", "JUNIOR SUITE": "JS-01", "PRESTIGE SUITE 01": "PS-01", "PRESTIGE SUITE 02": "PS-02",
       "PRESTIGE SUITE 02A": "PS-02A", "CORRIDOR L04 – L20": "CORRIDOR", "GUESTROOM LIFT LOBBY": "LIFT LOBBY"}


def chuan(x):
    x = x.strip()
    return "CORRIDOR" if x.startswith("Hành lang") else "LIFT LOBBY" if x.startswith("Sảnh thang") else x


def main():
    goc = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "_ACTIVE", "SOFITEL")
    p = td.P(goc).need()
    area = td.merge(p)[0]["AREA"]
    if not area:
        sys.exit("LỖI: chưa có bảng AREA, chạy td.py area trước")
    bang = {k: {chuan(x) for x in td.tach_cham_phay(r.get("Mã phòng"))} for k, r in area.items()}
    f = os.path.join(p.inp, "New folder", "20261031-DMVT SOFITEL BASE (1).xlsx")
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True)["DMVT"]
    n = dung = 0
    sai = []
    for i, r in enumerate(ws.iter_rows(min_row=8, max_row=51, min_col=2, max_col=3, values_only=True), 8):
        if not r[1]:
            continue
        n += 1
        dap_an = {TEN[x.strip()] for x in str(r[1]).split("\n") if x.strip()}
        ma = {x[0] for x in td.tach_ma(str(r[0])) if x[3] == "ffe"}
        may = set().union(*(bang.get(m[:2], set()) for m in ma)) if ma else set()
        if may == dap_an:
            dung += 1
        else:
            sai.append((i, sorted(ma), "thừa " + ",".join(sorted(may - dap_an)), "thiếu " + ",".join(sorted(dap_an - may))))
    print(f"Khớp đúng tập phòng: {dung}/{n} dòng có tên phòng (form hàng 8–51)")
    for x in sai:
        print("  hàng", *x)


if __name__ == "__main__":
    main()
