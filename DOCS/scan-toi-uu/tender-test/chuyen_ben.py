"""Chuyển `Nhóm thực hiện` + `NTP chi tiết` cũ sang `Bên thực hiện` (plan DOCS\\PLAN_TENDER-CODEID-DMVT_2026-10-31.md, mục 6).

Chạy:  py -I chuyen_ben.py xem <GOC>   chỉ báo sẽ đổi gì, không ghi
       py -I chuyen_ben.py ghi <GOC>   sao lưu rồi ghi; kiểm lại sau khi ghi, sai thì tự trả về bản sao lưu

Quy tắc: NTP → giá trị `NTP chi tiết` (không có thì giữ "NTP"); Nhà máy, Owner, MEP, Gói khác giữ nguyên; Chung bỏ.
Lấy theo kết quả đã ghép (có cả dòng mẫu `_loc` và thứ tự ưu tiên), ghi mỗi khóa 1 dòng vào SCOPE.ben.json.
Không để trong SCOPE.doichieu-*.json: SKILL.md mục 4 cho xóa JSON của bước khi chạy lại, bên thực hiện không được mất theo.
MAT chỉ xóa cột (plan mục 6). Sao lưu ở _data\\_backup\\truoc-gd3-<ngày giờ>.
"""
import datetime
import glob
import json
import os
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "..", ".."))
sys.path.insert(0, os.path.join(ROOT, "SKILL", "10-tender", "tool"))
import td  # noqa: E402

COT_CU = ("Nhóm thực hiện", "NTP chi tiết")
GIU = ("Nhà máy", "Owner", "MEP", "Gói khác")
HAI_COT = ("Bên thực hiện",) + COT_CU


def ben_theo_quy_tac(nhom, ntp, ghi_ro=False):
    """-> (bên hoặc None, ghi chú nếu dòng lạ).

    ghi_ro: chính dòng này ghi `NTP chi tiết` (vd vanity Nhà máy kèm gương NTP), không phải thừa hưởng từ dòng mẫu `_loc`
    rồi bị ghi đè Nhóm. Khi đó giữ cả 2, cùng dạng gộp nhiều bên của plan mục 6a; thừa hưởng thì bỏ phần chi tiết.
    """
    nhom, ntp = td.s(nhom), td.s(ntp)
    if nhom == "NTP":
        return ntp or "NTP", "NTP không có NTP chi tiết" if not ntp else None
    if nhom in GIU:
        if ntp and ghi_ro:
            return f"{nhom}; {ntp}", f"{nhom} và dòng ghi rõ NTP chi tiết '{ntp}', giữ cả 2"
        return nhom, f"{nhom}, NTP chi tiết '{ntp}' thừa hưởng từ dòng mẫu, bỏ" if ntp else None
    if nhom == "Chung" or not nhom:
        if ntp and ghi_ro:
            return ntp, f"Nhóm '{nhom}' và dòng ghi rõ NTP chi tiết '{ntp}', giữ phần chi tiết"
        return None, f"Nhóm '{nhom}', NTP chi tiết '{ntp}' thừa hưởng từ dòng mẫu, bỏ" if ntp else None
    return nhom, f"Nhóm lạ '{nhom}', giữ nguyên"


def tap_du_lieu(p):
    """Các file LLM/user có cột cũ, trừ -tool (tool không ghi 2 cột này)."""
    out = []
    for f in sorted(glob.glob(p.f("SCOPE.*.json")) + glob.glob(p.f("MAT.*.json"))):
        if "-tool" in os.path.basename(f) or os.path.basename(f) == "SCOPE.ben.json":
            continue
        rows = td.jload(f)["rows"]
        if any(c in r for r in rows for c in COT_CU):
            out.append((f, rows))
    return out


def bo_cot(rows, khoa):
    """Bỏ 2 cột cũ; dòng không còn giá trị nào thì bỏ luôn. -> (rows mới, số ô bỏ)."""
    moi, n = [], 0
    for r in rows:
        co = [c for c in COT_CU if c in r]
        if not co:
            moi.append(r)
            continue
        n += len(co)
        x = {c: v for c, v in r.items() if c not in COT_CU}
        if set(x) - {khoa, "_loc"}:
            moi.append(x)
    return moi, n


def kiem_lai(p, truoc, can):
    """Mỗi khóa SCOPE có đúng bên như quy tắc; mọi cột khác của SCOPE và MAT không đổi."""
    sau = td.merge(p)[0]
    loi = []
    if list(sau["SCOPE"]) != list(truoc["SCOPE"]):
        loi.append("tập khóa SCOPE đổi")
    for k, r in truoc["SCOPE"].items():
        x = sau["SCOPE"].get(k, {})
        if x.get("Bên thực hiện") != can.get(k):
            loi.append(f"SCOPE {k}: Bên '{x.get('Bên thực hiện')}', cần '{can.get(k)}'")
        for c in set(r) | set(x):
            if c not in HAI_COT and r.get(c) != x.get(c):
                loi.append(f"SCOPE {k}: cột {c} bị đổi")
    for k, r in truoc["MAT"].items():
        x = sau["MAT"].get(k, {})
        for c in set(r) | set(x):
            if c not in HAI_COT and r.get(c) != x.get(c):
                loi.append(f"MAT {k}: cột {c} bị đổi")
    con = [f for f, _ in tap_du_lieu(p)]
    if con:
        loi.append(f"còn cột cũ trong {con}")
    return loi


def main(a):
    if len(a) != 2 or a[0] not in ("xem", "ghi"):
        sys.exit(__doc__)
    p = td.P(a[1]).need()
    if os.path.exists(p.f("SCOPE.ben.json")):
        sys.exit(f"LỖI: đã có {p.f('SCOPE.ben.json')}, không chuyển lần 2.")
    du_lieu = tap_du_lieu(p)
    if not du_lieu:
        sys.exit("Không có file nào còn cột cũ, không có gì để chuyển.")
    truoc = td.merge(p)[0]
    can, la = {}, []
    ghi_ro = {td.s(r.get("Khóa")) for f, rows in du_lieu for r in rows if "_loc" not in r and COT_CU[1] in r}
    for k, r in truoc["SCOPE"].items():
        ben, ghi_chu = ben_theo_quy_tac(r.get(COT_CU[0]), r.get(COT_CU[1]), k in ghi_ro)
        if ben:
            can[k] = ben
        if ghi_chu:
            la.append(f"{k}: {ghi_chu}")
    tong_o = 0
    for f, rows in du_lieu:
        khoa = "Mã VL" if os.path.basename(f).startswith("MAT.") else "Khóa"
        moi, n = bo_cot(rows, khoa)
        tong_o += n
        print(f"{os.path.basename(f)}: {sum(1 for r in rows if any(c in r for c in COT_CU))} dòng có cột cũ, bỏ {n} ô; {len(rows)} → {len(moi)} dòng")
    tk = {}
    for v in can.values():
        tk[v] = tk.get(v, 0) + 1
    print(f"SCOPE {len(truoc['SCOPE'])} khóa: {len(can)} có Bên thực hiện, {len(truoc['SCOPE']) - len(can)} không (Chung hoặc chưa gán)")
    print("Theo bên:", dict(sorted(tk.items(), key=lambda x: -x[1])))
    print("MAT sẽ mất:", [(k, r[COT_CU[0]]) for k, r in truoc["MAT"].items() if r.get(COT_CU[0])])
    for x in la:
        print("XEM", x)
    if a[0] == "xem":
        print("Chưa ghi gì (lệnh xem).")
        return
    bk = p.f(os.path.join("_backup", f"truoc-gd3-{datetime.datetime.now():%Y%m%d-%H%M%S}"))
    os.makedirs(bk)
    for f, _ in du_lieu:
        shutil.copy2(f, bk)
    try:
        for f, rows in du_lieu:
            khoa = "Mã VL" if os.path.basename(f).startswith("MAT.") else "Khóa"
            td.jdump(f, {"rows": bo_cot(rows, khoa)[0]})
        td.jdump(p.f("SCOPE.ben.json"), {"rows": [{"Khóa": k, "Bên thực hiện": v} for k, v in can.items()]})
        loi = kiem_lai(p, truoc, can)
    except Exception as e:
        loi = [repr(e)]
    if loi:
        for f, _ in du_lieu:
            shutil.copy2(os.path.join(bk, os.path.basename(f)), f)
        if os.path.exists(p.f("SCOPE.ben.json")):
            os.remove(p.f("SCOPE.ben.json"))
        sys.exit("LỖI, đã trả lại dữ liệu cũ:\n  " + "\n  ".join(loi[:30]))
    print(f"Đã ghi {p.f('SCOPE.ben.json')} ({len(can)} dòng), sao lưu ở {bk}. Kiểm lại: đủ {len(truoc['SCOPE'])} khóa, các cột khác không đổi.")


if __name__ == "__main__":
    main(sys.argv[1:])
