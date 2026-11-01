# -*- coding: utf-8 -*-
"""Kiểm vá độ tin cậy 31/10: gemini --ma, Kiểm chéo Gemini số, danh sách ĐÃ LOẠI. Không gọi agy thật, chỉ ghi vào thư mục tạm.
A. Dự án giả: kiemke lưu file bỏ + lý do; gemini --ma ở mọi trạng thái.
B. Bản sao cache Blanca, so với pd bản trước khi vá:
   Gemini khớp hết -> Độ tin cậy y như bản cũ; chưa chạy -> không còn CAO ở mã có số; LỖI của check không đổi.
Chạy: py -I -X utf8 gem_so_test.py"""
import sys, os, json, shutil, tempfile, subprocess, importlib.util, collections, io, contextlib
HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = os.path.join(HERE, "..", "..", "SKILL", "01-scan", "tool")
sys.path.insert(0, TOOL)
import pd
import pymupdf
spec = importlib.util.spec_from_file_location("pd_cu", os.path.join(TOOL, "_backup", "pd_2026-10-31_truoc-vaGemini.py"))
pd_cu = importlib.util.module_from_spec(spec); spec.loader.exec_module(pd_cu); pd_cu.FORM = pd.FORM
sai = []


def kt(ten, ok, them=""):
    print(("  OK  " if ok else "  SAI ") + ten + (f" — {them}" if them else ""))
    if not ok: sai.append(ten)


def im(f, *a):
    with contextlib.redirect_stdout(io.StringIO()) as o: f(*a)
    return o.getvalue()


# ---------- A. dự án giả
print("A. Dự án giả")
goc = tempfile.mkdtemp(); inp = os.path.join(goc, "00_input", "D1")
for rel in ("ban ve\\a.pdf", "ban ve\\frame-1.pdf", "Archive\\cu.pdf", "CAD\\mep.pdf"):
    f = os.path.join(inp, rel); os.makedirs(os.path.dirname(f), exist_ok=True)
    d = pymupdf.open(); d.new_page().insert_text((50, 50), rel); d.save(f); d.close()
os.makedirs(os.path.join(goc, "01_scan"))
open(os.path.join(goc, "01_scan", "_bo_qua.txt"), "w", encoding="utf-8").write("# MEP không thuộc phạm vi\nCAD\\\\mep\n")
im(pd.cmd_kiemke, goc, "D1")
P = pd.paths(goc, "D1")
bo = {x["file"]: x["ly_do"] for x in pd.read_aux(P, "_kiemke_bo.json")["bo"]}
kt("3 file bỏ, đủ lý do", len(bo) == 3 and bo["CAD\\mep.pdf"] == "_bo_qua.txt: MEP không thuộc phạm vi" and "frame" in bo["ban ve\\frame-1.pdf"] and "archive" in bo["Archive\\cu.pdf"], str(bo))
FILES = json.load(open(os.path.join(P["data"], "FILES.kiemke.json"), encoding="utf-8"))["rows"]
kt("chỉ còn a.pdf trong FILES", [r["Tên file"] for r in FILES] == ["a.pdf"])

ms = [f"M{i}" for i in range(1, 21)]
QS = [{"Mã": m, "Trường": "Phạm vi nhà máy", "Giá trị": "KHÔNG THUỘC SCOPE NHÀ MÁY – thiết bị vệ sinh", "Trạng thái": "CÓ", "Nguồn": "BOQ"} for m in ms]
json.dump({"du_phong": ["F001"]}, open(os.path.join(P["data"], "_timkiem.kiemke.json"), "w"))
W = pd.ds_da_loai(P, {"FILES": FILES, "QUAN_SAT": QS}, set(ms))
txt = "\n".join(W)
kt("XEM có đủ file bỏ", all(os.path.basename(f) in txt for f in bo), f"{len(W)} dòng")
kt("XEM có dự phòng + tên file", "F001 a.pdf" in txt)
kt("XEM có đủ 20 mã ngoài scope (trước chỉ 12)", all(m in txt for m in ms) and "thiết bị vệ sinh" in txt)


class Kq:
    def __init__(s, out): s.stdout, s.stderr = out, ""


tra_loi = {}; goi = []
pd.agy_exe = lambda: "agy"
def gia_run(cmd, **kw):
    ten = [k for k in tra_loi if k in cmd[4]][0]; goi.append(ten)
    return Kq(tra_loi[ten])
run_that = subprocess.run; subprocess.run = gia_run


def crop(ten):
    f = os.path.join(goc, ten); open(f, "wb").write(b"x"); return f


def qs_ma(ma, gt, tt="CÓ", ng="HÌNH"):
    return {"Mã": ma, "Trường": "Kích thước BV", "Giá trị": gt, "Trạng thái": tt, "Tờ · Chi tiết": "p1", "Nguồn": ng}


def ghi_qs(rows): json.dump({"rows": rows}, open(os.path.join(P["data"], "QUAN_SAT.banve-1.json"), "w", encoding="utf-8"), ensure_ascii=False)


def gem(ma, *png):
    pd.cmd_gemini(list(png) + ["--ma", ma, "--goc", goc, "--dot", "D1"])


def tt(ma, rows): return pd.gemini_so(P, ma, rows)


# chưa ghi QUAN_SAT -> phải dừng, không gọi Gemini
goi.clear(); tra_loi["c0.png"] = '["900"]'
try: im(gem, "V0", crop("c0.png")); kt("chưa có QUAN_SAT thì dừng", False)
except SystemExit as e: kt("chưa có QUAN_SAT thì dừng, không gọi Gemini", not goi, str(e)[:60])

rows = [qs_ma("V1", "900 (R) × 450 × 2,375 (C); cánh 870"), qs_ma("V2", "600×300×1200"), qs_ma("V3", "500×500"), qs_ma("V4", "1000×2000"), qs_ma("V5", "KHÔNG THẤY", "KHÔNG THẤY")]
ghi_qs(rows)
by = collections.defaultdict(list)
for r in rows: by[r["Mã"]].append(r)
kt("V1 chưa chạy", tt("V1", by["V1"])["tt"] == "CHƯA CHẠY")
kt("V5 không có số", tt("V5", by["V5"])["tt"] == "KHÔNG CÓ SỐ")
kt("không lấy chuỗi diễn giải sau (870)", pd.claude_kt(by["V1"]) == ["450", "900", "2375"], str(pd.claude_kt(by["V1"])))
tra_loi.update({"c1.png": '```json\n["900", "450", "?"]\n```', "c1b.png": '["2,375", "50"]', "c2.png": '["600", "300", "120"]', "c3.png": "xin lỗi không đọc được", "c4.png": '["1000", "2000", "?"]'})
im(gem, "V1", crop("c1.png"))
kt("V1 thiếu 2375 khi mới 1 crop", tt("V1", by["V1"])["tt"] == "LỆCH")
im(gem, "V1", crop("c1b.png"))
kq = tt("V1", by["V1"]); kt("V1 khớp sau crop 2 (cộng dồn, chuẩn hóa 2,375)", kq["tt"] == "KHỚP" and kq["crops"] == ["c1.png", "c1b.png"], str(kq))
im(gem, "V2", crop("c2.png"))
kq = tt("V2", by["V2"]); kt("V2 lệch đúng số", kq["tt"] == "LỆCH" and kq["thieu"] == ["1200"], str(kq))
d = pd.gemini_so_dong("V2", kq); kt("dòng V2 = MÂU THUẪN, nguồn GEMINI", d["Trạng thái"] == "MÂU THUẪN" and d["Nguồn"] == "GEMINI" and "chỉ Claude: 1200" in d["Giá trị"], d["Giá trị"])
goi.clear(); im(gem, "V3", crop("c3.png"))
kt("V3 Gemini trả rác = LỖI, không phải LỆCH", tt("V3", by["V3"])["tt"] == "LỖI")
kt("rác không được cache", not os.path.exists(os.path.join(goc, "c3.png.gemini.txt")))
im(gem, "V4", crop("c4.png"))
kt("V4 khớp, '?' bị bỏ", tt("V4", by["V4"])["tt"] == "KHỚP")
# đổi số sau khi chạy: so theo QUAN_SAT lúc chạy, không phải lúc gọi gemini
kt("đổi QUAN_SAT sau khi chạy -> lệch ngay", tt("V4", [qs_ma("V4", "1000×2100")])["tt"] == "LỆCH")
# cache: cùng crop cho mã khác không gọi lại; crop bị ghi đè thì gọi lại
goi.clear(); ghi_qs(rows + [qs_ma("V6", "1000×2000")]); im(gem, "V6", os.path.join(goc, "c4.png"))
kt("crop dùng chung không gọi Gemini lại", not goi and tt("V6", [qs_ma("V6", "1000×2000")])["tt"] == "KHỚP")
f = os.path.join(goc, "c4.png"); os.utime(f, (os.path.getmtime(f) + 60, os.path.getmtime(f) + 60))
im(gem, "V6", f); kt("crop ghi đè (mới hơn cache) thì đọc lại", goi == ["c4.png"])
shutil.rmtree(goc, ignore_errors=True)
subprocess.run = run_that

# ---------- B. Blanca
print("B. Bản sao Blanca: bản cũ vs bản vá")
tmp = tempfile.mkdtemp(); d = os.path.join(tmp, "01_scan", "_data", "X")
shutil.copytree(os.path.join(HERE, "_cache", "Blanca", "_data"), d)
P = pd.paths(tmp, "X")


def doi(mod, ten):
    ra = os.path.join(tmp, ten); im(mod.cmd_doichieu, tmp, "X", ["--ra", ra])
    return json.load(open(os.path.join(ra, "_doichieu_tool.json"), encoding="utf-8"))["muc"], ra


cu, _ = doi(pd_cu, "cu")
_, _, M, _ = pd.load_merged(P, ("kiemke", "boq", "spec", "banve", "user"))
qs = collections.defaultdict(list)
for r in M["QUAN_SAT"]: qs[str(r["Mã"])].append(r)
co_so = {m for m, rs in qs.items() if pd.claude_kt(rs)}
moi, ra = doi(pd, "chua")
kt("chưa chạy Gemini số: mã có số không còn CAO", not [m for m in co_so if moi.get(m) == "CAO"], f"{len(co_so)} mã có số; cũ {collections.Counter(cu.values())} -> mới {collections.Counter(moi.values())}")
kt("mã không có số: mức không đổi", all(moi[m] == cu[m] for m in moi if m not in co_so))
tr = json.load(open(os.path.join(ra, "QUAN_SAT.doichieu-tool.json"), encoding="utf-8"))["rows"]
kt("sinh đúng 1 dòng cho mỗi mã có số trong scope", len(tr) == len(co_so - pd.ngoai_pham_vi(M["QUAN_SAT"])), str(len(tr)))

os.makedirs(os.path.join(d, "_gemini_so"))
for m in co_so:
    json.dump({"ma": m, "crops": {"gia.png": {"so": pd.claude_kt(qs[m]), "file": "gia.png"}}}, open(pd.gem_file(P, m), "w", encoding="utf-8"))
moi2, _ = doi(pd, "khop")
kt("Gemini khớp hết: Độ tin cậy y như bản cũ", moi2 == cu, f"khác {sum(moi2[m] != cu[m] for m in cu)} mã")
m0 = sorted(co_so - pd.ngoai_pham_vi(M["QUAN_SAT"]))[0]
g = json.load(open(pd.gem_file(P, m0), encoding="utf-8")); g["crops"]["gia.png"]["so"] = []
json.dump(g, open(pd.gem_file(P, m0), "w", encoding="utf-8"))
moi3, ra3 = doi(pd, "lech")
kt(f"{m0} lệch số -> THAP", moi3[m0] == "THAP", f"cũ {cu[m0]}")
kt(f"{m0} dòng tool = MÂU THUẪN", any(x["Mã"] == m0 and x["Trạng thái"] == "MÂU THUẪN" for x in json.load(open(os.path.join(ra3, "QUAN_SAT.doichieu-tool.json"), encoding="utf-8"))["rows"]))
# issue nháp chỉ cho mã chưa có issue nào: thử cả nhóm mã chưa có issue trong dữ liệu
co_is = {str(x.get("Mã")) for x in M.get("ISSUES", [])}
m1 = next((m for m in sorted(co_so - pd.ngoai_pham_vi(M["QUAN_SAT"])) if not any(m in c for c in co_is)), None)
if m1:
    g = json.load(open(pd.gem_file(P, m1), encoding="utf-8")); g["crops"]["gia.png"]["so"] = []
    json.dump(g, open(pd.gem_file(P, m1), "w", encoding="utf-8"))
    _, ra4 = doi(pd, "lech2")
    kt(f"{m1} (chưa có issue) có issue DT", any(x["Mã"] == m1 and "Kích thước BV" in x["Nội dung"] for x in json.load(open(os.path.join(ra4, "ISSUES.doichieu-tool.json"), encoding="utf-8"))["rows"]))
else: print("  --  không có mã chưa có issue để thử issue DT")


def chk(mod):
    wb, sheets, cols = mod.load_form(); raw, errs = mod.load_json(P["data"]); merged, e2, _ = mod.merge(sheets, cols, raw)
    return mod.run_checks(P, sheets, cols, merged, [], errs + e2, None)


rc, rm = chk(pd_cu), chk(pd)
kt("check: LỖI giữ nguyên", rc["E"] == rm["E"], f"{len(rc['E'])} LỖI")
them = [x for x in rm["W"] if x not in rc["W"]]; mat = [x for x in rc["W"] if x not in rm["W"]]
print(f"  XEM: cũ {len(rc['W'])} -> mới {len(rm['W'])}; thêm {len(them)}, mất {len(mat)}")
for x in them[:8]: print("    +", x[:170])
for x in mat[:8]: print("    -", x[:170])
shutil.copy(os.path.join(ra3, "QUAN_SAT.doichieu-tool.json"), d)
rm2 = chk(pd)
kt("ghép dòng tool vào QUAN_SAT không thêm LỖI", rm2["E"] == rm["E"], f"thêm {[x[:100] for x in rm2['E'] if x not in rm['E']][:3]}")
shutil.rmtree(os.path.join(d, "_gemini_so"))
rm3 = chk(pd)
kt("check: có dòng XEM mã chưa so số", any(x.startswith("45 mã có kích thước nhưng chưa so số") or "mã có kích thước nhưng chưa so số" in x for x in rm3["W"]), next((x[:90] for x in rm3["W"] if "chưa so số" in x), ""))
json.dump({"file_tong": ["F023"], "trang_trung": {"F001 p1": "F002 p1", "F001 p2": "F002 p2"}}, open(os.path.join(d, "_timkiem.json"), "w"))
t = "\n".join(pd.ds_da_loai(P, M, set()))
kt("XEM có file tổng + trang trùng", "ĐÃ LOẠI 1 file tổng" in t and "F023" in t and "F001 2 trang" in t)
shutil.rmtree(tmp, ignore_errors=True)
print("\nKẾT QUẢ:", "ĐẠT" if not sai else f"SAI {len(sai)}: {sai}")
sys.exit(1 if sai else 0)
