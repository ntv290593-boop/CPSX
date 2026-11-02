"""Kiểm thử skill 10-tender (plan DOCS\\PLAN_TENDER-CODEID-DMVT_2026-10-31.md, mục 8a).

Chạy:  py -I kiemthu.py            cả bộ (--sofitel tự bỏ qua nếu máy không có _ACTIVE)
       py -I kiemthu.py --mau      chỉ logic thuần, dữ liệu mẫu tự dựng
       py -I kiemthu.py --sofitel  chỉ phần cần dữ liệu thật

--mau không dùng số liệu dự án. --sofitel chỉ đọc hồ sơ thật, mọi thứ ghi ra thư mục tạm;
script kiểm chữ ký _data thật trước/sau để chắc không ghi vào 10_TENDER thật.
Mỗi GĐ thêm phần kiểm của mình vào đây. GĐ chỉ tính xong khi cả bộ vẫn đạt.
"""
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
import zipfile

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", ".."))
TD = os.path.join(ROOT, "SKILL", "10-tender", "tool", "td.py")
GOC_SOFITEL = os.path.join(ROOT, "_ACTIVE", "SOFITEL")

BO_KIEM = []


def kiem(gd, loai):
    def dk(fn):
        BO_KIEM.append((gd, loai, fn))
        return fn
    return dk


def td_that(*args):
    """Chạy td.py đúng tham số, trả (mã thoát, stdout + stderr)."""
    r = subprocess.run([sys.executable, "-I", TD, *args], capture_output=True, encoding="utf-8", errors="replace")
    return r.returncode, r.stdout + r.stderr


def td(*args):
    """Như td_that, nhưng build mặc định bỏ DMVT: các phép kiểm BOQ không cần mở Excel. Kiểm DMVT trong build thì dùng td_that."""
    if args[0] == "build" and "--khong-dmvt" not in args:
        args = (*args, "--khong-dmvt")
    return td_that(*args)


def td_ok(*args):
    rc, out = td(*args)
    assert rc == 0, f"td.py {args[0]} thoát {rc}:\n{out}"
    return out


def doc_index(goc):
    with open(os.path.join(goc, "10_TENDER", "_data", "FILE_INDEX.kiemke-tool.json"), encoding="utf-8") as fh:
        return json.load(fh)["rows"]


def doc_json(goc, ten):
    with open(os.path.join(goc, "10_TENDER", "_data", ten), encoding="utf-8") as fh:
        return json.load(fh)


def duong_dan(rows):
    return {r["Đường dẫn"] for r in rows}


def tao_file(goc_input, rel, noi_dung=b"x"):
    path = os.path.join(goc_input, *rel.split("/"))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(noi_dung)
    return path


def dung_mau(tmp):
    """Dự án giả: 1 thư mục cần loại ở gốc, 1 thư mục cùng tên nằm sâu hơn, 1 thư mục trùng tiền tố."""
    goc = os.path.join(tmp, "DA")
    inp = os.path.join(goc, "hoso")
    for rel in ("boq.xlsx", "tcq.xlsx", "HSMT/spec.pdf", "HSMT/New folder/giu.pdf",
                "New folder/khac_du_an_PVCV.xlsx", "New folder/con/sau.pdf", "New folder 2/giu.pdf"):
        tao_file(inp, rel)
    with zipfile.ZipFile(os.path.join(inp, "New folder", "bo.zip"), "w") as z:
        z.writestr("trong_zip.pdf", b"y")
    td_ok("init", goc, "--input", "hoso", "--boq", "boq.xlsx", "--tcq", "tcq.xlsx", "--du-an", "T")
    return goc


# ---------------------------------------------------------------- GĐ0
BI_LOAI = {"New folder/khac_du_an_PVCV.xlsx", "New folder/con/sau.pdf", "New folder/bo.zip"}
PHAI_GIU = {"boq.xlsx", "tcq.xlsx", "HSMT/spec.pdf", "HSMT/New folder/giu.pdf", "New folder 2/giu.pdf"}


@kiem(0, "mau")
def gd0_chua_cau_hinh_thi_van_quet_het(tmp):
    """Mốc đối chứng: thiếu dòng này thì các phép kiểm 'không có file' dưới đây vô nghĩa."""
    goc = dung_mau(tmp)
    td_ok("kiemke", goc)
    assert BI_LOAI | PHAI_GIU == duong_dan(doc_index(goc))


@kiem(0, "mau")
def gd0_loai_dung_thu_muc_ke_ca_trong_zip(tmp):
    goc = dung_mau(tmp)
    td_ok("init", goc, "--bo-thu-muc", "New folder")
    td_ok("kiemke", goc)
    rows = doc_index(goc)
    assert duong_dan(rows) == PHAI_GIU, duong_dan(rows)
    assert not any(r.get("Trong zip") for r in rows), "file trong zip của thư mục đã loại vẫn vào FILE_INDEX"
    assert set(doc_json(goc, "_chuky.json")) == {r["ID"] for r in rows}, "_chuky có file đã loại"


@kiem(0, "mau")
def gd0_khong_phan_biet_hoa_thuong_va_gach_cheo(tmp):
    goc = dung_mau(tmp)
    td_ok("init", goc, "--bo-thu-muc", "new FOLDER\\;HSMT\\New folder")
    td_ok("kiemke", goc)
    assert duong_dan(doc_index(goc)) == PHAI_GIU - {"HSMT/New folder/giu.pdf"}


@kiem(0, "mau")
def gd0_gia_tri_nguy_hiem_bi_tu_choi(tmp):
    """'.' hay '' mà lọt qua sẽ loại cả bộ hồ sơ; '..' thoát khỏi thư mục input."""
    goc = dung_mau(tmp)
    td_ok("kiemke", goc)
    truoc = doc_index(goc)
    duong_cauhinh = os.path.join(goc, "10_TENDER", "_data", "_cauhinh.json")
    goc_cfg = doc_json(goc, "_cauhinh.json")
    for xau in ("", ".", "..", "a/../b", "/", "C:/abc"):
        cfg = dict(goc_cfg, bo_thu_muc=[xau])
        with open(duong_cauhinh, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, ensure_ascii=False)
        rc, out = td("kiemke", goc)
        assert rc != 0 and "bo_thu_muc" in out, f"kiemke nhận {xau!r}: rc={rc}\n{out}"
        assert doc_index(goc) == truoc, f"{xau!r}: kiemke vẫn ghi FILE_INDEX"
    with open(duong_cauhinh, "w", encoding="utf-8") as fh:
        json.dump(goc_cfg, fh, ensure_ascii=False)
    rc, out = td("init", goc, "--bo-thu-muc", "..")
    assert rc != 0 and "bo_thu_muc" in out, f"init nhận '..': rc={rc}\n{out}"
    assert doc_json(goc, "_cauhinh.json") == goc_cfg, "init lỗi nhưng vẫn ghi cấu hình"


@kiem(0, "mau")
def gd0_go_sai_ten_phai_duoc_bao(tmp):
    """Gõ sai mà im lặng thì file dự án khác lọt vào FILE_INDEX mà không ai biết."""
    goc = dung_mau(tmp)
    td_ok("init", goc, "--bo-thu-muc", "New foldr")
    out = td_ok("kiemke", goc)
    assert "XEM" in out and "New foldr" in out, out


@kiem(0, "mau")
def gd0_init_lai_khong_lam_mat_cau_hinh(tmp):
    """SKILL.md mục 4: có TCQ mới thì chạy lại init --tcq; bo_thu_muc phải còn."""
    goc = dung_mau(tmp)
    td_ok("init", goc, "--bo-thu-muc", "New folder")
    td_ok("init", goc, "--tcq", "tcq.xlsx")
    assert doc_json(goc, "_cauhinh.json")["bo_thu_muc"] == ["New folder"]


@kiem(0, "mau")
def gd0_loai_sau_khi_da_quet_giu_nguyen_id(tmp):
    """Thêm bo_thu_muc khi thư mục đã có trong FILE_INDEX: ID file còn lại không đổi, file đã loại báo mất."""
    goc = dung_mau(tmp)
    td_ok("kiemke", goc)
    cu = {(r["Đường dẫn"], r.get("Trong zip", "")): r["ID"] for r in doc_index(goc)}
    td_ok("init", goc, "--bo-thu-muc", "New folder")
    td_ok("kiemke", goc)
    moi = {(r["Đường dẫn"], r.get("Trong zip", "")): r["ID"] for r in doc_index(goc)}
    assert all(cu[k] == v for k, v in moi.items()), "ID file còn lại bị đổi"
    td_doi = doc_json(goc, "_thaydoi.json")
    assert not td_doi["moi"], td_doi
    assert set(td_doi["mat"]) == {v for k, v in cu.items() if k[0] in BI_LOAI}, td_doi


def chu_ky_thu_muc(d):
    out = {}
    for ten in sorted(os.listdir(d)):
        p = os.path.join(d, ten)
        if os.path.isfile(p):
            with open(p, "rb") as fh:
                out[ten] = hashlib.md5(fh.read()).hexdigest()
    return out


def ban_sao_sofitel(tmp, bo_thu_muc=None):
    """Bản sao: chỉ chép dữ liệu kiemke vào thư mục tạm; hồ sơ giữ nguyên chỗ cũ, chỉ đọc."""
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    goc = os.path.join(tmp, "SOFITEL_copy")
    data = os.path.join(goc, "10_TENDER", "_data")
    os.makedirs(data)
    for ten in ("FILE_INDEX.kiemke-tool.json", "_chuky.json"):
        shutil.copy2(os.path.join(data_that, ten), data)
    with open(os.path.join(data_that, "_cauhinh.json"), encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg["input"] = os.path.abspath(os.path.join(GOC_SOFITEL, cfg["input"]))  # join với đường dẫn tuyệt đối trả về chính nó
    if bo_thu_muc is not None:
        cfg["bo_thu_muc"] = bo_thu_muc
    with open(os.path.join(data, "_cauhinh.json"), "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False)
    return goc, cfg


@kiem(0, "sofitel")
def gd0_sofitel_kiemke_khong_them_file_new_folder(tmp):
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    truoc = chu_ky_thu_muc(data_that)

    goc, cfg = ban_sao_sofitel(tmp)
    assert "New folder" in cfg.get("bo_thu_muc", []), "_cauhinh.json của SOFITEL chưa có bo_thu_muc"
    thu_muc = os.path.join(cfg["input"], "New folder")
    so_file = sum(len(f) for _, _, f in os.walk(thu_muc))
    assert so_file > 0, "New folder rỗng: phép kiểm này không chứng minh được gì"

    cu = {(r["Đường dẫn"], r.get("Trong zip", "")): r["ID"] for r in doc_index(goc)}
    td_ok("kiemke", goc)
    rows = doc_index(goc)
    lot = sorted(r["Đường dẫn"] for r in rows if r["Đường dẫn"].lower().startswith("new folder"))
    assert not lot, f"{len(lot)} file từ New folder vào FILE_INDEX: {lot[:5]}"
    moi = {(r["Đường dẫn"], r.get("Trong zip", "")): r["ID"] for r in rows}
    doi_id = {k: (cu[k], v) for k, v in moi.items() if k in cu and cu[k] != v}
    assert not doi_id, f"ID đổi: {list(doi_id.items())[:5]}"
    mat = [k for k in cu if k not in moi]
    assert not mat, f"file trong FILE_INDEX cũ bị mất: {mat[:5]}"

    # Đối chứng: cùng hồ sơ, không có bo_thu_muc thì New folder phải lọt vào.
    tmp2 = os.path.join(tmp, "doi_chung")
    goc2, _ = ban_sao_sofitel(tmp2, bo_thu_muc=[])
    td_ok("kiemke", goc2)
    lot2 = [r for r in doc_index(goc2) if r["Đường dẫn"].lower().startswith("new folder")]
    assert len(lot2) >= so_file, f"đối chứng chỉ có {len(lot2)} dòng, trên đĩa {so_file} file"

    assert chu_ky_thu_muc(data_that) == truoc, "_data thật đã bị ghi"
    print(f"      SOFITEL: {so_file} file trong New folder, 0 vào FILE_INDEX (không cấu hình: {len(lot2)}); "
          f"FILE_INDEX {len(cu)} → {len(moi)} dòng, ID cũ giữ nguyên")


# ---------------------------------------------------------------- GĐ1
sys.path.insert(0, os.path.join(ROOT, "SKILL", "_core"))


def ghi_cache(goc, rel, pages):
    """Ghi cache chữ giả của 1 file đã có trong FILE_INDEX (thay cho lệnh trang, khỏi cần PDF thật)."""
    fid = next(r["ID"] for r in doc_index(goc) if r["Đường dẫn"] == rel and not r.get("Trong zip"))
    d = os.path.join(goc, "10_TENDER", "_data", "_text")
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, fid + ".json"), "w", encoding="utf-8") as fh:
        json.dump({"ck": {}, "pages": [{"text": t, "so": "", "ten": "", "lop": "CÓ"} for t in pages]}, fh, ensure_ascii=False)
    return fid


def trang_tim(out):
    """Kết quả lệnh tim -> {'F0001 p3', ...}."""
    return {ln.split("\t")[0] for ln in out.splitlines() if "\t" in ln}


@kiem(1, "mau")
def gd1_chuan_hoa_ma(tmp):
    from scan_core import ma_chuan
    dung = {"51 - 500B": "51-500B", "51-500B": "51-500B", "51–500b": "51-500B", "51 500B": "51-500B",
            "ST 01": "ST-01", "ST-01": "ST-01", "st01": "ST-01", "ST-1": "ST-01", "CR-01A": "CR-01A", "zh 49b": "ZH-49B",
            "AREA53": "AREA 53", "AREA 53": "AREA 53", "Area_53": "AREA 53", "area-1": "AREA 01"}
    for vao, ra in dung.items():
        assert ma_chuan(vao) == ra, f"{vao!r} -> {ma_chuan(vao)!r}, cần {ra!r}"
    for vao in ("", None, "500-200", "ID-GA-L18-001", "51-500BX", "hello", "1-200", "SOFA 12", "51500"):
        assert ma_chuan(vao) is None, f"{vao!r} -> {ma_chuan(vao)!r}, cần None"


@kiem(1, "mau")
def gd1_regex_ma_khong_bat_nham(tmp):
    from scan_core import regex_ma
    bang = {
        "51-500B": (["51-500B", "51 - 500B", "51- 500B", "51 -500B", "53\n51 -\n500B", "x 51–500B y", "(51-500B)"],
                    ["51-500", "51-500C", "151-500B", "51-5001B", "51-500BX", "51 500B", "A-51-500B", "51-500B-2"]),
        "51-500": (["51-500", "51 - 500", "51-500 Bed"], ["51-500B", "51-5000", "44 51 500"]),
        "ST-01": (["ST-01", "ST 01", "ST01", "ST–01"], ["ST-011", "ST-01A", "ST-1", "st 01", "ST\n01", "XST-01", "ST-01-2"]),
        "AREA 53": (["AREA 53", "AREA53", "Area 53", "AREA_53", "AREA\n53"], ["AREA 530", "AREA 5", "SUBAREA 53"]),
    }
    for ma, (co, khong) in bang.items():
        rx = regex_ma(ma)
        for t in co:
            assert rx.search(t), f"{ma}: phải khớp {t!r}"
        for t in khong:
            assert not rx.search(t), f"{ma}: không được khớp {t!r}"


def dung_mau_tim(tmp):
    goc = dung_mau(tmp)
    inp = os.path.join(goc, "hoso")
    tao_file(inp, "HSMT/AREA53_spec.pdf")
    td_ok("kiemke", goc)
    a = ghi_cache(goc, "HSMT/AREA53_spec.pdf", ["51 - 500B", "51-500B", "51-500", "151-500B", "51 500B", "khong co"])
    b = ghi_cache(goc, "HSMT/spec.pdf", ["ITEM\n51 -\n500B", "AREA53 ST 01"])
    return goc, a, b


@kiem(1, "mau")
def gd1_tim_ma_du_trang_khong_thua(tmp):
    goc, a, b = dung_mau_tim(tmp)
    out = td_ok("tim", goc, "--ma", "51 - 500B")
    assert trang_tim(out) == {f"{a} p1", f"{a} p2", f"{b} p1"}, out


@kiem(1, "mau")
def gd1_tim_ma_cu_phap_loi_va_nhac(tmp):
    goc, a, b = dung_mau_tim(tmp)
    rc, out = td("tim", goc, "--ma", "khong phai ma")
    assert rc != 0 and "--ma" in out, f"mã sai không báo lỗi: rc={rc}\n{out}"
    # Regex cũ vẫn chạy như trước, nhưng truyền mã trơn thì nhắc vì dễ sót trang.
    out = td_ok("tim", goc, "51-500B")
    assert trang_tim(out) == {f"{a} p2", f"{a} p4"}, out  # regex trơn khớp cả chuỗi con 151-500B ở p4, đúng hành vi cũ
    assert "XEM" in out and "--ma" in out, out
    out = td_ok("tim", goc, r"51\s*-\s*500B")
    assert "XEM" not in out, out


@kiem(1, "mau")
def gd1_tim_file_coi_area53_la_area_53(tmp):
    goc, a, b = dung_mau_tim(tmp)
    for loc in ("AREA 53", "AREA53", "area_53"):
        out = td_ok("tim", goc, "--ma", "51-500B", "--file", loc)
        assert trang_tim(out) == {f"{a} p1", f"{a} p2"}, f"--file {loc}: {out}"


@kiem(1, "mau")
def gd1_nhan_file_pvcv_la_scope(tmp):
    goc = dung_mau(tmp)
    inp = os.path.join(goc, "hoso")
    la_scope = ["DA_PVCV BOHO.xlsx", "pvcv chung.pdf", "HSMT/261007-PVCV.xls", "PVCV as-built.pdf", "TCQ PVCV.xlsx", "PVCV.docx"]
    khong = {"SPVCV2.xlsx": None, "THU_MUC_PVCV/ban_ve.pdf": None, "mail PVCV.msg": "Mail", "PVCV anh.png": None}
    for rel in la_scope + list(khong):
        tao_file(inp, rel)
    with zipfile.ZipFile(os.path.join(inp, "bo.zip"), "w") as z:
        z.writestr("con/PVCV trong zip.xlsx", b"y")
        z.writestr("con/khac.xlsx", b"y")
    tao_file(inp, "bo_qua/x_PVCV.xlsx")
    td_ok("init", goc, "--bo-thu-muc", "bo_qua;New folder")
    td_ok("kiemke", goc)
    rows = doc_index(goc)
    vt = {r["Đường dẫn"]: r.get("Vai trò") for r in rows if not r.get("Trong zip")}
    for rel in la_scope:
        assert vt.get(rel) == "Scope", f"{rel}: {vt.get(rel)!r}"
    for rel, cho in khong.items():
        assert vt.get(rel) == cho, f"{rel}: {vt.get(rel)!r}, cần {cho!r}"
    zvt = {r["Trong zip"]: r.get("Vai trò") for r in rows if r.get("Trong zip")}
    assert zvt == {"con/PVCV trong zip.xlsx": "Scope", "con/khac.xlsx": None}, zvt
    assert "bo_qua/x_PVCV.xlsx" not in vt, "file PVCV trong thư mục đã loại vẫn được nhận"

    # check chấp nhận Scope; đối chứng: giá trị lạ phải bị bắt, không thì phép kiểm trên vô nghĩa.
    _, out = td("check", goc, "--den", "kiemke")
    assert "Scope" not in out, out
    fid = next(r["ID"] for r in rows if r["Đường dẫn"] == "SPVCV2.xlsx")
    with open(os.path.join(goc, "10_TENDER", "_data", "FILE_INDEX.kiemke.json"), "w", encoding="utf-8") as fh:
        json.dump({"rows": [{"ID": fid, "Vai trò": "XYZ"}]}, fh, ensure_ascii=False)
    _, out = td("check", goc, "--den", "kiemke")
    assert "XYZ" in out, f"check không bắt giá trị lạ:\n{out}"


@kiem(1, "mau")
def gd1_pdf_scope_duoc_doc_chu(tmp):
    """PVCV dạng PDF phải vào cache chữ và tim được, bước doichieu sau này mới đọc được."""
    goc = dung_mau(tmp)
    tao_file(os.path.join(goc, "hoso"), "PVCV goi.pdf")
    td_ok("kiemke", goc)
    fid = ghi_cache(goc, "PVCV goi.pdf", ["ST-01 thuoc nha thau"])
    assert trang_tim(td_ok("tim", goc, "--ma", "ST-01")) == {f"{fid} p1"}


def ban_sao_du_lieu(tmp):
    """Bản sao _data (cả cache chữ, bỏ _zip) để chạy lệnh đọc trên dữ liệu thật mà không chạm 10_TENDER."""
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    goc = os.path.join(tmp, "SOFITEL_copy")
    data = os.path.join(goc, "10_TENDER", "_data")
    # bỏ _dmvt*: có _dmvt.json thì build dựng cả DMVT bằng Excel; phép kiểm cần DMVT tự chép file cấu hình
    shutil.copytree(data_that, data, ignore=shutil.ignore_patterns("_zip", "_dmvt*"))
    cau_hinh = os.path.join(data, "_cauhinh.json")
    with open(cau_hinh, encoding="utf-8") as fh:
        cfg = json.load(fh)
    cfg["input"] = os.path.abspath(os.path.join(GOC_SOFITEL, cfg["input"]))
    with open(cau_hinh, "w", encoding="utf-8") as fh:
        json.dump(cfg, fh, ensure_ascii=False)
    return goc


@kiem(1, "sofitel")
def gd1_sofitel_tim_ma_du_7_trang(tmp):
    goc = ban_sao_du_lieu(tmp)
    out = td_ok("tim", goc, "--ma", "51-500B")
    can = {"F0046 p6", "F0046 p9", "F0046 p60", "F0046 p66", "F0046 p67", "F0046 p74", "F1167 p3"}
    co = trang_tim(out)
    assert can <= co, f"thiếu {sorted(can - co)}\n{out}"
    print(f"      SOFITEL: tim --ma \"51-500B\" ra {len(co)} trang (cần đủ {len(can)})")


@kiem(1, "sofitel")
def gd1_sofitel_khong_mat_trang_so_voi_tim_chu_y_nguyen(tmp):
    """Mỗi mã xuất hiện trong cache: regex_ma phải ra ít nhất các trang có đúng chữ đó."""
    import glob
    import re
    from scan_core import ma_chuan, regex_ma
    goc = ban_sao_du_lieu(tmp)
    trang = []
    for f in glob.glob(os.path.join(goc, "10_TENDER", "_data", "_text", "*.json")):
        with open(f, encoding="utf-8") as fh:
            trang.extend(x["text"] for x in json.load(fh)["pages"])
    tho = re.compile(r"(?<![\w-])(\d{2}-\d{3}[A-Z]?|[A-Z]{2,3}-\d{2}[A-Z]?)(?![\w-])")
    ma_co = {m for t in trang for m in tho.findall(t)}
    assert len(ma_co) > 100, f"chỉ thấy {len(ma_co)} mã, phép kiểm không đủ mẫu"
    thieu, n_tho, n_ma = [], 0, 0
    for ma in sorted(ma_co):
        chuan = ma_chuan(ma)
        assert chuan == ma, f"{ma} -> {chuan}"
        # GĐ4 (plan mục 4): mã con 52-500.1 không còn tính vào mã cha 52-500, nên chữ y nguyên cũng không đếm trang chỉ có mã con
        rx, rx_tho = regex_ma(ma), re.compile(rf"(?<![\w-]){re.escape(ma)}(?!\.\d)(?![\w-])")
        a = {i for i, t in enumerate(trang) if rx_tho.search(t)}
        b = {i for i, t in enumerate(trang) if rx.search(t)}
        n_tho, n_ma = n_tho + len(a), n_ma + len(b)
        if a - b:
            thieu.append(ma)
    assert not thieu, f"--ma sót trang của {thieu[:10]}"
    print(f"      SOFITEL: {len(ma_co)} mã; {n_tho} lượt trang theo chữ y nguyên → {n_ma} lượt trang theo --ma")


@kiem(1, "sofitel")
def gd1_sofitel_chua_co_pvcv_that(tmp):
    """SOFITEL chưa có PVCV thật (plan GĐ1); file Sarene trong New folder không được thành Scope của dự án này."""
    goc, _ = ban_sao_sofitel(tmp)
    td_ok("kiemke", goc)
    scope = [r["Đường dẫn"] for r in doc_index(goc) if r.get("Vai trò") == "Scope"]
    assert not scope, scope


# ---------------------------------------------------------------- GĐ2
BANG_MAU = [{"cap3": t, "ma": m} for t, m in (("Tên user", "US.ER"), ("Tên luật", "LU.AT"), ("Tên mô tả local", "MT.LO"),
                                              ("Tên mô tả chung", "MT.CH"), ("Tên đầu mục local", "DM.LO"),
                                              ("Tên đầu mục chung", "DM.CH"), ("Tên LLM", "LL.M"), ("Mã có cách", "LS. HA"))]
LUAT_MAU = {"rows": [{"ten": "bàn", "ma": "LU.AT", "dk": [{"cot": ["Mô tả BOQ"], "rx": "^Luat ap dung"}]}]}
DONG_MAU = {"Khóa": "S!1", "Sheet": "S", "Dòng": 1, "Item": "1", "Nhóm": "NHOM", "Mục phụ": "Dau muc",
            "Mô tả BOQ": "ST-04 Cai ban 12 cai", "ĐV": "Nr", "Mã VL": "ST-04"}
MT_KHOA = {"nhom": "nhom", "mo_ta": "st-# cai ban # cai"}
DM_KHOA = {"nhom": "nhom", "dau_muc": "dau muc", "tien_to": "ST"}


def ghi(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False)


def dung_code(tmp, dong=(DONG_MAU,), luat=LUAT_MAU):
    """Dự án giả + thư mục Code ID giả. TD_CODE_ID_DIR để không chạm _MASTER thật; main() xóa biến này sau mỗi phép kiểm."""
    goc = dung_mau(tmp)
    cd = os.path.join(tmp, "code_id")
    os.environ["TD_CODE_ID_DIR"] = cd
    ghi(os.path.join(cd, "bang_ma.json"), {"rows": BANG_MAU})
    ghi(os.path.join(cd, "luat.json"), luat)
    ghi(os.path.join(goc, "10_TENDER", "_data", "SCOPE.boq-tool.json"), {"rows": list(dong)})
    return goc, cd


def ket_qua(goc):
    rows = doc_json(goc, "SCOPE.phanloai-tool.json")["rows"]
    return {r["Khóa"]: r for r in rows}


@kiem(2, "mau")
def gd2_thu_tu_lop_tung_lop_mot(tmp):
    """Thêm dần từ lớp thấp lên lớp cao: mỗi lần thêm, lớp mới phải thắng. Sai thứ tự là ra mã sai trông như đúng."""
    goc, cd = dung_code(tmp)
    data = os.path.join(goc, "10_TENDER", "_data")
    local, chung = {"mo_ta": [], "dau_muc": []}, {"mo_ta": [], "dau_muc": []}

    def chay(ma_can, nguon_can):
        ghi(os.path.join(cd, "nho_local.json"), local)
        ghi(os.path.join(cd, "nho.json"), chung)
        td_ok("phanloai", goc)
        r = ket_qua(goc)["S!1"]
        ten = next((b["cap3"] for b in BANG_MAU if b["ma"] == ma_can), None)
        assert r.get("Mã cấp 3") == ten and r.get("Nguồn phân loại") == nguon_can, f"cần {ten}/{nguon_can}, ra {r}"

    chay(None, None)
    assert ket_qua(goc)["S!1"]["Cờ phân loại"] == "Chờ LLM"
    ghi(os.path.join(data, "_phanloai_llm.json"), {"rows": [{"Khóa": "S!1", "Mã cấp 3": "Tên LLM"}]})
    chay("LL.M", "LLM")
    chung["dau_muc"].append(dict(DM_KHOA, ma="DM.CH", du_an="T"))
    chay("DM.CH", "Nhớ-đầu-mục")
    local["dau_muc"].append(dict(DM_KHOA, ma="DM.LO", du_an="T"))
    chay("DM.LO", "Nhớ-đầu-mục")
    chung["mo_ta"].append(dict(MT_KHOA, ma="MT.CH", du_an="T"))
    chay("MT.CH", "Nhớ-mô-tả")
    local["mo_ta"].append(dict(MT_KHOA, ma="MT.LO", du_an="T"))
    chay("MT.LO", "Nhớ-mô-tả")
    ghi(os.path.join(cd, "luat.json"), {"rows": [{"ten": "x", "ma": "LU.AT", "dk": [{"cot": ["Mô tả BOQ"], "rx": "Cai ban"}]}]})
    chay("LU.AT", "Luật")
    ghi(os.path.join(data, "SCOPE.user.json"), {"rows": [{"Khóa": "S!1", "Mã cấp 3": "Tên user"}]})
    chay("US.ER", "User")


@kiem(2, "mau")
def gd2_user_vao_nho_local_con_llm_va_doan_thi_khong(tmp):
    """Quyết định 5: chỉ ô user sửa mới vào bộ nhớ; không ghi vào nho.json chung."""
    d2 = dict(DONG_MAU, **{"Khóa": "S!2", "Dòng": 2, "Mô tả BOQ": "ST-04 Cai ghe"})
    d3 = dict(DONG_MAU, **{"Khóa": "S!3", "Dòng": 3, "Mô tả BOQ": "ST-04 Cai tu", "Mục phụ": "Muc khac"})
    goc, cd = dung_code(tmp, (DONG_MAU, d2, d3))
    data = os.path.join(goc, "10_TENDER", "_data")
    ghi(os.path.join(data, "SCOPE.user.json"), {"rows": [{"Khóa": "S!1", "Mã cấp 3": "Tên user"}, {"Khóa": "S!3", "Mã cấp 3": "KHÔNG RÕ"}]})
    ghi(os.path.join(data, "_phanloai_llm.json"), {"rows": [{"Khóa": "S!2", "Mã cấp 3": "Tên LLM"}]})
    td_ok("phanloai", goc)
    with open(os.path.join(cd, "nho_local.json"), encoding="utf-8") as fh:
        local = json.load(fh)
    assert [x["ma"] for x in local["mo_ta"]] == ["US.ER"] and not local["dau_muc"], local
    assert local["mo_ta"][0]["mo_ta"] == MT_KHOA["mo_ta"] and not any(c.isdigit() for c in local["mo_ta"][0]["mo_ta"]), local
    assert not os.path.exists(os.path.join(cd, "nho.json")), "phanloai không được tự tạo nho.json"
    # Chạy lại không nhân đôi dòng nhớ; đổi ý thì cập nhật, không thêm.
    td_ok("phanloai", goc)
    ghi(os.path.join(data, "SCOPE.user.json"), {"rows": [{"Khóa": "S!1", "Mã cấp 3": "Tên luật"}]})
    td_ok("phanloai", goc)
    with open(os.path.join(cd, "nho_local.json"), encoding="utf-8") as fh:
        assert [x["ma"] for x in json.load(fh)["mo_ta"]] == ["LU.AT"]


@kiem(2, "mau")
def gd2_highlight_dung_cac_lop(tmp):
    """Plan mục 3. Sót highlight thì tender duyệt theo ngoại lệ sẽ bỏ qua dòng máy không chắc."""
    goc, cd = dung_code(tmp)
    data = os.path.join(goc, "10_TENDER", "_data")
    nho = os.path.join(cd, "nho.json")
    luat = os.path.join(cd, "luat.json")
    rong = {"rows": []}

    def co():
        td_ok("phanloai", goc)
        return ket_qua(goc)["S!1"].get("Cờ phân loại", "")

    ghi(luat, rong)
    ghi(nho, {"mo_ta": [dict(MT_KHOA, ma="MT.CH", du_an="T")]})
    assert co() == "", "nhớ mô tả cùng dự án không highlight"
    ghi(nho, {"mo_ta": [dict(MT_KHOA, ma="MT.CH", du_an="DA_KHAC")]})
    assert co() == "Dự án khác"
    ghi(nho, {"dau_muc": [dict(DM_KHOA, ma="DM.CH", du_an="T")]})
    assert co() == "Nhớ-đầu-mục", "nhớ đầu mục luôn highlight, kể cả cùng dự án"
    ghi(nho, {})
    assert co() == "Chờ LLM"
    ghi(os.path.join(data, "_phanloai_llm.json"), {"rows": [{"Khóa": "S!1", "Mã cấp 3": "Tên LLM"}]})
    assert co() == "LLM"
    ghi(os.path.join(data, "_phanloai_llm.json"), {"rows": [{"Khóa": "S!1", "Mã cấp 3": "KHÔNG RÕ"}]})
    assert co() == "KHÔNG RÕ"
    assert ket_qua(goc)["S!1"]["Mã cấp 3"] == "KHÔNG RÕ"
    ghi(luat, {"rows": [{"ten": "x", "ma": "LU.AT", "dk": [{"cot": ["Mô tả BOQ"], "rx": "Cai ban"}]}]})
    assert co() == "Luật chưa duyệt", "dự án mới: luật phải highlight"
    td_ok("phanloai", goc, "--duyet-luat")
    assert ket_qua(goc)["S!1"].get("Cờ phân loại", "") == ""
    assert doc_json(goc, "_cauhinh.json")["luat_da_duyet"] is True
    ghi(os.path.join(data, "SCOPE.user.json"), {"rows": [{"Khóa": "S!1", "Mã cấp 3": "Tên user"}]})
    assert co() == "", "user đã chọn thì không highlight"


@kiem(2, "mau")
def gd2_pham_vi_dong(tmp):
    """Quyết định 4: mọi dòng có đơn vị, kể cả loại trừ; GR&Prelims và Provisional sum để trống."""
    dong = [dict(DONG_MAU, **{"Khóa": f"{sh}!1", "Sheet": sh}) for sh in ("2. Finishing", "3. Furniture", "1. GR&Prelims", "4. Provisional sum")]
    khong_dv = {k: v for k, v in DONG_MAU.items() if k != "ĐV"}
    dong.append(dict(khong_dv, **{"Khóa": "2. Finishing!2", "Sheet": "2. Finishing", "Dòng": 2}))
    goc, _ = dung_code(tmp, dong)
    ghi(os.path.join(goc, "10_TENDER", "_data", "SCOPE.doichieu-1.json"), {"rows": [{"Khóa": "3. Furniture!1", "Trạng thái": "Loại trừ"}]})
    td_ok("phanloai", goc)
    assert set(ket_qua(goc)) == {"2. Finishing!1", "3. Furniture!1"}, set(ket_qua(goc))


@kiem(2, "mau")
def gd2_llm_chi_trong_danh_sach_dong(tmp):
    goc, _ = dung_code(tmp)
    data = os.path.join(goc, "10_TENDER", "_data")
    td_ok("phanloai", goc)
    truoc = doc_json(goc, "SCOPE.phanloai-tool.json")
    for xau in ({"Khóa": "S!1", "Mã cấp 3": "Tên lạ"}, {"Khóa": "S!99", "Mã cấp 3": "Tên LLM"}):
        ghi(os.path.join(data, "_phanloai_llm.json"), {"rows": [xau]})
        rc, out = td("phanloai", goc)
        assert rc != 0 and "_phanloai_llm.json" in out, f"{xau}: rc={rc}\n{out}"
        assert doc_json(goc, "SCOPE.phanloai-tool.json") == truoc, "LLM sai nhưng vẫn ghi kết quả"
    # check báo dòng chờ LLM; user chọn giá trị ngoài danh sách thì là LỖI.
    os.remove(os.path.join(data, "_phanloai_llm.json"))
    _, out = td("check", goc, "--den", "boq")
    assert "chưa có Mã cấp 3" in out and "phanloai-tool" not in out, out
    ghi(os.path.join(data, "SCOPE.user.json"), {"rows": [{"Khóa": "S!1", "Mã cấp 3": "Tên lạ"}]})
    _, out = td("check", goc, "--den", "boq")
    assert "Tên lạ" in out and "LỖI" in out, out


@kiem(2, "mau")
def gd2_dong_kich_thuoc_khong_dung_chung_cau_tra_loi(tmp):
    """'overall size #' ở 2 đầu mục là 2 món: không qua nhớ mô tả, câu LLM không lan sang đầu mục khác."""
    kt = {"Mô tả BOQ": "overall size 1460x180x2500mmH", "Mã VL": ""}
    a = dict(DONG_MAU, **kt, **{"Khóa": "S!1", "Mục phụ": "Tu ao"})
    b = dict(DONG_MAU, **kt, **{"Khóa": "S!2", "Dòng": 2, "Mục phụ": "Tu ao"})
    c = dict(DONG_MAU, **kt, **{"Khóa": "S!3", "Dòng": 3, "Mục phụ": "Vach"})
    goc, cd = dung_code(tmp, (a, b, c))
    ghi(os.path.join(cd, "nho.json"), {"mo_ta": [{"nhom": "nhom", "mo_ta": "overall size #x#x#mmh", "ma": "MT.CH", "du_an": "T"}]})
    ghi(os.path.join(goc, "10_TENDER", "_data", "_phanloai_llm.json"), {"rows": [{"Khóa": "S!1", "Mã cấp 3": "Tên LLM"}]})
    td_ok("phanloai", goc)
    kq = ket_qua(goc)
    assert kq["S!1"]["Mã cấp 3"] == kq["S!2"]["Mã cấp 3"] == "Tên LLM", kq
    assert "Mã cấp 3" not in kq["S!3"], kq


@kiem(2, "mau")
def gd2_ma_co_dau_cach_va_ma_khong_co_trong_bang(tmp):
    goc, cd = dung_code(tmp)
    ghi(os.path.join(cd, "nho.json"), {"mo_ta": [dict(MT_KHOA, ma="LS.HA", du_an="T")]})
    td_ok("phanloai", goc)
    assert ket_qua(goc)["S!1"]["Mã cấp 3"] == "Mã có cách", "'LS.HA' phải khớp 'LS. HA' của bảng mã"
    ghi(os.path.join(cd, "nho.json"), {"mo_ta": [dict(MT_KHOA, ma="XX.YY", du_an="T")]})
    out = td_ok("phanloai", goc)
    assert "XX.YY" in out and ket_qua(goc)["S!1"]["Cờ phân loại"] == "Chờ LLM", out


@kiem(2, "mau")
def gd2_ham_thuan(tmp):
    import code_id as ci
    for t in ("overall size 1460x180x2500mmH", "size 1500x2700mmH", "overall size (1485+830)x2500mmH", "100mm thick",
              "overall 1200x580x900mmH", "overall size D450x550mmH", "overall size 500x500; refer drawing no ID-GA-001"):
        assert ci.chi_kich_thuoc(t), f"phải là dòng kích thước: {t}"
    for t in ("", None, "to bathroom floor", "ST-04 Marble in honed finish", "Bench, overall size 900x450mmH",
              "51-500 L-shaped sofa, 1440x1800x820mmH", "overall size", "WD-01 to corridor"):
        assert not ci.chi_kich_thuoc(t), f"không phải dòng kích thước: {t}"
    bang = {"ST-04 Marble": "ST", "GL-03; MT-02": "GL"}
    for vl, tt in bang.items():
        assert ci.tien_to_ma({"Mã VL": vl}) == tt
    assert ci.tien_to_ma({"Mô tả BOQ": "51-400 Table"}) == "F4"
    assert ci.tien_to_ma({"Mô tả BOQ": "Plain"}) == ""
    # Bộ nhớ không giữ số; nhãn hòa nhau thì bỏ, không đoán.
    dong = [{"Nhóm": "N", "Mô tả BOQ": "ST-04 Marble 20mm", "Mục phụ": "A", "Mã VL": "ST-04", "_ma": m} for m in ("X.A", "X.B")]
    nho = ci.hoc(dong + [{"Nhóm": "N", "Mô tả BOQ": "Khac 12", "_ma": "X.C"}], "T")
    assert [x["ma"] for x in nho["mo_ta"]] == ["X.C"] and not nho["dau_muc"], nho
    assert not any(c.isdigit() for x in nho["mo_ta"] for c in x["mo_ta"]), nho


@kiem(2, "mau")
def gd2_du_lieu_master_hop_le(tmp):
    """Dữ liệu thật trong _MASTER, không phải số liệu dự án: bảng mã đọc được, luật chỉ trỏ tới mã có thật."""
    import code_id as ci
    bang = ci.nap_bang()
    assert bang and len(bang.rows) == 81, "cần 81 mã từ Datawork cột G"
    assert len(bang.ten) == len(bang.rows) == len(bang.ma), "D hoặc G bị trùng"
    assert any(r["ma"] == "LS. HA" for r in bang.rows), "quyết định 3: giữ đúng cột G"
    for x in ci._doc("luat.json")["rows"]:
        assert bang.co(x["ma"]), f"luật {x['ten']} trỏ mã không có: {x['ma']}"
        for d in x["dk"]:
            assert d["cot"] and re.compile(d["rx"]), x
    ban = ci.nap_luat()
    assert ban.xet({"Mô tả BOQ": "x", "Mục phụ": "51 - 200 Floor rug"})[0] == "LO.RUG"
    assert ban.xet({"Mô tả BOQ": "x", "Mục phụ": "52-200 wall to wall carpet"})[0] == "FF.CP"
    assert ban.xet({"Mô tả BOQ": "size 1x2", "Mục phụ": "51-300 Blackout"})[0] == "DE.CU"
    assert ban.xet({"Mô tả BOQ": "51-400 Table"})[0] is None, "plan chưa có luật 4xx"
    # Đúng cách chạy qua junction .claude\skills: lấy đúng _MASTER, không phải .claude\_MASTER.
    qua_junction = os.path.join(ROOT, ".claude", "skills", "10-tender", "tool", "td.py")
    if os.path.exists(qua_junction):
        os.environ.pop("TD_CODE_ID_DIR", None)
        r = subprocess.run([sys.executable, "-I", qua_junction, "cot", "--buoc", "phanloai-tool"], capture_output=True, encoding="utf-8", errors="replace")
        assert r.returncode == 0 and "Renovation" in r.stdout, r.stdout + r.stderr


def chu_ky_master():
    d = os.path.join(ROOT, "_MASTER", "code_id")
    return {t: hashlib.md5(open(os.path.join(d, t), "rb").read()).hexdigest() for t in sorted(os.listdir(d))}


@kiem(2, "sofitel")
def gd2_sofitel_dat_moc_hoc_cheo(tmp):
    """Mốc không nới sau khi thấy kết quả (quyết định 17). do_codeid.py đọc đáp án SSF, không dùng nho.json."""
    r = subprocess.run([sys.executable, "-I", os.path.join(os.path.dirname(os.path.realpath(__file__)), "do_codeid.py")],
                       capture_output=True, encoding="utf-8", errors="replace")
    dong = [x for x in r.stdout.splitlines() if x.startswith(("Mốc GĐ2", "GỘP"))]
    assert r.returncode == 0, "\n".join(dong) or r.stdout + r.stderr
    print("      SOFITEL:", dong[0])
    print("      SOFITEL:", dong[-1])


@kiem(2, "sofitel")
def gd2_sofitel_phanloai_tren_ban_sao(tmp):
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    truoc, truoc_master = chu_ky_thu_muc(data_that), chu_ky_master()
    goc = ban_sao_du_lieu(tmp)
    # Thư mục Code ID tạm: bảng mã + luật thật, chưa có nho.json (chưa nạp SSF).
    cd = os.path.join(tmp, "code_id")
    os.makedirs(cd)
    for t in ("bang_ma.json", "luat.json"):
        shutil.copy2(os.path.join(ROOT, "_MASTER", "code_id", t), cd)
    os.environ["TD_CODE_ID_DIR"] = cd
    llm = os.path.join(goc, "10_TENDER", "_data", "_phanloai_llm.json")  # phép kiểm này đo lúc chưa có bộ nhớ, chưa có câu LLM
    if os.path.exists(llm):
        os.remove(llm)
    _, truoc_check = td("check", goc)
    out = td_ok("phanloai", goc)
    kq = ket_qua(goc)
    assert len(kq) == 659, f"plan quyết định 4 có 659 dòng, ra {len(kq)}\n{out}"
    assert not [k for k in kq if k.startswith(("1. GR", "4. Prov"))], "GR&Prelims hoặc Provisional sum bị gán"
    assert all(r.get("Nguồn phân loại") == "Luật" or r.get("Cờ phân loại") == "Chờ LLM" for r in kq.values()), "chưa có bộ nhớ thì chỉ có Luật hoặc Chờ LLM"
    n_luat = sum(1 for r in kq.values() if r.get("Nguồn phân loại") == "Luật")
    assert all(r["Cờ phân loại"] == "Luật chưa duyệt" for r in kq.values() if r.get("Nguồn phân loại") == "Luật")
    _, sau_check = td("check", goc)
    loi_truoc = {x for x in truoc_check.splitlines() if x.startswith("LỖI")}
    loi_moi = [x for x in sau_check.splitlines() if x.startswith("LỖI") and x not in loi_truoc]
    assert not loi_moi, f"check thêm LỖI sau phanloai: {loi_moi[:5]}"
    assert chu_ky_thu_muc(data_that) == truoc, "_data thật đã bị ghi"
    assert chu_ky_master() == truoc_master, "_MASTER thật đã bị ghi"
    print(f"      SOFITEL: 659 dòng, {n_luat} theo Luật, {len(kq) - n_luat} chờ nhớ/LLM; check không thêm LỖI")


@kiem(2, "sofitel")
def gd2_sofitel_build_khong_doi_ngay_cap_nhat(tmp):
    """Phân loại không phải thay đổi nội dung: build sau phanloai không được đổi Ngày cập nhật của dòng nào."""
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    truoc = chu_ky_thu_muc(data_that)
    goc = ban_sao_du_lieu(tmp)
    cd = os.path.join(tmp, "code_id")
    os.makedirs(cd)
    for t in ("bang_ma.json", "luat.json"):
        shutil.copy2(os.path.join(ROOT, "_MASTER", "code_id", t), cd)
    os.environ["TD_CODE_ID_DIR"] = cd
    duong = os.path.join(goc, "10_TENDER", "_data", "_snapshot.json")
    snap = doc_json(goc, "_snapshot.json")
    snap.pop("boq_track", None)  # không kéo sửa tay từ file thật; lệch hàng ở file thật là chuyện riêng của build thật
    snap.pop("file_index", None)
    ghi(duong, snap)
    td_ok("build", goc)
    s1 = doc_json(goc, "_snapshot.json")
    td_ok("phanloai", goc)
    td_ok("build", goc)
    s2 = doc_json(goc, "_snapshot.json")
    moi, cu = s2["reg"]["SCOPE"], s1["reg"]["SCOPE"]
    co_ma = sum(1 for r in moi.values() if r.get("Mã cấp 3"))
    assert co_ma > 100, f"sheet SCOPE chỉ có {co_ma} dòng có Mã cấp 3: phép kiểm không chứng minh được gì"
    doi = [k for k in moi if moi[k].get("Ngày cập nhật") != cu[k].get("Ngày cập nhật")]
    assert not doi, f"{len(doi)} dòng đổi ngày: {doi[:5]}"
    for sh, info in s2["ct"].items():
        assert {"Mã cấp 3", "Bên thực hiện"} <= set(info["cols"]), f"{sh}: GĐ3 phải có 2 cột nhập ở sheet BOQ"
        assert not {"Nguồn phân loại", "Cờ phân loại", "Code ID"} & set(info["cols"]), f"{sh}: cột máy tính không được kéo ngược"
    assert chu_ky_thu_muc(data_that) == truoc, "_data thật đã bị ghi"
    print(f"      SOFITEL: {len(moi)} dòng SCOPE, {co_ma} có Mã cấp 3, 0 dòng đổi Ngày cập nhật")


# ---------------------------------------------------------------- GĐ3
CHUYEN = os.path.join(os.path.dirname(os.path.realpath(__file__)), "chuyen_ben.py")
COT_NOI_BO = ("Vùng", "Nhóm", "Mã phòng", "Mã cấp 3", "Code ID", "Bên thực hiện", "Mã cấu tạo", "Bản vẽ hiệu lực", "Link vật liệu",
              "Trạng thái", "Mã TCQ", "Nội dung cuối (VN)", "Loại thay đổi", "Việc còn mở", "Rev/ngày/đợt", "Ngày cập nhật")
SH = "2. Finishing"


def chuyen(goc, lenh="ghi"):
    r = subprocess.run([sys.executable, "-I", CHUYEN, lenh, goc], capture_output=True, encoding="utf-8", errors="replace")
    return r.returncode, r.stdout + r.stderr


def dung_cu(tmp):
    """Dữ liệu dạng cũ: đủ các ca của plan mục 6, kể cả dòng mẫu _loc bị dòng cụ thể ghi đè."""
    keys = [f"{SH}!{i}" for i in range(1, 9)]
    dong = [dict(DONG_MAU, **{"Khóa": k, "Sheet": SH, "Dòng": i, "Nhóm": "PHA DO" if i <= 2 else "NHOM"}) for i, k in enumerate(keys, 1)]
    goc, cd = dung_code(tmp, dong)
    data = os.path.join(goc, "10_TENDER", "_data")
    ghi(os.path.join(data, "SCOPE.doichieu-1.json"), {"rows": [
        {"_loc": {"Nhóm": "^PHA DO"}, "Nhóm thực hiện": "NTP", "NTP chi tiết": "Phá dỡ", "Trạng thái": "Base"},
        {"Khóa": keys[1], "Nhóm thực hiện": "Owner", "Trạng thái": "Loại trừ"},  # chi tiết thừa hưởng từ _loc: bỏ
        {"Khóa": keys[2], "Nhóm thực hiện": "NTP", "NTP chi tiết": "Đá", "Mã TCQ": "TCQ01-1"},
        {"Khóa": keys[3], "Nhóm thực hiện": "Nhà máy"},  # dòng chỉ có cột cũ: bỏ cả dòng
        {"Khóa": keys[4], "Nhóm thực hiện": "Gói khác", "Việc còn mở": "giữ"},
    ]})
    ghi(os.path.join(data, "SCOPE.doichieu-2.json"), {"rows": [
        {"Khóa": keys[5], "Nhóm thực hiện": "Chung", "Trạng thái": "Base"},
        {"Khóa": keys[6], "Nhóm thực hiện": "NTP", "Trạng thái": "Base"},  # NTP không có chi tiết: không được mất
        {"Khóa": keys[7], "Nhóm thực hiện": "Nhà máy", "NTP chi tiết": "Gương", "Trạng thái": "Base"},  # 2 bên ghi rõ: giữ cả 2
    ]})
    ghi(os.path.join(data, "MAT.mat.json"), {"rows": [{"Mã VL": "ST-04", "Mô tả": "giữ", "Nhóm thực hiện": "Nhà máy", "Trạng thái cuối": "Theo spec"}]})
    can = {keys[0]: "Phá dỡ", keys[1]: "Owner", keys[2]: "Đá", keys[3]: "Nhà máy", keys[4]: "Gói khác", keys[6]: "NTP", keys[7]: "Nhà máy; Gương"}
    return goc, data, keys, can


@kiem(3, "mau")
def gd3_chuyen_tung_gia_tri_khong_mat(tmp):
    """Sai ở đây là mất dữ liệu trong khi check vẫn báo 0 LỖI, nên so từng khóa, không chỉ đếm."""
    goc, data, keys, can = dung_cu(tmp)
    _, out = td("check", goc)
    assert "cột lạ 'Nhóm thực hiện'" in out, f"đối chứng: dữ liệu cũ phải bị check bắt\n{out}"
    with open(os.path.join(data, "SCOPE.doichieu-1.json"), encoding="utf-8") as fh:
        goc_1 = fh.read()
    rc, out = chuyen(goc, "xem")
    assert rc == 0 and not os.path.exists(os.path.join(data, "SCOPE.ben.json")), f"lệnh xem không được ghi\n{out}"
    with open(os.path.join(data, "SCOPE.doichieu-1.json"), encoding="utf-8") as fh:
        assert fh.read() == goc_1, "lệnh xem đã đổi file"
    rc, out = chuyen(goc)
    assert rc == 0, out
    ben = {r["Khóa"]: r["Bên thực hiện"] for r in doc_json(goc, "SCOPE.ben.json")["rows"]}
    assert ben == can, f"{ben}\ncần {can}"
    for t in ("SCOPE.doichieu-1.json", "SCOPE.doichieu-2.json", "MAT.mat.json"):
        with open(os.path.join(data, t), encoding="utf-8") as fh:
            noi = fh.read()
        assert "Nhóm thực hiện" not in noi and "NTP chi tiết" not in noi, t
    rows1 = doc_json(goc, "SCOPE.doichieu-1.json")["rows"]
    assert not any(r.get("Khóa") == keys[3] for r in rows1), "dòng chỉ còn khóa phải bỏ"
    assert any(r.get("Khóa") == keys[2] and r.get("Mã TCQ") == "TCQ01-1" for r in rows1), "cột khác bị mất"
    assert doc_json(goc, "MAT.mat.json")["rows"][0]["Mô tả"] == "giữ"
    _, out = td("check", goc)
    assert "cột lạ" not in out and "không được ghi" not in out, out
    # Sao lưu đúng bản cũ; chạy lần 2 phải từ chối.
    bk = os.listdir(os.path.join(data, "_backup"))
    assert len(bk) == 1
    with open(os.path.join(data, "_backup", bk[0], "SCOPE.doichieu-1.json"), encoding="utf-8") as fh:
        assert fh.read() == goc_1, "sao lưu sai"
    rc, out = chuyen(goc)
    assert rc != 0 and "lần 2" in out, out


@kiem(3, "mau")
def gd3_chay_lai_doichieu_khong_mat_ben(tmp):
    """SKILL.md mục 4 cho xóa JSON của bước khi chạy lại; bên thực hiện không được nằm trong file doichieu."""
    goc, data, keys, can = dung_cu(tmp)
    assert chuyen(goc)[0] == 0
    os.remove(os.path.join(data, "SCOPE.doichieu-1.json"))
    os.remove(os.path.join(data, "SCOPE.doichieu-2.json"))
    out = td_ok("loc", goc, "SCOPE", "--cot", "Khóa,Bên thực hiện", "--max", "99")
    co = {json.loads(x)["Khóa"]: json.loads(x)["Bên thực hiện"] for x in out.splitlines()[1:]}
    assert {k: v for k, v in co.items() if v} == can, co
    # doichieu ghi cột này là lỗi, tránh LLM gán lại.
    ghi(os.path.join(data, "SCOPE.doichieu-1.json"), {"rows": [{"Khóa": keys[0], "Bên thực hiện": "X"}]})
    _, out = td("check", goc)
    assert "không được ghi cột 'Bên thực hiện'" in out, out


def dung_boq(tmp, nho=True):
    """BOQ xlsx mẫu thật + bảng mã giả. 7 dòng cần mã (mỗi lớp 1 dòng), 1 dòng không đơn vị, 1 dòng prelims."""
    import openpyxl
    goc, cd = dung_code(tmp)
    inp = os.path.join(goc, "hoso")
    wb = openpyxl.Workbook()
    wb.active.title = SH
    ws = wb.active
    ws.append(["Tieu de"])
    ws.append([])
    ws.append(["Item", "Description", None, None, "Quantity", "Unit"])
    ws.append([])
    ws.append([None, "VUNG"])
    ws.append([None, None, "NHOM A"])
    ws.append([None, None, None, "Dau muc"])
    for i, (mt, kl, dv) in enumerate([("Luat ap dung 1", 5, "Nr"), ("Cai ghe go", 3, "m2"), ("Cai tu 12", 2, "Nr"), ("Cai ke", 1, "Nr"),
                                      ("Cai ke", 4, "Nr"), ("Chu thich khong don vi", None, None), ("Cai cua", 2, "Nr")], 1):
        ws.append([i, mt, None, None, kl, dv])
    w2 = wb.create_sheet("1. GR&Prelims")
    w2.append(["Item", "Description", None, None, "Quantity", "Unit"])
    w2.append([])
    w2.append([1, "Bao hiem", None, None, 1, "Item"])
    wb.save(os.path.join(inp, "boq.xlsx"))
    td_ok("boq", goc)
    if nho:
        ghi(os.path.join(cd, "nho.json"), {"mo_ta": [{"nhom": "nhom a", "mo_ta": "cai ghe go", "ma": "MT.CH", "du_an": "T"}]})
    data = os.path.join(goc, "10_TENDER", "_data")
    ghi(os.path.join(data, "SCOPE.ben.json"), {"rows": [{"Khóa": f"{SH}!8", "Bên thực hiện": "Nhà máy"}]})
    td_ok("phanloai", goc)
    return goc, cd, data


def file_track(goc):
    import glob
    return glob.glob(os.path.join(goc, "10_TENDER", "03_boq", "T_BOQ_TRACK_*.xlsx"))[0]


def cot_cua(ws, hang=3):
    return {c.value: c.column_letter for c in ws[hang] if c.value}


def co_mau(o):
    return o.fill is not None and o.fill.fgColor is not None and o.fill.fill_type == "solid"


@kiem(3, "mau")
def gd3_cot_moi_o_sheet_boq_va_scope(tmp):
    import openpyxl
    goc, cd, data = dung_boq(tmp)
    td_ok("build", goc)
    wb = openpyxl.load_workbook(file_track(goc))
    ws = wb[SH]
    c = cot_cua(ws)
    for x in COT_NOI_BO:
        assert x in c, f"thiếu cột {x}"
    thu_tu = [c[x] for x in ("Mã cấp 3", "Code ID", "Bên thực hiện", "Mã cấu tạo", "Bản vẽ hiệu lực")]
    assert [ws[f"{l}3"].column for l in thu_tu] == list(range(ws[f"{thu_tu[0]}3"].column, ws[f"{thu_tu[0]}3"].column + 5)), "sai thứ tự plan mục 3"
    for x in COT_NOI_BO:
        assert not ws.column_dimensions[c[x]].hidden, f"{x} bị ẩn, user muốn thấy đủ cột như BOQ_TONGHOP"
    # Dòng 8–11, 13–14 cần mã; 13 không có đơn vị ở fixture là hàng 13.
    m, ci = c["Mã cấp 3"], c["Code ID"]
    assert ws[f"{m}8"].value == "Tên luật" and not str(ws[f"{m}8"].value).startswith("="), "Mã cấp 3 phải là giá trị thật"
    assert ws[f"{ci}8"].value == f'=IFERROR(INDEX(_CODE!$G:$G,MATCH({m}8,_CODE!$D:$D,0))&"","")', ws[f"{ci}8"].value
    assert ws[f"{c['Bên thực hiện']}8"].value == "Nhà máy"
    dv = [d for d in ws.data_validations.dataValidation if d.formula1 and "_CODE!$D" in d.formula1]
    assert len(dv) == 1 and dv[0].type == "list" and dv[0].formula1 == f"=_CODE!$D$2:$D${len(BANG_MAU) + 2}", dv
    co_dv = {str(x) for x in dv[0].sqref.ranges}
    can_co = {f"{m}{i}" for i in (8, 9, 10, 11, 12, 14)}
    assert co_dv == can_co, f"droplist {sorted(co_dv)}, cần {sorted(can_co)}"
    w2 = wb["1. GR&Prelims"]
    c2 = cot_cua(w2, 1)
    assert not w2[f"{c2['Code ID']}3"].value and not w2[f"{c2['Mã cấp 3']}3"].value, "prelims phải để trống"
    assert not [d for d in w2.data_validations.dataValidation if d.formula1 and "_CODE" in d.formula1], "prelims không có droplist"
    with zipfile.ZipFile(file_track(goc)) as z:  # openpyxl đọc được nhưng Excel từ chối mở cả file
        rong = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet") and b'<dataValidations count="0"' in z.read(n)]
    assert not rong, f"thẻ droplist rỗng ở {rong}"
    # _CODE theo đúng cột Datawork; SCOPE có cùng droplist + công thức.
    wc = wb["_CODE"]
    assert (wc["D1"].value, wc["G1"].value) == ("Mã cấp 3", "Code ID")
    assert [(wc[f"D{i}"].value, wc[f"G{i}"].value) for i in range(2, len(BANG_MAU) + 2)] == [(b["cap3"], b["ma"]) for b in BANG_MAU]
    assert wc[f"D{len(BANG_MAU) + 2}"].value == "KHÔNG RÕ" and wc[f"G{len(BANG_MAU) + 2}"].value is None
    sc = wb["SCOPE"]
    cs = cot_cua(sc, 1)
    i = next(r for r in range(2, sc.max_row + 1) if sc[f"A{r}"].value == f"{SH}!8")
    assert sc[f"{cs['Code ID']}{i}"].value.startswith("=IFERROR(INDEX(_CODE"), "SCOPE: Code ID phải nhảy theo droplist"
    assert [d for d in sc.data_validations.dataValidation if "_CODE" in (d.formula1 or "")], "SCOPE thiếu droplist"
    assert "Nhóm thực hiện" not in cs and "NTP chi tiết" not in cs and "Nhóm thực hiện" not in cot_cua(wb["MAT"], 1)
    th = [x.value for x in wb["BOQ_TONGHOP"][1]]
    assert "Bên thực hiện" in th and "Code ID" in th and "Nhóm thực hiện" not in th, th
    i_ci = th.index("Code ID") + 1
    i_k = next(r for r in range(2, 20) if wb["BOQ_TONGHOP"].cell(r, 1).value == f"{SH}!8")
    assert wb["BOQ_TONGHOP"].cell(i_k, i_ci).value == "LU.AT", "BOQ_TONGHOP chỉ xem nên Code ID là giá trị"


def khoang_cot(path, ten_sheet):
    """Các khai báo <col> thô của 1 sheet trong file xlsx: [(min, max, hidden)]."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    i = wb.sheetnames.index(ten_sheet) + 1
    wb.close()
    with zipfile.ZipFile(path) as z:
        t = z.read(f"xl/worksheets/sheet{i}.xml").decode("utf-8")
    ra = []
    for m in re.findall(r"<col [^>]*>", t):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m))
        ra.append((int(a["min"]), int(a["max"]), a.get("hidden") in ("1", "true")))
    return sorted(ra)


@kiem(3, "mau")
def gd3_sheet_boq_du_cot_nhu_tonghop(tmp):
    """Sheet BOQ có mọi cột BOQ_TONGHOP, cùng thứ tự, trừ cột đã có trên chính hàng đó. Link vật liệu trỏ file gốc, không qua sheet MAT."""
    import openpyxl
    goc, cd, data = dung_boq(tmp)
    ghi(os.path.join(data, "SCOPE.doichieu.json"), {"rows": [{"Khóa": f"{SH}!8", "Mã VL": "ST-04", "Trạng thái": "Base"}]})
    ghi(os.path.join(data, "MAT.mat.json"), {"rows": [{"Mã VL": "ST-04", "Mô tả": "Da mau", "Trạng thái cuối": "Omit", "Nguồn spec": ["Spec p3", "spec.pdf", 3]}]})
    td_ok("build", goc)
    wb = openpyxl.load_workbook(file_track(goc))
    ws = wb[SH]
    c = cot_cua(ws)
    tren_hang = {"Dòng BOQ gốc", "Item", "Mô tả BOQ", "ĐV", "KL BOQ", "Mô tả NT – Base"}
    th = [x.value for x in wb["BOQ_TONGHOP"][1] if x.value not in tren_hang]
    thieu = [x for x in th if x not in c]
    assert not thieu, f"thiếu cột {thieu}"
    assert sorted(th, key=lambda x: ws[f"{c[x]}3"].column) == th, "khác thứ tự BOQ_TONGHOP"
    o = f"{c['Link vật liệu']}8"
    assert ws[o].value == "ST-04: Da mau [Omit]", ws[o].value
    assert ws[f"{c['Nhóm']}8"].value.startswith("=IFERROR(INDEX(SCOPE!"), "Nhóm phải tra SCOPE"
    for x in ("Mã VL", "Spec · VL", "Link spec", "Mã spec/FF&E"):
        assert x not in c, f"cột {x} đã gộp vào Link vật liệu"
    dong = [r for r in wb["_LINKS"].iter_rows(min_row=2, values_only=True) if r[0] == f"{SH}!{o}"]
    assert len(dong) == 1 and "spec.pdf" in dong[0][4] and dong[0][5] == 3, f"phải trỏ file gốc đúng trang: {dong}"
    assert "MAT" not in f"{ws[o].hyperlink.target} {ws[o].hyperlink.location}", "không được trỏ sheet MAT"


@kiem(3, "mau")
def gd3_ffe_san_pham_o_ban_ve_spec_o_vat_lieu(tmp):
    """Dòng FF&E: trang ảnh của tập mới nhất ở cột bản vẽ; trang spec và trang tập FF&E banve đã ghi ở Link vật liệu. Dòng thường không đổi; dữ liệu gốc không đổi."""
    import importlib.util
    sp = importlib.util.spec_from_file_location("td_mod", TD)
    td_mod = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(td_mod)
    moi, cu, bv = "F_MOI", "F_CU", "F_BV"
    goc_bv = [["MB FF&E", bv, 1], ["FF&E p1", moi, 1], ["FF&E p60", moi, 60]]
    res = {"AREA": {"51": {}}, "MAT": {"ST-04": {"Nguồn spec": ["Spec ST-04", "F_ST", 3]}},
           "FILE_INDEX": {moi: {}, cu: {}, bv: {"Số BV": "ID-1"}},
           "VL_XUATHIEN": {"a": {"Mã VL": "51-500", "Loại trang": "Định nghĩa", "Hiệu lực": "Mới nhất", "File": ["51-500 p60", moi, 60]},
                           "b": {"Mã VL": "51-500", "Loại trang": "Định nghĩa", "Hiệu lực": "Bản cũ", "File": ["51-500 cu p5", cu, 5]},
                           "c": {"Mã VL": "51-500", "Loại trang": "Ảnh", "File": ["51-500 p62", moi, 62]},
                           "d": {"Mã VL": "51-500", "Loại trang": "Ảnh", "File": ["51-500 cu p7", cu, 7]},
                           "e": {"Mã VL": "51-999", "Loại trang": "Định nghĩa", "Hiệu lực": "Mới nhất", "File": ["51-999 p9", moi, 9]}},
           "SCOPE": {"ffe": {"Mã spec/FF&E": "51-500", "Bản vẽ hiệu lực": goc_bv},
                     "khong_anh": {"Mã spec/FF&E": "51-999"},
                     "vl": {"Mã VL": "ST-04", "Bản vẽ hiệu lực": [["FF&E p1", moi, 1]]}}}
    td_mod.gan_link_vat_lieu(res)
    f = res["SCOPE"]["ffe"]
    assert [(l[1], l[2]) for l in td_mod.ban_ve_hien(f)] == [(moi, 62), (bv, 1)], td_mod.ban_ve_hien(f)
    assert [(l[1], l[2]) for l in f["Link vật liệu"]] == [(moi, 60), (moi, 1)], f["Link vật liệu"]
    assert f["Bản vẽ hiệu lực"] is goc_bv and len(goc_bv) == 3, "không được đổi dữ liệu gốc: goi đọc số BV từ đây"
    assert not f["_khong_sp"] and res["SCOPE"]["khong_anh"]["_khong_sp"], "phải đếm dòng FF&E không có link sản phẩm"
    v = res["SCOPE"]["vl"]
    assert td_mod.ban_ve_hien(v) == [["FF&E p1", moi, 1]] and [l[1] for l in v["Link vật liệu"]] == ["F_ST"], v


@kiem(3, "mau")
def gd3_khai_bao_cot_goc_khong_lan_cot_them(tmp):
    """File gốc ẩn 1 khoảng cột sau cột cuối: không được ẩn cột tool thêm, không được chồng khai báo, phần ngoài vẫn giữ."""
    import openpyxl
    from openpyxl.worksheet.dimensions import ColumnDimension
    goc, cd, data = dung_boq(tmp)
    src = os.path.join(goc, "hoso", "boq.xlsx")
    wb = openpyxl.load_workbook(src)
    ws = wb[SH]
    ws.column_dimensions["C"] = ColumnDimension(ws, index="C", min=3, max=4, width=9)  # không lấn, phải giữ
    ws.column_dimensions["I"] = ColumnDimension(ws, index="I", min=9, max=60, width=7, hidden=True)
    wb.save(src)
    td_ok("build", goc)
    import glob
    for f, cuoi_it in ((file_track(goc), 19), (glob.glob(os.path.join(goc, "10_TENDER", "04_out", "*_CDT.xlsx"))[0], 3)):
        wb = openpyxl.load_workbook(f)
        c = cot_cua(wb[SH])
        moi = [openpyxl.utils.column_index_from_string(c[x]) for x in ("Mô tả NT – Base", "Mô tả NT – Đề xuất", "KL – Đề xuất")]
        cuoi = max(openpyxl.utils.column_index_from_string(l) for x, l in c.items() if x not in ("Item", "Description", "Quantity", "Unit"))
        assert cuoi - min(moi) + 1 >= cuoi_it, f"{f}: chỉ có {cuoi - min(moi) + 1} cột thêm"
        kb = khoang_cot(f, SH)
        for a, b in zip(kb, kb[1:]):
            assert a[1] < b[0], f"{f}: khai báo cột chồng nhau {a} {b}"
        for lo, hi, an in kb:
            if an:
                assert hi < min(moi) or lo > cuoi, f"{f}: khoảng ẩn {lo}-{hi} lấn cột thêm {min(moi)}-{cuoi}"
        assert any(lo == cuoi + 1 and hi == 60 and an for lo, hi, an in kb), f"{f}: mất phần ẩn của CĐT ngoài cột thêm: {kb}"
        assert any(lo == 3 and hi == 4 for lo, hi, _ in kb), f"{f}: mất khai báo cột của CĐT: {kb}"


@kiem(3, "mau")
def gd3_highlight_khop_dashboard(tmp):
    """Sót highlight là tender bỏ qua dòng máy không chắc; DASHBOARD lệch thì duyệt theo ngoại lệ không tin được."""
    import openpyxl
    goc, cd, data = dung_boq(tmp)
    ghi(os.path.join(data, "_phanloai_llm.json"), {"rows": [{"Khóa": f"{SH}!10", "Mã cấp 3": "Tên LLM"}, {"Khóa": f"{SH}!14", "Mã cấp 3": "KHÔNG RÕ"}]})
    td_ok("build", goc)
    cho = {r["Khóa"]: r.get("Cờ phân loại", "") for r in doc_json(goc, "SCOPE.phanloai-tool.json")["rows"]}
    can = {f"{SH}!8": "Luật chưa duyệt", f"{SH}!9": "", f"{SH}!10": "LLM", f"{SH}!11": "Chờ LLM", f"{SH}!12": "Chờ LLM", f"{SH}!14": "KHÔNG RÕ"}
    assert cho == can, cho
    wb = openpyxl.load_workbook(file_track(goc))
    ws = wb[SH]
    m = cot_cua(ws)["Mã cấp 3"]
    to = {f"{SH}!{i}" for i in range(8, 15) if co_mau(ws[f"{m}{i}"])}
    assert to == {k for k, v in can.items() if v}, f"sheet BOQ tô {sorted(to)}"
    sc = wb["SCOPE"]
    cs = cot_cua(sc, 1)
    to2 = {sc[f"A{r}"].value for r in range(2, sc.max_row + 1) if co_mau(sc[f"{cs['Mã cấp 3']}{r}"])}
    assert to2 == to, f"SCOPE tô {sorted(to2)}"
    d = wb["DASHBOARD"]
    hang = [[c.value for c in r[:3]] for r in d.iter_rows()]
    tong = next(h for h in hang if h[0] == "Code ID highlight")
    assert tong[1] == f"{len(to)} dòng", tong
    dem = {h[1]: h[2] for h in hang if h[0] == "  Cờ phân loại"}
    assert dem == {"Luật chưa duyệt": 1, "LLM": 1, "Chờ LLM": 2, "KHÔNG RÕ": 1}, dem
    assert sum(dem.values()) == len(to)
    ben = {h[1]: h[2] for h in hang if h[0] == "  Bên thực hiện"}
    assert ben["Nhà máy"] == 1 and ben["(chưa gán)"] == 7, ben  # 7 dòng + 1 prelims - 1 có bên
    assert not any(h[0] == "  Nhóm thực hiện" for h in hang)
    # Duyệt luật xong thì hết highlight ở cả 2 nơi, số đếm giảm theo.
    td_ok("phanloai", goc, "--duyet-luat")
    td_ok("build", goc)
    wb = openpyxl.load_workbook(file_track(goc))
    assert not co_mau(wb[SH][f"{m}8"]), "duyệt luật rồi vẫn tô"
    hang = [[c.value for c in r[:3]] for r in wb["DASHBOARD"].iter_rows()]
    assert next(h for h in hang if h[0] == "Code ID highlight")[1] == f"{len(to) - 1} dòng"


def sua_o(goc, sheet, hang, cot, gt, hang_tieu_de=3):
    import openpyxl
    f = file_track(goc)
    wb = openpyxl.load_workbook(f)
    ws = wb[sheet]
    ws[f"{cot_cua(ws, hang_tieu_de)[cot]}{hang}"] = gt
    wb.save(f)


@kiem(3, "mau")
def gd3_sua_droplist_giu_qua_build_lai(tmp):
    """Plan GĐ3: đổi droplist → *.user.json + nho_local.json; build lại vẫn giữ; dòng cùng mô tả nhận mã mới; Code ID không nhập tay."""
    import openpyxl
    goc, cd, data = dung_boq(tmp)
    td_ok("build", goc)
    sua_o(goc, SH, 11, "Mã cấp 3", "Tên user")
    out = td_ok("build", goc)
    assert "sửa tay user giữ 1" in out, out
    u = {r["Khóa"]: r for r in doc_json(goc, "SCOPE.user.json")["rows"]}
    assert u == {f"{SH}!11": {"Khóa": f"{SH}!11", "Mã cấp 3": "Tên user"}}, u
    with open(os.path.join(cd, "nho_local.json"), encoding="utf-8") as fh:
        loc = json.load(fh)
    assert [(x["mo_ta"], x["ma"]) for x in loc["mo_ta"]] == [("cai ke", "US.ER")], loc
    assert not os.path.exists(os.path.join(cd, "nho_khac.json"))
    kq = ket_qua(goc)
    assert kq[f"{SH}!11"]["Nguồn phân loại"] == "User" and "Cờ phân loại" not in kq[f"{SH}!11"], kq[f"{SH}!11"]
    assert kq[f"{SH}!12"]["Mã cấp 3"] == "Tên user" and kq[f"{SH}!12"]["Nguồn phân loại"] == "Nhớ-mô-tả", "dòng cùng mô tả phải nhận mã ngay"
    wb = openpyxl.load_workbook(file_track(goc))
    ws = wb[SH]
    c = cot_cua(ws)
    assert ws[f"{c['Mã cấp 3']}11"].value == "Tên user" and not co_mau(ws[f"{c['Mã cấp 3']}11"]), "user đã chọn thì hết tô"
    assert ws[f"{c['Code ID']}11"].value.startswith("=IFERROR"), "Code ID phải còn công thức"
    # Build thêm lần nữa: không phát sinh sửa tay giả, giá trị còn nguyên.
    out = td_ok("build", goc)
    assert "sửa tay user giữ 0" in out, out
    assert load_gt(goc, SH, 11, "Mã cấp 3") == "Tên user"
    # Sửa ở sheet SCOPE cũng nhận.
    sc = openpyxl.load_workbook(file_track(goc))["SCOPE"]
    i = next(r for r in range(2, sc.max_row + 1) if sc[f"A{r}"].value == f"{SH}!10")
    sua_o(goc, "SCOPE", i, "Mã cấp 3", "Tên mô tả chung", 1)
    td_ok("build", goc)
    assert load_gt(goc, SH, 10, "Mã cấp 3") == "Tên mô tả chung" and ket_qua(goc)[f"{SH}!10"]["Nguồn phân loại"] == "User"


def load_gt(goc, sheet, hang, cot, hang_tieu_de=3):
    import openpyxl
    ws = openpyxl.load_workbook(file_track(goc))[sheet]
    return ws[f"{cot_cua(ws, hang_tieu_de)[cot]}{hang}"].value


@kiem(3, "mau")
def gd3_xoa_o_va_sua_ben(tmp):
    goc, cd, data = dung_boq(tmp)
    td_ok("build", goc)
    sua_o(goc, SH, 9, "Mã cấp 3", None)  # user xóa: máy không được đoán lại
    sua_o(goc, SH, 10, "Bên thực hiện", "NTP đá")
    td_ok("build", goc)
    assert not load_gt(goc, SH, 9, "Mã cấp 3"), "ô user xóa bị máy điền lại"
    assert f"{SH}!9" not in ket_qua(goc)
    assert load_gt(goc, SH, 10, "Bên thực hiện") == "NTP đá"
    assert load_gt(goc, SH, 8, "Bên thực hiện") == "Nhà máy", "giá trị chuyển từ dữ liệu cũ phải còn"
    u = {r["Khóa"]: r for r in doc_json(goc, "SCOPE.user.json")["rows"]}
    assert u[f"{SH}!9"]["Mã cấp 3"] == "(xóa)" and u[f"{SH}!10"]["Bên thực hiện"] == "NTP đá", u
    # Cột nội bộ khác là công thức tra: gõ đè không vào user.json, build lại trả về công thức.
    sua_o(goc, SH, 8, "Trạng thái", "Omit")
    td_ok("build", goc)
    u2 = {r["Khóa"]: r for r in doc_json(goc, "SCOPE.user.json")["rows"]}
    assert f"{SH}!8" not in u2, u2
    assert str(load_gt(goc, SH, 8, "Trạng thái")).startswith("=IFERROR(INDEX(SCOPE!"), "cột tra phải giữ công thức"


@kiem(3, "mau")
def gd3_chen_hang_thi_dung_khong_nhan_sua(tmp):
    import openpyxl
    goc, cd, data = dung_boq(tmp)
    td_ok("build", goc)
    f = file_track(goc)
    wb = openpyxl.load_workbook(f)
    wb[SH].insert_rows(9)
    wb.save(f)
    rc, out = td("build", goc)
    assert rc != 0 and "bị dời" in out, f"chèn hàng mà build vẫn chạy:\n{out}"
    assert not os.path.exists(os.path.join(data, "SCOPE.user.json")), "hàng lệch mà vẫn nhận sửa tay"


@kiem(3, "mau")
def gd3_ban_cdt_khong_co_noi_bo(tmp):
    """Bản gửi CĐT là chỗ lộ thông tin: kiểm cả nội dung XML, không chỉ tên cột."""
    import glob
    import openpyxl
    goc, cd, data = dung_boq(tmp)
    ghi(os.path.join(data, "SCOPE.ben.json"), {"rows": [{"Khóa": f"{SH}!8", "Bên thực hiện": "NTP-BI-MAT"}]})
    td_ok("build", goc)
    f = glob.glob(os.path.join(goc, "10_TENDER", "04_out", "T_BOQ_*_CDT.xlsx"))[0]
    wb = openpyxl.load_workbook(f)
    assert set(wb.sheetnames) == {SH, "1. GR&Prelims"}, wb.sheetnames
    ws = wb[SH]
    c = cot_cua(ws)
    for x in COT_NOI_BO + ("Nguồn phân loại", "Cờ phân loại"):
        assert x not in c, f"bản CĐT có cột {x}"
    assert "Mô tả NT – Base" in c, "bản CĐT phải còn cột chính thức"
    assert not ws.data_validations.dataValidation, "bản CĐT có droplist"
    with zipfile.ZipFile(f) as z:
        for n in z.namelist():
            if n.endswith(".xml"):
                t = z.read(n).decode("utf-8", "replace")
                for x in ("_CODE", "NTP-BI-MAT", "Bên thực hiện", "Mã cấp 3", "Tên luật", "LU.AT"):
                    assert x not in t, f"{n} chứa '{x}'"


@kiem(3, "mau")
def gd3_ra_khong_doi_du_lieu(tmp):
    """--ra để xuất thử: không kéo sửa tay, không chạy phanloai, không ghi nho_local."""
    goc, cd, data = dung_boq(tmp)
    td_ok("build", goc)
    sua_o(goc, SH, 11, "Mã cấp 3", "Tên user")
    truoc, truoc_cd = chu_ky_thu_muc(data), chu_ky_thu_muc(cd)
    td_ok("build", goc, "--ra", os.path.join(tmp, "ra"))
    assert chu_ky_thu_muc(data) == truoc, "--ra đã ghi _data"
    assert chu_ky_thu_muc(cd) == truoc_cd, "--ra đã ghi bộ nhớ"


@kiem(3, "mau")
def gd3_cot_moi_khong_doi_ngay_lan_dau(tmp):
    """Snapshot cũ chưa có cột mới: không đổi ngày. Đổi bên, đổi phân loại: không đổi ngày; đổi nội dung thì có."""
    goc, cd, data = dung_boq(tmp)
    td_ok("build", goc)
    duong = os.path.join(data, "_snapshot.json")
    snap = doc_json(goc, "_snapshot.json")
    snap.pop("boq_track")  # không so với file đã build, chỉ thử derive
    snap.pop("file_index")
    for r in snap["reg"]["SCOPE"].values():
        r["Ngày cập nhật"] = "2000-01-01"
        r.pop("Bên thực hiện")
    ghi(duong, snap)
    td_ok("build", goc)
    s2 = doc_json(goc, "_snapshot.json")
    assert {r["Ngày cập nhật"] for r in s2["reg"]["SCOPE"].values()} == {"2000-01-01"}, "cột mới làm đổi Ngày cập nhật"
    assert s2["reg"]["SCOPE"][f"{SH}!8"]["Bên thực hiện"] == "Nhà máy"
    s2.pop("boq_track")
    s2.pop("file_index")
    ghi(duong, s2)
    ghi(os.path.join(data, "SCOPE.ben.json"), {"rows": [{"Khóa": f"{SH}!8", "Bên thực hiện": "Owner"}]})
    ghi(os.path.join(data, "_phanloai_llm.json"), {"rows": [{"Khóa": f"{SH}!10", "Mã cấp 3": "Tên LLM"}]})
    ghi(os.path.join(data, "SCOPE.doichieu-1.json"), {"rows": [{"Khóa": f"{SH}!11", "Trạng thái": "Omit"}]})
    td_ok("build", goc)
    s3 = doc_json(goc, "_snapshot.json")["reg"]["SCOPE"]
    assert s3[f"{SH}!8"]["Ngày cập nhật"] == "2000-01-01", "đổi bên không phải thay đổi nội dung BOQ (danh sách bên sẽ còn đổi)"
    assert s3[f"{SH}!10"]["Ngày cập nhật"] == "2000-01-01", "đổi phân loại không phải thay đổi nội dung"
    assert s3[f"{SH}!11"]["Ngày cập nhật"] != "2000-01-01", "đổi nội dung mà không đổi ngày: phép kiểm vô nghĩa"


def mo_excel():
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return None
    pythoncom.CoInitialize()
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
    except Exception:
        return None
    xl.Visible, xl.DisplayAlerts = False, False
    return xl


@kiem(3, "mau")
def gd3_excel_doi_droplist_code_id_nhay(tmp):
    """Openpyxl không tính công thức: chỉ Excel mới chứng minh được Code ID nhảy và droplist liên sheet được chấp nhận."""
    goc, cd, data = dung_boq(tmp)
    td_ok("build", goc)
    xl = mo_excel()
    if not xl:
        print("      BỎ QUA: máy không có Excel COM")
        return
    f = file_track(goc)
    try:
        wb = xl.Workbooks.Open(f)
        ws = wb.Worksheets(SH)
        hdr = {ws.Cells(3, j).Value: j for j in range(1, 60) if ws.Cells(3, j).Value}
        m, ci = hdr["Mã cấp 3"], hdr["Code ID"]
        assert ws.Cells(8, ci).Value == "LU.AT", ws.Cells(8, ci).Value
        assert ws.Cells(8, m).Validation.Type == 3, "ô không có droplist"
        ws.Cells(8, m).Value = "Tên user"
        assert ws.Cells(8, ci).Value == "US.ER", f"Code ID không nhảy: {ws.Cells(8, ci).Value}"
        ws.Cells(10, m).Value = "Mã có cách"
        assert ws.Cells(10, ci).Value == "LS. HA", "phải giữ đúng cột G"
        ws.Cells(11, m).Value = "KHÔNG RÕ"
        assert (ws.Cells(11, ci).Value or "") == "", f"KHÔNG RÕ phải ra trống, đang là {ws.Cells(11, ci).Value!r}"
        wb.Saved = False
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    out = td_ok("build", goc)
    assert "sửa tay user giữ 3" in out, out
    assert load_gt(goc, SH, 8, "Mã cấp 3") == "Tên user"
    assert ket_qua(goc)[f"{SH}!8"]["Nguồn phân loại"] == "User"


def excel_mo_khong_loi(xl, f):
    """CorruptLoad=1 + đọc báo cáo: bản lành chỉ có dòng chung, lỗi thật có repairedRecord/removedPart (memory com-build-gotchas)."""
    import glob
    d = os.environ.get("TEMP", "")
    cu = set(glob.glob(os.path.join(d, "error*.xml")))
    wb = xl.Workbooks.Open(f, 0, True, None, None, None, True, None, None, None, None, None, None, None, 1)
    wb.Close(False)
    for x in set(glob.glob(os.path.join(d, "error*.xml"))) - cu:
        with open(x, encoding="utf-8", errors="replace") as fh:
            t = fh.read()
        assert "<repairedRecord" not in t and "<removedPart" not in t, f"{f}: {t[:600]}"


def ban_sao_gd3(tmp, nho=False):
    goc = ban_sao_du_lieu(tmp)
    cd = os.path.join(tmp, "code_id")
    os.makedirs(cd)
    for t in ("bang_ma.json", "luat.json") + (("nho.json",) if nho else ()):
        shutil.copy2(os.path.join(ROOT, "_MASTER", "code_id", t), cd)
    os.environ["TD_CODE_ID_DIR"] = cd
    return goc, cd


@kiem(3, "sofitel")
def gd3_sofitel_doi_chieu_gia_tri_cu(tmp):
    """Đối chiếu từng khóa: bản sao lưu dạng cũ, tính lại theo quy tắc độc lập, so với SCOPE.ben.json thật."""
    import glob
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    bk = sorted(glob.glob(os.path.join(data_that, "_backup", "truoc-gd3-*")))
    assert bk, "SOFITEL chưa chuyển: không có _backup\\truoc-gd3-*"
    truoc = chu_ky_thu_muc(data_that)
    goc, _ = ban_sao_gd3(tmp)
    data = os.path.join(goc, "10_TENDER", "_data")
    os.remove(os.path.join(data, "SCOPE.ben.json"))
    shutil.rmtree(os.path.join(data, "_backup"))
    n_cu = 0
    for f in glob.glob(os.path.join(bk[-1], "*.json")):
        shutil.copy2(f, data)
        with open(f, encoding="utf-8") as fh:
            n_cu += sum(1 for r in json.load(fh)["rows"] if "Nhóm thực hiện" in r or "NTP chi tiết" in r)
    out = td_ok("loc", goc, "SCOPE", "--cot", "Khóa,Nhóm thực hiện,NTP chi tiết", "--max", "9999")
    cu = {json.loads(x)["Khóa"]: json.loads(x) for x in out.splitlines()[1:]}
    ghi_ro = {r["Khóa"] for f in glob.glob(os.path.join(bk[-1], "SCOPE.*.json")) for r in json.load(open(f, encoding="utf-8"))["rows"]
              if "Khóa" in r and "NTP chi tiết" in r}
    with open(os.path.join(data_that, "SCOPE.ben.json"), encoding="utf-8") as fh:
        ben = {r["Khóa"]: r["Bên thực hiện"] for r in json.load(fh)["rows"]}
    sai = []
    for k, r in cu.items():
        n, c = r.get("Nhóm thực hiện"), r.get("NTP chi tiết")
        if n == "NTP":
            can = c or "NTP"
        elif n == "Chung" or not n:
            can = None
        else:
            can = f"{n}; {c}" if c and k in ghi_ro else n
        if ben.get(k) != can:
            sai.append((k, n, c, ben.get(k)))
    assert not sai, f"{len(sai)} dòng sai: {sai[:5]}"
    assert set(ben) <= set(cu)
    co = [k for k, r in cu.items() if r.get("Nhóm thực hiện") not in (None, "Chung")]
    assert set(co) == set(ben), f"khóa có nhóm khác khóa có bên: {sorted(set(co) ^ set(ben))[:5]}"
    assert ben.get("2. Finishing-L4-17!81") == "Owner", "plan mục 6: 3 dòng IR-01 giữ Owner"
    # Chạy lại công cụ trên bản cũ phải ra đúng file thật.
    rc, o = chuyen(goc)
    assert rc == 0, o
    assert {r["Khóa"]: r["Bên thực hiện"] for r in doc_json(goc, "SCOPE.ben.json")["rows"]} == ben, "chuyen_ben không tái lập được"
    assert chu_ky_thu_muc(data_that) == truoc, "_data thật đã bị ghi"
    print(f"      SOFITEL: {n_cu} dòng JSON có cột cũ, {len(cu)} khóa đối chiếu từng giá trị, {len(ben)} có Bên, 0 sai")


@kiem(3, "sofitel")
def gd3_sofitel_build_that_tren_ban_sao(tmp):
    """Plan GĐ3: check 0 LỖI; lần build đầu không đổi Ngày cập nhật do cột mới; bản CĐT sạch; file mở trong Excel không báo lỗi."""
    import glob
    import openpyxl
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    truoc, truoc_master = chu_ky_thu_muc(data_that), chu_ky_master()
    goc, cd = ban_sao_gd3(tmp)
    rc, out = td("check", goc)
    assert rc == 0, f"check sau chuyển phải 0 LỖI:\n{out[-1500:]}"
    td_ok("phanloai", goc)
    snap_cu = doc_json(goc, "_snapshot.json")
    snap_cu.pop("boq_track", None)  # không kéo sửa tay từ file thật; lệch hàng ở file thật là chuyện riêng của build thật
    snap_cu.pop("file_index", None)
    ghi(os.path.join(goc, "10_TENDER", "_data", "_snapshot.json"), snap_cu)
    out = td_ok("build", goc)
    assert "0 LỖI" in out, out[-800:]
    moi = doc_json(goc, "_snapshot.json")["reg"]["SCOPE"]
    cu = snap_cu["reg"]["SCOPE"]
    doi = [k for k in moi if moi[k]["Ngày cập nhật"] != cu[k]["Ngày cập nhật"]]
    assert not doi, f"{len(doi)} dòng đổi ngày: {doi[:5]}"
    f = glob.glob(os.path.join(goc, "10_TENDER", "03_boq", "SSP_BOQ_TRACK_*.xlsx"))[0]
    wb = openpyxl.load_workbook(f)
    assert "_CODE" in wb.sheetnames
    dv = to = 0
    for sh in ("2. Finishing-L4-17", "2.1 Finishing-L18-20", "3. Furniture-L4-17", "3.1 Furniture-suite-L18-20"):
        ws = wb[sh]
        for d in ws.data_validations.dataValidation:
            if "_CODE!$D" in (d.formula1 or ""):
                dv += sum((r.max_row - r.min_row + 1) * (r.max_col - r.min_col + 1) for r in d.sqref.ranges)
        h = cot_cua(ws, next(i for i in range(1, 25) if any(c.value == "Mã cấp 3" for c in ws[i])))
        to += sum(1 for row in ws.iter_rows(min_col=ws[f"{h['Mã cấp 3']}1"].column, max_col=ws[f"{h['Mã cấp 3']}1"].column) if co_mau(row[0]) and row[0].value != "Mã cấp 3" and row[0].row > 5)
    assert dv == 659, f"droplist {dv} ô, cần 659"
    kq = ket_qua(goc)
    n_co = sum(1 for r in kq.values() if r.get("Cờ phân loại"))
    n_ngoai = sum(1 for k, r in moi.items() if r.get("Cờ phân loại"))
    assert n_ngoai == n_co
    assert sum(1 for r in moi.values() if r.get("Bên thực hiện")) == 659, "659 dòng Bên thực hiện phải lên sheet"
    dash = [[c.value for c in r[:3]] for r in wb["DASHBOARD"].iter_rows()]
    assert next(h for h in dash if h[0] == "Code ID highlight")[1] == f"{n_co} dòng"
    fc = glob.glob(os.path.join(goc, "10_TENDER", "04_out", "SSP_BOQ_*_CDT.xlsx"))[0]
    wc = openpyxl.load_workbook(fc, read_only=True)
    assert not {"_CODE", "SCOPE", "DASHBOARD", "_LINKS"} & set(wc.sheetnames), wc.sheetnames
    xl = mo_excel()
    if xl:
        try:
            excel_mo_khong_loi(xl, f)
            excel_mo_khong_loi(xl, fc)
        finally:
            xl.Quit()
    else:
        print("      BỎ QUA kiểm mở Excel: máy không có Excel COM")
    assert chu_ky_thu_muc(data_that) == truoc, "_data thật đã bị ghi"
    assert chu_ky_master() == truoc_master, "_MASTER thật đã bị ghi"
    print(f"      SOFITEL: {len(moi)} dòng, droplist {dv} ô, tô {to} ô, highlight {n_co}, 0 dòng đổi Ngày cập nhật, check 0 LỖI")


# ---------------------------------------------------------------- GĐ3b
BEN_MAU = {"ds": ["Nhà máy", "NTP đá", "NTP sắt", "NCC giấy", "Owner"], "gop": {"Đá cũ": "NTP đá"},
           "loai": {"Đá": "NTP đá", "Gỗ/Veneer": "Nhà máy"}, "tam": True}


def dung_ben(tmp):
    """Dòng 8, 9: 2 dòng BOQ dùng chung mẫu WC-01 3 thành phần; 10: 1 bên; 11: chữ cũ cần gộp; 12: Loại trừ; 14: chưa có bên."""
    goc, cd, data = dung_boq(tmp)
    ghi(os.path.join(data, "_ben.json"), BEN_MAU)
    ghi(os.path.join(data, "SCOPE.ben.json"), {"rows": [{"Khóa": f"{SH}!10", "Bên thực hiện": "Nhà máy"}, {"Khóa": f"{SH}!11", "Bên thực hiện": "Đá cũ"},
                                                        {"Khóa": f"{SH}!12", "Bên thực hiện": "NTP đá"}]})
    ghi(os.path.join(data, "SCOPE.doichieu-1.json"), {"rows": [{"Khóa": f"{SH}!{i}", "Trạng thái": "Base"} for i in (8, 9, 10, 11, 14)]
                                                      + [{"Khóa": f"{SH}!12", "Trạng thái": "Loại trừ"}]})
    td_ok("build", goc)
    return goc, cd, data


def them_mau(goc):
    import openpyxl
    f = file_track(goc)
    wb = openpyxl.load_workbook(f)
    ws = wb["CAU_TAO"]
    c = cot_cua(ws, 1)
    for i, (k, tp, vl, b) in enumerate([("WC-01#1", "Giấy dán tường", "WC-01", "NCC giấy"), ("WC-01#2", "Khung gỗ", None, "Nhà máy"),
                                        ("WC-01#3", "Khung sắt", "MT-02", "NTP sắt")], 2):
        ws[f"A{i}"], ws[f"{c['Thành phần']}{i}"], ws[f"{c['Mã VL']}{i}"], ws[f"{c['Bên thực hiện']}{i}"] = k, tp, vl, b
    bq = wb[SH]
    cb = cot_cua(bq)
    bq[f"{cb['Mã cấu tạo']}8"] = "WC-01"
    bq[f"{cb['Mã cấu tạo']}9"] = "WC-01"
    wb.save(f)


@kiem("3b", "mau")
def gd3b_sua_mau_mot_lan_ca_hai_dong_doi(tmp):
    import openpyxl
    goc, cd, data = dung_ben(tmp)
    them_mau(goc)
    td_ok("build", goc)
    can = "NCC giấy; Nhà máy; NTP sắt"
    assert load_gt(goc, SH, 8, "Bên thực hiện") == can and load_gt(goc, SH, 9, "Bên thực hiện") == can, load_gt(goc, SH, 8, "Bên thực hiện")
    assert load_gt(goc, SH, 11, "Bên thực hiện") == "NTP đá", "chữ cũ phải gộp theo _ben.json"
    with open(os.path.join(data, "SCOPE.ben.json"), encoding="utf-8") as fh:
        assert "Đá cũ" in fh.read(), "không được ghi đè chữ gốc"
    # Sửa mẫu 1 lần: cả 2 dòng đổi; build lại vẫn giữ.
    f = file_track(goc)
    wb = openpyxl.load_workbook(f)
    ws = wb["CAU_TAO"]
    c = cot_cua(ws, 1)
    i = next(r for r in range(2, 9) if ws[f"A{r}"].value == "WC-01#3")
    ws[f"{c['Bên thực hiện']}{i}"] = "NTP đá"
    wb.save(f)
    td_ok("build", goc)
    td_ok("build", goc)
    can = "NCC giấy; Nhà máy; NTP đá"
    assert load_gt(goc, SH, 8, "Bên thực hiện") == can and load_gt(goc, SH, 9, "Bên thực hiện") == can
    wb = openpyxl.load_workbook(file_track(goc))
    ws = wb["CAU_TAO"]
    assert ws["B2"].value == "WC-01" and ws["A2"].value == "WC-01#1", "cột Mã cấu tạo phải tự có"
    # Sửa Bên ở dòng có mẫu: không tác dụng và bị báo.
    sua_o(goc, SH, 8, "Bên thực hiện", "Owner")
    out = td_ok("build", goc)
    assert load_gt(goc, SH, 8, "Bên thực hiện") == can, "sửa ở BOQ không được thắng mẫu"
    _, ck = td("check", goc)
    assert "không có tác dụng" in ck, ck
    # Xóa dòng mẫu trong xlsx thì mất thành phần đó.
    f = file_track(goc)
    wb = openpyxl.load_workbook(f)
    ws = wb["CAU_TAO"]
    i = next(r for r in range(2, 9) if ws[f"A{r}"].value == "WC-01#2")
    ws.delete_rows(i)
    wb.save(f)
    td_ok("build", goc)
    assert load_gt(goc, SH, 9, "Bên thực hiện") == "NCC giấy; NTP đá", load_gt(goc, SH, 9, "Bên thực hiện")


@kiem("3b", "mau")
def gd3b_phan_viec_va_dashboard(tmp):
    import openpyxl
    goc, cd, data = dung_ben(tmp)
    them_mau(goc)
    td_ok("build", goc)
    wb = openpyxl.load_workbook(file_track(goc))
    assert wb.sheetnames[:3] == ["DASHBOARD", "BOQ_TONGHOP", "PHAN_VIEC"], wb.sheetnames
    pv = [r for r in wb["PHAN_VIEC"].iter_rows(min_row=2, values_only=True)]
    d8 = [(r[0], r[4]) for r in pv if r[1] == f"{SH}!8"]
    assert d8 == [("NCC giấy", "Giấy dán tường"), ("Nhà máy", "Khung gỗ"), ("NTP sắt", "Khung sắt")], d8
    assert [(r[0], r[4]) for r in pv if r[1] == f"{SH}!10"] == [("Nhà máy", "Cả dòng")]
    assert [r[0] for r in pv if r[1] == f"{SH}!14"] == [None], "dòng chưa có bên phải hiện để lọc"
    assert "PHAN_VIEC" not in doc_json(goc, "_snapshot.json")["reg"], "PHAN_VIEC không được vào snapshot"
    hang = [[c.value for c in r[:3]] for r in wb["DASHBOARD"].iter_rows()]
    assert next(h for h in hang if h[0] == "Dòng Base chưa có bên")[1] == 1, "đếm sai dòng Base chưa có bên"
    assert next(h for h in hang if h[0] == "  Danh sách bên")[1] == "TẠM"
    ben = {h[1]: h[2] for h in hang if h[0] == "  Bên thực hiện"}
    assert ben["NCC giấy; Nhà máy; NTP sắt"] == 2, ben
    # Droplist bên, mã cấu tạo; _CODE có cả 2 danh sách.
    bq = wb[SH]
    c = cot_cua(bq)
    dv = {d.formula1: {str(x) for x in d.sqref.ranges} for d in bq.data_validations.dataValidation}
    assert f"{c['Bên thực hiện']}10" in dv[f"=_CODE!$L$2:$L${len(BEN_MAU['ds']) + 1}"], dv
    assert f"{c['Mã cấu tạo']}10" in dv["=_CODE!$N$2:$N$500"], dv
    assert not bq.column_dimensions[c["Mã cấu tạo"]].hidden
    wc = wb["_CODE"]
    assert wc["L2"].value == "Nhà máy" and wc["N2"].value == "WC-01"
    # Bản CĐT không có CAU_TAO, PHAN_VIEC.
    import glob
    fc = glob.glob(os.path.join(goc, "10_TENDER", "04_out", "T_BOQ_*_CDT.xlsx"))[0]
    wc = openpyxl.load_workbook(fc)
    assert not {"CAU_TAO", "PHAN_VIEC"} & set(wc.sheetnames) and "Mã cấu tạo" not in cot_cua(wc[SH])


@kiem("3b", "mau")
def gd3b_co_nhieu_ben(tmp):
    import code_id  # noqa: F401
    goc, cd, data = dung_ben(tmp)
    ghi(os.path.join(data, "MAT.mat.json"), {"rows": [{"Mã VL": "ST-04", "Loại": "Đá"}, {"Mã VL": "WD-01", "Loại": "Gỗ/Veneer"}, {"Mã VL": "WD-02", "Loại": "Gỗ/Veneer"}]})
    dong = doc_json(goc, "SCOPE.boq-tool.json")["rows"]
    theo = {r["Khóa"]: r for r in dong}
    theo[f"{SH}!8"]["Mã VL"] = "ST-04; WD-01"  # 2 mã VL khác bên
    theo[f"{SH}!10"]["Mã VL"] = "WD-01; WD-02"  # 2 mã VL cùng bên: không cờ
    theo[f"{SH}!11"]["Mô tả BOQ"] = "Painting to plywood/MDF, including framework"
    theo[f"{SH}!12"]["Mô tả BOQ"] = "Metal frame, consisting of"  # Loại trừ: không cờ
    ghi(os.path.join(data, "SCOPE.boq-tool.json"), {"rows": dong})
    td_ok("build", goc)
    snap = doc_json(goc, "_snapshot.json")["reg"]["SCOPE"]
    co = {k for k, r in snap.items() if r.get("Cờ bên")}
    assert co == {f"{SH}!8", f"{SH}!11"}, co
    # Đã chọn mẫu thì hết cờ.
    sua_o(goc, SH, 8, "Mã cấu tạo", "CT-09")
    td_ok("build", goc)
    assert not doc_json(goc, "_snapshot.json")["reg"]["SCOPE"][f"{SH}!8"].get("Cờ bên")
    _, ck = td("check", goc)
    assert "'CT-09' chưa có ở sheet CAU_TAO" in ck, ck


@kiem("3b", "mau")
def gd3b_nhap_ben_tu_mat(tmp):
    """Dự án mới: nháp từ Mã VL → Loại → bên, luôn gắn cờ; không có Mã VL thì trống; không đụng ô đã có."""
    goc, cd, data = dung_ben(tmp)
    os.remove(os.path.join(data, "SCOPE.ben.json"))
    ghi(os.path.join(data, "MAT.mat.json"), {"rows": [{"Mã VL": "ST-04", "Loại": "Đá"}]})
    dong = doc_json(goc, "SCOPE.boq-tool.json")["rows"]
    theo = {r["Khóa"]: r for r in dong}
    theo[f"{SH}!8"]["Mã VL"] = "ST-04"
    theo[f"{SH}!9"]["Mã VL"] = "ST-04"
    ghi(os.path.join(data, "SCOPE.boq-tool.json"), {"rows": dong})
    ghi(os.path.join(data, "SCOPE.user.json"), {"rows": [{"Khóa": f"{SH}!9", "Bên thực hiện": "Owner"}]})
    td_ok("ben", goc)
    kq = {r["Khóa"]: r for r in doc_json(goc, "SCOPE.ben-tool.json")["rows"]}
    assert list(kq) == [f"{SH}!8"] and kq[f"{SH}!8"] == {"Khóa": f"{SH}!8", "Bên thực hiện": "NTP đá", "Cờ bên": "Nháp"}, kq
    td_ok("build", goc)
    assert doc_json(goc, "_snapshot.json")["reg"]["SCOPE"][f"{SH}!8"]["Cờ bên"] == "Nháp"
    # User sửa đè nháp: build lại không ghi đè, hết cờ nháp.
    sua_o(goc, SH, 8, "Bên thực hiện", "Nhà máy")
    td_ok("build", goc)
    s8 = doc_json(goc, "_snapshot.json")["reg"]["SCOPE"][f"{SH}!8"]
    assert s8["Bên thực hiện"] == "Nhà máy" and not s8.get("Cờ bên"), s8


def doc_goi(goc, ben):
    import glob
    import openpyxl
    ten = re.sub(r"[^\w. -]+", "_", ben).strip()
    f = glob.glob(os.path.join(goc, "10_TENDER", "04_out", f"T_GOI_{ten}_*.xlsx"))[0]
    wb = openpyxl.load_workbook(f)
    assert wb.sheetnames == ["GOI"], wb.sheetnames
    rows = list(wb["GOI"].iter_rows(values_only=True))
    i = next(j for j, r in enumerate(rows) if r[0] == "Dòng BOQ")
    return f, rows[i], rows[i + 1:]


@kiem("3b", "mau")
def gd3b_goi_chi_co_phan_cua_ben(tmp):
    """Sai ở đây là lộ thông tin ra nhà thầu phụ: kiểm cả XML trong zip, không chỉ ô nhìn thấy."""
    goc, cd, data = dung_ben(tmp)
    them_mau(goc)
    # Mô tả có mã TCQ, dòng có giá trị nội bộ để thử lộ.
    ghi(os.path.join(data, "SCOPE.doichieu-2.json"), {"rows": [{"Khóa": f"{SH}!10", "Mô tả NT – Base": "Cai tu moi (standard, TCQ02-9); LED (TCQ04-10)",
                                                                "Nội dung cuối (VN)": "NOIBO-BIMAT", "Việc còn mở": "NOIBO-BIMAT", "Mã TCQ": "TCQ02-9"}]})
    td_ok("build", goc)
    ky_vong = {"NCC giấy": ("Giấy dán tường", 2), "NTP sắt": ("Khung sắt", 2), "Nhà máy": ("Khung gỗ", 3)}
    tat_ca = {b: td_ok("goi", goc, b) for b in ky_vong}
    for b, (pv, n) in ky_vong.items():
        f, hdr, rows = doc_goi(goc, b)
        assert len(rows) == n, f"{b}: {len(rows)} dòng, cần {n}\n{rows}"
        assert list(hdr) == ["Dòng BOQ", "Item BOQ", "Mô tả", "Phần việc của bên", "Mã VL", "Bản vẽ", "ĐV", "KL BOQ (tham khảo)", "Ghi chú", "ĐV NTP", "KL NTP", "Đơn giá"], hdr
        assert pv in {r[3] for r in rows}, rows
        for r in rows:
            assert r[9] is None and r[10] is None and r[11] is None, "cột của bên nhận phải trống"
        with zipfile.ZipFile(f) as z:
            for x in z.namelist():
                t = z.read(x).decode("utf-8", "replace")
                for cam in ("NOIBO-BIMAT", "TCQ", "LU.AT", "Tên luật", "Mã cấp 3", "Code ID", "Trạng thái", "_LINKS", "hyperlink", "10_TENDER"):
                    assert cam not in t, f"file {b}, {x} có '{cam}'"
                for khac in set(ky_vong) - {b}:
                    assert khac not in t, f"file {b} có tên bên khác '{khac}'"
    _, _, r = doc_goi(goc, "Nhà máy")
    d10 = next(x for x in r if x[0] == f"{SH}!10")
    assert d10[2] == "Cai tu moi (standard); LED" and d10[3] == "Cả dòng", d10
    # Loại trừ không vào; bên không có dòng thì báo.
    td_ok("goi", goc, "NTP đá")
    _, _, r = doc_goi(goc, "NTP đá")
    assert {x[0] for x in r} == {f"{SH}!11"}, "dòng Loại trừ không được gửi"
    rc, out = td("goi", goc, "Bên lạ")
    assert rc != 0 and "Nhà máy" in out, out
    # Máy soát phải bắt thật: đối chứng bằng cách nhét Code ID vào Mã VL.
    ghi(os.path.join(data, "SCOPE.doichieu-3.json"), {"rows": [{"Khóa": f"{SH}!14", "Mã VL": "LU.AT"}]})
    ghi(os.path.join(data, "SCOPE.ben.json"), {"rows": [{"Khóa": f"{SH}!14", "Bên thực hiện": "Owner"}]})
    rc, out = td("goi", goc, "Owner")
    assert rc != 0 and "Code ID" in out, f"máy soát không bắt:\n{out}"
    assert not [x for x in os.listdir(os.path.join(goc, "10_TENDER", "04_out")) if "GOI_Owner" in x], "file lỗi phải bị xóa"


@kiem("3b", "mau")
def gd3b_goi_ve_cho_va_nhieu_ben_khong_mau(tmp):
    goc, cd, data = dung_ben(tmp)
    ghi(os.path.join(data, "SCOPE.doichieu-1.json"), {"rows": [{"Khóa": f"{SH}!8", "Trạng thái": "VE"}, {"Khóa": f"{SH}!9", "Trạng thái": "Chờ"},
                                                               {"Khóa": f"{SH}!10", "Trạng thái": "Omit"}, {"Khóa": f"{SH}!11", "Trạng thái": "Base"}]})
    ghi(os.path.join(data, "SCOPE.ben.json"), {"rows": [{"Khóa": f"{SH}!{i}", "Bên thực hiện": "Nhà máy; NTP đá"} for i in (8, 9, 10, 11)]})
    td_ok("build", goc)
    td_ok("goi", goc, "NTP đá")
    _, _, r = doc_goi(goc, "NTP đá")
    assert {x[0]: x[8] for x in r} == {f"{SH}!8": "VE", f"{SH}!9": "Chờ xác nhận", f"{SH}!11": None}, r
    assert all(x[3] == "Một phần của dòng, xem mô tả" for x in r), r


@kiem("3b", "mau")
def gd3b_excel_mo_duoc_va_droplist_ben(tmp):
    goc, cd, data = dung_ben(tmp)
    them_mau(goc)
    td_ok("build", goc)
    td_ok("goi", goc, "Nhà máy")
    xl = mo_excel()
    if not xl:
        print("      BỎ QUA: máy không có Excel COM")
        return
    try:
        excel_mo_khong_loi(xl, file_track(goc))
        excel_mo_khong_loi(xl, doc_goi(goc, "Nhà máy")[0])
        wb = xl.Workbooks.Open(file_track(goc))
        ws = wb.Worksheets(SH)
        hdr = {ws.Cells(3, j).Value: j for j in range(1, 60) if ws.Cells(3, j).Value}
        assert ws.Cells(10, hdr["Bên thực hiện"]).Validation.Type == 3
        assert ws.Cells(10, hdr["Mã cấu tạo"]).Validation.Type == 3
        wb.Close(False)
    finally:
        xl.Quit()


@kiem("3b", "sofitel")
def gd3b_sofitel_danh_sach_tam_va_goi(tmp):
    """Danh sách tạm gộp hết chữ cũ; goi trên dữ liệu thật qua máy soát; chữ gốc không đổi."""
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    truoc, truoc_master = chu_ky_thu_muc(data_that), chu_ky_master()
    goc, cd = ban_sao_gd3(tmp, nho=True)
    snap = doc_json(goc, "_snapshot.json")
    snap.pop("boq_track", None)
    snap.pop("file_index", None)
    ghi(os.path.join(goc, "10_TENDER", "_data", "_snapshot.json"), snap)
    out = td_ok("build", goc)
    assert "0 LỖI, 1 XEM" in out, out[-600:]  # 1 XEM duy nhất: nhắc chưa đọc sâu (GĐ7)
    assert "XEM  Chưa đọc sâu" in td("check", goc)[1]
    moi = doc_json(goc, "_snapshot.json")["reg"]["SCOPE"]
    ds = set(doc_json(goc, "_ben.json")["ds"])
    la = {b for r in moi.values() for b in str(r.get("Bên thực hiện") or "").split("; ") if b and b not in ds}
    assert not la, f"bên ngoài danh sách: {la}"
    so_ben = sum(1 for r in moi.values() if r.get("Bên thực hiện"))
    assert so_ben == 659, so_ben
    tong = 0
    co = [b for b in sorted(ds) if not td("goi", goc, b)[0]]
    assert "NTP đá" in co and "Nhà máy" in co, co
    for b in co:
        _, _, r = doc_goi_du_an(goc, b)
        tong += len(r)
    n_co = sum(1 for r in moi.values() if "Nhiều bên?" in str(r.get("Cờ bên") or ""))
    assert chu_ky_thu_muc(data_that) == truoc, "_data thật đã bị ghi"
    assert chu_ky_master() == truoc_master, "_MASTER thật đã bị ghi"
    print(f"      SOFITEL: {len(ds)} bên tạm, {so_ben} dòng có bên, {len(co)} file goi = {tong} dòng, {n_co} dòng cờ nhiều bên; máy soát đạt cả {len(co)} file")


def doc_goi_du_an(goc, ben):
    import glob
    import openpyxl
    ten = re.sub(r"[^\w. -]+", "_", ben).strip()
    f = glob.glob(os.path.join(goc, "10_TENDER", "04_out", f"*_GOI_{ten}_*.xlsx"))[0]
    rows = list(openpyxl.load_workbook(f)["GOI"].iter_rows(values_only=True))
    i = next(j for j, r in enumerate(rows) if r[0] == "Dòng BOQ")
    return f, rows[i], rows[i + 1:]


# ---------------------------------------------------------------- GĐ4
TRANG_FFE_51 = "SPECIFICATION\nArea:\nStandard King\nItem:\n51 - 500B\nDate:\nFabric. Ghi chu kem PT-01 va 51 - 501\nLondon Postcode: SW10 0XE\nTel 44 - 207"


def dung_xh(tmp):
    """Dự án giả có đủ: FF&E AREA 51, IDD Area 51 (cùng số, khác khu), Schedule, thảm, TCQ, MAT."""
    goc, cd, data = dung_boq(tmp)
    inp = os.path.join(goc, "hoso")
    for rel in ("FFE/AREA 51 cu.pdf", "FFE/AREA 51 moi.pdf", "Sch/Finishes.pdf", "IDD/Area 51 - Lift Lobby/ID-LL-51-101.pdf", "TCQ/tcq.pdf"):
        tao_file(inp, rel)
    td_ok("kiemke", goc)
    fid = {r["Đường dẫn"]: r["ID"] for r in doc_index(goc)}
    f_cu = ghi_cache(goc, "FFE/AREA 51 cu.pdf", [TRANG_FFE_51, "MATERIAL REFERENCE\nITEM\n51 - 500B\nDesign No: X", "Image (s) provided for design intent\n51 - 500B"])
    f_moi = ghi_cache(goc, "FFE/AREA 51 moi.pdf", ["SPECIFICATION\nItem:\n51 - 500B\nRev moi", "ITEM SCHEDULE\n51 - 500B\n51 - 600\n52 - 500.1"])
    f_sch = ghi_cache(goc, "Sch/Finishes.pdf", ["DRAWING REFERENCE: ST-01\nMarble", "DRAWING REFERENCE:\nBasin\nDURAVIT\nSW-05 ghi cuoi trang", "Room schedule Standard Twin ST-01 90 keys"])
    f_idd = ghi_cache(goc, "IDD/Area 51 - Lift Lobby/ID-LL-51-101.pdf", ["ban ve khong co ma"])
    ghi_cache(goc, "TCQ/tcq.pdf", ["51 - 500B trong ban PDF TCQ"])
    idx = doc_json(goc, "FILE_INDEX.kiemke-tool.json")
    for r in idx["rows"]:
        r["Ngày issue"] = {f_cu: "2026-06-05", f_moi: "2026-10-19", f_sch: "2026-06-23"}.get(r["ID"], r.get("Ngày issue", "2026-01-01"))
        if r["ID"] == fid["TCQ/tcq.pdf"]:
            r["Vai trò"] = "Làm rõ"
        if r["ID"] == f_idd:
            r["Số BV"] = "ID-LL-51-101"
    ghi(os.path.join(data, "FILE_INDEX.kiemke-tool.json"), idx)
    ghi(os.path.join(data, "DWG.dwg-tool.json"), {"rows": [{"Số BV": "ID-LL-51-101", "Khu (thư mục)": "Area 51 - Lift Lobby", "File PDF": [["ID-LL-51-101 PDF", f_idd]], "Trạng thái": "Hiệu lực"}]})
    ghi(os.path.join(data, "KHU.dwg.json"), {"rows": [
        {"Khu": "Area 51 - Lift Lobby", "Bộ hồ sơ": "IDD", "Mã phòng": "Sảnh thang L4"},
        {"Khu": "FF&E AREA 51 (cu)", "Bộ hồ sơ": "FF&E", "Mã phòng": "SK-01; ST-01", "Nguồn": [["FF&E 51", f_cu, 1]]},
        {"Khu": "FF&E AREA 51 (moi)", "Bộ hồ sơ": "FF&E", "Mã phòng": "SK-01; ST-01", "Nguồn": [["FF&E 51 moi", f_moi, 1]]},
        {"Khu": "AREA51 tham", "Bộ hồ sơ": "Thảm", "Mã phòng": "SK-01; ST-01", "Nguồn": [["dan FF&E lam can cu", f_cu, 1]]},
        {"Khu": "Schedules", "Bộ hồ sơ": "Schedule", "Mã phòng": "Toàn gói", "Nguồn": [["Finishes", f_sch, 1]]}]})
    ghi(os.path.join(data, "MAT.mat.json"), {"rows": [
        {"Mã VL": "PT-01", "Mô tả": "Son", "Trạng thái cuối": "Theo spec", "Nguồn spec": ["Finishes PT-01", f_sch, 1]},
        {"Mã VL": "SW-10", "Mô tả": "Lavabo", "Trạng thái cuối": "Theo spec"}, {"Mã VL": "ST-01", "Mô tả": "Da", "Trạng thái cuối": "Theo spec"}]})
    ghi(os.path.join(data, "_boq_map.json"), dict(doc_json(goc, "_boq_map.json"), ma_phong=["SK-01", "ST-01"]))
    ghi(os.path.join(data, "SCOPE.boq.json"), {"rows": [
        {"Khóa": f"{SH}!8", "Mã phòng": "SK-01:2 | CR-09:1"},
        {"Khóa": f"{SH}!9", "Mã phòng": "ST-01:3"}]})
    ghi(os.path.join(data, "SCOPE.banve-1.json"), {"rows": [
        {"Khóa": f"{SH}!8", "Mã spec/FF&E": "51-500B", "Bản vẽ hiệu lực": [["ID-LL-51-101 p1", f_idd, 1]]},
        {"Khóa": f"{SH}!9", "Mã VL": "PT-01; ST-01"}]})
    td_ok("area", goc)
    td_ok("xuathien", goc)
    return goc, data, dict(cu=f_cu, moi=f_moi, sch=f_sch, idd=f_idd, tcq=fid["TCQ/tcq.pdf"])


@kiem(4, "mau")
def gd4_ma_con_khong_tinh_vao_ma_cha(tmp):
    from scan_core import ma_chuan, regex_ma, tach_ma
    assert ma_chuan("52 - 500.1") == "52-500.1"
    assert not regex_ma("52-500").search("52 - 500.1"), "mã con bị tính vào mã cha"
    assert regex_ma("52-500").search("52 - 500 Sofa"), "mã cha phải khớp"
    assert regex_ma("52-500.1").search("x 52 - 500.1 y") and not regex_ma("52-500.1").search("52-500.2")
    d = {x[0]: x[3] for x in tach_ma("51 - 500B, ST-01, GL 02, SW10 0XE, 52 - 500.1")}
    assert d == {"51-500B": "ffe", "ST-01": "gach", "GL-02": "cach", "SW-10": "lien", "52-500.1": "ffe"}, d


@kiem(4, "mau")
def gd4_area_khong_noi_khac_bo_ho_so(tmp):
    """IDD Area 51 = sảnh thang không được lẫn vào tiền tố 51 của bộ FF&E."""
    goc, data, f = dung_xh(tmp)
    a = {r["Tiền tố"]: r for r in doc_json(goc, "AREA.area-tool.json")["rows"]}
    assert set(a) == {"51"}, a
    assert a["51"]["Mã phòng"] == "SK-01; ST-01", a["51"]
    assert "Sảnh" not in a["51"]["Mã phòng"] and "IDD" not in a["51"]["Bộ hồ sơ"]
    v = {r["Khóa"]: r for r in doc_json(goc, "VL_XUATHIEN.xuathien-tool.json")["rows"]}
    bv = v[f"51-500B|{f['idd']}|1"]
    assert bv["Bộ hồ sơ"] == "IDD" and bv["Mã phòng"] == "Sảnh thang L4", bv
    tr = v[f"51-500B|{f['cu']}|1"]
    assert tr["Bộ hồ sơ"] == "FF&E" and "Sảnh" not in tr["Mã phòng"], tr


@kiem(4, "mau")
def gd4_bay_st01_ma_da_va_ma_phong(tmp):
    goc, data, f = dung_xh(tmp)
    v = {r["Khóa"]: r for r in doc_json(goc, "VL_XUATHIEN.xuathien-tool.json")["rows"]}
    da = v[f"ST-01|{f['sch']}|1"]
    assert da["Loại trang"] == "Định nghĩa" and da["Mã phòng"] == "Toàn gói", da
    assert f"ST-01|{f['sch']}|3" not in v, "trang room schedule: ST-01 là mã phòng, không được ghi là mã đá"
    # chiều ngược: mã phòng trong tập phòng của tiền tố là phòng thật, không phải vật liệu
    assert "ST-01" in doc_json(goc, "AREA.area-tool.json")["rows"][0]["Mã phòng"]
    assert all(r["Mã VL"] != "SK-01" for r in v.values()), "mã phòng SK-01 bị ghi thành mã VL"


@kiem(4, "mau")
def gd4_so_khong_co_dong_khong_nguon_khong_bat_nham(tmp):
    goc, data, f = dung_xh(tmp)
    rows = doc_json(goc, "VL_XUATHIEN.xuathien-tool.json")["rows"]
    assert rows and all(r.get("File") and r["File"][1] for r in rows), "có dòng không nguồn"
    ma = {r["Mã VL"] for r in rows}
    v = {r["Khóa"]: r for r in rows}
    assert f"SW-10|{f['cu']}|1" not in v, "SW10 0XE là mã bưu điện, không được ghi"
    assert "44-207" not in ma, "số điện thoại bị nhận là mã"
    assert not any(r["File"][1] == f["tcq"] for r in rows), "bản PDF TCQ không dùng, TCQ đọc từ TCQ_LOG"
    assert "52-500.1" in ma and "52-500" not in ma, ma
    assert v[f"51-500B|{f['idd']}|1"]["Cờ"].startswith("Qua dòng BOQ"), "bản vẽ qua dòng BOQ phải có cờ"
    # mã trên trang spec nhưng không phải mã của trang: Nhắc tới, không phải Định nghĩa
    assert v[f"PT-01|{f['cu']}|1"]["Loại trang"] == "Nhắc tới"
    assert v[f"51-500B|{f['cu']}|3"]["Loại trang"] == "Ảnh"
    assert v[f"SW-05|{f['sch']}|2"]["Loại trang"] == "Định nghĩa", "trang DR có đúng 1 mã nằm chỗ khác"


@kiem(4, "mau")
def gd4_lech_area_va_hieu_luc_theo_tung_ma(tmp):
    goc, data, f = dung_xh(tmp)
    v = {r["Khóa"]: r for r in doc_json(goc, "VL_XUATHIEN.xuathien-tool.json")["rows"]}
    assert v[f"51-500B|{f['moi']}|1"]["Hiệu lực"] == "Mới nhất"
    assert v[f"51-500B|{f['cu']}|1"]["Hiệu lực"] == "Bản cũ"
    assert v[f"51-500B|{f['cu']}|2"]["Hiệu lực"] == "Bản cũ", "trang MATERIAL REFERENCE của file cũ"
    assert v[f"51-500B|BOQ|{SH}!8"]["Cờ"] == "Lệch Area: CR-09", v[f"51-500B|BOQ|{SH}!8"]


@kiem(4, "mau")
def gd4_link_spec_va_sua_area_qua_build(tmp):
    import glob
    import openpyxl
    goc, data, f = dung_xh(tmp)
    out = td_ok("build", goc)
    assert "VL_XUATHIEN" in out
    fn = file_track(goc)
    wb = openpyxl.load_workbook(fn)
    assert {"AREA", "VL_XUATHIEN"} <= set(wb.sheetnames), wb.sheetnames
    ws = wb[SH]
    h = cot_cua(ws)
    o8, o9 = f"{SH}!{h['Link vật liệu']}8", f"{SH}!{h['Link vật liệu']}9"
    dong = [r for r in wb["_LINKS"].iter_rows(min_row=2, values_only=True)]
    d8 = [r for r in dong if r[0] == o8]
    d9 = [r for r in dong if r[0] == o9]
    assert d8 and all(r[5] == 1 and "AREA 51 moi" in r[4] for r in d8), f"dòng 8 phải mở trang mới nhất: {d8}"
    # PT-01 và ST-01 cùng ở trang 1 của Finishes: 1 trang 1 link, nhãn theo mã đầu
    assert len(d9) == 1 and d9[0][5] == 1 and "Finishes" in d9[0][4] and d9[0][3].startswith("PT-01"), f"MAT Nguồn spec phải thắng: {d9}"
    assert not ws[f"{h['Link vật liệu']}7"].value
    # sổ VL_XUATHIEN không vào snapshot: sửa tay không thành user.json
    snap = doc_json(goc, "_snapshot.json")
    assert "VL_XUATHIEN" not in snap["reg"]
    ngay = {k: r["Ngày cập nhật"] for k, r in snap["reg"]["SCOPE"].items()}
    # sửa Area ở sheet giữ qua build, kể cả khi chạy lại area
    wa = wb["AREA"]
    hd = [c.value for c in wa[1]]
    wa.cell(row=2, column=hd.index("Mã phòng") + 1, value="SK-01")
    wb.save(fn)
    td_ok("build", goc)
    a = {r["Tiền tố"]: r for r in doc_json(goc, "AREA.user.json")["rows"]}
    assert a["51"]["Mã phòng"] == "SK-01", a
    td_ok("area", goc)
    wb2 = openpyxl.load_workbook(file_track(goc))
    ga = {r[0]: r[hd.index("Mã phòng")] for r in wb2["AREA"].iter_rows(min_row=2, values_only=True)}
    assert ga["51"] == "SK-01", ga
    moi = doc_json(goc, "_snapshot.json")["reg"]["SCOPE"]
    assert all(moi[k]["Ngày cập nhật"] == ngay[k] for k in ngay), "Link vật liệu không được làm đổi Ngày cập nhật"
    wc = openpyxl.load_workbook(glob.glob(os.path.join(goc, "10_TENDER", "04_out", "*_CDT.xlsx"))[0])
    assert all(c.value != "Link vật liệu" for row in wc[SH].iter_rows() for c in row), "bản CĐT lộ Link vật liệu"
    assert not {"AREA", "VL_XUATHIEN"} & set(wc.sheetnames)


@kiem(4, "mau")
def gd4_hang_khong_gian(tmp):
    """Cột thêm vào sheet BOQ không wrap (wrap làm hàng của CĐT tự giãn); sheet của tool cao cố định, kể cả bản CĐT."""
    import glob
    import openpyxl
    goc, data, f = dung_xh(tmp)
    td_ok("build", goc)
    wb = openpyxl.load_workbook(file_track(goc))
    ws = wb[SH]
    h = cot_cua(ws)
    for c, l in h.items():
        if c in ("Mô tả NT – Base", "Mã cấp 3", "Bên thực hiện", "Bản vẽ hiệu lực", "Link vật liệu"):
            assert not ws[f"{l}8"].alignment.wrap_text, f"cột {c} đang wrap"
    assert ws.row_dimensions[8].height is None, "hàng BOQ bị đặt chiều cao"
    for sh in ("SCOPE", "BOQ_TONGHOP", "VL_XUATHIEN", "MAT"):
        assert wb[sh].row_dimensions[2].height == 15, f"{sh}: {wb[sh].row_dimensions[2].height}"
    wc = openpyxl.load_workbook(glob.glob(os.path.join(goc, "10_TENDER", "04_out", "*_CDT.xlsx"))[0])
    assert not any(c.alignment.wrap_text for row in wc[SH].iter_rows(min_row=8) for c in row if c.value is not None and c.column > 6)


def o_co_link(ws):
    """Sau khi nạp file, ws._hyperlinks rỗng; link nằm ở từng ô."""
    return [c for row in ws.iter_rows() for c in row if c.hyperlink]


def khong_link_toi_sheet_an(wb, ten_file):
    """Link tới sheet ẩn không nhảy và không báo gì: cả hyperlink thường lẫn dòng _LINKS trỏ trong chính file. Đếm để phép kiểm không rỗng."""
    an = {s for s in wb.sheetnames if wb[s].sheet_state != "visible"}
    n = 0
    for ws in wb.worksheets:
        for c in o_co_link(ws):
            l = c.hyperlink
            dich = l.location or (l.target[1:] if l.target and l.target.startswith("#") else None)
            if dich:
                n += 1
                assert dich.split("!")[0].strip("'") not in an, f"{ws.title}!{c.coordinate} link tới sheet ẩn: {dich}"
    for r in wb["_LINKS"].iter_rows(min_row=2, values_only=True):
        if r[4] == ten_file and r[6]:
            n += 1
            assert str(r[6]).split("!")[0].strip("'") not in an, f"_LINKS trỏ sheet ẩn: {r}"
    assert n and any(o_co_link(ws) for ws in wb.worksheets), "không có link nào trong file, phép kiểm rỗng"
    return an


@kiem(4, "mau")
def gd4_sheet_an_va_nhe_file(tmp):
    import openpyxl
    goc, data, f = dung_xh(tmp)
    td_ok("build", goc)
    wb = openpyxl.load_workbook(file_track(goc))
    an = khong_link_toi_sheet_an(wb, os.path.basename(file_track(goc)))
    assert an == {"SCOPE", "DWG", "KHU", "AREA", "VL_XUATHIEN", "_LINKS", "_CODE", "CAU_TAO", "PHAN_VIEC"}, an
    assert wb.active.title == "DASHBOARD"
    assert {"DASHBOARD", "BOQ_TONGHOP", "MAT", "TCQ_LOG", "VIEC_MO", SH} <= {s for s in wb.sheetnames if s not in an}
    for sh in ("SCOPE", "DWG", "KHU", "AREA", "VL_XUATHIEN", "_LINKS"):
        assert not o_co_link(wb[sh]), f"{sh} còn hyperlink"
    assert wb["_LINKS"].max_column == 7, wb["_LINKS"].max_column
    hd = [c.value for c in wb["VL_XUATHIEN"][1]]
    assert not {"Khóa", "Dòng BOQ", "Chữ gốc"} & set(hd) and "Mã VL" in hd, hd
    r0 = doc_json(goc, "VL_XUATHIEN.xuathien-tool.json")["rows"][0]
    assert "Khóa" in r0 and "File" in r0, "JSON phải giữ đủ"
    assert any(str(x[0]).startswith(SH) for x in wb["_LINKS"].iter_rows(min_row=2, values_only=True)), "link ở sheet BOQ phải còn"
    dash = {r[0] for r in wb["DASHBOARD"].iter_rows(values_only=True)}
    assert "SHEET ẨN" in dash and "  SCOPE" in dash
    # có mẫu cấu tạo: CAU_TAO và PHAN_VIEC hiện lại; sửa ở sheet ẩn vẫn được kéo vào
    w2 = wb["SCOPE"]
    h = [c.value for c in w2[1]]
    ghi(os.path.join(data, "CAU_TAO.user.json"), {"rows": [{"Khóa mẫu": "CT-01#1", "Thành phần": "Khung", "Bên thực hiện": "Nhà máy"}]})
    hang = next(i for i, r in enumerate(w2.iter_rows(min_row=2, values_only=True), 2) if r[0] == f"{SH}!8")
    w2.cell(row=hang, column=h.index("Mô tả NT – Base") + 1, value="Sua o sheet an")
    wb.save(file_track(goc))
    td_ok("build", goc)
    wb = openpyxl.load_workbook(file_track(goc))
    assert wb["CAU_TAO"].sheet_state == "visible" and wb["PHAN_VIEC"].sheet_state == "visible"
    assert wb["SCOPE"].sheet_state == "hidden"
    u = {r["Khóa"]: r for r in doc_json(goc, "SCOPE.user.json")["rows"]}
    assert u[f"{SH}!8"]["Mô tả NT – Base"] == "Sua o sheet an", "sửa ở sheet ẩn bị mất"


def trang_cache(goc, fid, pg):
    return doc_json(goc, os.path.join("_text", fid + ".json"))["pages"][pg - 1]["text"]


@kiem(4, "sofitel")
def gd4_sofitel_so_xuat_hien_va_link_spec(tmp):
    """Plan GĐ4: không dòng không nguồn; bẫy ST-01; không nối Area khác bộ hồ sơ; Link spec mở đúng trang (trang đích có chữ mã)."""
    import glob
    import openpyxl
    from scan_core import regex_ma
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    truoc = chu_ky_thu_muc(data_that)
    goc, cd = ban_sao_gd3(tmp, nho=True)
    snap = doc_json(goc, "_snapshot.json")
    cu = snap["reg"]["SCOPE"]
    snap.pop("boq_track", None)
    snap.pop("file_index", None)
    ghi(os.path.join(goc, "10_TENDER", "_data", "_snapshot.json"), snap)
    out = td_ok("build", goc)
    assert "0 LỖI, 1 XEM" in out, out[-800:]  # 1 XEM duy nhất: nhắc chưa đọc sâu (GĐ7)
    assert "XEM  Chưa đọc sâu" in td("check", goc)[1]
    rows = doc_json(goc, "VL_XUATHIEN.xuathien-tool.json")["rows"]
    assert all(r.get("File") and r["File"][1] for r in rows), "có dòng không nguồn"
    v = {r["Khóa"]: r for r in rows}
    # ST-01: F0003 p99 là Room schedule (mã phòng), F0042 p10 là trang đá
    assert "ST-01|F0003|99" not in v, "Room schedule bị ghi thành mã đá"
    assert v["ST-01|F0042|10"]["Loại trang"] == "Định nghĩa" and "ST-01" not in v["ST-01|F0042|10"]["Mã phòng"].split("; ")
    assert "ST-01|F0003|95" in v
    area = {r["Tiền tố"]: r for r in doc_json(goc, "AREA.area-tool.json")["rows"]}
    assert "Sảnh" not in area["51"]["Mã phòng"] and "IDD" not in area["51"]["Bộ hồ sơ"]
    assert area["51"]["Mã phòng"].split("; ") == ["SK-01", "SK-02", "ST-01", "CR-01A"], area["51"]
    dn51 = [r for r in rows if r["Mã VL"] == "51-500B" and r["Nguồn"] == "Spec"]
    assert dn51 and all("Sảnh" not in r.get("Mã phòng", "") for r in dn51), dn51
    assert not any(r["Mã VL"].startswith(("SW-10|", "ST-09")) for r in rows if r.get("Chữ gốc", "").find("0XE") >= 0)
    # 51-500B đủ 7 trang gốc của GĐ1: mọi trang spec có mã
    trang = {(r["File"][1], r["Trang"]) for r in rows if r["Mã VL"] == "51-500B" and r["Nguồn"] == "Spec"}
    assert {("F0046", 9), ("F0046", 66), ("F0046", 67), ("F1167", 3)} <= trang, trang
    # Link spec mở đúng trang
    f = glob.glob(os.path.join(goc, "10_TENDER", "03_boq", "SSP_BOQ_TRACK_*.xlsx"))[0]
    moi = doc_json(goc, "_snapshot.json")["reg"]["SCOPE"]
    doi = [k for k in moi if moi[k]["Ngày cập nhật"] != cu[k]["Ngày cập nhật"]]
    assert not doi, f"{len(doi)} dòng đổi ngày: {doi[:5]}"
    import importlib.util
    n_link = n_trang = 0
    sai = []
    sp = importlib.util.spec_from_file_location("td_mod", TD)
    td_mod = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(td_mod)
    data = td_mod.merge(td_mod.P(goc))[0]
    td_mod.gan_link_vat_lieu(data)
    co_ma = [r for r in data["SCOPE"].values() if td_mod.ma_cua_dong(r, data["AREA"])]
    for r in co_ma:
        for l in r.get("Link vật liệu") or []:
            n_link += 1
            ma = l[0].split(" · ")[0]
            if " · " not in l[0]:  # trang tập FF&E chuyển từ Bản vẽ hiệu lực, nhãn do bước banve đặt, không phải mã
                continue
            if len(l) > 2 and l[2] and os.path.exists(os.path.join(goc, "10_TENDER", "_data", "_text", l[1] + ".json")):
                n_trang += 1
                if not regex_ma(ma).search(trang_cache(goc, l[1], int(l[2]))):
                    sai.append((r["Khóa"], l))
    assert not sai, f"{len(sai)} link spec trỏ trang không có mã: {sai[:5]}"
    assert n_trang > 500, f"chỉ kiểm {n_trang} link, không đủ mẫu"
    trong = [r["Khóa"] for r in co_ma if not r.get("Link vật liệu")]
    # các link _LINKS của cột Link spec trong sheet BOQ: file đích tồn tại, trang đúng
    wb = openpyxl.load_workbook(f, read_only=True)
    assert {"SCOPE", "VL_XUATHIEN", "_LINKS", "_CODE"} <= {s for s in wb.sheetnames if wb[s].sheet_state == "hidden"}
    wb_day = openpyxl.load_workbook(f)
    khong_link_toi_sheet_an(wb_day, os.path.basename(f))
    wb_day.close()
    lk_dong = [r for r in wb["_LINKS"].iter_rows(min_row=2, values_only=True) if r[3] and " · " in str(r[3])]
    thieu = [r[4] for r in lk_dong[:400] if r[4] and not os.path.exists(os.path.join(os.path.dirname(f), r[4]))]
    assert not thieu, f"{len(thieu)} file đích không có: {thieu[:3]}"
    lech = {}
    xl = mo_excel()
    if xl:
        try:
            excel_mo_khong_loi(xl, f)
            goc_cfg = doc_json(goc, "_cauhinh.json")
            a = xl.Workbooks.Open(os.path.join(goc_cfg["input"], goc_cfg["boq"]), 0, True)
            fc = glob.glob(os.path.join(goc, "10_TENDER", "04_out", "SSP_BOQ_*_CDT.xlsx"))[0]
            for ff in (f, fc):
                b = xl.Workbooks.Open(ff, 0, True)
                for ws in a.Worksheets:
                    n = ws.UsedRange.Row + ws.UsedRange.Rows.Count
                    sb = b.Worksheets(ws.Name)
                    d = sum(1 for i in range(1, n + 1) if abs(ws.Rows(i).RowHeight - sb.Rows(i).RowHeight) > 0.5)
                    if d:
                        lech[f"{os.path.basename(ff)[:12]} {ws.Name}"] = d
                b.Close(False)
            a.Close(False)
        finally:
            xl.Quit()
    assert not lech, f"hàng BOQ khác chiều cao so với file gốc: {lech}"
    assert chu_ky_thu_muc(data_that) == truoc, "_data thật đã bị ghi"
    print(f"      SOFITEL: hàng BOQ giữ đúng chiều cao gốc; VL_XUATHIEN {len(rows)} dòng, {n_link} link spec, kiểm trang {n_trang}/{n_link} đúng, {len(trong)}/{len(co_ma)} dòng có mã chưa có link spec")


# ---------------------------------------------------------------- GĐ5: DMVT
def nap_dmvt():
    tool = os.path.join(ROOT, "SKILL", "10-tender", "tool")
    if tool not in sys.path:
        sys.path.insert(0, tool)
    import dmvt
    import dmvt_xlsx
    return dmvt, dmvt_xlsx


def res_mau():
    return {"AREA": {"51": {"Mã phòng": "SK-01; SK-02"}, "52": {"Mã phòng": "GK-01"}},
            "MAT": {"WD-05": {"Trạng thái cuối": "Omit", "Nội dung cuối (VN)": "WD-05 bỏ (TCQ9-9)"},
                    "ST-07": {"Trạng thái cuối": "Đổi theo TCQ", "Nội dung cuối (VN)": "Đổi mẫu (TCQ9-8)"},
                    "GL-03": {"Trạng thái cuối": "VE", "Nội dung cuối (VN)": "Dùng acrylic"},
                    "PT-01": {"Trạng thái cuối": "Theo spec"}, "MR-05": {"Trạng thái cuối": "Không thấy spec", "Mô tả": "Kính gân"}},
            "SCOPE": {"S!1": {"Mã VL": "PT-01", "Mã phòng": "SK-01:2 | GK-01:1"}, "S!2": {"Mã VL": "WD-05", "Mã phòng": "SK-02:1"}},
            "TCQ_LOG": {"TCQ9-1": {"Mã VL": "51-500A", "Bị thay bởi": ""}, "TCQ9-2": {"Mã VL": "51-500B", "Bị thay bởi": "TCQ9-3"}, "TCQ9-3": {"Mã VL": "", "Bị thay bởi": ""}}}


def nguon_dr(ma, mo_ta="Vật liệu mẫu"):
    return {"loai": "DR", "fid": "F1", "trang": 1, "ngay": "2026-01-01", "bo": "FINISHES", "mo_ta": mo_ta, "ref": ["Mẫu A"], "finish": ["Mờ"], "ncc": "NCC A", "vi_tri": "As shown", "note": ""}


@kiem(5, "mau")
def gd5_ap_trang_thai_mat(tmp):
    """Mã đã Omit hoặc đã đổi không được hiện như spec gốc; FF&E chưa có ở MAT có cờ, không suy (plan mục 5, 8a)."""
    dmvt, _ = nap_dmvt()
    res = res_mau()
    moi = {m: nguon_dr(m) for m in ("WD-05", "ST-07", "GL-03", "PT-01")}
    moi["51-500A"] = {"loai": "MR", "fid": "F2", "trang": 1, "ngay": "2026-01-01", "bo": "FF&E", "tieu_de": "Sofa - Fabric", "kv": [["Design No", "Mau"], ["Colour", "1"]]}
    moi["51-501A"] = {"loai": "MR", "fid": "F2", "trang": 2, "ngay": "2026-01-01", "bo": "FF&E", "tieu_de": "Stool - OMITTED", "kv": [["Design No", "Da"], ["Colour", "2"]]}
    moi["51-500B"] = {"loai": "MR", "fid": "F2", "trang": 3, "ngay": "2026-01-01", "bo": "FF&E", "tieu_de": "Sofa - Cushion", "kv": [["Design No", "X"], ["Colour", "3"]]}
    d = lambda m: dmvt.viet_dong([m], moi, res, {}, [])
    x = d("WD-05")
    assert x["Omit"] and "Omit" in x["Cờ"] and "OMITTED" in x["Remarks"] and x["Material"], x
    x = d("ST-07")
    assert not x["Omit"] and "Đổi theo TCQ" in x["Cờ"] and "Đổi mẫu" in x["Remarks"] and "Mẫu A" in x["Description"], "phải giữ nội dung spec, ghi kết luận vào Remarks"
    x = d("GL-03")
    assert "VE" in x["Cờ"] and "acrylic" in x["Remarks"]
    x = d("PT-01")
    assert not x["Remarks"] and "Omit" not in x["Cờ"]
    assert x["Area"] == "GK-01\nSK-01" or set(x["Area"].split()) == {"GK-01", "SK-01"}, x["Area"]
    assert "Area theo dòng BOQ" in x["Cờ"]
    x = d("51-500A")
    assert "Chưa qua MAT" in x["Cờ"] and "TCQ nhắc mã" in x["Cờ"] and "TCQ9-1" in x["Remarks"] and "Area theo tiền tố" in x["Cờ"]
    x = d("51-500B")
    assert "TCQ nhắc mã" not in x["Cờ"], "câu TCQ đã bị thay thì không tính"
    x = d("51-501A")
    assert x["Omit"], "spec ghi thẳng OMITTED thì giữ dòng và tô xám"
    x = dmvt.viet_dong(["MR-05"], moi, res, {}, [])
    assert "Không thấy spec" in x["Cờ"] and x["Material"] == "Kính gân"


@kiem(5, "mau")
def gd5_doc_khoi_va_gop_design_colour(tmp):
    dmvt, _ = nap_dmvt()
    t = "MATERIAL REFERENCE\nArea:\n51\nITEM\n51 - 300A\nDate: 30/01/2026\nRevised: \nBlind - Fabric\nDesign No:\nMercuro\nColour:\n44\nContent:\nPolyester\nWidth:\n255 cm\nRepeat:\n-\nSupplier:\nSilent - SG\nTel:\n+65 1\nFax:\nContact:\nPauline\nE-Mail:\np@x.sg\nITEM\n51 - 500A\nDate: 30/01/2026\nRevised: 05/06/2026 A\nSofa - Fabric\nDesign No:\nRuskin\nColour:\n7757/58\nSupplier:\nRomo\n"
    k = dmvt.ma_khoi_mr(t)
    assert [m for m, _ in k] == ["51-300A", "51-500A"], k
    td_, kv = dmvt.doc_khoi_mr(*k[0])
    assert td_ == "Blind - Fabric", td_
    assert dmvt.thong_so(kv) == "Design No: Mercuro\nColour: 44\nContent: Polyester\nWidth: 255 cm", dmvt.thong_so(kv)
    assert dmvt.lien_he(kv) == "Silent - SG\nContact: Pauline\nTel: +65 1\nE: p@x.sg", dmvt.lien_he(kv)
    sp = dmvt.doc_spec("SPECIFICATION\nItem:\n51 - 200\nDate:\n30/01/2026\nDescription: \nRug\nSupplier:\nRoyal\nRef:\nCustom\nContent:\n80% Wool\nSPECIAL INSTRUCTIONS:\nGlass top 6mm")
    assert sp["ma"] == "51-200" and sp["mo_ta"] == "Rug" and "Glass top" in sp["si"], sp
    assert dmvt.doc_spec("SPECIFICATION\nManufacturer to provide...\nItem\n51 - 707") is None, "trang tiếp theo của spec không phải trang định nghĩa"


@kiem(5, "mau")
def gd5_chon_anh_uu_tien_anh_chup_va_hash(tmp):
    from PIL import Image, ImageDraw
    import random
    dmvt, _ = nap_dmvt()
    net = Image.new("RGB", (800, 800), "white")
    ImageDraw.Draw(net).line((0, 0, 800, 800), fill="black", width=3)
    rnd = random.Random(1)
    chup = Image.new("RGB", (200, 200))
    chup.putdata([(rnd.randrange(255), rnd.randrange(255), rnd.randrange(255)) for _ in range(40000)])
    ung = [{"img": net}, {"img": chup}]
    assert dmvt.chon_mot(ung) == 1, "ưu tiên ảnh chụp dù bản vẽ lớn hơn (quyết định 12)"
    assert dmvt.chon_mot(ung, "lon_nhat") == 0
    assert dmvt.chon_mot([]) is None
    nen = Image.new("RGB", (400, 400), "white")
    ImageDraw.Draw(nen).rectangle((50, 50, 200, 300), fill="black")
    assert dmvt.lech(dmvt.phash(nen), dmvt.phash(nen.resize((230, 230)))) <= dmvt.HASH_LECH, "khác độ phân giải vẫn phải trúng"
    assert dmvt.lech(dmvt.phash(nen), dmvt.phash(nen.transpose(Image.ROTATE_90))) > dmvt.HASH_LECH


def dong_may(ma, **o):
    d = {"C": "SK-01", "D": "Mô tả", "E": "Vị trí", "F": "Spec", "G": "NCC", "I": ""}
    d.update(o)
    return {"Khóa": " / ".join(ma), "Mã": ma, "o": d, "Cờ": [], "Cờ ảnh": [], "Mục": "I", "Tên mục": "M", "Omit": False, "Ảnh": "a.png"}


def dong_cu(ma, hang=8, anh=None, sau=None, duoi=None, **o):
    d = {"C": "SK-01", "D": "Mô tả", "E": "Vị trí", "F": "Spec", "G": "NCC", "I": ""}
    d.update(o)
    return {"hang": hang, "loai": "ma", "tap": frozenset(ma), "ai": dict(d, A=1, B="\n".join(ma)), "sau": sau or {}, "duoi": duoi or [], "cao": 160, "anh": anh}


def snap_cua(ma, anh=b"", **o):
    d = {"C": "SK-01", "D": "Mô tả", "E": "Vị trí", "F": "Spec", "G": "NCC", "I": "", "_ma": ma}
    d.update(o)
    import hashlib as h
    d["anh"] = h.md5(anh).hexdigest() if anh else ""
    return {" / ".join(ma): d}


@kiem(5, "mau")
def gd5_giu_o_sua_tay_va_bao_spec_moi(tmp):
    _, x = nap_dmvt()
    A = ["51-300A"]
    # không ai sửa: lấy chữ máy mới
    ra, tk = x.gop_sua_tay([dong_may(A, F="Spec mới")], [dong_cu(A)], snap_cua(A))
    assert ra[0]["o"]["F"] == "Spec mới" and not ra[0]["spec_moi"], ra[0]
    # người sửa, spec không đổi: giữ, không báo
    ra, tk = x.gop_sua_tay([dong_may(A)], [dong_cu(A, F="Người sửa")], snap_cua(A))
    assert ra[0]["o"]["F"] == "Người sửa" and not ra[0]["spec_moi"] and tk["ô sửa tay giữ"] == 1
    # người sửa, spec đổi: giữ chữ người, báo spec mới, không đè
    ra, tk = x.gop_sua_tay([dong_may(A, F="Spec mới")], [dong_cu(A, F="Người sửa")], snap_cua(A))
    assert ra[0]["o"]["F"] == "Người sửa" and ra[0]["spec_moi"] == ["F: Spec mới"], ra[0]
    # J trở đi, dòng không mã đi theo dòng có mã phía trên
    ra, tk = x.gop_sua_tay([dong_may(A)], [dong_cu(A, sau={10: "91 days", 39: "cột tender thêm"}, duoi=[{"ai": {"C": "Ray trượt"}, "sau": {}, "cao": 50, "anh": None}])], snap_cua(A))
    assert ra[0]["sau"] == {10: "91 days", 39: "cột tender thêm"} and ra[0]["duoi"][0]["ai"]["C"] == "Ray trượt", ra[0]
    # thứ tự mã người gõ ở cột B không làm mất khớp
    ra, tk = x.gop_sua_tay([dong_may(["51-1", "52-1"], F="m")], [dong_cu(["52-1", "51-1"], F="tay")], {"51-1 / 52-1": snap_cua(["51-1", "52-1"])["51-1 / 52-1"]})
    assert ra[0]["o"]["F"] == "tay", ra[0]


@kiem(5, "mau")
def gd5_sua_tay_song_qua_nhieu_lan_build(tmp):
    """Mốc snapshot phải là chữ máy: ghi chữ người vào mốc thì build lần 2 tưởng chưa ai sửa và đè mất."""
    _, x = nap_dmvt()
    A = ["51-300A"]
    goc, nguoi = (b"anh may", "png"), (b"anh nguoi", "jpg")
    snap = {" / ".join(A): dict(snap_cua(A, anh=goc[0])[A[0]], hash="h1")}
    cu = [dong_cu(A, F="Người sửa", anh=nguoi)]
    for vong in range(3):
        may = dong_may(A, F="Spec mới")
        may["Hash ảnh"] = "h2"
        ra, tk = x.gop_sua_tay([may], cu, snap)
        r = ra[0]
        assert r["o"]["F"] == "Người sửa", f"vòng {vong}: mất chữ người"
        assert r["anh_cu"] == nguoi, f"vòng {vong}: mất ảnh người"
        assert r["spec_moi"] == ["F: Spec mới", "H: ảnh trong spec đã đổi"], f"vòng {vong}: báo spec mới phải còn tới khi người xử lý: {r['spec_moi']}"
        sn = x.moc_snapshot(r)
        assert sn["F"] == "Spec" and sn["anh"] == snap[A[0]]["anh"] and sn["hash"] == "h1", sn
        snap = {A[0]: sn}
    # người gõ đúng chữ spec mới: thôi báo
    ra, tk = x.gop_sua_tay([dong_may(A, F="Spec mới")], [dong_cu(A, F="Spec mới")], {A[0]: snap_cua(A)[A[0]]})
    assert not ra[0]["spec_moi"], ra[0]["spec_moi"]


@kiem(5, "mau")
def gd5_giu_anh_nguoi_thay(tmp):
    _, x = nap_dmvt()
    A = ["51-300A"]
    goc, moi = (b"anh may", "png"), (b"anh nguoi", "jpg")
    ra, tk = x.gop_sua_tay([dong_may(A)], [dong_cu(A, anh=goc)], snap_cua(A, anh=goc[0]))
    assert ra[0]["anh_cu"] is None, "ảnh không đổi thì dùng ảnh máy cắt"
    ra, tk = x.gop_sua_tay([dong_may(A)], [dong_cu(A, anh=moi)], snap_cua(A, anh=goc[0]))
    assert ra[0]["anh_cu"] == moi and tk["ảnh sửa tay giữ"] == 1, ra[0]
    ra, tk = x.gop_sua_tay([dong_may(A)], [dong_cu(A)], snap_cua(A, anh=goc[0]))
    assert ra[0]["anh_cu"] == ("xoa",), "người xóa ảnh thì build lại không chèn lại"


@kiem(5, "mau")
def gd5_excel_nen_lai_anh_khong_tinh_la_sua(tmp):
    """Excel nén lại ảnh khi lưu nên bytes đổi dù không ai sửa; tính nhầm thì ảnh bị đóng băng, spec đổi cũng không cập nhật."""
    import io
    from PIL import Image, ImageDraw
    _, x = nap_dmvt()

    def anh(mau, q):
        im = Image.new("RGB", (300, 300), mau)
        ImageDraw.Draw(im).ellipse((60, 60, 240, 240), fill=(40, 90, 160))
        b = io.BytesIO()
        im.save(b, "JPEG", quality=q)
        return b.getvalue()
    goc, nen, khac = anh((200, 180, 160), 95), anh((200, 180, 160), 60), anh((30, 160, 60), 95)
    assert goc != nen
    A = ["51-300A"]
    sn = dict(snap_cua(A)[A[0]], anh=x.dau_anh(goc))
    ra, _ = x.gop_sua_tay([dong_may(A)], [dong_cu(A, anh=(nen, "jpg"))], {A[0]: sn})
    assert ra[0]["anh_cu"] is None, "nén lại không phải sửa tay"
    ra, _ = x.gop_sua_tay([dong_may(A)], [dong_cu(A, anh=(khac, "jpg"))], {A[0]: sn})
    assert ra[0]["anh_cu"] == (khac, "jpg"), "ảnh khác thật phải được giữ"


@kiem(5, "mau")
def gd5_tach_gop_tap_ma_co_co_khong_tu_chuyen(tmp):
    _, x = nap_dmvt()
    cu = [dong_cu(["51-1", "52-1"], F="tay", sau={10: "91 days"})]
    snap = snap_cua(["51-1", "52-1"])
    ra, tk = x.gop_sua_tay([dong_may(["51-1"], F="m1"), dong_may(["52-1"], F="m2")], cu, snap)
    moi = {r["Khóa"]: r for r in ra}
    assert moi["51-1"]["o"]["F"] == "m1" and not moi["51-1"]["sau"], "tách mã: không tự chuyển sửa tay"
    assert any("Tập mã đổi" in c for c in moi["51-1"]["Cờ"]) and any("Tập mã đổi" in c for c in moi["52-1"]["Cờ"]), moi
    giu = [r for r in ra if r["Tên mục"].startswith("DÒNG CŨ")]
    assert len(giu) == 1 and giu[0]["sau"] == {10: "91 days"}, "công người phải còn ở dòng cũ giữ lại"
    # gộp
    ra, tk = x.gop_sua_tay([dong_may(["51-1", "52-1"])], [dong_cu(["51-1"]), dong_cu(["52-1"], hang=9)], {**snap_cua(["51-1"]), **snap_cua(["52-1"])})
    assert any("Tập mã đổi" in c for c in ra[0]["Cờ"]), ra[0]
    assert tk["dòng cũ bỏ"] == 2, "dòng cũ không có công người thì bỏ"


@kiem(5, "mau")
def gd5_dong_cu_mat_ma_va_dong_user_them(tmp):
    _, x = nap_dmvt()
    ra, tk = x.gop_sua_tay([dong_may(["51-1"])], [dong_cu(["51-1"]), dong_cu(["59-9"], hang=9), dong_cu(["58-8"], hang=10, sau={10: "x"}), dong_cu(["57-7"], hang=11)],
                           {**snap_cua(["51-1"]), **snap_cua(["59-9"]), **snap_cua(["58-8"])})
    k = {r["Khóa"]: r for r in ra}
    assert "59-9" not in k, "máy ghi, không ai sửa, không còn trong spec: bỏ"
    assert k["58-8"]["Cờ"] == ["Không còn trong spec"] and k["58-8"]["sau"] == {10: "x"}
    assert k["57-7"]["Cờ"] == ["Dòng user thêm"]


@kiem(5, "mau")
def gd5_form_sach_khong_co_du_lieu_du_an(tmp):
    f = os.path.join(ROOT, "SKILL", "10-tender", "form", "DMVT_form.xlsx")
    with zipfile.ZipFile(f) as z:
        ten = z.namelist()
        assert not [n for n in ten if "/media/" in n or "comments" in n], "form còn ảnh hoặc comment"
        chu = z.read("xl/sharedStrings.xml").decode("utf-8") + z.read("xl/workbook.xml").decode("utf-8")
    for x in ("SOFITEL", "Sofitel", "KAWAJUN", "Royal Thai", "CHECK VE"):
        assert x not in chu, x
    assert os.path.getsize(f) < 100_000
    with open(os.path.join(ROOT, ".gitignore"), encoding="utf-8") as fh:
        assert "!SKILL/10-tender/form/DMVT_form.xlsx" in fh.read(), "form phải lên git"


@kiem(5, "sofitel")
def gd5_sofitel_bao_toan_build_lai_va_link(tmp):
    """Vòng thật trên bản sao: build, người sửa bằng Excel, build lại; Omit tô xám, TCQ có Remarks; ảnh nằm đúng ô; Link vật tư tới đúng dòng."""
    import glob
    import openpyxl
    dmvt, x = nap_dmvt()
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    truoc = chu_ky_thu_muc(data_that)
    goc, cd = ban_sao_gd3(tmp, nho=True)
    data = os.path.join(goc, "10_TENDER", "_data")
    for t in ("_dmvt.json", "_dmvt_llm.json"):
        shutil.copy2(os.path.join(data_that, t), data)
    td_ok("dmvt", goc)
    f = glob.glob(os.path.join(goc, "10_TENDER", "03_boq", "*_DMVT_*.xlsx"))[0]
    wb = openpyxl.load_workbook(f)
    ws = wb["DMVT"]
    hang = {}
    for h in range(8, ws.max_row + 1):
        for m in x.tap_ma(ws.cell(h, 2).value):
            hang[m] = h
    # trạng thái MAT trên file thật
    for m, co in (("WD-05", "Omit"), ("ST-09", "Omit"), ("ST-07", "Đổi theo TCQ"), ("WD-01", "Đổi theo TCQ"), ("GL-03", "VE")):
        assert co in str(ws.cell(hang[m], 34).value), (m, ws.cell(hang[m], 34).value)
        assert ws.cell(hang[m], 9).value, f"{m} thiếu Remarks"
    assert str(ws.cell(hang["WD-05"], 1).fill.fgColor.rgb).endswith("D9D9D9"), "Omit phải tô xám"
    assert str(ws.cell(hang["PT-01"], 1).fill.fgColor.rgb) in ("00000000", "None") or not ws.cell(hang["PT-01"], 1).fill.fill_type
    assert "Omit" in str(ws.cell(hang["51-501A"], 34).value), "mã spec ghi OMITTED"
    assert set(wb.sheetnames) == {"DMVT", "_SOI_ANH", "_LINKS"} and wb.sheetnames[0] == "DMVT" and wb["_LINKS"].sheet_state == "hidden", wb.sheetnames
    # _LINKS: đích có thật, trang có chữ mã
    ws_l = wb["_LINKS"]
    so = 0
    for r in ws_l.iter_rows(min_row=2, values_only=True):
        assert os.path.exists(os.path.normpath(os.path.join(os.path.dirname(f), r[4]))), r
        so += 1
    assert so > 150, so
    n_anh = len(x.anh_trong_file(f))
    wb.close()
    # người sửa bằng Excel
    xl = mo_excel()
    assert xl, "cần Excel"
    h_f, h_a, h_i = hang["ST-07"], hang["PT-01"], hang["IR-07"]
    anh_nguoi = os.path.join(tmp, "nguoi.png")
    from PIL import Image
    Image.new("RGB", (120, 90), (200, 30, 30)).save(anh_nguoi)
    try:
        excel_mo_khong_loi(xl, f)
        wb = xl.Workbooks.Open(f)
        w = wb.Worksheets("DMVT")
        w.Cells(h_f, 6).Value = "SỬA TAY F"
        w.Cells(h_f, 10).Value = "91 days"
        w.Cells(h_a, 12).Value = "ghi chú giá"
        w.Rows(h_i + 1).Insert()
        w.Cells(h_i + 1, 3).Value = "Ray trượt cửa WC"
        for i in range(1, w.Shapes.Count + 1):
            sh = w.Shapes.Item(i)
            if sh.Type == 13 and sh.TopLeftCell.Row == h_a:
                sh.Delete()
                break
        c = w.Cells(h_a, 8)
        p = w.Shapes.AddPicture(anh_nguoi, False, True, c.Left + 2, c.Top + 2, 60, 45)
        p.Placement = 1
        wb.Saved = False
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    os.utime(f)
    out = td_ok("dmvt", goc)
    wb = openpyxl.load_workbook(f)
    ws = wb["DMVT"]
    hang = {}
    for h in range(8, ws.max_row + 1):
        for m in x.tap_ma(ws.cell(h, 2).value):
            hang[m] = h
    assert ws.cell(hang["ST-07"], 6).value == "SỬA TAY F" and ws.cell(hang["ST-07"], 10).value == "91 days"
    assert ws.cell(hang["PT-01"], 12).value == "ghi chú giá"
    assert ws.cell(hang["IR-07"] + 1, 3).value == "Ray trượt cửa WC", "dòng không mã phải đi theo dòng có mã phía trên"
    anh = x.anh_trong_file(f)
    im = Image.open(__import__("io").BytesIO(anh[hang["PT-01"]][0])).convert("RGB")
    assert im.getpixel((5, 5))[0] > 150 and im.getpixel((5, 5))[1] < 80, "ảnh người thay phải còn"
    assert len(anh) == n_anh, (len(anh), n_anh)
    wb.close()
    assert "'ô sửa tay giữ': 1" in out and "'ảnh sửa tay giữ': 1" in out, f"chỉ sửa 1 ô, 1 ảnh, nếu nhiều hơn là nhận nhầm: {out}"
    # build lại lần nữa: không có gì trôi
    out2 = td_ok("dmvt", goc)
    assert "'ô sửa tay giữ': 1" in out2 and "'ảnh sửa tay giữ': 1" in out2, out2
    wb = openpyxl.load_workbook(f)
    ws = wb["DMVT"]
    assert ws.cell(hang["ST-07"], 6).value == "SỬA TAY F"
    assert ws.cell(hang["IR-07"] + 1, 3).value == "Ray trượt cửa WC"
    im = Image.open(__import__("io").BytesIO(x.anh_trong_file(f)[hang["PT-01"]][0])).convert("RGB")
    assert im.getpixel((5, 5))[0] > 150 and im.getpixel((5, 5))[1] < 80, "ảnh người thay mất ở lần build thứ 2"
    wb.close()
    # Link vật tư: BOQ_TRACK trỏ đúng dòng DMVT
    td_ok("build", goc, "--khong-dmvt")
    t = glob.glob(os.path.join(goc, "10_TENDER", "03_boq", "*_BOQ_TRACK_*.xlsx"))
    t = max(t, key=os.path.getmtime)
    wb = openpyxl.load_workbook(t)
    wl = wb["_LINKS"]
    wd = openpyxl.load_workbook(f)["DMVT"]
    kq = n = 0
    for r in wl.iter_rows(min_row=2, values_only=True):
        if r[6] and "'DMVT'!A" in str(r[6]) and r[1] != "SCOPE":
            n += 1
            assert os.path.normcase(os.path.normpath(os.path.join(os.path.dirname(t), r[4]))) == os.path.normcase(os.path.normpath(f)), r
            ma = str(r[3]).split(" · ")[0]
            h = int(str(r[6]).split("!A")[1])
            assert ma in x.tap_ma(wd.cell(h, 2).value), (r, wd.cell(h, 2).value)
            kq += 1
    assert kq > 300, kq
    wb.close()
    assert chu_ky_thu_muc(data_that) == truoc, "_data thật đã bị ghi"
    print(f"      SOFITEL: {len(hang)} mã; giữ sửa tay qua 2 lần build; {kq} link vật tư trỏ đúng dòng")


# ---------------------------------------------------------------- GĐ6: đọc file qua _core
PD = os.path.join(ROOT, "SKILL", "01-scan", "tool", "pd.py")


def chay_goc(tool, args):
    """(mã thoát khác 0?, stdout dạng byte): so từng byte, không qua giải mã. Mã lỗi chỉ so có lỗi hay không, vì traceback khác nhau."""
    r = subprocess.run([sys.executable, "-I", tool, *args], capture_output=True)
    return r.returncode != 0, r.stdout


def so_voi_ban_cu(args, png=None, moi=None):
    """Chạy cùng lệnh qua pd.py và td.py (hoặc script đối chứng moi), cùng đầu vào. Trả chuỗi mô tả chỗ khác, giống hệt thì None."""
    kq = []
    for tool in (PD, moi or TD):
        if png and os.path.exists(png):
            os.remove(png)
        loi, ra = chay_goc(tool, args)
        anh = b""
        if png and os.path.exists(png):
            with open(png, "rb") as fh:
                anh = fh.read()
        kq.append((loi, ra, anh))
    (l0, r0, a0), (l1, r1, a1) = kq
    if l0 != l1:
        return f"{args}: pd.py lỗi={l0}, td.py lỗi={l1}"
    if r0 != r1:
        return f"{args}: stdout khác ({len(r0)} / {len(r1)} byte)"
    if a0 != a1:
        return f"{args}: ảnh khác ({len(a0)} / {len(a1)} byte)"
    return None


def so_nhieu(ca):
    """ca = [(args, png)], chạy song song vì mỗi lệnh phải khởi động 2 tiến trình. -> [chỗ khác]."""
    from concurrent.futures import ThreadPoolExecutor
    with ThreadPoolExecutor(8) as ex:
        return [k for k in ex.map(lambda c: so_voi_ban_cu(*c), ca) if k]


def ca_pdf(pdf, trang, tmp, ten):
    png = os.path.join(tmp, f"{ten}_{trang}.png")
    png2 = os.path.join(tmp, f"{ten}_{trang}_clip.png")
    return [
        (["chu", pdf, str(trang)], None),
        (["chu", pdf, str(trang), "--clip", "0.1,0.2,0.9,0.8"], None),
        (["anh", pdf, str(trang), png, "--dpi", "50"], png),
        (["anh", pdf, str(trang), png2, "--dpi", "72", "--clip", "0.5,0.5,1,1"], png2),
    ]


@kiem(6, "mau")
def gd6_tender_khong_con_goi_pd(tmp):
    goc = os.path.join(ROOT, "SKILL", "10-tender")
    sot = []
    for d, _, ten in os.walk(goc):
        if "__pycache__" in d:
            continue
        for t in ten:
            if t.lower().endswith((".md", ".py", ".json")):
                with open(os.path.join(d, t), encoding="utf-8") as fh:
                    noi_dung = fh.read()
                if "pd.py" in noi_dung or re.search(r"01-scan[\\/]+tool", noi_dung):
                    sot.append(os.path.relpath(os.path.join(d, t), goc))
    assert not sot, f"10-tender còn gọi pd.py: {sot}"
    for buoc, lenh in (("1-kiemke", "mail"), ("2-boq", "xem"), ("4-dwg", "chu"), ("5-mat", "chu"), ("6-banve", "anh")):
        with open(os.path.join(goc, "buoc", buoc + ".md"), encoding="utf-8") as fh:
            assert f'py -I "<TOOL>" {lenh} ' in fh.read(), f"{buoc}.md không còn hướng dẫn lệnh {lenh}"


@kiem(6, "mau")
def gd6_mau_giong_ban_cu(tmp):
    import openpyxl
    import pymupdf
    d = os.path.join(tmp, "thư mục có dấu")
    os.makedirs(d)
    pdf = os.path.join(d, "mẫu đọc.pdf")
    with pymupdf.open() as doc:
        doc.new_page().insert_text((72, 100), "A-101 chu mau 51 - 500B")
        xoay = doc.new_page(width=600, height=400)  # trang xoay: clip tính theo rect sau khi xoay
        xoay.insert_text((50, 50), "ST-01 trang xoay")
        xoay.set_rotation(90)
        doc.new_page()  # trang trắng
        doc.save(pdf)
    xlsx = os.path.join(d, "mẫu.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet có dấu"
    ws.append(["Item", None, "x" * 100, 1.5])
    ws.append([])
    ws.append([None, "Mô tả tiếng Việt", "", 0])
    wb.create_sheet("Trống")
    wb.save(xlsx)
    ca = [c for tr in (1, 2, 3) for c in ca_pdf(pdf, tr, d, "mau")]
    ca += [(a, None) for a in (
        ["xem", xlsx], ["xem", xlsx, "Sheet có dấu"], ["xem", xlsx, "Sheet có dấu", "B1:C3"], ["xem", xlsx, "Trống"],
        ["xem", xlsx, "Không có sheet này"], ["chu", pdf, "9"], ["anh", pdf, "9", os.path.join(d, "x.png")],
        ["chu", os.path.join(d, "không có.pdf"), "1"], ["mail", os.path.join(d, "không có.msg")])]
    khac = so_nhieu(ca)
    assert not khac, "\n".join(khac)

    # đối chứng: phép so phải bắt được khi đầu ra lệch, nếu không thì "giống hệt" ở trên không có giá trị
    gia = os.path.join(d, "gia.py")
    png = os.path.join(d, "dc.png")
    with open(gia, "w", encoding="utf-8") as fh:
        fh.write("import sys\nprint('khác')\n")
    assert so_voi_ban_cu(["xem", xlsx], moi=gia), "không bắt được stdout lệch"
    with open(gia, "w", encoding="utf-8") as fh:
        fh.write("import sys\nsys.exit(1)\n")
    assert "lỗi" in (so_voi_ban_cu(["xem", xlsx], moi=gia) or ""), "không bắt được lệnh báo lỗi"
    lenh_anh = ["anh", pdf, "1", png, "--dpi", "36", "--clip", "0,0,0.5,0.5"]
    _, ra_dung = chay_goc(PD, lenh_anh)
    with open(gia, "w", encoding="utf-8") as fh:  # stdout đúng từng byte, ảnh sai
        fh.write(f"import sys\nopen({png!r},'wb').write(b'x')\nsys.stdout.buffer.write({ra_dung!r})\n")
    assert "ảnh" in (so_voi_ban_cu(lenh_anh, png, gia) or ""), "không bắt được ảnh lệch"
    print(f"      {len(ca)} lệnh giống pd.py từng byte, kể cả ảnh; ca lỗi cùng báo lỗi; đối chứng bắt được lệch")


@kiem(6, "sofitel")
def gd6_sofitel_giong_ban_cu(tmp):
    data = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    with open(os.path.join(data, "_cauhinh.json"), encoding="utf-8") as fh:
        cfg = json.load(fh)
    inp = os.path.join(GOC_SOFITEL, cfg["input"])
    rows = doc_index(GOC_SOFITEL)
    msg = [os.path.join(inp, r["Đường dẫn"]) for r in rows if r["Đường dẫn"].lower().endswith(".msg") and not r.get("Trong zip")]
    assert msg, "không thấy file .msg"
    ca = [(["mail", m], None) for m in msg]
    n_mail = len(ca)

    n_xem = 0
    for rel in (cfg["boq"], cfg["tcq"]):
        x = os.path.join(inp, rel)
        loi, ra = chay_goc(TD, ["xem", x])
        assert not loi and ra, f"xem {rel} không ra gì"
        ten_sheet = re.findall(r"^\[(.+)\] ", ra.decode("utf-8"), re.M)
        ca.append((["xem", x], None))
        ca += [(["xem", x, t], None) for t in ten_sheet]
        ca += [(["xem", x, t, "A1:AP40"], None) for t in ten_sheet]
        n_xem += 1 + 2 * len(ten_sheet)

    pdfs = sorted(f for f in os.listdir(os.path.join(data, "_zip")) if f.lower().endswith(".pdf"))
    assert len(pdfs) >= 8, f"chỉ có {len(pdfs)} PDF trong _zip"
    import pymupdf
    chon = pdfs[:: len(pdfs) // 8][:8]  # rải đều: có spec, bản vẽ nhiều lớp, ảnh
    n_pdf = 0
    for i, f in enumerate(chon):
        p = os.path.join(data, "_zip", f)
        with pymupdf.open(p) as doc:
            n = doc.page_count
        for tr in sorted({1, (n + 1) // 2, n}):
            ca += ca_pdf(p, tr, tmp, f"p{i}")
            n_pdf += 4
    khac = so_nhieu(ca)
    assert not khac, f"{len(khac)}/{len(ca)} lệnh khác:\n" + "\n".join(khac[:20])
    print(f"      SOFITEL: {len(ca)} lệnh giống pd.py từng byte: mail {n_mail} file, xem {n_xem}, chu + anh {n_pdf} ({len(chon)} PDF)")


# ---------------------------------------------------------------- build ra đủ 2 file (user yêu cầu 01/11)
def file_dmvt(goc):
    import glob
    return glob.glob(os.path.join(goc, "10_TENDER", "03_boq", "*_DMVT_*.xlsx"))


@kiem(6, "mau")
def gd6_build_khong_cau_hinh_van_ra_2_file(tmp):
    goc, _, data = dung_boq(tmp)
    assert not os.path.exists(os.path.join(data, "_dmvt.json"))
    rc, out = td_that("build", goc)
    assert rc == 0, out
    assert file_dmvt(goc), f"thiếu DMVT:\n{out}"
    assert file_track(goc), "thiếu BOQ_TRACK"
    assert os.path.basename(file_dmvt(goc)[0]) in out.split("Đã ghi:")[-1], "phần báo kết quả phải nêu file DMVT"
    assert "Chưa có _data" in out, "dùng mặc định phải báo XEM"


@kiem(6, "mau")
def gd6_dmvt_hong_thi_thoat_loi_van_giu_boq(tmp):
    goc, _, data = dung_boq(tmp)
    with open(os.path.join(data, "_dmvt.json"), "w", encoding="utf-8") as fh:
        fh.write("{hỏng")
    rc, out = td_that("build", goc)
    assert rc != 0, "thiếu DMVT mà vẫn báo thành công"
    assert "chưa đủ 2 file" in out, out
    assert file_track(goc), "DMVT hỏng không được làm mất BOQ_TRACK"
    assert not file_dmvt(goc)


@kiem(6, "mau")
def gd6_khong_dmvt_chi_khi_yeu_cau(tmp):
    goc, _, _ = dung_boq(tmp)
    rc, out = td_that("build", goc, "--khong-dmvt")
    assert rc == 0 and not file_dmvt(goc) and "--khong-dmvt" in out, out


@kiem(6, "sofitel")
def gd6_sofitel_khong_cau_hinh_van_ra_2_file(tmp):
    """Giả lập dự án chưa tạo _dmvt.json: bản sao SOFITEL bỏ mọi file _dmvt*."""
    import openpyxl
    data_that = os.path.join(GOC_SOFITEL, "10_TENDER", "_data")
    truoc = chu_ky_thu_muc(data_that)
    goc, _ = ban_sao_gd3(tmp, nho=True)
    rc, out = td_that("build", goc)
    assert rc == 0, out[-1500:]
    import glob
    f = file_dmvt(goc)
    assert f and glob.glob(os.path.join(goc, "10_TENDER", "03_boq", "*_BOQ_TRACK_*.xlsx")), out[-1500:]
    wb = openpyxl.load_workbook(f[0], read_only=True)
    n = sum(1 for r in wb["DMVT"].iter_rows(min_row=8, max_col=2, values_only=True) if r[1])
    wb.close()
    assert n > 100, f"DMVT chỉ có {n} dòng có mã"
    assert chu_ky_thu_muc(data_that) == truoc, "_data thật đã bị ghi"
    print(f"      SOFITEL không có _dmvt.json: đủ 2 file, DMVT {n} hàng có chữ ở cột B (gồm hàng tiêu đề mục)")


# ---------------------------------------------------------------- chạy
# ---------------------------------------------------------------- GĐ7: đọc sâu (chuyển từ 01-scan)
def khoa_cua(data, mo_ta):
    with open(os.path.join(data, "SCOPE.boq-tool.json"), encoding="utf-8") as fh:
        return [r["Khóa"] for r in json.load(fh)["rows"] if r.get("Mô tả BOQ") == mo_ta]


def agy_gia(tmp, ten, dong):
    """agy giả: in đúng 1 dòng rồi thoát; dòng rỗng = thoát lỗi không in gì."""
    f = os.path.join(tmp, ten + ".bat")
    with open(f, "w", encoding="ascii") as fh:
        fh.write(f"@echo {dong}\n" if dong else "@exit /b 1\n")
    return f


def png_gia(tmp, ten):
    f = os.path.join(tmp, ten + ".png")
    with open(f, "wb") as fh:
        fh.write(b"x")
    return f


def dong_qs(cum, truong, gia_tri, tt="CÓ", nguon="HÌNH", to=True):
    r = {"Cụm đọc": cum, "Trường": truong, "Giá trị": gia_tri, "Trạng thái": tt, "Nguồn": nguon}
    if to:
        r["Tờ · Chi tiết"] = ["p1", "F0001", 1]
    return r


def du_an_doc_sau(tmp):
    goc, cd, data = dung_boq(tmp)
    ghi(os.path.join(data, "FILE_INDEX.kiemke-tool.json"), {"rows": [{"ID": "F0001", "Đường dẫn": "bv.pdf", "Vai trò": "Chính"}]})
    return goc, data


@kiem(7, "mau")
def gd7_ham_tach_so_giong_pd(tmp):
    import importlib.util
    sys.path.insert(0, os.path.join(ROOT, "SKILL", "_core"))
    spec = importlib.util.spec_from_file_location("pd_cu", os.path.join(ROOT, "SKILL", "01-scan", "tool", "pd.py"))
    pd = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(pd)
    import neo
    import gemini_doc
    mau = ["900×550×440", "1,200 x 600", "1.200x600mm", "W900 D550", "450+450=900", "1540 (R) × 2375 (C)", "900×?×440", "±130x200", "12/10/2026 900",
           "ID-SK-52-101 1200", "010 900", "Rev 3 2000", "WD-01 và PT1, GL 2", "WD1a 51-400 MK630 18mm F024 p11 900x650mmH", "", None, "+84 28 3824 1525 1200x300"]
    for x in mau:
        assert neo.neo_tach_so(x) == pd.neo_tach_so(x), x
        assert neo.neo_kt_gt(x) == pd.neo_kt_gt(x), x
        assert neo.neo_mavl(x, {"WD", "PT", "GL"}) == pd.neo_mavl(x, {"WD", "PT", "GL"}), x
        assert neo.neo_ma(x) == pd.neo_ma(x), x
    assert (gemini_doc.GEM_SO, gemini_doc.GEM_QS, gemini_doc.GEM_TRUONG, gemini_doc.GEM_MODEL) == (pd.GEM_SO, pd.GEM_QS, pd.GEM_TRUONG, pd.GEM_MODEL)
    assert gemini_doc.gem_doc('x ["900", "?"] y') == pd.gem_doc('x ["900", "?"] y') and gemini_doc.gem_doc("LỖI agy") is None


@kiem(7, "mau")
def gd7_so_giong_pd(tmp):
    import pymupdf
    f = os.path.join(tmp, "a.pdf")
    d = pymupdf.open()
    t = d.new_page()
    t.insert_text((50, 50), "900 x 450 1,200 WD-01 2375")
    d.save(f)
    d.close()
    a = subprocess.run([sys.executable, "-I", os.path.join(ROOT, "SKILL", "01-scan", "tool", "pd.py"), "so", f, "1"], capture_output=True).stdout
    b = subprocess.run([sys.executable, "-I", TD, "so", f, "1"], capture_output=True).stdout
    assert a.strip() and a.split() == b.split(), (a, b)


@kiem(7, "mau")
def gd7_gom_cum_dung_pham_vi(tmp):
    goc, data = du_an_doc_sau(tmp)
    ke = khoa_cua(data, "Cai ke")
    assert len(ke) == 2
    tu, cua = khoa_cua(data, "Cai tu 12")[0], khoa_cua(data, "Cai cua")[0]
    ghi(os.path.join(data, "SCOPE.doichieu.json"), {"rows": [{"Khóa": cua, "Trạng thái": "Omit"}]})
    td_ok("docsau", goc)
    cum = doc_json(goc, "_docsau.json")["cum"]
    assert cum[ke[0]]["khoa"] == ke, "2 dòng giống hệt phải 1 cụm, đại diện là dòng đầu"
    assert tu in cum and cua not in cum, "dòng Omit không đọc sâu"
    assert not any("Prelims" in k for k in cum), "sheet Prelims không đọc sâu"
    assert not khoa_cua(data, "Chu thich khong don vi")[0] in cum, "dòng không đơn vị không đọc sâu"
    # khác tờ là khác món: tách cụm
    ghi(os.path.join(data, "SCOPE.banve.json"), {"rows": [{"Khóa": ke[0], "Bản vẽ hiệu lực": [["A-101", "F0001", 1]]}, {"Khóa": ke[1], "Bản vẽ hiệu lực": [["A-102", "F0001", 2]]}]})
    td_ok("docsau", goc)
    cum = doc_json(goc, "_docsau.json")["cum"]
    assert ke[0] in cum and ke[1] in cum
    # lọc chỉ đổi phần giao, không đổi cụm
    n = len(cum)
    td_ok("docsau", goc, "--sheet", "không có")
    d = doc_json(goc, "_docsau.json")
    assert d["nhom"] == [] and len(d["cum"]) == n


@kiem(7, "mau")
def gd7_check_du_an_chua_doc_sau_khong_loi(tmp):
    import openpyxl
    goc, cd, data = dung_boq(tmp)
    ghi(os.path.join(data, "SCOPE.doichieu.json"), {"rows": [{"_loc": {"Khóa": "."}, "Trạng thái": "Base"}]})
    rc, out = td("check", goc)
    assert "Chưa đọc sâu" in out, out[-500:]
    assert not re.search(r"LỖI .*(QUAN_SAT|Độ tin cậy|Kiểm chéo)", out), out[-800:]
    td_ok("build", goc)
    wb = openpyxl.load_workbook(file_track(goc))
    assert not {"QUAN_SAT", "BOQ_MAP", "ISSUES", "RFI_LOG", "ASK_USER", "SPEC", "PAGES", "DAY_TUONG"} & set(wb.sheetnames), wb.sheetnames


@kiem(7, "mau")
def gd7_check_doi_du_quan_sat(tmp):
    goc, data = du_an_doc_sau(tmp)
    td_ok("docsau", goc)
    k = khoa_cua(data, "Cai tu 12")[0]
    rc, out = td("check", goc, "--den", "docsau")
    assert rc == 1 and f"{k}: chưa có dòng QUAN_SAT" in out, out[-600:]
    rc, out = td("check", goc, "--den", "doichieu")
    assert "chưa có dòng QUAN_SAT" not in out, "bước trước docsau không đòi đọc sâu"
    ghi(os.path.join(data, "QUAN_SAT.docsau-1.json"), {"rows": [dong_qs(k, "Kích thước BV", "900×450×2000", to=False), dong_qs(k, "Dày", "450+450=1000")]})
    rc, out = td("check", goc, "--den", "docsau")
    assert "không ghi tờ/chi tiết" in out and "chuỗi cộng sai" in out and f"{k}: thiếu dòng QUAN_SAT 'Kiểm chéo Gemini'" in out, out[-800:]
    # sai cột, sai giá trị danh sách vẫn bị bắt như mọi sheet
    ghi(os.path.join(data, "QUAN_SAT.docsau-2.json"), {"rows": [{"Cụm đọc": k, "Trường": "X", "Giá trị": "y", "Trạng thái": "có", "Nguồn": "HÌNH", "Lạ": 1}]})
    rc, out = td("check", goc, "--den", "docsau")
    assert "cột lạ 'Lạ'" in out and "'Trạng thái' = 'có'" in out, out[-800:]
    # BOQ_MAP không được thêm khóa mới
    ghi(os.path.join(data, "BOQ_MAP.docsau-1.json"), {"rows": [{"Khóa KL": "khong|co", "KL theo BV": 1}]})
    rc, out = td("check", goc, "--den", "docsau")
    assert "khóa 'khong|co' không có trong nháp tool" in out, out[-800:]


@kiem(7, "mau")
def gd7_gemini_doc_lap_va_tin_cay(tmp):
    goc, data = du_an_doc_sau(tmp)
    td_ok("docsau", goc)
    k = khoa_cua(data, "Cai tu 12")[0]
    ke = khoa_cua(data, "Cai ke")[0]
    os.environ["TD_AGY"] = agy_gia(tmp, "agy1", '["900","450","2000","?"]')
    try:
        rc, out = td("gemini", png_gia(tmp, "c0"), "--cum", k, "--goc", goc)
        assert rc != 0 and "Ghi số của Claude" in out, "phải buộc Claude ghi trước để Gemini độc lập: " + out[-300:]
        hang = [dong_qs(c, "Kích thước BV", "900×450×2000") for c in (k, ke)] + [dong_qs(c, "Kiểm chéo Gemini", "khớp: Có cánh; lệch: không; crop: p1", nguon="GEMINI", to=False) for c in (k, ke)]
        ghi(os.path.join(data, "QUAN_SAT.docsau-1.json"), {"rows": hang})
        ghi(os.path.join(data, "SCOPE.docsau-1.json"), {"rows": [{"Khóa": c, "Bản vẽ đối chứng": [["A-101", "F0001", 1]], "Kích thước BV": "900×450×2000"} for c in (k, ke)]})
        out = td_ok("gemini", png_gia(tmp, "c1"), "--cum", k, "--goc", goc)
        assert "khớp: 450, 900, 2000" in out, out[-300:]
        os.environ["TD_AGY"] = agy_gia(tmp, "agy2", '["900","450"]')
        out = td_ok("gemini", png_gia(tmp, "c2"), "--cum", ke, "--goc", goc)
        assert "chỉ Claude: 2000" in out, out[-300:]
        bm = doc_json(goc, "BOQ_MAP.docsau-tool.json")["rows"]
        ghi(os.path.join(data, "BOQ_MAP.docsau-1.json"), {"rows": [{"Khóa KL": r["Khóa KL"], "KL theo BV": r["KL BOQ"]} for r in bm if r["Khóa"] == k]})
        td_ok("tincay", goc)
        ct = doc_json(goc, "_tincay.json")
        assert ct["muc"][k] == "CAO", ct["ly_do"][k]
        assert ct["muc"][ke] == "THAP" and any("Gemini không đọc ra số: 2000" in x for x in ct["ly_do"][ke]), ct["ly_do"][ke]
        chua = khoa_cua(data, "Cai ghe go")[0]
        assert ct["muc"][chua] == "THAP" and any("chưa có quan sát" in x for x in ct["ly_do"][chua])
        iss = doc_json(goc, "ISSUES.tincay-tool.json")["rows"]
        assert any(r["Cụm đọc"] == ke and "2000" in r["Nội dung"] for r in iss), iss
        # chạy lại không nhân đôi issue, dòng Gemini số
        td_ok("tincay", goc)
        iss2, gs = doc_json(goc, "ISSUES.tincay-tool.json")["rows"], doc_json(goc, "QUAN_SAT.tincay-tool.json")["rows"]
        assert len(iss2) == len(iss) and len({r["Cụm đọc"] for r in gs}) == len(gs), (len(iss2), len(iss), gs)
        assert doc_json(goc, "_tincay.json")["muc"] == ct["muc"], "chạy lại phải cùng kết quả"
        # agy lỗi: không được CAO, không phải THAP vì chưa có bằng chứng lệch
        os.environ["TD_AGY"] = agy_gia(tmp, "agy3", "")
        os.remove(os.path.join(data, "_gemini_so", re.sub(r'[\\/:*?"<>|]', "_", k) + ".json"))
        td_ok("gemini", png_gia(tmp, "c3"), "--cum", k, "--goc", goc)
        td_ok("tincay", goc)
        ct = doc_json(goc, "_tincay.json")
        assert ct["muc"][k] == "TB" and any("LỖI" in x for x in ct["ly_do"][k]), ct["ly_do"][k]
    finally:
        os.environ.pop("TD_AGY", None)


@kiem(7, "mau")
def gd7_neo_bat_so_sai_tren_trang_co_chu(tmp):
    goc, data = du_an_doc_sau(tmp)
    td_ok("docsau", goc)
    k = khoa_cua(data, "Cai tu 12")[0]
    ke = khoa_cua(data, "Cai ke")[0]
    chu = "A-101 PLAN\n" + "\n".join(str(x) for x in (900, 450, 2000, 300, 600, 750, 1200, 1500, 1800, 2100, 2400)) + "\nWD-01 PT-02\n"
    ghi(os.path.join(data, "_text", "F0001.json"), {"ck": "x", "pages": [{"text": chu, "so": "A-101", "ten": "PLAN", "lop": "CÓ"}, {"text": "khong co gi o day " * 3, "so": "A-102", "ten": "", "lop": "CÓ"}]})
    ghi(os.path.join(data, "QUAN_SAT.docsau-1.json"), {"rows": [dong_qs(k, "Kích thước BV", "900×450×2000"), dong_qs(ke, "Kích thước BV", "900×777×2000"),
                                                               dong_qs(k, "Mã hoàn thiện trên BV", "WD-01"), dong_qs(ke, "Mã hoàn thiện trên BV", "PT-02"), dong_qs(ke, "Mã hoàn thiện trên BV", "WD-09")]})
    ghi(os.path.join(data, "SCOPE.docsau-1.json"), {"rows": [{"Khóa": k, "Bản vẽ đối chứng": [["A-101", "F0001", 1]]}, {"Khóa": ke, "Bản vẽ đối chứng": [["A-101", "F0001", 2]]}]})
    out = td_ok("neo", goc)
    n = doc_json(goc, "_neo.json")
    ly = " | ".join(l["ly_do"] for l in n["lech"])
    assert "777" in ly, out[-900:]
    assert "WD-09" in ly, out[-900:]
    assert "không có số tờ A-101" in ly, "link trỏ tờ khác phải LỆCH: " + out[-900:]
    assert n["theo_kieu"]["kt"]["CÓ NEO"] == 1, n["theo_kieu"]
    rc, x = td("check", goc, "--den", "docsau")
    assert "NEO LỆCH" in x, x[-500:]
    # link đổi: phải nhắc chạy lại, không in kết quả cũ
    ghi(os.path.join(data, "SCOPE.docsau-1.json"), {"rows": [{"Khóa": k, "Bản vẽ đối chứng": [["A-101", "F0001", 2]]}]})
    rc, x = td("check", goc, "--den", "docsau")
    assert "chạy lại lệnh neo" in x and "NEO LỆCH" not in x, x[-500:]


@kiem(7, "mau")
def gd7_build_co_doc_sau(tmp):
    import glob
    import openpyxl
    goc, data = du_an_doc_sau(tmp)
    ghi(os.path.join(data, "SCOPE.doichieu.json"), {"rows": [{"_loc": {"Khóa": "."}, "Trạng thái": "Base"}]})
    td_ok("docsau", goc)
    ke = khoa_cua(data, "Cai ke")
    ghi(os.path.join(data, "SCOPE.docsau-1.json"), {"rows": [{"Khóa": ke[0], "Kích thước BV": "900×450×2000", "Bản vẽ đối chứng": [["A-101", "F0001", 1]], "Phạm vi nhà máy": "TRONG SCOPE – mộc"}]})
    ghi(os.path.join(data, "QUAN_SAT.docsau-1.json"), {"rows": [dong_qs(ke[0], "Kích thước BV", "900×450×2000")]})
    ghi(os.path.join(data, "BOQ_MAP.docsau-1.json"), {"rows": [{"Khóa KL": f"{ke[0]}|(cả dòng)", "KL theo BV": 3}, {"Khóa KL": f"{ke[1]}|(cả dòng)", "KL theo BV": 4}]})
    ghi(os.path.join(data, "ASK_USER.tincay.json"), {"rows": [{"ID": "A01", "Câu hỏi": "Tính hay loại trừ?"}]})
    ghi(os.path.join(data, "RFI_LOG.tincay.json"), {"rows": [{"RFI": "RFI-01", "Câu hỏi": "Kích thước?", "Trạng thái": "Nháp"}]})
    td_ok("tincay", goc)
    assert doc_json(goc, "_tincay.json")["muc"][ke[0]] == "TB"
    td_ok("build", goc)
    f = file_track(goc)
    wb = openpyxl.load_workbook(f)
    an = {x for x in wb.sheetnames if wb[x].sheet_state == "hidden"}
    assert {"QUAN_SAT", "BOQ_MAP"} <= an, wb.sheetnames
    assert {"ISSUES", "RFI_LOG", "ASK_USER"} <= set(wb.sheetnames) - an, wb.sheetnames
    ws = wb["BOQ_TONGHOP"]
    c = {x.value: x.column for x in ws[1]}
    dong = {ws.cell(r, 1).value: r for r in range(2, ws.max_row + 1)}
    for k in ke:  # dòng thành viên nhận kết quả của đại diện
        assert ws.cell(dong[k], c["Kích thước BV"]).value == "900×450×2000", k
        assert ws.cell(dong[k], c["Độ tin cậy"]).value == "TB", k
    assert ws.cell(dong[ke[0]], c["Phạm vi nhà máy"]).value == "TRONG SCOPE – mộc"
    assert ws.cell(dong[ke[0]], c["KL theo BV"]).value == 3 and ws.cell(dong[ke[0]], c["Lệch KL"]).value == 2, "BOQ 1, BV 3"
    assert ws.cell(dong[ke[1]], c["Lệch KL"]).value == 0
    d = "\n".join(str(x.value) for r in wb["DASHBOARD"].iter_rows() for x in r if x.value is not None)
    assert "ĐỌC SÂU" in d and "Tính hay loại trừ?" in d, d[-500:]
    # sửa tay giữ qua build, không bị dòng đại diện đè
    sc = wb["SCOPE"]
    h = {x.value: x.column for x in sc[1]}
    r_ = next(r for r in range(2, sc.max_row + 1) if sc.cell(r, 1).value == ke[1])
    sc.cell(r_, h["Kích thước BV"], "800×400×1000")
    wa = wb["ASK_USER"]
    wa.cell(2, [x.value for x in wa[1]].index("Trả lời") + 1, "Tính")
    wb.save(f)
    td_ok("build", goc)
    wb = openpyxl.load_workbook(f)
    ws = wb["BOQ_TONGHOP"]
    assert ws.cell(dong[ke[1]], c["Kích thước BV"]).value == "800×400×1000" and ws.cell(dong[ke[0]], c["Kích thước BV"]).value == "900×450×2000"
    assert wb["ASK_USER"].cell(2, [x.value for x in wb["ASK_USER"][1]].index("Trả lời") + 1).value == "Tính"
    # bản gửi CĐT không lộ đọc sâu
    wc = openpyxl.load_workbook(glob.glob(os.path.join(goc, "10_TENDER", "04_out", "T_BOQ_*_CDT.xlsx"))[0])
    nd = {x.value for w in wc.worksheets for r in w.iter_rows() for x in r if isinstance(x.value, str)}
    lo = nd & {"Độ tin cậy", "Kích thước BV", "Giả định tính giá", "Tính hay loại trừ?", "900×450×2000"}
    assert not lo, lo
    assert not {"QUAN_SAT", "ISSUES", "ASK_USER", "RFI_LOG"} & set(wc.sheetnames)


def main():
    args = set(sys.argv[1:])
    if args - {"--mau", "--sofitel"}:
        sys.exit(__doc__)
    chon = {a.lstrip("-") for a in args} or {"mau", "sofitel"}
    co_du_lieu = os.path.isdir(os.path.join(GOC_SOFITEL, "10_TENDER", "_data"))
    if "sofitel" in chon and not co_du_lieu:
        if args:
            sys.exit(f"LỖI: không có dữ liệu {GOC_SOFITEL}")
        print("BỎ QUA --sofitel: máy này không có _ACTIVE\\SOFITEL")
        chon.discard("sofitel")
    dat = hong = 0
    for gd, loai, fn in BO_KIEM:
        if loai not in chon:
            continue
        tmp = tempfile.mkdtemp(prefix="tender_test_")
        try:
            fn(tmp)
            dat += 1
            print(f"ĐẠT   GĐ{gd} {fn.__name__}")
        except AssertionError as e:
            hong += 1
            print(f"HỎNG  GĐ{gd} {fn.__name__}\n      {e}")
        except Exception:
            hong += 1
            print(f"LỖI   GĐ{gd} {fn.__name__}\n{traceback.format_exc()}")
        finally:
            os.environ.pop("TD_CODE_ID_DIR", None)
            shutil.rmtree(tmp, ignore_errors=True)
    print(f"\n{dat} đạt, {hong} hỏng")
    sys.exit(1 if hong else 0)


if __name__ == "__main__":
    main()
