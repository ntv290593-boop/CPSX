r"""guilink.py — lệnh `td.py guilink`: chép BOQ_TRACK + DMVT ra ngoài dự án để gửi tender, viết lại đích link (_LINKS cột E)
theo cây thư mục hồ sơ của tender (zip OneDrive tải về hoặc thư mục), vì cây của tender khác cây 00_input máy build.

Khớp file theo NỘI DUNG (CRC32 + kích thước), không theo tên: tender hay đổi tên thư mục, file trùng nội dung nằm nhiều chỗ.
File nằm trong zip/rar/7z của cây tender: ghi đường dẫn như khi giải nén ra thư mục cùng tên ("Extract to <tên>\"), báo tender giải nén.
Đường dẫn ghi tương đối theo GỐC cây tender (không có ..\): đặt xlsx ở gốc thì mở thẳng; đặt sâu hơn thì add-in SRC_LINKS
bỏ ..\ rồi thử từng thư mục cha (SL_ResolvePath bước 3) nên vẫn tìm được. BASE_DIR (H2) xóa vì là đường dẫn máy build.
Lưu bằng Excel COM (giữ ảnh, richData), không lưu bằng openpyxl.
"""
import os, sys, re, io, glob, shutil, zipfile, zlib, subprocess, tempfile
from collections import Counter

BS = "\\"
NEN = (".zip", ".rar", ".7z")
BO_QUA = re.compile(r"(^|[\\/])(~\$|\.~lock)|\.(tmp|bak|dwl2?)$", re.I)
SEVENZ = [r"C:\Program Files\7-Zip\7z.exe", r"C:\Program Files (x86)\7-Zip\7z.exe"]


def _p(s):
    return s.replace("/", BS).strip(BS)


def _crc_file(path):
    c = 0
    with open(path, "rb") as h:
        for b in iter(lambda: h.read(1 << 20), b""):
            c = zlib.crc32(b, c)
    return c


def _crc_bytes(b):
    return zlib.crc32(b)


def _7z():
    for x in SEVENZ:
        if os.path.exists(x):
            return x
    return shutil.which("7z")


def _liet_ke_7z(file_nen):
    """[(đường dẫn trong file nén, size, crc)] qua 7z l -slt; không có 7-Zip thì trả None."""
    exe = _7z()
    if not exe:
        return None
    r = subprocess.run([exe, "l", "-slt", "-sccUTF-8", file_nen], capture_output=True)
    out, cur, ra = r.stdout.decode("utf-8", "replace"), {}, []
    for dong in out.splitlines() + [""]:
        if not dong.strip():
            if cur.get("Path") and cur.get("CRC") and cur.get("Folder", "-") != "+":
                ra.append((_p(cur["Path"]), int(cur.get("Size") or 0), int(cur["CRC"], 16)))
            cur = {}
            continue
        if " = " in dong:
            k, v = dong.split(" = ", 1)
            cur[k.strip()] = v.strip()
    return [x for x in ra if os.path.normcase(x[0]) != os.path.normcase(file_nen)]


class Cay:
    """Chỉ mục cây hồ sơ tender: (crc, size) -> [(đường dẫn trong cây, file nén cần giải hoặc None)].
    Chỉ tính CRC cho file có kích thước bằng 1 file đích (cây thư mục lớn, đỡ đọc toàn bộ)."""

    def __init__(self, nguon, sizes):
        self.idx, self.sizes, self.khong_doc = {}, set(sizes), []
        if os.path.isfile(nguon) and nguon.lower().endswith(".zip"):
            self._zip(zipfile.ZipFile(nguon), "", None)
        elif os.path.isdir(nguon):
            self._thu_muc(nguon)
        else:
            sys.exit(f"LỖI: {nguon} không phải file zip hay thư mục")

    def _them(self, crc, size, path, nen):
        self.idx.setdefault((crc, size), []).append((path, nen))

    def _goc_nen(self, path):
        return path[: -len(os.path.splitext(path)[1])]  # "Extract to <tên>\"

    def _zip(self, z, tien_to, nen):
        for i in z.infolist():
            if i.is_dir() or BO_QUA.search(i.filename):
                continue
            path = tien_to + _p(i.filename)
            if i.file_size in self.sizes:
                self._them(i.CRC, i.file_size, path, nen)
            ext = os.path.splitext(path)[1].lower()
            if ext in NEN and nen is None:  # 1 tầng nén lồng là đủ (OneDrive zip → zip/rar CĐT)
                self._nen_long(lambda: z.read(i), path, ext)

    def _nen_long(self, doc, path, ext):
        goc = self._goc_nen(path) + BS
        ten = os.path.basename(path)
        if ext == ".zip":
            try:
                self._zip(zipfile.ZipFile(io.BytesIO(doc())), goc, ten)
                return
            except zipfile.BadZipFile:
                pass
        with tempfile.TemporaryDirectory() as d:
            f = os.path.join(d, "x" + ext)
            with open(f, "wb") as h:
                h.write(doc())
            ds = _liet_ke_7z(f)
        if ds is None:
            self.khong_doc.append(path + " (cần 7-Zip để đọc)")
            return
        for ten_trong, size, crc in ds:
            if size in self.sizes and not BO_QUA.search(ten_trong):
                self._them(crc, size, goc + ten_trong, ten)

    def _thu_muc(self, goc):
        for d, _, fs in os.walk(goc):
            for f in fs:
                full = os.path.join(d, f)
                path = os.path.relpath(full, goc)
                if BO_QUA.search(path):
                    continue
                size = os.path.getsize(full)
                if size in self.sizes:
                    self._them(_crc_file(full), size, path, None)
                ext = os.path.splitext(f)[1].lower()
                if ext in NEN:
                    if ext == ".zip":
                        self._nen_long(lambda full=full: open(full, "rb").read(), path, ext)
                    else:
                        ds = _liet_ke_7z(full)
                        if ds is None:
                            self.khong_doc.append(path + " (cần 7-Zip để đọc)")
                            continue
                        g = self._goc_nen(path) + BS
                        for ten_trong, sz, crc in ds:
                            if sz in self.sizes and not BO_QUA.search(ten_trong):
                                self._them(crc, sz, g + ten_trong, f)


def chon(ung_vien, rel_cu):
    """Nhiều chỗ cùng nội dung: giống đuôi đường dẫn cũ nhất (giữ đúng đợt/bản hiệu lực đã chọn lúc build),
    rồi file có sẵn trước file phải giải nén, rồi đường dẫn ngắn."""
    cu = [x.lower() for x in _p(rel_cu).split(BS) if x not in ("..", ".")]

    def duoi(path):
        moi, n = [x.lower() for x in path.split(BS)], 0
        while n < min(len(moi), len(cu)) and moi[-1 - n] == cu[-1 - n]:
            n += 1
        return n
    return sorted(ung_vien, key=lambda u: (-duoi(u[0]), u[1] is not None, len(u[0]), u[0]))[0]


def doc_links(f):
    import openpyxl
    ws = openpyxl.load_workbook(f, read_only=True)["_LINKS"]
    rows = list(ws.iter_rows(min_row=1, max_col=8, values_only=True))
    base = rows[1][7] if len(rows) > 1 and rows[0][7] == "BASE_DIR" else None
    return [r[4] for r in rows[1:]], base


def tim_file_may(xlsx, rel, base):
    for d in (os.path.dirname(xlsx), base):
        if d:
            f = os.path.normpath(os.path.join(d, rel))
            if os.path.isfile(f):
                return f
    return None


def main(p, args, opt):
    if not args or args[0].startswith("--"):
        sys.exit('LỖI: guilink <GOC> <zip hoặc thư mục hồ sơ của tender> [--ra <dir>] [--ngay yymmdd]')
    nguon = os.path.abspath(args[0])
    d_boq = os.path.join(p.t, "03_boq")
    ngay = opt(args, "--ngay")
    tracks = sorted(glob.glob(os.path.join(d_boq, f"{p.cfg['du_an']}_BOQ_TRACK_{ngay or '*'}.xlsx")))
    if not tracks:
        sys.exit(f"LỖI: không có {p.cfg['du_an']}_BOQ_TRACK_{ngay or '<yymmdd>'}.xlsx trong {d_boq}")
    track = tracks[-1]
    stamp = re.search(r"_(\d{6})\.xlsx$", track).group(1)
    dmvt = os.path.join(d_boq, f"{p.cfg['du_an']}_DMVT_{stamp}.xlsx")
    gui = [f for f in (dmvt, track) if os.path.exists(f)]
    if dmvt not in gui:
        print(f"XEM: không có {os.path.basename(dmvt)}, chỉ gửi BOQ_TRACK (link Link vật tư sẽ không mở được)")
    ra = os.path.abspath(opt(args, "--ra") or os.path.join(os.path.expanduser("~"), "Downloads", f"{p.cfg['du_an']}_GUI_TENDER_{stamp}"))
    for cam in (p.inp, p.t):  # 00_input: kiemke sẽ quét nhầm như hồ sơ; 10_TENDER: lẫn với bản dùng ở máy build
        if os.path.normcase(ra + BS).startswith(os.path.normcase(os.path.abspath(cam) + BS)):
            sys.exit(f"LỖI: --ra {ra} nằm trong {cam}; bản gửi tender để ngoài dự án (mặc định Downloads)")
    os.makedirs(ra, exist_ok=True)
    ten_gui = {os.path.basename(f).lower() for f in gui}

    # 1. đích của mọi link, file trên máy build
    dich, may = {}, {}
    for f in gui:
        rels, base = doc_links(f)
        dich[f] = rels
        for r in rels:
            if r and r.lower() not in ten_gui and r not in may:
                may[r] = tim_file_may(f, r, base)
    thieu_may = sorted(r for r, x in may.items() if not x)
    sig = {r: (_crc_file(x), os.path.getsize(x)) for r, x in may.items() if x}
    print(f"Đích link: {len(may)} file ({len(thieu_may)} không thấy trên máy build). Đọc cây tender {nguon} ...")

    # 2. tìm trong cây tender
    cay = Cay(nguon, {s[1] for s in sig.values()})
    moi, can_giai = {}, Counter()
    for r, sg in sig.items():
        uv = cay.idx.get(sg)
        if uv:
            path, nen = chon(uv, r)
            moi[r] = path
            if nen:
                can_giai[nen] += 1
    khong_co = sorted(r for r in sig if r not in moi) + thieu_may

    # 3. ghi bản sao bằng Excel COM
    from dmvt_xlsx import mo_excel
    for f in gui:
        if os.path.exists(os.path.join(ra, "~$" + os.path.basename(f))):
            sys.exit(f"LỖI: {os.path.join(ra, os.path.basename(f))} đang mở trong Excel — đóng file rồi chạy lại.")
    xl = mo_excel()
    xl.EnableEvents = False
    dem = {}
    try:
        for f in gui:
            out = os.path.join(ra, os.path.basename(f))
            shutil.copy2(f, out)
            wb = xl.Workbooks.Open(out, UpdateLinks=0)
            if wb.ReadOnly:
                wb.Close(False)
                sys.exit(f"LỖI: {out} chỉ đọc")
            ws = wb.Worksheets("_LINKS")
            n = len(dich[f]) + 1
            c = Counter()
            ghi = []
            for r in dich[f]:
                if r in moi:
                    ghi.append((moi[r],)); c["đổi"] += 1
                elif r and r.lower() in ten_gui:
                    ghi.append((r,)); c["file gửi kèm"] += 1
                else:
                    ghi.append((r,)); c["không có trong cây tender" if r else "trống"] += 1
            if n > 1:
                ws.Range(f"E2:E{n}").Value = ghi
            if ws.Range("H1").Value == "BASE_DIR":
                ws.Range("H2").Value = ""
            wb.Saved = False
            wb.Save()
            wb.Close(False)
            dem[out] = c
    finally:
        xl.Quit()

    # 4. đọc lại file đã lưu
    for out, c in dem.items():
        rels, base = doc_links(out)
        con_cu = sum(1 for r in rels if r and r in moi)
        trung = sum(1 for r in rels if r in set(moi.values()))
        print(f"Đã ghi {out}: {len(rels)} dòng link — " + ", ".join(f"{k} {v}" for k, v in c.items())
              + f"; đọc lại: {trung} dòng đường dẫn mới, {con_cu} dòng còn đường dẫn cũ, BASE_DIR {'trống' if not base else base}")
        if con_cu:
            print(f"LỖI: {out} còn {con_cu} dòng đường dẫn cũ — Excel không lưu được")
    for r in khong_co:
        print("XEM không có trong cây tender:", r)
    for x in cay.khong_doc:
        print("XEM không đọc được file nén:", x)
    goc = sorted({v.split(BS)[0] for v in moi.values()})
    print(f"Tender đặt {len(gui)} file vào GỐC cây hồ sơ (cạnh: {' | '.join(goc)}), để chung 1 chỗ"
          + (" (BOQ_TRACK link sang DMVT)." if len(gui) > 1 else "."))
    print("  OneDrive đã đồng bộ về máy: đặt ở thư mục con nào trong cây cũng được; mở trên web chưa đồng bộ thì phải ở gốc.")
    for nen, k in sorted(can_giai.items()):
        print(f"  Giải nén '{nen}' ra thư mục cùng tên (Extract to {os.path.splitext(nen)[0]}\\): {k} file đích nằm trong đó.")
    print("  Máy tender cần add-in SRC_LINKS (_MASTER\\vba\\GUI_CHO_NGUOI_KHAC). Skill không tự up, không tự gửi.")
    return ra
