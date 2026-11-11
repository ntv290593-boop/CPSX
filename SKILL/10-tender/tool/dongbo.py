"""dongbo.py — chuyển trạng thái quét (_data) giữa 2 máy, không chép hồ sơ/xlsx.

  td.py dongbo xuat <GOC> [--ra <file.zip>]        gói <GOC>\\10_TENDER\\_data -> zip nhỏ (mặc định Downloads\\<DA>_DATA_<yymmdd-hhmm>.zip)
  td.py dongbo nhap <GOC> <file.zip> [--ghi-de]    giải vào _data máy này (máy chưa có _data thì tạo), rồi chạy kiemke để biết file đổi

Hai máy phải có cùng hồ sơ ở cùng đường dẫn tương đối (khóa file = đường dẫn tính từ thư mục input).
nhap từ chối khi _data máy này đã có sửa chưa xuất (so với lần xuat/nhap trước); --ghi-de: bỏ qua, bản cũ lưu zip rồi mới ghi.
"""
import os, sys, json, zipfile, hashlib, datetime, subprocess, socket

BO_THU_MUC = {"_zip"}          # file giải nén từ zip hồ sơ — dựng lại được, có thể nặng
BO_DUOI = (".tmp",)
STATE = "_dongbo_state.json"    # {tên file: md5} lúc xuat/nhap gần nhất — để biết máy này có sửa gì sau đó không
MANIFEST = "_dongbo.json"


def _md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def _files(d):
    kq = {}
    for root, dirs, fs in os.walk(d):
        rel = os.path.relpath(root, d).replace("\\", "/")
        dirs[:] = [x for x in dirs if not (rel == "." and x in BO_THU_MUC)]
        for fn in fs:
            if fn.endswith(BO_DUOI) or fn in (STATE, MANIFEST):
                continue
            r = fn if rel == "." else f"{rel}/{fn}"
            kq[r] = os.path.join(root, fn)
    return kq


def _git():
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        return subprocess.run(["git", "-C", here, "rev-parse", "--short", "HEAD"], capture_output=True, text=True, timeout=10).stdout.strip()
    except Exception:
        return ""


def _lech(d):
    """File _data máy này đã sửa/thêm/xóa so với lần xuat/nhap trước -> danh sách; None = chưa có mốc."""
    sp = os.path.join(d, STATE)
    if not os.path.exists(sp):
        return None
    st = json.load(open(sp, encoding="utf-8"))
    cur = _files(d)
    out = [f"sửa/mới: {r}" for r, p in cur.items() if st.get(r) != _md5(p)]
    out += [f"mất: {r}" for r in st if r not in cur]
    return out


def _ghi_state(d):
    with open(os.path.join(d, STATE), "w", encoding="utf-8") as fh:
        json.dump({r: _md5(p) for r, p in _files(d).items()}, fh, ensure_ascii=False, indent=0)


def xuat(p, args, opt):
    d = p.d
    if not os.path.isdir(d):
        sys.exit(f"LỖI: chưa có {d}")
    if not os.path.exists(os.path.join(d, "_chuky.json")):
        print("XEM: chưa có _chuky.json — chưa chạy kiemke, máy nhận sẽ không biết file nào đổi")
    fs = _files(d)
    now = datetime.datetime.now()
    da = (p.cfg.get("du_an") or os.path.basename(p.goc)).upper()
    ra = opt(args, "--ra") or os.path.join(os.path.expanduser("~"), "Downloads", f"{da}_DATA_{now:%y%m%d-%H%M}.zip")
    if os.path.abspath(ra).startswith(os.path.abspath(p.goc) + os.sep):
        sys.exit("LỖI: --ra không được nằm trong thư mục dự án (zip sẽ bị gói lại ở lần sau)")
    os.makedirs(os.path.dirname(os.path.abspath(ra)), exist_ok=True)
    fip = os.path.join(d, "FILE_INDEX.kiemke-tool.json")
    fi = json.load(open(fip, encoding="utf-8")).get("rows", []) if os.path.exists(fip) else []
    man = {"ngay": now.strftime("%Y-%m-%d %H:%M"), "may": socket.gethostname(), "du_an": da, "skill": _git(),
           "input": p.cfg.get("input"), "so_file": len(fs), "so_file_hoso": len(fi)}
    tmp = ra + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(MANIFEST, json.dumps(man, ensure_ascii=False, indent=1))
        for r, pth in fs.items():
            z.write(pth, r)
    os.replace(tmp, ra)
    _ghi_state(d)
    print(f"Đã gói {len(fs)} file của _data -> {ra} ({os.path.getsize(ra) / 1e6:.1f} MB)")
    print(f"Skill {man['skill'] or '(không đọc được git)'}, máy {man['may']}. Máy nhận: pull skill cùng bản, `dongbo nhap`, rồi `kiemke`.")


def nhap(goc, args, opt):
    if not args or not os.path.isfile(args[0]):
        sys.exit("LỖI: dongbo nhap <GOC> <file.zip> [--ghi-de]")
    zp = args[0]
    t = os.path.join(os.path.abspath(goc), "10_TENDER")
    d = os.path.join(t, "_data")
    with zipfile.ZipFile(zp) as z:
        if MANIFEST not in z.namelist():
            sys.exit("LỖI: zip không có _dongbo.json — không phải gói do `dongbo xuat` tạo")
        man = json.loads(z.read(MANIFEST))
        for n in z.namelist():  # chống zip-slip
            if os.path.isabs(n) or ".." in n.replace("\\", "/").split("/"):
                sys.exit(f"LỖI: tên file không an toàn trong zip: {n}")
        co = os.path.isdir(d) and bool(_files(d))
        if co:
            lech = _lech(d)
            if "--ghi-de" not in args:
                if lech is None:
                    sys.exit(f"LỖI: {d} đã có dữ liệu nhưng chưa từng xuat/nhap nên không biết có sửa gì chưa gửi. "
                             f"Nếu chắc chắn bản trong zip mới hơn: thêm --ghi-de (bản cũ được lưu zip trước).")
                if lech:
                    print("Máy này đã có thay đổi chưa xuất:")
                    for x in lech[:20]:
                        print("  ", x)
                    sys.exit(f"LỖI: {len(lech)} thay đổi sẽ bị mất. `dongbo xuat` trước rồi ghép tay, hoặc thêm --ghi-de.")
            bk = os.path.join(t, f"_data_truoc_nhap_{datetime.datetime.now():%y%m%d-%H%M%S}.zip")
            with zipfile.ZipFile(bk, "w", zipfile.ZIP_DEFLATED) as b:
                for r, pth in _files(d).items():
                    b.write(pth, r)
            print(f"Bản _data cũ lưu ở {bk}")
            ten = set(z.namelist())
            for r, pth in _files(d).items():  # zip là bản đúng: file không có trong zip thì bỏ
                if r not in ten:
                    os.remove(pth)
        os.makedirs(d, exist_ok=True)
        for n in z.namelist():
            if n != MANIFEST and not n.endswith("/"):
                out = os.path.join(d, n)
                os.makedirs(os.path.dirname(out), exist_ok=True)
                with z.open(n) as src, open(out + ".tmp", "wb") as fh:
                    fh.write(src.read())
                os.replace(out + ".tmp", out)
    _ghi_state(d)
    print(f"Đã nhập {man.get('so_file')} file từ {man.get('may')} ({man.get('ngay')}), dự án {man.get('du_an')}.")
    here = _git()
    if man.get("skill") and here and man["skill"] != here:
        print(f"XEM: skill máy này {here} ≠ máy gửi {man['skill']} — pull cho cùng bản trước khi chạy tiếp")
    cf = os.path.join(d, "_cauhinh.json")
    cfg = json.load(open(cf, encoding="utf-8")) if os.path.exists(cf) else None
    fi = os.path.join(d, "FILE_INDEX.kiemke-tool.json")
    if cfg and os.path.exists(fi):
        inp = os.path.join(os.path.abspath(goc), cfg["input"])
        rows = json.load(open(fi, encoding="utf-8")).get("rows", [])
        thieu = sorted({r["Đường dẫn"] for r in rows if not r.get("Trong zip") and not r["Đường dẫn"].startswith("http")
                        and not os.path.exists(os.path.join(inp, r["Đường dẫn"]))})
        if thieu:
            print(f"XEM: {len(thieu)} file hồ sơ trong FILE_INDEX không có ở {inp} (khác cây thư mục hoặc chưa tải) — kiemke sẽ coi là MẤT. Ví dụ:")
            for x in thieu[:8]:
                print("  ", x)
        else:
            print(f"Hồ sơ máy này khớp cây thư mục của FILE_INDEX ({len(rows)} dòng).")
    print("-> Chạy `td.py kiemke <GOC>` (đợt mới thì thêm --dot-moi), rồi `thaydoi`.")


def main(goc, args, opt, P):
    if not args or args[0] not in ("xuat", "nhap"):
        sys.exit("LỖI: dongbo xuat <GOC> [--ra f.zip] | dongbo nhap <GOC> <f.zip> [--ghi-de]")
    if args[0] == "xuat":
        return xuat(P(goc).need(), args[1:], opt)
    return nhap(goc, args[1:], opt)
