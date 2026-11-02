"""Trích tag (block có thuộc tính) và nhãn chữ từ DWG, theo từng tờ (layout). Dùng chung mọi dự án có DWG.

Chạy bằng python có ezdxf (_MASTER\\tools\\.venv):
  python -I dwg_tag.py <thư mục DWG> <file json ra> [--tam <thư mục tạm>] [--may 6] [--giu-dxf]
Tool sao DWG vào thư mục tạm, đổi sang DXF bằng ODA File Converter, đọc, rồi xóa DXF. Không ghi gì vào thư mục hồ sơ.

Kết quả: {"<đường dẫn tương đối của DWG>": {"to": {"<tên layout>": {"tag": [{"b": block, "a": {tag: chữ}, "xy": [x, y]}], "chu": [...]}},
          "model": {"tag": [...], "chu": [...]}}}
Tag trong block lồng (xref đã bind) của Model cũng được lấy, xy là tọa độ trong block đó.
"""
import json
import multiprocessing as mp
import os
import re
import shutil
import subprocess
import sys
import time

ODA = r"C:\Program Files\ODA\ODAFileConverter 27.9.0\ODAFileConverter.exe"
RX_MTEXT = re.compile(r"\\[A-Za-z][^;\\{}]*;|\\[PpNnLlOoKk~]|[{}]")


def chu_sach(t):
    return re.sub(r"\s+", " ", RX_MTEXT.sub(" ", t or "")).strip()


def doc_khong_gian(sp):
    tag, chu = [], []
    for e in sp:
        k = e.dxftype()
        if k == "INSERT":
            a = {x.dxf.tag: x.dxf.text.strip() for x in e.attribs if x.dxf.text.strip()}
            if a:
                tag.append({"b": e.dxf.name, "a": a, "xy": [round(e.dxf.insert[0], 1), round(e.dxf.insert[1], 1)]})
        elif k == "MULTILEADER":  # tag mặt đứng: chữ nằm trong block_attribs theo thứ tự, không có tên thuộc tính
            try:
                v = [x.text.strip() for x in e.block_attribs if x.text and x.text.strip()]
                c = e.context
                if not v and getattr(c, "mtext", None):
                    v = [chu_sach(c.mtext.default_content)]
                if v:
                    d = None
                    for ten in ("base_point", "plane_origin"):
                        p = getattr(c, ten, None)
                        if p is not None:
                            d = [round(p[0], 1), round(p[1], 1)]
                            break
                    mui = []  # đầu mũi tên: điểm tag chỉ vào
                    for ld in getattr(c, "leaders", []) or []:
                        for ln in getattr(ld, "lines", []) or []:
                            vs = getattr(ln, "vertices", None)
                            if vs:
                                mui.append([round(vs[0][0], 1), round(vs[0][1], 1)])
                    tag.append({"b": "MULTILEADER", "ml": v, "a": {}, "xy": d, "mui": mui})
            except AttributeError:
                pass
        elif k in ("TEXT", "MTEXT"):
            t = chu_sach(e.dxf.text if k == "TEXT" else e.text)
            if t and len(t) <= 120:
                chu.append(t)
    return tag, chu


def doc_dxf(args):
    f, rel = args
    import ezdxf
    try:
        doc = ezdxf.readfile(f)
    except Exception as e:
        return rel, {"loi": str(e)[:200]}
    out = {"to": {}}
    for lay in doc.layouts:
        tag, chu = doc_khong_gian(lay)
        da, hang = set(), [e.dxf.name for e in lay if e.dxftype() == "INSERT"]
        while hang:  # đi xuống các block được chèn (tag mặt đứng nằm trong block lồng ở paper space), mỗi block 1 lần
            n = hang.pop()
            if n in da or n not in doc.blocks or "TBLK" in n.upper():
                continue
            da.add(n)
            t2, c2 = doc_khong_gian(doc.blocks[n])
            for x in t2:
                x["long"] = n[:60]
            tag += t2
            chu += c2
            hang += [e.dxf.name for e in doc.blocks[n] if e.dxftype() == "INSERT"]
        if lay.name == "Model":
            out["model"] = {"tag": tag, "chu": chu}
        else:
            out["to"][lay.name] = {"tag": tag, "chu": chu}
            for e in lay:  # gốc tọa độ giấy của tờ có thể lệch: lưu vị trí khung tên để đổi sang tọa độ trang
                if e.dxftype() == "INSERT" and "TBLK" in e.dxf.name.upper():
                    out["to"][lay.name]["goc"] = [round(e.dxf.insert[0], 1), round(e.dxf.insert[1], 1), round(e.dxf.get("xscale", 1), 3)]
                    break
    # Tag ở Model space dùng chung nhiều tờ: gán về đúng tờ theo khung viewport. Tag nằm ngoài mọi khung thì không hiện trên tờ nào.
    khung = {}
    for lay in doc.layouts:
        if lay.name == "Model":
            continue
        for e in lay:
            if e.dxftype() == "VIEWPORT" and e.dxf.get("id", 2) != 1 and e.dxf.get("height", 0) > 0:
                cx, cy = e.dxf.view_center_point[0], e.dxf.view_center_point[1]
                h = e.dxf.view_height
                w = h * e.dxf.width / e.dxf.height
                k = e.dxf.height / h  # tỉ lệ model -> giấy
                px, py = e.dxf.center[0], e.dxf.center[1]
                khung.setdefault(lay.name, []).append((cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2, cx, cy, k, px, py))
                out["to"][lay.name].setdefault("vp", []).append([round(px - e.dxf.width / 2, 1), round(py - e.dxf.height / 2, 1), round(px + e.dxf.width / 2, 1), round(py + e.dxf.height / 2, 1)])
    for t in out["model"]["tag"]:
        if t.get("long") or not t.get("xy"):
            continue  # block lồng: xy là tọa độ trong block, không xét được
        x, y = t["xy"]
        trung = []
        for n, ds in khung.items():
            for a, b, c, d, cx, cy, k, px, py in ds:
                if a <= x <= c and b <= y <= d:
                    trung.append(n)
                    q = dict(t, tu_model=True)
                    q["xy"] = [round(px + (x - cx) * k, 1), round(py + (y - cy) * k, 1)]  # đổi sang tọa độ giấy của tờ
                    if t.get("mui"):
                        q["mui"] = [[round(px + (u - cx) * k, 1), round(py + (v - cy) * k, 1)] for u, v in t["mui"]]
                    out["to"][n]["tag"].append(q)
                    break
        t["gan"] = trung
    return rel, out


def main(a):
    src, ra = a[0], a[1]
    def opt(n, d=None):
        return a[a.index(n) + 1] if n in a else d
    tam = opt("--tam") or os.path.join(os.path.dirname(os.path.abspath(ra)), "_dwg_tam")
    may = int(opt("--may", 6))
    d_in, d_out = os.path.join(tam, "in"), os.path.join(tam, "out")
    dem_dxf = lambda: sum(x.lower().endswith(".dxf") for _, _, fs in os.walk(d_out) for x in fs)
    tong = sum(x.lower().endswith(".dwg") for _, _, fs in os.walk(src) for x in fs)
    if not os.path.isdir(d_out) or dem_dxf() < tong:
        shutil.copytree(src, d_in, dirs_exist_ok=True, ignore=lambda d, fs: [x for x in fs if not os.path.isdir(os.path.join(d, x)) and not x.lower().endswith(".dwg")])
        os.makedirs(d_out, exist_ok=True)
        print(f"Đổi {tong} DWG sang DXF…", flush=True)
        oda = subprocess.Popen([ODA, d_in, d_out, "ACAD2018", "DXF", "1", "1", "*.dwg"])
        while oda.poll() is None and dem_dxf() < tong:  # tờ nặng đổi lâu vài phút: chờ theo tiến trình, không theo thời gian
            time.sleep(10)
        time.sleep(20)  # file cuối còn đang ghi
        print(f"Có {dem_dxf()}/{tong} DXF", flush=True)
    if dem_dxf() < tong:
        sys.exit(f"LỖI: ODA chỉ đổi được {dem_dxf()}/{tong} DWG. Không đọc tiếp để khỏi thiếu tag.")
    viec = []
    for r, _, fs in os.walk(d_out):
        for x in fs:
            if x.lower().endswith(".dxf"):
                viec.append((os.path.join(r, x), os.path.relpath(os.path.join(r, x), d_out).replace("\\", "/")[:-4] + ".dwg"))
    kq = {}
    with mp.Pool(may) as pool:
        for i, (rel, o) in enumerate(pool.imap_unordered(doc_dxf, viec), 1):
            kq[rel] = o
            if i % 20 == 0:
                print(f"  đọc {i}/{len(viec)}", flush=True)
    json.dump(kq, open(ra, "w", encoding="utf-8"), ensure_ascii=False)
    print(f"Đã ghi {ra}: {len(kq)} DWG, lỗi {sum('loi' in v for v in kq.values())}")
    if "--giu-dxf" not in a:
        shutil.rmtree(tam, ignore_errors=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1:])
