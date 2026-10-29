"""Danh mục module + ô nhập của từng block mẫu (Template/Door/Kitchen) trong 1 file CPSX.

Quét đúng file của dự án (thư viện mẫu mỗi file có thể khác RE00). Block = dòng A có chữ và B hoặc E là số
(giống libBlocks trong cpsx_lib.js) đến dòng A="END". Ô nhập = ô hằng số (không phải công thức) ở B..R.
Nhãn `Lan` (#k) đếm giống applyEdits: norm(A).startswith(norm(nhãn)).

Dùng: py -I module_catalog.py <file CPSX> <json ra> [RULES_MASTER.xlsx]
"""
import sys, json, re, datetime
import openpyxl

LIB = ["Template", "Door", "Kitchen"]
COLS = "BCDEFGHIJKLMNOPQR"
NHAN_CONG = "chi phí nhân công"   # từ dòng này tới END: khóa (R0150)


def norm(v):
    return re.sub(r"\s+", " ", "" if v is None else str(v)).strip().lower()


def name_key(v):
    """So tên block với RULES: bỏ khoảng trắng quanh ngoặc, bỏ phần ngoặc ghi chú ("Vách ghép hộp ( Vách tivi..)")."""
    s = re.sub(r"\s*\(\s*", "(", norm(v)).replace(" )", ")")
    return s, re.sub(r"\(.*$", "", s).strip()


def rule_ids(rules_path):
    """[Template#1] Tủ bếp trên (~10–91) → {('Template','tủ bếp trên'): 'Template#1'}"""
    if not rules_path:
        return {}
    ws = openpyxl.load_workbook(rules_path, read_only=True)["RULES"]
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        m = re.match(r"\[(Template|Door|Kitchen)#(\d+)\]\s*(.+?)\s*\(~", str(r[3] or ""))
        if r[2] == "MAU" and m:
            out[(m[1], name_key(m[3])[0])] = {"id": f"{m[1]}#{m[2]}", "rule": r[0], "status": r[5]}
    return out


def find_id(ids, sn, name):
    full, short = name_key(name)
    if (sn, full) in ids:
        return ids[(sn, full)]
    hits = [v for (s, k), v in ids.items() if s == sn and k == short]
    return hits[0] if len(hits) == 1 else None


def scan(path):
    wf = openpyxl.load_workbook(path, read_only=True)                  # công thức
    wv = openpyxl.load_workbook(path, read_only=True, data_only=True)  # giá trị đã tính (nhãn A là công thức)
    out = {}
    for sn in LIB:
        if sn not in wf.sheetnames:
            continue
        F = list(wf[sn].iter_rows(min_col=1, max_col=18, values_only=True))
        V = list(wv[sn].iter_rows(min_col=1, max_col=18, values_only=True))
        blocks, cur = [], None
        for i, (f, v) in enumerate(zip(F, V)):
            lab = norm(v[0])
            if lab == "end":
                if cur:
                    cur["end"] = i + 1
                    blocks.append(cur)
                    cur = None
                continue
            if not cur and lab and (isinstance(v[1], (int, float)) or isinstance(v[4], (int, float))):
                cur = {"name": str(v[0]).strip(), "hdr": i + 1, "rows": []}
            if not cur:
                continue
            off = i + 1 - cur["hdr"]
            locked = any(norm(r["A"]).startswith(NHAN_CONG) for r in cur["rows"]) or lab.startswith(NHAN_CONG)
            inputs = {}
            for j, c in enumerate(COLS, start=1):
                x = f[j]
                if x not in (None, "") and not (isinstance(x, str) and x.startswith("=")):
                    inputs[c] = x if isinstance(x, (int, float, str)) else str(x)
            cur["rows"].append({"off": off, "A": "" if v[0] is None else str(v[0]).strip(),
                                "A_cong_thuc": isinstance(f[0], str) and f[0].startswith("="),
                                "inputs": inputs, "khoa": locked})
        for b in blocks:
            hdr = b["rows"][0]
            b["gia_theo"] = hdr["inputs"].get("M") if isinstance(hdr["inputs"].get("M"), str) else None
            b["nhom"] = hdr["inputs"].get("G")
            # 2 dòng "GIÁ GỐC NHÀ MÁY" trong 1 block = thiếu END ở block trước → chép block sẽ dài gấp đôi
            gg = [b["hdr"] + r["off"] for r in b["rows"] if norm(r["A"]) == "giá gốc nhà máy"]
            if len(gg) > 1:
                b["loi"] = f"Thiếu END sau dòng {gg[0]} (block gộp 2 mẫu)"
            # Lan (#k) và nhãn trùng phần đầu (applyEdits chọn nhầm hàng)
            labs = [norm(r["A"]) for r in b["rows"]]
            for idx, r in enumerate(b["rows"]):
                t = labs[idx]
                if not t:
                    continue
                r["lan"] = sum(1 for x in labs[:idx + 1] if x.startswith(t))
                lan_exact = sum(1 for x in labs[:idx + 1] if x == t)
                if r["lan"] != lan_exact:
                    r["trung_dau"] = [b["rows"][k]["A"] for k in range(idx) if labs[k].startswith(t) and labs[k] != t]
        out[sn] = blocks
    return out


def main():
    src, dst = sys.argv[1], sys.argv[2]
    ids = rule_ids(sys.argv[3] if len(sys.argv) > 3 else None)
    lib = scan(src)
    seen = set()
    for sn, blocks in lib.items():
        for b in blocks:
            hit = find_id(ids, sn, b["name"])
            b["mau"] = hit["id"] if hit else None
            b["rule"] = hit["rule"] if hit else None
            if hit:
                seen.add(hit["id"])
    missing = sorted({v["id"] for v in ids.values()} - seen, key=lambda s: (s.split("#")[0], int(s.split("#")[1])))
    doc = {"file": src, "ngay": datetime.date.today().isoformat(), "thieu_so_voi_rules": missing, "lib": lib}
    with open(dst, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
    sys.stdout.reconfigure(encoding="utf-8")
    for sn, blocks in lib.items():
        print(f"== {sn}: {len(blocks)} block")
        for b in blocks:
            nin = sum(len(r["inputs"]) for r in b["rows"] if not r["khoa"])
            tr = sum(1 for r in b["rows"] if r.get("trung_dau"))
            print(f"  {b['mau'] or '(không có trong RULES)':14} {b['hdr']:>5}-{b['end']:<5} {b['name'][:45]:45} giá theo={b['gia_theo'] or '-':4} ô nhập={nin:<4} trùng đầu={tr} {b.get('loi', '')}")
    print("Có trong RULES nhưng không có trong file:", ", ".join(missing) or "không")


if __name__ == "__main__":
    main()
