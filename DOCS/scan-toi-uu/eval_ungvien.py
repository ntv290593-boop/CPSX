# -*- coding: utf-8 -*-
"""Đo lại recall của ungvien trên Blanca. Chỉ đọc _data, không ghi gì vào dự án."""
import sys, os, re, json, glob, collections
sys.path.insert(0, r"D:\COSTING EXCEL\SKILL\01-scan\tool")
import pd

GOC, DOT = r"D:\COSTING EXCEL\_ACTIVE\Blanca_City", "261027_DV2.3-DV2.4"
P = pd.paths(GOC, DOT)
sheets = pd.load_json(P["data"])[0]


def merged_boq():
    m = collections.OrderedDict()
    for _, r in sheets["BOQ_TONGHOP"]:
        m.setdefault(str(r["Mã"]), {}).update({k: v for k, v in r.items() if v not in (None, "")})
    return m


def gold():
    """Mã -> set (fid, trang) từ 'Bản vẽ đối chứng' của bước banve."""
    g = collections.defaultdict(set)
    for b, r in sheets["BOQ_TONGHOP"]:
        if not b.startswith("banve"): continue
        v = r.get("Bản vẽ đối chứng") or []
        if isinstance(v, dict): v = v.get("links", [])
        if isinstance(v, str): continue
        for l in v:
            if isinstance(l, dict): l = [l.get("nhan"), l.get("file"), l.get("trang")]
            if isinstance(l, list) and len(l) >= 3 and l[2] and re.fullmatch(r"F\d+", str(l[1])):
                g[str(r["Mã"])].add((l[1], int(l[2])))
    return g


texts = {}
for f in sorted(glob.glob(os.path.join(P["data"], "_text", "F*.json"))):
    texts[os.path.basename(f)[:-5]] = json.load(open(f, encoding="utf-8"))
raw_texts = texts


def ungvien_current(boq):
    """Chép đúng logic cmd_ungvien -> Mã -> list (fid, trang) xếp theo điểm."""
    ft = {fid: [pd._fold(t) for t in pages] for fid, pages in texts.items()}
    out = {}
    for ma, r in boq.items():
        hm = r.get("Hạng mục") or ""
        kws = sorted({w for w in re.findall(r"[a-z0-9]{3,}", pd._fold(hm)) if w not in pd.STOP})
        mrx = re.compile(r"(?<![a-z0-9])" + re.escape(pd._fold(ma)) + r"(?![a-z0-9])") if ma else None
        res = []
        for fid, pages in ft.items():
            for i, t in enumerate(pages, 1):
                mh = len(mrx.findall(t)) if mrx else 0
                kh = [w for w in kws if w in t]
                if mh or len(kh) >= min(2, len(kws)) and kh:
                    res.append((mh * 5 + len(kh), fid, i, mh))
        res.sort(key=lambda x: -x[0])  # sort ổn định: hòa điểm giữ thứ tự file/trang
        out[ma] = res
    return out


def report(name, ranked, g, ks=(5, 12, 20, 40)):
    tot = sum(len(v) for v in g.values())
    inc = sum(1 for m, s in g.items() for p in s if p in {(x[1], x[2]) for x in ranked.get(m, [])})
    line = f"{name:28s} mã={len(g)} trang={tot}  trong ứng viên {inc}/{tot} ({inc/tot:.0%})"
    for k in ks:
        h = sum(1 for m, s in g.items() for p in s if p in {(x[1], x[2]) for x in ranked.get(m, [])[:k]})
        line += f"  @{k} {h/tot:.0%}"
    # mã có ít nhất 1 trang đúng trong top12
    a12 = sum(1 for m, s in g.items() if s & {(x[1], x[2]) for x in ranked.get(m, [])[:12]})
    line += f"  | mã có ≥1 trang đúng @12: {a12}/{len(g)}"
    sizes = sorted(len(ranked.get(m, [])) for m in g)
    line += f"  | trung vị ứng viên {sizes[len(sizes)//2]} max {sizes[-1]}"
    print(line)


if __name__ == "__main__":
    boq = merged_boq(); g = gold()
    print("BOQ mã:", len(boq), " mã có đáp án:", len(g), " trang đáp án:", sum(map(len, g.values())))
    allc = ungvien_current(boq)
    sizes = sorted(len(v) for v in allc.values())
    print("toàn bộ 148 mã: trung vị ứng viên", sizes[len(sizes)//2], "max", sizes[-1])
    by_code = sum(1 for m, v in allc.items() if any(x[3] for x in v))
    print("mã khớp theo chính mã:", by_code, "/", len(allc))
    report("ungvien hiện tại", allc, g)
    # đáp án gộp F023≡F024 (hai file trùng 42 trang)
    alias = lambda p: ("F024", p[1]) if p[0] == "F023" else p
    g2 = {m: {alias(p) for p in s} for m, s in g.items()}
    allc2 = {m: [(x[0], *alias((x[1], x[2])), x[3]) for x in v] for m, v in allc.items()}
    report("  (gộp F023≡F024)", allc2, g2)
