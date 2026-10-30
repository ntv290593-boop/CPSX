# -*- coding: utf-8 -*-
"""Bài H (bản N1, 30/10/2026): kiểm chuỗi kích thước theo THỨ TỰ trên trang PDF.
Sửa so với bản 29/10:
 - Ghép từ: get_text("words") gom theo (block, line), sắp theo x; tách ô khi khoảng trống > 0,6 × cao chữ
   (UNRULED_WORD_GAP của matchline). Chuỗi kích thước không vượt qua ranh giới ô.
 - Cùng 1 bộ tách số cho trang và cho giá trị cần kiểm (bản cũ tách chuỗi khác cách tách dòng số → 157 ca
   tự không khớp, đều là `Rev.00 -26/06/2026`).
 - Lọc rác: ngày, tỉ lệ 1:50 / 1/100, số điện thoại, mã dính chữ (IN-PS2-501, EN2-4, #1-02), số có 0 đứng đầu.
 - Dấu `-` chỉ là dấu chuỗi khi có đơn vị mm/cm (5–10mm). Bản cũ nhận mọi `a-b` nên 99% "chuỗi" là địa chỉ,
   mã hạng mục, tiêu chuẩn.
 - Tái lập: mọi tập đều sort trước khi random. Lưu ý `py -I` bỏ qua PYTHONHASHSEED (-I bao gồm -E).
Chạy: py -I -X utf8 neo_test7.py   (đọc PDF trong zip SOFITEL 10_TENDER; chỉ đọc)"""
src = open(r"D:\COSTING EXCEL\DOCS\scan-toi-uu\neo-test\neo_test5.py", encoding='utf-8').read().split("stat=collections")[0]
exec(src)
import re, os, random, collections, unicodedata, pymupdf  # đã có từ exec, khai lại cho rõ

GAP = 0.6
FOLD2 = {"×": "x", "X": "x", "–": "-", "—": "-", "−": "-", "＋": "+", "\u00a0": " "}

# rác: thay bằng ranh giới ô trước khi tách số
JUNK = [
    re.compile(r"\b\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4}\b"),                    # ngày
    re.compile(r"\b\d+\s*[:/]\s*\d+\b"),                                     # tỉ lệ, phân số, giờ
    re.compile(r"\+\s*\(?\d{1,3}\)?[\s.\-]?\(?\d{1,4}\)?(?:[\s.\-]?\d{2,4}){2,}"),  # điện thoại có +
    re.compile(r"\(\d{2,4}\)\s*\d{3,4}[\s.\-]\d{3,4}"),                      # điện thoại (028) 3824 1525
    re.compile(r"#\s*\d[\d\-]*"),                                            # #1-02
    re.compile(r"\b[A-Za-z][\w]*[\-.]\w*\d[\w\-.]*"),                        # mã dính chữ: IN-PS2-501, EN2-4, LD07755.LD
    re.compile(r"\b\d+(?:\.\d+){2,}\b"),                                     # 931.84.233
    re.compile(r"(?i)\brev\.?\s*\d+"),
]
DINH = r"0-9A-Za-wyz_À-ỹ"                   # ký tự "dính" số làm số thành mã; trừ x để 1200x600 vẫn tách được
PRE = r"(?<![" + DINH + r".,#/:])(?:[WDHLRS]|Ø)?"   # 1 chữ chiều đứng trước: W1200, D600
SO = r"\d+(?:\.\d+)?"
DV = r"(?:\s?(?:mm|cm))"
END = r"(?![" + DINH + r"#/:]|\.\d)"
NUM_RE = re.compile(PRE + "(" + SO + ")" + DV + "?" + END)
N0 = PRE + SO + DV + "?" + END
CHAIN_RE = re.compile(N0 + r"(?:\s*[x*+=]\s*" + N0 + r")+"            # 1200x600, 450+450=900
                      + "|" + PRE + SO + r"\s*-\s*" + PRE + SO + DV + END)  # dải 5-10mm: bắt buộc có đơn vị


def chuan(s):
    s = unicodedata.normalize("NFKC", str(s))
    s = "".join(FOLD2.get(c, c) for c in s)
    s = re.sub(r"(?<=\d),(?=\d{3}(?!\d))", "", s)                 # 1,200 -> 1200
    s = re.sub(r"(?<![\d.])(\d{1,3})\.(\d{3})(?![.\d])", r"\1\2", s)  # 1.200 -> 1200 (1 nhóm nghìn)
    for j in JUNK: s = j.sub(" ¦ ", s)
    return s


def so(tok):
    """'1200' / '1200.0' / '012' -> chuẩn; None nếu là rác (0 đứng đầu)."""
    if len(tok) > 1 and tok[0] == "0" and not tok.startswith("0."): return None
    return tok[:-2] if tok.endswith(".0") else tok


def tach_so(s):
    """Dòng số sạch theo thứ tự đọc."""
    out = []
    for cell in chuan(s).split("¦"):
        for m in NUM_RE.finditer(cell):
            v = so(m.group(1))
            if v: out.append(v)
    return out


def tach_chuoi(s):
    """Các chuỗi kích thước (>= 2 số) trong 1 ô/1 giá trị."""
    out = []
    for cell in chuan(s).split("¦"):
        for m in CHAIN_RE.finditer(cell):
            p = [so(x) for x in NUM_RE.findall(m.group(0))]
            if len(p) >= 2 and all(p): out.append((m.group(0).strip(), p))
    return out


def o_trang(words):
    """words -> danh sách ô (chuỗi) theo thứ tự block, line, x."""
    by = collections.defaultdict(list)
    for w in words: by[(w[5], w[6])].append(w)
    cells = []
    for k in sorted(by):
        ln = sorted(by[k], key=lambda w: w[0]); cur = [ln[0]]
        for a, b in zip(ln, ln[1:]):
            if b[0] - a[2] > GAP * max(a[3] - a[1], b[3] - b[1], 1e-3):
                cells.append(" ".join(w[4] for w in cur)); cur = []
            cur.append(b)
        cells.append(" ".join(w[4] for w in cur))
    return cells


def ord_ok(parts, nums, slack=3):
    n = len(nums)
    for i, v in enumerate(nums):
        if v != parts[0]: continue
        j = i; ok = True
        for p in parts[1:]:
            k = next((q for q in range(j + 1, min(n, j + 1 + slack)) if nums[q] == p), None)
            if k is None: ok = False; break
            j = k
        if ok: return True
    return False


def set_ok(parts, nums):
    S = set(nums); return all(p in S for p in parts)


if __name__ == "__main__":
    # ---- 0. bộ tách tự kiểm (ca đã gặp + dạng LLM hay viết)
    T = [("1,200 x 600 mm", ["1200", "600"]), ("1.200x600", ["1200", "600"]), ("W1200 x D600 x H750", ["1200", "600", "750"]),
         ("450+450=900", ["450", "450", "900"]), ("5–10mm", ["5", "10"]), ("1200mm x 600mm", ["1200", "600"]),
         ("Rev.00 -26/06/2026", None), ("30-31 Lyme Street", None), ("Item 53 - 700", None), ("IN-PS2-501", None),
         ("#1-02 One KL", None), ("931.84.233 - 12.06.2026", None), ("ISO 1452-2", None), ("+1(323)213-4626", None),
         ("1:50", None), ("W10 x H10", ["10", "10"]), ("600 × 600", ["600", "600"]), ("2.5 x 1200", ["2.5", "1200"])]
    bad = 0
    for s, exp in T:
        got = tach_chuoi(s); g = got[0][1] if got else None
        if g != exp: bad += 1; print("  SAI bộ tách:", repr(s), "->", got, "kỳ vọng", exp)
    print(f"0. bộ tách tự kiểm: {len(T) - bad}/{len(T)} đúng")

    random.seed(5)
    pages = []
    for r, sz in pick:
        zp = os.path.join(INP, r["Đường dẫn"])
        try: d = pymupdf.open(stream=zs[zp].read(r["Trong zip"]), filetype="pdf")
        except Exception: continue
        for i in range(d.page_count):
            cells = o_trang(d[i].get_text("words"))
            nums = tach_so(" ¦ ".join(cells))
            ch = [c for cell in cells for c in tach_chuoi(cell)]
            if ch and len(nums) >= 20: pages.append((r["ID"], i + 1, cells, nums, ch))
    print("1. trang có chuỗi kích thước & >= 20 số sạch:", len(pages), "(quét hết", len(pick), "file mẫu, không dừng ở 400)")

    real = [(fid, pg, txt, p, nums) for fid, pg, cells, nums, ch in pages for txt, p in ch]
    tot = len(real)
    fmt = collections.Counter(re.sub(r"\d+(?:\.\d+)?", "9", t) for _, _, t, _, _ in real)
    print("   chuỗi:", tot, "| bộ số khác nhau:", len(set(tuple(x[3]) for x in real)),
          "| file:", len(set(x[0] for x in real)), "| dạng hay gặp:", fmt.most_common(8))

    # ---- 2. dương tính: chuỗi thật tự khớp với dòng số của trang
    sp = sum(set_ok(p, n) for _, _, _, p, n in real); op = sum(ord_ok(p, n) for _, _, _, p, n in real)
    print(f"2. dương tính: tập {sp}/{tot}  thứ tự {op}/{tot} ({100 * op / max(tot, 1):.1f}%)")
    for x in [x for x in real if not ord_ok(x[3], x[4])][:5]: print("   không tự khớp:", x[:4])

    # ---- 3. dạng viết lại kiểu LLM: phải tách ra đúng các số đó
    def viet_lai(p):
        a = [f"{int(v):,}" if v.isdigit() and int(v) >= 1000 else v for v in p]
        return [" x ".join(p) + " mm", "x".join(p), " × ".join(a), "mm x ".join(p) + "mm"]
    rw = ok = 0
    for _, _, t, p, _ in real:
        if "+" in t or "=" in t or "-" in t: continue
        for s in viet_lai(p):
            rw += 1; g = tach_chuoi(s); ok += bool(g) and g[0][1] == p
    print(f"3. viết lại kiểu LLM (x, ×, 1,200, mm): tách đúng {ok}/{rw}")

    # ---- 4. đột biến khó: thay 1 số bằng số KHÁC có trên trang; cửa sổ 1 (liền nhau) / 2 / 3 số
    SL = (1, 2, 3)
    c_set = n = 0; c_ord = dict.fromkeys(SL, 0); ex = []
    for fid, pg, t, parts, nums in real:
        pool = sorted(set(x for x in nums if x not in parts))
        if not pool: continue
        for _ in range(3):
            k = random.randrange(len(parts)); m = list(parts); m[k] = random.choice(pool); n += 1
            if not set_ok(m, nums): c_set += 1
            for s in SL:
                if not ord_ok(m, nums, s): c_ord[s] += 1
                elif s == 1 and len(ex) < 5: ex.append((fid, pg, t, m))
    print(f"4. đột biến khó (n={n}): TẬP bắt {100 * c_set / max(n, 1):.0f}% | THỨ TỰ bắt "
          + " | ".join(f"cửa sổ {s}: {c_ord[s]} ({100 * c_ord[s] / max(n, 1):.1f}%)" for s in SL))
    for e in ex: print("   lọt (cửa sổ 1):", e)

    # ---- 5. đổi chỗ 2 chiều
    sw_set = m2 = 0; sw = dict.fromkeys(SL, 0)
    for fid, pg, t, parts, nums in real:
        if len(parts) < 2 or parts[0] == parts[1]: continue
        m = list(parts); m[0], m[1] = m[1], m[0]; m2 += 1
        sw_set += not set_ok(m, nums)
        for s in SL: sw[s] += not ord_ok(m, nums, s)
    print(f"5. đổi chỗ 2 số đầu (n={m2}): tập bắt {sw_set} | thứ tự bắt " + " | ".join(f"cửa sổ {s}: {sw[s]}" for s in SL))

    # ---- 6. trùng ngẫu nhiên: chuỗi của trang A khớp trên trang B (cùng file / khác file)
    #      "B không ghi chuỗi đó" = trùng tình cờ thật (bảng đèn các file F064x lặp cùng mẫu W10 x H10)
    for ten, cung, loai in (("cùng file", True, False), ("khác file", False, False),
                            ("khác file, B không ghi chuỗi đó", False, True)):
        hit = dict.fromkeys(SL, 0); n6 = 0
        for fid, pg, t, parts, nums in real:
            oth = [x for x in pages if (x[0] == fid) == cung and (x[0], x[1]) != (fid, pg)
                   and not (loai and any(p == parts for _, p in x[4]))]
            if not oth: continue
            b = random.choice(oth); n6 += 1
            for s in SL: hit[s] += ord_ok(parts, b[3], s)
        print(f"6. chuỗi trang A khớp trên trang ngẫu nhiên B {ten} (n={n6}): "
              + " | ".join(f"cửa sổ {s}: {hit[s]} ({100 * hit[s] / max(n6, 1):.0f}%)" for s in SL))

    # ---- 7. mẫu để xem tay rác
    print("7. mẫu 30 chuỗi để xem tay:")
    for x in random.sample(real, min(30, tot)): print("  ", x[0], x[1], repr(x[2]))
