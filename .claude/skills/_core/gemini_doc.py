# -*- coding: utf-8 -*-
r"""gemini_doc.py — Gemini (qua agy) đọc độc lập ảnh crop bản vẽ: người đọc thứ 2.
Chép từ 01-scan pd.py (đóng băng nên giữ bản riêng). Không biết đường dẫn dự án, form, bước.

  doc_anh(pngs, hoi, vat, par, cache_dir) -> [(png, chữ trả lời)]
      Cache theo NỘI DUNG: <cache_dir>\<sha1 byte PNG>_<sha1 model + câu hỏi>.txt, áp mọi chế độ (số, --quansat, --hoi).
      Cùng vùng cắt cùng dpi ra file khác (agent khác) vẫn trúng cache; khác dpi/vùng = ảnh khác = gọi lại.
      cache_dir mặc định: env TD_GEM_CACHE, không có thì thư mục _gemini_cache cạnh ảnh. Trả lời lỗi / không có JSON không lưu.
  gem_doc(chữ)                  -> danh sách số Gemini đọc, None nếu lỗi / không có JSON
Env TD_AGY = đường dẫn agy thay thế (chỉ để kiểm thử).
"""
import glob
import hashlib
import json
import os
import re
import shutil
import subprocess
import concurrent.futures as cf

GEM_MODEL = "gemini-3.8-flash-low"  # benchmark 24/10: low đủ chính xác (~13 s/crop); medium/high quá chậm
GEM_SO = ("It is a crop of an interior/joinery drawing (dimensions in millimetres). List EVERY dimension number printed along "
          "dimension lines in the crop, including vertical dimension lines, one entry per printed number, in reading order. "
          "Do NOT include detail numbers inside circles, drawing numbers, scale text, level markers (FFL/FCL/AFFL) or note text. "
          "If a number is unreadable write \"?\". Never guess. Output ONLY a JSON array of strings, nothing else.")
GEM_TRUONG = ["Lắp đặt (gắn tường/đứng rời/âm tường)", "Kích thước BV (R×S×C mm, chỉ số in trên bản vẽ)", "Có cánh (số cánh, mở/lùa)",
              "Ngăn kéo (số)", "Thanh treo", "Kệ/đợt (số)", "Hốc/ô tivi", "Mặt đá", "Gương", "Kính", "Kim loại",
              "Giấy/vải", "Mã hoàn thiện trên BV (nhãn WD/PT/LAM/SS… chỉ vào món này)", "Khung xương", "Cong/bo",
              "Phần ngoài phạm vi (chậu, vòi, thiết bị… thấy trên hình)", "Số chi tiết/hình đang xem"]
GEM_QS = ("It is a page or crop of an interior/joinery drawing. Inspect ONLY this item: {vat}. "
          "For each field below, report what is VISIBLE in the image. Output ONLY a JSON object: "
          "{{\"<field>\": {{\"value\": \"…\", \"status\": \"YES|NO|NOT VISIBLE\", \"where\": \"view/detail number + position in the image\"}}}}. "
          "YES = present, NO = clearly absent, NOT VISIBLE = cannot tell from this image. Count only what you can see; never guess, "
          "never infer from typical designs. Fields (keep these exact Vietnamese keys): {fields}")


def agy_exe():
    p = os.environ.get("TD_AGY") or shutil.which("agy")
    if p:
        return p
    g = glob.glob(os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "WinGet", "Packages", "Google.AntigravityCLI*", "**", "agy.exe"), recursive=True)
    if not g:
        raise SystemExit("Không thấy agy.exe (Antigravity CLI). Bỏ qua kiểm chéo Gemini, ghi vào kết quả trả về.")
    return g[0]


def gem_doc(txt):
    """Trả lời của Gemini ở chế độ số -> danh sách chuỗi, None nếu lỗi / không có JSON."""
    m = re.search(r"\[.*\]", txt, re.S)
    if txt.startswith("LỖI") or not m:
        return None
    try:
        return [str(x).strip() for x in json.loads(m.group(0))]
    except ValueError:
        return None


def _hash(b):
    return hashlib.sha1(b).hexdigest()[:20]


def _dang_luu(txt, che_do):
    """Trả lời đáng lưu cache: không lỗi; chế độ số phải có mảng JSON, quan sát phải có object JSON."""
    if not txt or txt.startswith("LỖI"):
        return False
    if che_do == "so":
        return gem_doc(txt) is not None
    if che_do == "qs":
        return "{" in txt and "}" in txt
    return True


def cau_hoi(hoi=None, vat=None):
    """-> (câu gửi Gemini, chế độ so|qs|hoi)."""
    if vat:
        return GEM_QS.format(vat=vat, fields=json.dumps(GEM_TRUONG, ensure_ascii=False)), "qs"
    if hoi:
        return hoi + " Answer only from what is visible in the image; if not visible say NOT VISIBLE. Never guess.", "hoi"
    return GEM_SO, "so"


def file_cache(png, cau, cache_dir=None):
    """Đường dẫn cache của 1 ảnh + 1 câu hỏi (theo nội dung, không theo tên file)."""
    png = os.path.abspath(png)
    with open(png, "rb") as fh:
        k_png = _hash(fh.read())
    d = cache_dir or os.environ.get("TD_GEM_CACHE") or os.path.join(os.path.dirname(png), "_gemini_cache")
    return os.path.join(d, f"{k_png}_{_hash((GEM_MODEL + chr(10) + cau).encode('utf-8'))}.txt")


def doc_anh(pngs, hoi=None, vat=None, par=4, cache_dir=None):
    """Đọc song song. vat = tên món (đủ bộ trường quan sát), hoi = câu hỏi riêng, không có cả hai = liệt kê số kích thước."""
    cau, che_do = cau_hoi(hoi, vat)
    exe = agy_exe()

    def run(png):
        png = os.path.abspath(png)
        out = file_cache(png, cau, cache_dir)
        if os.path.exists(out) and os.path.getsize(out) > 2:
            with open(out, encoding="utf-8") as fh:
                return png, fh.read()
        prompt = f"Do NOT run shell commands; use your file-viewing tool directly on this absolute path: {png} . " + cau
        try:
            r = subprocess.run([exe, "--model", GEM_MODEL, "--print", prompt, "--print-timeout", "280s"], cwd=os.path.dirname(png),
                               capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=320)
            txt = r.stdout.strip() or f"LỖI agy: {r.stderr.strip()[:200]}"
        except subprocess.TimeoutExpired:
            txt = "LỖI agy: quá 320 s"
        if _dang_luu(txt, che_do):
            os.makedirs(os.path.dirname(out), exist_ok=True)
            tmp = f"{out}.{os.getpid()}.{abs(hash(png))}.tmp"  # agent song song: ghi tạm rồi đổi tên, không ai đọc nửa file
            with open(tmp, "w", encoding="utf-8") as fh:
                fh.write(txt)
            os.replace(tmp, out)
        return png, txt

    with cf.ThreadPoolExecutor(par) as ex:
        return list(ex.map(run, pngs))
