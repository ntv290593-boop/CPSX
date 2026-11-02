# -*- coding: utf-8 -*-
r"""gemini_doc.py — Gemini (qua agy) đọc độc lập ảnh crop bản vẽ: người đọc thứ 2.
Chép từ 01-scan pd.py (đóng băng nên giữ bản riêng). Không biết đường dẫn dự án, form, bước.

  doc_anh(pngs, hoi, vat, par)  -> [(png, chữ trả lời)]; chế độ số lưu cache <ảnh>.gemini.txt
  gem_doc(chữ)                  -> danh sách số Gemini đọc, None nếu lỗi / không có JSON
Env TD_AGY = đường dẫn agy thay thế (chỉ để kiểm thử).
"""
import glob
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


def doc_anh(pngs, hoi=None, vat=None, par=4):
    """Đọc song song. vat = tên món (đủ bộ trường quan sát), hoi = câu hỏi riêng, không có cả hai = liệt kê số kích thước."""
    if vat:
        hoi = GEM_QS.format(vat=vat, fields=json.dumps(GEM_TRUONG, ensure_ascii=False))
    exe = agy_exe()

    def run(png):
        png = os.path.abspath(png)
        out = png + ".gemini.txt"
        if not hoi and os.path.exists(out) and os.path.getsize(out) > 2 and os.path.getmtime(out) >= os.path.getmtime(png):
            with open(out, encoding="utf-8") as fh:
                return png, fh.read()
        prompt = (f"Do NOT run shell commands; use your file-viewing tool directly on this absolute path: {png} . "
                  + ((hoi if vat else hoi + " Answer only from what is visible in the image; if not visible say NOT VISIBLE. Never guess.") if hoi else GEM_SO))
        try:
            r = subprocess.run([exe, "--model", GEM_MODEL, "--print", prompt, "--print-timeout", "280s"], cwd=os.path.dirname(png),
                               capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=320)
            txt = r.stdout.strip() or f"LỖI agy: {r.stderr.strip()[:200]}"
        except subprocess.TimeoutExpired:
            txt = "LỖI agy: quá 320 s"
        if not hoi and gem_doc(txt) is not None:
            with open(out, "w", encoding="utf-8") as fh:
                fh.write(txt)
        return png, txt

    with cf.ThreadPoolExecutor(par) as ex:
        return list(ex.map(run, pngs))
