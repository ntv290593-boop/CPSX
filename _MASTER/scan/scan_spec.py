"""Tach spec dang MasterFormat (PART 1/2/3, 1.01.., A., 1.) theo trang tu text layout.

Dau vao: text pdftotext -layout (trang tach bang \\f) hoac danh sach chuoi theo trang.
Tra ve danh sach muc: {ref: "3.02.A.3", title, text, page}.
"""
import re, sys, json

ART = re.compile(r"^\s{0,10}(\d\.\d\d)\s{2,}(.+)$")       # 3.02    TIMBER DOORS
PAR = re.compile(r"^\s{4,20}([A-Z])\.\s{2,}(.*)$")           # A.      General:
SUB = re.compile(r"^\s{8,30}(\d{1,2})\.\s{2,}(.*)$")          # 3.      Timber doors shall ...
KEYS = {
    "submittal": r"submit|shop drawing|sample|mock-?up",
    "qa": r"single-source|qualification|certif|test report",
    "material": r"\b\d+\s?mm\b|hardwood|plywood|veneer|lipping|MDF|laminate",
    "warranty": r"warrant|guarantee",
}

def parse(pages):
    out, art, par, toc = [], None, None, True
    for pn, txt in enumerate(pages, 1):
        for line in txt.splitlines():
            m = ART.match(line)
            if m and m.group(2).strip().isupper():
                toc = False
                art, par = m.group(1), None
                out.append({"ref": art, "title": m.group(2).strip(), "text": "", "page": pn})
                continue
            if toc or not art:
                continue
            m = PAR.match(line)
            if m:
                par = m.group(1)
                out.append({"ref": f"{art}.{par}", "title": m.group(2).strip(), "text": "", "page": pn})
                continue
            m = SUB.match(line)
            if m and par:
                out.append({"ref": f"{art}.{par}.{m.group(1)}", "title": "", "text": m.group(2).strip(), "page": pn})
                continue
            if out and line.strip() and not re.search(r"Page \d+ of \d+|DEVELOPMENT$", line):
                out[-1]["text"] = (out[-1]["text"] + " " + line.strip()).strip()
    for o in out:
        body = o["title"] + " " + o["text"]
        o["tags"] = [k for k, r in KEYS.items() if re.search(r, body, re.I)]
    return out

if __name__ == "__main__":
    pages = open(sys.argv[1], encoding="utf-8").read().split("\f")
    print(json.dumps(parse(pages), ensure_ascii=False, indent=1))
