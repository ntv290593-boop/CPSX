// CPSX_LIB v1 — thư viện recipe cho Claude trong Excel. Thân hàm: cần `context`, `Excel`; trả về L.
// Nạp: const L=new Function("context","Excel",libText)(context,Excel);  Gọi: return await L.map({});
const L = { VER: "2.1" };
const CT = "Chiet tinh", LOG = "_EDITLOG", YELLOW = "#FFFF00";
const wb = context.workbook;
const sync = () => context.sync();
const sh = (n) => wb.worksheets.getItem(n);
const colL = (n) => { let s = ""; n++; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }; // 0-based
const colN = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const norm = (v) => String(v == null ? "" : v).replace(/\s+/g, " ").trim().toLowerCase();
const q = (n) => /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'";
const rnd = (v) => typeof v === "number" ? Math.round(v * 100) / 100 : v;
const isErr = (v) => typeof v === "string" && /^#(N\/A|NAME\?|REF!|VALUE!|DIV\/0!|NUM!|NULL!|SPILL!|CALC!)/.test(v);
async function calc(mode) { wb.application.calculationMode = mode === "auto" ? Excel.CalculationMode.automatic : Excel.CalculationMode.manual; await sync(); }

// ---------- dịch công thức khi chép block ----------
const REF = /(?<![A-Za-z0-9_.$!'])((?:'(?:[^']|'')+'|[^\s'!(),;=+\-*/&<>^"{}:]+)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![A-Za-z0-9_(])/g;
// c = {srcSheet, s, e, off, boqSheet, srcBoqRow, boqRow}
function trans(f, c) {
  if (typeof f !== "string" || f[0] !== "=") return f;
  const keepAbs = /SUMIFS\(A[NU]\$\d+:A[NU]\$\d+/.test(f);
  return f.split(/("(?:[^"]|"")*")/).map((part, i) => i % 2 ? part : part.replace(REF, (m, pre, a1, c1, a2, r1, b1, c2, b2, r2) => {
    const sheet = pre ? pre.slice(0, -1).replace(/^'|'$/g, "").replace(/''/g, "'") : null;
    const rows = [[a2, +r1]]; if (r2) rows.push([b2, +r2]);
    let nr;
    if (!sheet || sheet === c.srcSheet) {
      const absIn = rows.filter(x => x[0]).every(x => x[1] >= c.s && x[1] <= c.e);
      nr = rows.map(([abs, r]) => !abs ? r + c.off : (!keepAbs && absIn ? r + c.off : r));
    } else if (sheet === c.boqSheet && c.srcBoqRow) {
      nr = rows.map(([, r]) => r === c.srcBoqRow ? c.boqRow : r);
    } else return m;
    return (pre || "") + a1 + c1 + a2 + nr[0] + (r2 ? ":" + b1 + c2 + b2 + nr[1] : "");
  })).join("");
}
L._trans = trans;

// ---------- bảng gộp tấm (SUMIFS nhiều dải, loại dòng tiêu đề block) ----------
const SEG = /SUMIFS\(AN\$(\d+):AN\$(\d+),AT\$\d+:AT\$\d+,([^,()]+),AL\$\d+:AL\$\d+,(\{[^}]*\})\)/g;
function segs(f) { const out = []; let m; SEG.lastIndex = 0; while ((m = SEG.exec(f))) out.push({ a: +m[1], b: +m[2], key: m[3], crit: m[4], i: m.index, len: m[0].length }); return out; }
function buildSegs(list, key, crit) { return list.map(([a, b]) => `SUMIFS(AN$${a}:AN$${b},AT$${a}:AT$${b},${key},AL$${a}:AL$${b},${crit})`).join(","); }
function replaceSegs(f, list) {
  const s = segs(f); if (!s.length) return f;
  const start = s[0].i, end = s[s.length - 1].i + s[s.length - 1].len;
  return f.slice(0, start) + buildSegs(list, s[0].key, s[0].crit) + f.slice(end);
}
function splitAt(list, h) { const o = []; for (const [a, b] of list) { if (h >= a && h <= b) { if (h - 1 >= a) o.push([a, h - 1]); if (b >= h + 1) o.push([h + 1, b]); } else o.push([a, b]); } return o; }
L._segs = segs; L._replaceSegs = replaceSegs; L._splitAt = splitAt;

// ---------- đọc cấu trúc ----------
// Cấu hình BoQ: tự nhận từ hàng tiêu đề; cấu hình đã lưu bằng L.setBoq (_LIB!B3) được ưu tiên.
// Khoá: sheet, hdrRow, code, label (tên hạng mục), qty, price, amt, desc, w, d, h (chữ cột)
async function savedBoq() {
  const s = wb.worksheets.getItemOrNullObject("_LIB"); await sync(); if (s.isNullObject) return null;
  const c = s.getRange("B3"); c.load("values"); await sync();
  try { return c.values[0][0] ? JSON.parse(c.values[0][0]) : null; } catch (e) { return null; }
}
function detectCols(values) {
  const o = {}; let hdr = 0;
  const put = (k, j, i) => { if (o[k] === undefined) { o[k] = colL(j); hdr = Math.max(hdr, i + 1); } };
  values.forEach((row, i) => row.forEach((v, j) => {
    const t = norm(v);
    if (/^(đơn giá|unit price)/.test(t)) put("price", j, i);
    else if (/^(thành tiền|amount)/.test(t)) put("amt", j, i);
    else if (/^(tổng kl|tổng khối lượng|quantity|q'?ty|số lượng)/.test(t)) put("qty", j, i);
    else if (/(ghi chú costing|mô tả costing|^remark)/.test(t)) put("desc", j, i);
    else if (/^(item|hạng mục|tên hạng mục)/.test(t)) put("name", j, i);
    else if (t === "w" || t === "dài") put("w", j, i);
    else if (t === "d" || t === "sâu" || t === "rộng") put("d", j, i);
    else if (t === "h" || t === "cao") put("h", j, i);
  }));
  return { o, hdr };
}
async function findBoq(name) {
  const saved = await savedBoq(); if (saved && (!name || saved.sheet === name)) return { name: saved.sheet, ...saved };
  const ws = wb.worksheets; ws.load("items/name"); await sync();
  const all = ws.items.map(w => w.name);
  const NOT_BOQ = /^(Matrix|Spec)$/; // sheet nội bộ, không phải BoQ của CĐT (v2.1)
  let cands = name ? [name] : all.filter(n => /^(matrix|boq)/i.test(n) && !NOT_BOQ.test(n));
  if (!name && all.includes(CT)) { // ưu tiên sheet BoQ mà Chiet tinh đang trỏ tới
    const a = sh(CT).getRange("A1:A400"); a.load("formulas"); await sync();
    const hit = a.formulas.map(r => String(r[0]).match(/^=\+?'?([^'!]+)'?!/)).find(m => m && all.includes(m[1]) && m[1] !== CT && !NOT_BOQ.test(m[1]));
    if (hit) cands = [hit[1], ...cands.filter(n => n !== hit[1])];
  }
  for (const n of cands) {
    const r = sh(n).getRange("A1:AZ15"); r.load("values"); await sync();
    const { o, hdr } = detectCols(r.values);
    if (!o.price) continue;
    const pj = colN(o.price), aj = o.amt ? colN(o.amt) : pj + 1;
    return { name: n, sheet: n, hdrRow: hdr, code: "A", label: o.name, qty: o.qty || colL(pj - 1), price: o.price, amt: colL(aj), desc: o.desc || colL(aj + 1), w: o.w, d: o.d, h: o.h };
  }
  throw new Error("Không tìm thấy sheet BoQ (cột 'Đơn giá'/'Unit Price'). Gọi L.setBoq({sheet,price,qty,amt,desc,...}).");
}
L.setBoq = async (cfg) => {
  const s = wb.worksheets.getItem("_LIB"); s.getRange("B3").numberFormat = [["@"]]; s.getRange("B3").values = [[JSON.stringify(cfg)]]; await sync();
  return { saved: cfg };
};
async function ctBlocks() {
  const s = sh(CT), u = s.getUsedRange(true); u.load("rowCount,rowIndex"); await sync();
  const n = u.rowIndex + u.rowCount + 2;
  const a = s.getRange(`A1:A${n}`); a.load("formulas,values");
  const h = s.getRange(`H1:H${n}`); h.load("values"); await sync();
  const blocks = []; let lastEnd = 0, open = null;
  for (let i = 0; i < n; i++) {
    const v = norm(a.values[i][0]), f = String(a.formulas[i][0]);
    if (v === "end") { if (open) { open.end = i + 1; blocks.push(open); open = null; } lastEnd = i + 1; continue; }
    if (!open && lastEnd && v !== "") {
      const m = f.match(/!\$?A\$?(\d+)/);
      open = { hdr: i + 1, code: a.values[i][0], boqRow: m ? +m[1] : null, price: rnd(h.values[i][0]) };
    }
  }
  return { blocks, lastEnd, n };
}
async function libBlocks(name) {
  const s = sh(name), u = s.getUsedRange(true); u.load("rowCount,rowIndex"); await sync();
  const n = u.rowIndex + u.rowCount; const a = s.getRange(`A1:E${n}`); a.load("values"); await sync();
  const out = []; let start = null; // dòng đầu block = A có chữ và B hoặc E là số; kết thúc ở END
  for (let i = 0; i < n; i++) {
    const r = a.values[i], v = String(r[0]).trim();
    if (norm(v) === "end") { if (start) { out.push([start.name, start.row, i + 1]); start = null; } continue; }
    if (!start && v && (typeof r[1] === "number" || typeof r[4] === "number")) start = { name: v.slice(0, 60), row: i + 1 };
  }
  return out;
}
async function blockEnd(sheetName, hdr) {
  const r = sh(sheetName).getRange(`A${hdr}:A${hdr + 600}`); r.load("values"); await sync();
  for (let i = 1; i < r.values.length; i++) if (norm(r.values[i][0]) === "end") return hdr + i;
  throw new Error(`Không thấy END dưới ${sheetName}!A${hdr}`);
}
async function resolveHdr(p) {
  if (p.hdr) return p.hdr;
  const { blocks } = await ctBlocks();
  const b = blocks.find(x => norm(x.code) === norm(p.code)); if (!b) throw new Error("Không thấy block Chiet tinh có mã " + p.code);
  return b.hdr;
}

// ---------- tô vàng có nhật ký (để bỏ vàng đúng màu cũ) ----------
async function logSheet() {
  let s = wb.worksheets.getItemOrNullObject(LOG); await sync();
  if (s.isNullObject) { s = wb.worksheets.add(LOG); s.getRange("A1:E1").values = [["Sheet", "Ô", "Màu cũ", "Thời gian", "Ghi chú"]]; s.visibility = Excel.SheetVisibility.veryHidden; await sync(); }
  return s;
}
async function mark(cells, tag) { // cells: [[sheet, "A1"], ...] (ô đơn hoặc vùng)
  if (!cells.length) return 0;
  const list = [];
  for (const [sn, addr] of cells) { const r = sh(sn).getRange(addr); const p = r.getCellProperties({ address: true, format: { fill: { color: true } } }); list.push([sn, p]); }
  await sync();
  const rows = [], now = new Date().toISOString().slice(0, 16);
  for (const [sn, p] of list) p.value.forEach(row => row.forEach(c => { const col = (c.format.fill.color || "").toUpperCase(); if (col !== YELLOW) rows.push([sn, c.address.split("!").pop().replace(/\$/g, ""), col, now, tag || ""]); }));
  for (const [sn, addr] of cells) sh(sn).getRange(addr).format.fill.color = YELLOW;
  if (rows.length) {
    const s = await logSheet(); const u = s.getUsedRange(); u.load("rowCount"); await sync();
    s.getRange(`A${u.rowCount + 1}:E${u.rowCount + rows.length}`).values = rows;
  }
  await sync(); return rows.length;
}
L.mark = async (p) => ({ marked: await mark(p.cells, p.tag) });

L.unmark = async (p = {}) => {
  if (p.legacy) { // ô vàng có trước khi có nhật ký: chép màu từ ô tương ứng của block mẫu
    const g = p.legacy, s = sh(g.sheet), src = sh(g.srcSheet || g.sheet);
    const r = s.getRange(g.range); r.load("rowIndex,columnIndex");
    const pr = r.getCellProperties({ address: true, format: { fill: { color: true } } }); await sync();
    const off = g.srcStartRow - (r.rowIndex + 1), hits = [];
    pr.value.forEach(row => row.forEach(c => { if ((c.format.fill.color || "").toUpperCase() === YELLOW) hits.push(c.address.split("!").pop().replace(/\$/g, "")); }));
    const srcP = hits.map(a => { const m = a.match(/([A-Z]+)(\d+)/); const x = src.getRange(m[1] + (+m[2] + off)); x.load("format/fill/color"); return x; });
    await sync(); let kept = 0;
    hits.forEach((a, i) => { const col = (srcP[i].format.fill.color || "").toUpperCase(); if (col === YELLOW) { kept++; return; } const t = s.getRange(a); if (!col || col === "#FFFFFF") t.format.fill.clear(); else t.format.fill.color = col; });
    await sync(); return { cleared: hits.length - kept, keptTemplateYellow: kept };
  }
  const s = wb.worksheets.getItemOrNullObject(LOG); await sync(); if (s.isNullObject) return { cleared: 0, note: "Chưa có _EDITLOG" };
  const u = s.getUsedRange(); u.load("values"); await sync();
  const keep = [u.values[0]]; let n = 0;
  for (const row of u.values.slice(1)) {
    const [sn, addr, col] = row; const rn = +String(addr).replace(/^[A-Z]+/, "");
    const hit = (!p.sheet || p.sheet === sn) && (!p.from || rn >= p.from) && (!p.to || rn <= p.to) && (!p.tag || p.tag === row[4]);
    if (!hit) { keep.push(row); continue; }
    const t = sh(sn).getRange(addr); if (!col || col === "#FFFFFF") t.format.fill.clear(); else t.format.fill.color = col; n++;
  }
  u.clear(); s.getRange(`A1:E${keep.length}`).values = keep; await sync();
  return { cleared: n, remainingInLog: keep.length - 1 };
};

// ---------- MAP: 1 lần gọi đầu phiên ----------
L.map = async (p = {}) => {
  const out = { lib: L.VER };
  const ws = wb.worksheets; ws.load("items/name,items/visibility"); await sync();
  const names = ws.items.map(w => w.name); out.sheets = names.filter(n => !/snapshot|_BK_/i.test(n));
  // quy tắc
  const t = wb.tables.getItemOrNullObject("tblRules"); await sync();
  if (!t.isNullObject) {
    const hd = t.getHeaderRowRange(), b = t.getDataBodyRange(); hd.load("values"); b.load("values"); await sync();
    const H = hd.values[0].map(norm), ci = (k) => H.indexOf(k);
    out.rules = b.values.filter(r => /active|pending/i.test(r[ci("trạng thái")])).map(r => `${r[ci("chủ đề")]}: ${r[ci("quy tắc")]}`);
  }
  const tm = wb.tables.getItemOrNullObject("tblRulesMasterView"); await sync();
  if (!tm.isNullObject) { const b = tm.getDataBodyRange(); b.load("values"); const hd = tm.getHeaderRowRange(); hd.load("values"); await sync();
    const H = hd.values[0].map(norm), qi = H.indexOf("quy tắc"), ti = H.indexOf("chủ đề"), si = H.indexOf("trạng thái");
    out.masterRules = b.values.filter(r => si < 0 || /active/i.test(r[si])).map(r => `${ti >= 0 ? r[ti] + ": " : ""}${r[qi]}`); }
  // BoQ
  try {
    const B = await findBoq(p.boqSheet); out.boq = B;
    const s = sh(B.name), u = s.getUsedRange(true); u.load("rowCount,rowIndex"); await sync();
    const last = u.rowIndex + u.rowCount, c0 = colN(B.qty), cl = colN(B.label || "C");
    const maxC = colL(Math.max(...[B.desc, B.qty, B.amt, B.label || "C"].map(colN)));
    const r = s.getRange(`A1:${maxC}${last}`); r.load("values"); const f = s.getRange(`${B.price}1:${B.price}${last}`); f.load("formulas"); await sync();
    const re = p.codes ? new RegExp(p.codes, "i") : null;
    out.items = r.values.map((row, i) => [i + 1, row[0], row]).filter(([i, code]) => code !== "" && i > (B.hdrRow || 0) && (!re || re.test(code)))
      .map(([i, code, row]) => [i, code, String(row[cl]).slice(0, p.descLen || 90), rnd(row[c0]), String(f.formulas[i - 1][0]).slice(0, 40), row[colN(B.desc)] ? "có mô tả" : ""]);
    if (!re && out.items.length > 300) { out.itemsNote = `${out.items.length} dòng, chỉ trả 300 dòng đầu — truyền codes:"^D0" để lọc`; out.items = out.items.slice(0, 300); }
  } catch (e) { out.boqError = String(e.message || e); }
  // Chiet tinh + thư viện mẫu
  if (names.includes(CT)) { const c = await ctBlocks(); out.chietTinh = { lastEnd: c.lastEnd, blocks: c.blocks.map(b => [b.hdr, b.end, b.code, b.boqRow, b.price]) }; }
  out.library = {};
  for (const n of ["Door", "Template", "Kitchen"]) if (names.includes(n)) out.library[n] = await libBlocks(n);
  const app = wb.application; app.load("calculationMode"); await sync(); out.calc = app.calculationMode;
  return out;
};

// ---------- áp sửa trong block theo nhãn cột A ----------
// edit: [nhãn | "hdr" | "+k", cột, giá trị]; nhãn "Đố đứng#2" = lần xuất hiện thứ 2; giá trị chuỗi "=..." là công thức, có {r} {h} {boq} {B}
async function applyEdits(sheetName, hdr, end, edits, ctx) {
  if (!edits || !edits.length) return [];
  const s = sh(sheetName), a = s.getRange(`A${hdr}:A${end}`); a.load("values"); await sync();
  const done = [];
  for (const [lab, col, val] of edits) {
    let row;
    if (lab === "hdr") row = hdr;
    else if (/^\+\d+$/.test(lab)) row = hdr + +lab.slice(1);
    else { const [t, k] = lab.split("#"); let cnt = 0; const want = +(k || 1);
      for (let i = 0; i < a.values.length; i++) if (norm(a.values[i][0]).startsWith(norm(t)) && ++cnt === want) { row = hdr + i; break; } }
    if (!row) throw new Error(`Không thấy hàng "${lab}" trong block ${hdr}-${end}`);
    const addr = col + row, cell = s.getRange(addr);
    if (typeof val === "string" && val[0] === "=") cell.formulas = [[val.replace(/\{r\}/g, row).replace(/\{h\}/g, hdr).replace(/\{boq\}/g, ctx.boqRow).replace(/\{B\}/g, q(ctx.boqSheet || ""))]];
    else cell.values = [[val]];
    done.push(addr);
  }
  await sync(); return done;
}
async function report(sheetName, hdr, end) {
  const s = sh(sheetName), r = s.getRange(`A${hdr}:BL${end}`); r.load("values"); await sync();
  const v = r.values, errors = [], rows = [];
  v.forEach((row, i) => {
    row.forEach((x, j) => { if (isErr(x)) { const known = (j === 1 && x === "#NAME?") || (j === 27 && norm(row[0]) === "cánh cửa"); errors.push(colL(j) + (hdr + i) + ":" + x + (known ? "(có sẵn)" : "")); } });
    if (i === 0 || (typeof row[4] === "number" && row[4] !== 0 && row[8] !== "" && row[8] !== 0))
      rows.push([hdr + i, String(row[0]).trim().slice(0, 32), ...[1, 2, 3, 4, 5, 6, 8].map(j => rnd(row[j]))]);
  });
  return { price: rnd(v[0][7]), errors: errors.slice(0, 40), nErrors: errors.length, rows: rows.slice(0, 70), cols: "row,A,B,C,D,E,F,G,I" };
}

// ---------- NEW ITEM: chép block → cuối Chiet tinh → link BoQ ----------
// p = {src:{sheet,row}|{code}, boqRow, boqSheet?, edits?, mark?=true, collapse?=true}
L.newItem = async (p) => {
  const notes = [];
  await calc("manual");
  try {
    const B = await findBoq(p.boqSheet);
    const ct = await ctBlocks();
    let src = p.src || {};
    if (src.code) { const b = ct.blocks.find(x => norm(x.code) === norm(src.code)); if (!b) throw new Error("Không thấy block mã " + src.code); src = { sheet: CT, row: b.hdr }; }
    const sEnd = await blockEnd(src.sheet, src.row), n = sEnd - src.row + 1;
    const d = ct.lastEnd + 1, dEnd = d + n - 1, off = d - src.row;
    const S = sh(src.sheet), T = sh(CT);
    const chk = T.getRange(`A${d}:BL${dEnd}`); chk.load("values");
    const sr = S.getRange(`A${src.row}:BL${sEnd}`); sr.load("formulas");
    const sa = S.getRange(`A${src.row}`); sa.load("formulas"); await sync();
    if (chk.values.flat().some(x => x !== "")) throw new Error(`Vùng đích Chiet tinh!A${d}:BL${dEnd} không trống`);
    const mb = String(sa.formulas[0][0]).match(/!\$?A\$?(\d+)/), srcBoqRow = mb ? +mb[1] : null;
    // 1) chép định dạng + giá trị
    T.getRange(`A${d}:BL${dEnd}`).copyFrom(sr, Excel.RangeCopyType.all); await sync();
    // 2) ghi công thức đã dịch theo từng cột
    const c = { srcSheet: src.sheet, s: src.row, e: sEnd, off, boqSheet: B.name, srcBoqRow, boqRow: p.boqRow };
    const F = sr.formulas.map(row => row.map(x => trans(x, c))), failCols = [];
    for (let j = 0; j < 64; j++) {
      const colF = F.map(r => [r[j]]); if (!colF.some(x => typeof x[0] === "string" && x[0][0] === "=")) continue;
      try { T.getRange(`${colL(j)}${d}:${colL(j)}${dEnd}`).formulas = colF; await sync(); } catch (e) { failCols.push(colL(j)); }
    }
    if (failCols.length) notes.push("Cột không ghi lại được (giữ bản copyFrom, kiểm nếu có tham chiếu tuyệt đối): " + failCols.join(","));
    // 3) dòng tiêu đề: nếu mẫu chưa link BoQ thì link A, E
    const set = [];
    if (!srcBoqRow) { // mẫu thư viện: link tên/mã, kích thước, SL từ dòng BoQ
      const Q = q(B.name), cv = sh(B.name).getRange(`A${p.boqRow}`); cv.load("values"); await sync();
      const useLabel = B.label && typeof cv.values[0][0] === "number";
      const hl = { A: `=${Q}!${useLabel ? B.label : "A"}${p.boqRow}`, E: `=${Q}!${B.qty}${p.boqRow}` };
      if (B.w) { hl.B = `=${Q}!${B.w}${p.boqRow}`; if (B.d) hl.C = `=${Q}!${B.d}${p.boqRow}`; if (B.h) hl.D = `=${Q}!${B.h}${p.boqRow}`; }
      else notes.push("BoQ không có cột W/D/H: đặt B (rộng), D (cao) qua edits");
      for (const k in hl) { T.getRange(k + d).formulas = [[hl[k]]]; set.push(k + d); }
    }
    else set.push(`A${d}`, `B${d}`, `D${d}`, `E${d}`);
    // 4) bảng gộp tấm
    const g = T.getRange(`G1:G${dEnd}`); g.load("formulas"); await sync();
    let canon = null;
    g.formulas.forEach((r, i) => { const row = i + 1; if (!canon && (row < d || row > dEnd)) { const sg = segs(String(r[0])); if (sg.length) canon = sg.map(x => [x.a, x.b]); } });
    let segFixed = 0, tooLong = [];
    g.formulas.forEach((r, i) => {
      const row = i + 1, f = String(r[0]); if (!segs(f).length) return;
      let list = (row >= d && row <= dEnd && canon) ? canon : segs(f).map(x => [x.a, x.b]);
      if (!canon) list = [[d + 1, dEnd]];
      list = splitAt(list, d);
      if (!list.some(([a, b]) => d + 1 >= a && d + 1 <= b)) list.push([d + 1, dEnd]);
      const nf = replaceSegs(f, list);
      if (nf !== f) { if (nf.length > 8000) { tooLong.push("G" + row); return; } T.getRange("G" + row).formulas = [[nf]]; segFixed++; }
    });
    await sync();
    if (tooLong.length) notes.push("Công thức gộp tấm quá dài, chưa sửa: " + tooLong.join(","));
    // 5) link đơn giá BoQ
    sh(B.name).getRange(B.price + p.boqRow).formulas = [[`='${CT}'!H${d}`]];
    // 6) sửa theo yêu cầu
    const edited = await applyEdits(CT, d, dEnd, p.edits, { boqRow: p.boqRow, boqSheet: B.name });
    // 7) chiều cao hàng + group
    const hs = []; for (let r = src.row; r <= sEnd; r++) { const x = S.getRange(`${r}:${r}`); x.load("format/rowHeight"); hs.push(x); }
    await sync();
    hs.forEach((x, i) => { if (x.format.rowHeight > 0) T.getRange(`${d + i}:${d + i}`).format.rowHeight = x.format.rowHeight; });
    try { T.getRange(`${d + 1}:${dEnd - 1}`).group(Excel.GroupOption.byRows); if (p.collapse !== false) T.getRange(`${d + 1}:${dEnd - 1}`).hideGroupDetails(Excel.GroupOption.byRows); await sync(); }
    catch (e) { notes.push("Chưa group được hàng: " + (e.message || e)); }
    // 8) tô vàng có nhật ký
    let marked = 0;
    if (p.mark !== false) marked = await mark([...new Set([...set, ...edited])].map(a => [CT, a]).concat([[B.name, B.price + p.boqRow]]), "newItem " + p.boqRow);
    if (p._batch) return { block: [d, dEnd], boqRow: p.boqRow, segFixed, edited, marked, notes };
    // 9) tính lại + báo cáo
    await calc("auto"); wb.application.calculate(Excel.CalculationType.full); await sync();
    const rep = await report(CT, d, dEnd);
    const bq = sh(B.name).getRange(`A${p.boqRow}:${B.desc}${p.boqRow}`); bq.load("values"); await sync();
    const bv = bq.values[0];
    return { block: [d, dEnd], from: `${src.sheet}!${src.row}-${sEnd}`, boq: { row: p.boqRow, code: bv[0], qty: rnd(bv[colN(B.qty)]), price: rnd(bv[colN(B.price)]), amount: rnd(bv[colN(B.amt)]) }, segFixed, edited, marked, notes, ...rep };
  } catch (e) { await calc("auto"); return { error: String(e.message || e), hint: /InvalidOperation|edit/i.test(String(e)) ? "Excel đang ở chế độ sửa ô: nhờ người dùng bấm Enter/Esc" : undefined, notes }; }
};

// ---------- NEW ITEMS: nhiều hạng mục, tính lại 1 lần ----------
// p = {items:[{src, boqRow, edits}], mark?, collapse?}
L.newItems = async (p) => {
  const done = [];
  for (const it of p.items) {
    const r = await L.newItem({ mark: p.mark, collapse: p.collapse, boqSheet: p.boqSheet, ...it, _batch: true });
    done.push(r); if (r.error) break;
  }
  await calc("auto"); wb.application.calculate(Excel.CalculationType.full); await sync();
  const B = await findBoq(p.boqSheet), out = [];
  for (const r of done) {
    if (r.error) { out.push(r); continue; }
    const rep = await report(CT, r.block[0], r.block[1]);
    const bq = sh(B.name).getRange(`A${r.boqRow}:${B.amt}${r.boqRow}`); bq.load("values"); await sync();
    const bv = bq.values[0];
    out.push({ block: r.block, boq: { row: r.boqRow, code: bv[0], qty: rnd(bv[colN(B.qty)]), price: rnd(bv[colN(B.price)]), amount: rnd(bv[colN(B.amt)]) }, notes: r.notes, edited: r.edited, price: rep.price, errors: rep.errors, rows: p.rows ? rep.rows : undefined });
  }
  return out;
};

// ---------- EDIT: sửa block có sẵn ----------
// p = {hdr|code, edits, mark?=true}
L.edit = async (p) => {
  await calc("manual");
  try {
    const hdr = await resolveHdr(p), end = await blockEnd(CT, hdr);
    const a = sh(CT).getRange(`A${hdr}`); a.load("formulas"); await sync();
    const m = String(a.formulas[0][0]).match(/^='?([^'!]+)'?!\$?A\$?(\d+)/);
    const edited = await applyEdits(CT, hdr, end, p.edits, { boqRow: m ? +m[2] : null, boqSheet: m ? m[1] : null });
    const marked = p.mark === false ? 0 : await mark(edited.map(x => [CT, x]), "edit " + hdr);
    await calc("auto"); wb.application.calculate(Excel.CalculationType.full); await sync();
    return { block: [hdr, end], edited, marked, ...(await report(CT, hdr, end)) };
  } catch (e) { await calc("auto"); return { error: String(e.message || e) }; }
};

// ---------- DESC: đọc/ghi mô tả BoQ ----------
// p = {get:["D01"]} | {items:[{row|code, text}], mark?=false, boqSheet?}
L.desc = async (p) => {
  const B = await findBoq(p.boqSheet), s = sh(B.name), u = s.getUsedRange(true); u.load("rowCount,rowIndex"); await sync();
  const last = u.rowIndex + u.rowCount, a = s.getRange(`A1:A${last}`); a.load("values"); await sync();
  const rowOf = (code) => { const i = a.values.findIndex(r => norm(r[0]) === norm(code)); if (i < 0) throw new Error("Không thấy mã " + code); return i + 1; };
  if (p.get) { const rs = p.get.map(c => { const r = s.getRange(B.desc + rowOf(c)); r.load("values"); return [c, r]; }); await sync(); return Object.fromEntries(rs.map(([c, r]) => [c, r.values[0][0]])); }
  const rows = p.items.map(it => [it.row || rowOf(it.code), it.text]), targets = new Set(rows.map(r => r[0]));
  const top = B.hdrRow ? B.hdrRow + 1 : 1, fr = s.getRange(`${B.desc}${top}:${B.desc}${last}`); fr.load("values"); await sync();
  let refRow = 0; fr.values.forEach((v, i) => { if (!refRow && v[0] !== "" && !targets.has(top + i)) refRow = top + i; }); // ô mô tả có sẵn để lấy định dạng
  for (const [r, t] of rows) { const c = s.getRange(B.desc + r); if (refRow) c.copyFrom(s.getRange(B.desc + refRow), Excel.RangeCopyType.formats); c.values = [[t]]; c.format.wrapText = true; }
  await sync();
  const marked = p.mark ? await mark(rows.map(([r]) => [B.name, B.desc + r]), "desc") : 0;
  const chk = rows.map(([r]) => { const c = s.getRange(B.desc + r); c.load("values"); return [r, c]; }); await sync();
  return { col: B.desc, written: chk.map(([r, c]) => [r, String(c.values[0][0]).slice(0, 50)]), marked };
};

// ---------- RULES: thêm quy tắc vào tblRules ----------
// p = {rows:[{topic, rule, source?, scope?:"RIENG"|"CHUNG"}]}
L.addRules = async (p) => {
  const t = wb.tables.getItem("tblRules"), hd = t.getHeaderRowRange(), b = t.getDataBodyRange(); hd.load("values"); b.load("values"); await sync();
  const H = hd.values[0].map(norm), today = Math.floor(Date.now() / 86400000) + 25569;
  const have = new Set(b.values.map(r => norm(r[H.indexOf("quy tắc")])));
  const add = p.rows.filter(x => !have.has(norm(x.rule))).map(x => H.map(h => ({ "ngày": today, "chủ đề": x.topic, "quy tắc": x.rule, "nguồn": x.source || "Claude", "trạng thái": "Pending", "phạm vi": x.scope || "RIENG" })[h] ?? ""));
  if (add.length) { t.rows.add(null, add); await sync(); }
  return { added: add.length, skippedDuplicate: p.rows.length - add.length };
};

return L;
