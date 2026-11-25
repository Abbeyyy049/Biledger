// Excel export (a tidy workbook plus a hidden full backup) and import.
// ExcelJS is loaded only when one of these is used.
const EXCELJS = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
let loading;
function lib() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  loading ??= new Promise((ok, no) => {
    const s = document.createElement('script');
    s.src = EXCELJS; s.onload = () => ok(window.ExcelJS); s.onerror = () => { loading = null; no(new Error('Excel 工具載入失敗，請檢查網路')); };
    document.head.appendChild(s);
  });
  return loading;
}

const GREEN = 'FF1D6A50', SOFT = 'FFE8F1EC', LINE = 'FFDDE3DD';
const FONT = 'Arial';
const DATE = 'yyyy/mm/dd';
const BACKUP = '_app';
const utc = s => { const [y, m, d] = s.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const isoOf = v => {
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);
  if (typeof v === 'string' && /^\d{4}[/-]\d{1,2}[/-]\d{1,2}$/.test(v.trim())) { const [y, m, d] = v.trim().split(/[/-]/).map(Number); return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`; }
  return '';
};
const val = c => { const v = c?.value; if (v && typeof v === 'object' && !(v instanceof Date)) return 'result' in v ? v.result : v.text ?? (v.richText ? v.richText.map(r => r.text).join('') : ''); return v ?? ''; };
const txt = c => String(val(c) ?? '').trim();
const r2 = n => Math.round(n * 100) / 100;
const q = t => String(t).replace(/"/g, '""');

function header(ws, row, cols) {
  cols.forEach(([h, w], i) => {
    const c = ws.getCell(row, i + 1);
    c.value = h; c.font = { name: FONT, bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
    c.alignment = { vertical: 'middle', horizontal: i ? 'center' : 'left' };
    ws.getColumn(i + 1).width = w;
  });
  ws.getRow(row).height = 22;
}
const cellStyle = (c, fmt, opts = {}) => {
  c.font = { name: FONT, size: 11, bold: !!opts.bold };
  c.border = { bottom: { style: 'thin', color: { argb: LINE } } };
  if (fmt) c.numFmt = fmt;
  if (opts.center) c.alignment = { horizontal: 'center' };
  if (opts.soft) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: SOFT } };
};
const titleCell = (c, text) => { c.value = text; c.font = { name: FONT, size: 14, bold: true }; };
const note = (c, text) => { c.value = text; c.font = { name: FONT, size: 10, color: { argb: 'FF6B7570' } }; };

export async function exportWorkbook(X) {
  const ExcelJS = await lib();
  const { S, CYC, NONE, NONE_LABEL } = X;
  const [, sym, dp] = X.cur, num = `"${sym}"#,##0${dp ? '.' + '0'.repeat(dp) : ''}`, MONEY = `${num};-${num};"-"`;
  const BUY = X.buyName;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Biledger';
  const aName = id => id ? X.assetOf(id).name : null;
  const cName = sub => X.catOf(sub)?.name || '';
  const src = e => e.src === NONE ? NONE_LABEL : (X.jarOf(e.src)?.name || NONE_LABEL);
  const paid = S.expenses.filter(e => e.status === 'paid');

  /* 開銷明細 */
  const ws = wb.addWorksheet('開銷明細', { views: [{ state: 'frozen', xSplit: 1, ySplit: 2, showGridLines: false }] });
  const cols = [['日期', 12], ['資產', 24], ['大類', 11], ['類別', 14], ['名稱', 22], ['金額', 13], ['使用期間（起）', 15], ['使用期間（迄）', 15], ['預付/後付', 10], ['從哪裡扣', 10], ['付款方式', 20], ['週期', 8], ['狀態', 8], ['備註', 40], ['收支', 8]];
  const ex = [...S.expenses].sort((a, b) => a.date < b.date ? 1 : a.date > b.date ? -1 : 0);
  const first = 3, last = 2 + Math.max(ex.length, 1);
  ws.getCell('A1').value = '篩選後合計'; ws.getCell('A1').font = { name: FONT, bold: true };
  const f1 = ws.getCell('F1');
  f1.value = { formula: `SUBTOTAL(109,F${first}:F${last})`, result: r2(ex.reduce((t, e) => t + X.sv(e), 0)) };
  f1.font = { name: FONT, bold: true }; f1.numFmt = MONEY;
  note(ws.getCell('G1'), '支出是負數、收入是正數；用表頭的 ▼ 篩選，合計會跟著變');
  header(ws, 2, cols);
  ex.forEach((e, i) => {
    const r = first + i;
    const v = [utc(e.date), aName(e.asset), cName(e.sub), X.subName(e.sub), e.name || X.subName(e.sub), X.sv(e),
      e.ps ? utc(e.ps) : null, e.pe ? utc(e.pe) : null, X.timing(e), src(e), e.method || '', CYC[e.cycle]?.[0] || '', e.status === 'paid' ? '已付' : '待繳', e.note || '', X.isIn(e) ? '收入' : '支出'];
    v.forEach((x, j) => { const c = ws.getCell(r, j + 1); c.value = x; cellStyle(c, [0, 6, 7].includes(j) ? DATE : j === 5 ? MONEY : null, { center: [0, 6, 7, 8, 9, 11, 12, 14].includes(j) }); });
  });
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: last, column: cols.length } };
  const R = c => `開銷明細!$${c}$${first}:$${c}$${last}`;
  const AMT = R('F'), DT = R('A'), AS = R('B'), CAT = R('C'), CYCR = R('L'), ST = R('M'), TYPE = R('O');

  /* 存錢罐 */
  for (const j of S.jars) {
    const w = wb.addWorksheet(`存錢罐-${j.name}`.slice(0, 31).replace(/[\\/?*[\]:]/g, ''), { views: [{ state: 'frozen', ySplit: 4, showGridLines: false }] });
    const rows = [
      ...S.moves.filter(m => m.jar === j.id).map(m => ({ d: m.date, lab: m.note || (m.kind === 'in' ? '存入' : '調整'), a: '', i: m.kind === 'in' ? +m.amount : null, o: m.kind === 'in' ? null : +m.amount })),
      ...paid.filter(e => e.src === j.id).map(e => ({ d: e.date, lab: e.name || X.subName(e.sub), a: aName(e.asset), i: X.isIn(e) ? +e.amount : null, o: X.isIn(e) ? null : +e.amount })),
    ].sort((a, b) => a.d < b.d ? -1 : a.d > b.d ? 1 : (a.i == null) - (b.i == null));
    titleCell(w.getCell('A1'), `${j.name}・存錢罐`);
    note(w.getCell('A2'), j.desc || '');
    w.getCell('D1').value = '目前餘額'; w.getCell('D1').font = { name: FONT, bold: true }; w.getCell('D1').alignment = { horizontal: 'right' };
    note(w.getCell('E2'), '負數 = 我先墊付的錢');
    header(w, 4, [['日期', 12], ['說明', 28], ['資產', 24], ['存入', 13], ['扣除', 13], ['餘額', 14]]);
    let bal = 0;
    rows.forEach((x, k) => {
      const r = 5 + k;
      [utc(x.d), x.lab, x.a, x.i, x.o].forEach((v, c) => { const cell = w.getCell(r, c + 1); cell.value = v; cellStyle(cell, c === 0 ? DATE : c > 2 ? MONEY : null, { center: c === 0 }); });
      bal = r2(bal + (x.i || 0) - (x.o || 0));
      const b = w.getCell(r, 6); b.value = { formula: k ? `F${r - 1}+D${r}-E${r}` : `D${r}-E${r}`, result: bal }; cellStyle(b, MONEY, { bold: true });
    });
    const e1 = w.getCell('E1'); e1.value = rows.length ? { formula: `F${4 + rows.length}`, result: bal } : 0; e1.font = { name: FONT, size: 14, bold: true }; e1.numFmt = MONEY;
  }

  /* 月度統計 */
  const order = S.cats.filter(c => c.kind !== 'in' && c.name !== BUY).map(c => c.name);
  const m = wb.addWorksheet('月度統計', { views: [{ state: 'frozen', xSplit: 1, ySplit: 3, showGridLines: false }] });
  titleCell(m.getCell('A1'), '每月收支（按付款日，只算已付；支出是負數）');
  header(m, 3, [['月份', 11], ...order.map(c => [c, 12]), [`小計（不含${BUY}）`, 17], [BUY, 14], ['支出合計', 14], ['收入', 14], ['收支淨額', 14]]);
  const ds = S.expenses.map(e => e.date).sort();
  if (ds.length) {
    let [y, mo] = ds[0].slice(0, 7).split('-').map(Number); const [ey, em] = ds.at(-1).slice(0, 7).split('-').map(Number);
    const months = []; while (y < ey || (y === ey && mo <= em)) { months.push([y, mo]); mo++; if (mo > 12) { mo = 1; y++; } }
    months.reverse().forEach(([y, mo], k) => {
      const r = 4 + k, key = `${y}-${String(mo).padStart(2, '0')}`, ny = mo === 12 ? y + 1 : y, nm = mo === 12 ? 1 : mo + 1;
      const rng = `${DT},">="&DATE(${y},${mo},1),${DT},"<"&DATE(${ny},${nm},1)`;
      const inMonth = paid.filter(e => e.date.startsWith(key));
      const c0 = m.getCell(r, 1); c0.value = `${y}/${String(mo).padStart(2, '0')}`; cellStyle(c0);
      let subtotal = 0;
      order.forEach((name, i) => { const v = r2(inMonth.filter(e => cName(e.sub) === name && !X.isIn(e)).reduce((t, e) => t + X.sv(e), 0)); subtotal += v;
        const c = m.getCell(r, i + 2); c.value = { formula: `SUMIFS(${AMT},${rng},${CAT},"${q(name)}",${ST},"已付",${TYPE},"支出")`, result: v }; cellStyle(c, MONEY); });
      const n = order.length, L = k2 => String.fromCharCode(64 + k2);
      const buy = r2(inMonth.filter(e => cName(e.sub) === BUY && !X.isIn(e)).reduce((t, e) => t + X.sv(e), 0)), inc = r2(inMonth.filter(X.isIn).reduce((t, e) => t + X.sv(e), 0));
      let c = m.getCell(r, n + 2); c.value = { formula: `SUM(B${r}:${L(n + 1)}${r})`, result: r2(subtotal) }; cellStyle(c, MONEY, { bold: true, soft: true });
      c = m.getCell(r, n + 3); c.value = { formula: `SUMIFS(${AMT},${rng},${CAT},"${q(BUY)}",${ST},"已付",${TYPE},"支出")`, result: buy }; cellStyle(c, MONEY);
      c = m.getCell(r, n + 4); c.value = { formula: `${L(n + 2)}${r}+${L(n + 3)}${r}`, result: r2(subtotal + buy) }; cellStyle(c, MONEY, { bold: true });
      c = m.getCell(r, n + 5); c.value = { formula: `SUMIFS(${AMT},${rng},${TYPE},"收入",${ST},"已付")`, result: inc }; cellStyle(c, MONEY);
      c = m.getCell(r, n + 6); c.value = { formula: `${L(n + 4)}${r}+${L(n + 5)}${r}`, result: r2(subtotal + buy + inc) }; cellStyle(c, MONEY, { bold: true, soft: true });
    });
  }

  /* 年度統計：固定開銷 / 一次性開銷 / 購買 */
  const y = wb.addWorksheet('年度統計', { views: [{ showGridLines: false }] });
  titleCell(y.getCell('A1'), '每年收支・依資產（按付款日，只算已付；支出是負數）');
  const years = [...new Set(ds.map(d => d.slice(0, 4)))].sort().reverse();
  // expenses without an asset get their own column
  const assets = [...S.assets, ...(S.expenses.some(e => !e.asset) ? [{ id: '', name: '未指定資產' }] : [])];
  let row = 3;
  const block = (label, hint, crit, test) => {
    y.getCell(row, 1).value = label; y.getCell(row, 1).font = { name: FONT, size: 12, bold: true };
    note(y.getCell(row + 1, 1), hint);
    header(y, row + 2, [['年份', 10], ...assets.map(a => [a.name, 20]), ['合計', 14]]);
    years.forEach((yy, k) => {
      const r = row + 3 + k;
      const c0 = y.getCell(r, 1); c0.value = yy; cellStyle(c0);
      let tot = 0;
      assets.forEach((a, i) => {
        const v = r2(paid.filter(e => e.date.startsWith(yy) && (e.asset || '') === a.id && test(e)).reduce((t, e) => t + X.sv(e), 0)); tot += v;
        const c = y.getCell(r, i + 2);
        c.value = { formula: `SUMIFS(${AMT},${DT},">="&DATE(${yy},1,1),${DT},"<"&DATE(${+yy + 1},1,1),${AS},"${a.id ? q(a.name) : '='}",${ST},"已付"${crit})`, result: v };
        cellStyle(c, MONEY);
      });
      const n = assets.length, L = String.fromCharCode(65 + n);
      const c = y.getCell(r, n + 2); c.value = { formula: `SUM(B${r}:${L}${r})`, result: r2(tot) }; cellStyle(c, MONEY, { bold: true, soft: true });
    });
    row += 4 + years.length;
  };
  const once = CYC.once[0];
  block('每年固定開銷', '週期是每月、每季、每半年或每年的帳單，例如水電網路、電信、管理費、房屋稅、保險', `,${CYCR},"<>${once}",${TYPE},"支出"`, e => e.cycle !== 'once' && !X.isIn(e));
  block(`一次性開銷（不含${BUY}）`, '週期是「一次性」的花費，例如維修、規費、印花稅、過戶費', `,${CYCR},"${once}",${CAT},"<>${q(BUY)}",${TYPE},"支出"`, e => e.cycle === 'once' && cName(e.sub) !== BUY && !X.isIn(e));
  block(BUY, '購屋、購車、訂金、家具家電等大筆購置', `,${CAT},"${q(BUY)}",${TYPE},"支出"`, e => cName(e.sub) === BUY && !X.isIn(e));
  if (paid.some(X.isIn)) block('收入', '收入類別的所有紀錄，例如薪資、租金收入、利息', `,${TYPE},"收入"`, X.isIn);

  /* 清單 */
  const l = wb.addWorksheet('清單', { views: [{ showGridLines: false }] });
  titleCell(l.getCell('A1'), '資產');
  header(l, 2, [['資料夾', 12], ['資產', 26], ['暱稱', 20], ['狀態', 8]]);
  let r = 3;
  for (const f of S.folders) for (const a of S.assets.filter(a => a.folder === f.id)) {
    [f.name, a.name, a.nick || '', a.archived ? '封存' : '使用中'].forEach((v, i) => { const c = l.getCell(r, i + 1); c.value = v; cellStyle(c); }); r++;
  }
  titleCell(l.getCell('F1'), '類別');
  [['大類', 12], ['類別', 22]].forEach(([h, w], i) => { const c = l.getCell(2, 6 + i); c.value = h; c.font = { name: FONT, bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } }; l.getColumn(6 + i).width = w; });
  r = 3; for (const c of S.cats) for (const s of c.subs) { [c.name, s.name].forEach((v, i) => { const x = l.getCell(r, 6 + i); x.value = v; cellStyle(x); }); r++; }
  titleCell(l.getCell('I1'), '付款方式');
  const mh = l.getCell(2, 9); mh.value = '付款方式'; mh.font = { name: FONT, bold: true, color: { argb: 'FFFFFFFF' } }; mh.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } }; l.getColumn(9).width = 24;
  S.methods.forEach((x, i) => { const c = l.getCell(3 + i, 9); c.value = x; cellStyle(c); });
  l.getColumn(5).width = 3; l.getColumn(8).width = 3;

  /* hidden full backup, read back by import */
  const bk = wb.addWorksheet(BACKUP, { state: 'veryHidden' });
  const cfg = { v: S.v || 3 }; ['folders', 'assets', 'cats', 'methods', 'fields', 'jars', 'bills', 'buffer', 'currency', 'font'].forEach(k => cfg[k] = S[k]);
  const json = JSON.stringify({ app: 'house-car-ledger', v: cfg.v, cfg, expenses: S.expenses, moves: S.moves });
  for (let i = 0, k = 1; i < json.length; i += 30000, k++) bk.getCell(k, 1).value = json.slice(i, i + 30000);

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

export async function parseWorkbook(file, X) {
  const ExcelJS = await lib();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());

  const bk = wb.getWorksheet(BACKUP);
  if (bk) {
    let s = ''; for (let k = 1; k <= bk.rowCount; k++) s += txt(bk.getCell(k, 1)) ? String(val(bk.getCell(k, 1))) : '';
    const d = JSON.parse(s);
    if (d.app !== 'house-car-ledger') throw new Error('不是 Biledger 匯出的備份');
    return { cfg: { ...X.defaultConfig(), ...d.cfg }, expenses: d.expenses || [], moves: d.moves || [], backup: true };
  }

  // No backup sheet: read the visible tables, matching names to the current setup.
  const ws = wb.getWorksheet('開銷明細');
  if (!ws) throw new Error('找不到「開銷明細」分頁');
  const cfg = JSON.parse(JSON.stringify({ ...X.defaultConfig(), ...pick(X.S) }));
  const folderFor = name => { let f = cfg.folders.find(f => f.name === name); if (!f) { f = { id: X.uid('f'), name }; cfg.folders.push(f); } return f; };
  const assetFor = (name, folder, nick) => {
    let a = cfg.assets.find(a => a.name === name);
    if (!a) { a = { id: X.uid('a'), name, nick: nick || '', folder: (folder ? folderFor(folder) : cfg.folders[0] || folderFor('個人')).id, kind: /mazda|toyota|honda|tesla|car|車/i.test(name) ? 'car' : 'house' }; cfg.assets.push(a); }
    return a;
  };
  const subFor = (catName, subName, income) => {
    let c = cfg.cats.find(c => c.name === catName) || cfg.cats.find(c => c.subs.some(s => s.name === subName));
    if (!c) { c = { id: X.uid('c'), name: catName || (income ? '收入' : '其他'), subs: [], ...(income ? { kind: 'in' } : {}) }; cfg.cats.push(c); }
    let s = c.subs.find(s => s.name === subName);
    if (!s) { s = { id: X.uid('s'), name: subName || '其他' }; c.subs.push(s); }
    return s;
  };
  const jarFor = name => { if (!name || name === X.NONE_LABEL) return X.NONE; let j = cfg.jars.find(j => j.name === name); if (!j) { j = { id: X.uid('jar'), name, desc: '', alert: 500, pin: true }; cfg.jars.push(j); } return j.id; };
  const cycOf = label => Object.entries(X.CYC).find(([, [l]]) => l === label)?.[0] || 'once';

  const list = wb.getWorksheet('清單');
  if (list) for (let r = 3; r <= list.rowCount; r++) {
    const f = txt(list.getCell(r, 1)), a = txt(list.getCell(r, 2));
    if (a) { const x = assetFor(a, f, txt(list.getCell(r, 3))); if (txt(list.getCell(r, 4)) === '封存') x.archived = true; }
    const c = txt(list.getCell(r, 6)), s = txt(list.getCell(r, 7)); if (c && s) subFor(c, s);
    const m = txt(list.getCell(r, 9)); if (m && !cfg.methods.includes(m)) cfg.methods.push(m);
  }

  const expenses = [];
  for (let r = 3; r <= ws.rowCount; r++) {
    const date = isoOf(val(ws.getCell(r, 1))), amount = Math.abs(+val(ws.getCell(r, 6))), income = txt(ws.getCell(r, 15)) === '收入';
    if (!date || !amount) continue;
    const sub = subFor(txt(ws.getCell(r, 3)), txt(ws.getCell(r, 4)), income);
    const name = txt(ws.getCell(r, 5)), method = txt(ws.getCell(r, 11));
    if (method && !cfg.methods.includes(method)) cfg.methods.push(method);
    const due = txt(ws.getCell(r, 13)) === '待繳';
    expenses.push({ _in: income, id: X.uid('e'), date, asset: txt(ws.getCell(r, 2)) ? assetFor(txt(ws.getCell(r, 2))).id : '', sub: sub.id, name: name === sub.name ? '' : name,
      amount: r2(amount), ps: isoOf(val(ws.getCell(r, 7))) || '', pe: isoOf(val(ws.getCell(r, 8))) || '', src: jarFor(txt(ws.getCell(r, 10))), method,
      cycle: cycOf(txt(ws.getCell(r, 12))), status: due ? 'due' : 'paid', due: due ? date : '', note: txt(ws.getCell(r, 14)), extra: {} });
  }

  const moves = [];
  wb.eachSheet(w => {
    if (!w.name.startsWith('存錢罐-')) return;
    const jid = jarFor(w.name.slice(4));
    for (let r = 5; r <= w.rowCount; r++) {
      const date = isoOf(val(w.getCell(r, 1))), inn = +val(w.getCell(r, 4)) || 0, out = +val(w.getCell(r, 5)) || 0, lab = txt(w.getCell(r, 2));
      if (!date || !(inn || out)) continue;
      const own = (amt, inc) => expenses.find(e => e.src === jid && e.date === date && Math.abs(e.amount - amt) < 0.005 && !!e._in === inc && !e._jar);
      if (inn) { const hit = own(inn, true); if (hit) hit._jar = true; else moves.push({ id: X.uid('m'), jar: jid, date, kind: 'in', amount: r2(inn), note: lab }); continue; }
      const hit = own(out, false);
      if (hit) hit._jar = true; else moves.push({ id: X.uid('m'), jar: jid, date, kind: 'out', amount: r2(out), note: lab });
    }
  });
  expenses.forEach(e => { delete e._jar; delete e._in; });
  return { cfg, expenses, moves, backup: false };
}

function pick(S) { const c = {}; ['folders', 'assets', 'cats', 'methods', 'fields', 'jars', 'bills', 'buffer', 'currency', 'font'].forEach(k => c[k] = S[k]); return c; }
