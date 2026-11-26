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

  const fs = +X.fyStart || 1, fyOf = d => (+d.slice(5, 7) >= fs ? +d.slice(0, 4) : +d.slice(0, 4) - 1);
  const fyName = fy => fs === 1 ? `${fy} 年` : `${fy}–${String(fy + 1).slice(2)} 財年`;
  const ds = S.expenses.map(e => e.date).sort();
  const years = [...new Set(ds.map(fyOf))].sort().reverse();
  const L = n => { let t = ''; for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) t = String.fromCharCode(65 + (n - 1) % 26) + t; return t; }; // 0 -> A
  const ov = wb.addWorksheet('總覽', { views: [{ showGridLines: false }] });

  /* 收支明細: one row per record, income and spending in separate columns */
  const ws = wb.addWorksheet('收支明細', { views: [{ state: 'frozen', xSplit: 1, ySplit: 2, showGridLines: false }] });
  const cols = [['日期', 12], ['資產', 24], ['大類', 11], ['類別', 14], ['名稱', 22], ['收入', 13], ['支出', 13], ['使用期間（起）', 15], ['使用期間（迄）', 15], ['預付/後付', 10], ['存錢罐', 10], ['付款方式', 20], ['週期', 8], ['狀態', 8], ['備註', 40]];
  const ex = [...S.expenses].sort((a, b) => a.date < b.date ? 1 : a.date > b.date ? -1 : 0);
  const first = 3, last = 2 + Math.max(ex.length, 1);
  const incAll = r2(ex.filter(X.isIn).reduce((t, e) => t + +e.amount, 0)), outAll = r2(ex.filter(e => !X.isIn(e)).reduce((t, e) => t + +e.amount, 0));
  ws.getCell('A1').value = '篩選後合計'; ws.getCell('A1').font = { name: FONT, bold: true };
  [['F1', `SUBTOTAL(109,F${first}:F${last})`, incAll], ['G1', `SUBTOTAL(109,G${first}:G${last})`, outAll], ['I1', 'F1-G1', r2(incAll - outAll)]].forEach(([a, f, v]) => { const c = ws.getCell(a); c.value = { formula: f, result: v }; c.font = { name: FONT, bold: true }; c.numFmt = MONEY; });
  ws.getCell('H1').value = '淨額（收入−支出）'; ws.getCell('H1').font = { name: FONT, bold: true }; ws.getCell('H1').alignment = { horizontal: 'right' };
  note(ws.getCell('J1'), '用表頭的 ▼ 篩選，上面的合計會跟著變');
  header(ws, 2, cols);
  ex.forEach((e, i) => {
    const r = first + i, inc = X.isIn(e);
    const v = [utc(e.date), aName(e.asset), cName(e.sub), X.subName(e.sub), e.name || X.subName(e.sub), inc ? +e.amount : null, inc ? null : +e.amount,
      e.ps ? utc(e.ps) : null, e.pe ? utc(e.pe) : null, X.timing(e), src(e), e.method || '', CYC[e.cycle]?.[0] || '', e.status === 'paid' ? '已付' : '待繳', e.note || ''];
    v.forEach((x, j) => { const c = ws.getCell(r, j + 1); c.value = x; cellStyle(c, [0, 7, 8].includes(j) ? DATE : (j === 5 || j === 6) ? MONEY : null, { center: [0, 7, 8, 9, 10, 12, 13].includes(j) }); });
  });
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: last, column: cols.length } };
  const R = c => `收支明細!$${c}$${first}:$${c}$${last}`;
  const INC = R('F'), OUT = R('G'), DT = R('A'), AS = R('B'), CAT = R('C'), CYCR = R('M'), ST = R('N');
  const paidIn = (list) => r2(list.filter(X.isIn).reduce((t, e) => t + +e.amount, 0)), paidOut = (list) => r2(list.filter(e => !X.isIn(e)).reduce((t, e) => t + +e.amount, 0));

  /* 存錢罐 */
  const jarRows = [];
  for (const j of S.jars) {
    const sheetName = `存錢罐-${j.name}`.slice(0, 31).replace(/[\\/?*[\]:]/g, '');
    const w = wb.addWorksheet(sheetName, { views: [{ state: 'frozen', ySplit: 4, showGridLines: false }] });
    const rows = [
      ...S.moves.filter(m => m.jar === j.id).map(m => ({ d: m.date, lab: m.note || (m.kind === 'in' ? '存入' : '調整'), a: '', i: m.kind === 'in' ? +m.amount : null, o: m.kind === 'in' ? null : +m.amount })),
      ...paid.filter(e => e.src === j.id).map(e => ({ d: e.date, lab: e.name || X.subName(e.sub), a: aName(e.asset), i: X.isIn(e) ? +e.amount : null, o: X.isIn(e) ? null : +e.amount })),
    ].sort((a, b) => a.d < b.d ? -1 : a.d > b.d ? 1 : (a.i == null) - (b.i == null));
    titleCell(w.getCell('A1'), `${j.name}・存錢罐`);
    note(w.getCell('A2'), j.desc || '');
    w.getCell('D1').value = '目前餘額'; w.getCell('D1').font = { name: FONT, bold: true }; w.getCell('D1').alignment = { horizontal: 'right' };
    note(w.getCell('E2'), '負數代表餘額不足');
    header(w, 4, [['日期', 12], ['說明', 28], ['資產', 24], ['存入', 13], ['扣除', 13], ['餘額', 14]]);
    let bal = 0;
    rows.forEach((x, k) => {
      const r = 5 + k;
      [utc(x.d), x.lab, x.a, x.i, x.o].forEach((v, c) => { const cell = w.getCell(r, c + 1); cell.value = v; cellStyle(cell, c === 0 ? DATE : c > 2 ? MONEY : null, { center: c === 0 }); });
      bal = r2(bal + (x.i || 0) - (x.o || 0));
      const b = w.getCell(r, 6); b.value = { formula: k ? `F${r - 1}+D${r}-E${r}` : `D${r}-E${r}`, result: bal }; cellStyle(b, MONEY, { bold: true });
    });
    const e1 = w.getCell('E1'); e1.value = rows.length ? { formula: `F${4 + rows.length}`, result: bal } : 0; e1.font = { name: FONT, size: 14, bold: true }; e1.numFmt = MONEY;
    jarRows.push({ name: j.name, ref: `'${sheetName}'!E1`, bal });
  }

  /* 月度統計: spending by main category, then totals, income and net */
  const order = S.cats.filter(c => c.kind !== 'in' && c.name !== BUY).map(c => c.name);
  const m = wb.addWorksheet('月度統計', { views: [{ state: 'frozen', xSplit: 1, ySplit: 3, showGridLines: false }] });
  titleCell(m.getCell('A1'), '每月收支（按付款日，只算已付）');
  header(m, 3, [['月份', 11], ...order.map(c => [c, 12]), [`支出小計（不含${BUY}）`, 19], [BUY, 13], ['支出合計', 14], ['收入合計', 14], ['淨額', 14]]);
  if (ds.length) {
    let [y, mo] = ds[0].slice(0, 7).split('-').map(Number); const [ey, em] = ds.at(-1).slice(0, 7).split('-').map(Number);
    const months = []; while (y < ey || (y === ey && mo <= em)) { months.push([y, mo]); mo++; if (mo > 12) { mo = 1; y++; } }
    months.reverse().forEach(([y, mo], k) => {
      const r = 4 + k, key = `${y}-${String(mo).padStart(2, '0')}`, ny = mo === 12 ? y + 1 : y, nm = mo === 12 ? 1 : mo + 1;
      const rng = `${DT},">="&DATE(${y},${mo},1),${DT},"<"&DATE(${ny},${nm},1),${ST},"已付"`;
      const inMonth = paid.filter(e => e.date.startsWith(key));
      const c0 = m.getCell(r, 1); c0.value = `${y}/${String(mo).padStart(2, '0')}`; cellStyle(c0);
      let sub = 0;
      order.forEach((name, i) => { const v = paidOut(inMonth.filter(e => cName(e.sub) === name)); sub += v;
        const c = m.getCell(r, i + 2); c.value = { formula: `SUMIFS(${OUT},${rng},${CAT},"${q(name)}")`, result: v }; cellStyle(c, MONEY); });
      const n = order.length, buy = paidOut(inMonth.filter(e => cName(e.sub) === BUY)), inc = paidIn(inMonth);
      const put = (col, f, v, o) => { const c = m.getCell(r, col); c.value = { formula: f, result: r2(v) }; cellStyle(c, MONEY, o); };
      put(n + 2, `SUM(B${r}:${L(n)}${r})`, sub, { bold: true, soft: true });
      put(n + 3, `SUMIFS(${OUT},${rng},${CAT},"${q(BUY)}")`, buy);
      put(n + 4, `${L(n + 1)}${r}+${L(n + 2)}${r}`, sub + buy, { bold: true });
      put(n + 5, `SUMIFS(${INC},${rng})`, inc);
      put(n + 6, `${L(n + 4)}${r}-${L(n + 3)}${r}`, inc - sub - buy, { bold: true, soft: true });
    });
  }

  /* 年度統計：固定開銷 / 一次性開銷 / 購置 / 收入 (by financial year) */
  const y = wb.addWorksheet('年度統計', { views: [{ showGridLines: false }] });
  titleCell(y.getCell('A1'), '每年收支・依資產（按付款日，只算已付）');
  const assets = [...S.assets, ...(S.expenses.some(e => !e.asset) ? [{ id: '', name: '未指定資產' }] : [])];
  let row = 3;
  const block = (label, hint, col, crit, test) => {
    y.getCell(row, 1).value = label; y.getCell(row, 1).font = { name: FONT, size: 12, bold: true };
    note(y.getCell(row + 1, 1), hint);
    header(y, row + 2, [['年度', fs === 1 ? 10 : 15], ...assets.map(a => [a.name, 20]), ['合計', 14]]);
    years.forEach((yy, k) => {
      const r = row + 3 + k;
      const c0 = y.getCell(r, 1); c0.value = fyName(yy); cellStyle(c0);
      let tot = 0;
      assets.forEach((a, i) => {
        const v = r2(paid.filter(e => fyOf(e.date) === yy && (e.asset || '') === a.id && test(e)).reduce((t, e) => t + +e.amount, 0)); tot += v;
        const c = y.getCell(r, i + 2);
        c.value = { formula: `SUMIFS(${col},${DT},">="&DATE(${yy},${fs},1),${DT},"<"&DATE(${yy + 1},${fs},1),${AS},"${a.id ? q(a.name) : '='}",${ST},"已付"${crit})`, result: v };
        cellStyle(c, MONEY);
      });
      const n = assets.length;
      const c = y.getCell(r, n + 2); c.value = { formula: `SUM(B${r}:${L(n)}${r})`, result: r2(tot) }; cellStyle(c, MONEY, { bold: true, soft: true });
    });
    row += 4 + years.length;
  };
  const once = CYC.once[0];
  block('每年固定開銷', '週期是每月、每季、每半年或每年的帳單，例如水電網路、電信、管理費、房屋稅、保險', OUT, `,${CYCR},"<>${once}"`, e => e.cycle !== 'once' && !X.isIn(e));
  block(`一次性開銷（不含${BUY}）`, '週期是「一次性」的花費，例如維修、規費、印花稅、過戶費', OUT, `,${CYCR},"${once}",${CAT},"<>${q(BUY)}"`, e => e.cycle === 'once' && cName(e.sub) !== BUY && !X.isIn(e));
  block(BUY, '購屋、購車、訂金、家具家電等大筆購置', OUT, `,${CAT},"${q(BUY)}"`, e => cName(e.sub) === BUY && !X.isIn(e));
  if (paid.some(X.isIn)) block('收入', '收入類別的所有紀錄，例如薪資、租金收入、利息', INC, '', X.isIn);

  /* 預算: the app's forecast for the current year and the savings plan (values, not formulas) */
  const BG = X.budget;
  if (BG) {
    const b = wb.addWorksheet('預算', { views: [{ showGridLines: false }] });
    titleCell(b.getCell('A1'), `${BG.yearName}預算`);
    note(b.getCell('A2'), `${BG.range}・由 Biledger 依固定帳單和歷史紀錄計算（不是 AI），預估方法寫在每一列`);
    const kv = (r, k, v, money = true) => { const a = b.getCell(r, 1), c = b.getCell(r, 2); a.value = k; a.font = { name: FONT, size: 11 }; c.value = v; cellStyle(c, money ? MONEY : null, { bold: true }); };
    kv(4, '年度預算', BG.budget); kv(5, '預估全年（實際＋預估剩餘）', BG.projected); kv(6, '實際已付', BG.actual); kv(7, '預估剩餘', BG.remaining);
    kv(8, '預算差異（正數＝在預算內）', BG.variance); kv(9, '每月平均', BG.monthly); kv(10, '緩衝比例', `${BG.buffer}%`, false); kv(11, '預算來源', BG.locked ? '已鎖定' : '自動計算', false);
    b.getColumn(1).width = 30; b.getColumn(2).width = 16;
    header(b, 13, [['帳單', 30], ['資產', 18], ['週期', 8], ['預估方法', 34], ['本年度實際', 14], ['預估剩餘次數', 12], ['預估剩餘', 14], ['全年', 14]]);
    BG.items.forEach((it, k) => { const r = 14 + k;
      [it.name, it.asset, it.cycle, it.method, it.actual, it.n, it.remaining, null].forEach((v, c) => { const x = b.getCell(r, c + 1); x.value = v; cellStyle(x, [4, 6].includes(c) ? MONEY : null, { center: [2, 5].includes(c) }); });
      const t = b.getCell(r, 8); t.value = { formula: `E${r}+G${r}`, result: r2(it.actual + it.remaining) }; cellStyle(t, MONEY, { bold: true }); });
    const tr = 14 + BG.items.length;
    b.getCell(tr, 1).value = '合計'; b.getCell(tr, 1).font = { name: FONT, bold: true };
    [5, 7, 8].forEach(c => { const x = b.getCell(tr, c); x.value = { formula: `SUM(${L(c - 1)}14:${L(c - 1)}${tr - 1})`, result: r2(BG.items.reduce((t, it) => t + (c === 5 ? it.actual : c === 7 ? it.remaining : it.actual + it.remaining), 0)) }; cellStyle(x, MONEY, { bold: true, soft: true }); });
    let sr = tr + 2;
    b.getCell(sr, 1).value = '存錢計畫（未來 12 個月）'; b.getCell(sr, 1).font = { name: FONT, size: 12, bold: true };
    kv(sr + 1, '帳單帳戶', BG.savings.jar || '未指定', false); kv(sr + 2, '存錢週期', BG.savings.every, false);
    kv(sr + 3, '建議每期存入', BG.savings.per); kv(sr + 4, '平均每期', BG.savings.avg); kv(sr + 5, '未來 12 個月帳單（含緩衝）', BG.savings.need);
    kv(sr + 6, '帳戶目前餘額', BG.savings.bal); kv(sr + 7, '還需要存', BG.savings.toSave);
    sr += 9;
    header(b, sr, [['月份', 30], ['帳單', 18], ['存入', 8], ['月底餘額', 34]]);
    BG.savings.months.forEach((x, k) => { const r = sr + 1 + k; [x.m, x.out, x.inn, x.bal].forEach((v, c) => { const cell = b.getCell(r, c + 1); cell.value = v; cellStyle(cell, c ? MONEY : null); }); });
  }

  /* 總覽 (first sheet): income, spending and net per year, and jar balances */
  titleCell(ov.getCell('A1'), `${X.ledgerName || 'Biledger'}・收支總覽`);
  note(ov.getCell('A2'), `匯出日期 ${new Date().toISOString().slice(0, 10).replace(/-/g, '/')}・貨幣 ${X.currencyCode || ''}・${fs === 1 ? '年度從 1 月開始' : `財年從 ${fs} 月 1 日開始`}・只算已付`);
  header(ov, 4, [['年度', 18], ['收入合計', 16], ['支出合計', 16], ['淨額', 16]]);
  years.forEach((yy, k) => { const r = 5 + k, rng = `${DT},">="&DATE(${yy},${fs},1),${DT},"<"&DATE(${yy + 1},${fs},1),${ST},"已付"`, inY = paid.filter(e => fyOf(e.date) === yy);
    const c0 = ov.getCell(r, 1); c0.value = fyName(yy); cellStyle(c0);
    const a = ov.getCell(r, 2); a.value = { formula: `SUMIFS(${INC},${rng})`, result: paidIn(inY) }; cellStyle(a, MONEY);
    const bq = ov.getCell(r, 3); bq.value = { formula: `SUMIFS(${OUT},${rng})`, result: paidOut(inY) }; cellStyle(bq, MONEY);
    const c = ov.getCell(r, 4); c.value = { formula: `B${r}-C${r}`, result: r2(paidIn(inY) - paidOut(inY)) }; cellStyle(c, MONEY, { bold: true }); });
  const tr = 5 + years.length;
  ov.getCell(tr, 1).value = '全部'; ov.getCell(tr, 1).font = { name: FONT, bold: true };
  [[2, paidIn(paid)], [3, paidOut(paid)], [4, paidIn(paid) - paidOut(paid)]].forEach(([c, v]) => { const x = ov.getCell(tr, c); x.value = { formula: c === 4 ? `B${tr}-C${tr}` : `SUM(${L(c - 1)}5:${L(c - 1)}${tr - 1})`, result: r2(v) }; cellStyle(x, MONEY, { bold: true, soft: true }); });
  if (jarRows.length) {
    header(ov, tr + 2, [['存錢罐', 18], ['目前餘額', 16]]);
    jarRows.forEach((j, k) => { const r = tr + 3 + k, a = ov.getCell(r, 1), c = ov.getCell(r, 2); a.value = j.name; cellStyle(a); c.value = { formula: j.ref, result: j.bal }; cellStyle(c, MONEY); });
  }
  if (BG) { const r = tr + 4 + jarRows.length; ov.getCell(r, 1).value = `${BG.yearName}預算`; ov.getCell(r, 1).font = { name: FONT, bold: true };
    const c = ov.getCell(r, 2); c.value = BG.budget; cellStyle(c, MONEY, { bold: true }); note(ov.getCell(r, 3), '詳細請看「預算」分頁'); }

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
  const cfg = { v: S.v || 3 }; ['folders', 'assets', 'cats', 'methods', 'fields', 'jars', 'bills', 'buffer', 'currency', 'fyStart', 'budget'].forEach(k => cfg[k] = S[k]);
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
  const ws = wb.getWorksheet('收支明細') || wb.getWorksheet('開銷明細');
  if (!ws) throw new Error('找不到「收支明細」分頁');
  // newer files: F 收入 / G 支出 and everything after shifts one column right
  const split = txt(ws.getCell(2, 6)) === '收入' && txt(ws.getCell(2, 7)) === '支出', o = split ? 1 : 0;
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
    const date = isoOf(val(ws.getCell(r, 1))), inc = split ? +val(ws.getCell(r, 6)) || 0 : 0;
    const amount = split ? (inc || +val(ws.getCell(r, 7)) || 0) : Math.abs(+val(ws.getCell(r, 6))), income = split ? inc > 0 : txt(ws.getCell(r, 15)) === '收入';
    if (!date || !amount) continue;
    const sub = subFor(txt(ws.getCell(r, 3)), txt(ws.getCell(r, 4)), income);
    const name = txt(ws.getCell(r, 5)), method = txt(ws.getCell(r, 11 + o));
    if (method && !cfg.methods.includes(method)) cfg.methods.push(method);
    const due = txt(ws.getCell(r, 13 + o)) === '待繳';
    expenses.push({ _in: income, id: X.uid('e'), date, asset: txt(ws.getCell(r, 2)) ? assetFor(txt(ws.getCell(r, 2))).id : '', sub: sub.id, name: name === sub.name ? '' : name,
      amount: r2(amount), ps: isoOf(val(ws.getCell(r, 7 + o))) || '', pe: isoOf(val(ws.getCell(r, 8 + o))) || '', src: jarFor(txt(ws.getCell(r, 10 + o))), method,
      cycle: cycOf(txt(ws.getCell(r, 12 + o))), status: due ? 'due' : 'paid', due: due ? date : '', note: txt(ws.getCell(r, 14 + o)), extra: {} });
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

function pick(S) { const c = {}; ['folders', 'assets', 'cats', 'methods', 'fields', 'jars', 'bills', 'buffer', 'currency', 'fyStart', 'budget'].forEach(k => c[k] = S[k]); return c; }
