// What a reminder notification should say for one ledger on one day.
// Kept free of Firebase so it can be tested on its own. The amounts are the simple version of the
// app's forecast: a manual amount, else the average of the last three payments, else the initial budget.

const CYC_MONTHS = { monthly: 1, quarterly: 3, half: 6, yearly: 12 };
const SYMBOL = { AUD: '$', TWD: 'NT$', USD: 'US$', NZD: 'NZ$', CAD: 'C$', HKD: 'HK$', SGD: 'S$', JPY: '¥', CNY: '¥', KRW: '₩', EUR: '€', GBP: '£' };
const NO_JAR = 'none';

const parse = s => { const [y, m, d] = s.split('-').map(Number); return Date.UTC(y, m - 1, d); };
const iso = t => new Date(t).toISOString().slice(0, 10);
export const daysBetween = (a, b) => Math.round((parse(b) - parse(a)) / 864e5);
export function addMonths(s, n) {
  const [y, m, d] = s.split('-').map(Number);
  const first = new Date(Date.UTC(y, m - 1 + n, 1)), last = new Date(Date.UTC(first.getUTCFullYear(), first.getUTCMonth() + 1, 0)).getUTCDate();
  return iso(Date.UTC(first.getUTCFullYear(), first.getUTCMonth(), Math.min(d, last)));
}
const md = s => `${+s.slice(5, 7)}/${+s.slice(8)}`;

// local date ("YYYY-MM-DD") and hour in a time zone
export function localNow(now, tz) {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-CA', { timeZone: tz || 'UTC', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', hourCycle: 'h23' })
    .formatToParts(now).map(x => [x.type, x.value]));
  return { date: `${p.year}-${p.month}-${p.day}`, hour: +p.hour };
}

/**
 * @param ledger { name, cfg, expenses: [], moves: [] }
 * @param today  "YYYY-MM-DD" in the person's time zone
 * @param types  { bills: bool, jars: bool }
 * @returns { bills: [{date, days, name, amount, estimate}], jars: [{name, balance, alert}] }
 */
export function remindersFor(ledger, today, types = { bills: true, jars: true }) {
  const cfg = ledger.cfg || {}, cats = cfg.cats || [], assets = cfg.assets || [];
  const subInfo = {}; cats.forEach(c => (c.subs || []).forEach(s => { subInfo[s.id] = { name: s.name, income: c.kind === 'in' }; }));
  const assetName = id => { const a = assets.find(x => x.id === id); return a ? (a.nick && ['house', 'apartment', 'land', 'shop'].includes(a.kind) ? a.nick : a.name) : ''; };
  const label = (sub, asset, name) => [name || subInfo[sub]?.name || '帳單', assetName(asset)].filter(Boolean).join('・');
  const paid = (ledger.expenses || []).filter(e => e.status === 'paid');
  const out = { bills: [], jars: [] };

  if (types.bills !== false) {
    const due = (ledger.expenses || []).filter(e => e.status === 'due');
    for (const e of due) {
      const d = e.due || e.date, days = daysBetween(today, d);
      const lead = (cfg.bills || []).find(b => b.asset === e.asset && b.sub === e.sub)?.lead ?? 3;
      if (days <= lead && days >= -7) out.bills.push({ date: d, days, name: label(e.sub, e.asset, e.name), amount: +e.amount || 0, estimate: false });
    }
    for (const b of cfg.bills || []) {
      const step = CYC_MONTHS[b.cycle];
      if (!b.on || !step || subInfo[b.sub]?.income) continue;
      if (due.some(e => e.asset === b.asset && e.sub === b.sub)) continue;
      const hist = paid.filter(e => e.asset === b.asset && e.sub === b.sub).sort((x, y) => (x.date < y.date ? -1 : 1));
      const next = hist.length ? addMonths(hist.at(-1).date, step) : b.first || null;
      if (!next) continue;
      const days = daysBetween(today, next);
      if (days > (b.lead ?? 5) || days < -7) continue;
      const recent = hist.slice(-3), amount = +b.manual > 0 ? +b.manual : recent.length ? recent.reduce((t, e) => t + +e.amount, 0) / recent.length : +b.init || 0;
      out.bills.push({ date: next, days, name: label(b.sub, b.asset), amount, estimate: !(+b.manual > 0) });
    }
    out.bills.sort((a, b) => (a.date < b.date ? -1 : 1));
  }
  if (types.jars !== false) {
    for (const j of cfg.jars || []) {
      if (!(+j.alert > 0)) continue;
      const moves = (ledger.moves || []).filter(m => m.jar === j.id).reduce((t, m) => t + (m.kind === 'in' ? +m.amount : -m.amount), 0);
      const recs = paid.filter(e => e.src === j.id && e.src !== NO_JAR).reduce((t, e) => t + (subInfo[e.sub]?.income ? +e.amount : -e.amount), 0);
      const balance = Math.round((moves + recs) * 100) / 100;
      if (balance < +j.alert) out.jars.push({ name: j.name, balance, alert: +j.alert });
    }
  }
  return out;
}

// The notification for one ledger, or null when there is nothing to say.
export function messageFor(ledgerName, r, currency) {
  const sym = SYMBOL[currency] || '$', dp = ['TWD', 'JPY', 'KRW'].includes(currency) ? 0 : 2;
  const fmt = n => `${n < 0 ? '−' : ''}${sym}${Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp })}`;
  const when = d => (d === 0 ? '今天' : d === 1 ? '明天' : d > 0 ? `${d} 天後` : `已過 ${-d} 天`);
  const lines = [
    ...r.bills.map(b => `${b.name} ${md(b.date)}（${when(b.days)}）${b.estimate ? '約 ' : ''}${fmt(b.amount)}`),
    ...r.jars.map(j => `存錢罐「${j.name}」餘額 ${fmt(j.balance)}，低於 ${fmt(j.alert)}`),
  ];
  if (!lines.length) return null;
  const title = r.bills.length
    ? `${ledgerName}：${r.bills.length} 筆帳單${r.bills.some(b => b.days < 0) ? '要繳（含逾期）' : '快到期'}`
    : `${ledgerName}：存錢罐餘額不足`;
  const more = lines.length > 4 ? `\n…還有 ${lines.length - 4} 項，打開 App 看全部` : '';
  return { title, body: lines.slice(0, 4).join('\n') + more };
}
