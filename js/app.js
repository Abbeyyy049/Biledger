import * as B from './backend.js';
import { exportWorkbook, parseWorkbook } from './excel.js';

const ICON = {
  house:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 10.5 12 3l9 7.5V20a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/></svg>',
  car:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M5 17h14M5 17v2M19 17v2M3 13l2-5.5A2 2 0 0 1 6.9 6h10.2a2 2 0 0 1 1.9 1.5L21 13v4H3z"/><circle cx="7.5" cy="13.5" r=".6"/><circle cx="16.5" cy="13.5" r=".6"/></svg>',
  apartment:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="3" width="16" height="18" rx="1"/><path d="M8 7h2M14 7h2M8 11h2M14 11h2M8 15h2M14 15h2M10 21v-3h4v3"/></svg>',
  moto:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="5.5" cy="16.5" r="3"/><circle cx="18.5" cy="16.5" r="3"/><path d="M8.5 16.5h6l3-6h-4l-2-3H9M14 7h3l1.5 3.5"/></svg>',
  shop:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l1.5-5h15L21 9M3 9v11h18V9M3 9h18M9 20v-6h6v6"/></svg>',
  land:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21s-7-6.2-7-11.5a7 7 0 0 1 14 0C19 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/></svg>',
  person:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 3.6-7 8-7s8 3 8 7"/></svg>',
  pet:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="7" cy="9" r="1.8"/><circle cx="12" cy="6.5" r="1.8"/><circle cx="17" cy="9" r="1.8"/><path d="M12 12c-3 0-5.5 3.2-5.5 5.3 0 1.6 1.4 2.2 2.8 1.8 1.1-.3 1.8-.6 2.7-.6s1.6.3 2.7.6c1.4.4 2.8-.2 2.8-1.8C17.5 15.2 15 12 12 12z"/></svg>',
  device:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="5" width="16" height="11" rx="1.5"/><path d="M2 19h20"/></svg>',
  invest:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M14.8 9.2c-.6-.8-1.6-1.2-2.8-1.2-1.7 0-2.8.8-2.8 2s1.1 1.7 2.8 2 2.8.8 2.8 2-1.1 2-2.8 2c-1.2 0-2.3-.5-2.9-1.3M12 6.5v11"/></svg>',
  other:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8 12 3 3 8v8l9 5 9-5z"/><path d="m3 8 9 5 9-5M12 13v8"/></svg>',
  gear:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>',
  check:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="m5 12 5 5 9-10"/></svg>',
  x:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12M18 6 6 18"/></svg>',
  trash:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg>',
  pick:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="5"/><path d="m8 12 3 3 5-6"/></svg>',
  in:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 19V5M5 12l7 7 7-7"/></svg>',
  out:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14M5 12l7-7 7 7"/></svg>',
  bell:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M6 8a6 6 0 1 1 12 0c0 7 3 9 3 9H3s3-2 3-9M10.3 21a1.9 1.9 0 0 0 3.4 0"/></svg>',
  jar:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8 3h8M9 3v3.2C6.6 7.4 5 9.6 5 12.5V19a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-6.5c0-2.9-1.6-5.1-4-6.3V3"/></svg>',
  pin:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 17v5M9 3h6l-1 6 4 4H6l4-4z"/></svg>',
};
const CYC = {once:['一次性',0],monthly:['每月',1],quarterly:['每季',3],half:['每半年',6],yearly:['每年',12]};
const NONE='none', NONE_LABEL='已結清';
const RENAME = ['elec','water','internet','phone','bodycorp','council','rego','carins','homeins'];

/* ---------- dates & money ---------- */
const pad=n=>String(n).padStart(2,'0');
const iso=d=>d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());
const parse=s=>{const [y,m,d]=s.split('-').map(Number);return new Date(y,m-1,d)};
const TODAY=iso(new Date());
const addMonths=(s,n)=>{const d=parse(s),day=d.getDate();d.setDate(1);d.setMonth(d.getMonth()+n);d.setDate(Math.min(day,new Date(d.getFullYear(),d.getMonth()+1,0).getDate()));return iso(d)};
const addDays=(s,n)=>{const d=parse(s);d.setDate(d.getDate()+n);return iso(d)};
const diff=(a,b)=>Math.round((parse(b)-parse(a))/864e5);
const slash=s=>s?s.replace(/-/g,'/'):'';
const md=s=>s?(+s.slice(5,7))+'/'+(+s.slice(8)):'';
const WEEK='日一二三四五六';
// currency code -> [name, symbol, decimals]; only changes how amounts are shown, nothing is converted
const CUR={AUD:['澳幣','$',2],TWD:['新台幣','NT$',0],USD:['美元','US$',2],NZD:['紐西蘭幣','NZ$',2],CAD:['加幣','C$',2],HKD:['港幣','HK$',2],SGD:['新加坡幣','S$',2],JPY:['日圓','¥',0],CNY:['人民幣','¥',2],KRW:['韓元','₩',0],EUR:['歐元','€',2],GBP:['英鎊','£',2]};
const cur=()=>CUR[S.currency]||CUR.AUD;
const zeroText=()=>cur()[2]?'0.00':'0';
const money=(n,sign)=>{n=+n||0;const [,sym,dp]=cur();const s=Math.abs(n).toLocaleString('en-US',{minimumFractionDigits:dp,maximumFractionDigits:dp});return (n<0?'−':(sign&&n>0?'+':''))+sym+s};
const money0=n=>{const [,sym]=cur();return (n<0?'−':'')+sym+Math.round(Math.abs(n)).toLocaleString('en-US')};
function guessCurrency(){const tz=(Intl.DateTimeFormat().resolvedOptions().timeZone||'');
  return tz.startsWith('Australia')?'AUD':tz==='Asia/Taipei'?'TWD':tz==='Pacific/Auckland'?'NZD':tz==='Asia/Hong_Kong'?'HKD':tz==='Asia/Tokyo'?'JPY':tz==='Asia/Singapore'?'SGD':tz.startsWith('Europe/London')?'GBP':'USD'}
// Chinese fonts from Google Fonts: key -> [label, family, css2 query]
const FONTS={system:['系統預設','',''],noto:['思源黑體','Noto Sans TC','Noto+Sans+TC:wght@400;500;700'],goround:['昭源圓體','Chiron GoRound TC','Chiron+GoRound+TC:wght@400;500;700'],
  huninn:['粉圓','Huninn','Huninn'],wenkai:['霞鶩文楷','LXGW WenKai TC','LXGW+WenKai+TC:wght@400;700']};
const BASE_UI='-apple-system,BlinkMacSystemFont,"PingFang TC","Microsoft JhengHei","Noto Sans TC","Plus Jakarta Sans",sans-serif';
function loadFonts(keys){const q=keys.map(k=>FONTS[k]?.[2]).filter(Boolean);if(!q.length)return;const href=`https://fonts.googleapis.com/css2?${q.map(x=>'family='+x).join('&')}&display=swap`;
  if(document.querySelector(`link[href="${href}"]`))return;const l=document.createElement('link');l.rel='stylesheet';l.href=href;document.head.appendChild(l)}
// A font file someone uploads stays in this device's browser storage (IndexedDB) only.
const IDB={open(){return this.p??=new Promise((ok,no)=>{const r=indexedDB.open('biledger',1);r.onupgradeneeded=()=>r.result.createObjectStore('files');r.onsuccess=()=>ok(r.result);r.onerror=()=>no(r.error)})},
  async run(mode,fn){const d=await this.open();return new Promise((ok,no)=>{const q=fn(d.transaction('files',mode).objectStore('files'));q.onsuccess=()=>ok(q.result);q.onerror=()=>no(q.error)})},
  get(k){return this.run('readonly',s=>s.get(k))},set(k,v){return this.run('readwrite',s=>s.put(v,k))},del(k){return this.run('readwrite',s=>s.delete(k))}};
let CUSTOM_FONT=null; // {name} once the file on this device is loaded
async function loadCustomFont(){
  try{const rec=await IDB.get('font');if(!rec)return false;
    const ff=new FontFace('BiledgerCustom',await rec.blob.arrayBuffer());await ff.load();document.fonts.add(ff);CUSTOM_FONT={name:rec.name};return true}
  catch(e){return false}
}
function applyFont(k){
  try{localStorage.setItem('font',k)}catch(e){}
  if(k==='custom'){
    const use=()=>document.documentElement.style.setProperty('--ui',`"BiledgerCustom",${BASE_UI}`);
    if(CUSTOM_FONT)use();else loadCustomFont().then(ok=>{if(ok)use();else document.documentElement.style.setProperty('--ui',BASE_UI)});
    return;
  }
  const f=FONTS[k]||FONTS.system;loadFonts([k]);document.documentElement.style.setProperty('--ui',f[1]?`"${f[1]}",${BASE_UI}`:BASE_UI);
}
try{applyFont(localStorage.getItem('font')||'system')}catch(e){}
loadCustomFont().then(ok=>{if(ok&&tab==='setup')scheduleRender()});
async function uploadFont(file){
  if(!/\.(ttf|otf|woff2?)$/i.test(file.name))throw new Error('請選 .ttf、.otf、.woff 或 .woff2 字型檔');
  try{await new FontFace('BiledgerTest',await file.arrayBuffer()).load()}catch(e){throw new Error('讀不到這個字型檔，可能檔案壞了或格式不支援')}
  await IDB.set('font',{name:file.name.replace(/\.\w+$/,''),blob:file});
  navigator.storage?.persist?.().catch(()=>{});
  CUSTOM_FONT=null;await loadCustomFont();
}
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const uid=p=>p+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
const sum=a=>a.reduce((t,e)=>t+(+e.amount||0),0);

/* ---------- state ---------- */
const DEFAULT_CATS=[
  {id:'life',name:'生活',subs:[['food','餐飲'],['grocery','日用品'],['shopping','購物'],['medical','醫療'],['edu','教育'],['fun','娛樂'],['subscr','訂閱服務']]},
  {id:'home',name:'住房',subs:[['rent','房租'],['mortgage','房貸'],['bodycorp','管理費'],['council','房屋稅／地價稅'],['homerepair','居家維修'],['garden','清潔／園藝']]},
  {id:'util',name:'水電網路',subs:[['elec','電費'],['water','水費'],['gas','瓦斯'],['internet','網路費'],['phone','電信費']]},
  {id:'car',name:'交通',subs:[['fuel','加油／充電'],['parking','停車'],['toll','過路費'],['transit','大眾運輸'],['service','汽車保養'],['carrepair','汽車維修'],['rego','牌照／Rego'],['caracc','汽車用品'],['license','駕照']]},
  {id:'ins',name:'保險',subs:[['homeins','房屋險'],['carins','汽車險'],['healthins','健康／醫療險'],['lifeins','人壽險']]},
  {id:'fee',name:'稅費規費',subs:[['stamp','印花稅／契稅'],['transferfee','過戶登記費'],['legal','律師／代書費'],['firb','政府規費'],['tax','其他稅金']]},
  {id:'buy',name:'購置',subs:[['housebuy','購屋'],['carbuy','購車'],['deposit','訂金'],['furniture','家具家電']]},
  {id:'misc',name:'其他',subs:[['other','其他']]},
];
const CFG_KEYS=['folders','assets','cats','methods','fields','jars','bills','buffer','currency','font'];
function defaultConfig(){
  return {v:3,folders:[{id:'f_me',name:'個人'}],assets:[],
    cats:DEFAULT_CATS.map(c=>({...c,subs:c.subs.map(([id,name])=>({id,name}))})),
    methods:['現金 Cash','信用卡 Credit Card','簽帳金融卡 Debit Card','銀行轉帳 Bank Transfer','自動扣款 Direct Debit','行動支付 Mobile Pay'],
    fields:[],jars:[],bills:[],buffer:0,currency:guessCurrency(),font:'system'};
}
// Accounts made before v3 used house/car-specific categories, plain-Chinese payment
// methods and a pseudo-asset for "no asset". Move them onto the general setup.
function migrateV2(st){
  const kindOf=id=>st.assets.find(a=>a.id===id)?.kind;
  const cats=defaultConfig().cats, known=new Set(cats.flatMap(c=>c.subs.map(x=>x.id)));
  const CATMAP={buy:'buy',fee:'fee',util:'util',ins:'ins',fix:'home',misc:'misc'};
  st.cats.forEach(c=>{const target=cats.find(x=>x.id===CATMAP[c.id]);
    c.subs.forEach(x=>{if(known.has(x.id)||x.id==='repair'||x.id==='purchase')return;
      if(target)target.subs.push(x);else{let t=cats.find(y=>y.id===c.id);if(!t){t={id:c.id,name:c.name,subs:[]};cats.push(t)}t.subs.push(x)}})});
  const subOf=(sub,asset)=>sub==='repair'?(kindOf(asset)==='car'?'carrepair':'homerepair'):sub==='purchase'?(kindOf(asset)==='car'?'carbuy':'housebuy'):sub;
  const METH={'銀行轉帳':'銀行轉帳 Bank Transfer','轉帳':'銀行轉帳 Bank Transfer','簽帳金融卡':'簽帳金融卡 Debit Card','信用卡':'信用卡 Credit Card','現金':'現金 Cash','自動扣款':'自動扣款 Direct Debit'};
  const mm=m=>METH[m]||m;
  const dropNone=st.assets.some(a=>a.id==='a_none'), fixAsset=a=>dropNone&&a==='a_none'?'':a;
  const cfg={v:3};CFG_KEYS.forEach(k=>cfg[k]=clean(st[k]??defaultConfig()[k]));
  Object.assign(cfg,{cats,methods:[...new Set(st.methods.map(mm))],assets:st.assets.filter(a=>!(dropNone&&a.id==='a_none')),
    bills:st.bills.map(b=>({...b,asset:fixAsset(b.asset),sub:subOf(b.sub,b.asset)})),currency:'AUD',font:st.font||'system'});
  const expenses=st.expenses.map(e=>({...e,sub:subOf(e.sub,e.asset),asset:fixAsset(e.asset),method:mm(e.method)}));
  return {cfg,expenses};
}
let S={...defaultConfig(),expenses:[],moves:[],icons:[]};
const clean=o=>JSON.parse(JSON.stringify(o));
const fail=e=>{console.error(e);toast('沒有存到雲端：'+(e.code||e.message))};
let cfgTimer, cfgDirty=false;
// settings changes are batched for 300 ms; until written, incoming copies must not overwrite them
function saveCfg(){cfgDirty=true;clearTimeout(cfgTimer);cfgTimer=setTimeout(()=>{cfgDirty=false;const c={v:3};CFG_KEYS.forEach(k=>c[k]=S[k]);B.saveConfig(clean(c)).catch(fail)},300)}

const subMap=()=>{const m={};S.cats.forEach(c=>c.subs.forEach(s=>m[s.id]={...s,cat:c}));return m};
const subName=id=>subMap()[id]?.name||'未分類';
const catOf=id=>subMap()[id]?.cat;
const assetOf=id=>S.assets.find(a=>a.id===id)||{name:'未指定資產',nick:'',kind:'other',folder:''};
const withAsset=(text,id)=>id?text+'・'+assetShort(id):text;
const matchAsset=(e,sel)=>sel==='all'||(sel==='none'?!e.asset:sel.startsWith('f:')?!!e.asset&&assetOf(e.asset).folder===sel.slice(2):e.asset===sel);
const buyName=()=>S.cats.find(c=>c.id==='buy')?.name||'購置';
// asset icons people can pick from; for places, the nickname is shown in lists instead of the long address
const KINDS=[['house','房子'],['apartment','公寓／大樓'],['land','土地'],['shop','店面／公司'],['car','汽車'],['moto','機車'],['person','家人'],['pet','寵物'],['device','電子設備'],['invest','投資'],['other','其他']];
const PLACE=['house','apartment','land','shop'];
// built-in icon key, or 'u:<id>' for an icon the person uploaded
function iconHTML(k){if(k?.startsWith('u:')){const ic=(S.icons||[]).find(i=>i.id===k.slice(2));if(ic)return `<img src="${ic.data}" alt="" style="width:70%;height:70%;object-fit:contain">`}return ICON[k]||ICON.other}
const assetShort=id=>{const a=assetOf(id);return a.nick&&PLACE.includes(a.kind)?a.nick:a.name};
const jarOf=id=>S.jars.find(j=>j.id===id);
const srcLabel=id=>id===NONE?NONE_LABEL:(jarOf(id)?.name||'已刪除的罐子');
const shortPeriod=e=>e.ps&&e.pe?md(e.ps)+'–'+md(e.pe):'';
// title shown in lists: own name if given, otherwise just the category
const title=e=>e.name||subName(e.sub);
// longer label for text exports (request slip)
const fullName=e=>e.name||(subName(e.sub)+(e.ps&&e.pe?' '+shortPeriod(e):' '+md(e.date)));
function timing(e){if(!e.ps||!e.pe||!e.date)return '';if(e.date<=e.ps)return '預付';if(e.date>=e.pe)return '後付';return '期間中';}
const paid=()=>S.expenses.filter(e=>e.status==='paid');
const notBuy=e=>catOf(e.sub)?.id!=='buy';

function jarLedger(jid){
  const rows=[...S.moves.filter(m=>m.jar===jid).map(m=>({date:m.date,amount:m.kind==='in'?+m.amount:-m.amount,label:m.note||(m.kind==='in'?'存入':'調整'),mv:m})),
    ...paid().filter(e=>e.src===jid).map(e=>({date:e.date,amount:-e.amount,label:fullName(e),exp:e}))];
  rows.sort((a,b)=>a.date<b.date?-1:a.date>b.date?1:(b.amount>0)-(a.amount>0));
  let bal=0; rows.forEach(r=>{bal+=r.amount;r.bal=bal}); return rows;
}
const jarBalance=jid=>{const l=jarLedger(jid);return l.length?l.at(-1).bal:0};

function upcoming(){
  const out=[];
  S.expenses.filter(e=>e.status==='due').forEach(e=>out.push({kind:'due',date:e.due||e.date,amount:+e.amount||0,asset:e.asset,sub:e.sub,exp:e,src:e.src}));
  S.bills.filter(b=>b.on&&CYC[b.cycle][1]).forEach(b=>{
    if(out.some(o=>o.asset===b.asset&&o.sub===b.sub)) return;
    const hist=paid().filter(e=>e.asset===b.asset&&e.sub===b.sub).sort((a,b)=>a.date<b.date?-1:1);
    const last=hist.at(-1); if(!last) return;
    const same=hist.filter(e=>e.cycle===b.cycle), recent=(same.length?same:hist).slice(-3);
    // yearly bills: last price is the best guess; shorter cycles: average of the last 3
    const avg=b.cycle==='yearly'?last.amount:sum(recent)/recent.length;
    out.push({kind:'bill',date:addMonths(last.date,CYC[b.cycle][1]),amount:last.amount,avg,asset:b.asset,sub:b.sub,bill:b,last,src:last.src});
  });
  return out.sort((a,b)=>a.date<b.date?-1:1);
}
// expected spend from one jar over the next N months, from recurring bills + bills marked 待繳
function forecast(jid,months){
  const end=addMonths(TODAY,months), items=[];
  upcoming().filter(u=>u.src===jid).forEach(u=>{
    if(u.kind==='due'){items.push({key:'d'+u.exp.id,label:withAsset(title(u.exp),u.asset),n:1,each:u.amount,total:u.amount,note:'已記待繳'});return}
    const step=CYC[u.bill.cycle][1];let d=u.date,n=0;while(d<=end){n++;d=addMonths(d,step)}
    if(n) items.push({key:u.bill.id,label:withAsset(subName(u.sub),u.asset),n,each:u.avg,total:u.avg*n,note:CYC[u.bill.cycle][0]});
  });
  return items.sort((a,b)=>b.total-a.total);
}

/* ---------- ui helpers ---------- */
const main=document.getElementById('main');
let tab='home';
const F={asset:'all',cat:'all',period:'year',src:'all',q:'',pick:false,sel:new Set(),from:'',to:''};
const PERIODS=[['all','全部時間'],['month','本月'],['lastmonth','上個月'],['year','今年'],['lastyear','去年'],['custom','自訂期間…']];
// [first day, last day] for the list's time filter; null = no limit
function periodRange(){
  const y=+TODAY.slice(0,4), som=TODAY.slice(0,7)+'-01', eom=s=>addDays(addMonths(s.slice(0,7)+'-01',1),-1);
  switch(F.period){
    case 'month': return [som,eom(som)];
    case 'lastmonth': {const a=addMonths(som,-1);return [a,eom(a)]}
    case 'year': return [`${y}-01-01`,`${y}-12-31`];
    case 'lastyear': return [`${y-1}-01-01`,`${y-1}-12-31`];
    case 'custom': return [F.from||'0000-01-01',F.to||'9999-12-31'];
    default: return null;
  }
}
const C={asset:'all',subs:['elec','water'],slot:{elec:0,water:1},mode:'service',range:12};
let setupTab='assets', curJar=null;
function toast(msg){const t=document.createElement('div');t.className='toast';t.textContent=msg;document.body.appendChild(t);setTimeout(()=>t.remove(),1800)}
function topbar(t,sub,right=''){return `<header class="top"><div class="grow"><h1>${t}</h1>${sub?`<div class="sub">${sub}</div>`:''}</div>${right}</header>`}
const gearBtn=`<button class="iconbtn" data-tab="setup" aria-label="設定">${ICON.gear}</button>`;
function assetOptions(sel,withAll){
  let h=withAll?`<option value="all"${sel==='all'?' selected':''}>全部資產</option>`:`<option value=""${!sel?' selected':''}>不指定資產</option>`;
  const groups=S.folders.map(f=>({f,as:S.assets.filter(a=>a.folder===f.id&&(withAll||!a.archived||a.id===sel))})).filter(g=>g.as.length);
  const label=a=>esc(a.name)+(a.nick&&PLACE.includes(a.kind)?'（'+esc(a.nick)+'）':'');
  groups.forEach(({f,as})=>{
    if(!withAll){h+=`<optgroup label="${esc(f.name)}">`+as.map(a=>`<option value="${a.id}"${a.id===sel?' selected':''}>${label(a)}</option>`).join('')+`</optgroup>`;return}
    // filters: the group is selectable itself and its assets are indented under it (one group: just the assets)
    const sub=groups.length>1;
    if(sub) h+=`<option value="f:${f.id}"${sel==='f:'+f.id?' selected':''}>${esc(f.name)}</option>`;
    h+=as.map(a=>`<option value="${a.id}"${a.id===sel?' selected':''}>${sub?'　└ ':''}${label(a)}</option>`).join('');});
  if(withAll) h+=`<option value="none"${sel==='none'?' selected':''}>未指定資產</option>`;
  return h;
}
function subOptions(sel,withAll){
  let h=withAll?`<option value="all"${sel==='all'?' selected':''}>全部類別</option>`:'';
  S.cats.forEach(c=>{
    if(withAll){h+=`<option value="c:${c.id}"${sel==='c:'+c.id?' selected':''}>${esc(c.name)}</option>`+c.subs.map(s=>`<option value="${s.id}"${s.id===sel?' selected':''}>　└ ${esc(s.name)}</option>`).join('');return}
    if(c.subs.length) h+=`<optgroup label="${esc(c.name)}">`+c.subs.map(s=>`<option value="${s.id}"${s.id===sel?' selected':''}>${esc(s.name)}</option>`).join('')+`</optgroup>`;});
  return h;
}
const srcTag=e=>e.src===NONE?`<span class="src none">${NONE_LABEL}</span>`:`<span class="src">${ICON.jar}${esc(srcLabel(e.src))}</span>`;
function expRow(e,pick){
  const a=assetOf(e.asset),on=F.sel.has(e.id),per=shortPeriod(e);
  return `<button class="item" data-act="${pick?'toggle':'edit'}" data-id="${e.id}">
    ${pick?`<span class="chk${on?' on':''}">${on?ICON.check:''}</span>`:`<span class="dot">${iconHTML(a.kind)}</span>`}
    <span class="grow"><span class="ell" style="display:block;font-weight:600">${esc(title(e))}${e.asset?`<span class="muted" style="font-weight:500">・${esc(assetShort(e.asset))}</span>`:''}</span>
      <span class="meta"><span class="tiny muted">${md(e.date)}${per?`・用 ${per}`:''}</span>
        ${e.status==='due'?'<span class="pill due">待繳</span>':''}${srcTag(e)}${e.flag?'<span class="pill neg">待確認</span>':''}</span></span>
    <span class="amt num">${money(e.amount)}</span></button>`;
}
function jarCard(j,mini){
  const bal=jarBalance(j.id);
  const note=bal<0?`你先墊了 <b class="num">${money(-bal)}</b>`:bal<j.alert?`快見底了，低於 <span class="num">${money0(j.alert)}</span>`:`還可以扣 <span class="num">${money(bal)}</span>`;
  return `<section class="card jar ${mini?'mini':''} ${bal<0?'neg':''}">
    <div class="row between"><span class="label">${ICON.jar.replace('<svg','<svg style="width:13px;height:13px;vertical-align:-2px;margin-right:4px"')}${esc(j.name)}・存錢罐</span><button class="btn sm ghost" data-act="openjar" data-id="${j.id}">明細 ›</button></div>
    <div class="bal num">${money(bal)}</div><div class="note">${note}</div>
    <div class="acts"><button class="btn primary sm" data-act="request" data-id="${j.id}">產生請款單</button><button class="btn sm" data-act="deposit" data-id="${j.id}">存入</button></div>
  </section>`;
}

/* ---------- home ---------- */
function renderHome(){
  const d=new Date(), y=TODAY.slice(0,4), ym=TODAY.slice(0,7);
  const ops=paid().filter(notBuy);
  const mT=sum(ops.filter(e=>e.date.startsWith(ym))), yT=sum(ops.filter(e=>e.date.startsWith(y)));
  const up=upcoming().filter(u=>diff(TODAY,u.date)<=45).slice(0,6);
  const flags=S.expenses.filter(e=>e.flag).length;
  const pinned=S.jars.filter(j=>j.pin);
  const upHtml=up.length?up.map(u=>{const n=diff(TODAY,u.date);
      const state=n<0?`<span class="pill neg">已過 ${-n} 天，還沒記</span>`:u.kind==='due'?`<span class="pill due">待繳</span>`:(u.bill.mode==='post'?'<span class="pill">後付・帳單快出來了</span>':'<span class="pill">預付</span>');
      return `<button class="item" data-act="${u.kind==='due'?'edit':'addfrom'}" data-id="${u.kind==='due'?u.exp.id:u.bill.id}">
        <span class="when"><b class="num">${md(u.date)}</b><span>${n===0?'今天':n>0?n+' 天後':'星期'+WEEK[parse(u.date).getDay()]}</span></span>
        <span class="grow"><span style="display:block;font-weight:600" class="ell">${esc(subName(u.sub))}${u.asset?`<span class="muted" style="font-weight:500">・${esc(assetShort(u.asset))}</span>`:''}</span><span class="meta">${state}</span></span>
        <span class="amt num muted">${u.kind==='due'?'':'約 '}${money(u.amount)}</span></button>`}).join(''):'<div class="empty">接下來 45 天沒有要繳的帳單</div>';
  const groups=S.folders.map(f=>({f,as:S.assets.filter(a=>a.folder===f.id&&!a.archived)})).filter(g=>g.as.length);
  const card=a=>{const es=paid().filter(e=>e.asset===a.id);const last=es.sort((x,z)=>x.date<z.date?1:-1)[0];
    return `<button class="asset" data-act="assetlist" data-id="${a.id}"><span class="row"><span class="dot" style="width:30px;height:30px">${iconHTML(a.kind)}</span><span class="tiny muted ell grow">${esc(a.nick||'')}</span></span>
      <span class="nm">${esc(a.name)}</span><span class="v num">${money0(sum(es.filter(e=>notBuy(e)&&e.date.startsWith(y))))}</span><span class="tiny faint">今年・${last?'最近 '+md(last.date)+' '+esc(subName(last.sub)):'還沒有紀錄'}</span></button>`};
  const folders=groups.length?groups.map(g=>`<div>${groups.length>1?`<div class="gh" style="padding-top:0"><span>${esc(g.f.name)}</span><span class="num">${money0(sum(ops.filter(e=>e.date.startsWith(y)&&g.as.some(a=>a.id===e.asset))))}／今年</span></div>`:''}<div class="agrid">${g.as.map(card).join('')}</div></div>`).join('')
    :`<div class="card empty">還沒有資產。<br>房子、車子或任何想分開記帳的東西都可以設成資產，這是選填的。<br><button class="btn sm ghost" style="margin-top:8px" data-act="goassets">新增資產</button></div>`;
  main.innerHTML=topbar(`${d.getMonth()+1}月${d.getDate()}日 星期${WEEK[d.getDay()]}`,syncLabel(),gearBtn)+`
  <div class="homegrid"><div class="stack">
    ${pinned.length?`<div class="jarcards">${pinned.map((j,i)=>jarCard(j,i>0)).join('')}</div>`:`<button class="card pad row" data-tab="jar" style="text-align:left"><span class="dot">${ICON.pin}</span><span class="grow small muted">首頁沒有釘選存錢罐。到存錢罐頁面按「釘選到首頁」</span><span class="muted">›</span></button>`}
    <div class="tiles">
      <div class="card tile"><div class="label">${d.getMonth()+1} 月花費</div><div class="v num">${money(mT)}</div><div class="tiny faint">不含${esc(buyName())}類</div></div>
      <div class="card tile"><div class="label">${y} 年花費</div><div class="v num">${money(yT)}</div><div class="tiny faint">不含${esc(buyName())}類</div></div>
    </div>
    ${flags?`<button class="card pad row" data-act="goreview" style="text-align:left;border-color:var(--warn)"><span class="dot" style="background:var(--warn-soft);color:var(--warn)">${ICON.bell}</span><span class="grow"><b>${flags} 筆匯入資料要請你確認</b><span class="tiny muted" style="display:block">匯入時發現可能有問題的資料，點進去逐筆確認</span></span><span class="muted">›</span></button>`:''}
    <section class="card"><div class="row between pad" style="padding-bottom:6px"><h2 class="h2">接下來要繳</h2><button class="btn sm ghost" data-act="gobills">提醒設定</button></div><div class="list">${upHtml}</div></section>
  </div>
  <div class="stack"><div class="row between" style="padding:4px 4px 0"><h2 class="h2">我的資產</h2><button class="btn sm ghost" data-act="goassets">管理</button></div><div class="folders">${folders}</div></div></div>`;
}

/* ---------- list ---------- */
function filtered(){
  const y=TODAY.slice(0,4), ym=TODAY.slice(0,7), q=F.q.trim().toLowerCase();
  const range=periodRange();
  return S.expenses.filter(e=>{
    if(!matchAsset(e,F.asset)) return false;
    if(F.cat!=='all'){ if(F.cat.startsWith('c:')){ if(catOf(e.sub)?.id!==F.cat.slice(2)) return false;} else if(e.sub!==F.cat) return false;}
    if(F.src!=='all'&&e.src!==F.src) return false;
    if(range&&(e.date<range[0]||e.date>range[1])) return false;
    if(q&&!(title(e)+' '+e.note+' '+(e.orig||'')+' '+subName(e.sub)+' '+assetOf(e.asset).name+' '+assetOf(e.asset).nick).toLowerCase().includes(q)) return false;
    return true;}).sort((a,b)=>a.date<b.date?1:a.date>b.date?-1:0);
}
function renderList(keepFocus){
  const rows=filtered();
  const groups={}; rows.forEach(e=>{(groups[e.date.slice(0,7)]??=[]).push(e)});
  const selRows=S.expenses.filter(e=>F.sel.has(e.id));
  main.innerHTML=topbar('明細',`${rows.length} 筆`,`<button class="iconbtn" data-act="pickmode" aria-pressed="${F.pick}" aria-label="勾選計算" style="${F.pick?'background:var(--accent);color:var(--accent-ink);border-color:var(--accent)':''}">${ICON.pick}</button>${gearBtn}`)+`
  <div class="stack">
    <input id="q" class="search" type="search" placeholder="搜尋名稱、備註、資產…" value="${esc(F.q)}">
    <div class="filters">
      <select id="fa" class="sel${F.asset!=='all'?' on':''}">${assetOptions(F.asset,true)}</select>
      <select id="fc" class="sel${F.cat!=='all'?' on':''}">${subOptions(F.cat,true)}</select>
      <select id="fp" class="sel${F.period!=='all'?' on':''}">${PERIODS.map(([v,l])=>`<option value="${v}"${F.period===v?' selected':''}>${l}</option>`).join('')}</select>
      <select id="fy" class="sel${F.src!=='all'?' on':''}"><option value="all">扣款來源：全部</option>${S.jars.map(j=>`<option value="${j.id}"${F.src===j.id?' selected':''}>存錢罐：${esc(j.name)}</option>`).join('')}<option value="${NONE}"${F.src===NONE?' selected':''}>${NONE_LABEL}</option></select>
    </div>
    ${F.period==='custom'?`<div class="inline"><span class="small muted">從</span><input id="ffrom" class="search" type="date" value="${F.from}" style="width:auto"><span class="small muted">到</span><input id="fto" class="search" type="date" value="${F.to}" style="width:auto"></div>`:''}
    <div class="card sumbar"><div><div class="label">篩選結果合計</div><div class="tiny faint">${rows.length} 筆${F.pick?'・點選項目來勾選':''}</div></div><div class="v num">${money(sum(rows))}</div></div>
    ${rows.length?Object.entries(groups).map(([m,es])=>`<div><div class="gh"><span>${m.slice(0,4)} 年 ${+m.slice(5)} 月</span><span class="num">${money(sum(es))}</span></div><div class="card list">${es.map(e=>expRow(e,F.pick)).join('')}</div></div>`).join(''):'<div class="card empty">沒有符合條件的開銷</div>'}
  </div>
  ${F.pick?`<div class="picktray"><div class="card sumbar"><div><div class="tiny muted">已勾選 ${selRows.length} 筆</div><div class="v num">${money(sum(selRows))}</div></div><div class="inline"><button class="btn sm" data-act="pickall">全選這些</button><button class="btn sm" data-act="pickclear">清除</button></div></div></div>`:''}`;
  const q=document.getElementById('q');
  q.oninput=()=>{F.q=q.value;renderList(true)};
  if(keepFocus){q.focus();q.setSelectionRange(q.value.length,q.value.length)}
  [['fa','asset'],['fc','cat'],['fy','src']].forEach(([id,k])=>document.getElementById(id).onchange=ev=>{F[k]=ev.target.value;renderList()});
  document.getElementById('fp').onchange=ev=>{F.period=ev.target.value;
    if(F.period==='custom'&&!F.from){F.from=TODAY.slice(0,7)+'-01';F.to=TODAY}
    renderList()};
  [['ffrom','from'],['fto','to']].forEach(([id,k])=>{const el=document.getElementById(id);if(el)el.onchange=()=>{F[k]=el.value;renderList()}});
}

/* ---------- chart ---------- */
function monthly(exps,mode){
  const m={};
  exps.forEach(e=>{
    if(mode==='service'&&e.ps&&e.pe&&e.pe>=e.ps){
      const n=diff(e.ps,e.pe)+1,per=e.amount/n;const d=parse(e.ps);
      for(let i=0;i<n;i++){const k=iso(d).slice(0,7);m[k]=(m[k]||0)+per;d.setDate(d.getDate()+1)}
    } else { const k=e.date.slice(0,7); m[k]=(m[k]||0)+(+e.amount); }
  });
  return m;
}
function monthsRange(range,first){
  const end=TODAY.slice(0,7);let start;
  if(range==='all') start=first||end; else start=addMonths(end+'-01',-(range-1)).slice(0,7);
  const out=[];let k=start+'-01';while(k.slice(0,7)<=end){out.push(k.slice(0,7));k=addMonths(k,1)}return out;
}
function renderChart(){
  const pool=paid().filter(e=>matchAsset(e,C.asset));
  const have=new Set(pool.map(e=>e.sub));
  const chips=S.cats.map(c=>{const subs=c.subs.filter(s=>have.has(s.id));if(!subs.length)return '';
    return `<div class="stack" style="gap:6px"><div class="tiny muted" style="font-weight:600">${esc(c.name)}</div><div class="chips">${subs.map(s=>{const on=C.subs.includes(s.id);return `<button class="chip" data-act="csub" data-id="${s.id}" aria-pressed="${on}" style="${on?'color:var(--s'+(C.slot[s.id]+1)+')':''}">${on?`<i style="background:var(--s${C.slot[s.id]+1})"></i>`:''}<span style="color:var(--ink)">${esc(s.name)}</span></button>`}).join('')}</div></div>`}).join('');
  const series=C.subs.filter(s=>have.has(s)).map(s=>({id:s,name:subName(s),color:`var(--s${C.slot[s]+1})`,m:monthly(pool.filter(e=>e.sub===s),C.mode)}));
  const firstAll=series.flatMap(s=>Object.keys(s.m)).sort()[0];
  const months=monthsRange(C.range,firstAll);
  while(months.length>2&&series.every(s=>!(s.m[months.at(-1)]>0)))months.pop();
  series.forEach(s=>{s.v=months.map(k=>s.m[k]||0);s.total=s.v.reduce((a,b)=>a+b,0);s.avg=s.total/months.length;s.max=Math.max(0,...s.v);s.maxAt=months[s.v.indexOf(s.max)]});
  main.innerHTML=topbar('圖表','同類別每月比較',gearBtn)+`
  <div class="stack">
    <div class="card pad stack">
      <div class="filters" style="flex-wrap:wrap">
        <select id="ca" class="sel on">${assetOptions(C.asset,true)}</select>
        <div class="seg" role="group" aria-label="計算方式"><button data-act="cmode" data-v="service" aria-pressed="${C.mode==='service'}">按使用期間</button><button data-act="cmode" data-v="paid" aria-pressed="${C.mode==='paid'}">按付款日</button></div>
        <div class="seg" role="group" aria-label="時間範圍">${[[12,'12 個月'],[24,'24 個月'],['all','全部']].map(([v,l])=>`<button data-act="crange" data-v="${v}" aria-pressed="${C.range==v}">${l}</button>`).join('')}</div>
      </div>
      <div class="hint">${C.mode==='service'?'有填使用期間的帳單，金額會平均分到它涵蓋的月份。季繳水費和月繳電費可以放在一起比。':'金額算在實際付款的月份，看的是每個月實際花出去多少。'}</div>
      ${chips||'<div class="empty">這個資產還沒有紀錄</div>'}
      <div class="hint">最多同時比較 4 項</div>
    </div>
    ${series.length?`<div class="card pad stack">
      <div class="legend">${series.map(s=>`<span><i style="background:${s.color}"></i>${esc(s.name)}</span>`).join('')}</div>
      <div class="chartwrap" id="cw"></div>
    </div>
    <div class="stats">${series.map(s=>`<div class="card tile"><div class="row"><i style="width:10px;height:10px;border-radius:50%;background:${s.color};display:inline-block"></i><span class="label">${esc(s.name)}</span></div><div class="v num">${money(s.total)}</div><div class="tiny muted">月平均 <span class="num">${money(s.avg)}</span>・最高 <span class="num">${s.maxAt?+s.maxAt.slice(5)+'月':''}</span></div></div>`).join('')}</div>
    <div class="card scrollx"><table class="tbl num"><thead><tr><th>月份</th>${series.map(s=>`<th>${esc(s.name)}</th>`).join('')}</tr></thead><tbody>
      ${months.map((k,i)=>({k,i})).reverse().map(({k,i})=>`<tr><td>${k.replace('-','/')}</td>${series.map(s=>{const v=s.v[i],p=i?s.v[i-1]:null;const dlt=p!=null&&p>0&&v>0?(v-p)/p*100:null;
        return `<td>${v?money(v):'<span class="faint">—</span>'}${dlt!=null&&Math.abs(dlt)>=1?` <span class="${dlt>0?'up':'down'}">${dlt>0?'▲':'▼'}${Math.abs(dlt).toFixed(0)}%</span>`:''}</td>`}).join('')}</tr>`).join('')}
    </tbody></table></div>`:''}
  </div>`;
  document.getElementById('ca').onchange=ev=>{C.asset=ev.target.value;renderChart()};
  if(series.length) drawChart(document.getElementById('cw'),months,series);
}
function niceMax(v){if(v<=0)return 100;const p=Math.pow(10,Math.floor(Math.log10(v)));const n=v/p;return (n<=1?1:n<=2?2:n<=2.5?2.5:n<=5?5:10)*p}
function drawChart(el,months,series){
  const W=Math.max(300,el.clientWidth),H=Math.round(Math.min(300,Math.max(210,W*.42))),L=52,R=10,T=10,B=26;
  const max=niceMax(Math.max(...series.flatMap(s=>s.v)));
  const x=i=>L+(months.length<2?0:i*(W-L-R)/(months.length-1)), y=v=>T+(H-T-B)*(1-v/max);
  let g='';for(let i=0;i<=4;i++){const v=max*i/4;g+=`<line x1="${L}" x2="${W-R}" y1="${y(v)}" y2="${y(v)}" stroke="var(--line)" stroke-width="1"/><text x="${L-8}" y="${y(v)+4}" text-anchor="end" font-size="11" fill="var(--faint)">${v>=1000?(v/1000).toFixed(v%1000?1:0)+'k':Math.round(v)}</text>`}
  const step=Math.ceil(months.length/(W<500?6:12));
  months.forEach((k,i)=>{if(i%step===0) g+=`<text x="${x(i)}" y="${H-6}" text-anchor="middle" font-size="11" fill="var(--faint)">${k.slice(5)==='01'||i===0?k.slice(2,4)+'/':''}${+k.slice(5)}</text>`});
  const lines=series.map(s=>{const pts=s.v.map((v,i)=>x(i).toFixed(1)+','+y(v).toFixed(1));const li=s.v.length-1;
    return `<polyline fill="none" stroke="${s.color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" points="${pts.join(' ')}"/>`+
      s.v.map((v,i)=>v>0&&months.length<=24?`<circle cx="${x(i)}" cy="${y(v)}" r="${i===li?4:2.5}" fill="${s.color}" stroke="var(--surface)" stroke-width="2"/>`:'').join('')}).join('');
  el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="每月金額折線圖">${g}${lines}<line id="xh" x1="0" x2="0" y1="${T}" y2="${H-B}" stroke="var(--muted)" stroke-dasharray="3 3" visibility="hidden"/><rect x="${L}" y="0" width="${W-L-R}" height="${H}" fill="transparent"/></svg><div class="tip" hidden></div>`;
  const svg=el.querySelector('svg'),tip=el.querySelector('.tip'),xh=el.querySelector('#xh');
  const move=ev=>{const r=svg.getBoundingClientRect(),px=(ev.clientX-r.left)*W/r.width;
    const i=Math.max(0,Math.min(months.length-1,Math.round((px-L)/((W-L-R)/Math.max(1,months.length-1)))));
    xh.setAttribute('x1',x(i));xh.setAttribute('x2',x(i));xh.setAttribute('visibility','visible');
    tip.hidden=false;tip.innerHTML=`<div style="font-weight:700;margin-bottom:4px">${months[i].replace('-',' 年 ')} 月</div>`+series.map(s=>`<div class="r"><span><i style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${s.color};margin-right:6px"></i>${esc(s.name)}</span><b class="num">${money(s.v[i])}</b></div>`).join('');
    const sx=x(i)*r.width/W;tip.style.left=Math.min(r.width-tip.offsetWidth,Math.max(0,sx+12>r.width-tip.offsetWidth?sx-tip.offsetWidth-12:sx+12))+'px';tip.style.top='8px'};
  svg.addEventListener('pointermove',move);svg.addEventListener('pointerdown',move);
  svg.addEventListener('pointerleave',()=>{tip.hidden=true;xh.setAttribute('visibility','hidden')});
}

/* ---------- jars ---------- */
function renderJar(){
  if(!jarOf(curJar)) curJar=S.jars[0]?.id||null;
  const j=jarOf(curJar);
  const picker=`<div class="chips">${S.jars.map(x=>`<button class="chip" data-act="pickjar" data-id="${x.id}" aria-pressed="${x.id===curJar}" style="${x.id===curJar?'color:var(--accent)':''}"><span style="color:var(--ink)">${esc(x.name)}</span><span class="num tiny ${jarBalance(x.id)<0?'':'muted'}" style="${jarBalance(x.id)<0?'color:var(--neg)':''}">${money0(jarBalance(x.id))}</span></button>`).join('')}<button class="chip" data-act="newjar">＋ 新增存錢罐</button></div>`;
  if(!j){main.innerHTML=topbar('存錢罐','',gearBtn)+`<div class="stack">${picker}<div class="card empty">還沒有存錢罐。按「新增存錢罐」建立第一個。</div></div>`;return}
  const l=jarLedger(j.id),bal=l.length?l.at(-1).bal:0;
  main.innerHTML=topbar('存錢罐','預先存好、專門用來付某些開銷的錢',gearBtn)+`
  <div class="stack">
    ${picker}
    <section class="card jar ${bal<0?'neg':''}">
      <div class="row between"><span class="label">${esc(j.name)}</span><div class="inline"><button class="btn sm ${j.pin?'primary':'ghost'}" data-act="pinjar" data-id="${j.id}">${j.pin?'已釘選到首頁':'釘選到首頁'}</button><button class="btn sm ghost" data-act="editjar" data-id="${j.id}">編輯</button></div></div>
      ${j.desc?`<div class="small muted">${esc(j.desc)}</div>`:''}
      <div class="bal num">${money(bal)}</div>
      <div class="note">${bal<0?`負數代表你先墊了 <b class="num">${money(-bal)}</b>`:`還可以扣 ${money(bal)}`}・低於 <span class="num">${money0(j.alert)}</span> 會提醒你</div>
      <div class="acts"><button class="btn primary sm" data-act="request" data-id="${j.id}">產生請款單</button><button class="btn sm" data-act="deposit" data-id="${j.id}">存入</button><button class="btn sm ghost" data-act="adjust" data-id="${j.id}">其他調整</button></div>
    </section>
    <div class="hint">記開銷時「從哪裡扣」選這個存錢罐，就會自動從這裡扣掉，不用再記第二次。</div>
    <div class="card list">${l.slice().reverse().map(r=>`<button class="item" data-act="${r.exp?'edit':'editmove'}" data-id="${r.exp?r.exp.id:r.mv.id}">
      <span class="grow"><span class="ell" style="display:block;font-weight:600">${esc(r.exp?title(r.exp):r.label)}${r.exp?.asset?`<span class="muted" style="font-weight:500">・${esc(assetShort(r.exp.asset))}</span>`:''}</span><span class="tiny muted">${slash(r.date)}${r.exp&&shortPeriod(r.exp)?'・用 '+shortPeriod(r.exp):''}${r.amount>0?'・存入':''}</span></span>
      <span style="text-align:right"><span class="amt num" style="display:block;${r.amount>0?'color:var(--accent)':''}">${money(r.amount,true)}</span><span class="tiny num ${r.bal<0?'':'faint'}" style="${r.bal<0?'color:var(--neg)':''}">餘 ${money(r.bal)}</span></span></button>`).join('')||'<div class="empty">還沒有紀錄</div>'}</div>
  </div>`;
}
function requestSheet(jid){
  const j=jarOf(jid),l=jarLedger(jid),bal=l.length?l.at(-1).bal:0;
  let lastIn=-1;l.forEach((r,i)=>{if(r.amount>0)lastIn=i});
  const since=l.slice(Math.max(0,lastIn));
  let months=j.months||12, off=new Set();
  openSheet('請款單',`<div class="form">
    <div class="fgroup">
      <div class="f"><label>預估未來</label><div class="seg" style="margin-left:auto" id="rm">${[3,6,12].map(m=>`<button type="button" data-m="${m}" aria-pressed="${m===months}">${m} 個月</button>`).join('')}</div></div>
      <div id="estlist"></div>
    </div>
    <div class="hint">依照每筆固定帳單過去 3 次的平均金額和週期來估算，只算從「${esc(j.name)}」扣的帳單。不想算進去的，點一下取消勾選。</div>
    <div class="card pad stack" style="gap:6px" id="calc"></div>
    <pre class="req" id="reqtxt"></pre>
    <button class="btn primary" id="copy">複製文字</button></div>`,root=>{
    const $=s=>root.querySelector(s);
    const draw=()=>{
      const items=forecast(jid,months), est=items.filter(i=>!off.has(i.key)).reduce((t,i)=>t+i.total,0), need=Math.max(0,est-bal);
      $('#estlist').innerHTML=items.map(i=>`<button class="est" data-k="${i.key}"><span class="chk${off.has(i.key)?'':' on'}">${off.has(i.key)?'':ICON.check}</span><span class="grow"><span class="ell" style="display:block;font-weight:600">${esc(i.label)}</span><span class="tiny muted">${i.note}・${i.n} 次 × 約 <span class="num">${money(i.each)}</span></span></span><span class="num" style="font-weight:600">${money(i.total)}</span></button>`).join('')||'<div class="empty">這段時間沒有預估的帳單</div>';
      $('#calc').innerHTML=`<div class="row between small"><span class="muted">未來 ${months} 個月預估開銷</span><span class="num">${money(est)}</span></div>
        <div class="row between small"><span class="muted">${bal<0?'加上你已經先墊的':'扣掉罐子裡還有的'}</span><span class="num">${bal<0?'+'+money(-bal):'−'+money(bal)}</span></div>
        <div class="row between" style="border-top:1px solid var(--line);padding-top:8px"><b>這次請款</b><b class="num" style="font-size:20px">${money(need)}</b></div>`;
      $('#reqtxt').textContent=`${j.name}存錢罐（${slash(since[0]?.date||TODAY)} 起）\n\n`+since.map(r=>`${slash(r.date).slice(5)}  ${r.exp?withAsset(fullName(r.exp),r.exp.asset):r.label}  ${money(r.amount,true)}`).join('\n')+
        `\n\n目前餘額：${money(bal)}${bal<0?'（我先墊付）':''}\n\n未來 ${months} 個月預估：\n`+items.filter(i=>!off.has(i.key)).map(i=>`・${i.label} ${i.n} 次 約 ${money(i.total)}`).join('\n')+
        `\n預估合計：${money(est)}\n\n這次請款：${money(need)}`;
      root.querySelectorAll('.est').forEach(b=>b.onclick=()=>{off.has(b.dataset.k)?off.delete(b.dataset.k):off.add(b.dataset.k);draw()});
    };
    root.querySelectorAll('#rm [data-m]').forEach(b=>b.onclick=()=>{months=+b.dataset.m;j.months=months;saveCfg();root.querySelectorAll('#rm [data-m]').forEach(x=>x.setAttribute('aria-pressed',x===b));draw()});
    $('#copy').onclick=()=>{const p=$('#reqtxt'),t=p.textContent;(navigator.clipboard?navigator.clipboard.writeText(t):Promise.reject()).then(()=>toast('已複製')).catch(()=>{const r=document.createRange();r.selectNodeContents(p);const s=getSelection();s.removeAllRanges();s.addRange(r);toast('已選取，請長按複製')})};
    draw();
  });
}
function moveForm(m,kind,jid){
  const isNew=!m; m=m||{id:uid('m'),date:TODAY,amount:'',kind,note:'',jar:jid};
  openSheet(isNew?(m.kind==='in'?'存入':'其他調整'):'編輯',`<div class="form">
    <div class="fgroup"><div class="amountin"><span>${cur()[1]}</span><input id="ja" type="number" inputmode="decimal" placeholder="${zeroText()}" value="${m.amount}"></div></div>
    <div class="fgroup">
      <div class="f"><label for="jj">存錢罐</label><select id="jj">${S.jars.map(j=>`<option value="${j.id}"${j.id===m.jar?' selected':''}>${esc(j.name)}</option>`).join('')}</select></div>
      <div class="f"><label for="jk">類型</label><select id="jk"><option value="in"${m.kind==='in'?' selected':''}>存入</option><option value="out"${m.kind==='out'?' selected':''}>扣除（不算開銷）</option></select></div>
      <div class="f"><label for="jd">日期</label><input id="jd" type="date" value="${m.date}"></div>
      <div class="f"><label for="jn">說明</label><input id="jn" type="text" value="${esc(m.note)}" placeholder="${m.kind==='in'?'例如：媽媽匯入':'例如：爸媽來澳現金'}"></div></div>
    <button class="btn primary" id="jsave">儲存</button>${isNew?'':'<div id="jdelwrap"><button class="btn danger" style="width:100%" id="jdel">刪除</button></div>'}</div>`,root=>{
    root.querySelector('#jsave').onclick=()=>{const a=+root.querySelector('#ja').value;if(!a){toast('請輸入金額');return}
      Object.assign(m,{amount:a,jar:root.querySelector('#jj').value,kind:root.querySelector('#jk').value,date:root.querySelector('#jd').value||TODAY,note:root.querySelector('#jn').value.trim()});
      B.putMove(clean(m)).catch(fail);if(isNew)S.moves.push(m);closeSheet();rerender();toast('已儲存')};
    const del=root.querySelector('#jdel');if(del)del.onclick=()=>confirmIn(root.querySelector('#jdelwrap'),()=>{S.moves=S.moves.filter(x=>x!==m);B.delMove(m.id).catch(fail);closeSheet();rerender();toast('已刪除')});
  });
}
function jarForm(j){
  const isNew=!j; j=j||{id:uid('jar'),name:'',desc:'',alert:500,pin:false};
  const used=!isNew&&(S.moves.some(m=>m.jar===j.id)||S.expenses.some(e=>e.src===j.id));
  openSheet(isNew?'新增存錢罐':'編輯存錢罐',`<div class="form">
    <div class="fgroup">
      <div class="f"><label for="jn">名稱</label><input id="jn" type="text" value="${esc(j.name)}" placeholder="例如：媽媽、爸爸、公司"></div>
      <div class="f col"><label for="jdsc">說明</label><textarea id="jdsc" placeholder="這個罐子的錢是誰的、用來做什麼">${esc(j.desc)}</textarea></div>
      <div class="f"><label for="jal">低於多少提醒</label><input id="jal" type="number" inputmode="decimal" value="${j.alert}" class="num"></div>
      <div class="f"><label>首頁顯示</label><div class="seg" style="margin-left:auto"><button type="button" data-pin="1" aria-pressed="${!!j.pin}">釘選</button><button type="button" data-pin="0" aria-pressed="${!j.pin}">不顯示</button></div></div>
    </div>
    <button class="btn primary" id="jsave">儲存</button>
    ${isNew?'':`<div id="jdelwrap">${used?'<div class="hint" style="text-align:center">這個罐子已經有紀錄，不能刪除。可以改成不釘選。</div>':'<button class="btn danger" style="width:100%" id="jdel">刪除這個存錢罐</button>'}</div>`}</div>`,root=>{
    let pin=!!j.pin;
    root.querySelectorAll('[data-pin]').forEach(b=>b.onclick=()=>{pin=b.dataset.pin==='1';root.querySelectorAll('[data-pin]').forEach(x=>x.setAttribute('aria-pressed',x===b))});
    root.querySelector('#jsave').onclick=()=>{const n=root.querySelector('#jn').value.trim();if(!n){toast('請輸入名稱');return}
      Object.assign(j,{name:n,desc:root.querySelector('#jdsc').value.trim(),alert:+root.querySelector('#jal').value||0,pin});
      if(isNew){S.jars.push(j);curJar=j.id}saveCfg();closeSheet();rerender();toast('已儲存')};
    const del=root.querySelector('#jdel');if(del)del.onclick=()=>confirmIn(root.querySelector('#jdelwrap'),()=>{S.jars=S.jars.filter(x=>x!==j);saveCfg();closeSheet();rerender();toast('已刪除')});
  });
}

/* ---------- expense form ---------- */
function lastOf(asset,sub,except){return paid().filter(e=>e.asset===asset&&e.sub===sub&&e.id!==except).sort((a,b)=>a.date<b.date?-1:1).at(-1)}
function expenseForm(e,preset){
  const isNew=!e;
  const lastNew=S.expenses.filter(x=>x.status==='paid').sort((a,b)=>a.date<b.date?-1:1).at(-1);
  e=e?{...e,extra:{...(e.extra||{})}}:{id:uid('e'),name:'',asset:lastNew?.asset||S.assets.find(a=>!a.archived)?.id||'',date:TODAY,amount:'',sub:lastNew?.sub||S.cats.find(c=>c.subs.length)?.subs[0]?.id||'',cycle:'monthly',method:S.methods[0],src:S.jars[0]?.id||NONE,note:'',ps:'',pe:'',status:'paid',due:'',extra:{},...(preset||{})};
  const orig=isNew?null:S.expenses.find(x=>x.id===e.id);
  const srcs=[...S.jars.map(j=>[j.id,j.name]),[NONE,NONE_LABEL]];
  openSheet(isNew?'記一筆':'編輯開銷',`<div class="form">
    <div class="fgroup"><div class="amountin"><span>${cur()[1]}</span><input id="xa" type="number" inputmode="decimal" placeholder="${zeroText()}" value="${e.amount}"></div><div class="hint" id="xlast" style="text-align:center;padding-bottom:12px"></div></div>
    <div class="fgroup">
      <div class="f"><label for="xs">類別</label><select id="xs">${subOptions(e.sub)}</select></div>
      <div class="f"><label for="xas">資產</label><select id="xas">${assetOptions(e.asset)}</select></div>
      <div class="f"><label for="xn">名稱</label><input id="xn" type="text" value="${esc(e.name)}" placeholder="選填，不填就顯示類別"></div>
    </div>
    <div class="fgroup">
      <div class="f"><label>狀態</label><div class="seg" style="margin-left:auto"><button type="button" data-st="paid" aria-pressed="${e.status==='paid'}">已付</button><button type="button" data-st="due" aria-pressed="${e.status==='due'}">待繳</button></div></div>
      <div class="f"><label for="xd" id="xdl">${e.status==='due'?'截止日':'付款日'}</label><input id="xd" type="date" value="${e.status==='due'?(e.due||e.date):e.date}"></div>
      <div class="f"><label>使用期間</label><div class="inline" style="flex:1;justify-content:flex-end;gap:4px"><input id="xps" type="date" value="${e.ps||''}" style="max-width:140px;flex:none"><span class="faint">–</span><input id="xpe" type="date" value="${e.pe||''}" style="max-width:140px;flex:none"></div></div>
      <div class="f" style="min-height:0;padding-top:0;border-top:0"><span class="hint" id="xtime" style="margin-left:auto"></span></div>
    </div>
    <div class="fgroup">
      <div class="f col"><label>從哪裡扣</label><div class="seg" id="xsrc">${srcs.map(([k,v])=>`<button type="button" data-src="${k}" aria-pressed="${e.src===k}">${esc(v)}</button>`).join('')}</div></div>
      <div class="f"><label for="xc">週期</label><select id="xc">${Object.entries(CYC).map(([k,[l]])=>`<option value="${k}"${e.cycle===k?' selected':''}>${l}</option>`).join('')}</select></div>
      <div class="f"><label for="xm">付款方式</label><select id="xm">${[...new Set([...S.methods,e.method].filter(Boolean))].map(m=>`<option${m===e.method?' selected':''}>${esc(m)}</option>`).join('')}</select></div>
    </div>
    <div class="fgroup">
      <div class="f col"><label for="xno">備註</label><textarea id="xno" placeholder="例如：帳單期間、幫誰代墊">${esc(e.note)}</textarea></div>
      ${S.fields.map((f,i)=>`<div class="f"><label for="xf${i}">${esc(f)}</label><input id="xf${i}" type="text" value="${esc(e.extra[f]||'')}"></div>`).join('')}
      <div class="f"><label>收據</label><span class="inline" style="margin-left:auto" id="xrwrap">${e.receipt?`<button type="button" class="btn sm" id="xrview">查看</button><button type="button" class="btn sm ghost" id="xrdel">移除</button>`:''}<label class="btn sm ghost"><input id="xr" type="file" accept="image/*,application/pdf" hidden><span id="xrlabel">${e.receipt?'換一張':'加照片或 PDF'}</span></label></span></div>
    </div>
    ${e.orig?`<div class="hint">Excel 原始名稱：${esc(e.orig)}</div>`:''}
    ${e.flag?`<div class="proto" id="xflag"><b>請確認：</b>${esc(e.flag)} <button class="btn sm" id="xok" style="margin-left:6px">沒問題</button></div>`:''}
    <button class="btn primary" id="xsave">儲存</button>
    ${isNew?'':'<div id="xdelwrap"><button class="btn danger" style="width:100%" id="xdel">刪除這筆</button></div>'}
  </div>`,root=>{
    const $=s=>root.querySelector(s);
    const sync=()=>{e.sub=$('#xs').value;e.asset=$('#xas').value;e.ps=$('#xps').value;e.pe=$('#xpe').value;e.date=$('#xd').value;
      const t=timing(e);$('#xtime').innerHTML=t?`這筆是 <b>${t}</b>（付款日在使用期間${t==='預付'?'之前':t==='後付'?'之後':'中間'}）`:'選填。有填的話圖表可以把金額平均分到每個月';
      const last=lastOf(e.asset,e.sub,e.id);$('#xlast').innerHTML=last?`上次：<b class="num">${money(last.amount)}</b>・${slash(last.date)}`:'';};
    const setSrc=k=>{if(!srcs.some(s=>s[0]===k))return;e.src=k;root.querySelectorAll('[data-src]').forEach(b=>b.setAttribute('aria-pressed',b.dataset.src===k))};
    const prefill=()=>{ if(!isNew) return; const last=lastOf($('#xas').value,$('#xs').value);
      if(last){ $('#xc').value=last.cycle; if(S.methods.includes(last.method))$('#xm').value=last.method; setSrc(last.src);
        $('#xa').placeholder=(+last.amount).toFixed(cur()[2]);
        if(last.ps&&last.pe&&CYC[last.cycle][1]){const ns=addDays(last.pe,1);$('#xps').value=ns;$('#xpe').value=addDays(addMonths(ns,CYC[last.cycle][1]),-1);} }
      sync(); };
    root.querySelectorAll('[data-src]').forEach(b=>b.onclick=()=>setSrc(b.dataset.src));
    root.querySelectorAll('[data-st]').forEach(b=>b.onclick=()=>{e.status=b.dataset.st;root.querySelectorAll('[data-st]').forEach(x=>x.setAttribute('aria-pressed',x===b));$('#xdl').textContent=e.status==='due'?'截止日':'付款日'});
    ['#xps','#xpe','#xd'].forEach(s=>$(s).oninput=sync);
    $('#xs').onchange=prefill;$('#xas').onchange=prefill;
    let rcpt;
    $('#xr').onchange=async ev=>{const f=ev.target.files[0];if(!f)return;$('#xrlabel').textContent='處理中…';
      try{rcpt=await readReceipt(f);$('#xrlabel').textContent=f.name}catch(err){rcpt=undefined;$('#xrlabel').textContent='加照片或 PDF';toast(err.message)}};
    if($('#xrview'))$('#xrview').onclick=()=>showReceipt(e.id);
    if($('#xrdel'))$('#xrdel').onclick=()=>{rcpt=null;$('#xrview')?.remove();$('#xrdel').remove();$('#xrlabel').textContent='加照片或 PDF';toast('按儲存後會移除')};
    if($('#xok'))$('#xok').onclick=()=>{e.flag=null;$('#xflag').remove();toast('按儲存後就不會再提醒')};
    $('#xsave').onclick=()=>{
      const a=$('#xa').value===''&&$('#xa').placeholder!==zeroText()?+$('#xa').placeholder:+$('#xa').value;
      if(!a){toast('請輸入金額');return}
      sync();
      Object.assign(e,{amount:a,cycle:$('#xc').value,method:$('#xm').value,note:$('#xno').value.trim(),name:$('#xn').value.trim()});
      e.due=e.status==='due'?e.date:'';
      S.fields.forEach((f,i)=>{e.extra[f]=$('#xf'+i).value});
      if(rcpt===null&&e.receipt){B.delReceipt(e.id).catch(fail);delete e.receipt}
      if(rcpt){B.putReceipt(e.id,rcpt).catch(fail);e.receipt={name:rcpt.name,type:rcpt.type}}
      if(orig)Object.assign(orig,e);else S.expenses.push(e);
      B.putExpense(clean(e)).catch(fail);closeSheet();rerender();toast(isNew?'已記下':'已更新');
    };
    if($('#xdel'))$('#xdel').onclick=()=>confirmIn($('#xdelwrap'),()=>{S.expenses=S.expenses.filter(x=>x.id!==e.id);B.delExpense(e.id).catch(fail);closeSheet();rerender();toast('已刪除')});
    isNew?prefill():sync();
    if(isNew)setTimeout(()=>$('#xa').focus(),250);
  });
}
function confirmIn(wrap,fn){wrap.innerHTML='<div class="inline" style="justify-content:center"><span class="small muted">確定要刪除嗎？</span><button class="btn sm" id="cn">取消</button><button class="btn sm danger" id="cy">刪除</button></div>';
  wrap.querySelector('#cy').onclick=fn;wrap.querySelector('#cn').onclick=()=>{closeSheet()}}

/* ---------- setup ---------- */
const delBtn=(act,id,label)=>`<button class="xbtn" data-act="${act}" data-id="${id}" aria-label="${label}">${ICON.trash}</button>`;
function renderSetup(){
  const tabs=[['assets','資產'],['cats','類別'],['bills','提醒'],['prefs','偏好'],['data','帳號']];
  let body='';
  if(setupTab==='assets'){
    body=S.folders.map(f=>`<div class="card"><div class="srow"><input type="text" value="${esc(f.name)}" data-edit="folder" data-id="${f.id}" style="font-weight:700" aria-label="分組名稱"><span class="tiny faint">分組</span>${S.folders.length>1?delBtn('delfolder',f.id,'刪除分組'):''}</div>
      ${S.assets.filter(a=>a.folder===f.id).map(a=>`<div class="srow" style="flex-wrap:wrap;padding-left:20px;${a.archived?'opacity:.5':''}"><button class="dot" data-act="pickicon" data-id="${a.id}" aria-label="換圖示" title="換圖示" style="width:34px;height:34px;border:1.5px dashed var(--faint)">${iconHTML(a.kind)}</button>
        <span class="stack" style="gap:0;flex:1 1 170px;min-width:0"><input type="text" value="${esc(a.name)}" data-edit="asset" data-id="${a.id}" aria-label="資產名稱"><input type="text" value="${esc(a.nick||'')}" placeholder="暱稱（選填）" data-edit="nick" data-id="${a.id}" class="small" style="color:var(--muted)" aria-label="暱稱"></span>
        <span class="inline" style="margin-left:auto"><select class="sel" data-edit="move" data-id="${a.id}" aria-label="移到分組" style="padding-block:4px;font-size:13px">${S.folders.map(x=>`<option value="${x.id}"${x.id===a.folder?' selected':''}>${esc(x.name)}</option>`).join('')}</select>
        <button class="btn sm ghost" data-act="archive" data-id="${a.id}">${a.archived?'恢復':'封存'}</button>${delBtn('delasset',a.id,'刪除資產')}</span></div>`).join('')}
      <div class="srow" style="padding-left:20px"><button class="btn sm" data-act="newasset" data-id="${f.id}">＋ 新增資產</button></div></div>`).join('')+
      `<button class="btn ghost" data-act="newfolder">＋ 新增分組</button><div class="hint">資產可以是房子、車子或任何想分開記帳的東西，記帳時是選填的。分組用來把資產分開，例如個人、公司、家人；只有一組時首頁不會顯示組名。封存的資產記帳時不會出現，舊紀錄會保留。</div>`;
  } else if(setupTab==='cats'){
    const cnt=id=>S.expenses.filter(e=>e.sub===id).length;
    body=S.cats.map(c=>`<div class="card"><div class="srow"><input type="text" value="${esc(c.name)}" data-edit="cat" data-id="${c.id}" style="font-weight:700" aria-label="大類名稱"><span class="tiny faint">大類</span>${c.subs.length?'':delBtn('delcat',c.id,'刪除大類')}</div>
      ${c.subs.map(s=>`<div class="srow" style="padding-left:28px"><input type="text" value="${esc(s.name)}" data-edit="sub" data-id="${s.id}" aria-label="小類名稱"><span class="tiny faint num">${cnt(s.id)} 筆</span>${cnt(s.id)?'':delBtn('delsub',s.id,'刪除小類')}</div>`).join('')}
      <div class="srow" style="padding-left:28px"><button class="btn sm" data-act="newsub" data-id="${c.id}">＋ 新增小類</button></div></div>`).join('')+
      `<button class="btn ghost" data-act="newcat">＋ 新增大類</button>
      <div class="card"><div class="srow"><b class="grow">付款方式</b></div>${S.methods.map((m,i)=>`<div class="srow" style="padding-left:28px"><input type="text" value="${esc(m)}" data-edit="method" data-id="${i}">${delBtn('delmethod',i,'刪除付款方式')}</div>`).join('')}<div class="srow" style="padding-left:28px"><button class="btn sm" data-act="newmethod">＋ 新增付款方式</button></div></div>
      <div class="card"><div class="srow"><b class="grow">自訂欄位</b><span class="tiny faint">會出現在記帳表單</span></div>${S.fields.map((f,i)=>`<div class="srow" style="padding-left:28px"><input type="text" value="${esc(f)}" data-edit="field" data-id="${i}">${delBtn('delfield',i,'刪除欄位')}</div>`).join('')}<div class="srow" style="padding-left:28px"><button class="btn sm" data-act="newfield">＋ 新增欄位</button></div></div>`;
  } else if(setupTab==='bills'){
    const up=upcoming();
    body=`<div class="hint">App 會用「上一次付款日 + 週期」推算下次大概什麼時候要繳。記了新的一筆，下次日期就會自動往後推。</div>
    <div class="card">${S.bills.map(b=>{const u=up.find(x=>x.asset===b.asset&&x.sub===b.sub);return `<div class="srow"><span class="grow"><b>${esc(subName(b.sub))}</b>・${esc(assetShort(b.asset))}<span class="tiny muted" style="display:block">${b.mode==='pre'?'預付':'後付'}・下次 ${u?slash(u.date):'—'}・提前 ${b.lead} 天提醒</span></span>
      <select class="sel" data-edit="bcycle" data-id="${b.id}" aria-label="週期" style="padding-block:4px;font-size:13px">${Object.entries(CYC).filter(([k])=>k!=='once').map(([k,[l]])=>`<option value="${k}"${b.cycle===k?' selected':''}>${l}</option>`).join('')}</select>
      <button class="btn sm ${b.on?'primary':'ghost'}" data-act="billon" data-id="${b.id}">${b.on?'開':'關'}</button></div>`}).join('')||'<div class="empty">還沒有固定帳單</div>'}<div class="srow"><button class="btn sm" data-act="newbill">＋ 新增固定帳單</button></div></div>
      <div class="card"><div class="srow"><b class="grow">存錢罐低於多少時提醒</b></div>${S.jars.map(j=>`<div class="srow" style="padding-left:28px"><span class="grow">${esc(j.name)}</span><input type="number" class="search num" style="width:120px;text-align:right;padding:6px 10px" value="${j.alert}" data-edit="jaralert" data-id="${j.id}" aria-label="${esc(j.name)} 提醒金額"></div>`).join('')}</div>
      <div class="hint">提醒目前顯示在首頁的「接下來要繳」。iPhone 推播通知會在下一版加上。</div>`;
  } else if(setupTab==='prefs'){
    loadFonts(Object.keys(FONTS));
    body=`<div class="card"><div class="srow"><b class="grow">貨幣</b><select class="sel" data-edit="currency" aria-label="貨幣">${Object.entries(CUR).map(([k,[n,sym]])=>`<option value="${k}"${S.currency===k?' selected':''}>${n} ${k}（${sym}）</option>`).join('')}</select></div></div>
      <div class="hint">只會改金額前面的符號和小數位數，不會換算匯率。</div>
      <div class="card"><div class="srow"><b class="grow">字體</b></div>${Object.entries(FONTS).map(([k,[n,fam]])=>`<button class="item" data-act="font" data-id="${k}" style="font-family:${fam?`'${fam}',`:''}${BASE_UI.replace(/"/g,"'")}">
        <span class="chk${(S.font||'system')===k?' on':''}" style="border-radius:50%">${(S.font||'system')===k?ICON.check:''}</span>
        <span class="grow"><span style="display:block;font-weight:600">${n}</span><span class="small muted">比記本・十一月電費 ${money(238.68)}・存錢罐</span></span></button>`).join('')}</div>
      <div class="card"><div class="srow"><b class="grow">自己的字型檔</b></div>
        ${CUSTOM_FONT?`<button class="item" data-act="font" data-id="custom" style="font-family:'BiledgerCustom',${BASE_UI.replace(/"/g,"'")}">
          <span class="chk${S.font==='custom'?' on':''}" style="border-radius:50%">${S.font==='custom'?ICON.check:''}</span>
          <span class="grow"><span style="display:block;font-weight:600">${esc(CUSTOM_FONT.name)}</span><span class="small muted">比記本・十一月電費 ${money(238.68)}・存錢罐</span></span></button>`
          :S.font==='custom'?'<div class="srow small muted">你在別的裝置選了自己的字型，這台裝置還沒上傳字型檔，所以先用系統字體。</div>':''}
        <div class="srow"><label class="btn sm"><input type="file" id="fontup" accept=".ttf,.otf,.woff,.woff2" hidden>${CUSTOM_FONT?'換一個字型檔':'上傳字型檔'}</label>${CUSTOM_FONT?'<button class="btn sm ghost" data-act="fontdel">從這台裝置移除</button>':''}<span id="fontmsg" class="tiny muted"></span></div></div>
      <div class="hint">字型檔只存在這台裝置，不會上傳到雲端，所以 iPhone、iPad、Mac 要各自上傳一次。建議把 App 加到主畫面，Safari 才不會在你很久沒打開時清掉它。</div>
      <div class="card"><div class="srow"><b class="grow">我的圖示</b><label class="btn sm"><input type="file" id="icup2" accept="image/svg+xml,image/png,image/jpeg,image/webp" multiple hidden>＋ 上傳</label></div>
        ${(S.icons||[]).map(ic=>`<div class="srow"><span class="dot" style="width:34px;height:34px">${iconHTML('u:'+ic.id)}</span><span class="grow">${esc(ic.name)}</span>${delBtn('delicon',ic.id,'刪除圖示')}</div>`).join('')||'<div class="srow small muted">還沒有自己的圖示。上傳後，在「資產」點圖示就可以選。</div>'}</div>`;
  } else {
    const flags=S.expenses.filter(e=>e.flag), u=USER||{}, linked=(u.providerData||[]).map(p=>p.providerId);
    body=`<div class="card"><div class="srow"><span class="grow"><b>${esc(u.displayName||u.email||'已登入')}</b><span class="tiny muted" style="display:block">${esc(u.email||'')}</span></span><button class="btn sm ghost" data-act="logout">登出</button></div>
      ${[['google','google.com'],['github','github.com']].map(([k,p])=>`<div class="srow"><span class="grow">${B.providerLabel(p)} 登入</span>${linked.includes(p)?'<span class="small" style="color:var(--accent);font-weight:600">已連結 ✓</span>':`<button class="btn sm" data-act="linkp" data-id="${k}">連結</button>`}</div>`).join('')}
      <div class="srow"><span class="grow">Email 和密碼登入</span>${linked.includes('password')?'<span class="small" style="color:var(--accent);font-weight:600">已設定 ✓</span>':'<button class="btn sm" data-act="addpw">設定密碼</button>'}</div></div>
      <div class="hint">連結多種登入方式後，用哪一種登入都會看到同一份資料。</div>
      <div class="card pad stack" style="gap:10px"><b>Excel</b><div class="small muted">匯出的檔案有：開銷明細、每個存錢罐、月度統計、年度統計、清單。檔案裡也藏了一份完整備份，之後可以用「匯入」還原。</div>
        <div class="inline"><button class="btn primary sm" data-act="exportx">匯出 Excel</button><label class="btn sm"><input type="file" id="impfile" accept=".xlsx" hidden>匯入 Excel</label></div><div id="impwrap"></div></div>
      ${flags.length?`<div class="card"><div class="srow"><b class="grow">匯入資料確認（${flags.length}）</b></div>
      ${flags.map(e=>`<button class="item" data-act="edit" data-id="${e.id}"><span class="grow"><b class="ell" style="display:block">${esc(withAsset(title(e),e.asset))}</b><span class="tiny muted">${esc(e.flag)}</span></span><span class="muted">›</span></button>`).join('')}</div>`:''}`;
  }
  main.innerHTML=topbar('設定','',`<button class="iconbtn" data-tab="home" aria-label="回首頁">${ICON.x}</button>`)+`<div class="stack"><div class="seg" role="tablist" style="align-self:flex-start">${tabs.map(([k,l])=>`<button data-act="stab" data-v="${k}" aria-pressed="${setupTab===k}">${l}</button>`).join('')}</div>${body}</div>`;
  const imp=document.getElementById('impfile');if(imp)imp.onchange=ev=>{const f=ev.target.files[0];if(f)importFlow(f,document.getElementById('impwrap'))};
  const fu=document.getElementById('fontup');if(fu)fu.onchange=async ev=>{const f=ev.target.files[0];if(!f)return;document.getElementById('fontmsg').textContent='讀取中…';
    try{await uploadFont(f);S.font='custom';applyFont('custom');saveCfg();toast('已套用字型');renderSetup()}catch(e){document.getElementById('fontmsg').textContent='';toast(e.message)}};
  const iu=document.getElementById('icup2');if(iu)iu.onchange=async ev=>{if(await uploadIcons(ev.target.files))renderSetup()};
}
function setupEdit(el){
  const id=el.dataset.id,v=el.value.trim(),k=el.dataset.edit;
  const sub=()=>{for(const c of S.cats){const s=c.subs.find(s=>s.id===id);if(s)return s}};
  if(k==='folder')S.folders.find(f=>f.id===id).name=v||'未命名';
  if(k==='asset')S.assets.find(a=>a.id===id).name=v||'未命名';
  if(k==='nick')S.assets.find(a=>a.id===id).nick=v;
  if(k==='move'){S.assets.find(a=>a.id===id).folder=v;renderSetup()}
  if(k==='cat')S.cats.find(c=>c.id===id).name=v||'未命名';
  if(k==='sub')sub().name=v||'未命名';
  if(k==='method'){const old=S.methods[+id];S.methods[+id]=v;S.expenses.forEach(e=>{if(e.method===old){e.method=v;B.putExpense(clean(e)).catch(fail)}})}
  if(k==='field')S.fields[+id]=v;
  if(k==='bcycle'){S.bills.find(b=>b.id===id).cycle=v;renderSetup()}
  if(k==='jaralert')jarOf(id).alert=+v||0;
  if(k==='currency'){S.currency=v;renderSetup()}
  if(k==='kind'){S.assets.find(a=>a.id===id).kind=v;renderSetup()}
  saveCfg();
}

/* ---------- sheet ---------- */
function openSheet(t,html,mount){
  const r=document.getElementById('sheetroot');
  r.innerHTML=`<div class="scrim" id="scrim"><div class="sheet" role="dialog" aria-modal="true" aria-label="${esc(t)}"><div class="grab"></div><div class="sh"><button class="btn sm ghost" id="shx">取消</button><h3>${esc(t)}</h3><span style="width:58px"></span></div>${html}</div></div>`;
  r.querySelector('#scrim').addEventListener('click',ev=>{if(ev.target.id==='scrim')closeSheet()});
  r.querySelector('#shx').onclick=closeSheet;
  mount&&mount(r);
}
function closeSheet(){document.getElementById('sheetroot').innerHTML=''}
document.addEventListener('keydown',ev=>{if(ev.key==='Escape')closeSheet()});

/* ---------- boot, routing & events ---------- */

const VIEWS={home:renderHome,list:renderList,chart:renderChart,jar:renderJar,setup:renderSetup};
function go(t){tab=VIEWS[t]?t:'home';
  document.querySelectorAll('.tab[data-tab]').forEach(b=>b.setAttribute('aria-current',b.dataset.tab===tab?'page':'false'));
  VIEWS[tab]();window.scrollTo(0,0);try{if(location.hash.slice(1)!==tab)history.replaceState(null,'','#'+tab)}catch(e){}}
function rerender(){const y=window.scrollY;VIEWS[tab]();window.scrollTo(0,y)}
document.addEventListener('click',ev=>{
  const t=ev.target.closest('[data-tab]');if(t&&!t.closest('#sheetroot')){go(t.dataset.tab);return}
  const b=ev.target.closest('[data-act]');if(!b||b.closest('#sheetroot .form'))return;const id=b.dataset.id,act=b.dataset.act;
  ({
    add:()=>expenseForm(),
    edit:()=>expenseForm(S.expenses.find(e=>e.id===id)),
    addfrom:()=>{const bl=S.bills.find(x=>x.id===id);expenseForm(null,{asset:bl.asset,sub:bl.sub,cycle:bl.cycle})},
    toggle:()=>{F.sel.has(id)?F.sel.delete(id):F.sel.add(id);rerender()},
    pickmode:()=>{F.pick=!F.pick;if(!F.pick)F.sel.clear();rerender()},
    pickall:()=>{filtered().forEach(e=>F.sel.add(e.id));rerender()},
    pickclear:()=>{F.sel.clear();rerender()},
    assetlist:()=>{F.asset=id;F.cat='all';F.period='year';F.q='';go('list')},
    request:()=>requestSheet(id||curJar), deposit:()=>moveForm(null,'in',id||curJar), adjust:()=>moveForm(null,'out',id||curJar),
    editmove:()=>moveForm(S.moves.find(m=>m.id===id)),
    openjar:()=>{curJar=id;go('jar')}, pickjar:()=>{curJar=id;renderJar()},
    newjar:()=>jarForm(), editjar:()=>jarForm(jarOf(id)),
    pinjar:()=>{const j=jarOf(id);j.pin=!j.pin;saveCfg();renderJar();toast(j.pin?'已釘選到首頁':'已從首頁移除')},
    goreview:()=>{setupTab='data';go('setup')}, gobills:()=>{setupTab='bills';go('setup')}, goassets:()=>{setupTab='assets';go('setup')},
    csub:()=>{const i=C.subs.indexOf(id);if(i>=0){C.subs.splice(i,1);delete C.slot[id]}else{if(C.subs.length>=4){toast('最多同時比較 4 項');return}
      const used=Object.values(C.slot);C.slot[id]=[0,1,2,3].find(n=>!used.includes(n));C.subs.push(id)}renderChart()},
    cmode:()=>{C.mode=b.dataset.v;renderChart()}, crange:()=>{C.range=b.dataset.v==='all'?'all':+b.dataset.v;renderChart()},
    stab:()=>{setupTab=b.dataset.v;renderSetup()},
    newfolder:()=>{S.folders.push({id:uid('f'),name:'新分組'});saveCfg();renderSetup()},
    delfolder:()=>{if(S.assets.some(a=>a.folder===id)){toast('這個分組裡還有資產，先把資產移到別組');return}S.folders=S.folders.filter(f=>f.id!==id);saveCfg();renderSetup()},
    newasset:()=>{S.assets.push({id:uid('a'),name:'新資產',nick:'',folder:id,kind:'house'});saveCfg();renderSetup()},
    fontdel:async()=>{await IDB.del('font').catch(()=>{});CUSTOM_FONT=null;if(S.font==='custom'){S.font='system';saveCfg()}applyFont(S.font);renderSetup()},
    delicon:()=>{S.icons=(S.icons||[]).filter(i=>i.id!==id);B.delIcon(id).catch(fail);renderSetup()},
    pickicon:()=>iconPicker(S.assets.find(a=>a.id===id)),
    delasset:()=>deleteAsset(S.assets.find(a=>a.id===id)),
    archive:()=>{const a=S.assets.find(a=>a.id===id);a.archived=!a.archived;saveCfg();renderSetup()},
    newcat:()=>{S.cats.push({id:uid('c'),name:'新大類',subs:[]});saveCfg();renderSetup()},
    delcat:()=>{S.cats=S.cats.filter(c=>c.id!==id);saveCfg();renderSetup()},
    newsub:()=>{S.cats.find(c=>c.id===id).subs.push({id:uid('s'),name:'新小類'});saveCfg();renderSetup()},
    delsub:()=>{S.cats.forEach(c=>c.subs=c.subs.filter(s=>s.id!==id));saveCfg();renderSetup()},
    newmethod:()=>{S.methods.push('新付款方式');saveCfg();renderSetup()},
    delmethod:()=>{S.methods.splice(+id,1);saveCfg();renderSetup()},
    newfield:()=>{S.fields.push('新欄位');saveCfg();renderSetup()},
    delfield:()=>{S.fields.splice(+id,1);saveCfg();renderSetup()},
    billon:()=>{const x=S.bills.find(x=>x.id===id);x.on=!x.on;saveCfg();renderSetup()},
    newbill:()=>billForm(),
    logout:()=>B.logout(),
    linkp:()=>B.linkProvider(id).then(()=>{toast('已連結');renderSetup()}).catch(e=>toast(e.message)),
    exportx:()=>doExport(),
    font:()=>{S.font=id;applyFont(id);saveCfg();renderSetup()},
    addpw:()=>passwordForm(),
  })[act]?.();
});
document.addEventListener('change',ev=>{const el=ev.target.closest('[data-edit]');if(el&&!el.closest('#sheetroot'))setupEdit(el)});
window.addEventListener('hashchange',()=>{const h=location.hash.slice(1);if(h&&h!==tab&&VIEWS[h])go(h)});
let rt;window.addEventListener('resize',()=>{clearTimeout(rt);rt=setTimeout(()=>{if(tab==='chart')renderChart()},150)});

/* ---------- cloud glue ---------- */
let USER=null, SYNCED=true, OFFLINE=false;
function syncLabel(){return `<span class="syncdot${SYNCED&&!OFFLINE?'':' off'}"></span>${OFFLINE?'離線中，連上網路後會同步':SYNCED?'已同步':'同步中…'}`}

// photos are shrunk so each receipt fits in one database document (< 1 MB)
async function readReceipt(f){
  const asData=blob=>new Promise((ok,no)=>{const r=new FileReader();r.onload=()=>ok(r.result);r.onerror=no;r.readAsDataURL(blob)});
  if(f.type==='application/pdf'){
    if(f.size>700000) throw new Error('PDF 太大了（上限 700 KB），可以改拍照片');
    return {name:f.name,type:f.type,data:await asData(f)};
  }
  if(!f.type.startsWith('image/')) throw new Error('只能放照片或 PDF');
  const img=await createImageBitmap(f).catch(()=>{throw new Error('讀不到這張照片，請換一張試試')});
  let max=1600,q=0.8,data='';
  for(let i=0;i<5;i++){
    const k=Math.min(1,max/Math.max(img.width,img.height)),c=document.createElement('canvas');
    c.width=Math.round(img.width*k);c.height=Math.round(img.height*k);c.getContext('2d').drawImage(img,0,0,c.width,c.height);
    data=c.toDataURL('image/jpeg',q); if(data.length<900000) break; max*=0.8;q-=0.1;
  }
  if(data.length>=900000) throw new Error('照片太大了，請換一張');
  return {name:f.name.replace(/\.\w+$/,'')+'.jpg',type:'image/jpeg',data};
}
async function showReceipt(id){
  openSheet('收據','<div class="spin"></div>',async root=>{
    const box=root.querySelector('.sheet');
    try{const r=await B.getReceipt(id);const body=box.querySelector('.spin');
      if(!r){body.outerHTML='<div class="empty">找不到這張收據</div>';return}
      if(r.type==='application/pdf'){const bin=atob(r.data.split(',')[1]),u=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)u[i]=bin.charCodeAt(i);
        const url=URL.createObjectURL(new Blob([u],{type:'application/pdf'}));body.outerHTML=`<div class="stack" style="align-items:center"><div>${esc(r.name)}</div><a class="btn primary" href="${url}" target="_blank" rel="noopener">打開 PDF</a></div>`}
      else body.outerHTML=`<div class="receipt"><img src="${r.data}" alt="收據"></div>`;
    }catch(e){box.querySelector('.spin').outerHTML=`<div class="empty">讀取失敗：${esc(e.message)}</div>`}
  });
}
function billForm(){
  if(!S.assets.length){toast('先到「資產」新增一個資產');return}
  openSheet('新增固定帳單',`<div class="form"><div class="fgroup">
    <div class="f"><label for="ba">資產</label><select id="ba">${assetOptions(S.assets[0].id)}</select></div>
    <div class="f"><label for="bs">類別</label><select id="bs">${subOptions(S.cats.find(c=>c.subs.length)?.subs[0]?.id)}</select></div>
    <div class="f"><label for="bc">週期</label><select id="bc">${Object.entries(CYC).filter(([k])=>k!=='once').map(([k,[l]])=>`<option value="${k}">${l}</option>`).join('')}</select></div>
    <div class="f"><label for="bm">付款方式</label><select id="bm"><option value="pre">預付（先付錢再使用）</option><option value="post">後付（用完才收帳單）</option></select></div>
    <div class="f"><label for="bl">提前幾天提醒</label><input id="bl" type="number" inputmode="numeric" value="5"></div></div>
    <div class="hint">下次繳費日會用這個資產、這個類別最近一次付款的日期加上週期來推算。</div>
    <button class="btn primary" id="bsave">新增</button></div>`,root=>{
    const $=s=>root.querySelector(s);
    $('#bsave').onclick=()=>{const b={id:uid('b'),asset:$('#ba').value,sub:$('#bs').value,cycle:$('#bc').value,mode:$('#bm').value,lead:+$('#bl').value||0,on:true};
      if(S.bills.some(x=>x.asset===b.asset&&x.sub===b.sub)){toast('這個帳單已經有了');return}
      S.bills.push(b);saveCfg();closeSheet();renderSetup();toast('已新增')};
  });
}
const xctx=()=>({S,CYC,NONE,NONE_LABEL,assetOf,subName,catOf,jarOf,title,timing,defaultConfig,uid,cur:cur(),buyName:buyName()});
function iconPicker(a){
  const tile=(k,l)=>`<button class="card pad stack" data-kind="${k}" style="align-items:center;gap:6px;${a.kind===k?'border-color:var(--accent);color:var(--accent)':''}"><span class="dot">${iconHTML(k)}</span><span class="small ell" style="max-width:100%">${esc(l)}</span></button>`;
  const grid=`style="grid-template-columns:repeat(auto-fill,minmax(96px,1fr))"`;
  openSheet('選擇圖示',`<div class="stats" ${grid}>${KINDS.map(([k,l])=>tile(k,l)).join('')}</div>
    <div class="hint" style="margin-top:8px">房子、公寓、土地、店面這幾種，列表裡會顯示暱稱，不會顯示完整地址。</div>
    <div class="gh">我的圖示</div>
    ${(S.icons||[]).length?`<div class="stats" ${grid}>${S.icons.map(ic=>tile('u:'+ic.id,ic.name)).join('')}</div>`:''}
    <label class="btn ghost" style="margin-top:10px;width:100%"><input type="file" id="icup" accept="image/svg+xml,image/png,image/jpeg,image/webp" multiple hidden>＋ 上傳自己的圖示</label>
    <div class="hint" style="margin-top:6px">可以選 SVG、PNG 或 JPG，一次選好幾個也可以。上傳後所有裝置都看得到；在「設定 → 偏好」可以刪除。</div>`,root=>{
    root.querySelector('#icup').onchange=async ev=>{const n=await uploadIcons(ev.target.files);if(n){closeSheet();iconPicker(a)}};
    root.querySelectorAll('[data-kind]').forEach(b=>b.onclick=()=>{a.kind=b.dataset.kind;saveCfg();closeSheet();renderSetup()});
  });
}
async function readIcon(f){
  const asData=blob=>new Promise((ok,no)=>{const r=new FileReader();r.onload=()=>ok(r.result);r.onerror=no;r.readAsDataURL(blob)});
  if(f.type==='image/svg+xml'){if(f.size>150000)throw new Error(`${f.name} 太大了（SVG 上限 150 KB）`);return asData(f)}
  const img=await createImageBitmap(f).catch(()=>{throw new Error(`讀不到 ${f.name}`)});
  const c=document.createElement('canvas');c.width=c.height=128;const k=Math.min(128/img.width,128/img.height);
  c.getContext('2d').drawImage(img,(128-img.width*k)/2,(128-img.height*k)/2,img.width*k,img.height*k);return c.toDataURL('image/png');
}
async function uploadIcons(files){
  let n=0;
  for(const f of files){try{const data=await readIcon(f);const id=uid('i');const ic={name:f.name.replace(/\.\w+$/,''),data};
    S.icons=[...(S.icons||[]),{id,...ic}];await B.putIcon(id,ic);n++}catch(e){toast(e.message)}}
  if(n)toast(`已加入 ${n} 個圖示`);return n;
}
function deleteAsset(a){
  const used=S.expenses.filter(e=>e.asset===a.id);
  openSheet('刪除資產',`<div class="form"><div class="card pad stack" style="gap:6px"><b>${esc(a.name)}</b>
    <div class="small muted">${used.length?`有 <b>${used.length}</b> 筆開銷屬於這個資產。刪除後，這些開銷會保留，但會變成「不指定資產」，這個資產的固定帳單提醒也會一起刪除。如果只是暫時不用，建議改用「封存」。`:'這個資產還沒有任何開銷，可以放心刪除。'}</div></div>
    <button class="btn danger" id="adel">刪除</button></div>`,root=>{
    root.querySelector('#adel').onclick=()=>{
      used.forEach(e=>{e.asset='';B.putExpense(clean(e)).catch(fail)});
      S.assets=S.assets.filter(x=>x!==a);S.bills=S.bills.filter(b=>b.asset!==a.id);
      saveCfg();closeSheet();renderSetup();toast('已刪除');
    };
  });
}
function passwordForm(){
  openSheet('設定密碼',`<div class="form"><div class="fgroup">
    <div class="f"><label for="pe">Email</label><input id="pe" type="email" autocomplete="email" value="${esc(USER?.email||'')}"></div>
    <div class="f"><label for="pp">密碼</label><input id="pp" type="password" autocomplete="new-password" placeholder="至少 6 個字"></div></div>
    <div class="hint">設定後，也可以用這個 Email 和密碼登入，看到的是同一份資料。</div>
    <button class="btn primary" id="psave">設定</button></div>`,root=>{
    root.querySelector('#psave').onclick=async ev=>{ev.target.disabled=true;
      try{await B.addPassword(root.querySelector('#pe').value,root.querySelector('#pp').value);closeSheet();toast('已設定密碼');renderSetup()}
      catch(e){ev.target.disabled=false;toast(e.message)}};
  });
}
async function doExport(){
  toast('正在產生 Excel…');
  try{const blob=await exportWorkbook(xctx());const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=`Biledger_${TODAY}.xlsx`;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},4000)}
  catch(e){console.error(e);toast('匯出失敗：'+e.message)}
}
async function importFlow(file,wrap){
  wrap.innerHTML='<div class="spin"></div>';
  let d;try{d=await parseWorkbook(file,xctx())}catch(e){console.error(e);wrap.innerHTML=`<div class="proto">讀不懂這個檔案：${esc(e.message)}</div>`;return}
  const hasData=S.expenses.length||S.moves.length;
  wrap.innerHTML=`<div class="stack" style="gap:8px"><div class="small">找到 <b>${d.expenses.length}</b> 筆開銷、<b>${d.moves.length}</b> 筆存錢罐存入／調整、<b>${d.cfg.assets.length}</b> 個資產${d.backup?'（完整備份）':''}。</div>
    ${hasData?'<div class="proto">匯入會<b>取代</b>目前雲端上的所有開銷、存錢罐和設定。收據照片會保留。</div>':''}
    <div class="inline"><button class="btn sm" id="imno">取消</button><button class="btn sm ${hasData?'danger':'primary'}" id="imyes">${hasData?'取代並匯入':'匯入'}</button></div></div>`;
  wrap.querySelector('#imno').onclick=()=>{wrap.innerHTML=''};
  wrap.querySelector('#imyes').onclick=async()=>{wrap.innerHTML='<div class="spin"></div><div class="tiny muted" style="text-align:center">匯入中…</div>';
    try{await B.replaceAll(clean(d.cfg),clean(d.expenses),clean(d.moves));toast('匯入完成');go('home')}
    catch(e){console.error(e);wrap.innerHTML=`<div class="proto">匯入失敗：${esc(e.code||e.message)}</div>`}};
}

/* ---------- sign-in gate ---------- */
const gate=document.getElementById('gate'), shell=document.getElementById('shell');
const GLOGO=`<div class="logo">${ICON.jar}</div>`;
const GICON={google:'<svg viewBox="0 0 24 24"><path fill="#4285F4" d="M22.5 12.3c0-.8-.1-1.5-.2-2.2H12v4.2h5.9a5 5 0 0 1-2.2 3.3v2.7h3.6c2-1.9 3.2-4.7 3.2-8z"/><path fill="#34A853" d="M12 23c3 0 5.5-1 7.3-2.7l-3.6-2.7c-1 .7-2.2 1.1-3.7 1.1-2.9 0-5.3-1.9-6.2-4.5H2.1v2.8A11 11 0 0 0 12 23z"/><path fill="#FBBC05" d="M5.8 14.2a6.6 6.6 0 0 1 0-4.3V7.1H2.1a11 11 0 0 0 0 9.9z"/><path fill="#EA4335" d="M12 5.4c1.6 0 3.1.6 4.2 1.7l3.2-3.2A11 11 0 0 0 2.1 7.1l3.7 2.8C6.7 7.3 9.1 5.4 12 5.4z"/></svg>',
  github:'<svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 .5a11.5 11.5 0 0 0-3.6 22.4c.6.1.8-.3.8-.6v-2c-3.2.7-3.9-1.5-3.9-1.5-.5-1.3-1.3-1.7-1.3-1.7-1-.7.1-.7.1-.7 1.2.1 1.8 1.2 1.8 1.2 1 1.8 2.8 1.3 3.5 1 .1-.8.4-1.3.7-1.6-2.6-.3-5.3-1.3-5.3-5.7 0-1.3.5-2.3 1.2-3.1-.1-.3-.5-1.5.1-3.1 0 0 1-.3 3.2 1.2a11 11 0 0 1 5.8 0c2.2-1.5 3.2-1.2 3.2-1.2.6 1.6.2 2.8.1 3.1.7.8 1.2 1.8 1.2 3.1 0 4.4-2.7 5.4-5.3 5.7.4.4.8 1.1.8 2.2v3.3c0 .3.2.7.8.6A11.5 11.5 0 0 0 12 .5z"/></svg>'};
function showGate(html){gate.hidden=false;shell.hidden=true;closeSheet();gate.innerHTML=`<div class="card">${html}</div>`}
function showLogin(err){
  const keep=gate.querySelector('#lem')?.value||'';
  showGate(`${GLOGO}<h1>Biledger <span style="font-weight:500;color:var(--muted)">比記本</span></h1><div class="small muted">記帳、管理資產和存錢罐，還會提醒你繳費。資料存在雲端，手機、平板、電腦都能用。</div>
    ${err?`<div class="${err.ok?'proto':'err'}">${esc(err.ok?err.ok:err)}</div>`:''}
    <button class="btn ghost" data-login="google">${GICON.google}用 Google 登入</button>
    <button class="btn ghost" data-login="github">${GICON.github}用 GitHub 登入</button>
    <div class="row tiny faint" style="gap:10px"><span class="grow" style="border-top:1px solid var(--line)"></span>或用 Email<span class="grow" style="border-top:1px solid var(--line)"></span></div>
    <form id="lform" class="stack" style="gap:8px">
      <input class="search" id="lem" type="email" autocomplete="email" placeholder="Email" value="${esc(keep)}">
      <input class="search" id="lpw" type="password" autocomplete="current-password" placeholder="密碼">
      <button class="btn primary" type="submit">登入</button>
      <div class="row between"><button type="button" class="btn sm ghost" id="lnew">建立新帳號</button><button type="button" class="btn sm ghost" id="lforgot">忘記密碼</button></div>
    </form>`);
  gate.querySelectorAll('[data-login]').forEach(b=>b.onclick=async()=>{b.disabled=true;try{await B.login(b.dataset.login)}catch(e){showLogin(e.message)}});
  const em=gate.querySelector('#lem'), pw=gate.querySelector('#lpw');
  const go2=async create=>{if(!em.value.trim()){showLogin('請輸入 Email');return}gate.querySelectorAll('button').forEach(b=>b.disabled=true);
    try{await B.emailLogin(em.value,pw.value,create)}catch(e){showLogin(e.message)}};
  gate.querySelector('#lform').onsubmit=ev=>{ev.preventDefault();go2(false)};
  gate.querySelector('#lnew').onclick=()=>{if(pw.value.length<6){showLogin('建立新帳號：請輸入 Email，並設定至少 6 個字的密碼');return}go2(true)};
  gate.querySelector('#lforgot').onclick=async()=>{if(!em.value.trim()){showLogin('請先輸入 Email，再按「忘記密碼」');return}
    try{await B.resetPassword(em.value);showLogin({ok:`已寄出重設密碼的信到 ${em.value.trim()}，請到信箱收信（也看一下垃圾信件匣）。`})}catch(e){showLogin(e.message)}};
}
function showWelcome(){
  showGate(`${GLOGO}<h1>歡迎</h1><div class="small muted">這個帳號還沒有資料。可以匯入之前匯出的 Excel，或從空白開始。</div>
    <label class="btn primary"><input type="file" id="gimp" accept=".xlsx" hidden>匯入 Excel</label><div id="gwrap"></div>
    <button class="btn ghost" id="gblank">從空白開始</button><button class="btn sm ghost" id="gout">登出</button>`);
  gate.querySelector('#gimp').onchange=ev=>{const f=ev.target.files[0];if(f)importFlow(f,gate.querySelector('#gwrap'))};
  gate.querySelector('#gblank').onclick=()=>{const c=defaultConfig();B.saveConfig(c).catch(e=>toast(e.message))};
  gate.querySelector('#gout').onclick=()=>B.logout();
}
let unsub=null, booted=false, renderT, got={}, migrating=false;
function tryMigrate(){
  if(migrating||!got.cfg||!got.exp||!got.mov||S.v>=3)return;
  migrating=true;const {cfg,expenses}=migrateV2(S);
  B.replaceAll(clean(cfg),clean(expenses),clean(S.moves)).then(()=>toast('已更新成新版的類別和付款方式')).catch(fail).finally(()=>{migrating=false});
}
// re-draw after cloud updates, but never while a field on the page is being edited
function scheduleRender(){clearTimeout(renderT);renderT=setTimeout(()=>{if(!booted)return;const a=document.activeElement;
  if(a&&main.contains(a)&&/^(INPUT|SELECT|TEXTAREA)$/.test(a.tagName)&&a.id!=='q'){a.addEventListener('blur',scheduleRender,{once:true});return}rerender()},60)}
function startApp(){gate.hidden=true;shell.hidden=false;if(!booted){booted=true;go(location.hash.slice(1)||'home')}else rerender()}
B.onUser((user,err)=>{
  if(err){showLogin(err.message);return}
  unsub?.();unsub=null;USER=user;booted=false;
  if(!user){showLogin();return}
  showGate('<div class="spin"></div><div class="small muted">讀取資料中…</div>');
  let have=null;got={cfg:false,exp:false,mov:false};
  unsub=B.subscribe(user.uid,{
    onConfig:(cfg,fromCache,e)=>{
      if(e){showGate(`<div class="err">讀不到資料：${esc(e.code||e.message)}。請確認 Firestore 的安全規則已經設定好。</div><button class="btn ghost" id="gout">登出</button>`);gate.querySelector('#gout').onclick=()=>B.logout();return}
      if(!cfg){if(!fromCache){have=false;showWelcome()}return}
      if(cfgDirty)return;
      have=true;got.cfg=true;Object.assign(S,defaultConfig(),cfg);S.v=cfg.v||2;if(!jarOf(curJar))curJar=S.jars[0]?.id||null;
      applyFont(S.font);startApp();tryMigrate();
    },
    onExpenses:list=>{S.expenses=list;got.exp=true;tryMigrate();scheduleRender()},
    onMoves:list=>{S.moves=list;got.mov=true;tryMigrate();scheduleRender()},
    onIcons:list=>{S.icons=list;scheduleRender()},
    onSync:(ok,cache)=>{SYNCED=ok;if(tab==='home'&&booted){const el=document.querySelector('.top .sub');if(el)el.innerHTML=syncLabel()}},
  });
});
window.addEventListener('online',()=>{OFFLINE=false;scheduleRender()});
window.addEventListener('offline',()=>{OFFLINE=true;scheduleRender()});
OFFLINE=!navigator.onLine;
if('serviceWorker' in navigator&&location.protocol==='https:')navigator.serviceWorker.register('sw.js').catch(()=>{});
