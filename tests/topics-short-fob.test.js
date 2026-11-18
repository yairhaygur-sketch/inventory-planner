/* ============ SHORT_ONE + FOB_COLS (החלטות המעתד 18.11) ============
   SHORT_ONE — בנושא «מלאי 0 / צפוי להיגמר»: פעולת אספקה אחת, בלי כמות,
   לפי הסדר: אין רכש → לקוח ממתין ויש ETA → רכש בלי ETA → ETA מאוחר → Back Order.
   FOB_COLS — «מחיר FOB» (BZ) ו«מטבע» (BE) בכל ארבעת הנושאים.
   תצוגה בלבד: המנוע, ההמלצות וקובץ SAP לא נוגעים. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx'),ETAX=path.join(SD,'zmrp-demo-eta.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');
 const tab=t=>p.evaluate(t=>{TP.tab=t;TP.cat={};TP.sort={};TP.q='';TP.f={};tpRender()},t);

 /* 1 · סדר העדיפויות — על רשומות סינתטיות, עם ובלי דוח ETA */
 const O=await p.evaluate(()=>{const save=ETA;const base={free:0,unsQty:0,cust:0,po:0,etaHas:false,etaFirst:null,mark:null,isOD:false,isPDItem:false,noStock:false,rate:5};
  const dd=new Date(ANCHOR);dd.setDate(dd.getDate()+40);const d=dd.toISOString().slice(0,10);
  const run=o=>tpShortAct({...base,...o});
  ETA=save||{parts:1};
  const res={
   noPo:run({cust:3}),                                        /* 1 גובר גם על לקוח */
   custEta:run({cust:3,po:5,etaHas:true,etaFirst:d}),          /* 2 */
   noEta:run({po:5}),                                         /* 3 */
   custNoEta:run({cust:3,po:5}),                              /* לקוח בלי ETA → 3 */
   late0:run({po:5,etaHas:true,etaFirst:d}),                  /* 4 — המדף ריק */
   lateMonth:run({free:4,unsQty:2,po:5,etaHas:true,etaFirst:d}), /* 4 — יחסר החודש */
   bo:run({free:400,po:5,etaHas:true,etaFirst:d}),            /* 5 */
  };
  ETA=null;res.noFile=run({cust:3,po:5,etaHas:true,etaFirst:d});  /* בלי דוח ETA אין תאריך → לבקש ETA */
  ETA=save;return res});
 ok('אין רכש פתוח → «לפתוח הזמנת רכש» (גובר גם על לקוח ממתין)',O.noPo==='לפתוח הזמנת רכש',O.noPo);
 ok('לקוח ממתין ויש ETA → «למסור ללקוח תאריך: …»',/^למסור ללקוח תאריך: \d/.test(O.custEta),O.custEta);
 ok('רכש בלי ETA → «לבקש ETA מהספק» (גם כשלקוח ממתין)',O.noEta==='לבקש ETA מהספק'&&O.custNoEta==='לבקש ETA מהספק');
 ok('יש ETA אבל המדף ריק או שיחסר החודש → «לזרז ספק»',O.late0==='לזרז ספק'&&O.lateMonth==='לזרז ספק');
 ok('כל השאר → «לבדוק Back Order אצל היצרן»',O.bo==='לבדוק Back Order אצל היצרן',O.bo);
 ok('בלי דוח ETA — אין תאריך, ולכן «לבקש ETA מהספק»',O.noFile==='לבקש ETA מהספק',O.noFile);

 /* 2 · הטבלה: כל תא «פעולה» = tpShortAct, ואין בו ספרה מלבד תאריך */
 await tab('short');
 const T=await p.evaluate(()=>{const ths=[...document.querySelectorAll('#lite thead th')].map(t=>t.textContent.trim());
  const trs=[...document.querySelectorAll('#lite tbody tr')],ia=ths.length-1;
  return {n:trs.length,bad:trs.map((tr,j)=>{const a=tr.children[ia].textContent.trim(),r=TP.shown[j];
   return a===tpShortAct(r)&&(!/\d/.test(a)||/^למסור ללקוח תאריך: /.test(a))?null:r.pn}).filter(Boolean),
   kinds:[...new Set(TP.shown.map(tpShortAct).map(a=>a.replace(/: .*/,'')))]}});
 ok('מלאי 0: כל תא «פעולה» = פעולת האספקה האחת, בלי כמות',T.n>0&&!T.bad.length,`${T.n} שורות · ${T.kinds.join(' / ')}`);

 /* 3 · הכרטיס: «מה לעשות» — פריט אחד בלבד, זהה לטבלה */
 const C=await p.evaluate(()=>{const r=TP.shown[0];tpdShow(r);const li=[...document.querySelectorAll('#tpd .qa.act li')].map(x=>x.textContent.trim());
  const res={li,exp:tpShortAct(r)};tpdClose();return res});
 ok('הכרטיס במלאי 0: «מה לעשות» — פעולה אחת, זהה לטבלה',C.li.length===1&&C.li[0]===C.exp,C.li.join(' | '));

 /* 4 · FOB_COLS בכל ארבעת הנושאים */
 const F={};
 for(const t of ['cust','trend','short','fix']){await tab(t);
  F[t]=await p.evaluate(()=>{const ths=[...document.querySelectorAll('#lite thead th')].map(t=>t.textContent.trim());
   const iP=ths.findIndex(t=>t.startsWith('מחיר FOB')),iC=ths.findIndex(t=>t==='מטבע'||t.startsWith('מטבע'));
   const trs=[...document.querySelectorAll('#lite tbody tr')];
   const bad=iP<0||iC<0?['no-col']:trs.map((tr,j)=>{const r=TP.shown[j],P=tr.children[iP].textContent.trim(),Cc=tr.children[iC].textContent.trim();
    const eP=r.priceMissing?'—':Number(r.price).toLocaleString('he-IL',{maximumFractionDigits:2}),eC=r.currency||'—';
    return P===eP&&Cc===eC?null:r.pn}).filter(Boolean);
   /* ה-CSV: כל כותרת פעם אחת, וההסרה נוגעת רק ל«מטבע» הכפול */
   const cols=TP_HEAD.concat(TOPICS[TP.tab].cols),raw=cols.flatMap(c=>c.x||[[c.l,c.v]]).map(x=>x[0]);
   const dup=[...new Set(raw.filter((h,i)=>raw.indexOf(h)!==i))];
   return {n:trs.length,bad,dup}})}
 ok('«מחיר FOB» ו«מטבע» בכל ארבעת הנושאים, ערכים = price/currency',Object.values(F).every(x=>x.n>0&&!x.bad.length),
   JSON.stringify(Object.fromEntries(Object.entries(F).map(([k,v])=>[k,v.bad.length?v.bad.slice(0,2):v.n]))));
 ok('ב-CSV הכותרת הכפולה היחידה היא «מטבע», והיא יוצאת פעם אחת',Object.values(F).every(x=>x.dup.every(h=>h==='מטבע')),
   JSON.stringify(Object.fromEntries(Object.entries(F).map(([k,v])=>[k,v.dup]))));
 {await tab('fix');const [dl]=await Promise.all([p.waitForEvent('download'),p.click('#lite [data-tp="csv"]')]);
  const H=fs.readFileSync(await dl.path(),'utf8').replace(/^﻿/,'').split('\r\n')[0].split(',');
  ok('CSV של טיוב נתונים: «מחיר FOB» פעם אחת, «מטבע» פעם אחת',H.filter(h=>h==='"מחיר FOB"').length===1&&H.filter(h=>h==='"מטבע"').length===1,H.length+' עמודות')}

 /* 5 · עם דוח ETA — אותם כללים, ואין שגיאות */
 if(fs.existsSync(ETAX)){await p.setInputFiles('#fe',ETAX);await p.waitForTimeout(2500);await tab('short');
  const E=await p.evaluate(()=>({n:TP.shown.length,eta:!!ETA,custEta:TP.shown.filter(r=>/^למסור ללקוח תאריך/.test(tpShortAct(r))).length,
   bad:TP.shown.filter(r=>{const a=tpShortAct(r);return /^למסור ללקוח תאריך/.test(a)&&!(r.cust>0&&r.etaHas&&r.po>0)}).length}));
  ok('עם דוח ETA: «למסור ללקוח תאריך» רק כשיש לקוח, רכש ו-ETA',E.eta&&E.n>0&&E.bad===0,JSON.stringify(E))}

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0);
})().catch(e=>{console.error(e);process.exit(1)});
