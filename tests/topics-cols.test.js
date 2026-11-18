/* ============ TP_COLS — סדר העמודות ומה מוצג, לכל נושא (החלטת המעתד 18.11) ============
   גרירת כותרת · חלון «עמודות» (סימון, ↑↓, חזרה לברירת המחדל) · שמירה בדפדפן לכל נושא
   בנפרד · «מק״ט» תמיד ראשון ותמיד מוצג · כללי ההקשר נשארים · ה-CSV לא משתנה. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 const load=async()=>{await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');await p.setInputFiles('#f',XL);
  await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table')};
 await load();await p.evaluate(()=>{try{localStorage.removeItem('planner_tpcols_v1')}catch(_){}});
 const tab=t=>p.evaluate(t=>{TP.tab=t;TP.cat={};TP.sort={};TP.q='';TP.f={};tpRender()},t);
 const heads=()=>p.evaluate(()=>[...document.querySelectorAll('#lite thead th')].map(t=>t.dataset.tpSort));
 const csvHead=async()=>{const [dl]=await Promise.all([p.waitForEvent('download'),p.click('#lite [data-tp="csv"]')]);
  return fs.readFileSync(await dl.path(),'utf8').replace(/^﻿/,'').split('\r\n')[0]};

 /* 1 · ברירת המחדל = בדיוק מה שהיה (הסדר בקוד, בלי hide:1) */
 await tab('short');
 const D=await p.evaluate(()=>{const exp=TP_HEAD.concat(TOPICS.short.cols).filter(c=>(ETA||!c.eta)&&!c.hide&&!(c.k==='planner'&&[...new Set(QF.all.map(r=>r.planner).filter(Boolean))].length<=1)).map(c=>c.k);
  return {exp,got:[...document.querySelectorAll('#lite thead th')].map(t=>t.dataset.tpSort)}});
 ok('ברירת המחדל זהה לסדר שבקוד',JSON.stringify(D.exp)===JSON.stringify(D.got),D.got.join(','));
 const csv0=await csvHead();

 /* 2 · חלון «עמודות»: הסתרה, ↑↓, הצגת עמודה מוסתרת (להזמין) */
 await p.click('#lite [data-tp="cols"]');
 ok('הכפתור «עמודות» פותח את החלון, ו«מק״ט» אינו ניתן להסרה',await p.evaluate(()=>!!document.querySelector('#lite .tpcols')&&document.querySelector('#lite [data-tpc-vis="pn"]').disabled));
 await p.click('#lite [data-tpc-vis="rate"]');
 let H=await heads();ok('הסתרת «קצב לחודש»',!H.includes('rate'),H.join(','));
 await p.click('#lite [data-tpc-vis="buy"]');
 H=await heads();ok('הצגת עמודה שמוסתרת כברירת מחדל («להזמין»)',H.includes('buy'),H.join(','));
 const before=await heads(),iFree=before.indexOf('free');
 await p.click('#lite [data-tp="coldn"][data-k="free"]');
 H=await heads();ok('↓ מזיז את «מלאי פנוי» אחרי העמודה המוצגת הבאה (מדלג על מוסתרת)',H.indexOf('free')===iFree+1&&H[iFree]===before[iFree+1],H.join(','));
 await p.click('#lite [data-tp="colup"][data-k="free"]');
 H=await heads();ok('↑ מחזיר אותו',H.indexOf('free')===iFree,H.join(','));
 ok('↑ על העמודה שאחרי «מק״ט» — כבוי («מק״ט» תמיד ראשון)',await p.evaluate(()=>{const st=tpColsState('short');return document.querySelector(`#lite [data-tp="colup"][data-k="${st[1].c.k}"]`).disabled}));

 /* 3 · גרירה */
 const g0=await heads(),from=g0[g0.length-1],to=g0[2];
 await p.dragAndDrop(`#lite th[data-tp-col="${from}"]`,`#lite th[data-tp-col="${to}"]`);
 H=await heads();ok('גרירת הכותרת האחרונה אל המקום השלישי',H[2]===from&&H[0]==='pn',`${from} → ${H.slice(0,4).join(',')}`);
 ok('לחיצה על כותרת עדיין ממיינת',await p.evaluate(()=>{document.querySelector('#lite th[data-tp-sort="free"]').click();return (TP.sort.short||{}).k==='free'}));

 /* 4 · שמירה: לכל נושא בנפרד, ושורדת טעינה מחדש */
 const sh=await heads();
 await tab('cust');const c0=await heads();
 ok('נושא אחר לא הושפע',!c0.includes('buy')&&c0.includes('free'),c0.join(','));
 await load();await tab('short');
 H=await heads();ok('אחרי טעינה מחדש — אותו סידור במלאי 0',JSON.stringify(H)===JSON.stringify(sh),H.join(','));
 ok('ה-CSV לא השתנה מהסידור',await csvHead()===csv0);

 /* 5 · כללי ההקשר: עמודת «מצב» יורדת בתת-קטגוריה גם אחרי הזזה */
 const cat=await p.evaluate(()=>{TP.cat.short='zero';tpRender();return [...document.querySelectorAll('#lite thead th')].map(t=>t.dataset.tpSort)});
 ok('תת-קטגוריה: «מצב» מוסתר גם בסידור אישי',!cat.includes('kind'),cat.join(','));

 /* 6 · עמודה חדשה שאינה בסידור השמור — נכנסת ליד השכנה שלה */
 const N=await p.evaluate(()=>{const all=JSON.parse(localStorage.getItem('planner_tpcols_v1'));
  all.short.order=all.short.order.filter(k=>k!=='zd');localStorage.setItem('planner_tpcols_v1',JSON.stringify(all));
  const st=tpColsState('short').map(x=>x.c.k),def=TP_HEAD.concat(TOPICS.short.cols).map(c=>c.k),prev=def[def.indexOf('zd')-1];
  return {i:st.indexOf('zd'),prev:st.indexOf(prev)}});
 ok('עמודה שלא הייתה בסידור השמור — מיד אחרי שכנתה',N.i===N.prev+1,JSON.stringify(N));

 /* 7 · חזרה לברירת המחדל, ושמירה חסומה אינה שוברת */
 await p.evaluate(()=>{TP.cat={};TP.colsOpen=true;tpRender()});
 await p.click('#lite [data-tp="colreset"]');
 H=await heads();ok('«חזרה לברירת המחדל»',JSON.stringify(H)===JSON.stringify(D.got),H.join(','));
 const blocked=await p.evaluate(()=>{const o=Storage.prototype.setItem,g=Storage.prototype.getItem;
  Storage.prototype.setItem=()=>{throw new Error('blocked')};Storage.prototype.getItem=()=>{throw new Error('blocked')};
  let okk=true;try{tpColsVis('short','rate',false);tpRender()}catch(e){okk=false}
  Storage.prototype.setItem=o;Storage.prototype.getItem=g;return okk&&!!document.querySelector('#lite table')});
 ok('אחסון חסום — הטבלה ממשיכה לעבוד בברירת המחדל',blocked);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await p.evaluate(()=>{try{localStorage.removeItem('planner_tpcols_v1')}catch(_){}});
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0);
})().catch(e=>{console.error(e);process.exit(1)});
