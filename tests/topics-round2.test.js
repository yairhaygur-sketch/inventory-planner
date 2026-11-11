/* ============ החלטות המעתד 2.11 — תצוגת 4 הנושאים, סבב 2 ============
   1 · חריגת צריכה: הסכום בדוח מוצג ליד הסכום המנוקה, וכרטיס הפריט אומר איזה
       חודש תוקן ולמה — מאותה cleanSeries של המנוע.
   2 · «ללא שינוי פרופיל»: נטרול משותף לשני הכלים, נשמר בנפרד מהסימונים, בגיבוי,
       וקבוע עד ביטול ידני. הפריט נשאר בשאר הנושאים.
   3 · PD ← ZM-C: נקודת הזמנה מינימלית של יח׳ אחת בהמלצה; כלי עבודה/אביזרים
       נשארים לבדיקה ידנית, בלי מספר. */
const {chromium}=require('playwright');const fs=require('fs');const path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx'),HTML='file://'+path.join(SD,'..','index.html');
let pass=0,fail=0;const ok=(n,c,d)=>{c?pass++:fail++;console.log(`${c?'PASS':'FAIL'} · ${n}${d?'  ['+d+']':''}`)};
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 /* הקשר אחד לשני הכלים — אותו localStorage, כמו שתי לשוניות של המתכנן */
 const ctx=await b.newContext({viewport:{width:1440,height:900},acceptDownloads:true});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const errs=[];
 const open=async q=>{const p=await ctx.newPage();p.on('pageerror',e=>errs.push(e.message));await p.goto(HTML+q);await p.waitForTimeout(500);
  await p.setInputFiles('#f',XL);await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB,null,{timeout:120000});
  await p.waitForTimeout(300);return p};

 /* ---- 1. חריגת צריכה: בדוח מול מנוקה ---- */
 const T=await open('?topics');await T.waitForSelector('#lite table');
 await T.click('[data-tp-tab="trend"]');
 const tr=await T.evaluate(()=>{const rows=TOPICS.trend.rows();
  const raw=(r,a)=>r.months.slice(a,a+3).reduce((x,y)=>x+y,0);
  const diff=rows.filter(r=>raw(r,0)!==r.tRec||raw(r,3)!==r.tPre);
  /* אותה סדרה שהמנוע ניקה — הסכום שלה הוא tRec/tPre בדיוק */
  const same=rows.every(r=>{const c=cleanSeries(r.months,r.A);return c.slice(0,3).reduce((x,y)=>x+y,0)===r.tRec&&c.slice(3,6).reduce((x,y)=>x+y,0)===r.tPre});
  const shown=TP.shown.length,cells=[...document.querySelectorAll('#lite tr.r')].filter(tr=>/בדוח/.test(tr.textContent)).length;
  const exp=TP.shown.filter(r=>raw(r,0)!==r.tRec||raw(r,3)!==r.tPre).length;
  const ex=diff[0],adj=ex?tpTrendAdj(ex):[];
  return {n:rows.length,diff:diff.length,same,cells,exp,shown,pn:ex&&ex.pn,adj,
   adjOk:!!ex&&adj.length>0&&adj.every(a=>a.raw===ex.months[a.i-1]&&a.clean===cleanSeries(ex.months,ex.A)[a.i-1]),
   heads:[...document.querySelectorAll('#lite th')].map(t=>t.textContent.trim()).filter(t=>/ח׳/.test(t))}});
 ok('המגמה = סכום הסדרה המנוקה (cleanSeries) — המנוע והתצוגה קוראים לאותה פונקציה',tr.same);
 ok(`«בדוח N» מופיע בדיוק בשורות שבהן הסכום בדוח שונה (${tr.exp} מתוך ${tr.shown})`,tr.exp>0&&tr.cells===tr.exp,`${tr.cells} · ${tr.exp}`);
 ok('הכותרות אומרות «מנוקה»',tr.heads.length===2&&tr.heads.every(h=>/מנוקה/.test(h)),tr.heads.join(' · '));
 ok('tpTrendAdj: כל חודש שתוקן — הערך בדוח והערך המנוקה, בדיוק כמו במנוע',tr.adjOk,JSON.stringify(tr.adj).slice(0,200));
 {await T.evaluate(pn=>{TP.q=pn},tr.pn);await T.evaluate(()=>tpRender());await T.click('#lite tr.r');
  const d=await T.evaluate(()=>document.getElementById('tpd').textContent);
  ok('כרטיס הפריט: «בדוח» + אילו חודשים תוקנו ולמה',/בדוח: /.test(d)&&/תוקנו:/.test(d)&&/(קפיצה|אזל|החזרה)/.test(d),d.match(/תוקנו:[^.]{0,90}/)?.[0]);
  await T.evaluate(()=>{tpdClose();TP.q='';tpRender()})}
 {const [dl]=await Promise.all([T.waitForEvent('download'),T.click('#lite [data-tp="csv"]')]);
  const lines=fs.readFileSync(await dl.path(),'utf8').replace(/^﻿/,'').split('\r\n'),H=lines[0].split(',').map(x=>x.replace(/"/g,''));
  ok('CSV: גם הסכום בדוח וגם המנוקה, לשני החלונות',['3 ח׳ אחרונים — מנוקה','3 ח׳ אחרונים — בדוח','3 ח׳ קודמים — מנוקה','3 ח׳ קודמים — בדוח'].every(h=>H.includes(h)),H.slice(0,12).join(','))}

 /* ---- 2. «ללא שינוי פרופיל» ---- */
 await T.click('[data-tp-tab="fix"]');
 const pick=await T.evaluate(()=>{const r=QF.all.find(x=>pdUpgrade(x)&&tpLtPrio(x)===0)||QF.all.find(pdUpgrade);
  return {pn:r.pn,key:markKey(r),inShort:TOPICS.short.rows().includes(r),inCust:TOPICS.cust.rows().includes(r),n0:profRows().filter(profActive).length}});
 await T.evaluate(pn=>{TP.cat.fix='prof';TP.q=pn;tpRender()},pick.pn);
 await T.click('#lite tr.r');await T.waitForSelector('#tpd [data-profoff="1"]');
 await T.click('#tpd [data-profoff="1"]');await T.waitForTimeout(400);
 const a1=await T.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);
  return {off:profOff(r),up:pdUpgrade(r),info:profInfo(r),inProf:TOPICS.fix.rows().filter(TOPICS.fix.cats[0][3]).includes(r),
   inShort:TOPICS.short.rows().includes(r),inCust:TOPICS.cust.rows().includes(r),n1:profRows().filter(profActive).length,
   undo:!!document.querySelector('#tpd [data-profoff="0"]'),stored:JSON.parse(localStorage.getItem('planner_profoff_v1')||'{}')[k],
   why:(r.why||[]).map(w=>w[1]).join(' | '),backup:Object.keys(stateBlob().keys).includes('planner_profoff_v1')}},pick.key);
 ok('נטרול מהכרטיס: אין המלצת פרופיל (pdUpgrade/profInfo), והפריט ירד מ«שינוי פרופיל MRP»',a1.off&&!a1.up&&a1.info===null&&!a1.inProf&&a1.n1===pick.n0-1,JSON.stringify({off:a1.off,up:a1.up,n:[pick.n0,a1.n1]}));
 ok('הפריט נשאר בשאר הנושאים',a1.inShort===pick.inShort&&a1.inCust===pick.inCust,JSON.stringify({before:[pick.inShort,pick.inCust],after:[a1.inShort,a1.inCust]}));
 ok('הנימוק במנוע אינו ממליץ עוד על ZM-C',!/ל-ZM-C בעקבות צריכה/.test(a1.why),a1.why.slice(0,120));
 ok('נשמר במפתח נפרד (planner_profoff_v1), נכלל בגיבוי, והכרטיס מציע ביטול',!!a1.stored&&a1.stored.ts>0&&a1.backup&&a1.undo);
 await T.evaluate(()=>{TP.q='';TP.cat.fix='';tpdClose();tpRender()});
 const lst=await T.evaluate(()=>{const d=document.querySelector('#lite details.off');return d?{t:d.querySelector('summary').textContent,n:d.querySelectorAll('li').length}:null});
 ok('«טיוב נתונים»: רשימת המנוטרלים עם ביטול',!!lst&&lst.n===1&&/ללא שינוי פרופיל/.test(lst.t),JSON.stringify(lst));
 /* הכלי המלא — אותו נטרול, ביטול מכרטיס הפריט שם */
 const F=await open('?nobrief=1');
 const f1=await F.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);setMode('prof');
  return {off:profOff(r),inProf:profRows().includes(r),sum:document.body.textContent.includes('ללא שינוי פרופיל (נוטרלו ידנית')}},pick.key);
 ok('הכלי המלא רואה את הנטרול: הפריט אינו במסך שינוי הפרופיל, והמונה אומר כמה נוטרלו',f1.off&&!f1.inProf&&f1.sum,JSON.stringify(f1));
 await F.evaluate(k=>detail(ALL.find(x=>markKey(x)===k)),pick.key);await F.waitForSelector('#detail [data-profoff="0"]');
 await F.click('#detail [data-profoff="0"]');await F.waitForTimeout(400);
 const f2=await F.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);return {off:profOff(r),up:pdUpgrade(r),btn:!!document.querySelector('#detail [data-profoff="1"]')}},pick.key);
 ok('ביטול בכרטיס הפריט בכלי המלא מחזיר את ההמלצה',!f2.off&&f2.up&&f2.btn,JSON.stringify(f2));
 const t2=await T.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);return {off:profOff(r)}},pick.key);
 ok('והלשונית של התצוגה המצומצמת רואה את הביטול (סנכרון בין לשוניות)',!t2.off);
 /* קבוע עד ביטול ידני (BUSINESS-DECISIONS 2.11) — גם אחרי שנה */
 const exp=await T.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);PROFOFF[k]={ts:Date.now()-400*864e5,pn:r.pn};
  const a={off:profOff(r),up:pdUpgrade(r)};delete PROFOFF[k];return a},pick.key);
 ok('נטרול אינו פג — גם אחרי יותר משנה, עד ביטול ידני',exp.off&&!exp.up);
 /* הורדת ZM (לא PD) — אותו כפתור */
 const zm=await T.evaluate(()=>{const r=QF.all.find(x=>profActive(x)&&!pdUpgrade(x));if(!r)return null;
  setProfOff(r,true);const r2=ALL.find(x=>x.pn===r.pn);const a={pn:r.pn,off:profOff(r2),act:profActive(r2)};setProfOff(r2,false);
  a.back=profActive(ALL.find(x=>x.pn===r.pn));return a});
 ok('הורדת פרופיל ZM — אותו נטרול ואותו ביטול',!!zm&&zm.off&&!zm.act&&zm.back,JSON.stringify(zm));

 /* ---- 3. נקודת הזמנה מינימלית ב-PD ← ZM-C ---- */
 const rm=await T.evaluate(()=>{const up=QF.all.filter(pdUpgrade).map(r=>[r,tpRec(r)]);
  const nums=up.filter(([,q])=>!q.na),min=up.filter(([,q])=>q.ropMin),od=up.filter(([r])=>OD_GROUPS.has(r.mkt1));
  return {n:up.length,nums:nums.length,low:nums.filter(([,q])=>q.rop<PD_ROP_MIN).map(([r])=>r.pn),
   min:min.map(([r,q])=>({pn:r.pn,rop:q.rop,calc:q.ropCalc,ss:q.ss})),odNa:od.every(([,q])=>q.na),od:od.length,
   simCol:min.length?profSimVal(min[0][0],'rop'):null,pn:min.length?min[0][0].pn:null}});
 ok(`כל המלצת PD ← ZM-C עם מספר — נקודת הזמנה ${'≥'} 1 (${rm.nums} מתוך ${rm.n})`,rm.nums>0&&!rm.low.length,rm.low.join(','));
 ok('פריט שהחישוב נתן לו 0 מקבל 1, והחישוב המקורי נשמר (ropCalc)',rm.min.length>0&&rm.min.every(x=>x.rop===1&&x.calc<1),JSON.stringify(rm.min));
 ok('כלי עבודה ואביזרים — נשארים בלי מספר (לבדיקה ידנית)',rm.od>0&&rm.odNa,`${rm.od} פריטים`);
 ok('מסך הפרופילים בכלי המלא מציג את אותו ROP (profSimVal)',rm.simCol===1,String(rm.simCol));
 {await T.evaluate(pn=>{TP.cat.fix='prof';TP.q=pn;tpRender()},rm.pn);await T.click('#lite tr.r');
  const d=await T.evaluate(()=>document.getElementById('tpd').textContent);
  ok('כרטיס הפריט אומר שזו נקודת הזמנה מינימלית, ומה החישוב נתן',/נקודת הזמנה מינימלית/.test(d)&&/החישוב אחרי ZM-C נתן 0/.test(d),d.match(/נקודת הזמנה מינימלית[^.]{0,80}/)?.[0])}

 /* ---- 4. «נתונים שבורים» — גם רמת שירות חסרה (2.11) ----
    בדוח הסינתטי אין פריט כזה, ולכן עותק שבו «רמת שרות» = 0 לשלושה פריטים:
    מנוהל עם צריכה (נכנס), לקוח ממתין שנעצר ברונג מוקדם במנוע (נכנס), PD (לא נכנס) */
 {const pk=await T.evaluate(()=>({mgd:QF.all.find(r=>!r.isOD&&!r.noStock&&r.rate>0&&!r.cust&&!r.ltBad&&!r.srvMissing).pn,
   burn:QF.all.find(r=>r.burn&&!r.isOD&&r.rate>0).pn,pd:QF.all.find(r=>r.isPDItem&&r.rate>0).pn}));
  const XLSX=require('xlsx'),os=require('os');const wb=XLSX.readFile(XL),ws=wb.Sheets[wb.SheetNames[0]];
  const aoa=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:''});const hi=aoa.findIndex(r=>r.includes('מק"ט מוביל'));
  const cp=aoa[hi].indexOf('מק"ט מוביל'),cs=aoa[hi].indexOf('רמת שרות');
  for(const r of aoa.slice(hi+1))if([pk.mgd,pk.burn,pk.pd].includes(String(r[cp]).trim()))r[cs]=0;
  const w2=XLSX.utils.book_new();XLSX.utils.book_append_sheet(w2,XLSX.utils.aoa_to_sheet(aoa),wb.SheetNames[0]);
  const f2=path.join(fs.mkdtempSync(path.join(os.tmpdir(),'srv-')),'srv.xlsx');XLSX.writeFile(w2,f2);
  const S=await ctx.newPage();S.on('pageerror',e=>errs.push(e.message));await S.goto(HTML+'?topics');await S.waitForTimeout(400);
  await S.setInputFiles('#f',f2);await S.waitForSelector('#lite table',{timeout:120000});
  const r=await S.evaluate(pk=>{const data=TOPICS.fix.rows().filter(TOPICS.fix.cats[2][3]).map(r=>r.pn);const by=pn=>ALL.find(x=>x.pn===pn);
   return {data,srv:[pk.mgd,pk.burn,pk.pd].map(pn=>by(pn).srvMissing),engBurn:!!by(pk.burn).serviceMissing,n:QF.all.filter(tpSrvMissing).length}},pk);
  ok('רמת שירות 0 → srvMissing בשלושת הפריטים (שדה המנוע)',r.srv.every(Boolean),JSON.stringify(r.srv));
  ok('«נתונים שבורים»: פריט מנוהל עם צריכה ורמת שירות חסרה — נכנס',r.data.includes(pk.mgd),r.data.join(','));
  ok('גם פריט עם לקוח ממתין, שהמנוע עצר ברונג מוקדם (בלי serviceMissing) — נכנס',r.data.includes(pk.burn)&&!r.engBurn);
  ok('PD — אין לו המלצת SS/ROP, ולכן אינו נכנס',!r.data.includes(pk.pd));
  await S.evaluate(pn=>{TP.tab='fix';TP.cat.fix='data';TP.q=pn;tpRender()},pk.mgd);await S.click('#lite tr.r');
  const d=await S.evaluate(()=>document.getElementById('tpd').textContent);
  ok('כרטיס הפריט: «רמת שירות חסרה» והפעולה ב-SAP',/רמת שירות חסרה או 0 ב-SAP/.test(d)&&/להזין רמת שירות מאושרת ב-SAP/.test(d));
  fs.rmSync(path.dirname(f2),{recursive:true,force:true})}

 /* ---- 5. ליד כל מק״ט: העתקה, «טופל», והדגם אחרי התיאור (11.11) ---- */
 {await ctx.grantPermissions(['clipboard-read','clipboard-write']).catch(()=>{});
  const P=await open('?topics');await P.waitForSelector('#lite table');await P.click('[data-tp-tab="cust"]');
  const row=await P.evaluate(()=>{const tr=document.querySelector('#lite tr.r'),r=TP.shown[0];
   return {pn:r.pn,model:r.model,key:markKey(r),n0:tpRows('cust',true).length,hasBtns:!!tr.querySelector('[data-tp="copy"]')&&!!tr.querySelector('[data-tp="done"]'),
    descCell:tr.children[1].textContent,heads:[...document.querySelectorAll('#lite th')].map(t=>t.textContent.trim())}});
  ok('בכל שורה, ליד המק״ט: כפתור העתקה וכפתור «טופל»',row.hasBtns&&await P.evaluate(()=>[...document.querySelectorAll('#lite tr.r')].every(tr=>tr.querySelector('[data-tp="copy"]')&&tr.querySelector('[data-tp="done"]'))));
  ok('הדגם מופיע אחרי התיאור',row.descCell.endsWith(row.model)&&row.heads.includes('תיאור · דגם'),row.descCell);
  await P.click('#lite tr.r [data-tp="copy"]');await P.waitForTimeout(250);
  const cp=await P.evaluate(async()=>({drawer:!document.getElementById('tpd').hidden,clip:await navigator.clipboard.readText().catch(()=>null)}));
  ok('העתקה: המק״ט בלוח, והכרטיס לא נפתח',!cp.drawer&&(cp.clip===null||cp.clip===row.pn),JSON.stringify(cp));
  await P.click('#lite tr.r [data-tp="done"]');await P.waitForTimeout(400);
  const d=await P.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);
   return {mark:MARKS[k]&&MARKS[k].t,n1:tpRows('cust',true).length,inAny:Object.keys(TOPICS).some(t=>tpRows(t,true).includes(r)),
    list:!!document.querySelector('#lite details.off [data-tp="undone"]')}},row.key);
  ok('«טופל»: סימון handled (של הכלי המלא), הפריט ירד מכל הנושאים, ומופיע ברשימת «טופלו»',d.mark==='handled'&&d.n1===row.n0-1&&!d.inAny&&d.list,JSON.stringify(d));
  const F2=await open('?nobrief=1');
  const fd=await F2.evaluate(pn=>{const r=ALL.find(x=>x.pn===pn);return {t:r.mark&&r.mark.t,cat:r.cat}},row.pn);
  ok('הכלי המלא רואה את אותו סימון «טופל»',fd.t==='handled',JSON.stringify(fd));
  await F2.close();
  await P.click('#lite details.off summary');await P.click('#lite [data-tp="undone"]');await P.waitForTimeout(400);
  const u=await P.evaluate(k=>({mark:MARKS[k]||null,n:tpRows('cust',true).length}),row.key);
  ok('«החזר לרשימה» מבטל את הסימון והפריט חוזר',!u.mark&&u.n===row.n0,JSON.stringify(u));
  const [dl]=await Promise.all([P.waitForEvent('download'),P.click('#lite [data-tp="csv"]')]);
  const H=fs.readFileSync(await dl.path(),'utf8').replace(/^\uFEFF/,'').split('\r\n')[0].split(',').map(x=>x.replace(/"/g,''));
  ok('CSV: «תיאור» ו«דגם» בעמודות נפרדות',H.indexOf('דגם')===H.indexOf('תיאור')+1,H.slice(0,5).join(','));
  await P.close()}

 ok('אפס שגיאות JS',!errs.length,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(`\n${pass} עברו · ${fail} נכשלו`);process.exit(fail?1:0)})();
