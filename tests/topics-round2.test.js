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

 ok('אפס שגיאות JS',!errs.length,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(`\n${pass} עברו · ${fail} נכשלו`);process.exit(fail?1:0)})();
