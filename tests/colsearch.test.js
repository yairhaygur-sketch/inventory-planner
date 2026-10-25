/* ============ D2C · חיפוש בכל ערכי סינון העמודה ============
   עד D2C בורר הסינון של עמודה הראה רק את 60 הערכים הנפוצים, וערך נדיר
   לא היה בו. ובפתיחה חוזרת הרשימה נבנתה מהשורות שכבר עברו את הסינון של
   אותה עמודה — כלומר הופיעו רק הערכים שכבר נבחרו. הבדיקה מוודאה:
     · כל הערכים מופיעים (מהשורות שעוברות את שאר הסינונים), עם חיפוש בתוכם
     · פתיחה וחיפוש מהירים גם בעמודה עם אלפי ערכים; גלילה טוענת עוד
     · הבחירה נשמרת בזמן חיפוש, ו«החל» מחיל בדיוק אותה
     · בפתיחה חוזרת — כל הערכים, והנבחרים מוצמדים למעלה ומסומנים
     · «סמן את כל התואמים», «נקה סימון», «נקה סינון עמודה»
     · סינון בעמודה אחרת מצמצם את הרשימה; ערך נבחר בלי שורות נשאר (0)
     · ערך עם תווים מיוחדים (&) מוצג ומסונן נכון
     · שום דבר מזה לא משנה נתונים, המלצות, בחירה ל-SAP; הייצוא = הסינון
     · טלפון: בטבלת התנועות (הטבלה היחידה עם כותרת בטלפון) הבורר בתוך
       המסך ועובד
   COLSEARCH_XLSX=<דוח> — על דוח אחר (למשל האמיתי, מקומית). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const XL=process.env.COLSEARCH_XLSX||path.join(SD,'zmrp-demo.xlsx');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
const N=n=>Number(n).toLocaleString('he-IL');
async function open(b,W,H){const ctx=await b.newContext({viewport:{width:W,height:H}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0,null,{timeout:180000});
 await p.waitForTimeout(600);
 await p.evaluate(()=>{window.__raf=()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
  window.__open=(ck,tbl)=>{const th=[...document.querySelectorAll(`#${tbl||'tbl'} thead th`)].find(t=>t.dataset.ck===ck);th.querySelector('.wfil').click()};
  window.__q=v=>{const q=document.getElementById('cpVq');q.value=v;q.dispatchEvent(new Event('input'))};
  window.__pick=v=>{const x=document.querySelector(`#cpVals input[value="${CSS.escape(v)}"]`);if(x)x.click();return !!x};
  window.__cap=[];XLSX.writeFile=(wb,name)=>{const sheets={};for(const sn of wb.SheetNames)
   sheets[sn]=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1}).length-1;window.__cap.push({name,sheets})}});
 return {ctx,p,errs}}
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 /* ---------------- מחשב ---------------- */
 {const {ctx,p,errs}=await open(b,1440,900);
  const R=await p.evaluate(async()=>{const o={},N=x=>x.toLocaleString('he-IL'),T=id=>(document.getElementById(id)||{}).textContent||'';
   const F=['cat','sugSS','sugROP','paramFix','tierFloor','sev','expCap','miss'];
   const snap=()=>JSON.stringify(ALL.map(r=>F.map(f=>r[f])));
   setMode('month');await __raf();for(const r of sapFiltered().slice(0,5))SAPSEL.add(r.pn);sapBarSync();
   const sel0=JSON.stringify([...SAPSEL].sort()),data0=snap();
   setMode('catalog');await __raf();
   /* 1 — כל הערכים, בלי מגבלת 60 */
   const uniq=new Set(currentRows().map(r=>String(colAcc('c:cPn')(r))));
   const t0=performance.now();__open('cPn');o.openMs=Math.round(performance.now()-t0);
   const box=document.getElementById('cpVals');o.uniq=uniq.size;o.meta=T('cpVn');o.drawn=box.querySelectorAll('label').length;
   /* גלילה טוענת עוד, עד הכול */
   let guard=0;while(box.querySelectorAll('label').length<uniq.size&&guard++<200){box.scrollTop=box.scrollHeight;box.dispatchEvent(new Event('scroll'))}
   o.all=box.querySelectorAll('label').length;
   /* 2 — ערך נדיר מסוף הרשימה, וחיפוש מהיר */
   const vals=[...uniq],rare=vals[vals.length-1],other=vals[0];
   const t1=performance.now();__q(rare);o.searchMs=Math.round(performance.now()-t1);
   o.found=__pick(rare);
   /* 3 — בחירה נשמרת בזמן חיפוש אחר */
   __q(other);__pick(other);o.meta2=T('cpVn');__q('');
   document.getElementById('cpOk').click();await __raf();
   o.rows=VL.tbl.items.length;o.expRows=currentRows().length;
   o.rowsOk=currentRows().every(r=>[rare,other].includes(String(colAcc('c:cPn')(r))));
   o.chip=[...document.querySelectorAll('#chipsTop .chip')].some(c=>/מק״ט/.test(c.textContent));
   /* 4 — פתיחה חוזרת: כל הערכים, הנבחרים למעלה ומסומנים */
   __open('cPn');const lab=[...document.querySelectorAll('#cpVals label input')];
   o.reMeta=T('cpVn');o.pinned=lab.slice(0,2).every(x=>x.checked)&&lab.slice(0,2).map(x=>x.value).sort().join()===[rare,other].sort().join();
   o.reAll=new RegExp(`^${N(uniq.size)} ערכים · 2 נבחרו`).test(T('cpVn'));
   /* 5 — נקה סימון + החל = הסינון יורד */
   document.getElementById('cpVnone').click();document.getElementById('cpOk').click();await __raf();
   o.unset=!colFilters['c:cPn']&&VL.tbl.items.length===QF.all.length;
   /* 6 — «סמן את כל התואמים» בחיפוש */
   const models=[...new Set(currentRows().map(r=>String(colAcc('c:model')(r))))];
   const sub=(models.find(m=>m.length>3)||models[0]).slice(0,3).toLowerCase();
   __open('model');__q(sub);const m=[...document.querySelectorAll('#cpVals input')].map(x=>x.value);
   document.getElementById('cpVall').click();document.getElementById('cpOk').click();await __raf();
   const want=QF.all.filter(r=>String(colAcc('c:model')(r)).toLowerCase().includes(sub)).length;
   o.allMatch={n:m.length,f:colFilters['c:model']?colFilters['c:model'].values.size:0,rows:VL.tbl.items.length,want,
    only:colFilters['c:model']&&[...colFilters['c:model'].values].every(v=>v.toLowerCase().includes(sub))};
   /* 7 — סינון בעמודה אחרת מצמצם את רשימת הערכים */
   __open('cPn');o.cascade=document.querySelectorAll('#cpVals label').length===Math.min(200,VL.tbl.items.length)
    &&new RegExp(`^${N(VL.tbl.items.length)} ערכים`).test(T('cpVn'));
   document.getElementById('cpVals').querySelector('input').click();const kept=document.getElementById('cpVals').querySelector('input').value;
   document.getElementById('cpOk').click();await __raf();
   /* 8 — ערך נבחר שאין לו שורות אחרי שינוי מסנן אחר (כאן החיפוש הכללי) נשאר, עם 0.
      בעמודות עצמן אי אפשר להגיע לזה: רשימת עמודה אחרת כבר מצומצמת לפי הבחירה הזו. */
   const otherPn=QF.all.map(r=>String(colAcc('c:cPn')(r))).find(v=>v!==kept&&!v.includes(kept)&&!kept.includes(v));
   setQuery(otherPn);await __raf();
   __open('cPn');{const first=document.querySelector('#cpVals label');
    o.zero=!!first&&first.querySelector('input').value===kept&&first.querySelector('input').checked&&first.querySelector('.c').textContent==='0'}
   document.getElementById('cpClear').click();await __raf();clearQuery();await __raf();
   __open('model');document.getElementById('cpClear').click();await __raf();
   /* 9 — ערך עם & (מוצג ומסונן נכון) */
   const amp=[...new Set(QF.all.map(r=>String(colAcc('c:model')(r))))].find(v=>/[&<>"]/.test(v));
   if(amp){__open('model');__q(amp);const lb=[...document.querySelectorAll('#cpVals label')].find(l=>l.querySelector('input').value===amp);
    o.ampShow=!!lb&&lb.querySelector('span').textContent===amp;lb.querySelector('input').click();document.getElementById('cpOk').click();await __raf();
    o.ampRows=VL.tbl.items.length===QF.all.filter(r=>String(colAcc('c:model')(r))===amp).length;
    __open('model');document.getElementById('cpClear').click();await __raf()}
   else{o.ampShow=o.ampRows=true;o.noAmp=true}
   /* 10 — שום דבר לא זז: נתונים, המלצות, בחירה ל-SAP; והייצוא */
   o.dataSame=snap()===data0;o.selSame=JSON.stringify([...SAPSEL].sort())===sel0;
   __open('model');__q(sub);document.getElementById('cpVall').click();document.getElementById('cpOk').click();await __raf();
   __cap.length=0;doExport();const ex=__cap[0]||{sheets:{}};
   /* כמו ב-main: שני הגיליונות הם הסינון הנוכחי, והקובץ נקרא «filtered» */
   o.exView=ex.sheets['התצוגה']===want&&currentRows().length===want;
   o.exAll=Object.entries(ex.sheets).some(([k,v])=>/^כל הנתונים/.test(k)&&v===want)&&/_filtered_/.test(ex.name||'');
   __open('model');document.getElementById('cpClear').click();await __raf();
   o.dataSame2=snap()===data0;o.filtersLeft=Object.keys(colFilters).length;
   SAPSEL.clear();sapBarSync();return o});
  ok(`מחשב · כל ${N(R.uniq)} הערכים בעמודה — לא רק 60 הנפוצים`,R.all===R.uniq&&R.uniq>60?true:R.all===R.uniq,`${R.meta} · נטענו ${N(R.all)}`);
  ok('מחשב · מצוירים 200 בכל פעם, וגלילה טוענת עוד',R.drawn===Math.min(200,R.uniq)&&R.all===R.uniq,`${R.drawn} → ${N(R.all)}`);
  ok('מחשב · פתיחה וחיפוש מהירים גם באלפי ערכים',R.openMs<400&&R.searchMs<100,`פתיחה ${R.openMs}ms · חיפוש ${R.searchMs}ms`);
  ok('מחשב · ערך נדיר מסוף הרשימה נמצא בחיפוש',R.found);
  ok('מחשב · הבחירה נשמרת בזמן חיפוש אחר, ו«החל» מחיל את שתיהן',/2 נבחרו/.test(R.meta2)&&R.rows===R.expRows&&R.rows>=2&&R.rowsOk,`${R.rows} שורות`);
  ok('מחשב · הסינון מסומן בשבב',R.chip);
  ok('מחשב · פתיחה חוזרת: כל הערכים, והנבחרים מוצמדים למעלה ומסומנים',R.pinned&&R.reAll,R.reMeta);
  ok('מחשב · «נקה סימון» ו«החל» — הסינון יורד',R.unset);
  ok('מחשב · «סמן את כל התואמים» בוחר בדיוק את התואמים לחיפוש',R.allMatch.f===R.allMatch.n&&R.allMatch.only&&R.allMatch.rows===R.allMatch.want,JSON.stringify(R.allMatch));
  ok('מחשב · סינון בעמודה אחרת מצמצם את רשימת הערכים',R.cascade);
  ok('מחשב · ערך נבחר שאין לו כרגע שורות נשאר ברשימה (0), כדי שאפשר יהיה להסירו',R.zero);
  ok('מחשב · ערך עם תווים מיוחדים (&) מוצג ומסונן נכון',R.ampShow&&R.ampRows,R.noAmp?'אין ערך כזה בדוח':'');
  ok('מחשב · הנתונים, ההמלצות והבחירה ל-SAP לא זזו',R.dataSame&&R.dataSame2&&R.selSame);
  ok('מחשב · הייצוא מכיל בדיוק את המסונן (כמו ב-main), ושם הקובץ «filtered»',R.exView&&R.exAll);
  ok('מחשב · «נקה סינון עמודה» מסיר את הסינון',R.filtersLeft===0,R.filtersLeft);
  ok('מחשב · אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
  await ctx.close()}
 /* ---------------- טלפון ---------------- */
 {const {ctx,p,errs}=await open(b,390,844);
  const R=await p.evaluate(async()=>{const o={};setMode('moves');await __raf();
   const th=[...document.querySelectorAll('#movtbl thead th')].find(t=>t.querySelector('.wfil')&&t.dataset.ck&&!/date|age|delta/i.test(t.dataset.ck));
   if(!th)return {none:true};const ck=th.dataset.ck;th.querySelector('.wfil').click();
   const pop=document.getElementById('colpop'),r=pop.getBoundingClientRect();
   o.inside=r.left>=0&&r.right<=innerWidth+1;o.ck=ck;
   const box=document.getElementById('cpVals');if(!box)return {...o,numeric:true};
   const v=box.querySelector('input').value;__q(v);__pick(v);document.getElementById('cpOk').click();await __raf();
   const acc=colAcc('c:'+ck);o.rows=VL.movtbl.items.length;o.want=movRows().filter(x=>String(acc(x))===v).length;
   o.match=movRows().every(x=>String(acc(x))===v);
   const th2=[...document.querySelectorAll('#movtbl thead th')].find(t=>t.dataset.ck===ck);th2.querySelector('.wfil').click();
   document.getElementById('cpClear').click();await __raf();o.cleared=!colFilters['c:'+ck];return o});
  if(R.none)ok('טלפון · אין בטבלת התנועות עמודה עם בורר',true);
  else{ok('טלפון · בורר הסינון בתוך המסך',R.inside,R.ck);
   if(!R.numeric)ok('טלפון · חיפוש ובחירת ערך מסננים את הטבלה',R.match&&R.rows===R.want&&R.rows>0,`${R.rows} שורות`);
   ok('טלפון · «נקה סינון עמודה» מסיר',R.cleared)}
  ok('טלפון · אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
  await ctx.close()}
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
