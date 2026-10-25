/* ============ L2 · קריאת הדוח ברקע (Web Worker) ============
   XLSX.read רץ ב-Worker שנבנה מהספרייה המשובצת, והדף ממשיך לעבוד בזמן
   הקריאה. נבדק כאן, במחשב ובטלפון:
     · הקריאה עוברת ב-Worker, גם מקובץ מקומי (file://), והדף מגיב בזמן
       הקריאה — מול קריאה רגילה על אותו קובץ
     · מצב טעינה: שם הקובץ, גודל, זמן שעובר, «בטל טעינה»
     · הנתונים זהים בין קריאה ב-Worker לקריאה רגילה (הגיבוי)
     · ביטול: שום דבר לא משתנה — לא הנתונים, לא הכותרת, לא ההדגמה
     · קובץ חדש באמצע טעינה — רק האחרון נכנס
     · קובץ פגום וקובץ בלי כותרת — אותן הודעות כמו קודם, ו«נסה שוב»
     · גיבוי: אין Worker, או Worker שנכשל לפני שהחזיר דבר — קריאה רגילה
     · טעינה שלא הסתיימה (הלשונית נסגרה/קרסה) — הודעה בפתיחה הבאה
   קובץ גדול (20,000 שורות) נבנה כאן מהדוח הסינתטי; אין נתונים אמיתיים.
   WORKER_XLSX=<דוח> — על דוח אחר (למשל האמיתי, מקומית). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path'),os=require('os');
const XLSX=require('xlsx');
const SD=__dirname;
const XL=process.env.WORKER_XLSX||path.join(SD,'zmrp-demo.xlsx');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
const TMP=fs.mkdtempSync(path.join(os.tmpdir(),'wl-'));
/* קובץ גדול: שורות הדוח הסינתטי משוכפלות, ובכל עותק נוסף מק״ט עם סיומת */
function bigFile(n){const wb=XLSX.readFile(path.join(SD,'zmrp-demo.xlsx'),{cellDates:true});const nm=wb.SheetNames[0];
 const a=XLSX.utils.sheet_to_json(wb.Sheets[nm],{header:1,blankrows:false,defval:''});
 const h=a.findIndex(r=>r.some(c=>/מק"ט מוביל|מק״ט מוביל/.test(String(c))));const pn=a[h].findIndex(c=>/מק"ט מוביל|מק״ט מוביל/.test(String(c)));
 const body=a.slice(h+1).filter(r=>String(r[pn]).trim());const rows=[];
 for(let i=0;i<n;i++){const k=Math.floor(i/body.length),r=body[i%body.length].slice();if(k)r[pn]=String(r[pn])+'-k'+k;rows.push(r)}
 const w2=XLSX.utils.book_new();XLSX.utils.book_append_sheet(w2,XLSX.utils.aoa_to_sheet(a.slice(0,h+1).concat(rows),{cellDates:true}),nm);
 const f=path.join(TMP,`big-${n}.xlsx`);XLSX.writeFile(w2,f,{compression:true});return f}
const BIG=bigFile(20000);
const BAD=path.join(TMP,'corrupt.xlsx');fs.writeFileSync(BAD,Buffer.from('PK\x03\x04 not really a workbook '.repeat(40)));
const NOHDR=path.join(TMP,'nohdr.xlsx');
{const w=XLSX.utils.book_new();XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet([['שלום','עולם'],[1,2],[3,4]]),'גיליון1');XLSX.writeFile(w,NOHDR)}
const HTML='file://'+path.join(SD,'..','index.html')+'?nobrief=1';
async function open(b,W,H){const ctx=await b.newContext({viewport:{width:W,height:H}});
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto(HTML);await p.waitForTimeout(500);
 return {ctx,p,errs}}
const loaded=p=>p.waitForFunction(()=>typeof LOADJOB!=='undefined'&&!LOADJOB&&typeof ALL!=='undefined'&&ALL&&ALL.length>0,null,{timeout:300000,polling:100});
const idle=p=>p.waitForFunction(()=>typeof LOADJOB!=='undefined'&&!LOADJOB,null,{timeout:300000,polling:100});
const SNAP=()=>{const F=['pn','cat','sev','sugSS','sugROP','ss','rop','paramFix','tierFloor','expCap','miss','orderQty'];
 return JSON.stringify(ALL.map(r=>F.map(f=>r[f])))+'|'+Object.keys(Q).map(k=>k+':'+Q[k].length).join(',')};
/* מודד את הפער הארוך ביותר בין שני טיקים של הדף, רק בזמן שלב הקריאה */
/* הפער נזקף לשלב שהיה פעיל כשהוא *התחיל*: בקריאה רגילה הטיק שאחרי ההקפאה
   מגיע כשהשלב כבר «מחשב», ובלי זה ההקפאה לא הייתה נמדדת כלל */
const STALL=()=>{window.__st={max:0,n:0,saw:0,box:'',cancel:false,secs:new Set()};let last=performance.now(),wasRead=false;
 window.__sti=setInterval(()=>{const now=performance.now();
  if(wasRead)__st.max=Math.max(__st.max,now-last);
  wasRead=typeof LOADJOB!=='undefined'&&!!LOADJOB&&LOADJOB.stage==='read';
  if(wasRead){__st.n++;
   const bx=document.getElementById('loadbox');if(bx&&!bx.hidden){__st.saw++;__st.box=bx.textContent;__st.cancel=__st.cancel||!!document.getElementById('lbCancel');
    const m=bx.textContent.match(/(\d+) שנ׳/);if(m)__st.secs.add(m[1])}}
  last=now},16)};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 /* ---------------- מחשב ---------------- */
 {const {ctx,p,errs}=await open(b,1440,900);
  /* 1 — Worker + דף מגיב + מצב טעינה, על הקובץ הגדול */
  await p.evaluate(STALL);await p.setInputFiles('#f',BIG);await loaded(p);
  const w=await p.evaluate(()=>{clearInterval(__sti);return {via:LOADVIA,n:ALL.length,max:Math.round(__st.max),ticks:__st.n,saw:__st.saw,box:__st.box,cancel:__st.cancel,secs:__st.secs.size,
   hidden:document.getElementById('loadbox').hidden,key:localStorage.getItem('planner_loading_v1')}});
  ok('מחשב · הקריאה עוברת ב-Worker, גם מקובץ מקומי',w.via==='worker'&&w.n===20000,`${w.via} · ${w.n} פריטים`);
  ok('מחשב · הדף מגיב בזמן הקריאה (הפער הארוך ביותר בין טיקים)',w.max<600&&w.ticks>20,`${w.max}ms · ${w.ticks} טיקים`);
  ok('מחשב · מצב טעינה: שם הקובץ, גודל, «בטל טעינה»',w.saw>0&&/big-20000\.xlsx/.test(w.box)&&/(KB|MB)/.test(w.box)&&w.cancel,w.box.slice(0,90));
  ok('מחשב · הזמן שעובר מתעדכן',w.secs>=2,`${w.secs} ערכים שונים`);
  ok('מחשב · בסוף הטעינה התיבה יורדת והרישום «טעינה בעיצומה» נמחק',w.hidden&&w.key==null);
  const snapW=await p.evaluate(SNAP);
  /* 2 — גיבוי: בלי Worker → קריאה רגילה; הנתונים זהים; והדף כן נתפס (ההוכחה שהמדידה אמיתית) */
  await p.evaluate(()=>{window.__W=window.Worker;window.Worker=undefined});
  await p.evaluate(STALL);await p.setInputFiles('#f',BIG);await loaded(p);
  const m=await p.evaluate(()=>{clearInterval(__sti);return {via:LOADVIA,max:Math.round(__st.max),box:__st.box}});
  const snapM=await p.evaluate(SNAP);
  ok('מחשב · בלי Worker — קריאה רגילה (גיבוי), ומצב הטעינה אומר זאת',m.via==='main'&&/קריאה רגילה/.test(m.box),m.via+' · '+m.box.slice(0,80));
  ok('מחשב · הנתונים, הסיווג והתורים זהים בין Worker לקריאה רגילה',snapW===snapM,`${snapW.length} תווים`);
  ok('מחשב · ובקריאה הרגילה הדף נתפס — כלומר המדידה מבחינה',m.max>w.max*4,`רגילה ${m.max}ms מול Worker ${w.max}ms`);
  /* 3 — Worker שנכשל לפני שהחזיר דבר → גיבוי */
  await p.evaluate(()=>{const W0=window.__W;window.Worker=function(){return new W0(URL.createObjectURL(new Blob(['throw new Error("boom")'],{type:'text/javascript'})))}});
  await p.setInputFiles('#f',XL);await loaded(p);
  const f3=await p.evaluate(()=>({via:LOADVIA,n:ALL.length,diag:document.getElementById('diag').textContent}));
  ok('מחשב · Worker שנכשל בהפעלה — הקובץ נקרא בכל זאת (גיבוי), בלי הודעת שגיאה',f3.via==='main'&&f3.n>0&&!/⛔/.test(f3.diag),`${f3.via} · ${f3.n}`);
  await p.evaluate(()=>{window.Worker=window.__W});
  /* 4 — ביטול באמצע: שום דבר לא משתנה */
  await p.setInputFiles('#f',XL);await loaded(p);
  await p.evaluate(()=>{window.__all0=ALL;window.__q0=Q;window.__upd0=document.getElementById('updated').textContent;window.__lm0=LOADMS});
  await p.setInputFiles('#f',BIG);
  await p.waitForFunction(()=>!!document.getElementById('lbCancel'),null,{timeout:30000});
  await p.click('#lbCancel');await idle(p);await p.waitForTimeout(300);
  const c=await p.evaluate(()=>({same:ALL===__all0&&Q===__q0,n:ALL.length,upd:document.getElementById('updated').textContent===__upd0,lm:LOADMS===__lm0,
   hidden:document.getElementById('loadbox').hidden,key:localStorage.getItem('planner_loading_v1'),toast:(document.getElementById('toast')||{}).textContent||''}));
  ok('מחשב · ביטול באמצע הקריאה — הנתונים, הסיווג והכותרת לא השתנו',c.same&&c.upd&&c.lm,`${c.n} פריטים`);
  ok('מחשב · אחרי ביטול: התיבה יורדת, הרישום נמחק, והודעה שהנתונים לא השתנו',c.hidden&&c.key==null&&/בוטלה/.test(c.toast),c.toast);
  /* 5 — קובץ חדש באמצע טעינה: רק האחרון נכנס */
  await p.setInputFiles('#f',BIG);await p.waitForTimeout(150);await p.setInputFiles('#f',XL);await loaded(p);await p.waitForTimeout(1500);
  const r5=await p.evaluate(()=>({n:ALL.length,lj:LOADJOB}));
  const nXL=await p.evaluate(()=>__all0.length);
  ok('מחשב · קובץ חדש באמצע טעינה — רק האחרון נטען',r5.n===nXL&&!r5.lj,`${r5.n} פריטים`);
  /* 6 — קובץ פגום וקובץ בלי כותרת: אותן הודעות כמו ב-main, ו«נסה שוב» */
  const msg=async(page,file)=>{await page.setInputFiles('#f',file);await page.waitForFunction(()=>/⛔/.test(document.getElementById('diag').textContent),null,{timeout:60000});
   await page.waitForTimeout(200);
   return page.evaluate(()=>{const d=document.getElementById('diag');const t=[...d.childNodes].filter(x=>x.nodeType===3).map(x=>x.textContent).join('');
    return {upd:document.getElementById('updated').textContent,txt:t,retry:!!d.querySelector('button'),n:ALL?ALL.length:0}})};
  await p.evaluate(()=>{window.__all0=ALL});
  const bad1=await msg(p,BAD);
  const kept=await p.evaluate(()=>ALL===__all0);
  /* אותו קובץ ב-main (git show origin/main), אם יש — אחרת מול loadErrTxt של אותה שגיאה */
  let mainBad=null,mainNo=null;
  try{const mh=require('child_process').execSync('git show origin/main:index.html',{cwd:path.join(SD,'..'),stdio:['ignore','pipe','ignore'],maxBuffer:64e6});
   const MF=path.join(TMP,'main.html');fs.writeFileSync(MF,mh);const q=await ctx.newPage();await q.goto('file://'+MF+'?nobrief=1');await q.waitForTimeout(500);
   mainBad=await msg(q,BAD);mainNo=await msg(q,NOHDR);await q.close()}catch(_){}
  ok('מחשב · קובץ פגום — הודעה וכותרת כמו קודם',mainBad?(bad1.upd===mainBad.upd&&bad1.txt===mainBad.txt):/⛔/.test(bad1.txt),bad1.upd+' | '+bad1.txt.slice(0,70)+(mainBad?' · זהה ל-main':''));
  ok('מחשב · ו«נסה שוב», והנתונים הקודמים נשארו',bad1.retry&&kept);
  await p.click('#diag button');await p.waitForTimeout(500);await idle(p);
  const again=await p.evaluate(()=>({diag:/⛔/.test(document.getElementById('diag').textContent),retry:!!document.querySelector('#diag button'),same:ALL===__all0}));
  ok('מחשב · «נסה שוב» מנסה שוב את אותו קובץ (ונכשל באותה דרך), בלי לגעת בנתונים',again.diag&&again.retry&&again.same);
  const no1=await msg(p,NOHDR);
  ok('מחשב · קובץ בלי שורת כותרות — אותה הודעה כמו קודם',mainNo?(no1.upd===mainNo.upd&&no1.txt===mainNo.txt):/מק״ט מוביל/.test(no1.txt),no1.upd+(mainNo?' · זהה ל-main':''));
  /* 7 — ביטול בתוך ההדגמה: נשארים בהדגמה */
  await p.evaluate(()=>demoStart());await loaded(p);await p.waitForTimeout(300);
  const d0=await p.evaluate(()=>({demo:DEMO,n:ALL.length}));
  await p.setInputFiles('#f',BIG);await p.waitForFunction(()=>!!document.getElementById('lbCancel'),null,{timeout:30000});
  await p.click('#lbCancel');await idle(p);
  const d1=await p.evaluate(()=>({demo:DEMO,n:ALL.length,bar:!document.getElementById('demobar').hidden}));
  ok('מחשב · ביטול בזמן הדגמה — נשארים בהדגמה, עם נתוני ההדגמה',d0.demo&&d1.demo&&d1.bar&&d1.n===d0.n,`${d1.n} פריטים`);
  await p.setInputFiles('#f',XL);await loaded(p);
  const d2=await p.evaluate(()=>({demo:DEMO,bar:!document.getElementById('demobar').hidden}));
  ok('מחשב · ודוח אמיתי שנטען עד הסוף — יוצא מההדגמה, כמו קודם',!d2.demo&&!d2.bar);
  /* 8 — טעינה שלא הסתיימה: הלשונית «נסגרה» עם הרישום → הודעה בפתיחה הבאה */
  await p.evaluate(()=>localStorage.setItem('planner_loading_v1',JSON.stringify({n:'דוח-גדול.xlsx',s:60*1048576,t:Date.now()})));
  await p.reload();await p.waitForTimeout(800);
  const s8=await p.evaluate(()=>({box:document.getElementById('loadbox').textContent,hidden:document.getElementById('loadbox').hidden,key:localStorage.getItem('planner_loading_v1')}));
  ok('מחשב · טעינה שלא הסתיימה — בפתיחה הבאה: שם הקובץ, גודל והסבר',!s8.hidden&&/לא הסתיימה/.test(s8.box)&&/דוח-גדול\.xlsx/.test(s8.box)&&/60MB/.test(s8.box)&&s8.key==null,s8.box.slice(0,80));
  await p.click('#lbClose');
  ok('מחשב · «סגור» מסיר אותה, והיא לא חוזרת',await p.evaluate(()=>document.getElementById('loadbox').hidden));
  await p.reload();await p.waitForTimeout(600);
  ok('מחשב · פתיחה רגילה — אין הודעה',await p.evaluate(()=>document.getElementById('loadbox').hidden));
  ok('מחשב · אפס שגיאות JS',!errs.length,errs.slice(0,3).join(' | '));
  await ctx.close()}
 /* ---------------- טלפון ---------------- */
 {const {ctx,p,errs}=await open(b,390,844);
  await p.evaluate(STALL);await p.setInputFiles('#f',BIG);
  await p.waitForFunction(()=>!!document.getElementById('lbCancel'),null,{timeout:30000});
  const g=await p.evaluate(()=>{const r=document.getElementById('loadbox').getBoundingClientRect(),c=document.getElementById('lbCancel').getBoundingClientRect();
   return {l:Math.round(r.left),r:Math.round(r.right),w:innerWidth,cv:c.width>0&&c.left>=0&&c.right<=innerWidth}});
  ok('טלפון · מצב הטעינה בתוך המסך, «בטל טעינה» נראה',g.l>=8&&g.r<=g.w-8&&g.cv,`${g.l}–${g.r} מתוך ${g.w}`);
  await loaded(p);
  const t=await p.evaluate(()=>{clearInterval(__sti);return {via:LOADVIA,n:ALL.length,max:Math.round(__st.max)}});
  ok('טלפון · נקרא ב-Worker והדף מגיב',t.via==='worker'&&t.n===20000&&t.max<600,`${t.max}ms`);
  await p.setInputFiles('#f',BIG);await p.waitForFunction(()=>!!document.getElementById('lbCancel'),null,{timeout:30000});
  await p.evaluate(()=>{window.__all0=ALL});await p.click('#lbCancel');await idle(p);
  ok('טלפון · ביטול — הנתונים לא השתנו',await p.evaluate(()=>ALL===__all0&&document.getElementById('loadbox').hidden));
  ok('טלפון · אפס שגיאות JS',!errs.length,errs.slice(0,3).join(' | '));
  await ctx.close()}
 await b.close();fs.rmSync(TMP,{recursive:true,force:true});
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
