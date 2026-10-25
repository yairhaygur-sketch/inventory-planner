/* ============ D2B · «בחר את כל N» — הבחירה לקובץ SAP על כל הסינון ============
   מאז D2A רק החלון שבגלילה מצויר, ולכן «הכול» צריך להיות מפורש:
     · התיבה שבכותרת — השורות שעל המסך בלבד
     · «בחר את כל N שבסינון» בפס הבחירה — כל מי שבסינון ויש לו תיבה,
       גם כאלה שלא נגללו אליהם
     · הפס אומר כמה נבחרו מתוך כמה, כמה ייכללו בקובץ — והמספר הזה הוא
       בדיוק מספר השורות בקובץ שיוצא ובתצוגה לפני הייצוא
     · נבחרים מחוץ לסינון הנוכחי (סינון עמודה, דלת אחרת) — כמה, וביטול
       שלהם בלבד; «נקה בחירה» מבטל הכול; ביטול ידני של אחת — חלקי
     · מעבר מסך והחזרה — הבחירה נשמרת; במסך בלי תיבות הפס יורד
     · טלפון (בלי כותרת): תיבה אחת מעלה את הפס, ומשם «בחר את כל N»
   SELALL_XLSX=<דוח> — על דוח אחר (למשל האמיתי, מקומית). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const XL=process.env.SELALL_XLSX||path.join(SD,'zmrp-demo.xlsx');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
const N=n=>Number(n).toLocaleString('he-IL');
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 for(const [W,H] of [[1440,900],[390,844]]){
  const ctx=await b.newContext({viewport:{width:W,height:H}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');await p.setInputFiles('#f',XL);
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0,null,{timeout:180000});
  await p.waitForTimeout(600);
  await p.evaluate(()=>{window.__raf=()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
   window.__cap=[];XLSX.writeFile=(wb,name)=>{const ws=wb.Sheets[wb.SheetNames[0]];
    window.__cap.push({name,rows:XLSX.utils.sheet_to_json(ws,{header:1}).length-1})}});
  const R=await p.evaluate(async()=>{const T=id=>(document.getElementById(id)||{}).textContent||'';
   const hidden=id=>document.getElementById(id).hidden;
   const o={};SAPSEL.clear();setMode('month');await __raf();
   const thead=document.querySelector('#tbl thead');o.head=!!thead&&getComputedStyle(thead).display!=='none';
   o.N=sapFiltered().length;o.barStart=hidden('sapbar');
   /* 1 — כניסה */
   /* השורות שעל המסך ברגע הלחיצה — אחריה הפס מופיע ודוחק את הטבלה */
   if(o.head){const vis=sapVisible().map(r=>r.pn);const a=document.querySelector('#tbl thead .sapall');a.checked=true;a.dispatchEvent(new Event('change'));
    o.vis=vis.length;o.onlyVis=SAPSEL.size===vis.length&&vis.every(x=>SAPSEL.has(x))}
   else{const b0=document.querySelector('#tbl tbody .sapsel');b0.checked=true;b0.dispatchEvent(new Event('change'));o.vis=1;o.onlyVis=SAPSEL.size===1}
   await __raf();o.bar1=!hidden('sapbar');o.allBtn=T('sapbarAll');o.info1=T('sapbarInfo');o.n1=SAPSEL.size;
   /* 2 — «בחר את כל N» */
   document.getElementById('sapbarAll').click();await __raf();
   const F=sapFiltered();o.sel2=SAPSEL.size;o.allF=F.every(r=>SAPSEL.has(r.pn))&&SAPSEL.size===F.length;
   o.allHidden=hidden('sapbarAll');o.info2=T('sapbarInfo');o.cnt2=T('sapbarN');
   const s=VL.tbl;s.box.scrollTop=s.box.scrollHeight;await __raf();await __raf();
   o.farAll=[...s.tb.querySelectorAll('.sapsel')].every(x=>x.checked);
   /* 3 — «ייכללו» = הקובץ = התצוגה לפני הייצוא */
   const P=sapPlan(undefined,sapSelRows(),QF.all.filter(r=>SAPSEL.has(r.pn)));o.plan=P.rows;o.planOut=P.out;
   __cap.length=0;doSapMass(undefined,sapSelRows());o.file=__cap.length?__cap[0].rows:0;
   document.getElementById('sapbarGo').click();await __raf();o.prevIn=T('sapIn');o.modal=document.getElementById('modalwrap').classList.contains('open');
   document.getElementById('sapCancel').click();
   /* 4 — סינון עמודה: נבחרים מחוץ לסינון */
   const col='c:model',vals={};for(const r of F){const v=String(colAcc(col)(r));vals[v]=(vals[v]||0)+1}
   const ent=Object.entries(vals).sort((a,b)=>b[1]-a[1]);o.multi=ent.length>1;
   if(o.multi){colFilters[col]={search:'',values:new Set([ent[0][0]]),min:null,max:null};render();await __raf();
    const fs=new Set(sapFiltered().map(r=>r.pn));o.expOut=sapSelRows().filter(r=>!fs.has(r.pn)).length;o.inF=sapFiltered().length;
    o.outShown=!hidden('sapbarOut');o.outTxt=T('sapbarOutN');o.info4=T('sapbarInfo');
    document.getElementById('sapbarOutClr').click();await __raf();
    o.after4=SAPSEL.size;o.keptIn=sapFiltered().every(r=>SAPSEL.has(r.pn));o.outHidden=hidden('sapbarOut');
    delete colFilters[col];render();await __raf();o.info4b=T('sapbarInfo');o.allBack4=!hidden('sapbarAll')}
   /* 5 — סינון במגירה: הבחירה לא חורגת מהנתונים המסוננים (כמו קודם) */
   document.getElementById('sapbarAll').click();await __raf();
   const k=FDEF[0][0],cnt={};for(const r of sapFiltered())cnt[r[k]]=(cnt[r[k]]||0)+1;const kv=Object.entries(cnt).sort((a,b)=>b[1]-a[1]);
   o.drawerMulti=kv.length>1;
   if(o.drawerMulti){SEL[k]=new Set([kv[0][0]]);apply();setMode('month');await __raf();
    const live=new Set(QF.all.map(r=>r.pn));o.subset=[...SAPSEL].every(x=>live.has(x));o.sel5=SAPSEL.size;
    SEL[k]=new Set();apply();setMode('month');await __raf()}
   /* 6 — מעבר מסך והחזרה */
   const before6=SAPSEL.size;setMode('catalog');await __raf();o.barCat=hidden('sapbar');
   setMode('month');await __raf();o.kept6=SAPSEL.size===before6&&before6>0;o.bar6=!hidden('sapbar');
   /* 7 — ביטול חלקי ידני, ואז מלא */
   document.getElementById('sapbarAll').click();await __raf();const full=SAPSEL.size;
   const bx=document.querySelector('#tbl tbody .sapsel:checked');bx.checked=false;bx.dispatchEvent(new Event('change'));await __raf();
   o.partial=SAPSEL.size===full-1;o.cnt7=T('sapbarN');o.allBack7=!hidden('sapbarAll');o.full=full;
   document.getElementById('sapbarClr').click();await __raf();o.cleared=SAPSEL.size===0&&hidden('sapbar')
    &&[...document.querySelectorAll('#tbl tbody .sapsel')].every(x=>!x.checked);
   /* 8 — דלת «רצפת מדרגה»: הפס מבטיח את הקובץ של הדלת הזאת */
   setMode('tier');await __raf();o.tierN=sapFiltered().length;
   if(o.tierN){for(const r of sapFiltered())SAPSEL.add(r.pn);sapBarSync();
    const PT=sapPlan(undefined,sapSelRows(),QF.all.filter(r=>SAPSEL.has(r.pn)));
    __cap.length=0;doSapMass(undefined,sapSelRows());o.tierPlan=PT.rows;o.tierFile=__cap.length?__cap[0].rows:0;
    o.tierInfo=new RegExp(`ייכללו בקובץ ${PT.rows.toLocaleString('he-IL')}(\\D|$)`).test(T('sapbarInfo'));SAPSEL.clear();sapBarSync()}
   return o});
  const tag=W>760?'מחשב':'טלפון';
  ok(`${tag} · הפס לא מוצג לפני שנבחר משהו`,R.barStart);
  ok(`${tag} · ${R.head?'התיבה שבכותרת בוחרת רק את השורות שעל המסך':'בלי כותרת — תיבה אחת מעלה את הפס'}`,R.onlyVis&&R.bar1,`${R.n1} נבחרו · ${R.N} בסינון`);
  ok(`${tag} · הפס מציע «בחר את כל N שבסינון» ואומר כמה מתוך כמה`,R.allBtn===`בחר את כל ${N(R.N)} שבסינון`&&new RegExp(`${N(R.n1)} מתוך ${N(R.N)} שבסינון`).test(R.info1),R.info1);
  ok(`${tag} · «בחר את כל N» בוחר בדיוק את כל מי שבסינון — גם מחוץ לחלון`,R.allF&&R.sel2===R.N&&R.farAll&&R.allHidden&&R.cnt2===N(R.N),R.info2);
  ok(`${tag} · «ייכללו בקובץ» = מספר השורות בקובץ = התצוגה לפני הייצוא`,
    R.plan===R.file&&R.prevIn===N(R.plan)&&R.modal&&new RegExp(`ייכללו בקובץ ${N(R.plan)}(\\D|$)`).test(R.info2),
    `${R.plan} · קובץ ${R.file} · תצוגה ${R.prevIn}${R.planOut?` · ${R.planOut} לא ייכללו`:''}`);
  if(R.multi){
   ok(`${tag} · סינון עמודה: הפס אומר כמה נבחרים מחוץ לסינון`,R.outShown&&R.expOut>0&&R.outTxt.startsWith(N(R.expOut)),R.outTxt);
   ok(`${tag} · «בטל אותם» מבטל רק את מי שמחוץ לסינון`,R.after4===R.inF&&R.keptIn&&R.outHidden,`${R.after4} נשארו`);
   ok(`${tag} · בלי הסינון — שוב «X מתוך N» ו«בחר את כל»`,R.allBack4&&/מתוך/.test(R.info4b),R.info4b)}
  if(R.drawerMulti)ok(`${tag} · סינון במגירה: הבחירה לא חורגת מהנתונים המסוננים (כמו לפני D2b)`,R.subset,`${R.sel5} נבחרו`);
  ok(`${tag} · מעבר מסך: במסך בלי תיבות הפס יורד, ובחזרה הבחירה נשמרת`,R.barCat&&R.kept6&&R.bar6);
  ok(`${tag} · ביטול ידני של אחת — חלקי, ו«בחר את כל» חוזר`,R.partial&&R.cnt7===N(R.full-1)&&R.allBack7,R.cnt7);
  ok(`${tag} · «נקה בחירה» מבטל הכול`,R.cleared);
  if(R.tierN)ok(`${tag} · «רצפת מדרגה»: הפס מבטיח את הקובץ של הדלת`,R.tierPlan===R.tierFile&&R.tierInfo,`${R.tierPlan}/${R.tierFile}`);
  ok(`${tag} · אפס שגיאות JS`,errs.length===0,errs.slice(0,2).join(' | '));
  await ctx.close()}
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
