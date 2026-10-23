/* ============ S3 · כל מספר אומר את מה שהוא סופר ============
   נמדד על הדוח האמיתי לפני S3:
     · הברכה «239 דורשים החלטה · 2,919 כסף תקוע» — והמסילה «מרכז עבודה 38»,
       «בריאות המלאי 2,779». שני מספרים לאותו שם.
     · «תכנון ו-MRP · 45 לעדכון ב-SAP» — וקובץ ה-SAP מאותו מסך: 33 שורות.
     · «מוצג 1-0 מתוך 0» בכל מסך ריק; «מתוך 7990» בלי מפריד בקטלוג.
     · «איכות נתונים»: «סה״כ ₪0» — סכום של מפתח כספי שאין למסלול.
     · ההיכרות תיארה מסלולים «לטיפול היום» ו«הקטלוג», שכבר אינם.

   הבדיקה עוברת על כל אחד מ-26 המסכים ועל כל מספר במסילה, ומשווה כל
   מספר לרשימה שהוא מייצג. COUNTS_XLSX=<דוח> — על דוח אחר, מקומית. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const XL=process.env.COUNTS_XLSX||path.join(SD,'zmrp-demo.xlsx');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0,null,{timeout:180000});
 await p.waitForTimeout(800);
 const R=await p.evaluate(()=>{
  const num=s=>{const m=String(s||'').match(/[\d,]+/);return m?+m[0].replace(/,/g,''):null};
  const T=s=>(document.querySelector(s)||{}).textContent?.replace(/\s+/g,' ').trim()||'';
  const bdg=sel=>Object.fromEntries([...document.querySelectorAll(sel)].map(e=>[e.dataset.m,num(e.querySelector('.bdg')?.textContent)]));
  setMode('home');
  const main=bdg('#tabs .tab[data-m]');
  const greet=T('#plannerGreeting'),qsub=T('#qsub');
  /* דלתות המשנה מצוירות רק לתחום הפעיל — עוברים על כל תחום */
  const sub={},title={};
  for(const a of Object.keys(main)){setMode(a);
   for(const e of document.querySelectorAll('#railsub .railsub2[data-m]')){
    sub[e.dataset.m]=num(e.querySelector('.bdg')?.textContent);title[e.dataset.m]=e.title}}
  const scr={};
  for(const m of Object.keys(MODESET)){setMode(m);
   let n;try{n=m==='moves'?movRows().length:currentRows().length}catch(_){n=null}
   scr[m]={n,pgL:T('#pgL'),pgR:T('#pgR'),movpg:T('#movpg'),
    act:m==='prof'?profRows().filter(profActive).length:null}}
  setMode('month');const v=currentRows();
  const monthFile=sapPlan(undefined,v,v).rows;
  const monthSub=[...document.querySelectorAll('#tabs .tab[data-m="month"] .m.stat')].map(e=>e.textContent).join('');
  return {main,sub,title,greet,qsub,scr,monthFile,monthSub,tour:TOUR.map(t=>t.x).join(' ')}});

 /* ---- הברכה = המסילה ---- */
 const g=R.greet.match(/([\d,]+) דורשים החלטה(?: · ([\d,]+) כסף תקוע)?/)||[];
 const gn=s=>s==null?0:+s.replace(/,/g,'');
 ok('הברכה: «דורשים החלטה» = «מרכז עבודה» במסילה',gn(g[1])===R.main.home,`${g[1]} מול ${R.main.home}`);
 ok('הברכה: «כסף תקוע» = «בריאות המלאי» במסילה',gn(g[2])===R.main.cap,`${g[2]||0} מול ${R.main.cap}`);
 const q=R.qsub.match(/([\d,]+) דורשים החלטה · ([\d,]+) כסף תקוע/)||[];
 ok('כותרת התור — אותם שני מספרים',gn(q[1])===R.main.home&&gn(q[2])===R.main.cap,R.qsub.slice(0,60));

 /* ---- כל מספר במסילה = הרשימה שנפתחת ---- */
 const all={...R.main,...R.sub};
 const bad1=Object.entries(all).filter(([m,n])=>m!=='moves'&&m!=='prof'&&R.scr[m]&&R.scr[m].n!==n);
 ok('כל מספר במסילה = מספר הפריטים במסך שנפתח',!bad1.length,
   bad1.length?bad1.map(([m,n])=>`${m}: ${n}≠${R.scr[m].n}`).join(' · '):`${Object.keys(all).length-2} מונים`);
 ok('«פרופיל MRP»: המונה = להעלות + לבדוק, והתיאור אומר את זה',
   R.sub.prof===R.scr.prof.act&&/להעלות \+ לבדוק/.test(R.title.prof||''),`${R.sub.prof} לפעולה מתוך ${R.scr.prof.n}`);

 /* ---- «לעדכון ב-SAP» = הקובץ ---- */
 const ms=R.monthSub.match(/([\d,]+) לעדכון ב-SAP/);
 ok('«תכנון ו-MRP · N לעדכון ב-SAP» = מספר השורות בקובץ «עדכון ל-SAP» מהמסך',
   R.monthFile===0?!ms:ms&&gn(ms[1])===R.monthFile,`${ms?ms[1]:'—'} מול ${R.monthFile} בקובץ`);

 /* ---- שורת הדפדוף ---- */
 const pg=Object.entries(R.scr).filter(([m])=>m!=='moves');
 const badPg=pg.filter(([m,s])=>s.n===0?s.pgL!=='אין פריטים':!(new RegExp(`מתוך ${s.n.toLocaleString('en-US')}$`).test(s.pgL)));
 ok('«מתוך N» = מספר הפריטים במסך, ו«אין פריטים» כשהוא ריק',!badPg.length,
   badPg.slice(0,4).map(([m,s])=>`${m}: «${s.pgL}» (${s.n})`).join(' · ')||`${pg.length} מסכים`);
 ok('אין «מוצג 1-0» באף מסך',!pg.some(([,s])=>/1-0 /.test(s.pgL)));
 const raw=Object.entries(R.scr).filter(([,s])=>/(^|[^\d,.])\d{4,}(?![\d,])/.test(s.pgL+' '+s.pgR+' '+s.movpg));
 ok('כל מספר בן 4 ספרות ומעלה — עם מפריד אלפים',!raw.length,raw.slice(0,3).map(([m,s])=>`${m}: ${(s.pgL+' | '+s.pgR).slice(0,70)}`).join(' · '));
 ok('אין «סה״כ ₪0» — סכום שלא מייצג כלום',!Object.values(R.scr).some(s=>/סה״כ ₪0(?!\d)/.test(s.pgR)),R.scr.qual.pgR.slice(0,60));

 /* ---- ההיכרות ---- */
 ok('ההיכרות לא מזכירה מסלולים שאינם («לטיפול היום»)',!/לטיפול היום/.test(R.tour));
 ok('ההיכרות מזכירה Enter ו-D',/Enter/.test(R.tour)&&/המקש D/.test(R.tour));
 ok('אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
