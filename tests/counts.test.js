/* ============ S3 · כל מספר אומר את מה שהוא סופר ============
   נמדד על הדוח האמיתי לפני S3:
     · הברכה «239 דורשים החלטה · 2,919 כסף תקוע» — והמסילה «מרכז עבודה 38»,
       «בריאות המלאי 2,779». שני מספרים לאותו שם.
     · «תכנון ו-MRP · 45 לעדכון ב-SAP» — וקובץ ה-SAP מאותו מסך: 33 שורות.
     · «מוצג 1-0 מתוך 0» בכל מסך ריק; «מתוך 7990» בלי מפריד בקטלוג.
     · «סה״כ ₪0» במסלול בלי מפתח כספי, וברשימה ריקה.
     · ההיכרות תיארה מסלולים «לטיפול היום» ו«הקטלוג», שכבר אינם; כותרות
       העזרה לא היו שמות הדלתות במסילה.

   מספרים שונים מותרים כשהם מדדים שונים — «תכנון ו-MRP 182» ולידו
   «33 לעדכון ב-SAP», «פרופיל MRP 116» מול רשימה של 5,046 — בתנאי
   שהתווית אומרת מה כל אחד סופר. הבדיקה אוכפת את שני הצדדים.

   כל הבדיקות רצות בחמישה מצבים: אחרי טעינה, אחרי סינון, כשהסינון
   מרוקן הכול, אחרי ניקוי הסינון, ואחרי טעינת דוח אחר.
   COUNTS_XLSX=<דוח> — על דוח אחר, מקומית. */
const {chromium}=require('playwright');const {execFileSync}=require('child_process');
const fs=require('fs'),path=require('path'),os=require('os');
const SD=__dirname,ROOT=path.join(SD,'..');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const XL=process.env.COUNTS_XLSX||path.join(SD,'zmrp-demo.xlsx');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
const gn=s=>s==null?0:+String(s).replace(/,/g,'');

/* דוח שני — אותו מחולל, זרע אחר, בתיקייה זמנית */
const tmp=fs.mkdtempSync(path.join(os.tmpdir(),'counts-'));
fs.copyFileSync(path.join(SD,'make-fixture.js'),path.join(tmp,'make-fixture.js'));
execFileSync(process.execPath,[path.join(tmp,'make-fixture.js')],
 {cwd:tmp,env:{...process.env,NODE_PATH:path.join(ROOT,'node_modules'),FIXTURE_SEED:'7'},stdio:'pipe'});
const XL2=path.join(tmp,'zmrp-demo.xlsx');

const snap=p=>p.evaluate(()=>{
 const num=s=>{const m=String(s||'').match(/[\d,]+/);return m?+m[0].replace(/,/g,''):null};
 const T=s=>(document.querySelector(s)||{}).textContent?.replace(/\s+/g,' ').trim()||'';
 const bdg=sel=>Object.fromEntries([...document.querySelectorAll(sel)].map(e=>[e.dataset.m,num(e.querySelector('.bdg')?.textContent)]));
 setMode('home');
 const main=bdg('#tabs .tab[data-m]');
 const greet=T('#plannerGreeting'),qsub=T('#qsub');
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
 const monthSub=T('#tabs .tab[data-m="month"] .m.stat');
 setMode('home');
 return {main,sub,title,greet,qsub,scr,monthFile,monthSub,all:QF.all.length}});

function invariants(tag,R){
 const g=R.greet.match(/([\d,]+) דורשים החלטה(?: · ([\d,]+) פריטים בכסף תקוע)?/)||[];
 ok(`${tag} · הברכה: «דורשים החלטה» = «מרכז עבודה», «פריטים בכסף תקוע» = «בריאות המלאי»`,
   !!g[0]&&gn(g[1])===R.main.home&&gn(g[2])===R.main.cap,`${R.greet.replace(/^.*?· /,'')} · מסילה ${R.main.home} / ${R.main.cap}`);
 const q=R.qsub.match(/([\d,]+) דורשים החלטה · ([\d,]+) פריטים בכסף תקוע/)||[];
 ok(`${tag} · כותרת התור — אותם שני מספרים`,gn(q[1])===R.main.home&&gn(q[2])===R.main.cap);
 const all={...R.main,...R.sub};
 const b1=Object.entries(all).filter(([m,n])=>m!=='moves'&&m!=='prof'&&R.scr[m]&&R.scr[m].n!==n);
 ok(`${tag} · כל מונה במסילה = הפריטים במסך שנפתח`,!b1.length,
   b1.length?b1.map(([m,n])=>`${m}: ${n}≠${R.scr[m].n}`).join(' · '):`${Object.keys(all).length-2} מונים`);
 const pr=R.scr.prof;
 ok(`${tag} · «פרופיל MRP»: המונה = «לפעולה», והמסך והתיאור אומרים זאת`,
   R.sub.prof===pr.act&&/להעלות \+ לבדוק/.test(R.title.prof||'')
   &&(pr.n===0||new RegExp(`${pr.act.toLocaleString('en-US')} לפעולה \\(להעלות \\+ לבדוק\\)`).test(pr.pgL)),
   `${R.sub.prof} · «${pr.pgL}»`);
 const ms=R.monthSub.match(/([\d,]+) לעדכון ב-SAP/);
 ok(`${tag} · «N לעדכון ב-SAP» = השורות בקובץ «עדכון ל-SAP» מהמסך`,
   R.monthFile===0?!ms:!!ms&&gn(ms[1])===R.monthFile,`${ms?ms[1]:'—'} מול ${R.monthFile}`);
 const pg=Object.entries(R.scr).filter(([m])=>m!=='moves');
 const b2=pg.filter(([m,s])=>s.n===0?s.pgL!=='אין פריטים'
   :!(new RegExp(`מתוך ${s.n.toLocaleString('en-US')}( ·|$)`).test(s.pgL)));
 ok(`${tag} · «מתוך N» = המסך; «אין פריטים» כשריק`,!b2.length,
   b2.slice(0,3).map(([m,s])=>`${m}: «${s.pgL}» (${s.n})`).join(' · ')||`${pg.length} מסכים`);
 const raw=Object.entries(R.scr).filter(([,s])=>/(^|[^\d,.])\d{4,}(?![\d,])/.test(s.pgL+' '+s.pgR+' '+s.movpg));
 ok(`${tag} · מפריד אלפים בכל מספר`,!raw.length,raw.slice(0,2).map(([m,s])=>`${m}: ${(s.pgL+' | '+s.pgR).slice(0,60)}`).join(' · '));
 const z=Object.entries(R.scr).filter(([,s])=>/סה״כ ₪0(?!\d)/.test(s.pgR));
 ok(`${tag} · אין «סה״כ ₪0»`,!z.length,z.map(([m])=>m).join(', '))}

(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(ROOT,'index.html')+'?nobrief=1');
 const load=async f=>{await p.setInputFiles('#f',f);
  await p.waitForFunction(n=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&ALL.length!==n,null,{timeout:180000});
  await p.waitForTimeout(800)};
 await p.evaluate(()=>window.__n0=0);await load(XL);

 /* ---- 1. אחרי טעינה ---- */
 const R1=await snap(p);invariants('טעינה',R1);
 /* ---- 2. אחרי סינון (המותג הגדול ביותר) ---- */
 const fk=await p.evaluate(()=>{const k=FDEF[0][0];const c={};for(const r of ALL)c[r[k]]=(c[r[k]]||0)+1;
  const v=Object.entries(c).sort((a,b)=>b[1]-a[1])[0][0];SEL[k]=new Set([v]);apply();return {k,n:c[v]}});
 const R2=await snap(p);invariants('סינון',R2);
 ok('סינון · המונים באמת התעדכנו',R2.all===fk.n&&R2.all<R1.all&&R2.main.catalog===fk.n,`${R1.all} → ${R2.all}`);
 /* ---- 3. סינון שמרוקן הכול: אפס פריטים ---- */
 await p.evaluate(k=>{SEL[k]=new Set(['__אין_כזה__']);apply()},fk.k);
 const R3=await p.evaluate(()=>{const T=s=>(document.querySelector(s)||{}).textContent?.replace(/\s+/g,' ').trim()||'';
  const res={};for(const m of Object.keys(MODESET)){setMode(m);res[m]={pgL:T('#pgL'),pgR:T('#pgR'),
   n:(()=>{try{return m==='moves'?movRows().length:currentRows().length}catch(_){return 0}})()}}
  setMode('home');
  return {greet:T('#plannerGreeting'),bdg:[...document.querySelectorAll('#tabs .tab[data-m] .bdg')].map(e=>e.textContent.trim()),res}});
 ok('אפס פריטים · הברכה אומרת 0',/(^|· )0 דורשים החלטה$/.test(R3.greet),R3.greet);
 ok('אפס פריטים · כל המונים במסילה 0',R3.bdg.every(t=>t==='0'||t==='▤'),R3.bdg.join(' '));
 const e3=Object.entries(R3.res).filter(([m,s])=>m!=='moves'&&(s.pgL!=='אין פריטים'||/₪0|\$0|1-0/.test(s.pgR)));
 ok('אפס פריטים · כל מסך: «אין פריטים», בלי «₪0» ובלי «1-0»',!e3.length,e3.slice(0,3).map(([m,s])=>`${m}: ${s.pgL} | ${s.pgR}`).join(' · '));
 /* ---- 4. ניקוי הסינון — חוזרים בדיוק למספרים של הטעינה ---- */
 await p.evaluate(k=>{SEL[k]=new Set();apply()},fk.k);
 const R4=await snap(p);
 ok('ניקוי סינון · כל מספר חזר לערך שלו אחרי הטעינה',
   JSON.stringify([R4.main,R4.sub,R4.greet,R4.monthSub])===JSON.stringify([R1.main,R1.sub,R1.greet,R1.monthSub]));
 /* ---- 5. דוח אחר ---- */
 await p.evaluate(()=>{for(const [k] of FDEF)SEL[k]=new Set()});
 await load(XL2);
 const R5=await snap(p);invariants('דוח חדש',R5);
 ok('דוח חדש · המספרים הם של הדוח החדש',R5.all!==R1.all||JSON.stringify(R5.main)!==JSON.stringify(R1.main),`${R1.all} → ${R5.all} פריטים`);

 /* ---- ההיכרות והעזרה מול הממשק ---- */
 const H=await p.evaluate(()=>{
  setMode('home');
  const tr=document.querySelector('#tbl tbody tr[data-i]');if(tr)tr.click();
  /* כל הטאבים, גם הסמויים: פריט עם המלצת פרמטר, כמו שההיכרות אומרת */
  const fx=QF.all.find(r=>r.paramFix);const det0=(document.getElementById('detail').textContent||'');
  if(fx)detail(fx);
  const names={};for(const [k,n] of NAV)names[k]=n;for(const a in SUBNAV)for(const [k,n] of SUBNAV[a])if(!names[k])names[k]=n;
  const nav=Object.entries(XP).filter(([k])=>k.startsWith('nav-')).map(([k,d])=>({k:k.slice(4),t:d.t,name:names[k.slice(4)]}));
  const det=(document.getElementById('detail').innerText||'');
  const tabs=Object.entries(XP).filter(([k])=>k.startsWith('dt-')).map(([k,d])=>({k,t:(d.t.match(/«(.+?)»/)||[])[1],
   here:!!document.querySelector(`#detail [data-xp="${k}"]`)}));
  return {nav,tabs,det,tour:TOUR.map(t=>({t:t.t,x:t.x,sel:t.sel,found:!!document.querySelector(t.sel)})),
   areas:NAV.length,body:det0+' '+document.body.textContent}});
 const badNav=H.nav.filter(x=>x.t!==x.name);
 ok('כותרת כל עזרה של דלת = שם הדלת במסילה',!badNav.length,badNav.map(x=>`${x.k}: «${x.t}» ≠ «${x.name}»`).join(' · ')||`${H.nav.length} דלתות`);
 ok('טאבי הכרטיס שהעזרה מזכירה — קיימים בכרטיס',H.tabs.every(x=>x.here&&H.det.includes(x.t)),H.tabs.map(x=>x.t).join(', '));
 ok('לכל צעד בהיכרות יש יעד בממשק',H.tour.every(s=>s.found),H.tour.filter(s=>!s.found).map(s=>s.t).join(', '));
 const tourAll=H.tour.map(s=>s.x).join(' ');
 const quoted=[...new Set([...tourAll.matchAll(/«([^»]+)»/g)].map(m=>m[1]))];
 /* «בטל» קיים רק בהודעה שאחרי סימון — נבדק ב-keys.test.js */
 const missing=quoted.filter(q=>q!=='בטל'&&!H.body.includes(q.replace(/^✓ /,'')));
 ok('כל שם שההיכרות מצטטת — מופיע בממשק',!missing.length,missing.join(', ')||quoted.join(', '));
 ok('ההיכרות: מספר התחומים נכון, ובלי «לטיפול היום»',
   !/לטיפול היום/.test(tourAll)&&(!/שבעה תחומים/.test(tourAll)||H.areas===7),`${H.areas} תחומים`);
 ok('ההיכרות מזכירה Enter ו-D',/Enter/.test(tourAll)&&/המקש D/.test(tourAll));

 ok('אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();fs.rmSync(tmp,{recursive:true,force:true});
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
