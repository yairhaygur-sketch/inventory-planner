/* ============ BASE_MONTH — חודש הבסיס, תאריך הייחוס, והיום ============
   הדרישה (המעתד, 1.11.2026): אותו קובץ, אותם נתונים ואותן הגדרות — אותן
   המלצות, גם אם נטען ביום אחר או בחודש אחר.

   נבדק כאן, על נתונים סינתטיים בלבד ובתאריכים קבועים ומפורשים:
     1. זיהוי חודש הבסיס מהנתונים — כל 12 החודשים, כולל דצמבר/ינואר ומעבר
        שנה; השנה מאומתת בנפרד; ספי ביטחון.
     2. הכותרת: תאריך אמין, עתידי, מוקדם מהפעילות, לא תקין, וסותר לנתונים.
     3. אותו קובץ בארבעה שעונים — כל שדה של כל פריט, כל תור, ייצוא לאקסל
        וקובצי SAP זהים. כולל גילי מלאי, סיווגים ומונים.
     4. פעילות אחרי תאריך הייחוס — גיל 0, מסומנת ומוסברת, ואינה «עתידית»
        ואינה מלאי מת; תאריך שמעבר לגבול — שגוי, כמו קודם.
     5. חודש לא מאומת — אזהרה, אין רישום הרצה/פרופיל/מעקב יישום, אין SAP;
        אישור או בחירה — נרשם פעם אחת, SAP נפתח.
     6. בחירה ידנית — נשמרת לתוכן הקובץ, לא לקובץ אחר; שינוי חודש לאותו
        תוכן אינו יוצר הרצה שנייה.
     7. היום נשאר לפעולות תפעוליות — ETA.
     8. ההדגמה שבתוך הכלי מאומתת. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path'),os=require('os');
const XLSX=require('xlsx');
const {FIXTURE_NOW}=require('./lib/fixture-date.js');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
const TMP=fs.mkdtempSync(path.join(os.tmpdir(),'bm-'));
const MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const dstr=d=>`${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`;

/* ---- וריאנטים של הדוח הסינתטי, כולם בתאריכים קבועים ---- */
const DEMO=(()=>{const wb=XLSX.readFile(path.join(SD,'zmrp-demo.xlsx'));const sn=wb.SheetNames[0];
 const a=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:''});
 const hi=a.findIndex(r=>r.includes('מק"ט מוביל'));return {sn,head:a.slice(0,hi),hdr:a[hi],rows:a.slice(hi+1)}})();
const H=k=>DEMO.hdr.indexOf(k);
/* k = מספר חודשי השנה הנוכחית בחלון; מחזיר «צר.השנה» עקבי */
function y0For(r,k){let s=+r[H('צר.החודש')]||0;for(let i=0;i<k;i++)s+=+r[H(MON[i])]||0;return s}
let VN=0;
function variant({head=null,k=null,last=null,mut=null,rows=null}={}){
 const R=(rows||DEMO.rows).map(r=>r.slice());
 if(k!=null)for(const r of R)r[H('צר.השנה')]=y0For(r,k);
 /* הפעילות האחרונה בקובץ: כל התאריכים מוזזים כך שהמאוחר שבהם הוא `last` */
 if(last){const P=s=>{const m=String(s).match(/(\d+)\.(\d+)\.(\d+)/);return m?new Date(+m[3],+m[2]-1,+m[1]):null};
  let mx=null;for(const r of R)for(const c of ["תאר' מכירה","תאר' כניסה"]){const d=P(r[H(c)]);if(d&&(!mx||d>mx))mx=d}
  const sh=Math.round((last-mx)/864e5);
  for(const r of R)for(const c of ["תאר' מכירה","תאר' כניסה"]){const d=P(r[H(c)]);if(d)r[H(c)]=dstr(new Date(d.getFullYear(),d.getMonth(),d.getDate()+sh))}}
 if(mut)mut(R);
 const aoa=[head||['דוח ZMRP — תכנון מלאי'],[],DEMO.hdr,...R];
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),DEMO.sn);
 const f=path.join(TMP,`v${++VN}.xlsx`);XLSX.writeFile(wb,f);return f}

async function page(b,now){const ctx=await b.newContext({viewport:{width:1440,height:900},timezoneId:'Asia/Jerusalem'});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.clock.setFixedTime(new Date(now));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 return {ctx,p,errs}}
async function load(p,f){const z0=await p.evaluate(()=>typeof ZLOAD==='undefined'?null:ZLOAD);
 await p.evaluate(()=>{window.__z=(typeof ZN!=='undefined'&&ZN)?ZN:null});
 await p.setInputFiles('#f',f);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB&&ZN&&ZN!==window.__z,null,{timeout:180000});
 await p.waitForTimeout(150)}
const baseOf=p=>p.evaluate(()=>BASE&&{y:BASE.y,m:BASE.m,src:BASE.src,ver:BASE.verified,
 state:BASE.state?dsigDate(BASE.state):null,ref:dsigDate(BASE.ref),assumed:BASE.refAssumed,lim:dsigDate(BASE.lim),
 last:BASE.last?dsigDate(BASE.last):null,warn:BASE.warn.slice(),cs:BASE.cs});

(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});

 /* ===== 1 · הזיהוי מהנתונים — הפונקציה עצמה, כל החודשים ===== */
 {const {ctx,p,errs}=await page(b,FIXTURE_NOW);
  const cases=[];
  for(let mon=1;mon<=12;mon++){
   /* הבסיס: mon/2026. הפעילות האחרונה: 10 בחודש שאחריו */
   const y=2026,n=mon===12?{y:2027,m:1}:{y,m:mon+1};
   cases.push({mon,y,k:mon===12?0:mon,last:[n.y,n.m-1,10]})}
  /* מעבר שנה: בסיס ינואר 2027, פעילות אחרונה בינואר עצמו */
  cases.push({mon:1,y:2027,k:1,last:[2027,0,28]});
  const R=await p.evaluate(({cases,rows,hdr})=>cases.map(c=>{
   const H=k=>hdr.indexOf(k),MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
   const L=new Date(c.last[0],c.last[1],c.last[2]);
   const rr=rows.map(r=>{const x=r.slice();let s=+x[H('צר.החודש')]||0;for(let i=0;i<c.k;i++)s+=+x[H(MON[i])]||0;x[H('צר.השנה')]=s;
    x[H("תאר' מכירה")]=x[H("תאר' כניסה")]=`${L.getDate()}.${L.getMonth()+1}.${L.getFullYear()}`;return x});
   const r=bmResolve(rr,hdr,[['ללא תאריך']],new Date(2027,5,1));
   return {want:`${c.y}-${c.mon}`,got:`${r.y}-${r.m}`,src:r.src,ok:r.data&&r.data.ok}}),{cases,rows:DEMO.rows,hdr:DEMO.hdr});
  const wrong=R.filter(x=>x.got!==x.want||x.src!=='data');
  ok('הנתונים מזהים את חודש הבסיס בכל 12 החודשים, כולל דצמבר (k=0) ומעבר שנה',!wrong.length,
   wrong.length?wrong.map(x=>`${x.want}→${x.got}/${x.src}`).join(' · '):`${R.length} מקרים`);
  /* השנה אינה נגזרת מהחודש לבדו */
  const Y=await p.evaluate(({rows,hdr})=>{const H=k=>hdr.indexOf(k),MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
   const rr=rows.map(r=>{const x=r.slice();let s=+x[H('צר.החודש')]||0;for(let i=0;i<9;i++)s+=+x[H(MON[i])]||0;x[H('צר.השנה')]=s;
    x[H("תאר' מכירה")]='15.03.2025';x[H("תאר' כניסה")]='15.03.2025';return x});
   const r=bmResolve(rr,hdr,[],new Date(2026,10,1));return {src:r.src,ver:r.verified,w:r.warn.join(' | ')}},{rows:DEMO.rows,hdr:DEMO.hdr});
  ok('חודש מזוהה, אבל הפעילות האחרונה בשנה אחרת — השנה אינה מאומתת והחודש «לא מאומת»',
   Y.src==='none'&&!Y.ver&&/השנה לא אומתה/.test(Y.w),Y.src+' · '+Y.w.slice(0,90));
  /* ספי ביטחון: נתונים אקראיים (הדוח הסינתטי) — לא מזוהים; מעט פריטים — לא מזוהים */
  const T=await p.evaluate(({rows,hdr})=>{const a=bmResolve(rows,hdr,[],new Date(2026,10,1));
   const H=k=>hdr.indexOf(k),MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
   const few=rows.slice(0,20).map(r=>{const x=r.slice();let s=+x[H('צר.החודש')]||0;for(let i=0;i<9;i++)s+=+x[H(MON[i])]||0;x[H('צר.השנה')]=s;return x});
   const c=bmResolve(few,hdr,[],new Date(2026,10,1));
   return {a:{src:a.src,rate:a.data.rate},c:{src:c.src,why:c.data.why}}},{rows:DEMO.rows,hdr:DEMO.hdr});
  ok('נתונים שאינם מתיישבים (הדוח הסינתטי) — לא מזוהים',T.a.src==='none'&&T.a.rate<0.95,`${T.a.src} · ${Math.round(T.a.rate*100)}%`);
  ok('פחות מ-30 פריטים עם צריכה — לא מזוהה, גם כשהם מתיישבים',T.c.src==='none'&&/מעט מדי/.test(T.c.why||''),T.c.why);
  ok('אין שגיאות JS',!errs.length,errs[0]);await ctx.close()}

 /* ===== 2 · הכותרת ===== */
 {const {ctx,p,errs}=await page(b,FIXTURE_NOW);
  const run=(rowsK,head,today)=>p.evaluate(({rows,hdr,k,head,today})=>{const H=x=>hdr.indexOf(x),MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
   const rr=k==null?rows:rows.map(r=>{const x=r.slice();let s=+x[H('צר.החודש')]||0;for(let i=0;i<k;i++)s+=+x[H(MON[i])]||0;x[H('צר.השנה')]=s;return x});
   const r=bmResolve(rr,hdr,[head],new Date(today));
   return {ym:`${r.y}-${r.m}`,src:r.src,state:r.state?dsigDate(r.state):null,ref:dsigDate(r.ref),assumed:r.refAssumed,w:r.warn.join(' | ')}},
   {rows:DEMO.rows,hdr:DEMO.hdr,k:rowsK,head,today});
  /* הדוח הסינתטי: הנתונים אינם מכריעים, והכותרת אמינה */
  const a=await run(null,['דוח','תאריך מצב: 23.10.2026'],'2026-11-01T08:00:00');
  ok('נתונים לא מכריעים + תאריך אמין בכותרת → הכותרת; גיל מלאי מתאריך המצב',
   a.src==='header'&&a.ym==='2026-9'&&a.state==='2026-10-23'&&a.ref==='2026-10-23'&&!a.assumed,JSON.stringify(a));
  /* הנתונים מזהים ספטמבר; הכותרת מתאימה — תאריך המצב משמש */
  const c=await run(9,['דוח','תאריך מצב: 23.10.2026'],'2026-11-01T08:00:00');
  ok('נתונים וכותרת מסכימים → מהנתונים, ותאריך המצב מהכותרת',c.src==='data'&&c.state==='2026-10-23'&&!c.assumed,JSON.stringify(c));
  /* הנתונים מזהים ספטמבר; הכותרת אומרת נובמבר → הנתונים קובעים, אזהרה, ותאריך ההפקה אינו תאריך ייחוס */
  const d=await run(9,['דוח','תאריך מצב: 05.12.2026'],'2027-01-10T08:00:00');
  ok('כותרת שסותרת את הנתונים → הנתונים קובעים, אזהרה שמבחינה בין תאריך ההפקה לחודש הבסיס',
   d.src==='data'&&d.ym==='2026-9'&&d.state===null&&d.ref==='2026-09-30'&&d.assumed&&/תאריך ההפקה/.test(d.w)&&/חודש בסיס/.test(d.w),JSON.stringify(d).slice(0,240));
  const e=await run(null,['דוח','תאריך מצב: 23.10.2027'],'2026-11-01T08:00:00');
  ok('תאריך עתידי בכותרת — לא נעשה בו שימוש, ונאמר למה',e.src==='none'&&/בעתיד/.test(e.w),e.w.slice(0,100));
  const f=await run(null,['דוח','תאריך מצב: 01.10.2026'],'2026-11-01T08:00:00');
  ok('תאריך בכותרת שמוקדם מהפעילות האחרונה — לא נעשה בו שימוש',f.src==='none'&&/מוקדם מהפעילות/.test(f.w),f.w.slice(0,100));
  const g=await run(null,['דוח','תאריך מצב: 31.02.2026'],'2026-11-01T08:00:00');
  ok('«31.02.2026» אינו תאריך — לא נעשה בו שימוש',g.src==='none'&&/אינו תאריך תקין/.test(g.w),g.w.slice(0,100));
  ok('אין שגיאות JS',!errs.length,errs[0]);await ctx.close()}

 /* ===== 3 · אותו קובץ, ארבעה שעונים ===== */
 /* בסיס ספטמבר 2026 מזוהה מהנתונים, בלי תאריך מצב: גיל נספר מ-30.9.2026.
    הפעילות האחרונה 22.10.2026 — כלומר יש פעילות אחרי תאריך הייחוס. */
 const FD=variant({k:9,last:new Date(2026,9,22),mut:R=>{
  /* שני פריטים מכוונים: מכירה אחרי תאריך הייחוס, ותאריך שגוי מעבר לגבול */
  const a=R[0],c=R[1];a[0]='BM-AFTER';a[H("תאר' מכירה")]='14.10.2026';a[H("תאר' כניסה")]='02.10.2026';
  c[0]='BM-BEYOND';c[H("תאר' מכירה")]='15.12.2026'}});
 const snap=async(p)=>p.evaluate(()=>{
  const prim=v=>v===null||['string','number','boolean'].includes(typeof v);
  const R=v=>typeof v==='number'?(Number.isFinite(v)?Math.round(v*1e6)/1e6:String(v)):v;
  /* שדות ההיסטוריה (hSeen/hDays…) מחוץ להשוואה: הם תלויים במה שנטען קודם בדפדפן, לא בקובץ */
  const skip=new Set(['hSeen','hNew','hDays','hWorse','apStatus','apPrev','apDrift']);
  const items={};for(const r of ALL){const o={};for(const k of Object.keys(r).sort()){if(skip.has(k))continue;const v=r[k];
   if(prim(v))o[k]=R(v);else if(Array.isArray(v)&&v.every(prim))o[k]=v.map(R);
   else if(v&&typeof v==='object'&&v.date instanceof Date)o[k]=dsigDate(v.date)}
   items[r.pn]=o}
  const queues={};for(const k of Object.keys(QF))if(Array.isArray(QF[k]))queues[k]=QF[k].map(x=>x.pn);
  const doors={};for(const m of ['ropGap','ropThin'])doors[m]=ropPick(ropAll(),m).map(x=>x.pn);
  window.__cap=[];XLSX.writeFile=(wb,name)=>{window.__cap.push({name,sheets:wb.SheetNames.map(n=>[n,XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:null})])})};
  window.Worker=undefined;
  setMode('catalog');for(const k of Object.keys(colFilters))delete colFilters[k];apply();
  doExport();doSapMass();doSapMass('slow');
  const files=window.__cap.map(f=>({name:f.name.replace(/\d{4}-\d{2}-\d{2}/g,'<date>'),sheets:f.sheets}));
  return {base:`${BASE.y}-${BASE.m}/${BASE.src}/${dsigDate(BASE.ref)}`,items,queues,doors,files,
   rd:document.getElementById('rd').value}});
 const clocks=['2026-10-16T08:00:00+03:00','2026-10-31T08:00:00+02:00','2026-11-01T08:00:00+02:00','2027-01-15T08:00:00+02:00'];
 const S=[];
 for(const c of clocks){const {ctx,p,errs}=await page(b,c);await load(p,FD);S.push({c,s:await snap(p),errs});await ctx.close()}
 const J=JSON.stringify;
 ok('בכל ארבעת השעונים: אותו חודש בסיס, אותו מקור, אותו תאריך ייחוס',S.every(x=>x.s.base===S[0].s.base)&&S[0].s.base==='2026-9/data/2026-09-30',S.map(x=>x.s.base).join(' · '));
 {const diffs=[];for(const x of S.slice(1)){for(const pn of Object.keys(S[0].s.items)){const A=S[0].s.items[pn],B=x.s.items[pn];
   if(J(A)!==J(B)){const ks=Object.keys(A).filter(k=>J(A[k])!==J((B||{})[k]));diffs.push(`${x.c.slice(0,10)} ${pn}: ${ks.join(',')}`)}}}
  ok('כל שדה של כל פריט זהה — כולל גילי מלאי, סיווג, ROP/SS',!diffs.length,diffs.slice(0,4).join(' · ')||`${Object.keys(S[0].s.items).length} פריטים × ${S.length}`)}
 ok('כל תור וכל דלת — אותם פריטים באותו סדר',S.every(x=>J(x.s.queues)===J(S[0].s.queues)&&J(x.s.doors)===J(S[0].s.doors)),
  Object.entries(S[0].s.queues).map(([k,v])=>k+':'+v.length).join(' '));
 {const names=S[0].s.files.map(f=>f.name);
  ok('נכתבו ייצוא לאקסל וקובץ SAP',names.length>=2&&/^Inventory/.test(names[0])&&names.some(n=>/^SAP_Params/.test(n)),names.join(' · '));
  const fd=S.slice(1).filter(x=>J(x.s.files)!==J(S[0].s.files)).map(x=>{
   for(let i=0;i<S[0].s.files.length;i++){const A=S[0].s.files[i],B=x.s.files[i]||{sheets:[]};
    for(let j=0;j<A.sheets.length;j++){const a=A.sheets[j],bb=B.sheets[j]||[,[]];
     for(let r=0;r<a[1].length;r++)if(J(a[1][r])!==J(bb[1][r]))return `${x.c.slice(0,10)} ${A.name}/${a[0]} שורה ${r+1}`}}return x.c});
  ok('ייצוא לאקסל, SAP ומלאי איטי — זהים תא בתא בכל השעונים',!fd.length,fd.slice(0,3).join(' · '))}
 ok('אין שגיאות JS בארבעת השעונים',S.every(x=>!x.errs.length),S.map(x=>x.errs[0]).filter(Boolean)[0]);

 /* ===== 4 · פעילות אחרי תאריך הייחוס ===== */
 {const it=S[0].s.items['BM-AFTER'],bey=S[0].s.items['BM-BEYOND'];
  ok('מכירה ב-14.10 בדוח שבסיסו ספטמבר: גיל 0, מסומנת, ואינה «עתידית»',
   it&&it.saleAge===0&&it.saleAfterRef===true&&it.saleFuture===false&&it.entAfterRef===true&&it.entAge===0,J(it&&{saleAge:it.saleAge,saleAfterRef:it.saleAfterRef,saleFuture:it.saleFuture,entAge:it.entAge}));
  ok('ואינה מסווגת כמלאי מת או איטי בגלל הגיל',it&&!it.dead3y&&it.ageTier!=='dead'&&it.cat!=='מלאי מת',it&&`${it.ageTier} · ${it.cat}`);
  ok('תאריך שמעבר לחודש שאחרי הבסיס (15.12) — שגוי, כמו קודם',bey&&bey.saleFuture===true&&bey.saleAge===null,bey&&`${bey.saleFuture}/${bey.saleAge}`);
  const n=Object.values(S[0].s.items).filter(x=>x.saleAfterRef).length;
  const plain=Object.values(S[0].s.items).filter(x=>!x.saleAfterRef&&!x.entAfterRef&&('saleAfterRef' in x||'entAfterRef' in x)).length;
  ok('השדות קיימים רק בפריטים שבהם זה קרה',n>1&&plain===0,`${n} פריטים עם מכירה אחרי תאריך הייחוס`)}
 {const {ctx,p}=await page(b,'2026-11-01T08:00:00+02:00');await load(p,FD);
  const t=await p.evaluate(()=>{const r=ALL.find(x=>x.pn==='BM-AFTER');detail(r);
   const tx=(document.getElementById('detail')||document.body).innerText;
   return {box:/אחרי תאריך הייחוס/.test(tx),z:document.getElementById('zchip').title,info:document.getElementById('bmInfo').textContent}});
  ok('כרטיס הפריט אומר שהמכירה אחרי תאריך הייחוס',t.box);
  ok('ההסבר בשבב: חודש הבסיס, המקור, תאריך הייחוס כהנחה, והפעילות האחרונה שאינה תאריך הדוח',
   /חודש הבסיס: ספטמבר 2026 — מהנתונים/.test(t.z)&&/הנחת חישוב: סוף חודש הבסיס/.test(t.z)&&/זו אינה תאריך הדוח/.test(t.z)&&/אחרי תאריך הייחוס/.test(t.z),t.z.split('\n').slice(0,5).join(' | '));
  ok('בתפריט הדוח: גיל מלאי נספר מ-30.9.2026, ושזו הנחה',/30\.9\.2026/.test(t.info)&&/הנחה/.test(t.info),t.info);
  await ctx.close()}

 /* ===== 5 · חודש לא מאומת ===== */
 const FU=variant({});                          /* הדוח הסינתטי בלי תאריך — הנתונים אינם מכריעים */
 {const {ctx,p,errs}=await page(b,'2026-11-01T08:00:00+02:00');
  await load(p,FU);
  const u=await p.evaluate(()=>{window.__cap=[];XLSX.writeFile=(wb,name)=>window.__cap.push(name);
   setMode('catalog');apply();doSapMass();doSapMass('slow');sapPreview();
   const sb=document.getElementById('sapXls');
   return {src:BASE.src,ver:BASE.verified,z:document.getElementById('zchip').textContent,fw:document.getElementById('fchip').classList.contains('warn'),
    ok:!document.getElementById('bmOk').hidden,files:window.__cap.length,dis:sb&&sb.getAttribute('aria-disabled'),title:sb&&sb.title,
    hist:localStorage.getItem('planner_history_v1'),par:localStorage.getItem('planner_params_v1'),prof:localStorage.getItem('planner_profile_v1'),
    runs:HIST.runs.length,hrun:HRUN,warn:DIAG.warn.join(' | '),hSeen:ALL.every(x=>x.hSeen===0)}});
  ok('לא מאומת: השבב בסרגל העליון אומר זאת, בטקסט ובצבע',!u.ver&&/חודש בסיס לא מאומת/.test(u.z)&&u.fw,u.z);
  ok('כפתור «אשר את חודש הבסיס» מוצג',u.ok);
  ok('אין קובץ SAP — לא עדכון, לא מלאי איטי, לא דרך התצוגה המקדימה',u.files===0,`${u.files} קבצים`);
  ok('וכפתור ה-SAP מסומן כלא זמין ואומר למה',u.dis==='true'&&/חודש הבסיס של הדוח לא אומת/.test(u.title||''),u.title);
  ok('לא נרשמה הרצה, לא מעקב יישום ולא פרופיל',u.hist===null&&u.par===null&&u.prof===null&&u.runs===0&&u.hrun===null,
   `hist=${!!u.hist} par=${!!u.par} prof=${!!u.prof} runs=${u.runs}`);
  ok('האזהרה אומרת מה לא קורה ואיך משחררים',/לא נרשמת הרצה/.test(u.warn)&&/אין קובץ SAP/.test(u.warn),u.warn.slice(0,120));
  /* אישור */
  await p.click('#fchip');await p.click('#bmOk');await p.waitForTimeout(300);
  const v=await p.evaluate(()=>{window.__cap=[];setMode('catalog');apply();doSapMass();
   const st=JSON.parse(localStorage.getItem('planner_basemonth_v1')||'{}');
   return {src:BASE.src,ver:BASE.verified,ym:`${BASE.y}-${BASE.m}`,files:window.__cap.length,runs:HIST.runs.length,
    hist:!!localStorage.getItem('planner_history_v1'),par:!!localStorage.getItem('planner_params_v1'),
    stored:st[BASE.cs]&&st[BASE.cs].ym,z:document.getElementById('zchip').textContent,dis:document.getElementById('sapXls').getAttribute('aria-disabled')}});
  ok('אחרי אישור: מאומת, ידני, אותו חודש',v.ver&&v.src==='manual'&&v.ym==='2026-10',`${v.src} ${v.ym}`);
  ok('ההרצה נרשמה פעם אחת, וגם מעקב היישום',v.runs===1&&v.hist&&v.par,`runs=${v.runs}`);
  ok('קובץ SAP יוצא, והכפתור זמין',v.files===1&&v.dis==='false',`${v.files} · ${v.dis}`);
  ok('הבחירה נשמרה לתוכן הקובץ',v.stored==='2026-10',v.stored);
  ok('השבב אומר «נבחר ידנית»',/בסיס אוקטובר 2026 · נבחר ידנית/.test(v.z),v.z);
  /* ===== 6 · בחירה ידנית: נשמרת לתוכן, לא לקובץ אחר, ואינה יוצרת הרצה שנייה ===== */
  await load(p,FU);
  const r1=await p.evaluate(()=>({src:BASE.src,ym:`${BASE.y}-${BASE.m}`,runs:HIST.runs.length,rep:HRUN&&HRUN.repeat}));
  ok('אותו קובץ שוב — הבחירה הידנית חוזרת, ואין הרצה חדשה',r1.src==='manual'&&r1.ym==='2026-10'&&r1.runs===1&&r1.rep===true,JSON.stringify(r1));
  const FU2=variant({mut:R=>{R[5][H('מלאי פנוי')]=(+R[5][H('מלאי פנוי')]||0)+1}});
  await load(p,FU2);
  const r2=await p.evaluate(()=>({src:BASE.src,runs:HIST.runs.length}));
  ok('קובץ אחר (תא אחד שונה) — הבחירה אינה חלה עליו, והוא לא מאומת',r2.src==='none',JSON.stringify(r2));
  await load(p,FU);
  await p.evaluate(()=>{const rd=document.getElementById('rd');rd.value='2026-08';rd.onchange({target:rd})});await p.waitForTimeout(300);
  const r3=await p.evaluate(()=>({src:BASE.src,ym:`${BASE.y}-${BASE.m}`,runs:HIST.runs.length,last:HIST.runs[HIST.runs.length-1].ym,
   stored:JSON.parse(localStorage.getItem('planner_basemonth_v1'))[BASE.cs].ym,lab:LABELS[0].name}));
  ok('שינוי חודש לאותו תוכן — הרשומה מתעדכנת לחודש החדש, בלי הרצה נוספת',r3.src==='manual'&&r3.ym==='2026-8'&&r3.last==='2026-8'&&r3.stored==='2026-08',JSON.stringify(r3));
  ok('ושמות החודשים מתעדכנים',r3.lab==='אוגוסט',r3.lab);
  ok('אין שגיאות JS',!errs.length,errs[0]);await ctx.close()}

 /* ===== 6ב · אותו קובץ בשני חודשים — הרצה אחת ===== */
 {const {ctx,p}=await page(b,'2026-10-20T08:00:00+03:00');await load(p,FD);
  await p.clock.setFixedTime(new Date('2026-11-03T08:00:00+02:00'));await load(p,FD);
  const h=await p.evaluate(()=>({runs:HIST.runs.length,rep:HRUN&&HRUN.repeat,ym:HIST.runs[0].ym}));
  ok('אותו קובץ באוקטובר ובנובמבר — הרצה אחת, «אותו דוח», באותו חודש בסיס',h.runs===1&&h.rep===true&&h.ym==='2026-9',JSON.stringify(h));
  await ctx.close()}

 /* ===== 7 · היום נשאר לפעולות תפעוליות ===== */
 {const {ctx,p}=await page(b,'2027-01-15T08:00:00+02:00');await load(p,FD);
  const t=await p.evaluate(()=>({eta:etaToday(),ref:dsigDate(bmRef())}));
  ok('ETA נמדד מהיום (15.1.2027), גיל מלאי מתאריך הייחוס (30.9.2026)',t.eta==='2027-01-15'&&t.ref==='2026-09-30',JSON.stringify(t));
  await ctx.close()}

 /* ===== 8 · ההדגמה שבתוך הכלי ===== */
 {const {ctx,p,errs}=await page(b,'2026-11-01T08:00:00+02:00');
  await p.evaluate(()=>{window.__z=null});await p.evaluate(()=>demoStart());
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB&&ZN,null,{timeout:60000});
  const d=await p.evaluate(()=>({ver:BASE.verified,src:BASE.src,ym:`${BASE.y}-${BASE.m}`,state:BASE.state&&dsigDate(BASE.state),warn:BASE.warn}));
  ok('ההדגמה מאומתת: חודש הבסיס אוקטובר 2026, תאריך המצב היום',d.ver&&d.ym==='2026-10'&&d.state==='2026-11-01'&&!d.warn.length,JSON.stringify(d));
  ok('אין שגיאות JS',!errs.length,errs[0]);await ctx.close()}

 /* ===== 9 · היסטוריית הפרופיל: קובץ אחד, חודש אחד ===== */
 const HD=['דוח ZMRP — תכנון מלאי','תאריך מצב: 23.10.2026'];
 const FA=variant({head:HD});                                   /* ספטמבר, מהכותרת */
 const FB=variant({head:HD,mut:R=>{for(const r of R.slice(0,300))for(let i=5;i<11;i++)r[H(MON[i])]=0}});
 const PH=p=>p.evaluate(()=>{const o={},src=PROFH['\u0000src']||{};
  for(const [pn,h] of Object.entries(PROFH))if(pn!=='\u0000src')o[pn]=h;
  const by={};for(const h of Object.values(o))for(const k of Object.keys(h))by[k]=(by[k]||0)+1;
  return {items:o,by,src:JSON.parse(JSON.stringify(src)),cs:BASE&&BASE.cs}});
 const pick=(P,ym)=>Object.fromEntries(Object.entries(P.items).filter(([k,h])=>h[ym]!=null).map(([k,h])=>[k,h[ym]]));
 const choose=async(p,ym)=>{await p.evaluate(v=>{const rd=document.getElementById('rd');rd.value=v;rd.onchange({target:rd})},ym);await p.waitForTimeout(500)};
 {const {ctx,p,errs}=await page(b,'2026-11-01T08:00:00+02:00');
  await load(p,FA);await p.waitForTimeout(400);
  /* סימון משתמש — לא אמור לזוז בשום שלב */
  await p.evaluate(()=>{setMarks([currentRows()[0]],'done','בדיקת היסטוריה');window.__mk=JSON.stringify(MARKS);window.__sug=JSON.stringify(ALL.map(x=>[x.pn,x.sugROP,x.sugSS]))});
  const p0=await PH(p);
  ok('אחרי טעינה: רשום רק תחת ספטמבר, והכותב הוא תוכן הקובץ',Object.keys(p0.by).join()==='2026-9'&&!!(p0.src['2026-9']&&p0.src['2026-9'].cs===p0.cs),JSON.stringify(p0.by));
  await choose(p,'2026-08');const p1=await PH(p);
  ok('ספטמבר → אוגוסט: רשום רק תחת אוגוסט, ואין כפילות',Object.keys(p1.by).join()==='2026-8'&&!p1.src['2026-9'],JSON.stringify(p1.by));
  await choose(p,'2026-09');const p2=await PH(p);
  ok('אוגוסט → ספטמבר: רשום רק תחת ספטמבר — ובדיוק כמו בטעינה',Object.keys(p2.by).join()==='2026-9'&&JSON.stringify(pick(p2,'2026-9'))===JSON.stringify(pick(p0,'2026-9')),JSON.stringify(p2.by));
  const keep=await p.evaluate(()=>({mk:JSON.stringify(MARKS)===window.__mk,sug:JSON.stringify(ALL.map(x=>[x.pn,x.sugROP,x.sugSS]))===window.__sug,
   stored:JSON.stringify(JSON.parse(localStorage.getItem('planner_profile_v1')))===JSON.stringify(PROFH)}));
  ok('הסימונים לא זזו, ו-ROP/SS חזרו בדיוק אחרי ספטמבר → אוגוסט → ספטמבר',keep.mk&&keep.sug,JSON.stringify(keep));
  ok('מה שבזיכרון הוא מה שנשמר',keep.stored);
  ok('אין שגיאות JS',!errs.length,errs[0]);await ctx.close()}

 /* ===== 9ב · שני קבצים באותו חודש, ואחד מהם עובר חודש ===== */
 {const {ctx,p,errs}=await page(b,'2026-11-01T08:00:00+02:00');
  await load(p,FA);await p.waitForTimeout(400);const a=await PH(p);const aSep=pick(a,'2026-9');
  await load(p,FB);await p.waitForTimeout(400);const bb=await PH(p);const bSep=pick(bb,'2026-9');
  const differ=Object.keys(aSep).filter(k=>aSep[k]!==bSep[k]).length+Object.keys(bSep).filter(k=>!(k in aSep)).length;
  ok('שני הקבצים שונים בהמלצות של ספטמבר (אחרת הבדיקה אינה בודקת דבר)',differ>0,`${differ} פריטים שונים`);
  ok('הקובץ השני נרשם לספטמבר, ומה שדרס של הראשון נשמר בצד',!!(bb.src['2026-9']&&bb.src['2026-9'].cs===bb.cs&&bb.src['2026-9'].prev&&bb.src['2026-9'].prev.cs===a.cs));
  await choose(p,'2026-08');const c=await PH(p);
  const cSep=pick(c,'2026-9');
  const wrong=Object.entries(cSep).filter(([k,v])=>aSep[k]!==v).length;
  const backB=Object.keys(bSep).filter(k=>!(k in aSep)&&(k in cSep)).length;
  ok('הקובץ השני עבר לאוגוסט: ספטמבר מחזיק את ההמלצות של הקובץ הראשון, לא של השני',Object.keys(cSep).length>0&&wrong===0&&backB===0,`${Object.keys(cSep).length} פריטים · שגויים ${wrong}`);
  ok('והכותב של ספטמבר חזר להיות הקובץ הראשון; אוגוסט שייך לשני',c.src['2026-9']&&c.src['2026-9'].cs===a.cs&&c.src['2026-8']&&c.src['2026-8'].cs===c.cs);
  ok('אין שגיאות JS',!errs.length,errs[0]);await ctx.close()}

 /* ===== 9ג · רישום ישן, שכותבו לא ידוע, אינו נמחק ===== */
 {let pn,OLDV;{const {ctx,p}=await page(b,'2026-11-01T08:00:00+02:00');await load(p,FA);await p.waitForTimeout(400);
   const m=pick(await PH(p),'2026-9');pn=Object.keys(m)[0];
   /* ערך ישן ששונה ממה שהקובץ עצמו כותב — אחרת «חזר» לא מוכיח דבר */
   OLDV=m[pn]==='A'?'B':'A';await ctx.close()}
  /* מצב «לפני התיקון»: רישום של ספטמבר בלי רישום מקור, בדפדפן נקי */
  const {ctx,p}=await page(b,'2026-11-01T08:00:00+02:00');
  await p.evaluate(a=>localStorage.setItem('planner_profile_v1',JSON.stringify({[a[0]]:{'2026-9':a[1]}})),[pn,OLDV]);
  await load(p,FA);await p.waitForTimeout(400);const l1=await PH(p);
  ok('רישום ישן בספטמבר — נשמר בצד כשהקובץ נרשם על אותו חודש (כותב לא ידוע)',!!(l1.src['2026-9']&&l1.src['2026-9'].prev&&l1.src['2026-9'].prev.cs===null&&l1.src['2026-9'].prev.v[pn]===OLDV));
  await choose(p,'2026-08');const l2=await PH(p);
  ok('והקובץ עבר לאוגוסט — הרישום הישן חזר לספטמבר כמו שהיה',!!(l2.items[pn]&&l2.items[pn]['2026-9']===OLDV),JSON.stringify(l2.items[pn]));
  await ctx.close()}

 /* ===== 9ד · שתי לשוניות לא דורסות זו את זו ===== */
 {const ctx=await b.newContext({viewport:{width:1440,height:900},timezoneId:'Asia/Jerusalem'});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const t1=await ctx.newPage(),t2=await ctx.newPage();
  for(const t of [t1,t2]){await t.clock.setFixedTime(new Date('2026-11-01T08:00:00+02:00'));await t.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1')}
  await Promise.all([load(t1,FA),load(t2,FB)]);await t1.waitForTimeout(800);
  const st=await t1.evaluate(()=>JSON.parse(localStorage.getItem('planner_profile_v1')||'{}')['\u0000src']);
  const csA=await t1.evaluate(()=>BASE.cs),csB=await t2.evaluate(()=>BASE.cs);
  const e=st&&st['2026-9'];
  ok('שתי לשוניות, שני קבצים באותו חודש: הכותב האחרון רשום, ומה שדרס של השני שמור — שום רישום לא אבד',
   !!e&&((e.cs===csA&&e.prev&&e.prev.cs===csB)||(e.cs===csB&&e.prev&&e.prev.cs===csA)),JSON.stringify(e&&{cs:e.cs===csA?'A':e.cs===csB?'B':'?',prev:e.prev&&(e.prev.cs===csA?'A':e.prev.cs===csB?'B':'?')}));
  /* הלשונית של הכותב האחרון עוברת חודש — השנייה מקבלת את הרישום שלה בחזרה */
  const last=e&&e.cs===csA?t1:t2,other=last===t1?csB:csA;
  await choose(last,'2026-08');await last.waitForTimeout(500);
  const st2=await t1.evaluate(()=>JSON.parse(localStorage.getItem('planner_profile_v1')||'{}')['\u0000src']);
  ok('הלשונית שעברה חודש מחזירה את ספטמבר ללשונית השנייה',st2&&st2['2026-9']&&st2['2026-9'].cs===other&&st2['2026-8']);
  await ctx.close()}

 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
