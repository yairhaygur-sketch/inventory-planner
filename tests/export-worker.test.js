/* ============ L3 · ייצוא Excel ב-Worker — dense, דחוס, תמונת מצב ============
   הייצוא לאקסל נכתב ב-Worker, בחלקים, מתמונת מצב שנלקחת ברגע הלחיצה.
   נבדק כאן, במחשב ובטלפון, על קובץ של 20,000 שורות שנבנה מהדוח הסינתטי:
     · הקובץ שיורד (Worker) זהה תא בתא — שמות גיליונות, סדר, ערכים, סוגי
       תאים, רוחבי עמודות — לספר שנבנה בדף, ולקובץ של main (אם זמין)
     · תמונת מצב: סינון ומעבר מסך בזמן הייצוא אינם משנים את הקובץ
     · ייצוא נוסף בזמן ייצוא — לא מתחיל, ואומר למה; יורד קובץ אחד
     · ביטול — אין קובץ, שום דבר לא משתנה; טעינת דוח מבטלת ייצוא שבהכנה
     · Worker שנכשל / אין Worker — הקובץ נכתב בדף, זהה
     · הדף מגיב בזמן הייצוא (מול כתיבה בדף על אותו קובץ)
     · אחרי הייצוא — הנתונים לא זזו והכלי עובד; ייצוא שלא הסתיים — הודעה
   EXPORT_XLSX=<דוח> — על דוח אחר (למשל האמיתי, מקומית). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path'),os=require('os');
const XLSX=require('xlsx');
const SD=__dirname;
const {resolveBase,short}=require('./lib/export-base.js');
const {cellDiffs,validateDecl,checkDeclared,fmt:fmtDiff}=require('./lib/export-expected.js');
const out=[];let bad=0,skip=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
const TMP=fs.mkdtempSync(path.join(os.tmpdir(),'xw-'));
function bigFile(n){const wb=XLSX.readFile(path.join(SD,'zmrp-demo.xlsx'),{cellDates:true});const nm=wb.SheetNames[0];
 const a=XLSX.utils.sheet_to_json(wb.Sheets[nm],{header:1,blankrows:false,defval:''});
 const h=a.findIndex(r=>r.some(c=>/מק"ט מוביל|מק״ט מוביל/.test(String(c))));const pn=a[h].findIndex(c=>/מק"ט מוביל|מק״ט מוביל/.test(String(c)));
 const body=a.slice(h+1).filter(r=>String(r[pn]).trim());const rows=[];
 for(let i=0;i<n;i++){const k=Math.floor(i/body.length),r=body[i%body.length].slice();if(k)r[pn]=String(r[pn])+'-k'+k;rows.push(r)}
 const w2=XLSX.utils.book_new();XLSX.utils.book_append_sheet(w2,XLSX.utils.aoa_to_sheet(a.slice(0,h+1).concat(rows),{cellDates:true}),nm);
 const f=path.join(TMP,`big-${n}.xlsx`);XLSX.writeFile(w2,f,{compression:true});return f}
const XL=process.env.EXPORT_XLSX||bigFile(20000);
/* תוכן קובץ להשוואה: שמות גיליונות בסדר, לכל גיליון — כל התאים (ערך+סוג) ורוחבי העמודות */
const sig=wb=>JSON.stringify(wb.SheetNames.map(n=>{const ws=wb.Sheets[n];
 const cells=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
 const types=Object.keys(ws).filter(k=>k[0]!=='!').sort().map(k=>k+':'+ws[k].t).join(',');
 return [n,cells,types,(ws['!cols']||[]).map(c=>c&&c.wch)]}));
const fileSig=f=>sig(XLSX.readFile(f,{cellDates:true}));
/* המקום הראשון שבו שני קבצים נבדלים: גיליון ותא — כדי שכשל יגיד מה זז */
function firstDiff(fa,fb){const A=XLSX.readFile(fa,{cellDates:true}),B=XLSX.readFile(fb,{cellDates:true});
 if(A.SheetNames.join('|')!==B.SheetNames.join('|'))return `גיליונות: ${A.SheetNames.join(',')} → ${B.SheetNames.join(',')}`;
 for(const n of A.SheetNames){const a=A.Sheets[n],b=B.Sheets[n];
  const ks=[...new Set([...Object.keys(a),...Object.keys(b)])].filter(k=>k[0]!=='!').sort();
  for(const k of ks){const x=a[k],y=b[k];
   if(!x||!y||x.t!==y.t||String(x.v)!==String(y.v))return `${n}!${k}: «${x?x.v:'∅'}» → «${y?y.v:'∅'}»`}
  const wa=(a['!cols']||[]).map(c=>c&&c.wch).join(','),wb=(b['!cols']||[]).map(c=>c&&c.wch).join(',');
  if(wa!==wb)return `${n}: רוחבי עמודות`}
 return ''}
const sheetsOf=f=>{const wb=XLSX.readFile(f);return wb.SheetNames.map(n=>`${n}:${XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1}).length}`).join(' · ')};
const HTML='file://'+path.join(SD,'..','index.html')+'?nobrief=1';
/* BASE_MONTH: השעון של הדף מקובע לתאריך המצב של הדוח הסינתטי. ה-main
   שלפני BASE_MONTH סופר גיל מלאי מהיום, והגרסה הנוכחית מתאריך מצב הדוח —
   בלי שעון משותף, ההשוואה תא-בתא הייתה מודדת את יום ההרצה ולא את הקוד, וכל
   יום הייתה נותנת main אחר. ההשוואה עצמה לא השתנתה: כל תא, כל גיליון. */
const {FIXTURE_NOW}=require('./lib/fixture-date.js');
async function open(b,W,H,html){const ctx=await b.newContext({viewport:{width:W,height:H},acceptDownloads:true});
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.clock.setFixedTime(FIXTURE_NOW);
 await p.goto(html||HTML);await p.waitForTimeout(500);
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof LOADMS!=='undefined'&&LOADMS>0&&(typeof LOADJOB==='undefined'||!LOADJOB),null,{timeout:300000,polling:200});
 await p.evaluate(()=>{setMode('catalog');for(const k of Object.keys(colFilters))delete colFilters[k];apply()});await p.waitForTimeout(300);
 return {ctx,p,errs}}
/* ייצוא עם הורדה אמיתית; מחזיר נתיב לקובץ (או null), וכמה הורדות היו */
async function exp(p,name,during){const files=[];const on=d=>files.push(d);p.on('download',on);
 await p.evaluate(()=>{doExport()});
 if(during)await during();
 await p.waitForFunction(()=>typeof XJOB==='undefined'||!XJOB,null,{timeout:300000,polling:100});
 const t=Date.now();while(!files.length&&Date.now()-t<5000)await p.waitForTimeout(100);
 await p.waitForTimeout(600);p.off('download',on);
 if(!files.length)return {f:null,n:0};
 const f=path.join(TMP,name+'.xlsx');await files[0].saveAs(f);return {f,n:files.length}}
const SNAP=()=>{const F=['pn','cat','sev','sugSS','sugROP','ss','rop','orderQty','expCap','miss'];
 return JSON.stringify(ALL.map(r=>F.map(f=>r[f])))+'|'+Object.keys(Q).map(k=>k+':'+Q[k].length).join(',')};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 /* ---------------- מחשב ---------------- */
 {const {ctx,p,errs}=await open(b,1440,900);
  const snap0=await p.evaluate(SNAP);
  /* 1 — הורדה דרך ה-Worker, והדף מגיב */
  await p.evaluate(()=>{window.__mx=0;let last=performance.now();(function t(){const n=performance.now();if(window.__xon)window.__mx=Math.max(window.__mx,n-last);last=n;setTimeout(t,16)})();window.__xon=true});
  const w1=await exp(p,'worker');
  const stallW=await p.evaluate(()=>{window.__xon=false;return Math.round(__mx)});
  ok('מחשב · הקובץ ירד דרך ה-Worker',!!w1.f&&w1.n===1,w1.f?sheetsOf(w1.f):'אין קובץ');
  /* 2 — אותו ייצוא כשהספר נבנה בדף (XLSX.writeFile מוחלף, כמו בשומר העסקי) */
  await p.evaluate(()=>{window.__wb=null;window.__xf=XLSX.writeFile;XLSX.writeFile=(wb,name,opt)=>{window.__wb=XLSX.write(wb,{type:'base64',bookType:'xlsx',compression:true});window.__opt=opt}});
  await p.evaluate(()=>doExport());
  const h=await p.evaluate(()=>({b64:__wb,opt:__opt,box:document.getElementById('loadbox').hidden}));
  await p.evaluate(()=>{XLSX.writeFile=__xf});
  const hf=path.join(TMP,'harness.xlsx');fs.writeFileSync(hf,Buffer.from(h.b64||'','base64'));
  ok('מחשב · הקובץ מה-Worker זהה תא בתא לספר שנבנה בדף (גיליונות, סדר, ערכים, סוגים, רוחבים)',!!w1.f&&fileSig(w1.f)===fileSig(hf));
  ok('מחשב · XLSX.writeFile שהוחלף מקבל את הספר בדף, סינכרוני — כמו קודם, עם דחיסה',!!h.b64&&h.opt&&h.opt.compression===true&&h.box);
  /* 3 — מול main: אותו ייצוא בגרסת הבסיס, תא בתא, והקובץ לא גדל.
     «קטן יותר» הוכיח את הדחיסה של L3 מול main שלפניו; מאז main כבר דחוס,
     ולכן התנאי הוא «לא גדול יותר». קובץ שגדל (למשל דחיסה שכבתה) נכשל,
     וכך גם כל שינוי בתאים, בסוגים או ברוחבים (fileSig).
     PR-A: עד כאן הבסיס נלקח מ-`git show origin/main`, וכל כשל — גם זה
     שב-CI קרה תמיד, וגם קריסה של הדף של main — נבלע ב-catch ונרשם כ-PASS.
     עכשיו resolveBase() מחליט (ראה tests/lib/export-base.js): compare —
     ההשוואה רצה, וכל כשל בה הוא FAIL עם השגיאה; skip — רק בשלושה מקרים,
     בשורת SKIP נפרדת; error — ההשוואה נדרשת ואי אפשר לבצע אותה: FAIL. */
  const base=resolveBase();
  if(base.kind==='skip'){skip++;out.push('SKIP · מחשב · ההשוואה ל-main לא בוצעה — '+base.reason)}
  else if(base.kind==='error')ok('מחשב · ההשוואה ל-main נדרשת — ולא ניתן לבצע אותה',false,base.reason);
  else{const who=`${base.src}${base.sha?' '+short(base.sha):''}`;
   let mainF=null,why='',m=null;
   try{const MF=path.join(TMP,'main.html');fs.writeFileSync(MF,base.html);
    m=await open(b,1440,900,'file://'+MF+'?nobrief=1');
    const files=[];m.p.on('download',d=>files.push(d));await m.p.evaluate(()=>doExport());
    const t=Date.now();while(!files.length&&Date.now()-t<120000)await m.p.waitForTimeout(200);
    if(files.length){mainF=path.join(TMP,'main.xlsx');await files[0].saveAs(mainF)}
    else why='הייצוא של הבסיס לא ירד בתוך 120 שניות';
    if(m.errs.length)why=(why?why+' · ':'')+'שגיאות JS בבסיס: '+m.errs.slice(0,2).join(' | ')}
   catch(e){why='הדף של הבסיס נכשל: '+String(e&&e.message||e).split('\n')[0]}
   finally{if(m)await m.ctx.close().catch(()=>{})}
   if(!mainF||why)ok(`מחשב · ההשוואה מול ${who} נכשלה בביצוע`,false,why);
   else{const sm=fs.statSync(mainF).size,sn=fs.statSync(w1.f).size,same=fileSig(mainF)===fileSig(w1.f);
    /* כשההבדל קטן מ-KB מעוגל — בבתים, אחרת «גדל» ליד שני מספרים זהים */
    const kb=Math.round(sm/1024)===Math.round(sn/1024)&&sm!==sn?`${sm.toLocaleString('he-IL')} → ${sn.toLocaleString('he-IL')} בתים`:`${Math.round(sm/1024)}KB → ${Math.round(sn/1024)}KB`;
    /* PR-B: ההבדל נבדק מול tests/export-expected.json. ברירת המחדל — הצהרה
       ריקה — היא בדיוק התנאי של PR-A: זהה תא בתא, והקובץ לא גדל. הצהרה
       לא ריקה מתירה רק את התאים שבה, מול ה-SHA שבה, ולא אחד יותר. */
    const decl=JSON.parse(fs.readFileSync(path.join(SD,'export-expected.json'),'utf8'));
    if(same){const de=validateDecl(decl);
     ok(`מחשב · ההשוואה בוצעה מול ${who} — זהה תא בתא, והקובץ ${sn<=sm?'אינו גדול יותר':'גדל'}`,
      sn<=sm&&!decl.changes.length&&!de.length,decl.changes.length?`${kb} · אבל tests/export-expected.json מצהיר על ${decl.changes.length} שינויים שלא קרו`:de.length?de.join(' · '):kb)}
    else{const r=checkDeclared(cellDiffs(XLSX.readFile(mainF,{cellDates:true}),XLSX.readFile(w1.f,{cellDates:true})),decl,base.sha||'');
     /* טיוטה להצהרה — בדיוק מה שזז, מחוץ לריפו. היא אינה היתר: מישהו צריך
        להעתיק אותה ל-tests/export-expected.json ולהסביר כל שורה ב-PR. */
     if(!r.ok&&!r.other.length&&base.sha){const cand=path.join(os.tmpdir(),`export-expected.candidate-${process.pid}.json`);
      fs.writeFileSync(cand,JSON.stringify({base:base.sha,changes:cellDiffs(XLSX.readFile(mainF,{cellDates:true}),XLSX.readFile(w1.f,{cellDates:true})).declarable},null,1));
      out.push('INFO · טיוטת הצהרה לכל ההבדלים שנמצאו (לבדיקה ידנית, לא להעתקה עיוורת): '+cand)}
     const why=[...r.errs,...r.other.slice(0,3),...r.undeclared.slice(0,3).map(d=>'לא מוצהר: '+fmtDiff(d)),...r.unused.slice(0,3).map(d=>'מוצהר ולא קרה: '+fmtDiff(d))];
     ok(`מחשב · ההשוואה בוצעה מול ${who} — ${r.ok?`השתנו בדיוק ${r.matched} התאים המוצהרים`:`תוכן שונה — ${decl.changes.length?'ההבדל אינו תואם את ההצהרה':'ואין הצהרה'}`}`,r.ok,
      r.ok?`${kb}`:`${firstDiff(mainF,w1.f)} · ${kb} · `+`${r.errs.length} שגיאות הצהרה · ${r.other.length} שינויים שאי אפשר להצהיר · ${r.undeclared.length} לא מוצהרים · ${r.unused.length} מוצהרים שלא קרו — ${why.join(' | ')}`)}}}
  /* 4 — תמונת מצב: סינון ומעבר מסך באמצע הייצוא לא משנים את הקובץ */
  const w2=await exp(p,'snapshot',async()=>{
   await p.waitForFunction(()=>!!XJOB,null,{timeout:10000}).catch(()=>{});
   await p.evaluate(()=>{const k=FDEF[0][0];SEL[k].add(ALL[0][k]);apply();setMode('month');setQuery('צינור')})});
  ok('מחשב · סינון, חיפוש ומעבר מסך בזמן הייצוא — הקובץ זהה לזה של רגע הלחיצה',!!w2.f&&fileSig(w2.f)===fileSig(w1.f));
  await p.evaluate(()=>{clearQuery();for(const [k] of FDEF)SEL[k].clear();apply();setMode('catalog')});await p.waitForTimeout(300);
  /* 5 — ייצוא נוסף בזמן ייצוא */
  const files5=[];const on5=d=>files5.push(d);p.on('download',on5);
  await p.evaluate(()=>doExport());await p.waitForFunction(()=>!!XJOB,null,{timeout:10000}).catch(()=>{});
  const t5=await p.evaluate(()=>{doExport();return (document.getElementById('toast')||{}).textContent||''});
  await p.waitForFunction(()=>!XJOB,null,{timeout:300000,polling:100});await p.waitForTimeout(1500);p.off('download',on5);
  ok('מחשב · ייצוא נוסף בזמן ייצוא — לא מתחיל, ואומר למה; יורד קובץ אחד',/עדיין בהכנה/.test(t5)&&files5.length===1,`${files5.length} קבצים · «${t5}»`);
  /* 6 — ביטול */
  const files6=[];const on6=d=>files6.push(d);p.on('download',on6);
  await p.evaluate(()=>doExport());await p.waitForFunction(()=>!!document.getElementById('xCancelBtn'),null,{timeout:10000});
  await p.click('#xCancelBtn');await p.waitForTimeout(2500);p.off('download',on6);
  const c6=await p.evaluate(()=>({job:!!XJOB,box:document.getElementById('loadbox').hidden,key:localStorage.getItem('planner_exporting_v1'),toast:(document.getElementById('toast')||{}).textContent||''}));
  ok('מחשב · ביטול — אין קובץ, התיבה יורדת, הרישום נמחק',!files6.length&&!c6.job&&c6.box&&c6.key==null&&/בוטל/.test(c6.toast),c6.toast);
  /* 7 — טעינת דוח בזמן ייצוא מבטלת את הייצוא */
  await p.evaluate(()=>doExport());await p.waitForFunction(()=>!!XJOB,null,{timeout:10000});
  await p.setInputFiles('#f',XL);
  await p.waitForFunction(()=>!XJOB&&!LOADJOB&&LOADMS>0,null,{timeout:300000,polling:200});
  ok('מחשב · טעינת דוח בזמן ייצוא — הייצוא מתבטל והדוח נטען',await p.evaluate(()=>!XJOB&&ALL.length>0));
  await p.evaluate(()=>{setMode('catalog');for(const k of Object.keys(colFilters))delete colFilters[k];apply()});await p.waitForTimeout(300);
  /* 8 — Worker שנכשל באמצע → נכתב בדף, זהה */
  await p.evaluate(()=>{window.__W=window.Worker;window.Worker=function(u){const w=new __W(URL.createObjectURL(new Blob(['let n=0;onmessage=()=>{if(++n>2)throw new Error("boom");postMessage({t:"ok"})}'],{type:'text/javascript'})));return w}});
  const w8=await exp(p,'fallback');
  ok('מחשב · Worker שנכשל באמצע — הקובץ נכתב בדף, זהה תא בתא',!!w8.f&&fileSig(w8.f)===fileSig(w1.f),await p.evaluate(()=>(document.getElementById('toast')||{}).textContent||''));
  /* 9 — אין Worker → בדף, זהה; והדף נתפס (המדידה מבחינה) */
  await p.evaluate(()=>{window.Worker=undefined;window.__mx=0;window.__xon=true});
  const w9=await exp(p,'noworker');
  const stallM=await p.evaluate(()=>{window.__xon=false;window.Worker=window.__W;return Math.round(__mx)});
  ok('מחשב · אין Worker — הקובץ נכתב בדף, זהה תא בתא',!!w9.f&&fileSig(w9.f)===fileSig(w1.f));
  ok('מחשב · הדף מגיב בזמן הייצוא ב-Worker (מול כתיבה בדף)',stallW<600&&stallM>stallW*3,`Worker ${stallW}ms · בדף ${stallM}ms`);
  /* 10 — אחרי כל אלה: הנתונים לא זזו, והכלי עובד */
  const after=await p.evaluate(async()=>{const k=FDEF[0][0];SEL[k].add(ALL[0][k]);apply();const f=currentRows().length;SEL[k].clear();apply();setMode('month');await new Promise(r=>setTimeout(r,200));return {f,n:currentRows().length}});
  ok('מחשב · אחרי הייצואים הנתונים והסיווג לא זזו, וסינון ומעבר מסך עובדים',(await p.evaluate(SNAP))===snap0&&after.f>0&&after.n>0,`${after.f} בסינון`);
  /* 11 — ייצוא שלא הסתיים (הלשונית נסגרה/קרסה) → הודעה בפתיחה הבאה */
  await p.evaluate(()=>localStorage.setItem('planner_exporting_v1',JSON.stringify({n:'Inventory_all_2026-10-25.xlsx',r:100000,t:Date.now()})));
  await p.reload();await p.waitForTimeout(800);
  const s11=await p.evaluate(()=>({box:document.getElementById('loadbox').textContent,hidden:document.getElementById('loadbox').hidden,key:localStorage.getItem('planner_exporting_v1')}));
  ok('מחשב · ייצוא שלא הסתיים — בפתיחה הבאה: שם הקובץ והסבר',!s11.hidden&&/הייצוא הקודם לא הסתיים/.test(s11.box)&&/Inventory_all/.test(s11.box)&&s11.key==null,s11.box.slice(0,70));
  ok('מחשב · אפס שגיאות JS',!errs.length,errs.slice(0,3).join(' | '));
  await ctx.close()}
 /* ---------------- טלפון ---------------- */
 {const {ctx,p,errs}=await open(b,390,844);
  const files=[];p.on('download',d=>files.push(d));
  await p.evaluate(()=>doExport());await p.waitForFunction(()=>!!document.getElementById('xCancelBtn'),null,{timeout:10000});
  const g=await p.evaluate(()=>{const r=document.getElementById('loadbox').getBoundingClientRect(),c=document.getElementById('xCancelBtn').getBoundingClientRect();return {l:Math.round(r.left),r:Math.round(r.right),w:innerWidth,cv:c.width>0&&c.left>=0&&c.right<=innerWidth}});
  ok('טלפון · מצב הייצוא בתוך המסך, «בטל ייצוא» נראה',g.l>=8&&g.r<=g.w-8&&g.cv,`${g.l}–${g.r} מתוך ${g.w}`);
  await p.waitForFunction(()=>!XJOB,null,{timeout:300000,polling:100});const t=Date.now();while(!files.length&&Date.now()-t<5000)await p.waitForTimeout(100);
  ok('טלפון · הקובץ יורד',files.length===1);
  ok('טלפון · אפס שגיאות JS',!errs.length,errs.slice(0,3).join(' | '));
  await ctx.close()}
 await b.close();fs.rmSync(TMP,{recursive:true,force:true});
 console.log(out.join('\n'));console.log(`\n${out.filter(l=>l.startsWith("PASS")).length} עברו · ${bad} נכשלו`+(skip?` · ${skip} דולג`:''));
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
