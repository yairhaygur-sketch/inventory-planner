/* ============ TERMS שלב 2 — רשימה אחת = קובץ SAP, זירוז ספקים, נתון חסר באמת (אושר 21.11) ============
   ההחלטה (docs/BUSINESS-DECISIONS.md, docs/DECISION-RULES.md 5ת5):
   · «המלצות לעדכון פרמטרים» = בדיוק השורות של קובץ SAP (sapMassUsable).
   · «ROP קצרה מזמן האספקה», «ROP נמוכה מהמומלץ» ו«פרמטר מעל מכירות 3 שנים»
     אינם דלתות: כפתורי סינון מעל הרשימה (MREASONS), עם מונה לכל אחד.
   · «מעקב אספקה» (תור follow) — דלת «זירוז ספקים» במעקב הזמנות רכש, ולא
     ברשימת הפרמטרים.
   · «נתוני אב חסרים» — רק פריט שחסר לו זמן אספקה או רמת שירות.
   הבדיקה טוענת את הפיקסצ׳ר ובודקת כל אחד מאלה מול שדות המנוע, ומוודאת שיש
   פריטים בכל ענף — אחרת היא אינה בודקת דבר. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML=process.argv[2]||path.join(SD,'..','index.html');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1366,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+HTML+'?nobrief=1');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});
 const R=await p.evaluate(()=>{
  const all=QF.all,fol=new Set(QF.follow||[]);
  setMode('month');const list=currentRows(),file=sapPlan(undefined,list,list).rows;
  const fileSet=all.filter(sapMassUsable);
  const qualRows=(()=>{setMode('qual');return currentRows()})();
  const followRows=(()=>{setMode('follow');return {rows:currentRows(),lt:document.getElementById('lt').textContent,
    view:colViewKey(),trs:document.querySelectorAll('#tbl tbody tr[data-i]').length,exp:!document.getElementById('expXls').hidden}})();
  /* כפתורי הסיבה */
  setMode('month');
  const btns=[...document.querySelectorAll('#mrRow [data-mr]')].map(e=>({k:e.dataset.mr,n:+e.querySelector('b').textContent.replace(/\D/g,''),on:e.classList.contains('on'),dis:e.disabled}));
  const rowVis=!document.getElementById('mrRow').hidden;
  const per={};for(const [k,,f] of MREASONS)per[k]={want:list.filter(f).length};
  for(const [k,,f] of MREASONS){if(!per[k].want)continue;
   document.querySelector(`#mrRow [data-mr="${k}"]`).click();
   const v=currentRows();per[k].got=v.length;per[k].allMatch=v.every(f);per[k].on=document.querySelector(`#mrRow [data-mr="${k}"]`).classList.contains('on');
   per[k].trs=document.querySelectorAll('#tbl tbody tr[data-i]').length}
  document.querySelector('#mrRow [data-mr=""]').click();const back=currentRows().length;
  /* הקישור הוותיק «נקודת הזמנה נמוכה מהנדרש» (UX4) פותח את אותו סינון */
  setMode('month','ropUp');const viaSub={mr:MREASON,n:currentRows().length,on:!!document.querySelector('#mrRow [data-mr="ropUp"].on')};
  setMode('line');const rowHidden=document.getElementById('mrRow').hidden;
  return {n:list.length,file,fileSet:fileSet.length,
   notInFile:list.filter(r=>!sapMassUsable(r)).length,missing:fileSet.filter(r=>!list.includes(r)).length,
   followInList:list.filter(r=>fol.has(r)&&!sapMassUsable(r)).length,folN:fol.size,
   follow:{n:followRows.rows.length,same:followRows.rows.length===fol.size&&followRows.rows.every(r=>fol.has(r)),lt:followRows.lt,view:followRows.view,trs:followRows.trs,exp:followRows.exp},
   qual:{n:qualRows.length,bad:qualRows.filter(r=>!(r.serviceMissing||r.ltMissing)).length,
    missed:(QF.quality||[]).filter(r=>(r.serviceMissing||r.ltMissing)&&!qualRows.includes(r)).length,
    quality:(QF.quality||[]).length},
   btns,rowVis,per,back,viaSub,rowHidden,
   gone:['ropGap','ropThin','over3y'].filter(k=>MODESET[k]),
   nav:{month:navCount('month').n,follow:navCount('follow').n,qual:navCount('qual').n},
   sapDue:sapDueN()}});
 ok('יש פריטים ברשימת הפרמטרים',R.n>0,R.n);
 ok('הרשימה = בדיוק קובץ SAP: כל שורה בקובץ, ואין פריט בלי שורה',R.notInFile===0&&R.missing===0&&R.file===R.n&&R.fileSet===R.n,`רשימה ${R.n} · קובץ ${R.file} · sapMassUsable ${R.fileSet}`);
 ok('המונה במסילה ומספר «לעדכון ב-SAP» = הרשימה',R.nav.month===R.n&&R.sapDue===R.n,JSON.stringify(R.nav)+' · '+R.sapDue);
 ok('יש פריטי «מעקב אספקה» בפיקסצ׳ר',R.folN>0,R.folN);
 ok('אין ברשימת הפרמטרים פריט מעקב אספקה שאין לו שורה בקובץ',R.followInList===0,R.followInList);
 ok('«זירוז ספקים» = תור follow של המנוע, ומונה המסילה זהה',R.follow.same&&R.nav.follow===R.folN,`${R.follow.n} · ${R.folN}`);
 ok('«זירוז ספקים»: כותרת, עמודות מסך החוסרים, שורות, ודוח זירוז',R.follow.lt==='מעקב הזמנות רכש · זירוז ספקים'&&(R.follow.view==='dec'||R.follow.view==='sup')&&R.follow.trs>0&&R.follow.exp,JSON.stringify(R.follow));
 ok('«נתוני אב חסרים»: רק זמן אספקה או רמת שירות חסרים, ואף אחד לא הוחמץ',R.qual.bad===0&&R.qual.missed===0&&R.nav.qual===R.qual.n,JSON.stringify(R.qual));
 ok('תור quality של המנוע גדול מהתחום — כלומר הסינון אכן עובד',R.qual.quality>R.qual.n,`${R.qual.quality} בתור · ${R.qual.n} בתחום`);
 ok('שורת הסיבה גלויה ברשימת הפרמטרים, ונסתרת במסך אחר',R.rowVis&&R.rowHidden,'');
 ok('שישה כפתורים: הכל + חמש סיבות, «הכל» פעיל ומונה = הרשימה',R.btns.length===6&&R.btns[0].k===''&&R.btns[0].on&&R.btns[0].n===R.n,JSON.stringify(R.btns));
 ok('מונה כל כפתור = הפריטים שעונים על הסיבה',R.btns.slice(1).every(x=>x.n===R.per[x.k].want&&x.dis===(x.n===0)),JSON.stringify(R.btns.slice(1).map(x=>[x.k,x.n,R.per[x.k].want])));
 const tested=Object.entries(R.per).filter(([,v])=>v.want>0);
 ok('לפחות שלוש סיבות עם פריטים בפיקסצ׳ר',tested.length>=3,tested.map(([k])=>k).join(','));
 ok('לחיצה על סיבה: בדיוק הפריטים שלה, הכפתור פעיל, והטבלה מציגה אותם',tested.every(([,v])=>v.got===v.want&&v.allMatch&&v.on&&v.trs===Math.min(v.want,v.trs)&&v.trs>0),JSON.stringify(R.per));
 ok('«הכל» מחזיר את כל הרשימה',R.back===R.n,R.back);
 ok('הקישור «נקודת הזמנה נמוכה מהנדרש» פותח את סינון «ROP נמוכה»',R.viaSub.mr==='ropUp'&&R.viaSub.on&&R.viaSub.n===R.per.ropUp.want,JSON.stringify(R.viaSub));
 ok('אין יותר דלתות ropGap / ropThin / over3y',!R.gone.length,R.gone.join(','));
 ok('0 שגיאות JS',!errs.length,errs.join('|'));
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})();
