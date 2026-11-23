/* ============ TERMS שלב 3 — «שווי עודף» במסך, ושם הסיווג (אושר 23.11) ============
   ההחלטה (docs/BUSINESS-DECISIONS.md):
   · «כסף תקוע» → «שווי עודף» — בעמודות, בכרטיס הפריט ובשורות ההסבר. כמו באקסל.
   · «למה הכסף תקוע» → «סוג העודף» · «% מצטבר מהכסף» → «% מצטבר מהשווי».
   · הסיווג «מספר לתקן ב-SAP» → «עדכון פרמטרים». אותו ענף במנוע — רק השם.
   הבדיקה עוברת על כל המסכים בפיקסצ׳ר ובודקת את מה שמוצג בפועל. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML=process.argv[2]||path.join(SD,'..','index.html');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+HTML+'?nobrief=1');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});
 const R=await p.evaluate(async()=>{const seen=[],old=[];const W=()=>new Promise(s=>setTimeout(s,30));
  for(const m of Object.keys(MODESET)){try{setMode(m)}catch(_){continue}await W();
   const t=document.body.innerText;if(/כסף תקוע|הכסף התקוע|למה הכסף|מצטבר מהכסף|מספר לתקן ב-SAP/.test(t))old.push(m);seen.push(m)}
  setMode('cap');await W();const capHeads=[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim());
  setMode('catalog');await W();const catHeads=[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim());
  /* כרטיס של פריט עם שווי עודף */
  const it=QF.all.find(r=>r.expCap>0&&!r.slowEst);detail(it);await W();
  const card=document.getElementById('detail').innerText;
  return {seen:seen.length,old,capHeads,catHeads,card:/שווי עודף/.test(card),cardOld:/כסף תקוע/.test(card),
   cats:QF.all.filter(r=>r.cat==='מספר לתקן ב-SAP').length,newCat:QF.all.filter(r=>r.cat==='עדכון פרמטרים').length,
   newCatFix:QF.all.filter(r=>r.cat==='עדכון פרמטרים').every(r=>r.paramFix)}});
 ok('נסרקו כל המסכים',R.seen>=20,R.seen);
 ok('אף מסך אינו מציג «כסף תקוע», «למה הכסף תקוע», «% מצטבר מהכסף» או «מספר לתקן ב-SAP»',!R.old.length,R.old.join(','));
 ok('מלאי עודף: «סוג העודף», «שווי עודף», «% מצטבר מהשווי»',['סוג העודף','שווי עודף','% מצטבר מהשווי'].every(h=>R.capHeads.includes(h)),R.capHeads.join('|'));
 ok('קטלוג: עמודת «שווי עודף»',R.catHeads.includes('שווי עודף'),R.catHeads.join('|'));
 ok('כרטיס פריט: «שווי עודף», בלי «כסף תקוע»',R.card&&!R.cardOld);
 ok('הסיווג «עדכון פרמטרים» קיים בפיקסצ׳ר, ו«מספר לתקן ב-SAP» לא',R.newCat>0&&R.cats===0,`${R.newCat} · ${R.cats}`);
 ok('«עדכון פרמטרים» — אותו ענף: כולם paramFix',R.newCatFix);
 ok('0 שגיאות JS',!errs.length,errs.join(' | '));
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
