/* ============ DESIGN-C — «מעקב רכש» = כל הרכש הפתוח, לפי ספק (אושר 21.11) ============
   הבעיה שנמדדה: המסך נפתח על 'stuck' — רשימה שכולה גם ב«חוסרים ולקוחות» (86 מתוך
   86 בדוח האמיתי; בפיקסצ׳ר אותם 168 פריטים ואותן עמודות).
   מה נעול כאן:
   · ברירת המחדל: כל פריט עם הזמנת רכש פתוחה, בקבוצות ספק, ומונה המסילה = הרשימה.
   · בתוך ספק: לקוח ממתין קודם, אחר כך החוסר הגדול, אחר כך הכמות הפתוחה.
   · עמודות: רכש פתוח, הגעה, מלאי, דרישת לקוח, חוסר חזוי, סיווג.
   · כותרת הספק: יח׳ ברכש פתוח ומספר עם לקוח ממתין — לא «יח׳ להזמנה».
   · הרצועה עדיין בוחרת את טבלת החוסרים, והכפתור «כל הרכש הפתוח» מחזיר.
   · מייל לספקים נשאר בירור על מה שאף משלוח ידוע לא מכסה (stuck), גם בברירת המחדל. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML=process.argv[2]||path.join(SD,'..','index.html');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+HTML+'?nobrief=1');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});
 await p.evaluate(()=>setMode('line'));await p.waitForTimeout(600);
 const R=await p.evaluate(()=>{
  const rows=currentRows(),want=QF.all.filter(r=>(r.po||0)>0);
  const SG=supplierGroups(rows);
  const okOrder=SG.every(([,l])=>l.every((r,i)=>!i||(()=>{const a=l[i-1];
    const ca=(a.cust||0)>0,cb=(r.cust||0)>0;if(ca!==cb)return ca;
    if((a.miss||0)!==(r.miss||0))return (a.miss||0)>(r.miss||0);return (a.po||0)>=(r.po||0)})()));
  const heads=[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim());
  const g=document.querySelector('#tbl tbody tr.grp.sup');
  const mail=lineMailText(),stuck=lineModel().stuck;
  const today=modeRows('today').map(r=>r.pn);
  return {sel:LINE_SEL,view:colViewKey(),n:rows.length,want:want.length,same:rows.every(r=>(r.po||0)>0),
   nav:navCount('line').n,groups:SG.length,sups:new Set(want.map(r=>r.supplier||'ללא ספק')).size,okOrder,heads,
   gtxt:g?g.textContent.replace(/\s+/g,' '):'',
   notSame:rows.filter(r=>!today.includes(r.pn)).length,
   btn:(e=>e?{on:e.getAttribute('aria-pressed'),txt:e.textContent.replace(/\s+/g,' ').trim()}:null)(document.querySelector('.lpo .lpob')),
   mailStuck:stuck.length===0||stuck.slice(0,5).every(r=>mail.includes(r.pn)),
   mailNotAll:want.filter(r=>!stuck.includes(r)).slice(0,20).every(r=>!mail.includes(r.pn)),
   listPct:Math.round(document.querySelector('.rows').getBoundingClientRect().height/innerHeight*100)}});
 ok('ברירת המחדל: כל הרכש הפתוח',R.sel==='po'&&R.view==='linepo',`${R.sel} · ${R.view}`);
 ok('הרשימה = כל פריט עם הזמנת רכש פתוחה, ומונה המסילה זהה',R.n===R.want&&R.same&&R.nav===R.n&&R.n>0,`${R.n} · ${R.want} · מסילה ${R.nav}`);
 ok('זו אינה עוד הרשימה של «חוסרים ולקוחות»',R.notSame>0,`${R.notSame} פריטים שאינם בחוסרים`);
 ok('בקבוצות ספק — קבוצה לכל ספק',R.groups===R.sups&&R.groups>0,`${R.groups} קבוצות · ${R.sups} ספקים`);
 ok('בתוך ספק: לקוח ממתין, אחר כך חוסר, אחר כך כמות פתוחה',R.okOrder,'');
 ok('עמודות: רכש פתוח · הגעה · מלאי · דרישת לקוח · חוסר חזוי · סיווג',['רכש פתוח','הגעה','מלאי','דרישת לקוח','חוסר חזוי','סיווג'].every(h=>R.heads.some(x=>x.replace(/\s+/g,' ')===h)),R.heads.join('|'));
 ok('כותרת הספק: יח׳ ברכש פתוח ולקוח ממתין — לא «להזמנה»',/ברכש פתוח/.test(R.gtxt)&&/עם לקוח ממתין/.test(R.gtxt)&&!/להזמנה/.test(R.gtxt),R.gtxt.slice(0,90));
 ok('הכפתור «כל הרכש הפתוח» לחוץ ונושא את המספר',!!R.btn&&R.btn.on==='true'&&R.btn.txt.includes(String(R.n)),JSON.stringify(R.btn));
 ok('מייל לספקים: רק מה שאף משלוח ידוע לא מכסה',R.mailStuck&&R.mailNotAll,'');
 ok('הרשימה עדיין מקבלת 62%+ מהמסך (הכפתור בשורת הרצועה)',R.listPct>=62,R.listPct+'%');
 /* הרצועה בוחרת את טבלת החוסרים, והכפתור מחזיר */
 const S=await p.evaluate(()=>{const e=document.querySelector('#line [data-ls="now"],#line [data-ls="stuck"]');const k=e.dataset.ls;e.click();
  const a={k,sel:LINE_SEL,view:colViewKey(),n:currentRows().length};
  document.querySelector('.lpo .lpob').click();return {a,back:{sel:LINE_SEL,n:currentRows().length,view:colViewKey()}}});
 ok('בחירה ברצועה פותחת את טבלת החוסרים',S.a.sel===S.a.k&&(S.a.view==='dec'||S.a.view==='sup'),JSON.stringify(S.a));
 ok('«כל הרכש הפתוח» מחזיר',S.back.sel==='po'&&S.back.view==='linepo'&&S.back.n===R.n,JSON.stringify(S.back));
 ok('0 שגיאות JS',!errs.length,errs.join('|'));
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})();
