/* ============ מצב הסבר ============
   שתי סכנות בתכונה כזאת, ושתיהן שקטות:
   1. כפתור שמסומן כ«מוסבר» ואין לו טקסט — לוחצים ומקבלים חלונית ריקה.
   2. מצב ההסבר מפעיל את הכפתור בטעות — לוחצים כדי ללמוד ומשנים מסלול.
   הבדיקה רודפת בדיוק אחרי השתיים. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1600,height:950}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');
 await p.waitForSelector('#tbl tbody tr[data-i]',{timeout:120000});
 await p.waitForTimeout(700);

 /* --- כבוי כברירת מחדל, בלי שום נוכחות --- */
 const off=await p.evaluate(()=>({
   on:document.body.classList.contains('xp'),
   pop:!!document.getElementById('xpPop')}));
 ok('מצב ההסבר כבוי כברירת מחדל',!off.on&&!off.pop);

 /* --- כל מה שמסומן כמוסבר חייב טקסט בשתי השאלות --- */
 const cover=await p.evaluate(()=>{
   const keys=[...document.querySelectorAll('[data-xp]')].map(e=>e.dataset.xp);
   const uniq=[...new Set(keys)];
   const missing=uniq.filter(k=>!XP[k]);
   const thin=Object.entries(XP).filter(([k,v])=>
     !v.t||!v.see||!v.chk||v.see.length<40||v.chk.length<25).map(([k])=>k);
   const unused=Object.keys(XP).filter(k=>!uniq.includes(k));
   return {found:uniq.length,missing,thin,unused,total:Object.keys(XP).length}});
 ok('לכל כפתור מסומן יש ערך במילון',cover.missing.length===0,
   cover.found+' מפתחות על המסך · חסרים: '+JSON.stringify(cover.missing));
 ok('אין ערך רזה — לכל אחד «מה רואים» ו«מה לבדוק»',cover.thin.length===0,
   JSON.stringify(cover.thin));

 /* --- שבעת המסלולים חייבים להיות מכוסים, זו עיקר הבקשה --- */
 const navCov=await p.evaluate(()=>{
   const tabs=[...document.querySelectorAll('#tabs .tab')];
   return {n:tabs.length,
     withXp:tabs.filter(t=>t.dataset.xp&&XP[t.dataset.xp]).length,
     names:tabs.map(t=>t.dataset.m)}});
 ok('כל טאב בסרגל הניווט מסביר את עצמו',navCov.n>0&&navCov.withXp===navCov.n,
   navCov.withXp+' מתוך '+navCov.n+' · '+navCov.names.join(' · '));

 /* --- הדלקה --- */
 await p.click('#xpBtn');await p.waitForTimeout(250);
 ok('הכפתור מדליק את המצב',
   await p.evaluate(()=>document.body.classList.contains('xp')));

 /* --- הסכנה האמיתית: לחיצה מסבירה ולא מפעילה --- */
 const before=await p.evaluate(()=>mode);
 const target=await p.evaluate(()=>{
   const t=[...document.querySelectorAll('#tabs .tab')].find(x=>x.dataset.m&&x.dataset.m!==mode);
   return t?t.dataset.m:null});
 await p.evaluate(m=>[...document.querySelectorAll('#tabs .tab')]
   .find(x=>x.dataset.m===m).click(),target);
 await p.waitForTimeout(350);
 const after=await p.evaluate(()=>({mode,
   popShown:(()=>{const e=document.getElementById('xpPop');return !!e&&!e.hidden})(),
   title:(document.querySelector('#xpPop .xph')||{}).textContent,
   see:(document.querySelector('#xpPop .xpb p')||{}).textContent||''}));
 ok('לחיצה במצב הסבר אינה מחליפה מסלול',after.mode===before,
   `היה ${before} · אחרי לחיצה על ${target} נשאר ${after.mode}`);
 ok('נפתחה חלונית עם כותרת',after.popShown&&after.title&&after.title.length>1,after.title);
 ok('החלונית מכילה טקסט אמיתי',after.see.length>40,after.see.slice(0,60)+'…');

 /* --- החלונית נשארת בתוך המסך --- */
 const fit=await p.evaluate(()=>{const e=document.getElementById('xpPop');
   const r=e.getBoundingClientRect();
   return {l:r.left,t:r.top,r:r.right,b:r.bottom,W:innerWidth,H:innerHeight}});
 ok('החלונית אינה חורגת מהמסך',
   fit.l>=0&&fit.t>=0&&fit.r<=fit.W+1&&fit.b<=fit.H+1,
   `x ${Math.round(fit.l)}–${Math.round(fit.r)} מתוך ${fit.W} · y ${Math.round(fit.t)}–${Math.round(fit.b)} מתוך ${fit.H}`);

 /* --- כיבוי מחזיר את ההתנהגות הרגילה --- */
 await p.click('#xpBtn');await p.waitForTimeout(250);
 await p.evaluate(m=>[...document.querySelectorAll('#tabs .tab')]
   .find(x=>x.dataset.m===m).click(),target);
 await p.waitForTimeout(500);
 ok('אחרי כיבוי הכפתור מפעיל שוב',
   target===await p.evaluate(()=>mode),
   'מסלול נוכחי: '+await p.evaluate(()=>mode));
 ok('החלונית נסגרה',
   await p.evaluate(()=>{const e=document.getElementById('xpPop');return !e||e.hidden}));

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 console.log(out.join('\n'));
 await b.close();
 process.exit(bad?1:0);
})();
