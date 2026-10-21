/* צילום של מה שהמנוע האמיתי החזיר — עותק אחד, דוח אחד.
   אין כאן שחזור של שום פרדיקט: נטען index.html, מוזן הדוח, ונקרא ALL.
   שרשרת classify בנויה מרונגים עם continue, ומדידה עצמאית של פרדיקט
   נתנה בעבר טעות של פי 7.
   שימוש:  node engine-snapshot.js <index.html> <report.xlsx> <out.json> */
const {chromium}=require('playwright');
const EXE=process.env.CHROMIUM_PATH||undefined;
(async()=>{
const [,,html,xlsx,out]=process.argv;
if(!html||!xlsx||!out){console.error('שימוש: engine-snapshot.js <html> <xlsx> <out.json>');process.exit(2)}
const b=await chromium.launch({executablePath:EXE});
const p=await (await b.newContext({viewport:{width:1512,height:900}})).newPage();
const errs=[];p.on('pageerror',e=>errs.push(e.message.split('\n')[0].slice(0,160)));
await p.goto('file://'+require('path').resolve(html)+'?nobrief=1');
await p.setInputFiles('#f',require('path').resolve(xlsx));
await p.waitForFunction(()=>typeof ALL!=='undefined'&&ALL&&ALL.length>0,null,{timeout:900000});
await p.waitForTimeout(4000);
const snap=await p.evaluate(()=>{
 const o={};
 /* כל התור שהמנוע בונה — כולל over3y ו-tier. תת-קבוצה כאן היא
    בדיוק איך מפספסים דלת שזזה. */
 for(const k of Object.keys(QF||{}))if(Array.isArray(QF[k])&&k!=='all')o['Q.'+k]=QF[k].map(x=>x.pn);
 const D={};
 const F={floor:()=>floorPick(floorAll(),'floor'),floorNR:()=>floorPick(floorAll(),'floorNR'),
  ropGap:()=>ropPick(ropAll(),'ropGap'),ropThin:()=>ropPick(ropAll(),'ropThin'),
  trend:()=>trendRows(),rise:()=>riseRows(),burn:()=>burnRows(),agenda:()=>agendaRows(),
  m1:()=>m1Rows(QF.all||[])};
 for(const k in F){try{D['D.'+k]=(F[k]()||[]).map(x=>x.pn)}catch(e){D['D.'+k]='ERR:'+e.message.slice(0,60)}}
 const M={};
 try{for(const m of MODES)M[m[0]]=modeRows(m[0]).length}catch(e){}
 return {n:ALL.length,queues:o,doors:D,modes:M,
  items:ALL.map(r=>({pn:r.pn,cat:r.cat,sev:r.sev,tier:r.ageTier,
   rop:r.rop,sugROP:r.sugROP,ss:r.ss,sugSS:r.sugSS,
   orderQty:r.orderQty==null?null:r.orderQty,miss:r.miss||0,excessQty:r.excessQty||0,
   expCap:Math.round(r.expCap||0),cur:r.currency||'',
   rate:+(r.rate||0).toFixed(3),cov:r.cov==null?null:+r.cov.toFixed(2),
   paramFix:!!r.paramFix,act:((r.act||[]).filter(Boolean)[0]||'')}))};
});
require('fs').writeFileSync(out,JSON.stringify({errs,...snap}));
console.log(`${out.split('/').pop()}: ${snap.n} פריטים · שגיאות JS: ${errs.length}`);
if(errs.length){console.log('  '+errs.slice(0,3).join(' | '));process.exitCode=1}
await b.close()})();
