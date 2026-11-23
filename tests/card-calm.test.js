/* ============ DESIGN-B — כרטיס פריט רגוע (ביקורת העיצוב, אושר 21.11) ============
   מה נעול כאן:
   · למעלה: ההחלטה, המספרים, עד שתי סיבות וכפתור «טופל».
   · פעולות נוספות, «גם ב...» ושאר הסיבות — בקיפול אחד «עוד N פרטים», סגור
     כברירת מחדל. שום דבר לא נמחק: כל הסיבות והפעולות ב-DOM.
   · תגית סוג ה-MRP אומרת שזה סוג MRP; תגית גיל המלאי רק כשהוא לא «פעיל».
   הבדיקה עוברת על פריטים מכמה מסכים בפיקסצ׳ר ומוודאת שיש פריט עם קיפול —
   אחרת היא אינה בודקת דבר. */
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
 const R=await p.evaluate(async()=>{const res=[];
  for(const m of ['home','today','month','cap','follow']){setMode(m);
   for(const r of currentRows().slice(0,20)){detail(r);await new Promise(s=>setTimeout(s,10));
    const d=document.getElementById('detail'),dcs=d.querySelector('.dcs'),fold=dcs.querySelector('.dcfold');
    const topWhy=[...dcs.querySelectorAll(':scope > .dcwhy .msg')].length;
    const allWhy=dcs.querySelectorAll('.dcwhy .msg').length;
    const acts=(r.act||[]).filter(Boolean).length;
    const tags=[...d.querySelectorAll('.ohead .tags .tag')].map(t=>t.textContent.trim());
    res.push({m,pn:r.pn,topWhy,allWhy,whySrc:(r.why||[]).length,fold:!!fold,open:fold?fold.open:null,
     sum:fold?fold.querySelector('summary').textContent.trim():'',
     foldItems:fold?(fold.querySelectorAll('.dcwhy .msg').length+(fold.querySelector('.dcmore')?1:0)+(fold.querySelector('.dcx')?1:0)):0,
     moreOutside:!!dcs.querySelector(':scope > .dcmore'),xOutside:!!dcs.querySelector(':scope > .dcx'),
     xo:xOther(r).length,acts,
     btn:!!dcs.querySelector('.dcb [data-dmark]'),
     mrpTag:tags.some(t=>t==='סוג MRP: '+(r.mrp||'—')),
     ageTag:tags.some(t=>/^גיל מלאי:/.test(t)),active:r.ageTier==='active'})}}
  return res});
 ok('נבדקו פריטים מחמישה מסכים',R.length>=40,R.length);
 const folded=R.filter(x=>x.fold);
 ok('יש בפיקסצ׳ר פריטים עם קיפול — אחרת הבדיקה ריקה',folded.length>0,folded.length);
 ok('למעלה לכל היותר שתי סיבות',R.every(x=>x.topWhy<=2),R.filter(x=>x.topWhy>2).map(x=>x.pn).join(','));
 ok('אף סיבה לא נמחקה: כל הסיבות ב-DOM',R.every(x=>x.allWhy>=Math.min(x.whySrc,x.allWhy)&&x.allWhy>=x.whySrc),R.filter(x=>x.allWhy<x.whySrc).map(x=>`${x.pn}:${x.allWhy}/${x.whySrc}`).join(','));
 ok('פעולות נוספות ו«גם ב...» אינם מעל הקיפול',R.every(x=>!x.moreOutside&&!x.xOutside),'');
 ok('«גם ב...» נמצא בקיפול בכל פריט שיש לו דלת נוספת',R.filter(x=>x.xo>0).every(x=>x.fold),R.filter(x=>x.xo>0&&!x.fold).map(x=>x.pn).join(','));
 ok('הקיפול סגור כברירת מחדל, והכיתוב סופר את מה שבתוכו',folded.every(x=>x.open===false&&(x.foldItems===1?x.sum==='עוד פרט אחד':x.sum===`עוד ${x.foldItems} פרטים`)),
   folded.filter(x=>x.open!==false||!(x.foldItems===1?x.sum==='עוד פרט אחד':x.sum===`עוד ${x.foldItems} פרטים`)).map(x=>`${x.pn}:${x.sum}/${x.foldItems}`).join(','));
 ok('בלי מה לקפל — אין קיפול',R.filter(x=>!x.fold).every(x=>x.whySrc<=2&&x.xo===0),'');
 ok('כפתור «טופל» בכל כרטיס',R.every(x=>x.btn),'');
 ok('תגית «סוג MRP: …» בכל כרטיס',R.every(x=>x.mrpTag),R.filter(x=>!x.mrpTag).map(x=>x.pn).join(','));
 ok('תגית גיל מלאי רק כשאינו «פעיל»',R.every(x=>x.ageTag===!x.active),R.filter(x=>x.ageTag!==!x.active).map(x=>x.pn).join(','));
 ok('0 שגיאות JS',!errs.length,errs.join('|'));
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})();
