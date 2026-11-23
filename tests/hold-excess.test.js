/* ============ HOLD_EXCESS — לא מעלים פרמטר על מלאי שמכסה שנה (אושר 22.11) ============
   הכלל (docs/BUSINESS-DECISIONS.md, docs/DECISION-RULES.md 5ת5):
   · פריט «עודף מלאי» או «מלאי איטי», קצב > 0, ומלאי (פנוי + רכש) שמכסה יותר
     מ-12 חודשים מעל נקודת ההזמנה המחושבת: אין «להעלות נקודת הזמנה» ואין «להעלות
     מלאי ביטחון». ההמלצה (sugROP / sugSS) = הערך שב-SAP; הקטנה אינה מושהית.
   · הכרטיס אומר «תקין לעת עתה», עם הערכים המחושבים — לא נעלמים.
   · classify רץ שוב בכל סינון — התוצאה זהה (holdRaw מוחזר בתחילת כל ריצה).
   הבדיקה עוברת על כל פריט בפיקסצ׳ר ומוודאת שיש פריטים בכל ענף. */
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
 const snap=()=>p.evaluate(()=>{
  const CATS=new Set(['עודף מלאי','מלאי איטי']);
  const cov=(r,rop)=>((r.free||0)+(r.po||0)-rop)/r.rate;
  const H=QF.all.filter(r=>r.holdRaw);
  const bad1=H.filter(r=>!CATS.has(r.cat)||!(r.rate>0)||!(cov(r,r.holdRaw.rop)>HOLD_MONTHS)).map(r=>r.pn);
  const badRop=H.filter(r=>r.holdRaw.rop>r.rop?r.sugROP!==r.rop:r.sugROP!==r.holdRaw.rop).map(r=>r.pn);
  const badSS=H.filter(r=>r.sugSS!==Math.min(r.holdRaw.ss>r.ss?r.ss:r.holdRaw.ss,r.sugROP)).map(r=>r.pn);
  const whyOk=H.every(r=>(r.why||[]).some(w=>/^תקין לעת עתה/.test(w[1]||'')&&String(w[1]).includes(String(r.holdRaw.rop))));
  const actUp=H.filter(r=>(r.act||[]).some(a=>/^להעלות (נקודת הזמנה|מלאי ביטחון)/.test(a))).map(r=>r.pn);
  /* הכיוון ההפוך: אף פריט בקטגוריה, מעל 12 חודשים, אינו ממליץ להעלות */
  const missed=QF.all.filter(r=>!r.holdRaw&&CATS.has(r.cat)&&r.rate>0&&((r.sugROP||0)>(r.rop||0)||(r.sugSS||0)>(r.ss||0))&&cov(r,r.sugROP||0)>HOLD_MONTHS).map(r=>r.pn);
  /* ולא נגעו בשום דבר אחר: פריט בקטגוריה שמכסה פחות — עדיין מקבל העלאה */
  const keepUp=QF.all.filter(r=>!r.holdRaw&&CATS.has(r.cat)&&r.rate>0&&(r.sugROP||0)>(r.rop||0)).length;
  const inSlow=H.filter(r=>sapSlowUsable(r)&&r.sugROP>=r.rop&&r.sugSS>=r.ss&&sapSsMin(r)===(r.ssMin||0)).map(r=>r.pn);
  const ssRop=QF.all.filter(r=>(r.sugSS||0)>(r.sugROP||0)&&!r.isOD&&(r.sugROP||0)>0).length;
  return {n:H.length,rop:H.filter(r=>r.holdRaw.rop>r.rop).length,ss:H.filter(r=>r.holdRaw.ss>r.ss).length,
   bad1,badRop,badSS,whyOk,actUp,missed,keepUp,inSlow,ssRop,pns:H.map(r=>r.pn).sort().join(',')}});
 const R=await snap();
 ok('יש פריטים מושהים בפיקסצ׳ר — אחרת הבדיקה ריקה',R.n>0,R.n);
 ok('ויש גם השהיית ROP וגם השהיית SS',R.rop>0&&R.ss>0,`ROP ${R.rop} · SS ${R.ss}`);
 ok('כל מושהה: עודף/איטי, קצב > 0, ומכסה יותר מ-12 חודשים מעל ה-ROP המחושב',!R.bad1.length,R.bad1.join(','));
 ok('ROP מוצע = הערך שב-SAP כשהמחושב גבוה ממנו; אחרת — המחושב',!R.badRop.length,R.badRop.join(','));
 ok('SS מוצע = הערך שב-SAP כשהמחושב גבוה ממנו, אחרת המחושב — ובכל מקרה לא מעל ה-ROP',!R.badSS.length,R.badSS.join(','));
 ok('הכרטיס אומר «תקין לעת עתה» עם הערך המחושב',R.whyOk,'');
 ok('אין «להעלות נקודת הזמנה / מלאי ביטחון» בפריט מושהה',!R.actUp.length,R.actUp.join(','));
 ok('אף פריט בקטגוריה מעל 12 חודשים לא הוחמץ',!R.missed.length,R.missed.join(','));
 ok('פריט בקטגוריה שמכסה פחות עדיין מקבל העלאה (הכלל צר)',R.keepUp>0,R.keepUp);
 ok('פריט מושהה בלי הקטנה — אין לו שורה בקובץ SAP של מלאי איטי',!R.inSlow.length,R.inSlow.join(','));
 ok('SS מוצע אינו עולה על ROP מוצע בשום פריט',R.ssRop===0,R.ssRop);
 /* סינון והסרתו — classify רץ פעמיים; התוצאה זהה */
 await p.evaluate(()=>{const k=FDEF[0][0];SEL[k]=new Set([ALL[0][k]]);apply();SEL[k]=new Set();apply()});
 const R2=await snap();
 ok('אחרי סינון והסרתו — אותם פריטים מושהים ואותם ערכים',R2.pns===R.pns&&!R2.badRop.length&&!R2.badSS.length,`${R.n} → ${R2.n}`);
 ok('0 שגיאות JS',!errs.length,errs.join('|'));
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})();
