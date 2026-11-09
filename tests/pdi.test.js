/* ============ PDI — הזמנת לקוח מול מלאי לא פנוי (אושר 2.11, 8.11) ============
   ההחלטה (docs/BUSINESS-DECISIONS.md):
   · מלאי לא פנוי = סה״כ − פנוי. כשהוא מכסה את מה שחסר להזמנת הלקוח אחרי
     המלאי הפנוי והרכש הפתוח — ההזמנה אינה «ללא כיסוי». מצב משלה: «מלאי לא
     פנוי מול ההזמנה — לבדוק: שמור ל-PDI או חסום», בלי המלצת רכש אוטומטית.
   · הדוח אינו מבדיל בין חסום לשמור, ולכן ההזמנה גם אינה נחשבת מכוסה.
   · כיסוי חלקי — נשאר «ללא כיסוי», ושורת מידע אחת בכרטיס.
   הבדיקה רצה על הפיקסצ׳ר (6 פריטים מכוסים במלא, 102 חלקית) ובודקת את
   ההחלטה על כל פריט, לא על דוגמה. נמדד על הדוח האמיתי 8.11: 0 במצב, 1 חלקי. */
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
 const R=await p.evaluate(()=>{const F=r=>Math.max(0,r.free||0),NF=r=>Math.max(0,(r.stock||0)-F(r)),
   GAP=r=>(r.cust||0)-(F(r)+Math.max(0,r.po||0));
  const all=QF.all,pdi=all.filter(r=>r.pdi),gapped=all.filter(r=>!r.mark&&(r.cust||0)>0&&GAP(r)>0&&r.custGap);
  const imm=new Set(QF.immediate);
  return {n:pdi.length,
   pdiOk:pdi.filter(r=>!((r.cust||0)>0&&GAP(r)>0&&NF(r)>=GAP(r))).map(r=>r.pn),
   pdiFlags:pdi.filter(r=>r.burn||r.custGap||r.cat!==PDI_CAT||!imm.has(r)||r.miss!==GAP(r)).map(r=>r.pn),
   pdiNoBuy:pdi.filter(r=>(r.act||[]).some(a=>/לפתוח הזמנת רכש|להשלים רכש/.test(a))||r.orderQty).map(r=>r.pn),
   pdiWhy:pdi.filter(r=>!(r.why||[]).some(w=>/אינו מבדיל בין מלאי חסום למלאי שמור/.test(w[1]))).map(r=>r.pn),
   fullLeft:gapped.filter(r=>NF(r)>=GAP(r)).map(r=>r.pn),
   part:gapped.filter(r=>NF(r)>0).length,
   partLine:gapped.filter(r=>NF(r)>0&&!(r.why||[]).some(w=>w[0]==='info'&&w[1]===`${NF(r)} יח׳ במלאי לא פנוי — לבדוק אם שמורות להזמנה`)).map(r=>r.pn),
   noLineWithout:gapped.filter(r=>NF(r)===0&&(r.why||[]).some(w=>/במלאי לא פנוי/.test(w[1]))).map(r=>r.pn),
   doors:{burn:burnRows().filter(r=>r.pdi).length,cust:custRows().filter(r=>r.pdi).length},
   topics:((QF&&QF.all)||[]).filter(r=>r.custGap&&r.pdi).length,
   inToday:pdi.filter(r=>!r.isOD&&!r.isPDItem&&!r.noStock&&(isT1(r)||(r.unsQty||0)>0)).filter(r=>!decisionList('short').includes(r)).map(r=>r.pn),
   expNoOrder:(()=>{const x=JSON.stringify(expediteRows());return pdi.filter(r=>(r.po||0)<=0&&x.includes(r.pn)).map(r=>r.pn)})()}});
 ok('יש פריטים במצב PDI בפיקסצ׳ר (אחרת הבדיקה אינה בודקת דבר)',R.n>0,R.n);
 ok('כל פריט במצב: הזמנת לקוח פתוחה, חסר אחרי פנוי + רכש, והמלאי הלא פנוי מכסה את החסר',!R.pdiOk.length,R.pdiOk.join(','));
 ok('ואף פריט ש«ללא כיסוי» אינו מכוסה במלא על ידי מלאי לא פנוי',!R.fullLeft.length,R.fullLeft.join(','));
 ok('המצב: לא burn, לא custGap, בתור immediate, החסר נשאר (אינו נחשב מכוסה)',!R.pdiFlags.length,R.pdiFlags.join(','));
 ok('אין המלצת רכש אוטומטית במצב הזה',!R.pdiNoBuy.length,R.pdiNoBuy.join(','));
 ok('הכרטיס אומר למה: הדוח אינו מבדיל בין חסום לשמור',!R.pdiWhy.length,R.pdiWhy.join(','));
 ok('אינו ב«אין רכש כלל» ואינו ב«כיסוי חלקי»',R.doors.burn===0&&R.doors.cust===0,JSON.stringify(R.doors));
 ok('ואינו ב«הזמנות לקוח ללא כיסוי» של 4 הנושאים',R.topics===0,R.topics);
 ok('נשאר ב«חוסרים ולקוחות» לפי הכלל הקיים של המסך',!R.inToday.length,R.inToday.join(','));
 ok('ואינו נכנס לדוח הזירוז כ«אין הזמנה בכלל»',!R.expNoOrder.length,R.expNoOrder.join(','));
 ok('כיסוי חלקי: יש פריטים כאלה בפיקסצ׳ר',R.part>0,R.part);
 ok('כיסוי חלקי: נשאר «ללא כיסוי», עם שורת מידע אחת וכמות נכונה',!R.partLine.length,R.partLine.join(','));
 ok('בלי מלאי לא פנוי — אין שורת מידע',!R.noLineWithout.length,R.noLineWithout.join(','));
 ok('0 שגיאות JS',!errs.length,errs.join('|'));
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})();
