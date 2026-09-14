/* ============ שאילתה בשפה חופשית ============
   הפרסר לא מריץ כלום לבד, ולכן הבדיקה היא על *מה שהוא מבין* ועל
   *מה שהוא מודה שלא הבין*. השנייה חשובה יותר: מילה שנבלעת בשקט
   הופכת רשימה מסוננת חלקית לרשימה שנראית שלמה. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');
 await p.waitForSelector('#tbl tbody tr[data-i]',{timeout:120000});
 await p.waitForTimeout(700);

 const parse=q=>p.evaluate(x=>{const r=nlqParse(x);
   return r&&{dims:r.dims,state:r.state?r.state.k:null,unknown:r.unknown,n:r.n,empty:r.empty}},q);

 /* שם אמיתי מהדוח — נלקח מהנתונים, לא מומצא */
 const real=await p.evaluate(()=>{
   const brands={},sups={};
   for(const it of ALL){if(it.brand)brands[it.brand]=(brands[it.brand]||0)+1;
     if(it.supplier)sups[it.supplier]=(sups[it.supplier]||0)+1}
   const top=o=>Object.entries(o).sort((a,b)=>b[1]-a[1])[0];
   return {brand:top(brands),supplier:top(sups)}});

 ok('יש מותג וספק בדוח לבדיקה',!!real.brand&&!!real.supplier,
   `${real.brand[0]} (${real.brand[1]}) · ${real.supplier[0]} (${real.supplier[1]})`);

 /* 1. מותג אמיתי — הספירה חייבת להיות בדיוק מה שיש בנתונים */
 const q1=await parse(real.brand[0]);
 ok('מותג מהדוח מזוהה ונספר נכון',
   q1&&q1.dims.brand&&q1.dims.brand.includes(real.brand[0])&&q1.n===real.brand[1],
   `זוהה ${JSON.stringify(q1&&q1.dims.brand)} · ${q1&&q1.n} מול ${real.brand[1]}`);

 /* 2. ביטוי מצב */
 const q2=await parse('לקוח ממתין');
 ok('«לקוח ממתין» ממופה למסלול cust',q2&&q2.state==='cust',q2&&q2.state);
 const q3=await parse('עודף');
 ok('«עודף» ממופה להון כלוא',q3&&q3.state==='catalog',q3&&q3.state);

 /* 3. שילוב: מותג + מצב */
 const q4=await parse(real.brand[0]+' לקוח ממתין');
 ok('מותג יחד עם מצב — שניהם מזוהים',
   q4&&q4.state==='cust'&&q4.dims.brand&&q4.dims.brand.includes(real.brand[0]),
   JSON.stringify(q4&&{s:q4.state,b:q4.dims.brand}));

 /* 4. הכלל החשוב: מילה שלא זוהתה חייבת להיות מדווחת */
 const q5=await parse('קסדות סגולות מרחפות');
 ok('מילים שלא זוהו מדווחות ולא נבלעות',
   q5&&q5.unknown.length>=2,JSON.stringify(q5&&q5.unknown));
 ok('שאילתה שלא זוהתה כלל מסומנת ריקה',q5&&q5.empty===true);

 /* 5. מילות קישור אינן נחשבות «לא זוהה» */
 const q6=await parse('תראה לי את כל הפריטים של '+real.brand[0]);
 ok('מילות קישור אינן מדווחות כשגיאה',
   q6&&q6.unknown.length===0&&q6.dims.brand,
   'לא זוהו: '+JSON.stringify(q6&&q6.unknown));

 /* 6. הספירה תואמת את מה שהחלה באמת תיתן */
 const q7=await parse(real.supplier[0]);
 const applied=await p.evaluate(v=>{
   const r=nlqParse(v);nlqApply(r);
   return {rows:(QF&&QF.all?QF.all:[]).length,promised:r.n}},real.supplier[0]);
 ok('הספירה שהובטחה היא הספירה שהתקבלה',
   applied.rows===applied.promised,`הובטח ${applied.promised} · התקבל ${applied.rows}`);

 /* 7. ניקוי מחזיר הכול */
 await p.evaluate(()=>{FDEF.forEach(([k])=>SEL[k]=new Set());apply()});
 const back=await p.evaluate(()=>(QF&&QF.all?QF.all:[]).length);
 ok('ניקוי הסינון מחזיר את כל הפריטים',back===await p.evaluate(()=>ALL.length).catch(()=>back),
   back+' פריטים');

 /* 8. ריק אינו מפיל */
 const q8=await parse('');
 ok('שאילתה ריקה אינה מפילה',q8===null||q8===undefined||q8.empty!==false);

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 console.log(out.join('\n'));
 await b.close();
 process.exit(bad?1:0);
})();
