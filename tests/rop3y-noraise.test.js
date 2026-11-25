/* ============ R3Y_NORAISE — קצב 3 שנים אינו מעלה ROP מעל SAP (אושר 24.11) ============
   פריט עם 3 יח׳ או פחות ב-11 החודשים מקבל את קצב 3 השנים. כ-ROP הוא מקפיץ: 2 יח׳ בשנה
   האחרונה, 90 בשלוש שנים, LT ‏120 יום → ROP 15 מול 6 ב-SAP. בנתיב הזה ה-ROP המוצע אינו עולה
   על SAP; SS ≤ ROP. rate, הכיסוי והסיווג לא זזים. rop3yCalc — רק בפריט שבו הכלל חל. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.evaluate(()=>{try{localStorage.clear()}catch(_){}});
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});

 const R=await p.evaluate(()=>{const A=ALL;
  const y3=A.filter(r=>r.rateFrom3y&&!r.isPDItem&&!r.isOD&&!r.ltBad);
  const capped=A.filter(r=>r.rop3yCalc!=null);
  return {y3:y3.length,capped:capped.length,
   up:y3.filter(r=>(r.sugROP||0)>(r.rop||0)).map(r=>r.pn),
   allCapped3y:capped.every(r=>r.rateFrom3y&&r.d12<=3),
   eqSAP:capped.every(r=>r.sugROP===Math.max(0,r.rop||0)),
   calcAbove:capped.every(r=>r.rop3yCalc>r.sugROP),
   ssLe:capped.every(r=>r.sugSS<=r.sugROP),
   noField:A.filter(r=>r.rop3yCalc==null).every(r=>!('rop3yCalc' in r)),
   rateKept:capped.every(r=>r.rate===r.rate3y),
   why:capped.every(r=>(r.why||[]).some(w=>/קצב 3 שנים אינו מעלה נקודת הזמנה/.test(w[1])&&w[1].includes(String(r.rop3yCalc)))),
   pick:capped[0]?capped[0].pn:null}});
 ok('בדוח הסינתטי יש פריטים בנתיב 3 השנים, ולפחות אחד שהכלל חל עליו',R.y3>0&&R.capped>0,`${R.y3} בנתיב · ${R.capped} נחתכו`);
 ok('אף פריט בנתיב 3 השנים לא מקבל ROP מוצע מעל SAP',R.up.length===0,R.up.slice(0,3).join(', '));
 ok('הכלל חל רק בנתיב 3 השנים (3 יח׳ או פחות ב-11 חודשים)',R.allCapped3y);
 ok('ה-ROP המוצע = ROP של SAP, וה-ROP שחושב (rop3yCalc) גבוה ממנו',R.eqSAP&&R.calcAbove);
 ok('SS מוצע ≤ ROP מוצע',R.ssLe);
 ok('הקצב (rate) נשאר קצב 3 השנים — הסיווג והכיסוי לא זזים',R.rateKept);
 ok('rop3yCalc קיים רק בפריט שבו הכלל חל',R.noField);
 ok('ב«למה»: «קצב 3 שנים אינו מעלה נקודת הזמנה», עם ה-ROP שחושב',R.why);

 /* «איך חושב» (explain) בכלי המלא: שורת החישוב אומרת שהכלל חל, ולא «כלל SS ≤ ROP» */
 const C=await p.evaluate(pn=>{const r=ALL.find(x=>x.pn===pn);detail(r);explain(r);const cl=document.body.cloneNode(true);cl.querySelectorAll("script,style").forEach(x=>x.remove());const t=cl.textContent;
  return {rule:/קצב 3 שנים אינו מעלה נקודת הזמנה \(R3Y_NORAISE\)/.test(t),noSSrule:!/כלל הברזל: SS מוצע ≤ ROP מוצע — לכן ROP הועלה/.test(t)}},R.pick);
 ok('בכרטיס: «קצב 3 שנים אינו מעלה נקודת הזמנה» בשורת החישוב, בלי «כלל הברזל» שגוי',C.rule&&C.noSSrule,JSON.stringify(C));

 /* פריט בנתיב 6 החודשים אינו נחתך גם כשהמוצע מעל SAP */
 const M=await p.evaluate(()=>{const L=ALL.filter(r=>!r.rateFrom3y&&!r.isPDItem&&!r.isOD&&(r.sugROP||0)>(r.rop||0));return {n:L.length,none:L.every(r=>r.rop3yCalc==null)}});
 ok('פריט עם צריכה עדכנית (נתיב 6 החודשים) — ההעלאה נשארת',M.n>0&&M.none,`${M.n} פריטים`);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
