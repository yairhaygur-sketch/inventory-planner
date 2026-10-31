/* ============ RATE_DECAY — הכלל עצמו, דרך build() ============
   קצב 3 השנים משמש ל-ROP/SS רק כשהביקוש לא דעך. כאן נבדק הכלל ברמת
   המנוע: הנוסחה של recentRate מול מימוש עצמאי, גבול 0.25 משני צדדיו,
   ומה שאסור לו לזוז — rate, הכיסוי, רצפות ה-SS, ופריטים מחוץ לנתיב.
   המסלול המלא עד קובץ SAP ואקסל נבדק ב-rate-decay-export.test.js. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.setInputFiles('#f',path.join(SD,'zmrp-demo.xlsx'));
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB,null,{timeout:180000});

 /* 1 · recentRate מול מימוש עצמאי, לכל חודש עוגן */
 const F=await p.evaluate(()=>{
  const ref=(clean,months,y3,A)=>{const raw=months.map(v=>Math.max(0,v));const seq=[];
   const inPrev=Math.max(0,11-A);
   if(inPrev>0){const n=12-inPrev,v=Math.max(0,y3[1]-raw.slice(11-inPrev).reduce((a,b)=>a+b,0))/n;for(let k=0;k<n;k++)seq.push(v)}
   else{for(let k=0;k<12;k++)seq.push(y3[1]/12);const n=A-11,v=n>0?Math.max(0,y3[0]-raw.reduce((a,b)=>a+b,0))/n:0;for(let k=0;k<n;k++)seq.push(v)}
   seq.push(...clean.slice().reverse());
   /* משקל 0.5^(גיל/4) לכל חודש, ממומש כרקורסיה אחרת: סכום משוקלל של סדרה שמתחילה מהערך הראשון */
   const a=1-Math.pow(0.5,1/4);let L=seq[0];for(const v of seq)L=(1-a)*L+a*v;return L};
  let s=7;const R=()=>{s=(s*1103515245+12345)&0x7fffffff;return s/0x7fffffff};
  let worst=0,n=0;
  for(let A=1;A<=12;A++)for(let t=0;t<25;t++){
   const months=Array.from({length:11},()=>R()<.3?Math.round(R()*5):0),clean=months.slice();
   const y3=[Math.round(R()*30),Math.round(R()*200),Math.round(R()*200)];
   worst=Math.max(worst,Math.abs(recentRate(clean,months,y3,A)-ref(clean,months,y3,A)));n++}
  return {worst,n,dec:RATE_DECAY,half:RATE_HALF}});
 ok('recentRate = מימוש עצמאי, בכל 12 חודשי העוגן',F.worst<1e-9,`${F.n} מקרים · סטייה מרבית ${F.worst.toExponential(1)}`);
 ok('הקבועים: RATE_DECAY 0.25 · RATE_HALF 4 חודשים',F.dec===0.25&&F.half===4,`${F.dec} · ${F.half}`);

 /* 2 · דרך build(): פריטים נבנים מאותה כותרת של הדוח שנטען */
 const B=await p.evaluate(()=>{
  const hdr=ALL_RAW.hdr,m=mapCols(hdr),tpl=ALL_RAW.rows[0];
  const mk=(pn,o)=>{const r=tpl.slice();const set=(i,v)=>{if(i>=0)r[i]=v};
   set(m.pn,pn);set(m.status,'01');set(m.statusTxt,'פעיל');set(m.mrp,'ND');set(m.srv,80);
   set(m.ss,o.ss);set(m.rop,o.rop||0);set(m.ssMin,o.ssMin);set(m.lt,120);set(m.ltAdd,0);
   set(m.price,o.price||100);set(m.currency,'USD');set(m.free,o.free||0);set(m.stock,o.free||0);set(m.po,0);
   if(m.openDel>=0)set(m.openDel,0);if(m.transfer>=0)set(m.transfer,0);if(m.cust>=0)set(m.cust,0);
   m.months.forEach((ix,k)=>set(ix,(o.m||[])[k]||0));
   [m.y0,m.y1,m.y2].forEach((ix,k)=>set(ix,o.y[k]));
   if(o.st!=null)set(m.status,o.st);
   return r};
  const one=(pn,o,h)=>build([mk(pn,o)],h||hdr)[0];
  /* פריט בסיס: 1 יח׳ בחודש 4, 2025 = 6. λ אינו תלוי ב-2024, ולכן y2 קובע
     את קצב 3 השנים ומזיז אותו סביב הגבול בלי לגעת ב-λ. */
  const base={ss:11,ssMin:11,rop:18,m:[0,0,0,1],free:21};
  const ref=one('B-REF',{...base,y:[1,6,202]});
  const lam=ref.rateRecent,win=24+(ANCHOR.getMonth()+1);
  const y2for=ratio=>lam/ratio*win-1-6;            /* rate3y = (1+6+y2)/win = λ/ratio */
  const below=one('B-LOW',{...base,y:[1,6,y2for(0.249)]});   /* λ/rate3y = 0.249 → דעיכה */
  const above=one('B-HIGH',{...base,y:[1,6,y2for(0.251)]});  /* 0.251 → אין דעיכה */
  /* בלי עמודה שנתית אחת: has3y=false — הנתיב לא קיים, ואין דעיכה */
  const hdrNo=hdr.slice();hdrNo[m.y2]='עמודה אחרת';
  const no3y=one('B-NO3Y',{...base,y:[1,6,202]},hdrNo);
  /* 04 — אין המלצה בכלל, גם כשהקצב דעך */
  const st04=one('B-04',{...base,y:[1,6,202],st:'04'});
  /* צריכה סדירה — d12>3 */
  const reg=one('B-REG',{ss:4,ssMin:4,rop:12,m:Array(11).fill(2),y:[18,30,30],free:30});
  const pick=x=>x&&{rate:x.rate,rateROP:x.rateDecay?x.rateROP:x.rate,rateDecay:!!x.rateDecay,rateRecent:x.rateRecent==null?null:x.rateRecent,rate3y:x.rate3y,
   dLT:x.dLT,ltM:x.ltM,sugSS:x.sugSS,sugROP:x.sugROP,ssStat:x.ssStat,ssMin:x.ssMin,cov:x.cov,free:x.free,rateFrom3y:x.rateFrom3y,has3y:x.has3y};
  return {ref:pick(ref),below:pick(below),above:pick(above),no3y:pick(no3y),st04:pick(st04),reg:pick(reg),lam}});
 const {ref,below,above,no3y,st04,reg}=B;
 const near=(a,b)=>Math.abs(a-b)<1e-9;
 ok('המבנה של PN-07: הביקוש דעך — ROP 12, SS 11 (היה 37)',ref.rateDecay&&ref.sugROP===12&&ref.sugSS===11,
   `λ ${ref.rateRecent.toFixed(3)} · קצב 3 שנים ${ref.rate3y.toFixed(2)} · ROP ${ref.sugROP} · SS ${ref.sugSS}`);
 ok('ROP/SS מחושבים על הקצב העדכני: ביקוש בזמן אספקה = rateROP × LT',near(ref.dLT,ref.rateROP*ref.ltM)&&near(ref.rateROP,ref.rateRecent));
 ok('rate והכיסוי לא זזו — נשארים על קצב 3 השנים',near(ref.rate,ref.rate3y)&&near(ref.cov,ref.free/ref.rate)&&ref.rateFrom3y===true,
   `rate ${ref.rate.toFixed(2)} · כיסוי ${ref.cov.toFixed(1)}`);
 ok('רצפת SS נשמרת: SS = max(סטטיסטי, ssMin), ssMin כפי שהוא',ref.sugSS===Math.max(ref.ssStat,ref.ssMin)&&ref.ssMin===11);
 ok('גבול: יחס 0.249 — דעיכה',below.rateDecay===true&&near(below.rateROP,below.rateRecent),`יחס ${(below.rateRecent/below.rate3y).toFixed(4)}`);
 ok('גבול: יחס 0.251 — אין דעיכה, קצב 3 השנים נשאר ל-ROP',above.rateDecay===false&&near(above.rateROP,above.rate)&&near(above.dLT,above.rate*above.ltM),
   `יחס ${(above.rateRecent/above.rate3y).toFixed(4)}`);
 ok('עמודה שנתית חסרה — אין נתיב 3 שנים ואין דעיכה',no3y.has3y===false&&no3y.rateDecay===false&&no3y.rateRecent===null);
 ok('סטטוס 04 — אין המלצה, גם כשהקצב דעך',st04.sugROP===0&&st04.sugSS===0);
 ok('צריכה סדירה — הנתיב אינו פעיל, הנוסחה כמו קודם',reg.rateRecent===null&&reg.rateDecay===false&&near(reg.dLT,reg.rate*reg.ltM)&&reg.sugROP===Math.ceil(reg.dLT+reg.sugSS),
   `ROP ${reg.sugROP}`);

 /* 3 · הדוח הסינתטי עצמו: אין בו פריט בנתיב הדעיכה, ולכן אף פריט לא זז */
 const D=await p.evaluate(()=>({n:ALL.length,dec:ALL.filter(x=>x.rateDecay).length,
  diff:ALL.filter(x=>!x.rateDecay&&Math.abs(x.dLT-x.rate*x.ltM)>1e-9).length}));
 ok('בכל פריט שאינו בדעיכה: ביקוש בזמן אספקה = rate × LT, כמו קודם',D.diff===0,`${D.n} פריטים · ${D.dec} בדעיכה`);
 ok('אין שגיאות JS',!errs.length,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
