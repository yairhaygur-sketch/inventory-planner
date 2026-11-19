/* ============ SHARP_DROP — ירידה חדה בתוך חלון 6 החודשים ============
   כשסכום 3 החודשים האחרונים ≤ RATE_DECAY × סכום השלושה שלפניהם (ובסיס
   של SHARP_MIN_PRE יח׳ לפחות, בלי מכירה בחודש הנוכחי): ROP על קצב 3
   החודשים האחרונים, ה-SS הסטטיסטי על התנודה שלהם, רצפות ה-SS נשמרות.
   rate, הכיסוי והמגמה לא זזים. וגם DROP_NO_RAISE: בדלת «ירידה בביקוש»
   אין הוראה «לשקול העלאת נקודת הזמנה». */
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

 const B=await p.evaluate(()=>{
  const hdr=ALL_RAW.hdr,m=mapCols(hdr),tpl=ALL_RAW.rows[0];
  const mk=(pn,o)=>{const r=tpl.slice();const set=(i,v)=>{if(i>=0)r[i]=v};
   set(m.pn,pn);set(m.status,'01');set(m.statusTxt,'פעיל');set(m.mrp,'ND');set(m.srv,80);
   set(m.ss,o.ss);set(m.rop,o.rop||0);set(m.ssMin,o.ssMin);set(m.lt,60);set(m.ltAdd,0);
   set(m.price,100);set(m.currency,'USD');set(m.free,o.free);set(m.stock,o.free);set(m.po,0);
   if(m.openDel>=0)set(m.openDel,0);if(m.transfer>=0)set(m.transfer,0);if(m.cust>=0)set(m.cust,0);
   m.months.forEach((ix,k)=>set(ix,o.m[k]||0));
   const d=o.m.reduce((a,v)=>a+v,0);[m.y0,m.y1,m.y2].forEach((ix,k)=>set(ix,[d,d+40,400][k]));
   if(o.st!=null)set(m.status,o.st);
   if(m.curMonth>=0)set(m.curMonth,o.cm||0);
   return r};
  const one=(pn,o)=>build([mk(pn,o)],hdr)[0];
  /* אחרי הירידה: 3,4,3 · לפניה 40,50,45 — יחס 10/135 */
  const tail=[40,40,40,40,40];
  const base={ss:10,ssMin:0,rop:60,free:300,m:[3,4,3,40,50,45,...tail]};
  const sh=one('S-SHARP',base);
  /* גבול: 3 אחרונים = 0.25 × 3 שלפניהם בדיוק (30 מול 120) — חל; 31 — לא */
  const at=one('S-AT',{...base,m:[10,10,10,40,40,40,...tail]});
  const over=one('S-OVER',{...base,m:[11,10,10,40,40,40,...tail]});
  /* בסיס זעיר: 1 מול 5 — לא חל */
  const tiny=one('S-TINY',{...base,m:[0,1,0,2,2,1,1,1,1,1,1]});
  /* מכירה בחודש הנוכחי — לא חל */
  const curm=one('S-CURM',{...base,cm:5});
  /* סטטוס 04 — אין המלצה */
  const st04=one('S-04',{...base,st:'04'});
  /* ssMin=40 — מ-19.11 (SS_STAT) אינו רצפה לפריט עם צריכה */
  const flo=one('S-FLOOR',{...base,ssMin:40});
  /* צריכה סדירה */
  const reg=one('S-REG',{...base,m:Array(11).fill(20)});
  const pick=x=>({statOnly:!!x.ssStatOnly,sharp:!!x.sharpDrop,rate:x.rate,rateDrop:x.rateDrop,rateROP:x.rateROP,dLT:x.dLT,ltM:x.ltM,sugSS:x.sugSS,sugROP:x.sugROP,
   ssStat:x.ssStat,ssMin:x.ssMin,lowFloor:x.lowFloor,sdLT:x.sdLT,sdTot:x.sdTot,zEff:x.zEff??x.z,ltvP:x.ltvP??LTV_P_DEF,cov:x.cov,free:x.free,
   tRec:x.tRec,tPre:x.tPre,trendF:x.trendF,ropRate:ropRate(x),clean:cleanSeries(x.months,x.A),rateFrom3y:x.rateFrom3y,
   keys:['sharpDrop','rateDrop','sd3'].filter(k=>k in x).length});
  return {sh:pick(sh),at:pick(at),over:pick(over),tiny:pick(tiny),curm:pick(curm),st04:pick(st04),flo:pick(flo),reg:pick(reg),
   k:{min:SHARP_MIN_PRE,dec:RATE_DECAY}}});
 const {sh,at,over,tiny,curm,st04,flo,reg}=B;
 const near=(a,b)=>Math.abs(a-b)<1e-9;
 const sdS=a=>{const n=a.length,mu=a.reduce((x,y)=>x+y,0)/n;return Math.sqrt(a.reduce((x,y)=>x+(y-mu)**2,0)/(n-1))};

 ok('הקבועים: SHARP_MIN_PRE 6 · סף 0.25 (RATE_DECAY)',B.k.min===6&&B.k.dec===0.25);
 ok('ירידה חדה מזוהה: 10 יח׳ ב-3 החודשים האחרונים מול 135',sh.sharp&&sh.tRec===10&&sh.tPre===135,`${sh.tRec} מול ${sh.tPre}`);
 const r3=sh.clean.slice(0,3).reduce((a,b)=>a+b,0)/3;
 ok('קצב ל-ROP = סכום 3 החודשים האחרונים ÷ 3',near(sh.rateDrop,r3)&&near(sh.rateROP,r3),sh.rateDrop.toFixed(3));
 ok('ביקוש בזמן אספקה = קצב 3 החודשים × LT',near(sh.dLT,sh.rateDrop*sh.ltM));
 const sdLT=sdS(sh.clean.slice(0,3))*Math.sqrt(sh.ltM),sdSup=sh.rateDrop*sh.ltvP*sh.ltM;
 const ssStat=Math.ceil(sh.zEff*Math.sqrt(sdLT*sdLT+sdSup*sdSup));
 ok('SS סטטיסטי = ⌈Z × √(σ₃²·LT + (קצב₃·LTV·LT)²)⌉, σ₃ מ-3 החודשים בלבד',near(sh.sdLT,sdLT)&&sh.ssStat===ssStat,`SS ${sh.ssStat}`);
 ok('SS = max(סטטיסטי, ssMin, רצפת מדיניות) · ROP = ⌈ביקוש_LT + SS⌉',sh.sugSS===Math.max(sh.ssStat,sh.ssMin,sh.lowFloor||0)&&sh.sugROP===Math.max(sh.sugSS,Math.ceil(sh.dLT+sh.sugSS)),
   `SS ${sh.sugSS} · ROP ${sh.sugROP}`);
 ok('ROP יורד מתחת ל-SAP (60) — לא מעל',sh.sugROP<60,`ROP ${sh.sugROP}`);
 const r6=sh.clean.slice(0,6).reduce((a,b)=>a+b,0)/6;
 ok('rate, הכיסוי והמגמה לא זזו — נשארים על 6 החודשים',near(sh.rate,r6)&&near(sh.cov,sh.free/sh.rate)&&sh.trendF<1,`rate ${sh.rate.toFixed(2)} · trendF ${sh.trendF}`);
 ok('ROP קצר מזמן האספקה בודק על אותו קצב (ropRate = קצב 3 החודשים)',near(sh.ropRate,sh.rateDrop));
 ok('גבול: 30 מול 120 (0.25 בדיוק) — חל',at.sharp&&near(at.dLT,at.rateDrop*at.ltM),`ROP ${at.sugROP}`);
 ok('גבול: 31 מול 120 — לא חל, הנוסחה כמו קודם',!over.sharp&&near(over.dLT,over.rate*over.ltM)&&over.keys===0,`ROP ${over.sugROP}`);
 ok('בסיס זעיר (5 יח׳ < 6) — לא חל',!tiny.sharp&&tiny.keys===0);
 ok('מכירה בחודש הנוכחי — לא חל',!curm.sharp&&near(curm.dLT,curm.rate*curm.ltM)&&curm.keys===0);
 ok('סטטוס 04 — אין המלצה ואין סימון',!st04.sharp&&st04.sugROP===0&&st04.sugSS===0&&st04.keys===0);
 /* SS_STAT (19.11): לפני — «רצפת ssMin נשמרת: SS = 40». עכשיו פריט עם צריכה: SS סטטיסטי (עם רצפת המדיניות) */
 ok('ssMin=40 אינו רצפה לפריט עם צריכה: SS = max(סטטיסטי, רצפת מדיניות), ROP = ⌈ביקוש_LT + SS⌉',flo.sharp&&flo.statOnly&&flo.sugSS===Math.max(flo.ssStat,flo.lowFloor)&&flo.sugSS<40&&flo.sugROP===Math.max(flo.sugSS,Math.ceil(flo.dLT+flo.sugSS)),`SS ${flo.sugSS} · ROP ${flo.sugROP}`);
 ok('צריכה סדירה — לא חל, הנוסחה כמו קודם',!reg.sharp&&near(reg.dLT,reg.rate*reg.ltM)&&reg.keys===0);

 /* בדוח הסינתטי: כל פריט שאינו בירידה חדה ואינו בדעיכה — dLT = rate × LT */
 const D=await p.evaluate(()=>({n:ALL.length,sh:ALL.filter(x=>x.sharpDrop).length,
  diff:ALL.filter(x=>!x.sharpDrop&&!x.rateDecay&&x.rateW==null&&Math.abs(x.dLT-x.rate*x.ltM)>1e-9).length}));
 ok('בכל פריט אחר בדוח הסינתטי: ביקוש בזמן אספקה = rate × LT',D.diff===0,`${D.n} פריטים · ${D.sh} בירידה חדה`);

 /* DROP_NO_RAISE — דרך classify: פריט בדלת «ירידה בביקוש» שהמנוע מחשב לו ROP מעל SAP */
 const C=await p.evaluate(()=>{
  const hdr=ALL_RAW.hdr,m=mapCols(hdr),tpl=ALL_RAW.rows[0];
  const FREE=40;const mk=(pn,mo,rop)=>{const r=tpl.slice();const set=(i,v)=>{if(i>=0)r[i]=v};
   set(m.pn,pn);set(m.status,'01');set(m.statusTxt,'פעיל');set(m.mrp,'ND');set(m.srv,80);
   set(m.ss,0);set(m.rop,rop);set(m.ssMin,0);set(m.lt,60);set(m.ltAdd,0);
   set(m.price,100);set(m.currency,'USD');set(m.free,FREE);set(m.stock,FREE);set(m.po,0);
   if(m.openDel>=0)set(m.openDel,0);if(m.transfer>=0)set(m.transfer,0);if(m.cust>=0)set(m.cust,0);
   m.months.forEach((ix,k)=>set(ix,mo[k]||0));
   const d=mo.reduce((a,v)=>a+v,0);[m.y0,m.y1,m.y2].forEach((ix,k)=>set(ix,[d,d+40,400][k]));
   if(m.curMonth>=0)set(m.curMonth,0);return r};
  /* ירידה מתונה (לא חדה): ROP המנוע נשאר מעל SAP */
  const mo=[12,12,12,12,12,40,40,40,40,40,40];
  const items=build([mk('C-UP',mo,2),mk('C-DOWN',mo,400)],hdr);
  classify(items);
  return items.map(x=>({cat:x.cat,act:(x.act||[])[0]||'',sugROP:x.sugROP,rop:x.rop,sharp:!!x.sharpDrop}))});
 const [up,down]=C;
 ok('DROP_NO_RAISE: «ירידה בביקוש» עם ROP מחושב מעל SAP — אין «העלאת», ההוראה להשאיר',
   up.cat==='ירידה בביקוש'&&up.sugROP>up.rop&&!/העלאת/.test(up.act)&&/^להשאיר את נקודת ההזמנה ב-SAP \(2\)/.test(up.act),`${up.cat} · ${up.act}`);
 ok('DROP_NO_RAISE: הכיוון המקטין לא השתנה — «לשקול הקטנת»',
   down.cat==='ירידה בביקוש'&&down.sugROP<down.rop&&down.act===`לשקול הקטנת נקודת הזמנה מ-400 ל-${down.sugROP}`,`${down.cat} · ${down.act}`);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0);
})().catch(e=>{console.error(e);process.exit(1)});
