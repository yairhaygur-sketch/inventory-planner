const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];const ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html'));
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(2600);

 const st=await p.evaluate(()=>{
  const smp=ALL.slice(0,400);
  return {
   noSell:      smp.every(r=>r.sell===undefined),
   noExpRev:    smp.every(r=>r.expRev===undefined),
   noExpShort:  smp.every(r=>r.expShort===undefined),
   riskIsCapOnly: smp.every(r=>(r.risk||0)===(r.expCap||0)),
   shortByQty:  smp.some(r=>r.miss>0),
   noRevUnpr:   smp.every(r=>r.revUnpriced===undefined),
   noLostPast:  smp.every(r=>r.lostPast===undefined),
   noSellMiss:  smp.every(r=>r.sellMissing===undefined),
   hasCurrency: smp.some(r=>r.currency==='USD'),
   priceIsBZ:   smp.filter(r=>r.price>0).length,
   tracks:      TRACKS.map(t=>t[0]).join(','),
   trackDefault:track,
   diag:        DIAG.warn.join(' | '),
   kpiLabels:   [...document.querySelectorAll('#kpis .kpi .l')].map(e=>e.textContent),
  }});
 ok('שדה sell הוסר מכל פריט',st.noSell);
 ok('expRev הוסר',st.noExpRev);
 ok('expShort הוסר — אין מדד כספי לחוסרים',st.noExpShort);
 ok('risk = הון כלוא בלבד',st.riskIsCapOnly);
 ok('חוסרים נמדדים בכמות',st.shortByQty);
 ok('revUnpriced הוסר',st.noRevUnpr);
 ok('lostPast הוסר',st.noLostPast);
 ok('sellMissing הוסר',st.noSellMiss);
 ok('BE (מטבע FOB) נקרא ונשמר',st.hasCurrency);
 ok('BZ (מחיר FOB) הוא מקור המחיר',st.priceIsBZ>0,st.priceIsBZ+' פריטים עם מחיר');
 ok('מסלולים: short/cap/fix',st.tracks==='short,cap,fix',st.tracks);
 ok('ברירת מחדל = short',st.trackDefault==='short');
 ok('אין אריח "מכירות אבודות"',!st.kpiLabels.some(l=>l.includes('מכירות אבודות')),st.kpiLabels.length+' אריחים');
 ok('אזהרת DIAG על FOB בלבד',!/מחיר מכירה/.test(st.diag),st.diag.slice(0,90)||'(אין אזהרות)');

 // הלוגיקה המוגנת לא זזה
 const prot=await p.evaluate(()=>{
  const smp=ALL.slice(0,400);
  const withSS=smp.filter(r=>r.sugSS>0).length, withROP=smp.filter(r=>r.sugROP>0).length;
  const cov=smp.filter(r=>r.covA!==null).length;
  const dead=smp.filter(r=>r.ageTier==='dead').length, slow=smp.filter(r=>r.ageTier==='slow').length;
  const exc=smp.filter(r=>r.expCap>0).length;
  return {withSS,withROP,cov,dead,slow,exc,
    capUsesPrice:smp.filter(r=>r.expCap>0).every(r=>r.price>0)}});
 ok('SS/ROP עדיין מחושבים',prot.withSS>0&&prot.withROP>0,`SS ${prot.withSS} · ROP ${prot.withROP}`);
 ok('כיסוי עדיין מחושב',prot.cov>0,prot.cov+' פריטים');
 ok('מלאי מת/איטי עדיין מסווגים',prot.dead>0||prot.slow>0,`מת ${prot.dead} · איטי ${prot.slow}`);
 ok('הון כלוא מחושב לפי BZ',prot.exc>0&&prot.capUsesPrice,prot.exc+' פריטים');

 // מעבר בין המסלולים
 for(const t of ['short','cap','fix']){
  await p.evaluate(k=>{track=k;render()},t);await p.waitForTimeout(350);
  const n=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
  ok('מסלול '+t+' מרנדר',n>0,n+' שורות');
 }
 await p.evaluate(()=>{track='short';render()});await p.waitForTimeout(300);
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(400);
 const dt=await p.evaluate(()=>document.getElementById('detail').innerText);
 ok('כרטיס הפריט ללא "הכנסה בסיכון"',!dt.includes('הכנסה בסיכון'));
 ok('כרטיס הפריט ללא "הון נדרש לכיסוי"',!dt.includes('הון נדרש לכיסוי'));
 ok('כרטיס הפריט מציג כמות חסרה',dt.includes('כמות חסרה'));
 ok('כרטיס הפריט מציג מטבע',dt.includes('USD'));

 // ייצוא
 await p.evaluate(()=>{window.__x=[];const o=XLSX.writeFile;XLSX.writeFile=(wb)=>{window.__x=wb};});
 await p.click('#exportXls');await p.waitForTimeout(900);
 const xh=await p.evaluate(()=>{const wb=window.__x;if(!wb||!wb.SheetNames)return null;
   const ws=wb.Sheets[wb.SheetNames[0]];return XLSX.utils.sheet_to_json(ws,{header:1})[0]||[]});
 ok('ייצוא: אין עמודות הכנסה/מכירה',xh&&!xh.some(h=>/הכנסה בסיכון|מחיר מכירה|שאבדו|הון נדרש/.test(h)),
    xh?xh.filter(h=>/FOB|הון|חסרה/.test(h)).join(' · '):'(לא נקרא)');

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 await p.screenshot({path:SD+'/06-bz.png'});
 await b.close();console.log(out.join('\n'));
 process.exit(out.some(l=>l.startsWith('FAIL'))?1:0)})();
