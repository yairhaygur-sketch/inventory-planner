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
 ok('מסלולים: short/m1/cap/fix',st.tracks==='short,m1,cap,fix',st.tracks);
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

 // ── פריטי PD: נרכשים לפי דרישה, ולכן אפס המלצות מלאי ──
 const pd=await p.evaluate(()=>{
  const smp=ALL.slice(0,900);
  const pds=smp.filter(r=>r.isPDItem), rest=smp.filter(r=>!r.isPDItem);
  return {n:pds.length,tot:smp.length,
   noSS:  pds.every(r=>r.sugSS===0),
   noROP: pds.every(r=>r.sugROP===0),
   srv50: pds.every(r=>r.srv===50),
   restHasSS: rest.some(r=>r.sugSS>0)}});
 ok('פריטי PD מזוהים',pd.n>0,`${pd.n}/${pd.tot} (${Math.round(pd.n/pd.tot*100)}%)`);
 ok('פריט PD אינו מקבל המלצת מלאי ביטחון',pd.noSS);
 ok('פריט PD אינו מקבל המלצת נקודת הזמנה',pd.noROP);
 ok('PD מצומד לרמת שרות 50',pd.srv50);
 ok('פריטים שאינם PD כן מקבלים המלצות',pd.restHasSS);
 ok('כיסוי עדיין מחושב',prot.cov>0,prot.cov+' פריטים');
 ok('מלאי מת/איטי עדיין מסווגים',prot.dead>0||prot.slow>0,`מת ${prot.dead} · איטי ${prot.slow}`);
 ok('הון כלוא מחושב לפי BZ',prot.exc>0&&prot.capUsesPrice,prot.exc+' פריטים');

 // ── מיפוי מאומת מול המסמך העסקי (ZMRP_COLUMNS) ──
 const map=await p.evaluate(()=>{
  const smp=ALL.slice(0,600);
  return {
   noAccCol:   !/n\.indexOf\('אביזרים'\)/.test(document.documentElement.innerHTML),
   accFromU:   smp.some(r=>r.isAcc),
   transfer:   smp.some(r=>r.transfer>0),
   availSum:   smp.every(r=>r.avail===Math.max(0,r.free)+Math.max(0,r.po)+(r.transfer||0)),
   matType:    [...new Set(smp.map(r=>r.matType))].sort().join(','),
   diag:       DIAG.warn.join(' | ')}});
 ok('אביזר מזוהה מהיררכייה3 (U)',map.accFromU);
 ok('החיפוש המת אחרי עמודת "אביזרים" הוסר',map.noAccCol);
 ok('מלאי בהעברה (CK) נקרא',map.transfer);
 ok('זמין = פנוי + רכש + בהעברה',map.availSum);
 ok('סוג חומר ממופה ליבוא/מקומי',/יבוא/.test(map.matType)&&/רכש מקומי/.test(map.matType),map.matType);
 ok('אין אזהרת DIAG על עמודה חסרה',!/לא נמצאה/.test(map.diag),map.diag.slice(0,70)||'(נקי)');
 // ── Z: נבדק ישירות על הפונקציה, לא דרך הדוח. הדוח מכיל 50/80/95 בלבד,
 //    ובדיקה שנשענת עליו הייתה עוברת בריק על כל שאר הטווח. ──
 const zt=await p.evaluate(()=>{
  const legacy={50:0,80:0.84,90:1.28,95:1.65,99:2.33};
  const drift=Object.keys(legacy).filter(k=>zFor(+k)!==legacy[k]);
  const grid=[];for(let v=50;v<=99.9;v=Math.round((v+0.1)*10)/10)grid.push([v,zFor(v)]);
  const nonMono=grid.filter(([v,z],i)=>i>0&&z<grid[i-1][1]-1e-9).map(([v])=>v);
  const spot=[93,96,97,98].map(v=>v+'→'+zFor(v)).join(' ');
  return {drift,nonMono,spot,z95:zFor(95),z98:zFor(98),z50:zFor(50)}});
 ok('Z זהה לטבלה ההיסטורית ב-50/80/90/95/99',zt.drift.length===0,zt.drift.join(',')||'אפס סטיות');
 ok('Z מונוטוני על כל הטווח 50→99.9',zt.nonMono.length===0,zt.nonMono.slice(0,3).join(',')||'500 נקודות');
 ok('Z ב-98 גבוה מ-95 (הבאג המקורי)',zt.z98>zt.z95,`95→${zt.z95} · 98→${zt.z98}`);
 ok('רמות ביניים מקבלות ערך אמיתי ולא 1.28',zt.spot.split(' ').every(x=>+x.split('→')[1]!==1.28),zt.spot);
 ok('רמת שרות 50 נותנת Z=0',zt.z50===0); 
 // מסלול "אוזל החודש" — צריכה ברבעון מול חודש קדימה, ZM בלבד
 const m1=await p.evaluate(()=>{const L=decisionList('m1');
  const c3ok=L.every(r=>r.c3>=M1_MIN3), zm=L.every(r=>!r.isPDItem), sh=L.every(r=>r.free<r.r3);
  const sorted=L.every((r,i)=>i===0||Math.ceil(L[i-1].r3-L[i-1].free)>=Math.ceil(r.r3-r.free));
  /* פריט שהרכש הפתוח מכסה חייב להישאר ברשימה — אין ETA, ולכן הרכש אינו כיסוי */
  const withPO=L.filter(r=>r.po>0&&r.avail>=r.r3).length;
  /* פריט PD שעונה על שאר התנאים ובכל זאת אינו ברשימה */
  const pdSkipped=ALL.filter(r=>r.isPDItem&&r.c3>=M1_MIN3&&r.free<r.r3).length;
  const c3=ALL.filter(r=>r.c3>0).length;
  return {n:L.length,c3ok,zm,sh,sorted,withPO,pdSkipped,c3}});
 ok('צריכת 3 חודשים מחושבת',m1.c3>0,m1.c3+' פריטים עם צריכה ברבעון');
 ok('מסלול "אוזל החודש" — כל פריט מעל סף הצריכה',m1.c3ok,m1.n+' פריטים');
 ok('ורק ZM — פריטי PD מוחרגים',m1.zm&&m1.pdSkipped>0,m1.pdSkipped+' פריטי PD עומדים בתנאים ולא נכללו');
 ok('כל פריט: מלאי פנוי קטן מהצריכה החודשית',m1.sh);
 ok('רכש פתוח אינו מוציא פריט מהרשימה (אין ETA)',m1.withPO>0,m1.withPO+' פריטים עם רכש שמכסה — ונשארו');
 ok('ממוין לפי הכמות החסרה לחודש',m1.sorted);
 const noFollow=await p.evaluate(()=>({
   bad:Q.follow.filter(isM1).length, follow:Q.follow.length,
   /* מסלול m1 יכול לנחות רק בתורי היום: לקוח ממתין, פער כיסוי או מניעת חוסר */
   todayOnly:decisionList('m1').every(r=>Q.prevent.includes(r)||Q.waiting.includes(r)||Q.immediate.includes(r)),
   dist:Object.entries(decisionList('m1').reduce((a,r)=>{const q=Object.keys(Q).find(k=>Q[k].includes(r))||'-';a[q]=(a[q]||0)+1;return a},{})).map(([k,v])=>k+'='+v).join(' '),
   inactive:decisionList('m1').filter(r=>/לא פעיל|גמר המלאי/.test(r.stTxt||'')).length,
   pdSafe:Q.passive.filter(r=>r.isPDItem).length}));
 ok('אף פריט במעקב אספקה אינו חוסר של החודש',noFollow.bad===0,`${noFollow.follow} במעקב`);
 ok('כל פריטי «אוזל החודש» מסווגים לטיפול היום',noFollow.todayOnly,noFollow.dist);
 ok('פריטי PD לא נשאבו לרונג החדש',noFollow.pdSafe>0,noFollow.pdSafe+' פריטי PD ברקע');
 ok('פריט לא פעיל / גמר מלאי אינו "אוזל"',noFollow.inactive===0,noFollow.inactive+' פריטים לא פעילים במסלול');



 // מעבר בין המסלולים
 for(const t of ['short','m1','cap','fix']){
  await p.evaluate(k=>{track=k;render()},t);await p.waitForTimeout(350);
  const n=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
  ok('מסלול '+t+' מרנדר',n>0,n+' שורות');
 }
 await p.evaluate(()=>{track='short';render()});await p.waitForTimeout(300);
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(400);
 const dt=await p.evaluate(()=>document.getElementById('detail').innerText);
 const dcur=await p.evaluate(()=>CURRENT_DETAIL&&CURRENT_DETAIL.currency);
 ok('כרטיס הפריט ללא "הכנסה בסיכון"',!dt.includes('הכנסה בסיכון'));
 ok('כרטיס הפריט ללא "הון נדרש לכיסוי"',!dt.includes('הון נדרש לכיסוי'));
 ok('כרטיס הפריט מציג כמות חסרה',dt.includes('כמות חסרה'));
 ok('כרטיס הפריט מציג את המטבע של הפריט עצמו',!!dcur&&dt.includes(dcur),dcur||'ללא מטבע');

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
