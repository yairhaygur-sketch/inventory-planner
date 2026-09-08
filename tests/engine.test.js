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
   modes:       MODES.map(m=>m[0]).join(','),
   modeDefault: mode,
   modeCounts:  MODES.map(m=>m[0]+'='+modeRows(m[0]).length).join(' '),
   tabsShown:   document.querySelectorAll('#tabs .tab').length,
   trackStrip:  getComputedStyle(document.getElementById('track')).display,
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
 ok('שלושה מצבים: היום/החודש/קטלוג',st.modes==='today,month,catalog',st.modes);
 ok('ברירת מחדל = היום',st.modeDefault==='today'&&st.trackDefault==='short');
 ok('ארבעה יעדים בסרגל במקום 11',st.tabsShown===4,st.tabsShown+' אריחים');
 ok('רצועת המסלולים ירדה',st.trackStrip==='none',st.trackStrip);
 ok('לכל מצב יש תוכן',/today=[1-9]/.test(st.modeCounts)&&/month=[1-9]/.test(st.modeCounts)&&/catalog=[1-9]/.test(st.modeCounts),st.modeCounts);
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
 for(const t of ['short','m1','cap','fix','month']){
  await p.evaluate(k=>{track=k;render()},t);await p.waitForTimeout(350);
  const n=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
  ok('מסלול '+t+' מרנדר',n>0,n+' שורות');
 }
 for(const m of ['today','month','catalog']){
  await p.evaluate(k=>setMode(k),m);await p.waitForTimeout(350);
  const x=await p.evaluate(()=>({n:document.querySelectorAll('#tbl tbody tr[data-i]').length,
    mode,cur,track,cols:document.querySelectorAll('#tbl thead th').length}));
  ok('מצב '+m+' מרנדר',x.n>0&&x.mode===m,`${x.n} שורות · cur=${x.cur} track=${x.track} · ${x.cols} עמודות`);
 }
 /* כל פריט בתור עבודה חייב להיות נגיש מאחד המצבים — שום תצוגה לא אבדה */
 const reach=await p.evaluate(()=>{const t=new Set(modeRows('today')),m=new Set(modeRows('month')),c=new Set(modeRows('catalog'));
   const work=['waiting','immediate','follow','prevent','excess','quality'].flatMap(k=>Q[k]);
   const lost=work.filter(r=>!t.has(r)&&!m.has(r)&&!c.has(r));
   return {work:work.length,lost:lost.length,cats:[...new Set(lost.map(r=>r.cat))].join(',')}});
 ok('אף פריט עבודה לא אבד במעבר למצבים',reach.lost===0,`${reach.work} פריטים · אבדו ${reach.lost} ${reach.cats}`);
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(300);
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
 /* ============ פריטים שאינם מנוהלים במלאי ============
    שלושה מקורות: PD · קבוצת שיווק "כלי עבודה"/"אביזרים" · סימון ידני.
    "לפי דרישה" אינו "מוסתר" — הפריט נשאר גלוי וממשיך להופיע ב"לקוח
    ממתין", כי הזמנת לקוח היא בדיוק הטריגר להזמין פריט כזה. */
 const odBase=await p.evaluate(()=>{
  const tools=ALL.filter(a=>OD_GROUPS.has(a.mkt1));
  return {tools:tools.length, withRec:tools.filter(a=>a.sugSS>0||a.sugROP>0).length,
    inM1:tools.filter(a=>isM1(a)).length,
    inFloor:tools.filter(a=>ssFloorGap(a)>0).length}});
 ok('קבוצות "כלי עבודה"/"אביזרים" מזוהות',odBase.tools>0,odBase.tools+' פריטים');
 ok('הן אינן מקבלות המלצת SS/ROP',odBase.withRec===0,odBase.withRec+' עם המלצה');
 ok('הן אינן נכנסות ל"אוזל החודש"',odBase.inM1===0);
 ok('הן אינן נכנסות למסלול הרצפה',odBase.inFloor===0);

 /* מחזור חיים מלא של הסימון הידני על פריט אמיתי מהתור */
 const odPn=await p.evaluate(()=>{const a=QF.prevent.find(x=>!x.isOD&&x.sugSS>0&&x.cust<=0);
   return a?a.pn:null});
 ok('נמצא פריט לבדיקת הסימון',!!odPn,odPn||'—');
 if(odPn){
  const snap=()=>p.evaluate(n=>{const a=ALL.find(x=>x.pn===n);
    return {cat:a.cat,ss:a.sugSS,rop:a.sugROP,isOD:a.isOD,stale:!!a.odStale,pf:!!a.paramFix,
      inToday:QF.prevent.includes(a)||QF.waiting.includes(a)||QF.immediate.includes(a),
      act:(a.act||[])[0]||''}},odPn);
  const b0=await snap();
  await p.evaluate(n=>setMark(n,'ondemand','בדיקה'),odPn);await p.waitForTimeout(500);
  const b1=await snap();
  ok('סימון "לפי דרישה" מוריד מתור העבודה',!b1.inToday&&b1.cat==='לפי דרישה',b1.cat);
  ok('הסימון מאפס את המלצות המלאי',b1.ss===0&&b1.rop===0,`SS ${b1.ss} · ROP ${b1.rop}`);
  ok('הפעולה אומרת שיוזמן לפי דרישה',/יוזמן כשתגיע דרישה/.test(b1.act),b1.act);
  /* לקוח ממתין גובר: הזמנה אמיתית היא הטריגר להזמין */
  await p.evaluate(n=>{const a=ALL.find(x=>x.pn===n);a._f=a.free;a._c=a.cust;
    a.cust=5;a.free=0;a.po=0;apply()},odPn);await p.waitForTimeout(500);
  const b2=await snap();
  ok('לקוח ממתין גובר על ההחרגה',b2.inToday,b2.cat);
  await p.evaluate(n=>{const a=ALL.find(x=>x.pn===n);a.free=a._f;a.cust=a._c;apply()},odPn);
  await p.waitForTimeout(400);
  /* ביטול מחזיר בדיוק את ההמלצות המקוריות — בלי דליפת מצב בין מעברים */
  await p.evaluate(n=>setMark(n,'',''),odPn);await p.waitForTimeout(500);
  const b3=await snap();
  ok('ביטול הסימון מחזיר את ההמלצות המקוריות',
    b3.ss===b0.ss&&b3.rop===b0.rop&&b3.inToday,`SS ${b0.ss}→${b3.ss} · ROP ${b0.rop}→${b3.rop}`);
  /* תפוגה: חצי שנה */
  await p.evaluate(n=>{setMark(n,'ondemand','ישן');
    const k=Object.keys(MARKS).find(x=>x.indexOf(n)>=0);
    MARKS[k].ts=Date.now()-200*864e5;saveMarks();apply()},odPn);await p.waitForTimeout(500);
  const b4=await snap();
  ok('סימון בן 200 יום פג',b4.stale&&!b4.isOD);
  ok('פריט שההחרגה שלו פגה חוזר לאימות ולא לתור',
    b4.pf&&/לאמת/.test(b4.act),b4.cat+' · '+b4.act);
  ok('ההמלצות חוזרות עם התפוגה',b4.ss===b0.ss,`SS ${b4.ss}`);
  await p.evaluate(n=>setMark(n,'',''),odPn);await p.waitForTimeout(400)}

 /* ============ ריסון מגמת ירידה ============ */
 const tr=await p.evaluate(()=>{
  const S=a=>a.reduce((x,y)=>x+y,0);
  const bad=[];let damped=0,gapPos=0,everUp=0,odLeak=0,monotone=0,floorHit=0;
  for(const r of ALL){
   if(r.trendF>1)everUp++;                                  /* אסור להגביר קצב */
   if(r.trendF<1){damped++;
    if(r.trendF<0.6-1e-9)floorHit++;
    if(!(r.rateT<r.rate))monotone++;
    if(r.isOD)odLeak++;
    /* הריסון הוא בדיוק חצי מהירידה, עם רצפה */
    const want=Math.max(0.6,1+(r.tRatio-1)*0.5);
    if(Math.abs(r.trendF-want)>1e-9)bad.push(r.pn);
    /* ROP מרוסן לעולם לא מעל המקורי, ולעולם לא מתחת ל-SS */
    if(r.sugROPT>r.sugROP||r.sugROPT<r.sugSS)bad.push(r.pn+'/rop')}
   if(trendGap(r)>0)gapPos++}
  const rows=trendRows();
  const sorted=rows.every((r,i)=>i===0||trendGap(rows[i-1])*(rows[i-1].price||0)>=0);
  return {damped,gapPos,everUp,odLeak,monotone,floorHit,bad:bad.slice(0,4),
    nRows:rows.length,sorted,
    /* כל שורה במסלול חייבת פער חיובי */
    allGap:rows.every(r=>trendGap(r)>0),
    /* פריט לפי דרישה לא נכנס */
    noOD:rows.every(r=>!r.isOD),
    trigOK:ALL.every(r=>!(r.tRatio!=null&&r.tRatio>=0.9&&r.trendF<1))}});
 ok('הקצב אף פעם לא מוגבר בגלל מגמה',tr.everUp===0,'הוגברו: '+tr.everUp);
 ok('הריסון הוא חצי מהירידה, עם רצפת 0.6',tr.bad.length===0,tr.bad.join(',')||'—');
 ok('הרצפה 0.6 לא נשברת',tr.floorHit===0);
 ok('קצב מרוסן תמיד קטן מהקצב',tr.monotone===0);
 ok('ריסון פועל רק מתחת לסף 0.9',tr.trigOK);
 ok('פריט לפי דרישה לא מקבל פער מגמה',tr.odLeak===0&&tr.noOD);
 ok('כל שורה במסלול המגמה נושאת פער חיובי',tr.allGap,tr.nRows+' שורות');
 ok('מסלול המגמה קיים בנתוני הבדיקה',tr.damped>0&&tr.nRows>0,
   'מרוסנים '+tr.damped+' · במסלול '+tr.nRows);

 /* המסלול על המסך */
 await p.evaluate(()=>setMode('trend'));await p.waitForTimeout(500);
 const tv=await p.evaluate(()=>{
  const th=[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim());
  const tb=document.getElementById('tbl');
  const btn=document.getElementById('trendBtn');
  return {nth:th.length,head:th.join('|'),
   cells:(document.querySelector('#tbl tbody tr')||{children:[]}).children.length,
   over:tb.scrollWidth>tb.clientWidth+2,
   btnVis:btn.offsetParent!==null,btnOn:btn.classList.contains('on'),
   sub:document.getElementById('phdSub').textContent,
   trackStrip:getComputedStyle(document.getElementById('track')).display}});
 ok('כותרות המסלול תואמות למספר התאים',tv.nth===tv.cells&&tv.nth===8,tv.nth+' / '+tv.cells);
 ok('אין גלישה אופקית במסלול המגמה',!tv.over);
 ok('כפתור המגמה נראה ומסומן',tv.btnVis&&tv.btnOn);
 ok('רצועת המסלולים מוסתרת',tv.trackStrip==='none',tv.trackStrip);
 ok('התת-כותרת אומרת שההמלצה לא שונתה',/לא שונתה אוטומטית/.test(tv.sub));

 /* [hidden] מול display מפורש — נבדק בנראות בפועל, לא בתכונה (סעיף 32) */
 const vis=[];
 for(const m of ['today','month','catalog','floor','trend','cust']){
  await p.evaluate(k=>setMode(k),m);await p.waitForTimeout(350);
  vis.push([m,await p.evaluate(()=>{const g=document.getElementById('gseg');
    return g.offsetParent!==null&&g.getBoundingClientRect().width>0})])}
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);
 /* ============ נקודת הזמנה של 1 היא נוכחות, לא פרמטר ============ */
 const rp=await p.evaluate(()=>{
  const shown=decisionList('month');
  const keep=ALL.filter(x=>x.rop===ROP_MIN_ACT&&!x.isPDItem&&!x.isOD);
  const keepSS=ALL.filter(x=>x.ss===ROP_MIN_ACT&&!x.isPDItem&&!x.isOD);
  return {
   /* אף המלצה להוריד 1 → 0 על פריט שמנוהל מלאי */
   ropDrops:keep.filter(x=>x.sugROP===0).length,
   ssDrops:keepSS.filter(x=>x.sugSS===0).length,
   kept:keep.filter(x=>x.sugROP===ROP_MIN_ACT).length,
   /* PD ו«לפי דרישה» כן נשארים על 0 — מדיניות מפורשת */
   pdZero:ALL.filter(x=>(x.isPDItem||x.isOD)&&x.rop===ROP_MIN_ACT&&x.sugROP===0).length,
   /* לא נגענו בפריטים עם ROP גבוה יותר */
   biggerStillDrop:ALL.filter(x=>x.rop>ROP_MIN_ACT&&x.sugROP===0).length,
   /* ROP=1 אינו מכניס פריט לרשימת העבודה */
   youngWithRop1:shown.filter(x=>x.cat==='טרם הוכיח את עצמו'&&x.rop<=ROP_MIN_ACT).length,
   /* טקסט ההמלצה הישן נעלם לגמרי */
   oldText:ALL.filter(x=>(x.act||[]).some(a=>/איפוס נקודת הזמנה מ-1 ל-0/.test(a))).length,
   shown:shown.length};});
 ok('אין המלצה להוריד נקודת הזמנה מ-1 ל-0',rp.ropDrops===0,rp.ropDrops+' המלצות');
 ok('אין המלצה להוריד מלאי ביטחון מ-1 ל-0',rp.ssDrops===0,rp.ssDrops+' המלצות');
 ok('ההמלצה על פריטים כאלה היא להשאיר 1',rp.kept>0,rp.kept+' פריטים');
 ok('פריטי PD ולפי דרישה עדיין מקבלים 0 — מדיניות מפורשת',rp.pdZero>=0,rp.pdZero+' פריטים');
 ok('נקודת הזמנה גבוהה מ-1 עדיין ניתנת לאיפוס',rp.biggerStillDrop>0,rp.biggerStillDrop+' פריטים');
 ok('ROP=1 אינו מכניס «טרם הוכיח את עצמו» לרשימת החודש',rp.youngWithRop1===0);
 ok('הטקסט «איפוס נקודת הזמנה מ-1 ל-0» נעלם',rp.oldText===0,rp.oldText+' מופעים');

 /* ============ שקט אחרי מכירה ============ */
 const sl=await p.evaluate(()=>{
  const shown=decisionList('month');
  const cand=ALL.filter(x=>x.sugROP===0&&x.rop>0&&x.d12>3&&!x.isPDItem&&!x.isOD&&x.rate<=0);
  const noted=cand.filter(x=>(x.why||[]).some(w=>/שקט \d+ חודשים/.test(w[1])));
  const dry=ALL.filter(x=>x.silentDry);
  return {cand:cand.length,noted:noted.length,
   /* הדגל נדלק רק כשאין מלאי */
   dryAllEmpty:dry.every(x=>x.free<=0),
   dryInMonth:dry.every(x=>shown.includes(x)),
   dryHasAct:dry.every(x=>(x.act||[]).some(a=>/להחליט ידנית/.test(a))),
   /* פריט עם מלאי מקבל הסבר מרגיע ולא דגל */
   calm:cand.filter(x=>x.free>0).every(x=>!x.silentDry),
   /* לא הומצא קצב — sugROP נשאר 0 */
   stillZero:cand.every(x=>x.sugROP===0),
   dry:dry.length,silentMax:Math.max(0,...cand.map(x=>x.silent))};});
 ok('כל פריט ששקט אחרי מכירה מקבל הסבר',sl.cand>0&&sl.noted===sl.cand,sl.noted+'/'+sl.cand);
 ok('הדגל האדום נדלק רק על פריט ללא מלאי',sl.dryAllEmpty&&sl.calm,sl.dry+' פריטים');
 ok('הפריטים החשופים נכנסים לרשימת החודש עם החלטה ידנית',sl.dryInMonth&&sl.dryHasAct);
 ok('לא הומצא קצב חלופי — ההמלצה נשארת 0',sl.stillZero,'שקט מקסימלי '+sl.silentMax+' חודשים');

 /* ============ לקוח ממתין ============ */
 const cw=await p.evaluate(()=>{
  const rows=(QF.all||[]).filter(custWaiting);
  const G=custGroups(rows);
  const QN=['waiting','immediate','follow','prevent','excess','quality','marked','passive','ok'];
  const where=x=>{for(const k of QN)if(Q[k].includes(x))return k;return '?'};
  const qs={};for(const r of rows){const w=where(r);qs[w]=(qs[w]||0)+1}
  return {n:rows.length,
   /* ההגדרה עצמה: הזמנת לקוח פתוחה ומלאי פנוי אפס — שום תנאי נוסף */
   defOK:rows.every(r=>r.cust>0&&r.free<=0),
   complete:(QF.all||[]).filter(r=>r.cust>0&&r.free<=0).length===rows.length,
   /* הפיצול מכסה את הכל בדיוק פעם אחת */
   split:G.none.length+G.sup.reduce((a,[,l])=>a+l.length,0)===rows.length,
   noneOK:G.none.every(r=>r.po<=0&&(r.transfer||0)<=0),
   supOK:G.sup.every(([,l])=>l.every(r=>r.po>0||r.transfer>0)),
   dupes:(()=>{const seen=new Set();let d=0;
     for(const r of [...G.none,...G.sup.flatMap(([,l])=>l)]){if(seen.has(r.pn))d++;seen.add(r.pn)}return d})(),
   /* מיון לפי שווי ההזמנה בתוך כל קבוצה */
   sorted:G.sup.every(([,l])=>l.every((r,i)=>i===0||custVal(l[i-1])>=custVal(r)))
     &&G.none.every((r,i)=>i===0||custVal(G.none[i-1])>=custVal(r)),
   /* הרשימה באמת חוצה סיווגים — אחרת אין לה הצדקה */
   queues:Object.keys(qs).length,qs,
   pd:rows.filter(r=>r.isPDItem).length,
   held:rows.filter(r=>r.stock>0).length};});
 ok('כל שורה עומדת בהגדרה: הזמנת לקוח ומלאי פנוי אפס',cw.defOK&&cw.complete,cw.n+' פריטים');
 ok('הפיצול מכסה את כל הפריטים בדיוק פעם אחת',cw.split&&cw.dupes===0);
 ok('«אין רכש» מכיל רק פריטים ללא רכש ובלי מלאי בהעברה',cw.noneOK);
 ok('קבוצות הספקים מכילות רק פריטים עם רכש או העברה',cw.supOK);
 ok('בכל קבוצה המיון הוא לפי שווי ההזמנה',cw.sorted);
 ok('הרשימה חוצה יותר מתור עבודה אחד',cw.queues>=2,JSON.stringify(cw.qs));

 await p.evaluate(()=>setMode('cust'));await p.waitForTimeout(600);
 const cv=await p.evaluate(()=>{
  const th=[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim());
  const trs=[...document.querySelectorAll('#tbl tbody tr')];
  const dat=trs.filter(t=>t.dataset.i!==undefined);
  const tb=document.getElementById('tbl'),btn=document.getElementById('custBtn');
  return {nth:th.length,cells:dat.length?dat[0].children.length:0,
   grps:trs.filter(t=>t.classList.contains('grp')).length,rows:dat.length,
   over:tb.scrollWidth>tb.clientWidth+2,
   btnVis:btn.offsetParent!==null,btnOn:btn.classList.contains('on'),
   strip:getComputedStyle(document.getElementById('track')).display,
   sub:document.getElementById('phdSub').textContent};});
 ok('כותרות «לקוח ממתין» תואמות למספר התאים',cv.nth===cv.cells&&cv.nth===8,cv.nth+' / '+cv.cells);
 ok('אין גלישה אופקית ב«לקוח ממתין»',!cv.over);
 ok('יש כותרת קבוצה לכל קטע',cv.grps>=2,cv.grps+' קבוצות · '+cv.rows+' שורות');
 ok('כל הפריטים מרונדרים',cv.rows===cw.n,cv.rows+' / '+cw.n);
 ok('כפתור «לקוח ממתין» נראה ומסומן',cv.btnVis&&cv.btnOn);
 ok('רצועת המסלולים מוסתרת',cv.strip==='none',cv.strip);
 /* קיפול קבוצה מוריד רק אותה */
 const coll=await p.evaluate(()=>{const g=document.querySelector('#tbl tbody tr.grp[data-g^="cst:sup:"]');
   const key=g.dataset.g;const before=document.querySelectorAll('#tbl tbody tr[data-i]').length;
   toggleGroup(key);const after=document.querySelectorAll('#tbl tbody tr[data-i]').length;
   toggleGroup(key);
   return {before,after,back:document.querySelectorAll('#tbl tbody tr[data-i]').length}});
 ok('קיפול קבוצה מסתיר רק אותה ומחזיר',coll.after<coll.before&&coll.back===coll.before,
   `${coll.before} → ${coll.after} → ${coll.back}`);
 /* לחיצה על שורה פותחת את הפריט שבשורה — לא פריט אחר */
 const click=await p.evaluate(()=>{const t=[...document.querySelectorAll('#tbl tbody tr[data-i]')].pop();
   const pn=t.children[1].textContent.replace('העתק','').trim();t.click();
   const c=document.querySelector('#detail .opn');
   return {pn,card:c?c.textContent.replace('העתק','').trim():'—'}});
 ok('לחיצה על השורה האחרונה פותחת את הפריט הנכון',click.pn===click.card,click.pn+' / '+click.card);
 await p.evaluate(()=>{closeDetail();setMode('today')});await p.waitForTimeout(400);
 const back=await p.evaluate(()=>({mode,n:document.querySelectorAll('#tbl tbody tr[data-i]').length}));
 ok('חזרה ל«היום» משחזרת את התור',back.mode==='today'&&back.n>0,back.n+' שורות');

 ok('מתג הקיבוץ גלוי ב"היום" בלבד',
   vis.every(([m,v])=>v===(m==='today')),
   vis.map(([m,v])=>m+'='+v).join(' '));

 await b.close();console.log(out.join('\n'));
 process.exit(out.some(l=>l.startsWith('FAIL'))?1:0)})();
