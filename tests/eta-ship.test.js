/* ============ ETA_SHIP — מספר המשלוח מדוח ה-ETA (החלטת המעתד 21.11) ============
   «משלוח» (בגיליון של המעתד עמודה A) הוא מספר המשלוח; «אספקה» הוא מספר אספקה בן 10 ספרות,
   וכמה אספקות יושבות במשלוח אחד. עד כאן הכלי קרא רק את «אספקה» והציג אותו בכותרת «משלוח».
   עכשיו מוצג מספר המשלוח — בלוח ההגעות ובכרטיס (בשני הכלים), בבועית «הגעה» של מסכי החוסרים
   ובבועית «אספקה הבאה (ETA)» בתצוגת 4 הנושאים. בקובץ בלי עמודת «משלוח» — מספר האספקה, בשם
   «אספקה». במסך בלבד: אף קובץ שיוצא מהכלי לא השתנה (בחירת המעתד). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');const XLSX=require('xlsx');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx'),ETAX=path.join(SD,'zmrp-demo-eta.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
/* מימוש עצמאי: הקובץ עצמו, לפי שמות הכותרות — מק״ט → «כמות|משלוח|אספקה» */
const want=(()=>{const wb=XLSX.readFile(ETAX),a=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:''});
 const h=a[0].map(String),ix=k=>h.indexOf(k),m={};
 for(const r of a.slice(1)){const pn=String(r[ix('חומר')]).replace(/[‎‏⁦-⁩\s]/g,'').toUpperCase();if(!pn)continue;
  (m[pn]=m[pn]||[]).push(`${+r[ix('כמות באספקה')]}|${r[ix('משלוח')]}|${r[ix('אספקה')]}`)}
 for(const k in m)m[k].sort();return {m,colA:h[0]}})();
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 const wait=ms=>p.waitForTimeout(ms);
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.evaluate(()=>{try{localStorage.clear()}catch(_){}});
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB,null,{timeout:180000});
 await p.setInputFiles('#fe',ETAX);await wait(2500);

 /* 1 · הקריאה: העמודה «משלוח» לפי שם הכותרת, שדה חמישי בכל שורת ETA */
 const P=await p.evaluate(w=>{const got={};for(const k in ETA.byPn)got[k]=ETA.byPn[k].map(l=>`${l[0]}|${l[4]}|${l[2]}`).sort();
  const ks=Object.keys(w);return {has:ETA.hasShip===true,same:ks.length===Object.keys(got).length&&ks.every(k=>JSON.stringify(got[k])===JSON.stringify(w[k])),
   n:Object.values(ETA.byPn).flat().length,ships:new Set(Object.values(ETA.byPn).flat().map(l=>l[4])).size}},want.m);
 ok('בקובץ ההדגמה עמודה A היא «משלוח», כמו בגיליון של המעתד',want.colA==='משלוח',want.colA);
 ok('הכלי קורא את «משלוח»: לכל שורת ETA — הכמות, המשלוח והאספקה כמו בקובץ (מימוש עצמאי)',P.has&&P.same&&P.n>0,`${P.n} שורות · ${P.ships} משלוחים`);

 /* 2 · לוח ההגעות בכלי המלא: העמודה «משלוח» = מספר המשלוח של השורה */
 await p.evaluate(()=>setMode('line'));await wait(400);await p.click('#arrBtn');await wait(500);
 await p.click('#arrb details.arr:not(.past) summary');await wait(300);
 const A=await p.evaluate(()=>{const d=document.querySelector('#arrb details.arr[open]'),x=ARR_D[+d.dataset.ad];
  const th=[...d.querySelectorAll('thead th')].map(t=>t.textContent),i=th.indexOf('משלוח');
  const trs=[...d.querySelectorAll('tbody tr')];
  const cells=trs.map((tr,j)=>({t:i<0?'':tr.children[i].textContent.trim(),e:x.sorted[j]}));
  const lineOf=e=>(ETA.byPn[etaKey(e.r.pn)]||[]).filter(l=>l[1]===x.d&&l[0]===e.q&&(l[3]||'')===(e.po||''));
  return {th:th.join('|'),i,n:cells.length,ok:cells.every(c=>c.t===c.e.ref&&lineOf(c.e).some(l=>l[4]===c.t)&&/^SH-\d+$/.test(c.t)),
   noDlv:!/419000\d{4}/.test(d.textContent)}});
 ok('לוח ההגעות (כלי מלא): כותרת «משלוח», ובכל שורה מספר המשלוח של שורת ה-ETA',A.i>=0&&A.n>0&&A.ok,`${A.n} שורות · ${A.th}`);
 ok('לוח ההגעות (כלי מלא): מספר האספקה לא מוצג',A.noDlv);
 await p.click('#arrBtn');await wait(300);

 /* 3 · הכרטיס בכלי המלא: «משלוח X» ליד התאריך */
 const C=await p.evaluate(()=>{const r=QF.all.find(x=>(ETA.byPn[etaKey(x.pn)]||[]).length);detail(r);
  const li=[...document.querySelectorAll('#detail .etad li')].map(x=>x.textContent.replace(/\s+/g,' ').trim());
  const ls=ETA.byPn[etaKey(r.pn)];closeDetail();
  return {li,ok:ls.every(l=>li.some(t=>t.startsWith(tpdDate(l[1]))&&t.includes('משלוח '+l[4]))),noDlv:li.every(t=>!/419000\d{4}/.test(t))}});
 ok('כרטיס (כלי מלא): «משלוח» ומספר המשלוח ליד כל תאריך',C.li.length>0&&C.ok,C.li[0]);
 ok('כרטיס (כלי מלא): מספר האספקה לא מוצג',C.noDlv);

 /* 4 · בועית «הגעה» במסכי החוסרים: מספרי המשלוח של ההגעות הקרובות ושל השורות שתאריכן עבר */
 const T=await p.evaluate(()=>{const t=etaToday(),tip=r=>{const d=document.createElement('table');d.innerHTML='<tr>'+doorEtaCell(r)+'</tr>';return d.querySelector('td').title};
  const L=QF.all.filter(r=>(ETA.byPn[etaKey(r.pn)]||[]).length);
  const fut=L.filter(r=>(r.etaQty||0)>0),past=L.filter(r=>(r.etaPast||0)>0);
  const shipsOf=(r,f)=>[...new Set(ETA.byPn[etaKey(r.pn)].filter(l=>f(l[1])).map(l=>l[4]))];
  return {nf:fut.length,np:past.length,
   fut:fut.every(r=>{const s=tip(r);return shipsOf(r,d=>d>=t).every(x=>s.includes(x))&&/משלוח/.test(s)}),
   past:past.every(r=>{const s=tip(r);return shipsOf(r,d=>d<t).every(x=>s.includes('('+'משלוח '+x)||s.includes(x))}),
   sample:fut[0]?tip(fut[0]):'',noDlv:L.every(r=>!/419000\d{4}/.test(tip(r)))}});
 ok('בועית «הגעה»: מספרי המשלוח של ההגעות הקרובות',T.nf>0&&T.fut,`${T.nf} פריטים · ${T.sample.slice(0,80)}`);
 ok('בועית «הגעה»: גם מספרי המשלוח של השורות שתאריכן עבר',T.np>0&&T.past,`${T.np} פריטים`);
 ok('בועית «הגעה»: מספר האספקה לא מוצג',T.noDlv);
 const V=await p.evaluate(()=>{const t=etaToday(),tip=r=>{const d=document.createElement('table');d.innerHTML='<tr>'+shortCovered(r)+'</tr>';return d.querySelector('td').title};
  const L=QF.all.filter(r=>(r.etaQty||0)>0);
  return {n:L.length,ok:L.every(r=>{const s=tip(r);return [...new Set(ETA.byPn[etaKey(r.pn)].filter(l=>l[1]>=t).map(l=>l[4]))].every(x=>s.includes(x))}),
   noDlv:L.every(r=>!/419000\d{4}/.test(tip(r))),sample:L[0]?tip(L[0]):''}});
 ok('בועית «מכוסה ברכש» («חוסרים לטיפול»): מספרי המשלוח של ההגעות הקרובות, בלי מספר האספקה',V.n>0&&V.ok&&V.noDlv,`${V.n} פריטים · ${V.sample.slice(0,70)}`);

 /* 5 · אף קובץ שיוצא מהכלי לא השתנה: אין מספר משלוח באף ייצוא */
 const X=await p.evaluate(()=>{const has=v=>/SH-\d+/.test(String(v));const res={};
  const sheet=m=>{setMode(m);const v=viewExportRows();return v.aoa.some(row=>row.some(has))||v.aoa[0].includes('משלוח')};
  for(const m of ['today','line','stale','follow'])res[m]=sheet(m);
  const er=expediteRows().rows;res.exp=er.some(x=>EXPCOLS.some(c=>{try{return has(c[1](x))}catch(_){return false}}))||EXPCOLS.some(c=>c[0]==='משלוח');
  res.stDoc=QF.all.some(r=>has(SCOL.stDoc.xval(r)));
  return res});
 ok('ייצוא בלי שינוי: אין מספר משלוח בגיליון «התצוגה» (חוסרים, רכש פתוח, באיחור, זירוז ספקים), בדוח הזירוז ובעמודת המסמכים',Object.values(X).every(v=>v===false),JSON.stringify(X));

 /* 6 · קובץ בלי עמודת «משלוח» (וגם דוח שנשמר לפני השינוי): מספר האספקה, בשם «אספקה» */
 const F=await p.evaluate(()=>{const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['אספקה','חומר','כמות באספקה','תארי.אספקה'],['4190009999','PN-X',3,'01.01.2030']]),'א');
  const a=etaParse(wb);
  const wb2=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet([['אספקה','חומר','כמות באספקה','תארי.אספקה','סטאטוס משלוח','משלוח'],['4190009999','PN-X',3,'01.01.2030','X','SH-9']]),'א');
  const c=etaParse(wb2);
  const keep=ETA.hasShip;ETA.hasShip=false;
  const r=QF.all.find(x=>(ETA.byPn[etaKey(x.pn)]||[]).length);detail(r);
  const li=[...document.querySelectorAll('#detail .etad li')].map(x=>x.textContent.replace(/\s+/g,' ').trim());closeDetail();
  const ls=ETA.byPn[etaKey(r.pn)];
  setMode('line');ARR_ON=true;arrSync();const d=document.querySelector('#arrb details.arr');
  if(!d.querySelector('table'))d.insertAdjacentHTML('beforeend',arrDayHTML(+d.dataset.ad));
  const th=[...d.querySelectorAll('thead th')].map(t=>t.textContent).join('|'),x=ARR_D[+d.dataset.ad];
  const cells=[...d.querySelectorAll('tbody tr')].map(tr=>tr.children[3].textContent.trim());
  ARR_ON=false;arrSync();ETA.hasShip=keep;
  return {a:a.hasShip===false&&a.byPn['PN-X'][0][4]==='',c:c.hasShip===true&&c.byPn['PN-X'][0][4]==='SH-9',
   card:ls.every(l=>li.some(t=>t.startsWith(tpdDate(l[1]))&&t.includes('אספקה '+l[2])))&&li.every(t=>!/SH-/.test(t)),
   board:/אספקה/.test(th)&&!/משלוח/.test(th)&&cells.every((t,j)=>t===x.sorted[j].dl&&/^419000\d{4}$/.test(t))}});
 ok('קובץ בלי «משלוח» נטען, והשדה ריק',F.a);
 ok('«משלוח» נקרא לפי שם הכותרת גם כשאינו בעמודה A (ולא «סטאטוס משלוח»)',F.c);
 ok('בלי «משלוח»: בכרטיס — «אספקה» ומספר האספקה',F.card);
 ok('בלי «משלוח»: בלוח ההגעות — הכותרת «אספקה» ומספר האספקה',F.board);

 /* 7 · תצוגת 4 הנושאים */
 const q=await ctx.newPage();q.on('pageerror',e=>errs.push(e.message));
 await q.goto('file://'+path.join(SD,'..','index.html')+'?topics');await q.setInputFiles('#f',XL);
 await q.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await q.waitForSelector('#lite table');
 await q.setInputFiles('#fe',ETAX);await q.waitForTimeout(2500);
 const Q=await q.evaluate(()=>{
  /* בועית בעמודה של נושא. בדוח ההדגמה לפריטי «הזמנות לקוח» אין שורות ETA, ותאריכי ה-ETA אחרי «סוף החודש»
     של הדוח — לכן לשתי השורות הראשונות מוצמדות זמנית שורות ETA (ו«סוף החודש» מוזז זמנית), ואחר כך מוחזרות.
     בכל שורה כזו: בבועית — כל המשלוחים של השורות שהתא מסכם, ורק הם; התא עצמו כמו קודם */
  const tipCol=(tab,key,prep,keep,txt)=>{TP.tab=tab;TP.cat={};TP.sort={};TP.q='';TP.f={};TP.view='';tpRender();
   const pick=(TP.list||[]).slice(0,2),undo=pick.map(prep);tpRender();
   const ths=[...document.querySelectorAll('#lite thead th')],i=ths.findIndex(t=>t.dataset.tpSort===key);
   const trs=[...document.querySelectorAll('#lite tbody tr')],L=TP.list||[];
   const rows=trs.map((tr,j)=>({r:L[j],c:i<0?null:tr.children[i]})).filter(x=>x.r&&x.c&&pick.includes(x.r));
   const res={i,n:rows.length,ok:rows.every(x=>{const s=(x.c.querySelector('span')||{}).title||'',ls=ETA.byPn[etaKey(x.r.pn)]||[];
    const inn=[...new Set(ls.filter(l=>keep(x.r,l[1])).map(l=>l[4]))],outn=ls.filter(l=>!keep(x.r,l[1])).map(l=>l[4]);
    return inn.length>0&&inn.every(v=>s.includes(v))&&outn.every(v=>!s.includes(v))&&x.c.textContent.trim()===txt(x.r)})};
   undo.forEach(f=>f());tpRender();return res};
  const put=(r,ls,f)=>{const k=etaKey(r.pn),o=ETA.byPn[k],s={etaHas:r.etaHas,etaFirst:r.etaFirst,etaMonth:r.etaMonth};
   ETA.byPn[k]=ls;Object.assign(r,f);return ()=>{if(o)ETA.byPn[k]=o;else delete ETA.byPn[k];Object.assign(r,s)}};
  const E1=tipCol('cust','eta',r=>put(r,[[2,'2030-01-10','D-9','','SH-91'],[1,'2030-01-10','D-8','','SH-92'],[4,'2030-02-01','D-7','','SH-93']],{etaHas:true,etaFirst:'2030-01-10'}),
   (r,d)=>d===r.etaFirst,r=>tpdDate(r.etaFirst));
  const E0=etaEom;etaEom=()=>'2030-12-31';
  const E2=tipCol('short','etaM',r=>put(r,[[2,'2030-01-10','D-9','','SH-81'],[3,'2031-03-01','D-8','','SH-82']],{etaMonth:2}),
   (r,d)=>d>=etaToday()&&d<='2030-12-31',r=>tpN(r.etaMonth));
  etaEom=E0;
  const col=TOPICS.cust.cols.find(c=>c.k==='eta'),r0={etaHas:true,etaFirst:'2030-01-10'};
  /* ה-CSV של שני הנושאים: אותן עמודות, ואין בו מספר משלוח */
  const xsOf=k=>TP_HEAD.concat(TOPICS[k].cols).flatMap(c=>c.x||[[c.l,c.v]]);
  const xs=xsOf('cust').concat(xsOf('short'));
  const csv=['cust','short'].flatMap(k=>tpSorted(k,tpRows(k,false)).flatMap(r=>xsOf(k).map(x=>String(x[1](r)??''))));
  /* לוח ההגעות של 4 הנושאים */
  TP.view='arr';tpRender();const d=document.querySelector('#lite details.arr');d.open=true;
  const ath=[...d.querySelectorAll('thead th')].map(t=>t.textContent),ai=ath.indexOf('משלוח');
  const acells=[...d.querySelectorAll('tbody tr')].map(tr=>ai<0?'':tr.children[ai].textContent.trim());
  TP.view='';tpRender();
  /* הכרטיס של 4 הנושאים */
  const r=QF.all.find(x=>(ETA.byPn[etaKey(x.pn)]||[]).length);tpdShow(r);const li=[...document.querySelectorAll('#tpd ul.eta li')].map(x=>x.textContent);tpdClose();
  return {E1:{i:E1.i,n:E1.n,ok:E1.ok},E2:{i:E2.i,n:E2.n,ok:E2.ok},csvShip:csv.some(v=>/SH-\d+/.test(v)),csvHead:xs.some(x=>x[0]==='משלוח'),
   arr:ai>=0&&acells.length>0&&acells.every(t=>/^SH-\d+$/.test(t)),
   card:(ETA.byPn[etaKey(r.pn)]||[]).every(l=>li.some(t=>t.includes('משלוח '+l[4]))),vSame:!!r0&&col.v(r0)===r0.etaFirst}});
 ok('4 הנושאים, «הזמנות לקוח» › «אספקה הבאה (ETA)»: התאריך כמו קודם, והמשלוח של אותו תאריך בבועית',Q.E1.i>=0&&Q.E1.n>0&&Q.E1.ok&&Q.vSame,`${Q.E1.n} שורות`);
 ok('4 הנושאים, «מלאי 0 / צפוי להיגמר» › «אספקה עד סוף החודש (ETA)»: הכמות כמו קודם, והמשלוחים שלה בבועית',Q.E2.i>=0&&Q.E2.n>0&&Q.E2.ok,`${Q.E2.n} שורות`);
 ok('4 הנושאים: ה-CSV של שני הנושאים בלי שינוי — אין עמודת «משלוח» ואין מספר משלוח',!Q.csvShip&&!Q.csvHead);
 ok('4 הנושאים: לוח ההגעות — «משלוח» ומספר המשלוח',Q.arr);
 ok('4 הנושאים: הכרטיס — «משלוח» ומספר המשלוח',Q.card);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
