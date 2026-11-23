/* ============ TP_COLFILT — מיון וסינון בכותרות בתצוגת 4 הנושאים (הערת משתמשים 23.11, אושר) ============
   כמו בכלי המלא: לחיצה על כותרת ממיינת; ▾ פותח חלונית — מיון עולה/יורד, חיפוש בעמודה, ערכים לסימון
   (טקסט) או מ-עד (מספר), «החל» ו«נקה». סינון לכל נושא בנפרד, עם שבב להסרה. כפתור «סינון» נשאר.
   בלוח ההגעות — מיון בלבד. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx'),ETAX=path.join(SD,'zmrp-demo-eta.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');
 const reset=()=>p.evaluate(()=>{TP.tab='cust';TP.cat={};TP.sort={};TP.q='';TP.f={};TP.cf={};TP.view='';tpRender()});
 const pop=()=>p.evaluate(()=>{const e=document.getElementById('colpop');return {open:e.classList.contains('open'),vis:getComputedStyle(e).display!=='none'}});
 await reset();

 /* 1 · ▾ בכל כותרת, והחלונית נפתחת */
 const H=await p.evaluate(()=>{const ths=[...document.querySelectorAll('#lite thead th')];return {n:ths.length,fil:ths.filter(t=>t.querySelector('[data-tp-fil]')).length,ftog:!!document.querySelector('#lite [data-tp="ftog"]')}});
 ok('בכל כותרת בטבלה יש ▾ לסינון, וכפתור «סינון» נשאר',H.n>3&&H.fil===H.n&&H.ftog,`${H.fil}/${H.n}`);
 await p.click('#lite th[data-tp-sort="free"] [data-tp-fil]');
 let P=await pop();
 ok('▾ פותח את חלונית העמודה (גם בתצוגה המצומצמת)',P.open&&P.vis,JSON.stringify(P));
 const M=await p.evaluate(()=>({asc:!!document.getElementById('tpcAsc'),desc:!!document.getElementById('tpcDesc'),s:!!document.getElementById('tpcS'),rng:!!document.getElementById('tpcMin')&&!!document.getElementById('tpcMax'),vals:!!document.getElementById('tpcVals')}));
 ok('בעמודת מספר: מיון עולה/יורד, חיפוש, ו«מ-עד» (בלי רשימת ערכים)',M.asc&&M.desc&&M.s&&M.rng&&!M.vals,JSON.stringify(M));

 /* 2 · מיון מתוך החלונית */
 await p.click('#tpcAsc');
 const S1=await p.evaluate(()=>{const c=tpCol('cust','free'),L=TP.list.map(r=>c.v(r)).filter(v=>v!=null);return {s:TP.sort.cust,mono:L.every((v,i)=>!i||L[i-1]<=v),open:document.getElementById('colpop').classList.contains('open')}});
 ok('«▲ עולה»: הטבלה ממוינת לפי העמודה, והחלונית נסגרת',S1.s&&S1.s.k==='free'&&S1.s.dir===1&&S1.mono&&!S1.open,JSON.stringify(S1.s));

 /* 3 · מ-עד */
 await p.click('#lite th[data-tp-sort="free"] [data-tp-fil]');
 await p.fill('#tpcMin','1');await p.fill('#tpcMax','5');await p.click('#tpcOk');
 const R=await p.evaluate(()=>{const c=tpCol('cust','free'),all=tpRows('cust',false);
  return {n:all.length,inR:all.every(r=>{const v=c.v(r);return typeof v==='number'&&v>=1&&v<=5}),
   rows:document.querySelectorAll('#lite tbody tr.r').length,filt:document.querySelector('#lite th[data-tp-sort="free"]').classList.contains('filt'),
   chip:[...document.querySelectorAll('#lite .fchip')].map(x=>x.textContent).join('|'),tab:+document.querySelector('[data-tp-tab="cust"] .n').textContent.replace(/[^\d]/g,'')}});
 ok('«מ-1 עד 5»: רק שורות שהערך שלהן בטווח, הכותרת מסומנת, ושבב «מלאי פנוי: מ-1 עד 5»',R.n>0&&R.inR&&R.rows===Math.min(R.n,200)&&R.filt&&/מלאי פנוי: מ-1 עד 5/.test(R.chip),`${R.n} שורות · ${R.chip}`);
 ok('המונה בלשונית הנושא = השורות אחרי הסינון',R.tab===R.n,`${R.tab}/${R.n}`);
 /* CSV — השורות שבטבלה */
 const C=await p.evaluate(()=>{const t='cust',cols=TP_HEAD.concat(TOPICS[t].cols);return tpSorted(t,tpRows(t,false)).length});
 ok('הייצוא ל-CSV — אותן שורות כמו בטבלה',C===R.n,`${C}/${R.n}`);

 /* 4 · הסינון לכל נושא בנפרד */
 const T=await p.evaluate(()=>{const before=tpRows('trend',false).length;TP.cf.trend=undefined;return {other:before===tpRows('trend',false).length,cf:Object.keys(TP.cf.cust||{})}});
 ok('הסינון חל רק על הנושא שבו הוגדר',T.other&&T.cf.join()==='free',T.cf.join());

 /* 5 · שבב ✕ מסיר */
 await p.click('#lite .fchip [data-tp="cfunset"]');
 const U=await p.evaluate(()=>({n:tpRows('cust',false).length,all:TOPICS.cust.rows().filter(r=>!tpDone(r)&&!isNeutral(r)).length,cf:Object.keys(TP.cf.cust||{}).length}));
 ok('✕ בשבב מסיר את הסינון, והרשימה חוזרת למלואה',U.cf===0&&U.n===U.all,`${U.n}/${U.all}`);

 /* 6 · עמודת טקסט: רשימת ערכים לסימון, וחיפוש בעמודה */
 await reset();
 await p.click('#lite th[data-tp-sort="kind"] [data-tp-fil]');
 const V=await p.evaluate(()=>{const ls=[...document.querySelectorAll('#tpcVals label')].map(l=>({v:l.querySelector('input').value,n:+l.querySelector('.c').textContent.replace(/[^\d]/g,'')}));
  return {ls,rng:!!document.getElementById('tpcMin')}});
 ok('בעמודת טקסט: רשימת הערכים עם מספר השורות לכל ערך (בלי «מ-עד»)',V.ls.length>=2&&!V.rng&&V.ls.every(x=>x.n>0),V.ls.map(x=>`${x.v}:${x.n}`).join(', '));
 const pickV=V.ls[0].v,pickN=V.ls[0].n;
 await p.click(`#tpcVals input[value="${pickV}"]`);await p.click('#tpcOk');
 const W=await p.evaluate(v=>{const c=tpCol('cust','kind'),L=tpRows('cust',false);return {n:L.length,all:L.every(r=>tpCfStr(c.v(r))===v)}},pickV);
 ok('סימון ערך אחד ו«החל»: רק השורות עם הערך, ובדיוק המספר שהוצג ליד הערך',W.all&&W.n===pickN,`${pickV}: ${W.n}/${pickN}`);
 await reset();
 await p.click('#lite th[data-tp-sort="desc"] [data-tp-fil]');
 /* מילה שמצמצמת: המילה הנדירה ביותר מתיאור השורה הראשונה */
 const word=await p.evaluate(()=>{const ws=String(TP.list[0].desc||'').split(/\s+/).filter(w=>w.length>1);const L=tpRows('cust',false);
  return ws.map(w=>[w,L.filter(r=>String(r.desc||'').toLowerCase().includes(w.toLowerCase())).length]).sort((a,b)=>a[1]-b[1])[0][0]});
 await p.fill('#tpcS',word);await p.press('#tpcS','Enter');
 const Q=await p.evaluate(w=>{const L=tpRows('cust',false);return {n:L.length,all:L.every(r=>String(r.desc||'').toLowerCase().includes(w.toLowerCase()))}},word);
 ok('חיפוש בעמודה (Enter): רק שורות שהעמודה מכילה את הטקסט — והרשימה מצטמצמת',Q.n>0&&Q.all&&Q.n<131,`«${word}»: ${Q.n}`);
 /* «נקה סינון עמודה» */
 await p.click('#lite th[data-tp-sort="desc"] [data-tp-fil]');await p.click('#tpcClear');
 ok('«נקה סינון עמודה» מסיר את הסינון',await p.evaluate(()=>!Object.keys(TP.cf.cust||{}).length));
 /* לחיצה על הכותרת עצמה — עדיין ממיינת */
 await p.click('#lite th[data-tp-sort="cust"]');
 ok('לחיצה על הכותרת (לא על ▾) ממיינת, כמו קודם',await p.evaluate(()=>TP.sort.cust&&TP.sort.cust.k==='cust'));

 /* 7 · לוח ההגעות — מיון בלבד */
 await p.setInputFiles('#fe',ETAX);await p.waitForTimeout(2500);
 await p.evaluate(()=>{TP.view='arr';TP.arrSort=null;tpRender();const ds=[...document.querySelectorAll('#lite details.arr')];const d=ds.sort((a,b)=>b.querySelectorAll('tbody tr').length-a.querySelectorAll('tbody tr').length)[0];d.open=true});
 const A0=await p.evaluate(()=>{const d=document.querySelector('#lite details.arr[open]');return {d:d.dataset.d,fil:!!d.querySelector('[data-tp-fil]'),hs:[...d.querySelectorAll('thead th')].map(t=>t.dataset.tpAsort)}});
 await p.click('#lite details.arr[open] th[data-tp-asort="pn"]');
 const A1=await p.evaluate(d=>{const x=document.querySelector(`#lite details.arr[data-d="${d}"]`);const pns=[...x.querySelectorAll('tbody tr')].map(tr=>tr.querySelector('[data-tp="arrpn"]').textContent);
  return {open:x.open,sorted:pns.every((v,i)=>!i||pns[i-1].localeCompare(v,'he')<=0),n:pns.length,s:TP.arrSort}},A0.d);
 ok('לוח ההגעות: כל כותרת ממיינת, בלי ▾ (מיון בלבד)',A0.hs.length===6&&A0.hs.every(Boolean)&&!A0.fil,A0.hs.join(','));
 ok('לחיצה על «מק״ט» ממיינת את טבלת התאריך, והתאריך נשאר פתוח',A1.open&&A1.sorted&&A1.s.k==='pn'&&A1.s.dir===1&&A1.n>1,`${A1.n} שורות`);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
