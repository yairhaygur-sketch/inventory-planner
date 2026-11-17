/* ============ TOPICS_UX — ביקורת מסך 4 הנושאים (אושר 15.11): קריטי וגבוה, ואחריהם בינוני ונמוך ============
   תצוגה בלבד: אין חישוב, סף או סיווג חדשים — כל מספר נבדק כאן מול השדה של המנוע.
     #1 פעולה לפי הנושא (ב-CSV נשארת גם «פעולה» של המנוע)
     #2 מיון ברירת מחדל לפי דחיפות + שווי ההזמנה בהזמנות לקוח
     #3 «להזמין» במלאי 0 = הכמות שבפעולה (actQty)
     #4 סרגל אחד: שבבי תת-קטגוריה, המסננים מקופלים
     #5 שורת הסבר אחת, ההסבר המלא מאחורי ⓘ
     #6 שורה אחת לפריט
     #7 פס «אין דוח ETA», ועמודות ETA מוסתרות בלעדיו
     #8 C העתקה · D «טופל» ומעבר לבא */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx'),ETAX=path.join(SD,'zmrp-demo-eta.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML=process.argv[2]||path.join(SD,'..','index.html');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+HTML+'?topics');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');
 const tab=t=>p.evaluate(t=>{TP.tab=t;TP.cat={};TP.sort={};TP.q='';TP.f={};tpRender()},t);

 /* #1 — הפעולה בכל נושא שייכת לנושא */
 const A=await p.evaluate(()=>{const R={};
  for(const t of ['cust','trend','short','fix']){TP.tab=t;TP.cat={};tpRender();
   const list=TP.list,cells=[...document.querySelectorAll('#lite tbody tr')].map(tr=>tr.lastElementChild.textContent.trim());
   R[t]={n:list.length,cellsMatch:list.slice(0,cells.length).every((r,i)=>cells[i]===tpTopicAct(t,r)),
    trendBad:t==='trend'?list.filter(r=>!/^לבחון (העלאת|הורדת) נק׳ הזמנה ל-/.test(tpTopicAct(t,r))).map(r=>r.pn):[],
    trendNum:t==='trend'?list.filter(r=>{const a=tpTopicAct(t,r);return riseGap(r)>0?!a.includes(tpN(r.sugROPU)):!a.includes(tpN(r.sugROPT))}).map(r=>r.pn):[],
    fixBad:t==='fix'?list.filter(r=>/הזמנת רכש|רכש יזום|להשלים/.test(tpTopicAct(t,r))).map(r=>r.pn):[],
    fixCat:t==='fix'?{prof:list.filter(r=>profActive(r)&&!(tpLtMissing(r)||tpSrvMissing(r))&&!(tpInFix(r)&&!recNA(r))).every(r=>/פרופיל:/.test(tpTopicAct(t,r))),
     data:list.filter(r=>tpLtMissing(r)||tpSrvMissing(r)).every(r=>/^להזין /.test(tpTopicAct(t,r)))}:null,
    same:(t==='cust'||t==='short')?list.every(r=>tpTopicAct(t,r)===tpAct(r)):true,
    engineDiffers:list.filter(r=>tpTopicAct(t,r)!==tpAct(r)).length}}
  return R});
 ok('#1 כל תא «פעולה» בטבלה = הפעולה של הנושא',Object.values(A).every(x=>x.cellsMatch),JSON.stringify(Object.fromEntries(Object.entries(A).map(([k,v])=>[k,v.cellsMatch]))));
 ok('#1 חריגת צריכה: «לבחון העלאת/הורדת נק׳ הזמנה ל-N», N = נקודת ההזמנה המייעצת של המנוע',A.trend.n>0&&!A.trend.trendBad.length&&!A.trend.trendNum.length,[...A.trend.trendBad,...A.trend.trendNum].slice(0,3).join(','));
 ok('#1 טיוב נתונים: אף פעולת רכש; פרופיל ← «… פרופיל:», נתונים שבורים ← «להזין …»',A.fix.n>0&&!A.fix.fixBad.length&&A.fix.fixCat.prof&&A.fix.fixCat.data,JSON.stringify(A.fix.fixCat));
 ok('#1 הזמנות לקוח ומלאי 0: הפעולה של המנוע כמו שהיא',A.cust.same&&A.short.same);
 ok('#1 ההבדל מהמנוע אמיתי (יש שורות בטיוב ובחריגה שבהן הפעולה הוחלפה)',A.fix.engineDiffers>0&&A.trend.engineDiffers>0,`טיוב ${A.fix.engineDiffers} · חריגה ${A.trend.engineDiffers}`);

 /* #1 + CSV: «פעולה בנושא» ו«פעולה» של המנוע */
 await tab('fix');
 {const [dl]=await Promise.all([p.waitForEvent('download'),p.click('#lite [data-tp="csv"]')]);
  const H=fs.readFileSync(await dl.path(),'utf8').replace(/^﻿/,'').split('\r\n')[0].split(',').map(x=>x.replace(/"/g,''));
  ok('#1 CSV: «פעולה בנושא» וגם «פעולה» של המנוע',H.includes('פעולה בנושא')&&H.includes('פעולה'),H.slice(-4).join(','))}

 /* #2 — מיון ברירת מחדל לפי דחיפות, ושווי ההזמנה */
 const S=await p.evaluate(()=>{const R={};
  TP.tab='cust';TP.cat={};TP.sort={};tpRender();
  {const L=TP.list,bIdx=L.findIndex(r=>!r.burn);
   R.burnFirst=bIdx<0||L.slice(bIdx).every(r=>!r.burn);
   const blk=L.filter(r=>r.burn&&!r.priceMissing);let mono=true;
   for(let i=1;i<blk.length;i++){const a=blk[i-1],c=blk[i];if(curPos(a)===curPos(c)&&custVal(a)<custVal(c))mono=false;if(curPos(a)>curPos(c))mono=false}
   R.valueDesc=mono;R.custN=L.length;
   R.valCol=[...document.querySelectorAll('#lite thead th')].some(th=>th.textContent.trim().startsWith('שווי ההזמנה'));
   const tr=document.querySelector('#lite tbody tr'),r0=L[0],ths=[...document.querySelectorAll('#lite thead th')].map(t=>t.textContent.trim()),i=ths.findIndex(t=>t.startsWith('שווי ההזמנה'));
   R.valCell=tr.children[i].textContent.trim()===(r0.priceMissing?'—':money(custVal(r0),r0.currency));
   R.label=(document.querySelector('#lite .sortl')||{}).textContent||''}
  TP.tab='short';TP.cat={};TP.sort={};tpRender();
  {const L=TP.list,c=r=>shelfQty(r)<=0?0:(r.unsQty||0)>0?1:2;let ok=true;for(let i=1;i<L.length;i++)if(c(L[i-1])>c(L[i]))ok=false;R.shortOrder=ok;R.shortN=L.length}
  TP.tab='trend';TP.cat={};TP.sort={};tpRender();
  {const L=TP.list,g=r=>Math.max(riseGap(r),trendGap(r));let ok=true;for(let i=1;i<L.length;i++)if(g(L[i-1])<g(L[i]))ok=false;R.trendOrder=ok}
  TP.tab='cust';TP.sort={cust:{k:'free',dir:1}};tpRender();R.manualHidesLabel=!document.querySelector('#lite .sortl');TP.sort={};
  return R});
 ok('#2 הזמנות לקוח: «אין רכש כלל» קודם',S.custN>0&&S.burnFirst);
 ok('#2 ובתוכן — שווי ההזמנה יורד, בבלוקים לפי מטבע (בלי להשוות שקל לדולר)',S.valueDesc);
 ok('#2 עמודת «שווי ההזמנה» = custVal במטבע הפריט',S.valCol&&S.valCell);
 ok('#2 מלאי 0: «מלאי 0 עכשיו» קודם, אחריו «צפוי להיגמר»',S.shortN>0&&S.shortOrder);
 ok('#2 חריגת צריכה: לפי ההשפעה על נקודת ההזמנה, מהגדולה',S.trendOrder);
 ok('#2 הסרגל אומר לפי מה ממוין, ונעלם כשממיינים ידנית',/^ממוין: /.test(S.label)&&S.manualHidesLabel,S.label);

 /* #3 — «להזמין» = הכמות שבפעולה */
 await tab('short');
 const Q=await p.evaluate(()=>{const ths=[...document.querySelectorAll('#lite thead th')].map(t=>t.textContent.trim()),i=ths.findIndex(t=>t.startsWith('להזמין'));
  const trs=[...document.querySelectorAll('#lite tbody tr')];
  return {i,noUns:!ths.some(t=>t.startsWith('לא יסופק')),bad:trs.map((tr,j)=>{const r=TP.shown[j],q=actQty(r),c=tr.children[i].textContent.trim();return c===(q>0?tpN(q):'—')?null:r.pn}).filter(Boolean)}});
 ok('#3 מלאי 0: עמודת «להזמין» = הכמות שבפעולה (actQty), ו«לא יסופק החודש» ירד מהטבלה',Q.i>=0&&Q.noUns&&!Q.bad.length,Q.bad.slice(0,3).join(','));
 {const [dl]=await Promise.all([p.waitForEvent('download'),p.click('#lite [data-tp="csv"]')]);
  const H=fs.readFileSync(await dl.path(),'utf8').replace(/^﻿/,'').split('\r\n')[0];
  ok('#3 ה-CSV שומר גם «לא יסופק החודש»',/להזמין/.test(H)&&/לא יסופק החודש/.test(H))}

 /* #4 #5 #6 — סרגל אחד, שורת הסבר, שורה אחת */
 await tab('cust');
 const L=await p.evaluate(()=>{const d=document.querySelector('#lite details.lead'),bar=document.querySelector('#lite .bar');
  const rows=[...document.querySelectorAll('#lite tbody tr')].slice(0,30).map(tr=>Math.round(tr.getBoundingClientRect().height));
  return {lead:!!d&&!d.open&&d.querySelector('summary').textContent.includes(TOPICS.cust.short)&&d.querySelector('p').textContent===TOPICS.cust.lead,
   tilesInBar:!!bar.querySelector('.tiles .tile'),fHidden:document.querySelector('#lite .bar.flt').hidden,
   maxRow:Math.max(...rows),minRow:Math.min(...rows),btns:[...document.querySelectorAll('#lite tbody tr')].every(tr=>tr.querySelector('[data-tp="copy"]')&&tr.querySelector('[data-tp="done"]'))}});
 ok('#5 שורת הסבר אחת; ההסבר המלא בתוך ⓘ, סגור כברירת מחדל',L.lead);
 ok('#4 שבבי תת-הקטגוריה בתוך הסרגל, והמסננים מקופלים',L.tilesInBar&&L.fHidden);
 ok('#6 כל שורה בגובה שורה אחת (≤ 36px), עם כפתורי העתקה ו«טופל»',L.maxRow<=36&&L.btns,`${L.minRow}–${L.maxRow}px`);

 /* #7 — בלי דוח ETA: פס, ועמודות ETA מוסתרות. עם דוח: אין פס, והעמודות חוזרות */
 const E0=await p.evaluate(()=>({eta:!!ETA,banner:!!document.querySelector('#lite .etab'),cols:[...document.querySelectorAll('#lite thead th')].map(t=>t.textContent.trim()).filter(t=>/ETA/.test(t))}));
 await p.setInputFiles('#fe',ETAX);await p.waitForFunction(()=>!!ETA,null,{timeout:60000});await tab('short');
 const E1=await p.evaluate(()=>({banner:!!document.querySelector('#lite .etab'),cols:[...document.querySelectorAll('#lite thead th')].map(t=>t.textContent.trim()).filter(t=>/ETA/.test(t))}));
 ok('#7 בלי דוח ETA: פס אזהרה, ואין עמודות ETA בטבלה',!E0.eta&&E0.banner&&!E0.cols.length,JSON.stringify(E0));
 ok('#7 אחרי העלאת דוח ETA: הפס נעלם והעמודות חוזרות',!E1.banner&&E1.cols.length===2,JSON.stringify(E1));

 /* #8 — C ו-D */
 await tab('cust');
 await ctx.grantPermissions(['clipboard-read','clipboard-write']).catch(()=>{});
 const k0=await p.evaluate(()=>({pn0:TP.list[0].pn,pn1:TP.list[1].pn,n:TP.list.length}));
 await p.focus('#lite tr.r[data-tp-row="0"]');await p.keyboard.press('c');await p.waitForTimeout(200);
 const clip=await p.evaluate(async()=>await navigator.clipboard.readText().catch(()=>null));
 ok('#8 C על שורה ממוקדת מעתיק את המק״ט',clip===null||clip===k0.pn0,String(clip));
 await p.keyboard.press('d');await p.waitForTimeout(400);
 const k1=await p.evaluate(pn=>({mark:(MARKS[markKey(ALL.find(r=>r.pn===pn))]||{}).t,n:TP.list.length,focus:document.activeElement&&TP.shown[+document.activeElement.dataset.tpRow]&&TP.shown[+document.activeElement.dataset.tpRow].pn}),k0.pn0);
 ok('#8 D על שורה ממוקדת: «טופל» (handled), והמיקוד עובר לפריט הבא',k1.mark==='handled'&&k1.n===k0.n-1&&k1.focus===k0.pn1,JSON.stringify(k1));
 await p.click('#lite tr.r[data-tp-row="0"]');await p.keyboard.press('d');await p.waitForTimeout(400);
 const k2=await p.evaluate(pn=>({mark:(MARKS[markKey(ALL.find(r=>r.pn===pn))]||{}).t,open:TP.open&&TP.open.pn,first:TP.list[0].pn}),k0.pn1);
 ok('#8 D בכרטיס פתוח: «טופל», והכרטיס עובר לפריט הבא',k2.mark==='handled'&&k2.open===k2.first&&k2.open!==k0.pn1,JSON.stringify(k2));
 await p.keyboard.press('Escape');
 await p.click('#lite details.lead summary');await p.click('#lite details.lead summary');
 const k3=await p.evaluate(()=>Object.keys(MARKS).length);await p.keyboard.press('d');await p.waitForTimeout(200);
 ok('#8 בלי פריט ממוקד — D אינו עושה דבר',await p.evaluate(()=>Object.keys(MARKS).length)===k3);

 /* ============ בינוני ונמוך (#9–#14) ============ */
 await p.evaluate(()=>{for(const k of Object.keys(MARKS))delete MARKS[k];apply()});
 await tab('cust');
 const M=await p.evaluate(()=>{const R={};
  /* #9 — «גם ב…» בדיוק בפריטים שמופיעים בנושא נוסף, ובאילו */
  const sets=Object.fromEntries(Object.keys(TOPICS).map(k=>[k,new Set(tpRows(k,true))]));
  const bad9=[];let n9=0;
  for(const t of Object.keys(TOPICS)){TP.tab=t;TP.cat={};tpRender();
   [...document.querySelectorAll('#lite tbody tr')].forEach((tr,i)=>{const r=TP.shown[i],exp=Object.keys(sets).filter(k=>k!==t&&sets[k].has(r));
    const tag=tr.querySelector('.also'),txt=tag?tag.textContent:'';if(exp.length)n9++;
    if(!!tag!==!!exp.length||exp.some(k=>!txt.includes(TP_SHORT[k])))bad9.push(t+':'+r.pn)})}
  R.n9=n9;R.bad9=bad9;
  /* #10 — מתכנן אחד: אין עמודת «מתכנן»; «סוג MRP» לא בהזמנות לקוח; «זמן אספקה» לא במלאי 0 */
  const heads=t=>{TP.tab=t;TP.cat={};tpRender();return [...document.querySelectorAll('#lite thead th')].map(x=>x.textContent.trim())};
  const pls=new Set(QF.all.map(r=>r.planner).filter(Boolean)).size;
  R.pls=pls;R.noPlanner=pls>1||!heads('cust').some(h=>h.startsWith('מתכנן'));
  R.noMrp=!heads('cust').some(h=>h.startsWith('סוג MRP'));R.noLt=!heads('short').some(h=>h.startsWith('זמן אספקה'));
  /* #11 — «מצב» רק כשהרשימה מעורבת; בלשונית — החלק הדחוף באדום */
  TP.tab='cust';TP.cat={cust:'none'};tpRender();R.kindHidden=![...document.querySelectorAll('#lite thead th')].some(x=>x.textContent.trim().startsWith('מצב'));
  TP.cat={};tpRender();R.kindShown=[...document.querySelectorAll('#lite thead th')].some(x=>x.textContent.trim().startsWith('מצב'));
  const u=k=>{const e=document.querySelector(`[data-tp-tab="${k}"] .u`);return e?+((e.textContent.match(/[\d,]+/)||['0'])[0].replace(/,/g,'')):0};
  R.urg={cust:[u('cust'),tpRows('cust',true).filter(r=>r.burn).length],short:[u('short'),tpRows('short',true).filter(TOPICS.short.cats[0][3]).length],
   fix:[u('fix'),tpRows('fix',true).filter(TOPICS.fix.cats[2][3]).length]};
  /* #12 — חריגת צריכה: מספר אחד בתא, וה«בדוח» בנקודה */
  TP.tab='trend';TP.cat={};tpRender();
  const ths=[...document.querySelectorAll('#lite thead th')],ix=ths.findIndex(t=>t.dataset.tpSort==='tRec');
  const bad12=[];let n12=0;[...document.querySelectorAll('#lite tbody tr')].forEach((tr,i)=>{const r=TP.shown[i],raw=tpRawSum(r,0),td=tr.children[ix];
   const vis=[...td.childNodes].filter(n=>!(n.classList&&n.classList.contains('vh'))).map(n=>n.textContent).join('');
   if(raw!==r.tRec){n12++;if(!td.querySelector('.tdot')||/בדוח/.test(vis)||!td.querySelector('.vh'))bad12.push(r.pn)}});
  R.n12=n12;R.bad12=bad12;
  /* #13 — שמות אחידים */
  const H={};for(const t of ['cust','trend','short','fix'])H[t]=heads(t);
  R.names={kind:['cust','trend','short'].every(t=>H[t].some(h=>h.startsWith('מצב'))),noSug:!H.trend.some(h=>h==='נק׳ הזמנה מוצעת'||h.startsWith('לפי המגמה (מייעץ)')),
   fixPairs:H.fix.some(h=>h.startsWith('נק׳ הזמנה (SAP ← מוצעת)'))&&H.fix.some(h=>h.startsWith('מלאי ביטחון (SAP ← מוצע)'))};
  /* #14 — חריגה: ההשפעה על נק׳ ההזמנה לפני האחוז; טיוב: שווי 0 ← «—» */
  /* 17.11: הכותרת אומרת מי ההמלצה ומי המייעץ — «נק׳ הזמנה: המלצה ← אם המגמה תימשך» */
  R.impBeforePct=H.trend.findIndex(h=>h.startsWith('נק׳ הזמנה: המלצה ← אם המגמה תימשך'))>=0&&H.trend.findIndex(h=>h.startsWith('נק׳ הזמנה: המלצה ← אם המגמה תימשך'))<H.trend.findIndex(h=>h.startsWith('שינוי'));
  TP.tab='fix';TP.cat={};tpRender();const vi=[...document.querySelectorAll('#lite thead th')].findIndex(t=>t.dataset.tpSort==='val');
  R.zeroDash=[...document.querySelectorAll('#lite tbody tr')].every((tr,i)=>{const v=tpRecVal(TP.shown[i]),c=tr.children[vi].textContent.trim();return v===0?c==='—':!/^[+−]?[€$₪]0$/.test(c)});
  return R});
 ok('#9 «גם ב…» בדיוק בפריטים שמופיעים בנושא נוסף, ושם הנושא נכון',M.n9>0&&!M.bad9.length,`${M.n9} שורות · ${M.bad9.slice(0,3).join(',')}`);
 ok('#10 עמודות שלא מכריעות: «מתכנן» (ערך יחיד), «סוג MRP» בלקוחות, «זמן אספקה» במלאי 0 — לא בטבלה',M.noPlanner&&M.noMrp&&M.noLt,JSON.stringify({pls:M.pls,mrp:M.noMrp,lt:M.noLt}));
 ok('#11 עמודת «מצב» רק כשהרשימה מעורבת',M.kindHidden&&M.kindShown);
 ok('#11 בלשוניות — החלק הדחוף באדום, ובמספר הנכון',Object.values(M.urg).every(([a,b])=>a===b)&&M.urg.cust[0]>0,JSON.stringify(M.urg));
 ok('#12 חריגת צריכה: מספר אחד בתא; הסכום בדוח בנקודה ובטקסט לקורא מסך',M.n12>0&&!M.bad12.length,`${M.n12} · ${M.bad12.slice(0,3).join(',')}`);
 ok('#13 שמות אחידים: «מצב» בכל הנושאים, «(SAP ← מוצעת)» בטיוב',M.names.kind&&M.names.noSug&&M.names.fixPairs,JSON.stringify(M.names));
 ok('#14 חריגה: ההשפעה על נק׳ ההזמנה לפני האחוז; טיוב: שינוי 0 מוצג «—»',M.impBeforePct&&M.zeroDash,JSON.stringify({imp:M.impBeforePct,zero:M.zeroDash}));
 /* #9 — הודעת «טופל» אומרת מאילו נושאים נוספים הפריט ירד */
 {const k=await p.evaluate(()=>{for(const t of Object.keys(TOPICS)){TP.tab=t;TP.cat={};tpRender();const r=TP.shown.find(x=>tpOthers(x).length);if(r)return {t,i:TP.shown.indexOf(r),oth:tpOthers(r).map(x=>TOPICS[x].t)}}return null});
  if(k){await p.focus(`#lite tr.r[data-tp-row="${k.i}"]`);await p.keyboard.press('d');await p.waitForTimeout(300);
   const msg=await p.evaluate(()=>(document.getElementById('toast')||{}).textContent||'');
   ok('#9 «טופל» על פריט בכמה נושאים — ההודעה אומרת שירד גם מהם',k.oth.every(o=>msg.includes(o)),msg.slice(0,120))}
  else ok('#9 יש פריט בכמה נושאים לבדיקת ההודעה',false)}
 /* #14 — כרטיס: PD ← ZM-C שבו מלאי הביטחון יורד אומר למה */
 {const c=await p.evaluate(()=>{const r=QF.all.find(x=>{if(!pdUpgrade(x))return false;const q=tpRec(x);return !q.na&&q.src==='sim'&&q.ss<x.ss});
   if(!r)return null;TP.tab='fix';tpRender();tpdShow(r);const t=document.getElementById('tpd').textContent;tpdClose();return {pn:r.pn,ok:/נקבע מחדש מהצריכה/.test(t)&&/מלאי ביטחון מינימלי/.test(t)}});
  ok('#14 כרטיס: PD ← ZM-C עם מלאי ביטחון נמוך מ-SAP — שורה שמסבירה',!!c&&c.ok,JSON.stringify(c))}

 ok('0 שגיאות JS',!errs.length,errs.join('|'));
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})();
