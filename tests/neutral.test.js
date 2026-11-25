/* ============ NEUTRAL — נטרול «כללי» / «עונתיות» / «קמפיין» (החלטות המעתד 18.11) ============
   פריט מנוטרל יוצא מכל התורים, מארבעת הנושאים, מ«בוער עכשיו», מקובץ העדכון ל-SAP
   ומהמלצת שינוי הפרופיל — בשני הכלים — עד שהמעתד מבטל (בלי תפוגה). «מנוטרלים (N)»
   בכותרת מציג אותם ומבטל. «טופל» ושאר הסימונים — כמו קודם. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');
 await p.evaluate(()=>{for(const k of Object.keys(MARKS))delete MARKS[k]});

 /* 1 · קבועים: שלושה סוגים, בלי תפוגה; «טופל» — 7 ימים כמו קודם */
 const K=await p.evaluate(()=>({t:NEUTRAL_T,ttl:NEUTRAL_T.map(markTTL),handled:markTTL('handled')/864e5,od:markTTL('ondemand')/864e5,
  exp:NEUTRAL_T.map(t=>markExpired({t,ts:Date.now()-5*365*864e5})),lbl:NEUTRAL_T.map(t=>MARK_LABEL[t])}));
 ok('שלושה סוגי נטרול: נטרול — כללי · עונתיות · קמפיין',K.lbl.join('|')==='נטרול — כללי|עונתיות|קמפיין',K.lbl.join(' / '));
 ok('בלי תפוגה — גם אחרי 5 שנים הסימון חי',K.ttl.every(v=>v===Infinity)&&K.exp.every(v=>v===false));
 ok('«טופל» 7 ימים ו«לפי דרישה» חצי שנה — לא השתנו',K.handled===7&&Math.round(K.od)===183,`${K.handled} · ${Math.round(K.od)}`);

 /* 2 · פריט בקובץ העדכון ל-SAP, בהזמנות לקוח ובשינוי פרופיל — נוטרל, ויוצא מכולם */
 for(const t of ['neutral','season','campaign']){
  const R=await p.evaluate(t=>{const inAll=k=>Object.keys(TOPICS).filter(k=>k!=='all').some(x=>tpRows(x,true).some(r=>markKey(r)===k));
   const fix=decisionList('fix').find(r=>!r.mark&&inAll(markKey(r)));const cust=tpBurnRows().find(r=>!r.mark);
   const prof=QF.all.find(r=>!r.mark&&profActive(r));
   const keys=[fix,cust,prof].map(r=>r&&markKey(r));
   const before={fix:!!fix,cust:!!cust,prof:!!prof};
   for(const r of [fix,cust,prof])if(r)setMark(r,t,'');
   const g=k=>ALL.find(x=>markKey(x)===k);
   const after={fixOut:!decisionList('fix').some(r=>markKey(r)===keys[0]),topicsOut:keys.every(k=>!k||!inAll(k)),
    burnOut:!tpBurnRows().some(r=>markKey(r)===keys[1]),profOut:!keys[2]||!profInfo(g(keys[2])),
    marked:keys.every(k=>!k||(QF.marked||[]).some(r=>markKey(r)===k)),listed:keys.every(k=>!k||tpNeuRows().some(r=>markKey(r)===k)),
    isN:keys.every(k=>!k||isNeutral(g(k)))};
   for(const k of keys)if(k)setMark(g(k),'','');
   const back={fixBack:!keys[0]||decisionList('fix').some(r=>markKey(r)===keys[0]),burnBack:!keys[1]||tpBurnRows().some(r=>markKey(r)===keys[1]),
    profBack:!keys[2]||!!profInfo(g(keys[2]))};
   return {before,after,back}},t);
  const L={neutral:'נטרול — כללי',season:'עונתיות',campaign:'קמפיין'}[t];
  ok(`«${L}»: יוצא מקובץ העדכון ל-SAP, מארבעת הנושאים, מ«בוער עכשיו» ומשינוי פרופיל`,R.before.fix&&R.before.cust&&R.before.prof&&R.after.fixOut&&R.after.topicsOut&&R.after.burnOut&&R.after.profOut,JSON.stringify(R.after));
  ok(`«${L}»: בתור «מסומנים» וברשימת «מנוטרלים»`,R.after.marked&&R.after.listed&&R.after.isN);
  ok(`«${L}»: ביטול — חוזר לקובץ, ל«בוער עכשיו» ולפרופיל`,R.back.fixBack&&R.back.burnBack&&R.back.profBack,JSON.stringify(R.back))}

 /* 3 · «טופל» אינו נטרול */
 const H=await p.evaluate(()=>{const r=tpBurnRows()[0];setMark(r,'handled','');const n=isNeutral(ALL.find(x=>markKey(x)===markKey(r)))||tpNeuRows().length>0;setMark(r,'','');return n});
 ok('«טופל» אינו נטרול ואינו ברשימת «מנוטרלים»',!H);

 /* 4 · ממשק ארבעת הנושאים: כפתור בכרטיס, «מנוטרלים (N)», ביטול */
 await p.evaluate(()=>{TP.tab='cust';TP.cat={};tpRender();tpdShow(TP.shown[0])});
 const U1=await p.evaluate(()=>{const r=TP.open,k=markKey(r),btns=[...document.querySelectorAll('#tpd [data-tpd="neu"]')].map(x=>x.dataset.t);
  document.querySelector('#tpd [data-tpd="neu"][data-t="season"]').click();
  document.getElementById('nAskOk').click();   /* NEUTRAL_NOTE (25.11): שדה הערה לא חובה — אישור */
  const head=document.querySelector('#lite [data-tp="neulist"]').textContent;
  return {k,btns,gone:!TP.list.some(x=>markKey(x)===k),moved:!TP.open||markKey(TP.open)!==k,head,mark:(MARKS[k]||{}).t}});
 ok('כרטיס: שלושה כפתורי נטרול',U1.btns.join()==='neutral,season,campaign',U1.btns.join());
 ok('לחיצה: הפריט יוצא מהרשימה, הכרטיס עובר הלאה, והסימון נשמר',U1.gone&&U1.moved&&U1.mark==='season',JSON.stringify(U1));
 ok('הכותרת: «מנוטרלים (1)»',/מנוטרלים \(1\)/.test(U1.head),U1.head);
 const U2=await p.evaluate(k=>{tpdClose();document.querySelector('#lite [data-tp="neulist"]').click();
  const rows=[...document.querySelectorAll('#lite .neul tbody tr')].map(tr=>tr.textContent);
  document.querySelector(`#lite [data-tp="neuoff"][data-key="${CSS.escape(k)}"]`).click();
  return {rows,mark:MARKS[k]||null,back:TOPICS.cust.rows().filter(r=>!isNeutral(r)).some(r=>markKey(r)===k),
   head:document.querySelector('#lite [data-tp="neulist"]').textContent}},U1.k);
 ok('«מנוטרלים»: שורה עם הסיבה (עונתיות)',U2.rows.length===1&&/עונתיות/.test(U2.rows[0]),U2.rows.join(' | ').slice(0,90));
 ok('«בטל נטרול»: הסימון נמחק, הפריט חוזר, והמונה 0',U2.mark===null&&U2.back&&/מנוטרלים \(0\)/.test(U2.head),U2.head);

 /* 5 · הכלי המלא: כפתורי הסימון והסבר «עד ביטול» */
 const F=await p.evaluate(()=>{const r=QF.all[0];const a=marksPanel(r);setMark(r,'neutral','');const r2=ALL.find(x=>markKey(x)===markKey(r));
  const b=marksPanel(r2);setMark(r2,'','');return {btn:/data-t="neutral"/.test(a)&&/data-t="season"/.test(a)&&/data-t="campaign"/.test(a),
   legend:/עד שתבטל את הסימון/.test(b)&&!/Infinity/.test(b),bulk:BULKMARKS.map(x=>x[0]).filter(k=>NEUTRAL_T.includes(k)).join()}});
 ok('הכלי המלא: כפתורי «נטרול — כללי», «עונתיות», «קמפיין» בכרטיס ובסימון הקבוצתי',F.btn&&F.bulk==='neutral,season,campaign',F.bulk);
 ok('הכלי המלא: ההסבר אומר «עד שתבטל», בלי מספר ימים',F.legend);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0);
})().catch(e=>{console.error(e);process.exit(1)});
