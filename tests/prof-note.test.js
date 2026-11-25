/* ============ PROF_NOTE + NEUTRAL_NOTE + PROF_OFF_LIFT (בקשות המעתד 25.11, אושרו) ============
   1. «ללא שינוי פרופיל» — שדה הערה לא חובה; בתצוגת 4 הנושאים מעבר לפריט הבא (כמו «טופל»), בכלי המלא
      נשאר על הפריט. ההערה ברשימת «ללא שינוי פרופיל», בגיליון שלה (profExclNote) ובכרטיס.
   2. נטרול («כללי» / «עונתיות» / «קמפיין») — שדה הערה לא חובה; ההערה ברשימת «מנוטרלים» ובכרטיס.
   3. נטרול PD ← ZM-C מתבטל כשיש מכירה אחרי יום הנטרול — «בוטל — מכירה חדשה» בכרטיס, בשני הכלים. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const open=async q=>{const p=await ctx.newPage();p.errs=[];p.on('pageerror',e=>p.errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+q);await p.setInputFiles('#f',XL);
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});return p};
 const T=await open('?topics');await T.evaluate(()=>{try{localStorage.clear()}catch(_){}});await T.waitForSelector('#lite table');

 /* 1 · «ללא שינוי פרופיל» בתצוגת 4 הנושאים */
 const P0=await T.evaluate(()=>{TP.tab='fix';TP.cat={fix:'prof'};TP.sort={};tpRender();const L=TP.list;tpdShow(L[0]);return {pn:L[0].pn,k:markKey(L[0]),next:L[1]&&L[1].pn}});
 await T.click('#tpd [data-profoff="1"]');await T.waitForSelector('#nAskN');
 const lbl=await T.evaluate(()=>document.querySelector('#colpop label').textContent);
 ok('«ללא שינוי פרופיל» פותח שדה הערה (לא חובה)',/למה לא לשנות פרופיל/.test(lbl),lbl);
 await T.press('#nAskN','Escape');
 const esc1=await T.evaluate(k=>({off:!!PROFOFF[k],open:TP.open&&markKey(TP.open)}),P0.k);
 ok('Esc — ביטול: לא נוטרל, הכרטיס על אותו פריט',!esc1.off&&esc1.open===P0.k);
 await T.click('#tpd [data-profoff="1"]');await T.fill('#nAskN','פריט קמפיין חד-פעמי');await T.press('#nAskN','Enter');await T.waitForTimeout(300);
 const P1=await T.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);return {note:(PROFOFF[k]||{}).note,off:profOff(r),open:TP.open&&TP.open.pn,inList:TP.list.includes(r),excl:profExclNote(r)}},P0.k);
 ok('Enter: נוטרל, וההערה נשמרה',P1.off&&P1.note==='פריט קמפיין חד-פעמי',P1.note);
 ok('הפריט יצא מהרשימה והכרטיס עבר לפריט הבא',!P1.inList&&P1.open===P0.next,`${P0.pn} → ${P1.open}`);
 ok('הגיליון «ללא שינוי פרופיל» (profExclNote): ההערה',P1.excl==='פריט קמפיין חד-פעמי',P1.excl);
 const LST=await T.evaluate(()=>{tpdClose();tpRender();const d=document.querySelector('#lite details.off');return d?d.textContent:''});
 ok('רשימת «ללא שינוי פרופיל» בטיוב נתונים: ההערה',LST.includes('פריט קמפיין חד-פעמי'),LST.slice(0,120));
 const C1=await T.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);TP.tab='all';TP.cat={};tpRender();tpdShow(r);const t=document.querySelector('#tpd').textContent;tpdClose();return t},P0.k);
 ok('הכרטיס של הפריט: «מנוטרל» עם ההערה',C1.includes('שינוי פרופיל — מנוטרל')&&C1.includes('פריט קמפיין חד-פעמי'));

 /* 3 · ביטול בגלל מכירה חדשה */
 const L1=await T.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k),e=PROFOFF[k];e.ts=Date.now()-10*864e5;r.sales={date:new Date(e.ts+3*864e5)};   /* נוטרל לפני 10 ימים, מכירה לפני 7 */PFSIM=null;DL_CACHE={};apply();
  TP.tab='fix';TP.cat={fix:'prof'};tpRender();const inList=TP.list.includes(r);tpdShow(r);const lift=(document.querySelector('#tpd .tplift')||{}).textContent||'';tpdClose();
  return {pd:isPD(r.mrp),inList,lift,off:profOff(r),stored:!!PROFOFF[k],offList:profOffRows().includes(r)}},P0.k);
 ok('מכירה אחרי יום הנטרול: הנטרול אינו פעיל, הפריט חזר ל«שינוי פרופיל MRP» ויצא מרשימת המנוטרלים',L1.pd&&!L1.off&&L1.inList&&!L1.offList&&L1.stored,JSON.stringify({pd:L1.pd,off:L1.off,in:L1.inList}));
 ok('בכרטיס: «ללא שינוי פרופיל» בוטל — מכירה חדשה, עם תאריך הנטרול וההערה',/בוטל — מכירה חדשה ב-/.test(L1.lift)&&L1.lift.includes('פריט קמפיין חד-פעמי'),L1.lift);
 /* נטרול מחדש — תאריך חדש, פעיל שוב */
 const R1=await T.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);setProfOff(r,true,'שוב');return {off:profOff(r),lifted:!!profOffLifted(r)}},P0.k);
 ok('נטרול מחדש אחרי הביטול — פעיל שוב (תאריך חדש)',R1.off&&!R1.lifted);

 /* 2 · נטרול עם הערה */
 const N0=await T.evaluate(()=>{TP.tab='cust';TP.cat={};tpRender();tpdShow(TP.list[0]);return {k:markKey(TP.list[0])}});
 await T.click('#tpd [data-tpd="neu"][data-t="neutral"]');await T.waitForSelector('#nAskN');
 const nl=await T.evaluate(()=>document.querySelector('#colpop label').textContent);
 ok('נטרול «כללי» פותח שדה «סיבת הנטרול» (לא חובה)',/סיבת הנטרול/.test(nl)&&/כללי/.test(nl),nl);
 await T.fill('#nAskN','פריט מוחלף ב-PN-02');await T.press('#nAskN','Enter');await T.waitForTimeout(300);
 const N1=await T.evaluate(k=>({mk:MARKS[k],moved:!TP.open||markKey(TP.open)!==k}),N0.k);
 ok('נשמר נטרול «כללי» עם ההערה, והכרטיס עבר הלאה',N1.mk&&N1.mk.t==='neutral'&&N1.mk.note==='פריט מוחלף ב-PN-02'&&N1.moved,JSON.stringify(N1.mk));
 const NL=await T.evaluate(()=>{tpdClose();TP.neuOpen=true;tpRender();const s=document.querySelector('#lite .neul');const v={h:[...s.querySelectorAll('th')].map(x=>x.textContent),t:s.textContent};TP.neuOpen=false;tpRender();return v});
 ok('רשימת «מנוטרלים»: עמודת «הערה» עם ההערה',NL.h.includes('הערה')&&NL.t.includes('פריט מוחלף ב-PN-02'),NL.h.join('|'));
 const NC=await T.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);TP.tab='all';TP.cat={};tpRender();tpdShow(r);const f=document.querySelector('#tpd .db').firstElementChild;const v={cls:f.className,t:f.textContent};tpdClose();return v},N0.k);
 ok('כרטיס של פריט מנוטרל: תיבה עם ההערה וסוג הנטרול',/tpprev/.test(NC.cls)&&/neu/.test(NC.cls)&&NC.t.includes('פריט מוחלף ב-PN-02')&&NC.t.includes('נטרול — כללי'),NC.t.slice(0,120));
 /* שלושת הסוגים */
 const N3=await T.evaluate(()=>{TP.tab='cust';TP.cat={};tpRender();tpdShow(TP.list[0]);const out=[];
  for(const t of ['season','campaign']){document.querySelector(`#tpd [data-tpd="neu"][data-t="${t}"]`).click();out.push(document.querySelector('#colpop label').textContent);closeColPop()}
  tpdClose();return out});
 ok('גם «עונתיות» ו«קמפיין» — שדה הערה',N3.length===2&&N3.every(x=>/סיבת הנטרול/.test(x)),N3.join(' | '));

 /* הכלי המלא — «ללא שינוי פרופיל» עם הערה, נשאר על הפריט; ושורת הביטול במכירה חדשה */
 const F=await open('?nobrief=1');
 const k2=await F.evaluate(()=>{const r=QF.all.find(x=>pdUpgrade(x)&&!profOff(x));detail(r);return markKey(r)});
 await F.click('#detail [data-profoff="1"]');await F.waitForSelector('#nAskN');await F.fill('#nAskN','לא עכשיו');await F.press('#nAskN','Enter');await F.waitForTimeout(400);
 const F1=await F.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);return {note:(PROFOFF[k]||{}).note,still:!!document.querySelector('#detail [data-profoff="0"]')}},k2);
 ok('כלי מלא: הערה נשמרת, והכרטיס נשאר על הפריט עם «החזר המלצה»',F1.note==='לא עכשיו'&&F1.still,JSON.stringify(F1));
 const F2=await F.evaluate(k=>{const r=ALL.find(x=>markKey(x)===k);r.sales={date:new Date(PROFOFF[k].ts+2*864e5)};PFSIM=null;DL_CACHE={};apply();detail(ALL.find(x=>markKey(x)===k));
  return (document.querySelector('#detail').textContent.match(/«ללא שינוי פרופיל» בוטל[^)]*\)/)||[''])[0]},k2);
 ok('כלי מלא: «ללא שינוי פרופיל» בוטל — מכירה חדשה, בכרטיס',/בוטל — מכירה חדשה/.test(F2)&&F2.includes('לא עכשיו'),F2);

 ok('אין שגיאות JS',!T.errs.length&&!F.errs.length,[...T.errs,...F.errs].slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
