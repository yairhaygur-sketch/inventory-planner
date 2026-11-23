/* ============ TP_DONE_NEXT + TP_DONE_NOTE — «טופל», הערה, ומעבר לפריט הבא (בקשות המעתד 23.11) ============
   «✓ טופל» (בכרטיס או בשורה) פותח שדה הערה לא חובה; Enter מסמן ועובר לבא. כשהסימון פג והפריט חוזר —
   ההערה והתאריך בכרטיס ובבועית בשורה. «טופלו» — מהאחרון לראשון, עם שעה, הערה וימים עד החזרה.
   כפתור «✓ טופל» בכרטיס: מסמן «טופל» ופותח את הפריט הבא ברשימה. «✓ טופל» בשורה כשהכרטיס של אותה
   שורה פתוח: הכרטיס נשאר פתוח ועובר לפריט הבא (עד כאן נסגר). מקש D — כמו קודם, דרך אותה פונקציה. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');
 await p.evaluate(()=>{try{localStorage.clear()}catch(_){}});
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');
 const st=()=>p.evaluate(()=>({open:TP.open?TP.open.pn:null,list:(TP.list||[]).map(r=>r.pn)}));

 /* 1 · הכפתור בכרטיס */
 await p.evaluate(()=>{TP.tab='cust';TP.cat={};tpRender()});
 await p.click('#lite tr.r[data-tp-row="0"]');
 const A0=await st();
 ok('בכרטיס יש «✓ טופל» ליד «הקודם / הבא»',await p.evaluate(()=>!!document.querySelector('#tpd .dnav [data-tpd="done"]')));
 await p.click('#tpd [data-tpd="done"]');
 const P1=await p.evaluate(()=>({open:document.getElementById('colpop').classList.contains('open'),inp:!!document.getElementById('tpdnN')}));
 ok('«✓ טופל» פותח שדה הערה (לא חובה)',P1.open&&P1.inp,JSON.stringify(P1));
 await p.fill('#tpdnN','אין אפשרות לזרז, אין מקום ל-VOR');await p.press('#tpdnN','Enter');await p.waitForTimeout(300);
 const A1=await st();
 const note1=await p.evaluate(pn=>{const r=QF.all.find(x=>x.pn===pn);return (MARKS[markKey(r)]||{}).note},A0.open);
 ok('Enter: ההערה נשמרת על הסימון',note1==='אין אפשרות לזרז, אין מקום ל-VOR',note1);
 const m1=await p.evaluate(pn=>{const r=QF.all.find(x=>x.pn===pn);return MARKS[markKey(r)]&&MARKS[markKey(r)].t},A0.open);
 ok('לחיצה: הפריט מסומן «טופל» ויוצא מהרשימה',m1==='handled'&&!A1.list.includes(A0.open),m1);
 ok('והכרטיס נשאר פתוח — על הפריט שהיה הבא ברשימה',A1.open&&A1.open===A0.list[1],`${A0.open} → ${A1.open}`);

 /* 2 · «✓ טופל» בשורה, כשהכרטיס של אותה שורה פתוח */
 const B0=await st();const i=B0.list.indexOf(B0.open);
 await p.evaluate(pn=>{const k=markKey(QF.all.find(x=>x.pn===pn));const btn=document.querySelector(`#lite [data-tp="done"][data-key="${k}"]`);btn.click()},B0.open);await p.waitForTimeout(150);await p.press('#tpdnN','Enter');await p.waitForTimeout(300);
 const B1=await st();
 ok('«✓ טופל» בשורה של הפריט הפתוח: הכרטיס לא נסגר — עובר לפריט הבא',B1.open&&B1.open===B0.list[i+1]&&!B1.list.includes(B0.open),`${B0.open} → ${B1.open}`);

 /* 3 · «✓ טופל» בשורה אחרת — הכרטיס הפתוח לא זז */
 const C0=await st();const other=C0.list.find(pn=>pn!==C0.open);
 await p.evaluate(pn=>{const k=markKey(QF.all.find(x=>x.pn===pn));document.querySelector(`#lite [data-tp="done"][data-key="${k}"]`).click()},other);await p.waitForTimeout(150);await p.click('#tpdnOk');await p.waitForTimeout(300);
 const C1=await st();
 ok('«✓ טופל» בשורה אחרת: הכרטיס הפתוח נשאר על אותו פריט',C1.open===C0.open&&!C1.list.includes(other));

 /* 4 · מקש D — עדיין «טופל» ומעבר לבא */
 const D0=await st();const j=D0.list.indexOf(D0.open);
 await p.keyboard.press('d');await p.waitForTimeout(300);
 const D1=await st();
 ok('מקש D: «טופל» ומעבר לפריט הבא, כמו קודם',D1.open===D0.list[j+1]&&!D1.list.includes(D0.open),`${D0.open} → ${D1.open}`);

 /* 5 · «כל הפריטים»: הפריט נשאר ברשימה, ו«הבא» הוא השורה שאחריו */
 await p.evaluate(()=>{tpdClose();TP.tab='all';TP.cat={};TP.sort={};tpRender()});
 await p.click('#lite tr.r[data-tp-row="0"]');
 const E0=await st();await p.click('#tpd [data-tpd="done"]');await p.waitForTimeout(150);await p.press('#tpdnN','Enter');await p.waitForTimeout(300);
 const E1=await st();
 ok('«כל הפריטים»: הפריט נשאר ברשימה (מסומן), והכרטיס עובר לשורה שאחריו',E1.list.includes(E0.open)&&E1.open===E0.list[E0.list.indexOf(E0.open)+1],`${E0.open} → ${E1.open}`);

 /* 6 · Esc בשדה — ביטול, בלי סימון */
 await p.evaluate(()=>{tpdClose();TP.tab='cust';TP.cat={};tpRender()});await p.click('#lite tr.r[data-tp-row="0"]');
 const F0=await st();await p.click('#tpd [data-tpd="done"]');await p.press('#tpdnN','Escape');await p.waitForTimeout(200);
 const F1=await p.evaluate(pn=>{const r=QF.all.find(x=>x.pn===pn);return {mark:!!MARKS[markKey(r)],open:TP.open&&TP.open.pn}},F0.open);
 ok('Esc בשדה: ביטול — הפריט לא סומן, והכרטיס נשאר עליו',!F1.mark&&F1.open===F0.open);
 /* 7 · פריט שחזר אחרי שהסימון פג: ההערה בכרטיס ובשורה */
 const G=await p.evaluate(()=>{const r=TP.list[0];r.markStale={t:'handled',note:'הזמנה שלא ניתן לזרז',ts:Date.now()-8*864e5};tpRender();tpdShow(r);
  const c=(document.querySelector('#tpd .tpprev')||{}).textContent||'';const i=document.querySelector('#lite tbody tr.r .tpnote');const ti=i?i.title:'';tpdClose();delete r.markStale;tpRender();return {c,ti}});
 ok('פריט שחזר: בכרטיס «סומן טופל ב-… : ההערה — וחזר לרשימה»',/סומן «טופל» ב-/.test(G.c)&&G.c.includes('הזמנה שלא ניתן לזרז')&&/חזר לרשימה/.test(G.c),G.c.slice(0,80));
 ok('ובשורה — 📝 עם ההערה בבועית',G.ti.includes('הזמנה שלא ניתן לזרז'),G.ti);
 /* 8 · «טופלו»: מהאחרון לראשון, עם שעה, הערה וימים עד החזרה */
 const H=await p.evaluate(()=>{tpRender();const lis=[...document.querySelectorAll('#lite details.off li')];
  const L=QF.all.filter(tpDone).sort((a,b)=>b.mark.ts-a.mark.ts).map(r=>r.pn);
  return {order:lis.map(li=>li.querySelector('b').textContent),want:L,first:lis[0]?lis[0].textContent:'',any:lis.some(li=>li.textContent.includes('אין אפשרות לזרז'))}});
 ok('«טופלו» ממוינת מהאחרון שסומן לראשון',H.order.length>2&&JSON.stringify(H.order)===JSON.stringify(H.want),H.order.slice(0,3).join(','));
 ok('בכל שורה: תאריך ושעה, ההערה, וכמה ימים עד שהפריט חוזר',/\d{1,2}:\d{2}/.test(H.first)&&/חוזר בעוד \d+ ימים/.test(H.first)&&H.any,H.first.slice(0,90));

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
