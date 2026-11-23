/* ============ SHORT_ACT_FULL — «מלאי 0» בכלי המלא: פעולת אספקה אחת במילים (החלטת המעתד 20.11) ============
   כמו בתצוגת 4 הנושאים (tpShortAct): במסכי החוסרים — «חוסרים לטיפול» ו«רכש ואספקות» —
   העמודה «הפעולה הבאה» אומרת פעולה אחת במילים, בלי כמות (מספר בפעולה מאומץ בלי בדיקה).
   הכרטיס נשאר עם הכמות, ושאר המסכים כמו קודם (בחירת המעתד). באקסל — המשפט המלא של המנוע,
   כמו קודם. «כמות להזמנה» ירדה מברירת המחדל של «לפי ספק» (נשארת בבורר ובגיליון «כל הנתונים»). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx'),ETAX=path.join(SD,'zmrp-demo-eta.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 const wait=ms=>p.waitForTimeout(ms);
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.evaluate(()=>{try{localStorage.clear()}catch(_){}});
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB,null,{timeout:180000});
 await p.setInputFiles('#fe',ETAX);await wait(2200);
 /* תא «הפעולה הבאה» בכל שורה מצוירת, מול הפריט שלה — לפי המק״ט שבשורה (במסכי הכסף data-i
    מצביע על הרשימה המקוצרת לפי קו ה-80%, ולא על LAST_VIEW) */
 const scan=()=>p.evaluate(()=>{const i=[...document.querySelectorAll('#tbl thead th')].findIndex(t=>t.dataset.ck==='act');
  const by=new Map(QF.all.map(r=>[String(r.pn),r])),trs=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
  const pnOf=tr=>{const o=tr.querySelector('.obj');return o?o.textContent.trim():''};
  const rows=trs.map(tr=>({t:i<0?null:tr.children[i].textContent.trim(),ttl:i<0?'':tr.children[i].title,r:by.get(pnOf(tr))}));
  const kinds={};rows.forEach(x=>{const k=String(x.t).replace(/: .*$/,'');kinds[k]=(kinds[k]||0)+1});
  return {view:colViewKey(),short:SHORTLIKE(),i,n:rows.length,
   worded:rows.every(x=>x.r&&x.t===tpShortAct(x.r)&&x.ttl===x.t),
   engine:rows.every(x=>x.r&&x.t===(actShort(x.r)||'—')),
   qty:rows.filter(x=>/\d+\s*יח׳/.test(x.t)).length,kinds,
   th:(document.querySelector('#tbl thead th[data-ck="act"] .thc')||{}).title||''}});

 /* 1 · «חוסרים לטיפול» */
 await p.evaluate(()=>setMode('today'));await wait(500);
 let S=await scan();
 ok('«חוסרים לטיפול»: «הפעולה הבאה» = פעולת האספקה במילים (tpShortAct), בכל שורה',S.view==='dec'&&S.short&&S.n>50&&S.worded,`${S.n} שורות`);
 ok('אף תא בעמודה אינו נוקב בכמות',S.qty===0,String(S.qty));
 ok('בדמו מופיעות לפחות שלוש מהפעולות: לפתוח הזמנת רכש · לבקש ETA · למסור ללקוח תאריך',
  S.kinds['לפתוח הזמנת רכש']>0&&S.kinds['לבקש ETA מהספק']>0&&S.kinds['למסור ללקוח תאריך']>0,JSON.stringify(S.kinds));
 ok('הכותרת מסבירה: בלי כמות, הכמות בכרטיס, באקסל המשפט המלא',/בלי כמות/.test(S.th)&&/בכרטיס/.test(S.th)&&/באקסל/.test(S.th),S.th.slice(0,70));
 const X=await p.evaluate(()=>{const v=viewExportRows(),i=v.keys.indexOf('act');const L=LAST_VIEW||currentRows();
  return {i,head:v.aoa[0][i],same:v.aoa.slice(1).every((row,j)=>row[i]===String((L[j].act||[])[0]||'')),qty:v.aoa.slice(1).filter(row=>/\d+\s*יח׳/.test(row[i])).length,n:v.aoa.length-1}});
 ok('הייצוא לא השתנה: בגיליון «התצוגה» — המשפט המלא של המנוע, כמו קודם',X.i>=0&&X.head==='הפעולה הבאה'&&X.same&&X.n===S.n,`${X.n} שורות · ${X.qty} עם כמות`);
 /* המיון והסינון קוראים ל-val: במסכי החוסרים — מה שעל המסך (המיון עצמו פועל בתוך כל קבוצה ובלוק מטבע) */
 await p.evaluate(()=>{document.querySelector('#tbl thead th[data-ck="act"] .thc').click()});await wait(350);
 const SO=await p.evaluate(()=>({ws:WSORT.col,val:currentRows().every(r=>SCOL.act.val(r)===tpShortAct(r)),x:currentRows().every(r=>SCOL.act.xval(r)===String((r.act||[])[0]||''))}));
 ok('מיון וסינון לפי «הפעולה הבאה» הולכים לפי מה שעל המסך; הייצוא לפי המשפט המלא',SO.ws==='act'&&SO.val&&SO.x,JSON.stringify(SO));
 await p.evaluate(()=>{WSORT={col:null,dir:'desc'};render()});await wait(300);

 /* 2 · «רכש ואספקות» */
 /* DESIGN-C (21.11): «מעקב רכש» נפתח על כל הרכש הפתוח (עמודות משלו, בלי «הפעולה הבאה»);
    פעולת האספקה במילים חלה על טבלת החוסרים שבתוכו — הבחירה 'stuck' */
 await p.evaluate(()=>{setMode('line');lineSelect('stuck')});await wait(500);
 S=await scan();
 ok('«רכש ואספקות»: אותה פעולה במילים, בלי כמות',S.short&&S.n>0&&S.worded&&S.qty===0,`${S.n} שורות`);

 /* 3 · שאר המסכים — כמו קודם */
 for(const m of ['home','cap']){await p.evaluate(m=>setMode(m),m);await wait(500);
  S=await scan();ok(`«${m==='home'?'מרכז עבודה':'בריאות המלאי'}»: המשפט של המנוע, כמו קודם`,!S.short&&S.i>=0&&S.engine,`${S.n} שורות`)}

 /* 4 · הכרטיס נשאר עם הכמות */
 await p.evaluate(()=>setMode('today'));await wait(400);
 const C=await p.evaluate(()=>{const r=currentRows().find(x=>/\d+\s*יח׳/.test((x.act||[])[0]||''));detail(r);
  const d=document.querySelector('#detail').textContent;return {has:d.includes((r.act||[])[0]),act:(r.act||[])[0]}});
 ok('כרטיס הפריט לא השתנה: ההחלטה עם הכמות',C.has,C.act);
 await p.evaluate(()=>closeDetail());await wait(200);

 /* 5 · «לפי ספק»: «כמות להזמנה» ירדה מברירת המחדל, ונשארת לבחירה */
 await p.evaluate(()=>{groupBy='supplier';render()});await wait(500);
 const U=await p.evaluate(()=>({view:colViewKey(),heads:[...document.querySelectorAll('#tbl thead th[data-ck]')].map(t=>t.dataset.ck),
  def:SCOL_DEF.sup.slice(),offered:colChoices('sup').includes('buy'),sheet:viewExportRows().aoa[0]}));
 ok('«לפי ספק»: בלי «כמות להזמנה» בברירת המחדל',U.view==='sup'&&!U.def.includes('buy')&&!U.heads.includes('buy'),U.heads.join(','));
 ok('ובגיליון «התצוגה» של «לפי ספק» — גם בלעדיה (אושר)',!U.sheet.includes('כמות להזמנה'),U.sheet.join('|'));
 ok('«כמות להזמנה» עדיין מוצעת בחלון «עמודות»',U.offered);
 await p.click('#colsBtn');await wait(300);
 await p.click('#colspick input[data-ck="buy"]');await wait(450);
 ok('סימון בחלון מחזיר אותה לטבלה',await p.evaluate(()=>!!document.querySelector('#tbl thead th[data-ck="buy"]')));
 await p.evaluate(()=>{colsOpen(false);colReset('sup');groupBy='decision';render()});await wait(400);
 const M=await p.evaluate(()=>{window.__cap=[];const o=XLSX.writeFile;XLSX.writeFile=(wb,name)=>{const sh={};for(const sn of wb.SheetNames){const a=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1});sh[sn]=(a.find(r=>r&&r.some(c=>/^מק[״"']ט/.test(String(c))))||[])}window.__cap.push(sh)};
  /* XLS_REFORM (23.11): ב«נתונים עיקריים» שורת קבוצות מעל הכותרת — הכותרת היא השורה עם «מק״ט» */
  try{doExport()}finally{XLSX.writeFile=o}const s=window.__cap[0]||{};return Object.entries(s).filter(([n,h])=>h.includes('כמות להזמנה')).map(([n])=>n)});
 ok('הכמות להזמנה נשארת באקסל, בגיליון «נתונים עיקריים»',M.includes('נתונים עיקריים'),M.join(' · '));

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
