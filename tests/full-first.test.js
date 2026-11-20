/* ============ FIRST_FULL — «צריכה ראשונה» בכלי המלא (החלטת המעתד 20.11) ============
   אותו כלל כמו בתצוגת 4 הנושאים (tpFirst): פריט שלא נצרך מעולם — לא ב-11 החודשים ולא
   ב-3 השנים שבדוח — ויש לו עכשיו צריכה ראשונה (חודש מלא אחרון או החודש הנוכחי).
   דלת ב«תכנון ו-MRP» ליד «ביקוש יורד» (בחירת המעתד), רשימה, ושורה בכרטיס. בלי שינוי
   במנוע, בתורים ובקובץ SAP; שתי עמודות הכמות במסך בלבד (noX, החלטה 1.11). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
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
 await wait(600);

 /* 1 · הכלל: בדיוק אותה אוכלוסייה כמו בתת-הקטגוריה של 4 הנושאים */
 const R=await p.evaluate(()=>{const L=firstRows(),cat=TOPICS.trend.cats.find(c=>c[0]==='first')[3];
  const want=QF.all.filter(r=>cat(r));
  const sane=L.every(r=>r.has3y&&(r.months||[]).slice(1).every(v=>!(v>0))&&(firstM0(r)+firstCm(r))>0&&(r.cons3y||0)-firstM0(r)-firstCm(r)<=0);
  return {n:L.length,same:L.length===want.length&&L.every(r=>want.includes(r)),sane}});
 ok('אותה אוכלוסייה כמו «צריכה ראשונה» בתצוגת 4 הנושאים',R.n>0&&R.same,String(R.n));
 ok('כל פריט ברשימה: בלי צריכה קודמת, ועם צריכה עכשיו',R.sane);

 /* 2 · הדלת במסילה */
 await p.evaluate(()=>setMode('month'));await wait(500);
 const D=await p.evaluate(()=>{const subs=[...document.querySelectorAll('#railsub .railsub2')].map(b=>b.dataset.m);
  const b=document.querySelector('#railsub .railsub2[data-m="first"]');
  return {subs,i:subs.indexOf('first'),t:b&&b.querySelector('.t').textContent,n:b&&+b.querySelector('.bdg').dataset.count,ttl:b&&b.title}});
 ok('«צריכה ראשונה» במסילה של «תכנון ו-MRP», מיד אחרי «ביקוש יורד»',D.i>0&&D.subs[D.i-1]==='trend'&&D.t==='צריכה ראשונה',D.subs.join(','));
 ok('המונה בדלת = אורך הרשימה',D.n===R.n,`${D.n}/${R.n}`);
 ok('ההסבר בדלת: לא נצרך מעולם, לבדוק אם להגדיר תכנון',/לא נצרך מעולם/.test(D.ttl)&&/תכנון/.test(D.ttl),D.ttl.slice(0,70));

 /* 3 · המסך */
 await p.click('#railsub .railsub2[data-m="first"]');await wait(600);
 const S=await p.evaluate(()=>{const rows=currentRows();
  const heads=[...document.querySelectorAll('#tbl thead th[data-ck]')].map(t=>t.dataset.ck);
  const trs=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
  const tot=rows.map(r=>firstM0(r)+firstCm(r));
  const cellsOk=trs.every((tr,i)=>{const r=LAST_VIEW[+tr.dataset.i]||rows[i];if(!r)return false;
   const a=tr.querySelector('td[data-ck="fM0"]').textContent.trim(),c=tr.querySelector('td[data-ck="fCm"]').textContent.trim();
   const f=n=>n?n.toLocaleString('he-IL'):'—';return a===f(firstM0(r))&&c===f(firstCm(r))});
  return {mode,track,view:colViewKey(),title:document.getElementById('lt').textContent,sub:(document.getElementById('phdSub')||{}).textContent||'',
   heads,n:rows.length,trs:trs.length,sorted:tot.every((v,i)=>i===0||tot[i-1]>=v),cellsOk,pgR:document.getElementById('pgR').textContent,
   m0:rows.reduce((a,r)=>a+firstM0(r),0),cm:rows.reduce((a,r)=>a+firstCm(r),0),po:rows.filter(r=>r.po>0).length,
   btns:['exportXls','sapXls','sapSlowXls','expXls','lmailBtn','arrBtn'].filter(id=>!document.getElementById(id).hidden)}});
 ok('המסך: «תכנון ו-MRP · צריכה ראשונה»',S.mode==='first'&&S.track==='first'&&S.view==='first'&&S.title==='תכנון ו-MRP · צריכה ראשונה',S.title);
 ok('תת-הכותרת אומרת את הכלל ואת הפעולה',/לא נצרך מעולם/.test(S.sub)&&/3 השנים/.test(S.sub)&&/להגדיר לו תכנון/.test(S.sub),S.sub.slice(0,60));
 ok('העמודות: מק״ט, דגם, תיאור, שני החודשים, מלאי, בדרך, MRP, סיווג',JSON.stringify(S.heads)===JSON.stringify(['pnPlain','model','desc','fM0','fCm','free','poPlain','mrp','cat']),S.heads.join(','));
 ok('כל הפריטים ברשימה, מהצריכה הגדולה לקטנה',S.n===R.n&&S.trs===S.n&&S.sorted,`${S.trs}/${S.n}`);
 ok('תאי שני החודשים — הכמות, או מקף',S.cellsOk);
 ok('שורת הסיכום: יחידות בשני החודשים ורכש פתוח, בלי סייג מחיר',S.pgR.includes(`${S.m0.toLocaleString('he-IL')} יח׳ בחודש המלא האחרון`)&&S.pgR.includes(`${S.cm.toLocaleString('he-IL')} יח׳ בחודש הנוכחי`)&&S.pgR.includes(`${S.po.toLocaleString('he-IL')} עם רכש פתוח`)&&!/בלי מחיר/.test(S.pgR),S.pgR);
 ok('בכותרת רק «ייצוא לאקסל» — אין כאן קובץ SAP',JSON.stringify(S.btns)===JSON.stringify(['exportXls']),S.btns.join(','));
 const X=await p.evaluate(()=>{const v=viewExportRows();return {keys:v.keys,head:v.aoa[0],n:v.n}});
 ok('גיליון «התצוגה»: בלי שתי עמודות הכמות החדשות (במסך בלבד), שאר העמודות כמו במסך',!X.keys.includes('fM0')&&!X.keys.includes('fCm')&&JSON.stringify(X.keys)===JSON.stringify(['pnPlain','model','desc','free','poPlain','mrp','cat'])&&X.n===S.n,X.head.join('|'));
 await p.evaluate(()=>{document.querySelector('#tbl thead th[data-ck="fCm"] .thc').click()});await wait(300);
 ok('מיון לפי «החודש הנוכחי»',await p.evaluate(()=>WSORT.col==='fCm'));
 await p.evaluate(()=>{WSORT={col:null,dir:'desc'};render()});await wait(300);

 /* 4 · חיפוש מצמצם */
 const pn=await p.evaluate(()=>currentRows()[0].pn);
 await p.evaluate(v=>setQuery(v),pn);await wait(500);
 ok('חיפוש מק״ט מצמצם את הרשימה',await p.evaluate(v=>{const L=currentRows();return L.length>=1&&L.every(r=>String(r.pn).includes(v))},pn));
 await p.evaluate(()=>clearQuery());await wait(400);

 /* 5 · הכרטיס */
 const C=await p.evaluate(()=>{const r=currentRows().find(x=>firstM0(x)>0)||currentRows()[0];detail(r);
  const e=document.querySelector('#detail .anom.f');const t=e?e.textContent:'';
  const o=QF.all.find(x=>!tpFirst(x)&&(x.d12||0)>0);detail(o);const none=!document.querySelector('#detail .anom.f');
  return {t,none,m0:firstM0(r),cm:firstCm(r)}});
 ok('כרטיס: «צריכה ראשונה» עם הכמויות, «לא נרשמה צריכה קודמת» והפעולה',/צריכה ראשונה/.test(C.t)&&/לא נרשמה צריכה קודמת/.test(C.t)&&/לבדוק אם להגדיר תכנון לפריט/.test(C.t)
  &&(C.m0?C.t.includes(`${C.m0.toLocaleString('he-IL')} יח׳ בחודש המלא האחרון`):true)&&(C.cm?C.t.includes(`${C.cm.toLocaleString('he-IL')} יח׳ בחודש הנוכחי`):true),C.t.slice(0,90));
 ok('כרטיס של פריט שנצרך קודם — בלי השורה',C.none);
 await p.evaluate(()=>closeDetail());await wait(200);

 /* 6 · בלי עמודות 3 השנים הכלל אינו חל, והמסך אומר למה */
 const E=await p.evaluate(()=>{const keep=QF.all.map(r=>r.has3y);QF.all.forEach(r=>r.has3y=false);DL_CACHE={};render();
  const t=(document.querySelector('#tbl tbody td.empty')||{}).textContent||'';const n=currentRows().length;
  QF.all.forEach((r,i)=>r.has3y=keep[i]);DL_CACHE={};render();return {t,n,back:currentRows().length}});
 ok('בלי עמודות 3 השנים: רשימה ריקה, וההסבר אומר למה',E.n===0&&/אין עמודות 3 השנים/.test(E.t)&&E.back===R.n,E.t.slice(0,60));

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
