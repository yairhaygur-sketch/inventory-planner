/* ============ ARRIVALS_FULL — לוח הגעות ו-ETA לפי תאריך בכרטיס, בכלי המלא (החלטת המעתד 20.11) ============
   כמו בתצוגת 4 הנושאים. כפתור «לוח הגעות» בכותרת של «רכש ואספקות» מחליף את הטבלה בלוח ובחזרה.
   בלוח: כל הפריטים שיש להם שורות בדוח ה-ETA, אחרי החיפוש והסינון; תאריכים שעברו בראש; תאריך
   נפתח לפריטים שלו; מק״ט פותח את הכרטיס. בכרטיס: «אספקות לפי תאריך» — שורה לכל תאריך.
   קריאה בלבד: בלי שינוי במונים ובייצוא. */
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
 await wait(600);
 const st=()=>p.evaluate(()=>({on:ARR_ON,btn:!document.getElementById('arrBtn').hidden,
  pressed:document.getElementById('arrBtn').getAttribute('aria-pressed'),
  box:!document.getElementById('arrb').hidden,rows:!document.querySelector('#w_queue .rows').hidden,
  pager:!document.querySelector('#w_queue .pager').hidden}));

 /* 1 · הכפתור רק ב«רכש ואספקות» */
 const vis={};for(const m of ['home','today','line','month','catalog']){await p.evaluate(m=>setMode(m),m);await wait(300);vis[m]=(await st()).btn}
 ok('הכפתור «לוח הגעות» מוצג רק ב«רכש ואספקות»',vis.line&&!vis.home&&!vis.today&&!vis.month&&!vis.catalog,JSON.stringify(vis));
 await p.evaluate(()=>setMode('line'));await wait(400);
 let S=await st();
 ok('ברירת המחדל: הטבלה, בלי הלוח',!S.on&&!S.box&&S.rows&&S.pager&&S.pressed==='false',JSON.stringify(S));

 /* 2 · בלי דוח ETA — הלוח אומר את זה ומציע להעלות */
 await p.click('#arrBtn');await wait(400);
 S=await st();
 const E=await p.evaluate(()=>{const bx=document.getElementById('arrb');return {t:bx.textContent,up:!!bx.querySelector('label[for="fe"]')}});
 ok('לחיצה מחליפה את הטבלה בלוח (והכפתור לחוץ)',S.on&&S.box&&!S.rows&&!S.pager&&S.pressed==='true',JSON.stringify(S));
 ok('בלי דוח ETA: «אין דוח ETA» וכפתור העלאה',/אין דוח ETA/.test(E.t)&&E.up);

 /* 3 · עם דוח ETA: הלוח נבנה מחדש מעצמו, וכל השורות בו */
 await p.setInputFiles('#fe',ETAX);await wait(2500);
 const D=await p.evaluate(()=>{const base=QF.all.filter(passColFilters);let q=0,n=0;const dates=new Set(),items=new Set();
  for(const r of base){const ls=ETA.byPn[etaKey(r.pn)];if(!ls)continue;for(const l of ls)if(l[1]){q+=l[0];n++;dates.add(l[1]);items.add(r)}}
  const sm=[...document.querySelectorAll('#arrb details.arr')].map(d=>({past:d.classList.contains('past'),t:d.querySelector('summary').textContent}));
  const iso=x=>x.d,ds=ARR_D.map(iso);
  return {q,n,dates:dates.size,items:items.size,boardQ:ARR_D.reduce((a,x)=>a+x.q,0),boardN:ARR_D.reduce((a,x)=>a+x.ls.length,0),
   nd:sm.length,sorted:ds.every((d,i)=>i===0||ds[i-1]<d),pastFirst:sm.findIndex(x=>!x.past)===sm.filter(x=>x.past).length,
   pastTxt:sm.filter(x=>x.past).every(x=>/לבדוק אם נקלט/.test(x.t)),npast:sm.filter(x=>x.past).length,
   perDay:ARR_D.every(x=>x.n===new Set(x.ls.map(e=>e.r)).size&&x.q===x.ls.reduce((a,e)=>a+e.q,0)),
   head:(document.querySelector('#arrb .arrh')||{}).textContent||'',lazy:!document.querySelector('#arrb details.arr table')}});
 ok('הלוח נבנה מחדש אחרי טעינת ה-ETA, והוא עדיין פתוח',(await st()).on&&D.nd>0,String(D.nd));
 ok('כל שורות ה-ETA של פריטי הדוח בלוח — יחידות ושורות',D.boardQ===D.q&&D.boardN===D.n&&D.n>0,`${D.boardQ}/${D.q} יח׳ · ${D.boardN}/${D.n} שורות`);
 ok('שורה לכל תאריך, בסדר עולה; תאריכים שעברו בראש, עם «לבדוק אם נקלט»',D.nd===D.dates&&D.sorted&&D.pastFirst&&D.pastTxt&&D.npast>0,`${D.nd} תאריכים · ${D.npast} שעברו`);
 ok('בכל תאריך: מספר הפריטים והיחידות נכונים',D.perDay);
 ok('הכותרת אומרת מה בלוח',/לוח הגעות/.test(D.head)&&/אחרי החיפוש והסינון/.test(D.head),D.head.slice(0,70));
 ok('הטבלאות נבנות רק כשתאריך נפתח',D.lazy);

 /* 4 · פתיחת תאריך, ומק״ט שפותח את הכרטיס */
 await p.click('#arrb details.arr:not(.past) summary');await wait(300);
 const O=await p.evaluate(()=>{const d=document.querySelector('#arrb details.arr[open]'),x=ARR_D[+d.dataset.ad];
  const trs=[...d.querySelectorAll('tbody tr')],q=trs.map(tr=>+tr.children[2].textContent.replace(/[^\d]/g,''));
  return {n:trs.length,want:x.ls.length,desc:q.every((v,i)=>i===0||q[i-1]>=v),heads:[...d.querySelectorAll('thead th')].map(t=>t.textContent).join('|'),
   pn:trs[0].querySelector('.arrpn').textContent,first:x.sorted[0].r.pn}});
 ok('תאריך שנפתח: שורה לכל שורת ETA, מהכמות הגדולה לקטנה',O.n===O.want&&O.desc,`${O.n}/${O.want} · ${O.heads}`);
 await p.click('#arrb details.arr[open] .arrpn');await wait(500);
 ok('לחיצה על מק״ט פותחת את הכרטיס שלו',await p.evaluate(()=>CURRENT_DETAIL&&CURRENT_DETAIL.pn)===O.first,O.first);
 await p.evaluate(()=>closeDetail());await wait(250);

 /* 5 · החיפוש מצמצם את הלוח */
 const pn=await p.evaluate(()=>ARR_D[ARR_D.length-1].ls[0].r.pn);
 await p.evaluate(v=>setQuery(v),pn);await wait(500);
 const Q=await p.evaluate(v=>({only:ARR_D.every(x=>x.ls.every(e=>e.r.pn===v)),n:ARR_D.length}),pn);
 ok('חיפוש מק״ט: בלוח רק התאריכים שלו',Q.only&&Q.n>=1,`${Q.n} תאריכים`);
 await p.evaluate(()=>clearQuery());await wait(400);

 /* 6 · הייצוא לא תלוי בלוח */
 const xOn=await p.evaluate(()=>{const v=viewExportRows();return v?JSON.stringify(v.aoa):null});
 await p.click('#arrBtn');await wait(400);
 S=await st();
 ok('לחיצה נוספת מחזירה את הטבלה ואת שורת המונה',!S.on&&!S.box&&S.rows&&S.pager&&S.pressed==='false'&&await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length>0),JSON.stringify(S));
 const xOff=await p.evaluate(()=>{const v=viewExportRows();return v?JSON.stringify(v.aoa):null});
 ok('גיליון «התצוגה» זהה עם הלוח ובלעדיו',xOn!=null&&xOn===xOff);

 /* 7 · יציאה מהמסך, ובחירת תחנה ברצועה, מחזירות לטבלה */
 await p.click('#arrBtn');await wait(300);
 await p.evaluate(()=>setMode('today'));await wait(400);
 S=await st();
 ok('יציאה מ«רכש ואספקות» סוגרת את הלוח',!S.on&&!S.box&&S.rows&&!S.btn,JSON.stringify(S));
 await p.evaluate(()=>setMode('line'));await wait(400);
 ok('וחזרה למסך — בטבלה',!(await st()).on);
 await p.click('#arrBtn');await wait(300);
 const stn=await p.evaluate(()=>{const b=document.querySelector('#line [data-ls]');return b?b.dataset.ls:null});
 if(stn){await p.click(`#line [data-ls="${stn}"]`);await wait(400)}
 ok('בחירת תחנה ברצועה מציגה את הטבלה שלה',!!stn&&!(await st()).on,stn);

 /* 8 · הכרטיס: «אספקות לפי תאריך» */
 const C=await p.evaluate(()=>{const r=QF.all.find(x=>x.po>0&&ETA.byPn[etaKey(x.pn)]);const k=etaKey(r.pn),keep=ETA.byPn[k];
  const y=(etaToday().slice(0,4)|0)+1;
  /* ETA_SHIP (21.11): בדוח ההדגמה יש עמודת «משלוח» — המזהה שמוצג הוא מספר המשלוח (השדה החמישי), לא האספקה */
  ETA.byPn[k]=[[5,`${y}-01-10`,'D-1','PO-1','SH-1'],[3,`${y}-01-10`,'D-2','PO-1','SH-2'],[4,`${y}-02-15`,'D-3','PO-2','SH-3'],[2,'2000-01-05','D-0','PO-0','SH-0']];
  detail(r);const li=[...document.querySelectorAll('#detail .etad li')].map(x=>x.textContent.replace(/\s+/g,' ').trim());
  const sec=document.querySelector('#detail .etad'),dt=sec&&sec.dataset.dt;
  ETA.byPn[k]=keep;return {li,dt,h:sec&&sec.querySelector('h3').textContent,y}});
 ok('כרטיס: שורה לכל תאריך, בסדר עולה',C.li.length===3&&/^05\.01\.2000/.test(C.li[0])&&C.li[1].startsWith(`10.01.${C.y}`)&&C.li[2].startsWith(`15.02.${C.y}`),C.li.join(' | '));
 ok('כרטיס: שתי שורות באותו תאריך — סכום, מספר השורות ומספרי המשלוחים',/8 יח׳/.test(C.li[1])&&/2 שורות/.test(C.li[1])&&/משלוחים SH-1, SH-2/.test(C.li[1])&&!/D-1|D-2/.test(C.li[1]),C.li[1]);
 ok('כרטיס: שורה אחת — המשלוח שלה; תאריך שעבר — «לבדוק אם נקלט»',/4 יח׳/.test(C.li[2])&&/משלוח SH-3/.test(C.li[2])&&!/D-3/.test(C.li[2])&&/לבדוק אם נקלט/.test(C.li[0])&&!/לבדוק/.test(C.li[2]),C.li[0]);
 ok('כרטיס: המקטע בלשוניות «החלטה» ו«נתונים», עם כותרת',C.dt==='why data'&&/אספקות לפי תאריך/.test(C.h),`${C.dt} · ${C.h}`);
 const C2=await p.evaluate(()=>{const a=QF.all.find(x=>!(x.po>0)&&!x.etaHas);detail(a);const none=!document.querySelector('#detail .etad');
  const c=QF.all.find(x=>x.po>0&&!ETA.byPn[etaKey(x.pn)]);detail(c);const t=(document.querySelector('#detail .etad')||{}).textContent||'';
  const keep=ETA;ETA=null;detail(c);const t2=(document.querySelector('#detail .etad')||{}).textContent||'';ETA=keep;
  return {none,t,t2}});
 ok('כרטיס: פריט בלי רכש ובלי ETA — בלי מקטע',C2.none);
 ok('כרטיס: רכש פתוח בלי שורות בדוח ה-ETA — אומר את זה',/אין שורות בדוח ה-ETA/.test(C2.t),C2.t.slice(0,80));
 ok('כרטיס: בלי דוח ETA — «אין תאריך הגעה בדוח ZMRP»',/אין תאריך הגעה בדוח ZMRP/.test(C2.t2),C2.t2.slice(0,80));

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
