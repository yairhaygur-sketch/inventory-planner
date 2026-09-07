const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];const ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html'));
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(2500);

 // 0. מצב ההתחלה — לפני שנגענו בכלום
 const fresh=await p.evaluate(()=>{const d=document.querySelector('.wdetail').getBoundingClientRect();
   return {open:document.body.classList.contains('dopen'),onScreen:d.right>0&&d.left<innerWidth}});

 // 1. ניווט מקלדת ↓
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(250);
 let pn1=await p.evaluate(()=>document.querySelector('#detail .opn')?.textContent.replace('העתק','').trim());
 ok('חץ למטה בוחר פריט ומעדכן את הכרטיס',!!pn1,pn1);
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(250);
 let pn2=await p.evaluate(()=>document.querySelector('#detail .opn')?.textContent.replace('העתק','').trim());
 ok('חץ נוסף מתקדם לפריט הבא',pn2&&pn2!==pn1,pn1+' → '+pn2);
 await p.keyboard.press('ArrowUp');await p.waitForTimeout(250);
 ok('חץ למעלה חוזר אחורה',
   (await p.evaluate(()=>document.querySelector('#detail .opn')?.textContent.replace('העתק','').trim()))===pn1);

 // 2. Enter מסמן כטופל וממשיך
 const before=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
 await p.keyboard.press('Enter');await p.waitForTimeout(400);
 const after=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
 const stillSel=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr.sel').length);
 ok('Enter מסמן כטופל ומסיר מהתור',after<before,before+' → '+after);
 ok('הבחירה עוברת אוטומטית לפריט הבא',stillSel===1);

 // 3. / ממקד חיפוש
 await p.keyboard.press('/');await p.waitForTimeout(150);
 ok('"/" ממקד את שדה החיפוש',await p.evaluate(()=>document.activeElement.id==='q'));
 await p.keyboard.press('Escape');await p.waitForTimeout(150);
 ok('Escape יוצא מהשדה',await p.evaluate(()=>document.activeElement.id!=='q'));

 // 4. טאב תנועות
 await p.evaluate(()=>[...document.querySelectorAll('#tabs .tab')].find(t=>t.dataset.m==='moves')?.click());
 await p.waitForTimeout(500);
 const mv=await p.evaluate(()=>({rows:document.querySelectorAll('#movtbl tbody tr').length,
   panelShown:getComputedStyle(document.getElementById('movesPanel')).display!=='none',
   queueHidden:getComputedStyle(document.querySelector('#w_queue .qpanel')).display==='none',
   h:document.querySelector('#movesPanel .rows')?.clientHeight}));
 ok('טאב תנועות מציג את הטבלה במקום תור העבודה',mv.panelShown&&mv.queueHidden&&mv.rows>0,
    mv.rows+' שורות, גובה '+mv.h+'px');
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(300);
 ok('ניווט מקלדת עובד גם בטאב תנועות',
   await p.evaluate(()=>document.querySelectorAll('#movtbl tbody tr.sel').length===1));
 // חזרה
 await p.evaluate(()=>[...document.querySelectorAll('#tabs .tab')].find(t=>t.dataset.m==='today')?.click());
 await p.waitForTimeout(400);
 ok('חזרה לתור העבודה משחזרת את הרשימה',
   await p.evaluate(()=>getComputedStyle(document.querySelector('#w_queue .qpanel')).display!=='none'
     &&document.querySelectorAll('#tbl tbody tr[data-i]').length>0));

 // 5. גרפים מתקפלים
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(400);
 const d0=await p.evaluate(()=>{const o=document.querySelector('#detail .opad');return {h:o.clientHeight,sh:o.scrollHeight,ovf:getComputedStyle(o).overflow}});
 await p.evaluate(()=>document.querySelector('#detail details[data-chart="c1"]').open=true);
 await p.waitForTimeout(400);
 const drawn=await p.evaluate(()=>{const c=document.getElementById('c1');return c&&c.width>0&&
   c.getContext('2d').getImageData(0,0,c.width,c.height).data.some(v=>v!==0)});
 ok('גרף הצריכה מצויר בפתיחת המקטע',drawn);
 ok('כרטיס הפריט גולל ולא נחתך',d0.ovf==='auto',d0.ovf);
 ok('כרטיס הפריט: פחות גלילה מהבסיס',d0.sh<950,'תוכן '+d0.sh+'px בחלון '+d0.h+'px (בסיס: 1066/328)');

 // שורת הכותרת לא נשברת, והפילטרים קיימים — נבדק בשלושה רוחבי מסך
 for(const w of [1920,1512,1280]){
  await p.setViewportSize({width:w,height:860});await p.waitForTimeout(350);
  const h=await p.evaluate(()=>{const e=document.querySelector('#w_queue .phd');
   return {h:Math.round(e.getBoundingClientRect().height),ov:Math.round(e.scrollWidth-e.clientWidth)}});
  ok(`שורת הכותרת בשורה אחת ב-${w}px`,h.h<=48&&h.ov<=1,`גובה ${h.h}px · גלישה ${h.ov}px`);
 }
 await p.setViewportSize({width:1512,height:860});await p.waitForTimeout(350);
 const f=await p.evaluate(()=>({keys:FDEF.map(x=>x[0]),
   labels:[...document.querySelectorAll('#sidefilters .tl')].map(e=>e.textContent),
   modelOpts:optionsFor('model').length}));
 // שום טקסט בטאב לא חורג מגבולותיו — גלישה כאן דורסת את הטאב השכן.
 // נכשל כשפיצול המטבעות הכניס "$39M · €11M · ₪8.8M" לטאב של 90px.
 for(const w of [1920,1512,1280]){
  await p.setViewportSize({width:w,height:860});await p.waitForTimeout(300);
  const sp=await p.evaluate(()=>{const out=[];
   document.querySelectorAll('#tabs .tab').forEach(t=>{const tb=t.getBoundingClientRect();
    t.querySelectorAll('*').forEach(c=>{const cb=c.getBoundingClientRect();
     if(cb.width>0&&(cb.right>tb.right+1||cb.left<tb.left-1))
       out.push(((t.querySelector('.t')||{}).textContent||'')+'/'+c.className)})});
   return out});
  ok(`אין דריסת טקסט בטאבים ב-${w}px`,sp.length===0,sp.slice(0,2).join(' · ')||'נקי');
 }
 await p.setViewportSize({width:1512,height:860});await p.waitForTimeout(300);
 ok('פילטר "דגם" קיים בסרגל',f.keys.includes('model')&&f.labels.includes('דגם'),f.labels.join(' · '));
 ok('ולפילטר הדגם יש ערכים אמיתיים',f.modelOpts>1,f.modelOpts+' דגמים');
 // חותמת הגרסה — הדרך היחידה לדעת איזו גרסה הדפדפן מגיש
 const bld=await p.evaluate(()=>{const e=document.getElementById('bld');
   const b=e.getBoundingClientRect(),hd=document.querySelector('.top').getBoundingClientRect();
   return {txt:e.textContent,over:b.right>hd.right+1||b.left<hd.left-1,h:hd.height}});
 ok('חותמת גרסה מוצגת בכותרת',/^\d{2}\.\d{2} \d{2}:\d{2}$/.test(bld.txt),bld.txt);
 ok('החותמת אינה חורגת משורת הכותרת',!bld.over&&bld.h<=60,`גובה ${bld.h}`);
 // מגירת כרטיס הפריט — מוסתרת עד שבוחרים, והרשימה מקבלת את הרוחב
 await p.evaluate(()=>closeDetail());await p.waitForTimeout(250);
 const dr=await p.evaluate(()=>{const d=document.querySelector('.wdetail').getBoundingClientRect();
   const rows=document.querySelector('.rows'),tbl=document.getElementById('tbl');
   return {onScreen:d.right>0&&d.left<innerWidth,open:document.body.classList.contains('dopen'),
     rowsW:rows.clientWidth,tblW:tbl.scrollWidth,
     railW:Math.round((document.querySelector('.wfilters')||{getBoundingClientRect:()=>({width:0})}).getBoundingClientRect().width),
     cols:getComputedStyle(document.querySelector('.dash')).gridTemplateColumns.split(' ').length}});
 ok('בטעינה, לפני שנבחר פריט, המגירה סגורה ומחוץ למסך',!fresh.open&&!fresh.onScreen);
 ok('סגירה מוציאה את המגירה מהמסך',!dr.onScreen&&!dr.open);
 /* מסילת הסינון פתוחה כברירת מחדל — תשעת הסינונים בשימוש יומיומי.
    היא נקפלת ונפתחת, והמצב נשמר. הבדיקה הקודמת כאן קיבעה בטעות
    שהמסילה מוסתרת; זו הייתה רגרסיה, לא כוונה. */
 ok('הרשימה חולקת את הגריד עם מסילת הסינון',dr.cols===2,dr.cols+' עמודות');
 ok('מסילת הסינון פתוחה כברירת מחדל',dr.railW>150,`רוחב מסילה ${dr.railW}`);
 ok('תשעת הסינונים קיימים',
   9===await p.evaluate(()=>document.querySelectorAll('.wfilters .ddbtn').length));
 const railVis=()=>p.evaluate(()=>{const r=document.querySelector('.wfilters');
   return !!r&&r.offsetParent!==null});
 ok('כפתור «סינון» בסרגל העליון גלוי',
   await p.evaluate(()=>{const b=document.getElementById('filtBtn');return !!b&&b.offsetParent!==null}));
 await p.click('.wfilters .whead .wx');await p.waitForTimeout(350);
 ok('✕ מקפל את המסילה',!(await railVis()));
 ok('הקיפול משחרר את עמודת הרשת',
   await p.evaluate(()=>document.body.classList.contains('nofilters')));
 await p.click('#filtBtn');await p.waitForTimeout(350);
 ok('הכפתור בסרגל מחזיר את המסילה',await railVis());
 await p.click('.wfilters .whead .wx');await p.waitForTimeout(300);
 await p.reload();await p.waitForTimeout(1600);
 ok('הקיפול שורד רענון',!(await railVis()));
 await p.click('#filtBtn');await p.waitForTimeout(300);
 await p.reload();await p.waitForTimeout(1600);
 ok('הפתיחה שורדת רענון',await railVis());
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(4000);
 ok('אין גלילה אופקית בטבלה',dr.tblW<=dr.rowsW+2,`טבלה ${dr.tblW} בתוך ${dr.rowsW}`);
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(300);
 const dop=await p.evaluate(()=>{const d=document.querySelector('.wdetail').getBoundingClientRect();
   return {open:document.body.classList.contains('dopen'),onScreen:d.right>0&&d.left<innerWidth,
     w:Math.round(d.width),txt:document.getElementById('detail').innerText.length,
     rowsW:document.querySelector('.rows').clientWidth}});
 ok('לחיצה על שורה פותחת את המגירה',dop.open&&dop.onScreen&&dop.txt>50,`רוחב ${dop.w}`);
 ok('המגירה מרחפת ואינה מכווצת את הרשימה',dop.rowsW===dr.rowsW,`${dop.rowsW} מול ${dr.rowsW}`);
 await p.keyboard.press('Escape');await p.waitForTimeout(300);
 const dcl=await p.evaluate(()=>({open:document.body.classList.contains('dopen'),
   sel:document.querySelectorAll('#tbl tbody tr.sel').length}));
 ok('Esc סוגר את המגירה ומנקה את הבחירה',!dcl.open&&dcl.sel===0);
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(250);
 await p.click('#dclose');await p.waitForTimeout(250);
 const dx=await p.evaluate(()=>document.body.classList.contains('dopen'));
 ok('כפתור ה-✕ סוגר את המגירה',!dx);
// קבוצות בתוך "היום"
 await p.evaluate(()=>{GCOLL.clear();setMode('today')});await p.waitForTimeout(350);
 const g0=await p.evaluate(()=>({
   groups:[...document.querySelectorAll('#tbl tbody tr.grp')].map(t=>t.dataset.g),
   names:[...document.querySelectorAll('#tbl tbody tr.grp b')].map(t=>t.textContent),
   counts:[...document.querySelectorAll('#tbl tbody tr.grp .gn')].map(t=>+t.textContent),
   rows:document.querySelectorAll('#tbl tbody tr[data-i]').length}));
 ok('שלוש קבוצות ב"היום"',g0.groups.join()==='wait,prevent,m1',g0.names.join(' · '));
 ok('סכום הקבוצות = מספר הפריטים',g0.counts.reduce((a,b)=>a+b,0)===g0.rows,
    `${g0.counts.join('+')} = ${g0.rows}`);
 /* השורה שנלחצת חייבת להיות הפריט שנפתח בכרטיס — גם כשקבוצה מקופלת */
 const align=async n=>p.evaluate(i=>{const tr=[...document.querySelectorAll('#tbl tbody tr[data-i]')][i];
   if(!tr)return null;tr.click();
   return {row:tr.querySelector('.obj').textContent.trim(),
           card:CURRENT_DETAIL&&CURRENT_DETAIL.pn.replace(/[\u200e\u200f]/g,'').trim()}},n);
 const a1=await align(4);await p.waitForTimeout(250);
 ok('לחיצה על שורה פותחת את הפריט הנכון',a1&&a1.row===a1.card,a1?`${a1.row} / ${a1.card}`:'—');
 await p.evaluate(()=>toggleGroup('m1'));await p.waitForTimeout(350);
 const g1=await p.evaluate(()=>({coll:document.querySelectorAll('#tbl tbody tr.grp.coll').length,
   rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
   groups:document.querySelectorAll('#tbl tbody tr.grp').length,
   saved:JSON.parse(localStorage.getItem('planner_groups_v1')||'[]').join()}));
 ok('קיפול קבוצה מסתיר את שורותיה',g1.coll===1&&g1.rows<g0.rows&&g1.groups===3,
    `${g1.rows} מתוך ${g0.rows} · ${g1.groups} כותרות`);
 ok('מצב הקיפול נשמר',g1.saved==='m1',g1.saved);
 const a2=await align(2);await p.waitForTimeout(250);
 ok('ההתאמה שורה↔כרטיס נשמרת גם כשקבוצה מקופלת',a2&&a2.row===a2.card,a2?`${a2.row} / ${a2.card}`:'—');
 await p.evaluate(()=>toggleGroup('m1'));await p.waitForTimeout(300);
 const g2=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
 ok('פתיחה מחזירה את כל השורות',g2===g0.rows,`${g2} מתוך ${g0.rows}`);
 await p.evaluate(()=>closeDetail());await p.waitForTimeout(200);

 // תקציב הגובה — הכרום ירד, הרשימה קיבלה שורות
 const bud=await p.evaluate(()=>{const rows=document.querySelector('.rows'),rb=rows.getBoundingClientRect();
   return {vis:[...document.querySelectorAll('#tbl tbody tr[data-i]')]
     .filter(tr=>{const b=tr.getBoundingClientRect();return b.top>=rb.top-1&&b.bottom<=rb.bottom+1}).length,
    kpi:getComputedStyle(document.querySelector('.wkpi')).display!=='none',
    diagH:Math.round(document.getElementById('diag').getBoundingClientRect().height),
    chip:!document.getElementById('dchip').hidden,
    listPct:Math.round(rb.height/innerHeight*100)}});
 ok('שורת המדדים מוסתרת כברירת מחדל',!bud.kpi);
 ok('אזהרות מקופלות לשבב ולא לפס',bud.diagH===0&&bud.chip,`פס ${bud.diagH}px · שבב ${bud.chip}`);
 ok('הרשימה מקבלת 75%+ מהמסך',bud.listPct>=75,bud.listPct+'%');
 ok('16+ שורות גלויות ב-1512x860',bud.vis>=16,bud.vis+' שורות (היה 11 לפני העיצוב מחדש)');
 await p.click('#dchip');await p.waitForTimeout(300);
 const dOpen=await p.evaluate(()=>Math.round(document.getElementById('diag').getBoundingClientRect().height));
 ok('לחיצה על השבב פותחת את פירוט האזהרות',dOpen>0,dOpen+'px');
 await p.click('#dchip');await p.waitForTimeout(250);
 const dShut=await p.evaluate(()=>Math.round(document.getElementById('diag').getBoundingClientRect().height));
 ok('לחיצה נוספת סוגרת',dShut===0);
 /* שורת המדדים ניתנת להחזרה — ההסתרה היא ברירת מחדל, לא מחיקה */
 await p.evaluate(()=>toggleWidget('kpis'));await p.waitForTimeout(300);
 const back=await p.evaluate(()=>getComputedStyle(document.querySelector('.wkpi')).display!=='none');
 ok('אפשר להחזיר את שורת המדדים מתפריט הווידג׳טים',back);
 await p.evaluate(()=>toggleWidget('kpis'));await p.waitForTimeout(250);

 // אין גלילה אופקית בשום מצב ובשום רוחב
 for(const w of [1920,1512,1280,1180]){
  await p.setViewportSize({width:w,height:820});await p.waitForTimeout(200);
  for(const m of ['today','month','catalog']){
   await p.evaluate(k=>setMode(k),m);await p.waitForTimeout(250);
   const r=await p.evaluate(()=>{const rows=document.querySelector('.rows'),tbl=document.getElementById('tbl');
     return {h:tbl.scrollWidth>rows.clientWidth+2,t:Math.round(tbl.scrollWidth),c:rows.clientWidth}});
   ok(`${w}px · מצב ${m} · ללא גלילה אופקית`,!r.h,`טבלה ${r.t} ברשימה ${r.c}`);
  }}
 /* ============ הדלת לפריטים שסומנו "טופל" ============
    סימון מוציא פריט מתור העבודה ברונג הראשון של classify. המנגנון
    והאחסון היו קיימים, אבל לא הייתה דרך להגיע אליהם כדי לבטל. */
 /* בדיקות קודמות בקובץ מסמנות פריטים — מתחילים מלוח נקי */
 await p.evaluate(()=>{Object.keys(MARKS).forEach(k=>delete MARKS[k]);saveMarks();apply()});
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(600);
 const dHid=()=>p.evaluate(()=>document.getElementById('doneBtn').hidden);
 ok('בלי סימונים הכפתור "טופלו" מוסתר',await dHid());
 const n0=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
 for(let k=0;k<3;k++){
  await p.evaluate(()=>document.querySelector('#tbl tbody tr[data-i] .dn[data-done]').click());
  await p.waitForTimeout(400)}
 ok('אחרי סימון הכפתור מופיע',!(await dHid()));
 ok('המונה מציג את המספר הנכון',
   '3'===await p.evaluate(()=>document.getElementById('doneN').textContent));
 ok('הפריטים ירדו מתור העבודה',
   (await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length))===n0-3);
 await p.click('#doneBtn');await p.waitForTimeout(600);
 const dv=await p.evaluate(()=>({rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
   heads:[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim()).join(' · '),
   undo:document.querySelectorAll('#tbl .dn[data-undo]').length,
   when:/\d\d\.\d\d/.test((document.querySelectorAll('#tbl tbody tr[data-i] td')[4]||{}).textContent||''),
   pgR:document.getElementById('pgR').textContent}));
 ok('המסך מציג את שלושת הפריטים',dv.rows===3);
 ok('עמודות ייעודיות ולא של "כל הפריטים"',dv.heads.includes('הסימון'),dv.heads);
 ok('כפתור "בטל סימון" בכל שורה',dv.undo===3);
 ok('חותמת זמן הסימון מוצגת',dv.when);
 ok('הפוטר אינו כספי במסלול הזה',!/₪|\$/.test(dv.pgR),dv.pgR);
 await p.click('#tbl .dn[data-undo]');await p.waitForTimeout(700);
 ok('ביטול מוריד את הפריט מהמסך',
   2===await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length));
 await p.click('#doneBtn');await p.waitForTimeout(500);
 ok('לחיצה שנייה מחזירה ל"היום"','today'===await p.evaluate(()=>mode));
 ok('הפריט שבוטל חזר לתור העבודה',
   (await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length))===n0-2);
 await p.evaluate(()=>{Object.keys(MARKS).forEach(k=>delete MARKS[k]);saveMarks();apply()});
 await p.waitForTimeout(600);
 ok('ניקוי הסימונים מסתיר את הכפתור שוב',await dHid());

 /* ============ מלאי · בדרך · לקוח ממתין ============
    שלוש העמודות שמכריעות אם לפתוח הזמנה. "בדרך" מעומעם בכוונה:
    בדוח האמיתי הוא >= החוסר ב-37 מתוך 48 השורות ואין בו ETA, ולכן
    מספר מלא היה נקרא ככיסוי ומבטל את מסלול "אוזל החודש". */
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(500);
 const dec=await p.evaluate(()=>{
  const h=[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim());
  const tr=document.querySelector('#tbl tbody tr[data-i]');
  const otw=document.querySelector('#tbl td.otw span');
  const zero=document.querySelector('#tbl td.zero');
  const gs=n=>{const e=document.querySelector(n);return e?getComputedStyle(e):null};
  const so=gs('#tbl td.otw'),sz=gs('#tbl td.zero'),sn=gs('#tbl tbody tr[data-i] td.num');
  return {heads:h, hasStock:h.includes('מלאי'), hasCust:h.includes('לקוח'),
    hasOtw:h.some(x=>x.indexOf('בדרך')===0),
    otwMarked:!!otw&&/⌛/.test(otw.textContent),
    otwTitleHasEta:!!otw&&/אין תאריך הגעה/.test(otw.getAttribute('title')||''),
    otwDim:so&&sn?so.color!==sn.color:false,
    zeroRed:!!sz&&sz.fontWeight>=600}});
 ok('עמודת מלאי קיימת',dec.hasStock,dec.heads.join(' · '));
 ok('עמודת לקוח ממתין קיימת',dec.hasCust);
 ok('עמודת "בדרך" קיימת',dec.hasOtw);
 ok('"בדרך" מסומן ⌛ ולא נקרא ככיסוי',dec.otwMarked);
 ok('ההסבר אומר במפורש שאין ETA',dec.otwTitleHasEta);
 ok('"בדרך" מעומעם ביחס למספרים האחרים',dec.otwDim);
 ok('מלאי אפס בולט',dec.zeroRed);

 /* סיווג שחורג מרוב הקבוצה נכתב בשורה. בדוח האמיתי אלה שמונה פריטי
    «חשד אזילה» בתוך «אוזל החודש» — הבחנה שנמחקה בטעות עם העמודה. */
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(500);
 const oc=await p.evaluate(()=>{
  const G={};for(const tr of document.querySelectorAll('#tbl tbody tr')){
   if(tr.classList.contains('grp')){G.cur=tr.textContent.trim();continue}}
  const notes=[...document.querySelectorAll('#tbl td.note .ocat')].map(x=>x.textContent.trim());
  return {shown:notes.length,uniq:[...new Set(notes)]}});
 ok('סיווג חריג נכתב בשורה ולא נבלע בכותרת',oc.shown>0,
   `${oc.shown} שורות · ${oc.uniq.join(' · ')}`);
 const noDup=await p.evaluate(()=>{
  let bad=0;
  for(const tr of document.querySelectorAll('#tbl tbody tr[data-i]')){
   const o=tr.querySelector('td.note .ocat');if(!o)continue;
   let h=tr.previousElementSibling;while(h&&!h.classList.contains('grp'))h=h.previousElementSibling;
   if(h&&h.textContent.includes(o.textContent.trim()))bad++}
  return bad});
 ok('סיווג שכבר כתוב בכותרת אינו חוזר בשורה',noDup===0,noDup+' חזרות');

 await p.setViewportSize({width:1512,height:860});

 /* גרף הצריכה — ברמת המסך, לא קבור בתוך <details> בכרטיס הפריט */
 await p.evaluate(()=>setMode('catalog'));await p.waitForTimeout(700);
 const ch=await p.evaluate(()=>{const c=document.getElementById('ctop');
  if(c.hidden)return {hidden:true};
  const sv=c.querySelector('svg'),b=sv.getBoundingClientRect();
  const vb=sv.getAttribute('viewBox').split(' ').map(Number);
  return {hidden:false,inView:b.top>=0&&b.bottom<=innerHeight,
   bars:c.querySelectorAll('rect').length,labs:c.querySelectorAll('.clab').length,
   vals:c.querySelectorAll('.cval').length,
   /* יחס הצירים חייב להישמר, אחרת הטקסט נמתח */
   skew:Math.abs((b.width/b.height)-(vb[2]/vb[3])),
   ticks:[...c.querySelectorAll('.cax')].map(t=>t.textContent),
   tallest:Math.max(...[...c.querySelectorAll('rect')].map(r=>+r.getAttribute('height'))),
   /* שטח הציור נגזר מהציור עצמו: מקו הבסיס עד קו הרשת העליון */
   plotH:Math.max(...[...c.querySelectorAll('rect')].map(r=>+r.getAttribute('y')+ +r.getAttribute('height')))
        -Math.min(...[...c.querySelectorAll('line')].map(l=>+l.getAttribute('y1')))}});
 ok('הגרף נראה במצב הקטלוג בלי גלילה ובלי לחיצה',!ch.hidden&&ch.inView);
 ok('עמודה לכל חודש בדוח',ch.bars===11,ch.bars+' עמודות');
 ok('רק השיא והחודש האחרון מתויגים',ch.vals===2,ch.vals+' תוויות ערך');
 ok('יחס הצירים נשמר — הטקסט אינו נמתח',ch.skew<0.05,'סטייה '+ch.skew.toFixed(3));
 ok('תווי הסקאלה עגולים',ch.ticks.every(t=>/^\d+(\.\d)?k?$/.test(t)),ch.ticks.join(' · '));
 ok('העמודה הגבוהה ממלאת את רוב שטח הציור',ch.tallest>=ch.plotH*0.6,
    `${Math.round(ch.tallest)} מתוך ${Math.round(ch.plotH)}`);
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);
 ok('הגרף אינו מופיע במצב "היום"',await p.evaluate(()=>document.getElementById('ctop').hidden));

 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(300);
  ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 await p.screenshot({path:SD+'/04-after.png'});
 await b.close();console.log(out.join('\n'));
 process.exit(out.some(l=>l.startsWith('FAIL'))?1:0)})();
