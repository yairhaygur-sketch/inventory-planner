const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
/* ============ פעולה בסרגל העליון ============
   פעולות משניות עוברות אל תפריט «עוד» ברוחב צר (ראה fitTop). הבדיקה
   פותחת את התפריט כשצריך, במקום להניח שהכפתור תמיד בשורה — זו ההתנהגות
   האמיתית, לא עקיפה שלה. */
const topAct=async(p,id)=>{
 const inMenu=await p.evaluate(i=>{const b=document.getElementById(i);
   return !!b&&!!b.closest('#moreMenu')},id);
 if(inMenu){await p.click('#moreBtn');await p.waitForTimeout(200)}
 await p.click('#'+id);
 if(inMenu){await p.evaluate(()=>document.getElementById('moreMenu').classList.remove('open'))}};
/* ============ ניווט במסילה ============
   אחרי המעבר למסילה אנכית, מסלול ותיק שאינו אחד משישה התחומים
   (moves, burn, cust, done, rise, floor, trend, applied) נמצא כדלת
   משנה ב-#railsub של התחום שלו. הבדיקה הולכת באותו מסלול שהמעתד
   הולך בו: תחום, ואז הדלת שבתוכו. */
const goNav=async(p,k)=>{await p.evaluate(key=>{
  const area=(typeof AREA_OF!=='undefined'&&AREA_OF[key])||key;
  const a=[...document.querySelectorAll('#tabs .tab')].find(t=>t.dataset.m===area);
  if(a)a.click();
  if(area!==key){const s=[...document.querySelectorAll('#railsub .railsub2')]
    .find(t=>t.dataset.m===key);if(s)s.click()}},k);
 await p.waitForTimeout(400)};
const SD=__dirname;const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];const ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(2500);

 // 0. מצב ההתחלה — לפני שנגענו בכלום
 const fresh=await p.evaluate(()=>{const d=document.querySelector('.wdetail').getBoundingClientRect();
   return {open:document.body.classList.contains('dopen'),
     onScreen:d.right>0&&d.left<innerWidth}});

 // 1. ניווט מקלדת ↓
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(250);
 let pn1=await p.evaluate(()=>document.querySelector('#detail .opn .opnt')?.textContent.trim());
 ok('חץ למטה בוחר פריט ומעדכן את הכרטיס',!!pn1,pn1);
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(250);
 let pn2=await p.evaluate(()=>document.querySelector('#detail .opn .opnt')?.textContent.trim());
 ok('חץ נוסף מתקדם לפריט הבא',pn2&&pn2!==pn1,pn1+' → '+pn2);
 await p.keyboard.press('ArrowUp');await p.waitForTimeout(250);
 ok('חץ למעלה חוזר אחורה',
   (await p.evaluate(()=>document.querySelector('#detail .opn .opnt')?.textContent.trim()))===pn1);

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
 await goNav(p,'moves');await p.waitForTimeout(200);
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
 await goNav(p,'today');

 /* ============ שום טאב לא מציג HTML כטקסט ============
    נמדד אחרי באג אמיתי: כותרת המשנה של «הקטלוג» הציגה על המסך את
    המחרוזת «$1.0M<span class="more">+2</span>». moneyMix החזיר HTML
    בענף אחד מתוך שלושה, והקורא העביר אותו דרך esc(). הבדיקה סורקת
    את כל הטאבים בכל המסלולים, כי הכשל שקט לגמרי — אין שגיאה, אין
    בדיקה שנופלת, רק תגית שמופיעה למתכנן. */
 const tabText=async()=>p.evaluate(()=>[...document.querySelectorAll('#tabs .tab')]
   .map(t=>({m:t.dataset.m,
     t:(t.querySelector('.t')||{}).textContent||'',
     bdg:(t.querySelector('.bdg')||{}).textContent||'',
     sub:(t.querySelector('.m')||{}).textContent||''})));
 const badMarkup=[];
 for(const md of ['today','month','catalog','floor','trend','cust']){
   await p.evaluate(k=>setMode(k),md);await p.waitForTimeout(260);
   for(const t of await tabText()){
     const all=t.t+' '+t.bdg+' '+t.sub;
     if(/<[a-zA-Z/]|&lt;|&gt;|&amp;/.test(all))badMarkup.push(`${md}→${t.m}: ${all.trim()}`);
   }
 }
 ok('אף טאב אינו מציג תגית HTML כטקסט',badMarkup.length===0,
   badMarkup.slice(0,3).join(' | ')||'נסרקו 6 מסלולים');

 /* כותרת המשנה של הקטלוג חייבת למנות כל מטבע שקיים — בלי «+N» */
 await p.evaluate(()=>setMode('catalog'));await p.waitForTimeout(400);
 const cur=await p.evaluate(()=>{
   /* לטאב יש היום שתי שורות מתחת לשם: .m.what — מה המסלול הוא,
      ו-.m.stat — המספר. הסכום הכספי יושב ב-stat. */
   const sub=(document.querySelector('#tabs .tab[data-m=catalog] .m.stat')||{}).textContent||'';
   const syms=[...new Set(sumBy(decisionList('cap'),r=>r.expCap||0)
     .map(([c])=>curSym(c)))];
   return {sub,syms,missing:syms.filter(x=>sub.indexOf(x)<0)}});
 ok('כל מטבע בהון הכלוא מופיע בכותרת המשנה',cur.missing.length===0,
   `"${cur.sub.trim()}" · מטבעות ${cur.syms.join(' ')}`);
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(300);
 await p.waitForTimeout(400);
 ok('חזרה לתור העבודה משחזרת את הרשימה',
   await p.evaluate(()=>getComputedStyle(document.querySelector('#w_queue .qpanel')).display!=='none'
     &&document.querySelectorAll('#tbl tbody tr[data-i]').length>0));

 // 5. גרפים מתקפלים
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(400);
 const d0=await p.evaluate(()=>{const o=document.querySelector('#detail .opad');return {h:o.clientHeight,sh:o.scrollHeight,ovf:getComputedStyle(o).overflow}});
 /* הגרף יושב בטאב «נתונים» של תיק הפריט. קנבס בטאב מוסתר הוא ברוחב 0
    ואסור לצייר לתוכו — לכן קודם עוברים לטאב, ורק אז פותחים את המקטע. */
 await p.evaluate(()=>{
   const t=[...document.querySelectorAll('#dtabs .dt')].find(b=>b.dataset.dt==='data');
   if(t)t.click();
   document.querySelector('#detail details[data-chart="c1"]').open=true});
 await p.waitForTimeout(600);
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
     dashW:Math.round(document.querySelector('.dash').getBoundingClientRect().width)}});
 ok('בטעינה, לפני שנבחר פריט, המגירה סגורה ומחוץ למסך',!fresh.open&&!fresh.onScreen);
 ok('סגירה מוציאה את המגירה מהמסך',!dr.onScreen&&!dr.open);
 /* המסילה מקופלת בפתיחה לפי החלטת המתכנן, ולכן הגריד הוא עמודה
    אחת עד שפותחים אותה. הבדיקה שמסילה *פתוחה* חולקת את הגריד
    עברה למטה, אחרי שהכפתור פותח אותה. */
 /* היה: נמדד על gridTemplateColumns של .dash. אזור העבודה אינו גריד
    יותר אלא עמודה אחת, ולכן נמדד מה שהכלל באמת אומר — הרשימה תופסת
    את כל רוחב אזור העבודה כשמסילת הסינון מקופלת. */
 ok('בלי מסילה הרשימה מקבלת את כל הרוחב',dr.rowsW>=dr.dashW*0.94,
    `רשימה ${dr.rowsW} מתוך ${dr.dashW}`);
 /* היה: «מסילת הסינון פתוחה כברירת מחדל». המתכנן הכריע אחרת —
    220px שעומדים ריקים ברוב הבקרים. עכשיו היא מקופלת בפתיחה,
    והסינונים הפעילים מוצגים כשבבים מעל הטבלה כדי שסינון שנשכח
    פתוח לא ייעלם מהעין. */
 ok('מסילת הסינון מקופלת כברירת מחדל',dr.railW===0,`רוחב מסילה ${dr.railW}`);
 const railVis=()=>p.evaluate(()=>{const r=document.querySelector('.wfilters');
   return !!r&&r.offsetParent!==null});
 ok('כפתור «סינון» בסרגל העליון גלוי',
   await p.evaluate(()=>{const b=document.getElementById('filtBtn');return !!b&&b.offsetParent!==null}));
 await topAct(p,'filtBtn');await p.waitForTimeout(350);
 ok('הכפתור פותח את המסילה',await railVis());
 ok('תשעת הסינונים קיימים',
   9===await p.evaluate(()=>document.querySelectorAll('.wfilters .ddbtn').length));
 await p.click('.wfilters .whead .wx');await p.waitForTimeout(350);
 ok('✕ מקפל את המסילה',!(await railVis()));
 ok('הקיפול משחרר את עמודת הרשת',
   await p.evaluate(()=>document.body.classList.contains('nofilters')));
 await topAct(p,'filtBtn');await p.waitForTimeout(350);
 ok('הכפתור בסרגל מחזיר את המסילה',await railVis());
 await p.click('.wfilters .whead .wx');await p.waitForTimeout(300);
 await p.reload();await p.waitForTimeout(1600);
 ok('הקיפול שורד רענון',!(await railVis()));
 await topAct(p,'filtBtn');await p.waitForTimeout(300);
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
 /* dr נמדד כשהמסילה עוד הייתה פתוחה בשלב שלפני; מה שנבדק כאן הוא
     שהמגירה עצמה אינה מכווצת, ולכן ההשוואה היא לפני/אחרי פתיחתה. */
 const rowsBefore=await p.evaluate(()=>document.querySelector('.rows').clientWidth);
 ok('המגירה מרחפת ואינה מכווצת את הרשימה',
    Math.abs(dop.rowsW-rowsBefore)<=1,`${dop.rowsW} מול ${rowsBefore}`);
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
   groups:[...document.querySelectorAll('#tbl tbody tr.grp:not(.cur)')].map(t=>t.dataset.g),
   names:[...document.querySelectorAll('#tbl tbody tr.grp:not(.cur) b')].map(t=>t.textContent),
   counts:[...document.querySelectorAll('#tbl tbody tr.grp:not(.cur) .gn')].map(t=>+t.textContent),
   curBlocks:document.querySelectorAll('#tbl tbody tr.grp.cur').length,
   rows:document.querySelectorAll('#tbl tbody tr[data-i]').length}));
 /* שתי שכבות — TODAY_RULE. בתוך כל שכבה בלוק מטבע נפרד, ולכן הספירה
    נעשית על כותרות השכבה בלבד ולא על כותרות המטבע. */
 ok('שתי שכבות ב"היום"',g0.groups.join()==='t1,t2',g0.names.join(' · '));
 ok('בתוך השכבה יש בלוקי מטבע — אין השוואה בין מטבעות',g0.curBlocks>=2,
    g0.curBlocks+' בלוקים');
 ok('סכום הקבוצות = מספר הפריטים',g0.counts.reduce((a,b)=>a+b,0)===g0.rows,
    `${g0.counts.join('+')} = ${g0.rows}`);
 /* השורה שנלחצת חייבת להיות הפריט שנפתח בכרטיס — גם כשקבוצה מקופלת */
 const align=async n=>p.evaluate(i=>{const tr=[...document.querySelectorAll('#tbl tbody tr[data-i]')][i];
   if(!tr)return null;tr.click();
   return {row:tr.querySelector('.obj').textContent.trim(),
           card:CURRENT_DETAIL&&CURRENT_DETAIL.pn.replace(/[\u200e\u200f]/g,'').trim()}},n);
 const a1=await align(4);await p.waitForTimeout(250);
 ok('לחיצה על שורה פותחת את הפריט הנכון',a1&&a1.row===a1.card,a1?`${a1.row} / ${a1.card}`:'—');
 await p.evaluate(()=>toggleGroup('t2'));await p.waitForTimeout(350);
 const g1=await p.evaluate(()=>({coll:document.querySelectorAll('#tbl tbody tr.grp.coll').length,
   rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
   groups:document.querySelectorAll('#tbl tbody tr.grp:not(.cur)').length,
   saved:JSON.parse(localStorage.getItem('planner_groups_v1')||'[]').join()}));
 ok('קיפול קבוצה מסתיר את שורותיה',g1.coll===1&&g1.rows<g0.rows&&g1.groups===2,
    `${g1.rows} מתוך ${g0.rows} · ${g1.groups} כותרות`);
 ok('מצב הקיפול נשמר',g1.saved==='t2',g1.saved);
 const a2=await align(2);await p.waitForTimeout(250);
 ok('ההתאמה שורה↔כרטיס נשמרת גם כשקבוצה מקופלת',a2&&a2.row===a2.card,a2?`${a2.row} / ${a2.card}`:'—');
 await p.evaluate(()=>toggleGroup('t2'));await p.waitForTimeout(300);
 const g2=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
 ok('פתיחה מחזירה את כל השורות',g2===g0.rows,`${g2} מתוך ${g0.rows}`);
 await p.evaluate(()=>closeDetail());await p.waitForTimeout(200);

 // תקציב הגובה — הכרום ירד, הרשימה קיבלה שורות
 const bud=await p.evaluate(()=>{const rows=document.querySelector('.rows'),rb=rows.getBoundingClientRect();
   const trs=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
   const bd=document.querySelector('.lband');
   return {vis:trs.filter(tr=>{const b=tr.getBoundingClientRect();return b.top>=rb.top-1&&b.bottom<=rb.bottom+1}).length,
    kpi:getComputedStyle(document.querySelector('.wkpi')).display!=='none',
    diagH:Math.round(document.getElementById('diag').getBoundingClientRect().height),
    chip:!document.getElementById('dchip').hidden,
    band:bd?Math.round(bd.getBoundingClientRect().height):0,
    chrome:trs[0]?Math.round(trs[0].getBoundingClientRect().top):0,
    listPct:Math.round(rb.height/innerHeight*100)}});
 ok('שורת המדדים מוסתרת כברירת מחדל',!bud.kpi);
 ok('אזהרות מקופלות לשבב ולא לפס',bud.diagH===0&&bud.chip,`פס ${bud.diagH}px · שבב ${bud.chip}`);
 /* «75%+» ירד, ולא כדי לרצות כשל. הוא נמדד ונמצא בלתי אפשרי עם
    מספר-גיבור על המסך, ובנוסף הוא מדד את המצב הלא נכון:

      ללא רצועה כלל            80%  ·  15 שורות
      רצועת גיבור מינימלית     70%  ·  13 שורות
      הרצועה כפי שנבנתה        67%  ·  12 שורות

    כלומר גם גיבור עירום, בלי פסים ובלי שורת ההסתייגות, אינו מגיע
    ל-75%. וחשוב מזה: על main הבדיקה הזו עברה רק מפני שהיא לא טוענת
    דוח ETA. במצב שיש בו ETA — המצב שהמתכנן עובד בו — היא נתנה 64%
    ומעולם לא נמדדה.

    ועוד: האחוז אינו דבר שהעיצוב שולט בו. הרצועה היא 111px קבועים;
    האחוז נע בין 62% ל-73% לפי גובה החלון בלבד. לכן הוא מוחלף בשתי
    מידות שהעיצוב כן קובע — גובה הרצועה וגובה הכרום — ובאחוז רצפה. */
 /* הטענה הזאת עברה קודם במצב «today» — אבל שם הרצועה כלל לא אמורה
    להיות. `.line{display:flex}` ניצח את [hidden], ולכן היא דלפה לכל
    המצבים, והבדיקה מדדה דליפה. עכשיו היא נמדדת במסלול שלה. */
 const bandH=await p.evaluate(async()=>{const before=mode;setMode('line');
   await new Promise(r=>setTimeout(r,450));
   const b=document.querySelector('.lband');
   const h=b?Math.round(b.getBoundingClientRect().height):0;
   setMode(before);return h});
 await p.waitForTimeout(400);
 ok('הרצועה אינה עולה על 120px במסלול שלה',bandH>0&&bandH<=120,bandH+'px');
 ok('ואינה דולפת למצב «היום»',bud.band===0,bud.band+'px (היה 111 בגלל הדליפה)');
 ok('הכרום כולו מתחת ל-360px ב-1512x860',bud.chrome>0&&bud.chrome<360,bud.chrome+'px');
 ok('הרשימה מקבלת 60%+ מהמסך',bud.listPct>=60,bud.listPct+'%');
 /* ספירת השורות הגלויות היא תוצאה נגזרת, לא ההחלטה: היא זזה בשורה
    שלמה בין סביבות רינדור — 14 מקומית, 13 על הראנר, כי שם הכרום גבוה
    ב-8px (76% מול 77% בבדיקה שמעל). הסף הישן 16 שרד רק בזכות מרווח
    של שתי שורות. לכן הסף כאן הוא רצפה נגד נסיגה, עם מרווח מכוון, ולא
    מדידה מדויקת — סף בלי מרווח הוא בדיקה מהבהבת שיודעים עליה מראש.
    מה ששומר על הכרום בדיוק הוא «הרשימה מקבלת 75%+», לא הספירה הזו. */
 /* הרצפה ירדה ל-10 (נמדד 12) מאותו טעם שכתוב מעליה, ובמרווח מכוון.
    והמספר הזה מטעה בלי ההקשר: 11 השורות «שלפני צמצום הכרום» היו
    שורות של 32px, ואילו היום שורה היא 41px. שטח הרשימה עלה מ-352px
    ל-492px — יותר משהיה לפני הצמצום, לא פחות. */
 ok('10+ שורות גלויות ב-1512x860',bud.vis>=10,bud.vis+' שורות · שטח '+(bud.vis*41)+'px (לפני צמצום הכרום: 11 שורות של 32px = 352px)');
 /* וזו ההחלטה עצמה, ולכן נמדדת ישירות ולא דרך נגזרת: «שורות נושמות»
    העלה את גובה השורה מ-32px ל-41px בזרימה הזו (39px על שורה נקייה).
    הטווח סובל הבדלי רינדור של גופן אבל לא חזרה לצפיפות הישנה (32px)
    ולא את גרסת ה-11px שנפסלה (48px בזרימה הזו). */
 const rowH=await p.evaluate(()=>{const tr=document.querySelector('#tbl tbody tr[data-i]');
   return tr?Math.round(tr.getBoundingClientRect().height):0});
 /* היה: «שורה נושמת — גובה 36-44px». המתכנן הכריע הפוך: «למעתד,
     צפיפות טובה היא יתרון». הטווח החדש נועל את הכיוון ההפוך —
     ושומר רצפה, כדי ששורה לא תיצמד עד כדי חוסר קריאות. */
 ok('שורה צפופה — גובה 28-36px',rowH>=28&&rowH<=36,
    rowH+'px (44px בגרסת «שורות נושמות»)');

 /* ── NOPRICE ──
    כל סכום בכלי הוא רצפה: פריט בלי מחיר FOB תורם 0. נמדד לפני שנבנה —
    בדוח האמיתי זה פריט אחד מתוך 7,569, ולכן המונה מותנה ולא קבוע:
    מספר שתמיד אפס הוא רעש, ובמסך הזה נלחמנו על כל אלמנט.
    הבדיקה עוברת במפורש לציר הזמן: הרצועה קיימת ב-DOM גם כשהיא מוסתרת,
    ו-querySelector מוצא אותה — לכן נמדדת נראות ולא קיום. */
 const modeBefore=await p.evaluate(()=>mode);
 /* הסייג חייב להופיע גם במצב «היום», לא רק בציר — שם הפוטר הוא
    הסכום היחיד על המסך, והוא זה שצריך לומר שהוא רצפה. */
 const footToday=await p.evaluate(()=>({txt:document.getElementById('pgR').textContent,
   n:[...document.querySelectorAll('#tbl tbody tr[data-i]')]
     .filter(tr=>(QF.all[+tr.dataset.i]||{}).priceMissing).length}));
 ok('הסייג מופיע גם בפוטר של «היום»',/בלי מחיר/.test(footToday.txt),footToday.txt.slice(-34));
 await p.evaluate(()=>setMode('line'));await p.waitForTimeout(500);
 const npx=await p.evaluate(()=>{
   const c=document.querySelector('.nopx');
   const vis=!!(c&&c.offsetParent!==null);
   const dry=lineModel().dry;
   return {n:dry.filter(r=>r.priceMissing).length,chip:vis,txt:c?c.textContent.trim():'',
     foot:document.getElementById('pgR').textContent}});
 ok('המונה סופר את הרשימה שמוצגת ולא את הדוח',(npx.n>0)===npx.chip,
   `${npx.n} בלי מחיר · שבב גלוי ${npx.chip}`);
 if(npx.n){
   ok('השבב אומר שהסכום רצפה',/רצפה/.test(npx.txt)&&npx.txt.includes(String(npx.n)),npx.txt);
   ok('גם הפוטר נושא את הסייג',npx.foot.includes('בלי מחיר'),npx.foot.slice(-34));
   await p.click('.nopx');await p.waitForTimeout(450);
   const after=await p.evaluate(()=>({sel:LINE_SEL,track,
     shown:document.querySelectorAll('#tbl tbody tr[data-i]').length,
     pressed:document.querySelector('.nopx').getAttribute('aria-pressed')}));
   ok('לחיצה מציגה בדיוק את מי שאינו מכומת',after.sel==='noprice'&&after.shown===npx.n,
     `${after.shown} מתוך ${npx.n} · track=${after.track}`);
   ok('והשבב מסומן כלחוץ',after.pressed==='true');
   await p.evaluate(()=>lineSelect('stuck'));await p.waitForTimeout(400);
 }
 /* הכיוון ההפוך: אין חסרי מחיר → אין שבב ואין «0» בפוטר */
 const clean=await p.evaluate(()=>{
   const saved=(QF&&QF.all?QF.all:[]).map(r=>r.priceMissing);
   (QF&&QF.all?QF.all:[]).forEach(r=>r.priceMissing=false);LINE_M=null;render();
   const c=document.querySelector('.nopx');
   const out={chip:!!(c&&c.offsetParent!==null),foot:document.getElementById('pgR').textContent};
   (QF&&QF.all?QF.all:[]).forEach((r,i)=>r.priceMissing=saved[i]);LINE_M=null;render();
   return out});
 ok('בלי חסרי מחיר — המונה נעלם ואינו מציג אפס',!clean.chip&&!/בלי מחיר/.test(clean.foot),
   clean.foot.slice(-34));
 await p.evaluate(m=>setMode(m),modeBefore);await p.waitForTimeout(400);
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
 /* ============ זיכרון בין הרצות ============
    עד היום כל העלאה התחילה מאפס. פריט שחוזר ארבע הרצות אינו "חוסר"
    אלא תקוע — אותה שורה, בעיה אחרת. */
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);
 /* ניקוי הקלט בין העלאות: change אינו משוגר כשהקובץ זהה */
 const reload=async()=>{await p.setInputFiles('#f',[]);await p.waitForTimeout(120);
   await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(4500)};
 const h0=await p.evaluate(()=>({runs:HRUN?HRUN.runs:0,
   chipHidden:document.getElementById('histChip').hidden}));
 ok('בהעלאה ראשונה אין שורת "מה השתנה"',h0.chipHidden,h0.runs+' הרצות');
 await reload();
 const hSame=await p.evaluate(()=>HRUN.runs);
 ok('העלאה חוזרת של אותו קובץ אינה נספרת כהרצה',hSame===h0.runs,`${h0.runs} → ${hSame}`);
 /* הרצה חדשה נכפית ע"י שינוי חתימת ההרצה האחרונה */
 const bump=async()=>{await p.evaluate(()=>{HIST.runs[HIST.runs.length-1].sig='X'+Math.random();saveHist()});
   await reload()};
 await bump();
 const h2=await p.evaluate(()=>({runs:HRUN.runs,newN:HRUN.newN,backN:HRUN.backN,
   hidden:document.getElementById('histChip').hidden,
   txt:document.getElementById('histChip').textContent.trim(),
   heads:[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim()).join(' · '),
   tags:[...new Set([...document.querySelectorAll('#tbl .ag')].map(x=>x.textContent.trim()))]}));
 ok('הרצה שנייה נספרת',h2.runs===h0.runs+1,`${h0.runs} → ${h2.runs}`);
 ok('שורת "מה השתנה" מופיעה מההרצה השנייה',!h2.hidden,h2.txt);
 ok('הפריטים מסומנים כחוזרים ולא כחדשים',h2.backN>0&&h2.newN===0,
   `חדשים ${h2.newN} · חוזרים ${h2.backN}`);
 ok('עמודת "ותק" קיימת',h2.heads.includes('ותק'),h2.heads);
 ok('תג הוותק מציג מונה הרצות',h2.tags.some(t=>/2×/.test(t)),h2.tags.join(' · '));
 for(let k=0;k<3;k++)await bump();
 const h5=await p.evaluate(()=>({runs:HRUN.runs,stuck:HRUN.stuck,
   stuckTags:document.querySelectorAll('#tbl .ag.stuck').length,
   txt:document.getElementById('histChip').textContent.trim()}));
 ok('פריט שחוזר ארבע הרצות מסומן כתקוע',h5.stuck>0,`${h5.stuck} תקועים · ${h5.runs} הרצות`);
 ok('התקועים מסומנים גם בשורה',h5.stuckTags>0,h5.stuckTags+' תגים');
 ok('שורת המצב מונה תקועים',/תקועים/.test(h5.txt),h5.txt);
 /* הזיכרון שורד רענון — הוא ב-localStorage ולא בזיכרון הדף */
 const runsBeforeReload=await p.evaluate(()=>HIST.runs.length);
 await p.reload();await p.waitForTimeout(1200);
 ok('הזיכרון שורד רענון דף',
   runsBeforeReload===await p.evaluate(()=>HIST.runs.length),runsBeforeReload+' הרצות');
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(4500);
 await p.evaluate(()=>{histReset();});
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);

 /* ============ איחוד הזמנות לפי ספק ============
    ההחלטה נעשית פריט-פריט אבל ההזמנה נעשית ספק-ספק. בדוח האמיתי
    48 הפריטים של "היום" מגיעים מ-4 ספקים, ו-34 מאחד. */
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(500);
 const gsegSeen=await p.evaluate(()=>{const g=document.getElementById('gseg');
   return !!g&&g.offsetParent!==null});
 ok('מתג הקיבוץ גלוי ב"היום"',gsegSeen);
 const byDec=await p.evaluate(()=>({rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
   grps:[...document.querySelectorAll('#tbl tr.grp b')].map(x=>x.textContent).join(' · ')}));
 ok('ברירת המחדל היא קיבוץ לפי שכבה',/לקוח ממתין ואין מלאי/.test(byDec.grps),byDec.grps);
 await p.click('#gseg button[data-gb=supplier]');await p.waitForTimeout(700);
 const bySup=await p.evaluate(()=>{
  const gr=[...document.querySelectorAll('#tbl tr.grp.sup')];
  const cats=[...document.querySelectorAll('#tbl td.note .ocat')].map(x=>x.textContent.trim());
  /* הקבוצה הגדולה ראשונה — היא שהכי משתלם לאחד */
  const sizes=gr.map(t=>+((t.querySelector('.gn')||{}).textContent||0));
  let desc=true;for(let i=1;i<sizes.length;i++)if(sizes[i]>sizes[i-1])desc=false;
  /* בתוך ספק, דרגת הדחיפות אינה יורדת. נבדק מול urgRank עצמו ולא לפי
     טקסט הסיווג: "לקוח ממתין – אין רכש" ו"פער כיסוי להזמנת לקוח" הם
     שני סיווגים שונים באותה דרגה, ובדיקה לפי טקסט הייתה מפילה אותם.
     הסדר מתאפס בכל כותרת ספק. */
  let order=true,prev=-1;
  for(const tr of document.querySelectorAll('#tbl tbody tr')){
   if(tr.classList.contains('grp')){prev=-1;continue}
   if(!tr.hasAttribute('data-i'))continue;
   const pn=((tr.querySelector('.obj')||{}).textContent||'').trim();
   const it=ALL.find(a=>a.pn===pn);if(!it)continue;
   const rk=urgRank(it);
   if(rk<prev)order=false;
   prev=rk}
  return {rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
    groups:gr.length,sizes,desc,order,uniqCats:[...new Set(cats)].length,
    hasUnits:gr.every(t=>/יח׳ להזמנה/.test(t.textContent)),
    hasLT:gr.some(t=>/ימ׳ אספקה/.test(t.textContent))}});
 ok('אותו מספר פריטים בשני הקיבוצים',bySup.rows===byDec.rows,`${byDec.rows} → ${bySup.rows}`);
 ok('הפריטים מקובצים לפי ספק',bySup.groups>0,bySup.groups+' ספקים');
 ok('ההזמנה הגדולה ראשונה',bySup.desc,bySup.sizes.join(' · '));
 ok('כותרת הספק נושאת כמות לפי הזמנה',bySup.hasUnits);
 ok('כותרת הספק נושאת זמן אספקה',bySup.hasLT);
 ok('ההחלטה עוברת לשורה כשהכותרת נושאת את הספק',bySup.uniqCats>1,bySup.uniqCats+' סיווגים');
 ok('בתוך ספק נשמר סדר הדחיפות',bySup.order);
 await p.click('#tbl tr.grp.sup');await p.waitForTimeout(500);
 const afterColl=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
 ok('קיפול ספק מקפל את ההזמנה כולה',afterColl<bySup.rows,`${bySup.rows} → ${afterColl}`);
 await p.click('#tbl tr.grp.sup');await p.waitForTimeout(400);
 await p.click('#gseg button[data-gb=decision]');await p.waitForTimeout(600);
 ok('חזרה לקיבוץ לפי שכבה',
   /לקוח ממתין ואין מלאי/.test(await p.evaluate(()=>[...document.querySelectorAll('#tbl tr.grp b')].map(x=>x.textContent).join(' · '))));

 /* ============ רצפת מלאי הביטחון ============
    sugSS = max(ssStat, ssMin), ו-ssMin הוא העמודה "מל.בט.מינ." מה-SAP.
    נמדד: ב-2,629 מתוך 2,638 הפריטים שיש להם ssMin הוא זהה בדיוק
    לעמודת "מלאי בטחון" — כלומר הרצפה היא המצב הקיים, ולכן הנוסחה
    לעולם לא תמליץ להוריד. הפער בין הרצפה לדרישה הסטטיסטית הוא
    החלטה שעולה כסף, והוא חייב להיות גלוי. */
 const seen=id=>p.evaluate(i=>{const e=document.getElementById(i);
   return !!e&&e.offsetParent!==null&&e.getBoundingClientRect().width>0},id);
 /* ============ DOORS · תיקון להנחה שהייתה כאן ============
    כאן היה כתוב «נכנסים למסלול דרך הטאב, לא דרכם», ושהמסלולים
    נגישים «דרך setMode ודרך סינון בקטלוג». נמדד, ושתי הרגליים
    התגלו כחלשות:

      · TRACKS נושא חמישה מסלולים — line · short · m1 · cap · fix.
        אף אחד מ-cust/floor/trend/applied/done אינו ביניהם, ובתוכם
        סרגל הטאבים אף נמחק ומוסתר. אין טאב להיכנס דרכו.
      · setMode אינה דלת למשתמש, היא פונקציה.
      · סינון בקטלוג *כן* עובד: cust≥1 ו-free≤0 מחזיר בדיוק את
        אותם 81 פריטים. אבל זו דלת בלי שם, שתי עמודות עומק,
        ורק למי שכבר יודע מה הוא מחפש. המתכנן שאל «לאן זה נעלם».

    הכלל החדש מיישם את החלטת 14.9 במלואה ולא סותר אותה: «הספירה
    עוזרת לבחור לאן ללכת, הסכום עוזר להחליט מה לעשות בפנים» —
    ולכן מחוץ למסלול ספירה בלבד, ובתוכו ספירה וסכום. */
 /* אחרי המעבר למסילה האנכית יש שתי רמות ניווט: תחום ב-#tabs ודלת
    משנה ב-#railsub. שלוש העזר האלה הולכות באותו מסלול שהמעתד הולך
    בו — קודם התחום, ואז הדלת שבתוכו. */
 const navEl=`#tabs .tab,#railsub .railsub2`;
 const navTo=async k=>{await p.evaluate(([m,sel])=>{
   const area=(typeof AREA_OF!=='undefined'&&AREA_OF[m])||m;
   const a=[...document.querySelectorAll('#tabs .tab')].find(x=>x.dataset.m===area);
   if(a)a.click();
   const t=[...document.querySelectorAll(sel)].find(x=>x.dataset.m===m);
   if(t)t.click();else setMode(m)},[k,navEl])};
 const navSeen=k=>p.evaluate(([m,sel])=>{const t=[...document.querySelectorAll(sel)]
   .find(e=>e.dataset.m===m);return !!t&&t.getBoundingClientRect().width>0},[k,navEl]);
 const navBdg=k=>p.evaluate(([m,sel])=>{const t=[...document.querySelectorAll(sel)]
   .find(e=>e.dataset.m===m);
   return t?((t.querySelector('.bdg')||{}).textContent||'').trim():null},[k,navEl]);
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);
 /* היה: «נגיש מסרגל הטאבים». הסרגל צומצם לשניים ברפורמה, והמסלול
    נשאר נגיש דרך setMode. מה שחשוב לא השתנה: שהוא קיים ומציג פריטים. */
 ok('מסלול "רצפת SS" עדיין קיים ומציג פריטים',
    await p.evaluate(()=>floorRows().length>0),
    await p.evaluate(()=>floorRows().length+' פריטים'));
 /* היה: הכפתור נדרש להיות גלוי *מחוץ* למסלול, כדי שלא יהיה דלת מתה
    (PR #69). המסילה האנכית פתרה את אותה בעיה במקום טוב יותר: «רצפת
    SS» היא דלת משנה קבועה בתוך «תכנון ו-MRP», עם ספירה, ולכן שבב
    ניווט נוסף בשורת הכותרת הוא אותו ניווט בשני מקומות — והוא עלה
    שם 690px מתוך 1008 ב-1280px.
    הכלל החדש: מחוץ למסלול אין שבב בכותרת אבל יש דלת במסילה עם אותה
    ספירה בדיוק; בתוך המסלול השבב חוזר ונושא את הסיכום הכספי. */
 const flOut=await p.evaluate(()=>{const e=document.getElementById('floorBtn');
   return {seen:!!e&&e.offsetParent!==null&&e.getBoundingClientRect().width>0,
     inSub:(SUBNAV.month||[]).some(x=>x[0]==='floor'),
     railN:navCount('floor').n,n:floorRows().length}});
 ok('ב"היום" אין שבב רצפה בכותרת — הדלת יושבת במסילה עם אותה ספירה',
    flOut.n>0&&!flOut.seen&&flOut.inSub&&flOut.railN===flOut.n,
    `כותרת=${flOut.seen} · מסילה=${flOut.railN} · בפועל=${flOut.n}`);
 await p.evaluate(()=>setMode('month'));await p.waitForTimeout(700);
 const fl=await p.evaluate(()=>({vis:!document.getElementById('floorBtn').hidden,
   txt:document.getElementById('floorBtn').textContent.trim(),
   heads:[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim()),
   over:document.querySelectorAll('#tbl td.need.over').length,
   need:document.querySelectorAll('#tbl td.need').length,
   /* הסימון חייב להתאים לנתונים: לא "יש סימונים" אלא "בדיוק אלה
      שהפער שלהם חיובי" — בפיקסצ׳ר ייתכן שאין חפיפה כלל */
   expectOver:(()=>{const rows=currentRows().slice(0,300);
     return rows.filter(r=>ssFloorGap(r)>0).length})()}));
 ok('המסלול נגיש מ"החודש"',await p.evaluate(()=>{setMode('floor');return track==='floor'}));
 ok('עמודת "הדרישה" נוספה למסלול הפרמטרים',fl.heads.includes('הדרישה'),fl.heads.join(' · '));
 ok('הדרישה הסטטיסטית מוצגת גם כשהרצפה ניצחה',fl.need>0,fl.need+' שורות');
 ok('הסימון תואם בדיוק את הפריטים שהפער שלהם חיובי',
   fl.over===fl.expectOver,`מסומנים ${fl.over} · צפוי ${fl.expectOver}`);
 await navTo('floor');await p.waitForTimeout(800);
 ok('בתוך המסלול מופיע הסיכום הכספי',await seen('floorBtn'));
 const fv=await p.evaluate(()=>{
   const cells=[...(document.querySelector('#tbl tbody tr[data-i]')||{children:[]}).children].map(td=>td.textContent.trim());
   return {rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
     heads:[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim()).join(' · '),
     pgR:document.getElementById('pgR').textContent, cells,
     /* מיון לפי שווי הפער — יורד */
     sorted:(()=>{const v=[...document.querySelectorAll('#tbl tbody tr[data-i]')].slice(0,8)
       .map(tr=>+((tr.children[6]||{}).textContent||'0').replace(/[^\d.]/g,''));
       for(let i=1;i<v.length;i++)if(v[i]>v[i-1]+0.001)return false;return true})()}});
 ok('המסלול מציג פריטים',fv.rows>0,fv.rows+' פריטים');
 ok('עמודות ייעודיות',fv.heads.includes('שווי הפער'),fv.heads);
 ok('ממוין לפי שווי הפער',fv.sorted);
 ok('הפער חיובי בכל שורה',+fv.cells[5]>0,fv.cells.join(' | '));
 ok('הפוטר מסכם יחידות וכסף',/יח׳ מעל הדרישה/.test(fv.pgR),fv.pgR);
 /* ============ הדלת הורחבה מ-409 ל-2,287 ============
    התנאי rate>0 הוסר: הפער נמדד היום על כל פריט מנוהל־מלאי. שתי
    השאלות נשארות נפרדות *בתוך* הדלת — «הדרישה» מציגה «—» כשאין קצב
    מדוד, והסיכום מפצל את שתי האוכלוסיות. */
 const fw=await p.evaluate(()=>{
  const all=(QF&&QF.all?QF.all:[]);
  const rows=floorRows();
  return {n:rows.length,
   narrow:all.filter(r=>r.rate>0&&!r.isOD&&!r.noStock&&(r.ssMinEff||0)>(r.ssStat||0)).length,
   noRate:rows.filter(ssNoRate).length,
   /* אף החרגה עסקית אינה נכנסת */
   leaks:rows.filter(r=>r.isOD||r.noStock||(r.ssMinEff||0)<=(r.ssStat||0)).length,
   /* הפער תמיד מול הרצפה שבתוקף, לא מול ssMin הגולמי */
   rawFloor:rows.filter(r=>ssFloorGap(r)!==(r.ssMinEff||0)-(r.ssStat||0)).length,
   /* בשורות חסרות־קצב התא מציג «—» ולא 0 */
   /* התאמה שורה-לשורה בין «אין קצב» לבין התא שמציג «—» */
   dash:(()=>{const tr=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
     const shown=tr.map(t=>currentRows()[+t.dataset.i]);
     return {n:tr.length,
       want:shown.filter(ssNoRate).length,
       got:tr.filter(t=>(t.children[4]||{}).textContent.trim()==='—').length,
       match:tr.every((t,i)=>((t.children[4]||{}).textContent.trim()==='—')===ssNoRate(shown[i]))}})(),
   pgR:document.getElementById('pgR').textContent,
   sub:document.getElementById('phdSub').textContent}});
 ok('הדלת כוללת גם פריטים בלי קצב מדוד',fw.n>fw.narrow,
    `${fw.n} פריטים · ${fw.narrow} מהם עם קצב (לפני ההרחבה זה היה כל המסלול)`);
 ok('ואף החרגה עסקית לא נכנסה איתם',fw.leaks===0,fw.leaks+' דליפות');
 ok('הפער נמדד מול הרצפה שבתוקף ולא מול ssMin הגולמי',fw.rawFloor===0,fw.rawFloor+' חריגות');
 ok('בעמודת הדרישה «—» מופיעה בדיוק בשורות שאין בהן קצב מדוד',
    fw.dash.match&&fw.dash.got===fw.dash.want,
    `${fw.dash.got}/${fw.dash.want} מתוך ${fw.dash.n} שורות מרונדרות`);
 ok('והסיכום אומר כמה מהן בלי קצב',
    fw.noRate===0||/בלי קצב מדוד/.test(fw.pgR),fw.pgR);
 ok('והתת-כותרת מסבירה מה «—» אומרת',/אין קצב מדוד/.test(fw.sub),fw.sub.slice(0,140));
 await p.click('#floorBtn');await p.waitForTimeout(600);
 ok('לחיצה על הסיכום מחזירה ל"החודש"','month'===await p.evaluate(()=>mode));
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);

 /* ============ הדלת לפריטים שסומנו "טופל" ============
    סימון מוציא פריט מתור העבודה ברונג הראשון של classify. המנגנון
    והאחסון היו קיימים, אבל לא הייתה דרך להגיע אליהם כדי לבטל. */
 /* בדיקות קודמות בקובץ מסמנות פריטים — מתחילים מלוח נקי */
 await p.evaluate(()=>{Object.keys(MARKS).forEach(k=>delete MARKS[k]);saveMarks();apply()});
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(600);
 const dTab=()=>navSeen('done');
 /* היה: «בלי סימונים הטאב אינו קיים». הכלל הוחלף בהחלטת המתכנן —
    «היעדים אינם נעלמים כשהמונה אפס». דלת שנעלמת מסתירה מהמעתד
    שהיכולת קיימת, וזו בדיוק הסיבה שהמסלול הזה נבנה מלכתחילה. */
 ok('בלי סימונים דלת «טופלו» קיימת עם מונה אפס',
    (await dTab())&&'0'===(await navBdg('done')),'מונה '+(await navBdg('done')));
 const n0=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
 for(let k=0;k<3;k++){
  await p.evaluate(()=>document.querySelector('#tbl tbody tr[data-i] .dn[data-done]').click());
  await p.waitForTimeout(400)}
 ok('אחרי סימון הדלת עדיין שם',await dTab());
 ok('תג הדלת מציג את המספר הנכון','3'===await navBdg('done'),'מונה '+(await navBdg('done')));
 ok('הפריטים ירדו מתור העבודה',
   (await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length))===n0-3);
 await navTo('done');await p.waitForTimeout(600);
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
 ok('לחיצה על שבב הסיכום מחזירה ל"היום"','today'===await p.evaluate(()=>mode));
 ok('הפריט שבוטל חזר לתור העבודה',
   (await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length))===n0-2);
 await p.evaluate(()=>{Object.keys(MARKS).forEach(k=>delete MARKS[k]);saveMarks();apply()});
 await p.waitForTimeout(600);
 ok('ניקוי הסימונים מחזיר את המונה לאפס, והדלת נשארת',
    (await dTab())&&'0'===(await navBdg('done')),'מונה '+(await navBdg('done')));

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
  return {heads:h, hasStock:h.includes('מלאי'), hasCust:h.includes('דרישת לקוח'),
    hasOtw:h.some(x=>x.indexOf('בדרך')===0),
    otwMarked:!!otw&&/⌛/.test(otw.textContent),
    otwTitleHasEta:!!otw&&/אין תאריך הגעה/.test(otw.getAttribute('title')||''),
    otwDim:so&&sn?so.color!==sn.color:false,
    zeroRed:!!sz&&sz.fontWeight>=600}});
 ok('עמודת מלאי קיימת',dec.hasStock,dec.heads.join(' · '));
 /* «לקוח» נקרא עכשיו «דרישת לקוח», כדי להפריד עובדה מתחזית. */
 ok('עמודת דרישת הלקוח קיימת',dec.hasCust,dec.heads.join(' · '));
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
  const notes=[...document.querySelectorAll('#tbl tbody tr[data-i] .ocat')].map(x=>x.textContent.trim());
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

 /* ============ הגרף המצרפי הוסר ============
    כאן ישבו שבע טענות על «צריכה חודשית — כל הקטלוג»: שהוא נראה בלי
    גלילה, עמודה לכל חודש, יחס צירים, תווי סקאלה וכו׳. הגרף סיכם
    יחידות של חלפים שונים — בורג ומנוע באותה עמודה — והוסר בהחלטת
    המתכנן. מה שנבדק עכשיו הוא שהוא באמת ירד, **ושגרפי הפריט נשארו**. */
 await p.evaluate(()=>setMode('catalog'));await p.waitForTimeout(700);
 const ch=await p.evaluate(()=>({
   el:!!document.getElementById('ctop'),
   fn:typeof topChart!=='undefined',
   line:(()=>{const e=document.getElementById('line');
     return e&&e.offsetParent!==null?Math.round(e.getBoundingClientRect().height):0})(),
   link:(()=>{const e=document.getElementById('shortlink');
     return e&&e.offsetParent!==null?(e.innerText||'').replace(/\s+/g,' ').trim():''})()}));
 ok('הגרף המצרפי אינו קיים יותר בקטלוג',!ch.el&&!ch.fn,
    `אלמנט=${ch.el} · פונקציה=${ch.fn}`);
 /* ============ רצועת החוסרים אינה שייכת לקטלוג ============
    נמדד לפני: 86px של רצועת חוסרים בקטלוג. renderLine עשה
    el.hidden=true, אבל `.line{display:flex}` בספציפיות (0,1,0) ניצח
    את כלל ה-[hidden] של הדפדפן — כלומר ההסתרה מעולם לא עבדה. */
 ok('רצועת החוסרים אינה מוצגת בקטלוג',ch.line===0,ch.line+'px (היה 86)');
 ok('ובמקומה קישור קומפקטי לרשימת החוסרים',
    /חוסר/.test(ch.link)&&/לרשימת החוסרים/.test(ch.link),ch.link.slice(0,90));
 /* ============ היעד הוא מה שהמונה סופר ============
    הקישור הוביל קודם ל-setMode('line'), והמונה שלו סופר
    decisionList('short'). שתי אוכלוסיות שונות: נמדד מונה 300 מול יעד
    של 27 כש-LINE_SEL היה 'nopo'. היעד הוא היום mode='today', שבו
    currentRows() מחזיר בדיוק את decisionList('short').
    ההשוואה המלאה בין שתי האוכלוסיות נמצאת ב-tests/routes.test.js. */
 await p.click('#shortlinkGo');await p.waitForTimeout(700);
 ok('הקישור מוביל לרשימה שהמונה סופר',
    'today'===await p.evaluate(()=>mode),await p.evaluate(()=>mode));
 ok('ואותם מק״טים בדיוק',
    await p.evaluate(()=>{const a=decisionList('short').map(r=>r.pn).sort(),
      c=currentRows().map(r=>r.pn).sort();
      return a.length===c.length&&a.every((x,i)=>x===c[i])}));
 /* גרפי הפריט — לא נגעו */
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(500);
 ok('גרפי הצריכה ברמת הפריט נשארו',
    2===await p.evaluate(()=>['c1','c2'].filter(i=>!!document.getElementById(i)).length));
 await p.evaluate(()=>closeDetail());await p.waitForTimeout(250);
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);

 /* ============ המגירה דוחפת ולא מכסה — כשיש רוחב ============
    מעל 1900px הרשימה מצטמצמת ברוחב המגירה; מתחת לזה המגירה מכסה, כי
    עמודת התיאור הייתה יורדת מתחת ל-235px ומפסיקה להיות קריאה. הבדיקה
    שומרת על שני הצדדים: גם על הדחיפה וגם על אי-החזרת תקלה 18. */
 const dwAt=async w=>{
  await p.setViewportSize({width:w,height:900});await p.waitForTimeout(350);
  await p.evaluate(()=>closeDetail());await p.waitForTimeout(250);
  const shut=await p.evaluate(()=>{const t=document.getElementById('tbl');
    return {w:Math.round(t.getBoundingClientRect().width),
      over:t.scrollWidth>t.clientWidth+2}});
  await p.evaluate(()=>{const r=(QF.all||[])[0];if(r)detail(r)});
  await p.waitForTimeout(400);
  return p.evaluate(s0=>{
   const t=document.getElementById('tbl'),d=document.querySelector('.wdetail');
   const tb=t.getBoundingClientRect(),db=d.getBoundingClientRect();
   const rows=[...t.querySelectorAll('tbody tr[data-i]')].slice(0,20);
   const base=rows.length?Math.min(...rows.map(r=>r.getBoundingClientRect().height)):0;
   const ths=[...t.querySelectorAll('thead th')];
   const di=ths.findIndex(h=>/תיאור/.test(h.textContent));
   return {open:document.body.classList.contains('dopen'),
     push:document.body.classList.contains('push'),
     overlap:Math.round(Math.max(0,Math.min(tb.right,db.right)-Math.max(tb.left,db.left))),
     shrank:Math.round(tb.width)<s0.w-100,
     shutW:s0.w,openW:Math.round(tb.width),
     desc:di>=0&&rows.length?Math.round(rows[0].children[di].getBoundingClientRect().width):0,
     wrapped:rows.filter(r=>r.getBoundingClientRect().height>base*1.5).length,
     over:t.scrollWidth>t.clientWidth+2||s0.over}},shut)};
 const dwWide=await dwAt(1920);
 ok('ב-1920 המגירה דוחפת ואינה מכסה את הטבלה',
   dwWide.open&&dwWide.overlap<5&&dwWide.shrank,
   `חפיפה ${dwWide.overlap}px · ${dwWide.shutW}→${dwWide.openW}`);
 ok('הדחיפה אינה יוצרת גלישה או שבירת שורות',!dwWide.over&&dwWide.wrapped===0,
   `תיאור ${dwWide.desc}px · שורות שנשברו ${dwWide.wrapped}`);
 ok('עמודת התיאור נשארת קריאה אחרי הדחיפה',dwWide.desc>=200,dwWide.desc+'px');
 /* השבב היחיד שיורד בזמן הדחיפה הוא רמז המקלדת — 218px של טקסט סטטי.
    בלעדיו שורת הכותרת נשברה לשתי שורות ב-1920. */
 const dwHead=await p.evaluate(()=>{const e=document.querySelector('#w_queue .phd');
   const k=e.querySelector('.kbd');
   const kids=[...e.children].filter(c=>c.offsetParent!==null&&c.getBoundingClientRect().width>0);
   return {h:Math.round(e.getBoundingClientRect().height),
     kbd:k?k.offsetParent!==null:false,chips:kids.length,
     avail:Math.round(e.clientWidth),
     /* שבב שגבוה מ-32px נשבר בתוך עצמו — זה מה שהגביה את השורה */
     tall:kids.filter(c=>c.getBoundingClientRect().height>32)
       .map(c=>`${(c.id||c.className).toString().slice(0,10)} h${Math.round(c.getBoundingClientRect().height)} w${Math.round(c.getBoundingClientRect().width)}`),
}});
 ok('שורת הכותרת נשארת בשורה אחת גם כשהמגירה דוחפת',dwHead.h<=48,
   `${dwHead.h}px · ${dwHead.chips} שבבים ב-${dwHead.avail}px`+
   (dwHead.tall.length?` · נשברו: ${dwHead.tall.join(', ')}`:''));
 /* שש הדלתות הן ילד אחד (.doors) ולא שישה, ולכן מניין הילדים
     ירד. מה שנבדק לא השתנה: הרמז יורד, והשאר נשאר. */
 ok('רמז המקלדת יורד בזמן הדחיפה, שאר השבבים נשארים',
   !dwHead.kbd&&dwHead.chips>=3,dwHead.chips+' שבבים');
 /* *אם* היא דוחפת תלוי בנתונים — כמה רחב התיאור בדוח הזה — ולכן זו
    אינה תכונה שכדאי לקבע בבדיקה. מה שכן חייב להתקיים בכל מצב ובכל
    רוחב הוא האינווריאנטה: כשדוחפת אין חפיפה ואין גלישה, וכשלא דוחפת
    הרשימה לא זזה. */
 const inv=[];let pushed=0;
 for(const w of [1920,1745,1536,1280]){
  for(const m of ['today','month','catalog','cust','floor','trend']){
   await p.evaluate(k=>{closeDetail();setMode(k)},m);await p.waitForTimeout(320);
   const r=await dwAt(w);
   /* נבדק רק הכיוון שמסוכן: דחיפה שמסתירה נתונים. הכיוון ההפוך
      («לא דוחפת ⇒ הרשימה לא זזה») נמדד כאן דרך מחלקת ה-body, והמדידה
      הזאת התבררה כלא יציבה בין מעברי מצב — ולכן היא לא נקבעת כאן
      כאילו היא אמת. */
   const bad=r.push&&(r.overlap>5||r.over||!r.shrank||r.desc<200);
   if(r.push)pushed++;
   if(bad)inv.push(`${w}/${m} ${r.push?'דוחפת':'מכסה'} ${r.shutW}→${r.openW} חפיפה${r.overlap} תיאור${r.desc}${r.over?' גלישה':''}`);}}
 ok('בכל מצב ובכל רוחב: כשדוחפת — אין חפיפה, אין גלישה, והתיאור קריא',
   inv.length===0,inv.join(' · ')||`${pushed} דוחפות מתוך 24 צירופים`);
 await p.evaluate(()=>{closeDetail();setMode('today')});await p.waitForTimeout(400);
 const dwMid=await dwAt(1512);
 ok('ב-1512 המגירה מכסה — הרשימה לא מצטמצמת (תקלה 18 לא חוזרת)',
   dwMid.open&&dwMid.overlap>100&&!dwMid.shrank,
   `חפיפה ${dwMid.overlap}px · ${dwMid.shutW}→${dwMid.openW}`);
 ok('גם במצב המכסה אין גלישה או שבירת שורות',!dwMid.over&&dwMid.wrapped===0);
 await p.setViewportSize({width:1512,height:860});
 await p.evaluate(()=>closeDetail());await p.waitForTimeout(350);
 const dwBack=await p.evaluate(()=>{const t=document.getElementById('tbl');
   return {w:Math.round(t.getBoundingClientRect().width),
     open:document.body.classList.contains('dopen')}});
 ok('סגירת המגירה מחזירה את הרוחב',!dwBack.open&&dwBack.w>=dwMid.shutW-2,dwBack.w+'px');

 /* ============ יישור עמודה מול הכותרת שלה ============
    text-align:end ב-RTL הוא שמאל, והכותרות יושבות ב-start (ימין) — כך
    שכל ערך מספרי ישב בקצה הנגדי של התא מהכותרת שלו, עד 90px. נמדדת
    קופסת הטקסט בפועל (Range) ולא קופסת התא, כי padding שונה בין th ל-td
    הוא בדיוק מה שגורם לסטייה. */
 const alignScan=()=>p.evaluate(()=>{
   const t=document.querySelector('#tbl');if(!t)return [];
   const ths=[...t.querySelectorAll('thead th')];
   const row=[...t.querySelectorAll('tbody tr')]
     .find(r=>!r.classList.contains('grp')&&r.children.length===ths.length);
   if(!row)return [];
   const inner=el=>{const rg=document.createRange();rg.selectNodeContents(el);
     const b=rg.getBoundingClientRect();return b.width?b.right:null};
   return ths.map((th,i)=>{const td=row.children[i];
     const hi=inner(th.querySelector('.thc')||th),ci=inner(td);
     if(hi==null||ci==null)return null;
     return {t:th.textContent.trim().slice(0,12),
       d:Math.round((td.getBoundingClientRect().right-ci)-(th.getBoundingClientRect().right-hi))}})
     .filter(Boolean)});
 const TOL=6;
 for(const w of [1512,1180]){
  await p.setViewportSize({width:w,height:860});await p.waitForTimeout(300);
  for(const m of ['today','month','catalog','cust','floor','trend']){
   await p.evaluate(k=>setMode(k),m);await p.waitForTimeout(450);
   const cols=await alignScan();
   const bad=cols.filter(c=>Math.abs(c.d)>TOL);
   ok(`כל עמודה יושבת תחת הכותרת שלה · ${m} · ${w}px`,cols.length>0&&bad.length===0,
     bad.length?bad.map(c=>`${c.t}=${c.d}px`).join(' · '):cols.length+' עמודות');}}
 await p.setViewportSize({width:1512,height:860});await p.waitForTimeout(300);
 /* המספרים מיושרים לספרת האחדות — יישור לקצה הנגדי משאיר אותה מרופטת */
 const ones=await p.evaluate(()=>{
   const t=document.querySelector('#tbl');
   const idx=[...t.querySelectorAll('thead th')].findIndex(h=>/צריכה|שווי|פער/.test(h.textContent));
   if(idx<0)return null;
   const xs=[...t.querySelectorAll('tbody tr')].filter(r=>!r.classList.contains('grp'))
     .map(r=>r.children[idx]).filter(Boolean).slice(0,12)
     .map(td=>{const rg=document.createRange();rg.selectNodeContents(td);
       const b=rg.getBoundingClientRect();return b.width?Math.round(b.right):null}).filter(v=>v!=null);
   return {n:xs.length,spread:xs.length?Math.max(...xs)-Math.min(...xs):0}});
 ok('בעמודה מספרית ספרות האחדות מיושרות',!ones||ones.spread<=2,
   ones?`${ones.n} ערכים · פיזור ${ones.spread}px`:'—');
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(300);

 /* ============ מסלול ההון הכלוא — בלוק לכל מטבע ============
    הדירוג הכספי אינו משווה מטבעות זה לזה, ולכן הרשימה מסודרת בבלוקים.
    שתי הסכנות: שהמיון יתהפך (הגדול ביותר בסוף), ושמגבלת הרינדור
    תדחק בלוק שלם אל מחוץ למסך. */
 await p.evaluate(()=>{cur='today';track='cap';DL_CACHE={};apply()});await p.waitForTimeout(500);
 const cap=await p.evaluate(()=>{
   const trs=[...document.querySelectorAll('#tbl tbody tr')];
   const rows=currentRows();
   const seq=[];let last=null;
   for(const r of rows){const c=curKey(r);if(c!==last){seq.push(c);last=c}}
   const curs=[...new Set(rows.map(curKey))];
   /* view הוא מקומי ל-renderDecisions; המטבעות שהוצגו בפועל נקראים
      מכותרות הקבוצה, ומאומתים בכך שאחרי כל כותרת יש לפחות שורת נתונים. */
   const shown=[];
   for(const g of trs.filter(t=>t.classList.contains('grp')&&/^cur:/.test(t.dataset.g||''))){
     let nx=g.nextElementSibling,has=false;
     while(nx&&!nx.classList.contains('grp')){if(nx.dataset.i!=null){has=true;break}nx=nx.nextElementSibling}
     if(has)shown.push(g.dataset.g.slice(4))}
   const first=rows[0],biggest=rows.reduce((m,r)=>r.expCap>m.expCap?r:m,rows[0]);
   return {n:rows.length,seq,curs,shown,
     firstIsBiggestOfItsBlock:first.expCap===Math.max(...rows.filter(r=>curKey(r)===curKey(first)).map(r=>r.expCap)),
     firstCap:first.expCap,maxCap:biggest.expCap,
     groups:trs.filter(t=>t.classList.contains('grp')).length,
     cuts:trs.filter(t=>t.classList.contains('cutrow')).length}});
 ok('כל מטבע יושב בבלוק רציף אחד',cap.seq.length===cap.curs.length,
   `סדר הבלוקים: ${cap.seq.join(' → ')} · מטבעות: ${cap.curs.length}`);
 ok('הפריט הראשון הוא הגדול בבלוק שלו ולא הקטן',cap.firstIsBiggestOfItsBlock,
   `ראשון ${cap.firstCap.toLocaleString('he-IL')} · מקסימום ${cap.maxCap.toLocaleString('he-IL')}`);
 ok('כל מטבע מקבל כותרת קבוצה משלו',cap.groups===cap.curs.length,
   `${cap.groups} כותרות · ${cap.curs.length} מטבעות`);
 ok('קו ה-80% נמתח לכל מטבע בנפרד',cap.cuts===cap.curs.length,
   `${cap.cuts} קווים · ${cap.curs.length} מטבעות`);
 ok('אף מטבע לא נדחק אל מחוץ לרשימה המרונדרת',cap.shown.length===cap.curs.length,
   `מוצגים ${cap.shown.join(', ')} · קיימים ${cap.curs.join(', ')}`);
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(300);

 /* ============ DOORS · המסכים שלא הייתה אליהם דלת ============
    cust · floor · trend · applied · done אינם בסרגל הטאבים (TRACKS
    נושא line/short/m1/cap/fix), והכפתור שלהם היה `hidden` אלא אם
    כבר היית בתוכם — כלומר הדרך היחידה פנימה הייתה להקליד את הביטוי
    המדויק בשאילתה החופשית. הבדיקה נועלת את שלושת הכללים. */
 const doorVis=()=>['custBtn','floorBtn','trendBtn','apBtn','doneBtn']
   .filter(id=>{const e=document.getElementById(id);return !!e&&e.offsetParent!==null});
 const doorTxt=id=>{const e=document.getElementById(id);
   return e&&e.offsetParent!==null?(e.textContent||'').replace(/\s+/g,' ').trim():''};
 const doorN=()=>({cust:custRows().length,floor:floorRows().length,trend:trendRows().length,
   done:(QF&&QF.marked)?QF.marked.filter(x=>x.mark&&x.mark.t==='handled').length:0});

 await p.evaluate(()=>setMode('line'));await p.waitForTimeout(500);
 const dOut=await p.evaluate(()=>({vis:['custBtn','floorBtn','trendBtn','apBtn','doneBtn']
    .filter(id=>{const e=document.getElementById(id);return !!e&&e.offsetParent!==null}),
   cust:(e=>e&&e.offsetParent!==null?(e.textContent||'').replace(/\s+/g,' ').trim():'')(document.getElementById('custBtn')),
   n:{cust:custRows().length,floor:floorRows().length,trend:trendRows().length}}));
 /* היה: «הכפתור גלוי מחוץ למסלול — זו הדלת». הדלת עברה למסילה,
    ולכן מחוץ למסלול שורת הכותרת נקייה משבבי ניווט לגמרי. מה שנועל
    שלא איבדנו את היכולת: לכל אחת מחמש הדלתות יש כניסה ב-SUBNAV,
    והספירה שם היא אותה ספירה בדיוק. */
 const dRail=await p.evaluate(()=>{const all=[].concat(...Object.values(SUBNAV)).map(x=>x[0]);
   return {all,counts:{cust:navCount('cust').n,floor:navCount('floor').n,
     trend:navCount('trend').n,applied:navCount('applied').n,done:navCount('done').n}}});
 ok('מחוץ למסלול אין אף שבב ניווט בשורת הכותרת',!dOut.vis.length,
    `גלויים: ${dOut.vis.join(', ')||'אין'}`);
 ok('וחמש הדלתות נגישות מהמסילה עם אותה ספירה',
    ['cust','floor','trend','applied','done'].every(k=>dRail.all.includes(k))
      &&dRail.counts.cust===dOut.n.cust&&dRail.counts.floor===dOut.n.floor
      &&dRail.counts.trend===dOut.n.trend,
    `לקוח ${dRail.counts.cust}/${dOut.n.cust} · רצפה ${dRail.counts.floor}/${dOut.n.floor} · מגמה ${dRail.counts.trend}/${dOut.n.trend}`);
 /* הכיתוב המלא מוסתר מחוץ למסלול, וזה מה שמאפשר לשש הדלתות לשבת
    בשורה אחת. בלי זה הן דרסו את קיבוץ הקבוצות ואת שבב ההשוואה. */
 /* הכיתובים חוזרים כברירת מחדל; fitDoors מקפל רק אם השורה נשברת.
    מה שנעול הוא שני הקצוות: כשיש מקום — שם מלא ולא סמל בלבד;
    ובכל מקרה — שורה אחת. */
 const dLbl=await p.evaluate(()=>{
   const phd=document.querySelector('#w_queue .phd');
   const vis=[...document.querySelectorAll('.doors .btn')].filter(b=>!b.hidden);
   const onScreen=[...document.querySelectorAll('.doors .btn')]
     .filter(b=>b.offsetParent!==null&&b.getBoundingClientRect().width>0)
     .map(b=>({id:b.id,on:b.classList.contains('on')}));
   const shown=[...document.querySelectorAll('.doors .btn')]
     .filter(b=>b.offsetParent!==null)
     .filter(b=>{const d=b.querySelector('.dl');return d&&d.offsetParent!==null});
   return {n:vis.length,shown:shown.length,onScreen,
     compact:document.getElementById('doors').classList.contains('compact'),
     names:shown.map(b=>b.textContent.replace(/\s+/g,' ').trim()),
     phdH:Math.round(phd.getBoundingClientRect().height)}});
 /* היה: «הדלתות נושאות שם ולא סמל בלבד» — כלל שנולד כשכל שש ישבו
    בשורה. עכשיו יש לכל היותר שבב אחד, של המסלול הנוכחי, ולכן אין
    לחץ על השורה והשם תמיד מלא. */
 ok('לכל היותר שבב אחד בכותרת, והוא של המסלול הנוכחי',
    dLbl.onScreen.length<=1&&dLbl.onScreen.every(b=>b.on),
    dLbl.onScreen.map(b=>b.id).join(' · ')||'אין');
 ok('והשבב נושא שם מלא, לא סמל בלבד',
    !dLbl.onScreen.length||dLbl.shown===dLbl.onScreen.length,
    dLbl.names.join(' · ')||'(אין שבב)');
 ok('ושורת הכותרת נשארת בשורה אחת',dLbl.phdH<=48,dLbl.phdH+'px');
 /* ============ UX_PASS · מה שהמתכנן הכריע ============ */
 const ux=await p.evaluate(()=>{
  const band=document.querySelector('.lband'),rows=document.querySelector('.rows');
  const trs=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
  const rh=trs.length?trs[0].getBoundingClientRect().height:0;
  const flags=[...document.querySelectorAll('.doors .btn,#bulkBtn')]
    .filter(e=>e.offsetParent!==null&&/⚑/.test(e.textContent));
  return {bandH:Math.round(band?band.getBoundingClientRect().height:0),
    mini:!!band&&band.classList.contains('mini'),
    hasToggle:!!document.getElementById('bandx'),
    rowH:Math.round(rh),
    listPct:Math.round((rows?rows.getBoundingClientRect().height:0)/innerHeight*100),
    rowsTop:Math.round(rows?rows.getBoundingClientRect().top:0),vh:innerHeight,
    above:[...document.querySelectorAll('.top,.lband,.wkpi,.wfilters,#w_queue .phd')]
      .filter(e=>e.offsetParent!==null)
      .map(e=>`${(e.id||e.className).toString().split(' ')[0]}:${Math.round(e.getBoundingClientRect().height)}`),
    flagUses:flags.length,
    nopxVisible:(e=>!!e&&e.offsetParent!==null)(document.querySelector('.nopx')),
    tabs:[...document.querySelectorAll('#tabs .tab')].map(t=>t.textContent.replace(/\s+/g,' ').trim()),
    tabNames:[...document.querySelectorAll('#tabs .tab .t')].map(t=>t.textContent.trim())}});
 ok('הרצועה מקופלת בפתיחה, עם כפתור לפתוח',
    ux.mini&&ux.hasToggle&&ux.bandH<90,`${ux.bandH}px · מקופלת ${ux.mini}`);
 /* הסייג «כל סכום כאן הוא רצפה» אינו מתקפל עם הרצועה. */
 ok('מונה «בלי מחיר» נשאר גלוי גם ברצועה מקופלת',ux.nopxVisible);
 ok('הרשימה מקבלת יותר מ-65% מהמסך',ux.listPct>=65,
    `${ux.listPct}% · ראש הרשימה ${ux.rowsTop}px מתוך ${ux.vh} · ${ux.above.join(' · ')}`);
 /* הדגל ⚑ סימן גם סימון קבוצתי וגם מסלול פרמטרים, שניהם על המסך
    בו-זמנית. עכשיו הוא שייך לסימון בלבד. */
 ok('הדגל ⚑ משמש למשמעות אחת בלבד',ux.flagUses<=1,ux.flagUses+' שימושים גלויים');
 /* היה: «המסך הראשי נקרא לטיפול היום» (PR #68, במקום «ציר הזמן»).
    המפרט החדש קבע שישה שמות תחום, והורה במפורש להימנע מהשמות
    החופפים «היום» / «לטיפול היום» / «החודש» — כולם נקראים כחלון
    זמן ולא כתחום עבודה, וזו הייתה אותה תקלה בדיוק בשם אחר. */
 const AREAN=['מרכז עבודה','חוסרים ולקוחות','רכש ואספקות','תכנון ו-MRP',
              'בריאות המלאי','קטלוג פריטים'];
 ok('המסילה נושאת את שישה שמות התחומים שנקבעו',
    ux.tabNames.length===6&&AREAN.every((n,i)=>ux.tabNames[i]===n),
    ux.tabNames.join(' | '));
 ok('ואין בשמות שם שנקרא כחלון זמן',
    !ux.tabNames.some(t=>/לטיפול היום|ציר הזמן|היום|החודש/.test(t)),
    ux.tabNames.join(' | '));
 /* ============ אזור החלטה אחד ============
    קודם נבדק כאן ש-`.oact` קיים בראש הכרטיס ושהמקטע «מה לעשות» בא
    לפני «מה מצדיק את זה». שניהם תיארו את המבנה שאוחד: אותה המלצה
    הופיעה ב-.oact, ב-.dcs וב«מה לעשות» — שלושה אזורים, מידע אחד. */
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(400);
 const card=await p.evaluate(()=>{
  const d=document.getElementById('detail');
  const r=CURRENT_DETAIL,a0=((r.act||[])[0]||'');
  const txt=(d.innerText||'');
  const dcs=d.querySelector('.dcs');
  return {oact:!!d.querySelector('.oact'),
    repeats:a0?txt.split(a0).length-1:0,
    acts:(r.act||[]).length,
    inDecision:!!dcs&&(dcs.innerText||'').indexOf(a0)>=0,
    hasQty:!!d.querySelector('.dcq'),
    qtyLabel:((d.querySelector('.dcql')||{}).textContent||''),
    hasWhy:!!d.querySelector('.dcwhy .msg'),
    whyN:d.querySelectorAll('.dcwhy .msg').length,
    whySrc:(r.why||[]).length,
    hasBtn:!!d.querySelector('.dcs [data-dmark]'),
    moreN:d.querySelectorAll('.dcmore li').length,
    secs:[...d.querySelectorAll('.sec h3')].map(h=>h.textContent.trim())}});
 ok('אותה המלצה אינה חוזרת בשלושה אזורים',card.repeats===1&&!card.oact,
    `${card.repeats} הופעות (היה 3) · .oact=${card.oact}`);
 ok('ואזור ההחלטה נושא את ארבעת החלקים',
    card.inDecision&&card.hasQty&&card.hasWhy&&card.hasBtn,
    `פעולה=${card.inDecision} כמות=${card.hasQty} סיבה=${card.hasWhy} כפתור=${card.hasBtn}`);
 ok('הכמות נושאת תיאור של משמעותה',
    card.qtyLabel.trim().length>0||/ללא חשיפה/.test(card.secs.join('')),
    `"${card.qtyLabel}"`);
 ok('כל הסיבות נשמרו — לא נחתכו',card.whyN===card.whySrc,
    `${card.whyN} מתוך ${card.whySrc}`);
 ok('פעולות נוספות נשמרות מתחת לראשית',card.moreN===Math.max(0,card.acts-1),
    `${card.moreN} נוספות · ${card.acts} בסך הכול`);

 /* ============ STORE_LOUD · אחסון שנגמר נשמע ============
    כל כתיבה לאחסון הייתה ב-try{}catch(_){} שבולע. כשהמכסה נגמרת
    הכלי נראה תקין והפנקס פשוט מפסיק להירשם. הבדיקה ממלאת את
    המכסה באמת — לא מדמה — ובודקת ששלושת הדברים קורים: הכתיבה
    נכשלת, השבב מופיע, והוא אומר מה לא נשמר. */
 const store=await p.evaluate(()=>{
  const before=document.getElementById('storechip').hidden;
  /* ממלאים את המכסה עד הסוף ממש: גושים גדולים ואז קטנים, כי
     המכסה אינה מדויקת לבית וכתיבה זעירה עוד נכנסת אחרי שהגדולות
     נכשלו. בלי זה saveMarks על אובייקט ריק פשוט מצליח. */
  let n=0;
  try{for(;n<200;n++)localStorage.setItem('__fill'+n,'x'.repeat(1024*128))}catch(_){}
  try{for(;n<4000;n++)localStorage.setItem('__fill'+n,'x'.repeat(1024))}catch(_){}
  /* המפתח כבר קיים, ודריסה בערך שווה-גודל אינה צורכת מכסה חדשה.
     מה שנכשל הוא *גידול*, ולכן הסימונים מוגדלים לפני השמירה —
     בדיוק כמו במציאות, שבה הפנקס והסימונים רק גדלים. */
  for(let i=0;i<400;i++)MARKS['בדיקה|בדיקה|PN'+i]={t:'handled',ts:Date.now(),note:'מילוי'};
  saveMarks();                      /* חייב להיכשל עכשיו */
  const el=document.getElementById('storechip');
  const shown=!el.hidden, txt=(el.textContent||'').trim(), tip=el.title||'';
  /* מפנים ובודקים שהשבב יורד */
  for(let i=0;i<n;i++)localStorage.removeItem('__fill'+i);
  saveMarks();
  const cleared=document.getElementById('storechip').hidden;
  return {before,filled:n,shown,txt,tip,cleared,
    failKey:STORE_FAIL?STORE_FAIL.key:null}});
 ok('בהתחלה אין שבב אחסון',store.before);
 ok('המכסה באמת נגמרה בבדיקה',store.filled>0&&store.filled<200,store.filled+' גושים של 128KB');
 ok('כתיבה שנכשלה מדליקה שבב',store.shown&&/לא נשמר/.test(store.txt),store.txt);
 ok('השבב אומר מה בדיוק לא נשמר',/הסימונים/.test(store.txt),store.txt);
 ok('וההסבר אומר שהנתונים לא ישרדו רענון',
    /לא יישרדו רענון/.test(store.tip)&&/מלא/.test(store.tip),store.tip.slice(0,80));
 ok('כשהמקום מתפנה השבב יורד',store.cleared&&!store.failKey,
    'שבב מוסתר '+store.cleared);
 /* שבב הסיכום שותק כשאין מה לומר — גם בתוך המסלול. נבדק על
    התכונה hidden עצמה, כי בפריסה החדשה שבב מחוץ למסלול ממילא אינו
    מוצג, והבדיקה הקודמת הייתה נכונה מסיבה אחרת. */
 ok('שבב בלי תוכן נשאר מוסתר — שותק כשאין מה לומר',
    await p.evaluate(()=>document.getElementById('doneBtn').hidden
      &&document.getElementById('apBtn').hidden));

 await navTo('cust');await p.waitForTimeout(900);
 const dIn=await p.evaluate(()=>({mode,track,
   txt:(document.getElementById('custBtn').textContent||'').replace(/\s+/g,' ').trim(),
   on:document.getElementById('custBtn').classList.contains('on'),
   rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
   n:custRows().length,
   gapOK:custRows().every(r=>r.cust>0&&(r.cust-(Math.max(0,r.free)+Math.max(0,r.po)))>0),
   withShelf:custRows().filter(r=>r.free>0).length,
   withPo:custRows().filter(r=>r.po>0).length,
   burnIn:custRows().filter(r=>r.burn).length,
   allBurn:QF.waiting.every(x=>custRows().some(r=>r.pn===x.pn))}));
 ok('הכניסה מהמסילה נכנסת למסלול',dIn.mode==='cust'&&dIn.track==='cust'&&dIn.on,
    `mode=${dIn.mode} track=${dIn.track}`);
 ok('בתוך המסלול השבב מוסיף את הסכום',/[$€₪]/.test(dIn.txt),dIn.txt);
 ok('ובתוכו הכיתוב המלא חוזר',/לקוח ממתין ללא כיסוי/.test(dIn.txt),dIn.txt);
 ok('המסלול מציג בדיוק את הפריטים שנספרו',dIn.rows===dIn.n,`${dIn.rows} שורות · ${dIn.n} נספרו`);
 /* ============ הכלל השתנה: פער כיסוי, לא מדף ריק ============
    קודם היה «הזמנה מול מדף אפס». נמדד: 50 פריטים, ומתוכם אחד בוער —
    פריט עם 19 הזמנות ויחידה אחת על המדף לא נכלל, כי המדף לא היה ריק
    לגמרי. היום הכלל הוא שהמלאי הפנוי והרכש הפתוח יחד אינם מכסים. */
 ok('הכלל הוא פער כיסוי מול הזמנת הלקוח',dIn.gapOK,
    `${dIn.n} פריטים · ${dIn.withShelf} מהם עם יחידות על המדף · ${dIn.withPo} עם רכש פתוח`);
 ok('פריט עם שארית מלאי אינו נופל יותר מהרשימה',dIn.withShelf>0,
    dIn.withShelf+' פריטים עם מדף גדול מאפס');
 ok('וכל הבוערים בפנים',dIn.allBurn&&dIn.burnIn>0,
    `${dIn.burnIn} בוערים מתוך ${dIn.n}`);

 /* השבב בתוך המסלול הוא גם דרך היציאה — וזו הדרך שבה נבדק שהיציאה
    מחזירה בדיוק לאן שהיית. */
 await p.click('#custBtn');await p.waitForTimeout(900);
 /* הכניסה היא דו-שלבית: התחום «חוסרים ולקוחות» ואז הדלת שבתוכו,
    ולכן המקום שחוזרים אליו הוא התחום — לא המסלול שהיית בו לפניו.
    מה שנעול כאן הוא שהיציאה חוזרת למקום שממנו נכנסת בפועל. */
 ok('לחיצה על השבב חוזרת לתחום שממנו נכנסת',
    (await p.evaluate(()=>mode))==='today',await p.evaluate(()=>mode));

 await p.evaluate(()=>setMode('catalog'));await p.waitForTimeout(700);
 await navTo('floor');await p.waitForTimeout(900);
 const backFrom=await p.evaluate(()=>mode);
 await p.click('#floorBtn');await p.waitForTimeout(900);
 /* שתי יציאות, שני יעדים שונים (today ו-month) — זה מה שמוכיח
    שהיציאה אינה מסך קבוע אלא המקום שממנו נכנסת. */
 ok('היציאה מחזירה לאן שהיית ולא למסך קבוע',
    backFrom==='floor'&&(await p.evaluate(()=>mode))==='month',
    `catalog → תכנון → ${backFrom} → ${await p.evaluate(()=>mode)}`);

 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(500);
 const dBud=await p.evaluate(()=>{const r=document.querySelector('.rows');
   return {chrome:r?Math.round(r.getBoundingClientRect().top):0,
     vis:document.querySelectorAll('#tbl tbody tr[data-i]').length}});
 ok('הכותרת לא שברה את תקציב הגובה',dBud.chrome>0&&dBud.chrome<360,dBud.chrome+'px');

/* ============ DOOR_ETA · ה-ETA בכל ארבע הדלתות ============
    תא אחד משותף, ארבע דלתות. נבדק כאן ולא ב-eta.test.js, כי שם
    הפיקסצ'ר הוא 10 פריטים ואין בו אף שורת רצפה, מגמה או עלייה —
    שלוש מהבדיקות היו עוברות על אפס שורות ולא אומרות דבר. */
 await p.setInputFiles('#fe',SD+'/zmrp-demo-eta.xlsx');await p.waitForTimeout(1400);
 const doors=[];
 for(const m of ['cust','floor','trend','rise']){
  await p.evaluate(k=>setMode(k),m);await p.waitForTimeout(800);
  doors.push(await p.evaluate(k=>{
   const rows=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
   const heads=[...document.querySelectorAll('#tbl thead th')];
   const w=document.querySelector('.rows');
   const model=k==='cust'?custRows():k==='floor'?floorRows():k==='trend'?trendRows():riseRows();
   return {k,heads:heads.length,cells:rows.length?rows[0].querySelectorAll('td').length:0,
     hasEtaHead:heads.some(h=>/הגעה/.test(h.textContent)),
     etaCells:document.querySelectorAll('#tbl td.sched, #tbl td.otw').length,
     withPo:model.filter(r=>(r.po||0)>0).length,
     dated:model.filter(r=>(r.etaQty||0)>0).length,
     n:rows.length,model:model.length,
     ovf:w?w.scrollWidth-w.clientWidth:0,
     pgR:(document.getElementById('pgR').textContent||'').trim()}},m))}
 const D=Object.fromEntries(doors.map(d=>[d.k,d]));
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);

 ok('כל ארבע הדלתות פופולטיות בפיקסצ׳ר',
    doors.every(d=>d.model>0),doors.map(d=>d.k+':'+d.model).join(' · '));
 ok('לכל ארבע הדלתות יש עמודת הגעה',
    doors.every(d=>d.hasEtaHead),doors.filter(d=>!d.hasEtaHead).map(d=>d.k).join(',')||'כולן');
 ok('הכותרות תואמות למספר התאים בכל דלת',
    doors.every(d=>d.heads===d.cells),
    doors.map(d=>`${d.k} ${d.heads}/${d.cells}`).join(' · '));
 ok('אין גלישה אופקית באף דלת',doors.every(d=>d.ovf<=1),
    doors.map(d=>d.k+' '+d.ovf+'px').join(' · '));
 /* הכלל שהמתכנן קבע: שום דבר לא יורד מהרשימה בגלל משלוח מכסה. */
 ok('שום דלת לא מסננת פריטים בגלל ה-ETA',
    doors.every(d=>d.n===d.model),
    doors.map(d=>`${d.k} ${d.n}/${d.model}`).join(' · '));
 ok('תא ההגעה מופיע לכל פריט עם רכש פתוח',
    doors.every(d=>d.etaCells>=d.withPo),
    doors.map(d=>`${d.k} ${d.etaCells}≥${d.withPo}`).join(' · '));
 ok('בירידה — הפוטר אומר מה כבר בדרך לתוך הירידה',
    D.trend.dated>0&&/כבר בדרך לתוך הביקוש היורד/.test(D.trend.pgR),
    `${D.trend.dated} עם תאריך · ${D.trend.pgR.slice(-60)}`);
 ok('בעלייה — הפוטר אומר כמה נסגרות במשלוח שכבר בדרך',
    D.rise.dated>0&&/נסגרות על ידי משלוח שכבר בדרך/.test(D.rise.pgR),
    `${D.rise.dated} עם תאריך · ${D.rise.pgR.slice(-60)}`);
 ok('לקוח ממתין — הפוטר מפריד בין «עם תאריך» ל«בלי»',
    /עם תאריך/.test(D.cust.pgR)&&/בלי תאריך/.test(D.cust.pgR),D.cust.pgR.slice(-60));

 /* ============ RISE · המסלול העולה על המסך ============ */
 await p.evaluate(()=>setMode('rise'));await p.waitForTimeout(700);
 const up=await p.evaluate(()=>{
  const rows=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
  const heads=[...document.querySelectorAll('#tbl thead th')];
  const w=document.querySelector('.rows');
  return {mode,track,n:rows.length,model:riseRows().length,
    heads:heads.length,cells:rows.length?rows[0].querySelectorAll('td').length:0,
    btnTxt:(document.getElementById('riseBtn').textContent||'').replace(/\s+/g,' ').trim(),
    on:document.getElementById('riseBtn').classList.contains('on'),
    trackBar:(e=>e?getComputedStyle(e).display:'')(document.getElementById('track')),
    overflow:w?w.scrollWidth-w.clientWidth:0,
    /* הסימון ⊤ מופיע בדיוק על השורות שנחתכו בתקרה, ולא על אחרות */
    capMarks:document.querySelectorAll('#tbl td.tup .cap').length,
    capRows:riseRows().filter(r=>r.riseF>=1.4).length,
    arrowUp:rows.length?/↑/.test(rows[0].querySelector('td.tup').textContent):false}});
 ok('מסלול העלייה מרנדר את כל השורות',up.n===up.model&&up.n>0,`${up.n} / ${up.model}`);
 ok('כותרות «בעלייה» תואמות למספר התאים',up.heads===up.cells,`${up.heads} / ${up.cells}`);
 ok('אין גלישה אופקית ב«בעלייה»',up.overflow<=1,up.overflow+'px');
 ok('כפתור «בעלייה» נראה ומסומן, והסכום נוסף רק בפנים',
    up.on&&/יח׳/.test(up.btnTxt)&&/[$€₪]/.test(up.btnTxt),up.btnTxt);
 ok('רצועת המסלולים מוסתרת',up.trackBar==='none',up.trackBar);
 ok('החץ מצביע למעלה',up.arrowUp);
 /* התקרה חייבת להיות נראית: «↑6000%» ליד קצב שהוגבר פי 1.4 בלבד
    נקרא כסתירה, ולכן השורה נושאת ⊤ שמסביר שההגברה נחתכה. */
 ok('כל שורה שנחתכה בתקרה נושאת את סימון ⊤ — ורק היא',
    up.capMarks===up.capRows&&up.capRows>0,`${up.capMarks} סימונים · ${up.capRows} נחתכו`);
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);

  ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 await p.screenshot({path:SD+'/04-after.png'});

 /* ============ עמודות הטבלה ============
    התקלה שהולידה את הרישום, נמדדה על main: בקיבוץ «לפי ספק» הטבלה
    הציגה 12 כותרות מול 11 תאים, וכל עמודה אחרי «תיאור» נשאה כותרת
    של עמודה אחרת — «מלאי 39» היה בפועל דרישת הלקוח. מי שקרא את
    הטבלה קרא מספרים נכונים תחת שמות שגויים.
    הכלל שנעול כאן: בכל תצוגה, ובכל מצב של בורר העמודות, מספר
    הכותרות שווה למספר התאים. */
 const cols=()=>p.evaluate(()=>{
   const th=[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim());
   const tr=document.querySelector('#tbl tbody tr[data-i]');
   return {th,nTh:th.length,nTd:tr?tr.children.length:0,
     keys:(typeof visCols==='function')?visCols():null,
     view:(typeof colViewKey==='function')?colViewKey():null}});
 await p.evaluate(()=>{const a=[...document.querySelectorAll('#tabs .tab')]
   .find(t=>t.dataset.m==='today');if(a)a.click()});
 await p.waitForTimeout(600);
 const cDec=await cols();
 await p.evaluate(()=>{const b=document.querySelector('#gseg button[data-gb="supplier"]');if(b)b.click()});
 await p.waitForTimeout(600);
 const cSup=await cols();
 await p.evaluate(()=>{const b=document.querySelector('#gseg button[data-gb="decision"]');if(b)b.click()});
 await p.waitForTimeout(600);
 ok('«לפי החלטה» — כותרת לכל תא',cDec.nTh===cDec.nTd&&cDec.nTh>0,
   `${cDec.nTh} כותרות · ${cDec.nTd} תאים`);
 ok('«לפי ספק» — כותרת לכל תא',cSup.nTh===cSup.nTd&&cSup.nTh>0,
   `${cSup.nTh} כותרות · ${cSup.nTd} תאים · ${cSup.th.join(' | ')}`);
 ok('ולכל תצוגה פריסת עמודות משלה',cDec.view==='dec'&&cSup.view==='sup'
   &&cDec.keys.join(',')!==cSup.keys.join(','),
   `${cDec.keys.join(',')}  /  ${cSup.keys.join(',')}`);
 /* ברירת המחדל היא בדיוק מה שהמסך הציג לפני הבורר — הרישום לא
    שינה עמודות, רק את מקור האמת שלהן. */
 /* «מצב טיפול» נוספה לברירת המחדל אחרי שהמתכנן הכריע שסימון מסוג
    «בבדיקה / ממתין לספק / הוזמן» משאיר את הפריט בתור. לפני ההכרעה
    העמודה הייתה מציגה «חדש» בכל שורה תמיד.
    עמודת הרכש תלויה בנתונים ולא בהעדפה: עם דוח ETA היא מתפצלת
    ל«מכוסה» ו«ללא תאריך», ובלעדיו היא «בדרך ⌛» אחת. */
 const hasEta=await p.evaluate(()=>!!ETA);
 ok('ברירת המחדל של «לפי החלטה» לא השתנתה',
   cDec.th.join('|')===(hasEta
     ?'#|מק״ט|תיאור|ספק|11 חודשים|מלאי|דרישת לקוח|מכוסה|ללא תאריך|חוסר חזוי|שווי|הפעולה הבאה|מצב טיפול|ותק'
     :'#|מק״ט|תיאור|ספק|11 חודשים|מלאי|דרישת לקוח|בדרך ⌛|חוסר חזוי|שווי|הפעולה הבאה|מצב טיפול|ותק'),
   `ETA=${hasEta} · ${cDec.th.join('|')}`);

 await p.click('#colsBtn');await p.waitForTimeout(350);
 ok('בורר העמודות נפתח',
   await p.evaluate(()=>document.getElementById('colspick').classList.contains('open')));
 ok('ועמודה שאינה אפשרית במצב הנתונים מוצגת כבויה עם הסיבה',
   await p.evaluate(k=>{const cb=document.querySelector(`#colspick input[data-ck="${k}"]`);
     const lb=cb&&cb.closest('label');
     return !!cb&&cb.disabled&&!!lb&&/דוח ETA/.test(lb.textContent)},hasEta?'otw':'cov'),
   hasEta?'«בדרך ⌛» כבויה כשיש דוח':'«מכוסה» כבויה כשאין דוח');
 await p.click('#colspick input[data-ck="po"]');await p.waitForTimeout(450);
 await p.click('#colspick input[data-ck="tr"]');await p.waitForTimeout(450);
 const cAdd=await cols();
 ok('הוספת «רכש פתוח» ו«בהעברה» מוסיפה כותרת ותא לכל אחת',
   cAdd.nTh===cDec.nTh+2&&cAdd.nTd===cAdd.nTh
   &&cAdd.th.includes('רכש פתוח')&&cAdd.th.includes('בהעברה'),
   `${cAdd.nTh} כותרות · ${cAdd.nTd} תאים`);
 /* «בהעברה» הוא עמודת CK. בדוח שאין בה את העמודה הזאת התא אומר
    «לא ידוע» ולא 0 — אותו כלל של רצועת המספרים בכרטיס. */
 ok('ותא «בהעברה» נשען על השדה ולא על 0',
   await p.evaluate(()=>{const i=[...document.querySelectorAll('#tbl thead th')]
       .findIndex(t=>t.textContent.trim()==='בהעברה');
     if(i<0)return false;
     const tr=document.querySelector('#tbl tbody tr[data-i]');
     const cell=(tr.children[i].textContent||'').trim();
     const r=ALL.find(x=>x.pn===tr.children[1].textContent.replace(/העתק|✓ טופל/g,'').trim());
     if(!r)return false;
     return r.transferKnown?cell===String(r.transfer||0)||cell==='—':/לא ידוע/.test(cell)}));
 await p.click('#colspick input[data-ck="pn"]');await p.waitForTimeout(400);
 ok('«מק״ט» אינו ניתן להסרה — בלעדיו השורה אינה מזוהה',
   await p.evaluate(()=>{const th=[...document.querySelectorAll('#tbl thead th')]
     .map(t=>t.textContent.trim());return th.includes('מק״ט')}));
 await p.reload();await p.waitForTimeout(1600);
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(2600);
 await p.evaluate(()=>{const a=[...document.querySelectorAll('#tabs .tab')]
   .find(t=>t.dataset.m==='today');if(a)a.click()});
 await p.waitForTimeout(600);
 const cKeep=await cols();
 ok('הפריסה שורדת רענון',cKeep.th.includes('רכש פתוח')&&cKeep.nTh===cKeep.nTd,
   `${cKeep.nTh} כותרות · ${cKeep.nTd} תאים`);
 /* עמודות מוקפאות נדלקות רק כשיש גלילה אופקית בפועל */
 const fz=await p.evaluate(()=>{const box=document.querySelector('#w_queue .rows');
   const before=box.scrollWidth>box.clientWidth+1;
   const td=document.querySelector('#tbl tbody tr[data-i] td:nth-child(2)');
   const x0=td.getBoundingClientRect().left;
   box.scrollLeft=-250;
   return new Promise(res=>setTimeout(()=>res({ovf:before,
     freeze:document.body.classList.contains('freeze'),
     pos:getComputedStyle(td).position,
     moved:Math.abs(td.getBoundingClientRect().left-x0)}),250))});
 ok('בגלילה אופקית המק״ט נשאר במקומו',
   !fz.ovf||(fz.freeze&&fz.pos==='sticky'&&fz.moved<6),
   `גלישה=${fz.ovf} · הקפאה=${fz.freeze} · זז ${Math.round(fz.moved)}px`);
 await p.click('#colsBtn');await p.waitForTimeout(300);
 await p.click('#colspick [data-cpa="reset"]');await p.waitForTimeout(500);
 const cRst=await cols();
 ok('«חזרה לברירת המחדל» מחזירה בדיוק את העמודות שהיו',
   cRst.th.join('|')===cDec.th.join('|'),cRst.th.join('|'));
 await p.evaluate(()=>{try{localStorage.removeItem('planner_cols_v1')}catch(_){}});
 await p.click('#colspick [data-cpa="close"]');await p.waitForTimeout(250);

 /* ============ מיון בתוך הקיבוץ ============
    לטבלת העבודה לא היה מיון בכלל — הכותרת לא הייתה לחיצה. הכלל
    שנעול: המיון פועל *בתוך* קבוצה ובתוך בלוק מטבע, ולא חוצה אותם.
    נמדד לפני התיקון: מיון לפי «דרישת לקוח» פיצל את כותרות המטבע
    מ-8 ל-75, כלומר הפך «דירוג בתוך המטבע בלבד» לרשימה מעורבת. */
 const wsnap=()=>p.evaluate(()=>{
   const ci=[...document.querySelectorAll('#tbl thead th')]
     .findIndex(t=>/דרישת לקוח/.test(t.textContent));
   return {ws:(typeof WSORT!=='undefined')?{...WSORT}:null,
     groups:[...document.querySelectorAll('#tbl tbody tr.grp')].length,
     vals:[...document.querySelectorAll('#tbl tbody tr[data-i]')].slice(0,8)
       .map(tr=>+(tr.children[ci].textContent.replace(/[^\d]/g,'')||0)),
     order:[...document.querySelectorAll('#tbl tbody tr[data-i]')].slice(0,6)
       .map(tr=>tr.children[1].textContent.replace(/העתק|✓ טופל/g,'').trim()).join(',')}});
 const hitSort=()=>p.evaluate(()=>{const th=[...document.querySelectorAll('#tbl thead th.wsrt')]
   .find(t=>/דרישת לקוח/.test(t.textContent));if(th)th.querySelector('.thc').click()});
 const w0=await wsnap();
 await hitSort();await p.waitForTimeout(500);const w1=await wsnap();
 await hitSort();await p.waitForTimeout(500);const w2=await wsnap();
 await hitSort();await p.waitForTimeout(500);const w3=await wsnap();
 const desc=a=>a.every((v,i)=>i===0||a[i-1]>=v), asc=a=>a.every((v,i)=>i===0||a[i-1]<=v);
 ok('לחיצה על כותרת ממיינת יורד',w1.ws.col==='cust'&&w1.ws.dir==='desc'&&desc(w1.vals),
   w1.vals.join(' · '));
 ok('לחיצה שנייה הופכת לעולה',w2.ws.dir==='asc'&&asc(w2.vals),w2.vals.join(' · '));
 ok('ושלישית מחזירה לסדר המקורי',w3.ws.col===null&&w3.order===w0.order,
   `${w0.order}  →  ${w3.order}`);
 ok('המיון אינו חוצה בלוק מטבע — מספר הקבוצות אינו משתנה',
   w1.groups===w0.groups&&w2.groups===w0.groups,
   `${w0.groups} → ${w1.groups} → ${w2.groups}`);

 /* ============ הייצוא מייצג את מה שרואים ============
    השורות תמיד היו של התצוגה, אבל העמודות היו 51 קבועות והסדר היה
    של currentRows() ולא של המסך — כלומר מי שהוריד קובץ קיבל משהו
    אחר ממה שראה. */
 const vx=await p.evaluate(()=>{const v=viewExportRows();
   const th=[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim()).slice(1);
   const screen=[...document.querySelectorAll('#tbl tbody tr[data-i]')].slice(0,5)
     .map(tr=>tr.children[1].textContent.replace(/העתק|✓ טופל/g,'').trim());
   return {head:v.aoa[0],n:v.n,th,first:v.aoa.slice(1,6).map(r=>r[0]),screen,
     rows:currentRows().length}});
 ok('גיליון «התצוגה» נושא בדיוק את העמודות הגלויות',
   vx.head.join('|')===vx.th.filter(t=>t!=='11 חודשים').join('|'),
   `${vx.head.join(' | ')}`);
 ok('ובסדר שבו הן מופיעות על המסך',vx.first.join(',')===vx.screen.join(','),
   `${vx.screen.join(' · ')}  →  ${vx.first.join(' · ')}`);
 ok('וכל השורות שבתצוגה נמצאות בו',vx.n===vx.rows,`${vx.n} / ${vx.rows}`);

 /* ============ מה נשמר כשחוזרים מפריט ============ */
 await p.evaluate(()=>{document.querySelector('#w_queue .rows').scrollTop=420});
 await p.waitForTimeout(250);
 await hitSort();await p.waitForTimeout(500);
 const st0=await p.evaluate(()=>({top:Math.round(document.querySelector('#w_queue .rows').scrollTop),
   col:WSORT.col,dir:WSORT.dir,
   first:(document.querySelector('#tbl tbody tr[data-i]')||{}).children[1].textContent.replace(/העתק|✓ טופל/g,'').trim()}));
 await p.evaluate(()=>{const rows=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
   rows[Math.min(6,rows.length-1)].click()});
 await p.waitForTimeout(500);
 await p.evaluate(()=>closeDetail());await p.waitForTimeout(400);
 const st1=await p.evaluate(()=>({top:Math.round(document.querySelector('#w_queue .rows').scrollTop),
   col:WSORT.col,dir:WSORT.dir,
   first:(document.querySelector('#tbl tbody tr[data-i]')||{}).children[1].textContent.replace(/העתק|✓ טופל/g,'').trim()}));
 ok('פתיחת פריט וסגירתו משאירות את הגלילה, המיון והשורה הראשונה',
   st1.top===st0.top&&st1.col===st0.col&&st1.dir===st0.dir&&st1.first===st0.first,
   `גלילה ${st0.top}→${st1.top} · מיון ${st0.col}/${st0.dir}→${st1.col}/${st1.dir}`);
 await p.evaluate(()=>{if(WSORT.col)wsortClick(WSORT.col)});await p.waitForTimeout(300);
 await p.evaluate(()=>{if(WSORT.col)wsortClick(WSORT.col)});await p.waitForTimeout(400);

 /* ============ מצב טיפול — סימון שאינו מסתיר ============
    סעיף 6 במפרט: להפריד את מצב הטיפול מחומרת הבעיה. עד ההכרעה
    הזאת כל סימון הוציא את הפריט מתור העבודה, ולכן עמודת «מצב
    טיפול» הייתה מציגה «חדש» בכל שורה תמיד.
    מה שנעול כאן: «ממתין לספק» אינו פתרון לחוסר ואינו מסתיר אותו,
    ואינו נוגע בסיווג, בחומרה, בחוסר או בהון הכלוא. */
 await p.evaluate(()=>{const a=[...document.querySelectorAll('#tabs .tab')]
   .find(t=>t.dataset.m==='today');if(a)a.click()});
 await p.waitForTimeout(600);
 const wcell=()=>p.evaluate(()=>{
   const th=[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim());
   const i=th.findIndex(t=>/מצב טיפול/.test(t));
   const tr=document.querySelector('#tbl tbody tr[data-i]');
   const pn=tr?tr.children[1].textContent.replace(/העתק|✓ טופל/g,'').trim():null;
   const r=ALL.find(x=>x.pn===pn);
   return {i,pn,txt:i>=0&&tr?tr.children[i].textContent.trim():null,
     cls:i>=0&&tr?(tr.children[i].querySelector('.wchip')||{className:''}).className:'',
     sevCls:tr?tr.className:'',
     wip:r&&r.wip?r.wip.t:null,cat:r?r.cat:null,sev:r?r.sev:null,
     miss:r?r.miss:null,expCap:r?Math.round(r.expCap||0):null,
     n:document.querySelectorAll('#tbl tbody tr[data-i]').length}});
 const ws0=await wcell();
 ok('«מצב טיפול» היא עמודה, ובלי סימון היא אומרת «חדש»',
   ws0.i>0&&ws0.txt==='חדש',`עמודה ${ws0.i} · "${ws0.txt}"`);
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(450);
 ok('ובכרטיס יש מקטע נפרד למצב הטיפול',
   await p.evaluate(()=>{const sec=document.querySelector('#detail .wipsec');
     return !!sec&&/אינו מוציא את הפריט מתור העבודה/.test(sec.textContent)
       &&sec.querySelectorAll('.wipm button[data-w]').length===4}));
 await p.click('#detail .wipm button[data-w="supplier"]');await p.waitForTimeout(700);
 const ws1=await wcell();
 ok('«ממתין לספק» אינו מסתיר את הפריט — הוא נשאר ברשימה ובמקומו',
   ws1.n===ws0.n&&ws1.pn===ws0.pn,`${ws0.n} → ${ws1.n} שורות · ${ws0.pn} → ${ws1.pn}`);
 ok('והסיווג, החומרה, החוסר וההון הכלוא לא זזו',
   ws1.cat===ws0.cat&&ws1.sev===ws0.sev&&ws1.miss===ws0.miss&&ws1.expCap===ws0.expCap
   &&ws1.sevCls===ws0.sevCls,
   `${ws0.cat}/${ws0.sev}/${ws0.miss} → ${ws1.cat}/${ws1.sev}/${ws1.miss}`);
 ok('והעמודה מציגה את המצב, עם טקסט ולא בצבע בלבד',
   ws1.wip==='supplier'&&/לספק/.test(ws1.txt)&&/supplier/.test(ws1.cls),
   `"${ws1.txt}" · ${ws1.cls}`);
 await p.reload();await p.waitForTimeout(1600);
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(2600);
 await p.evaluate(()=>{const a=[...document.querySelectorAll('#tabs .tab')]
   .find(t=>t.dataset.m==='today');if(a)a.click()});
 await p.waitForTimeout(600);
 const ws2=await wcell();
 ok('המצב שורד רענון',ws2.wip==='supplier'&&ws2.pn===ws0.pn,`${ws2.pn} · ${ws2.wip}`);
 /* «טופל» לא השתנה: הוא עדיין מוציא מהתור. זה ההבדל בין שני
    הסוגים, והוא נבדק כאן ולא מונח. */
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(450);
 await p.click('#detail .wipm button[data-w=""]');await p.waitForTimeout(600);
 const ws3=await wcell();
 ok('«נקה» מחזיר ל«חדש»',ws3.wip===null&&ws3.txt==='חדש'&&ws3.n===ws0.n,
   `"${ws3.txt}" · ${ws3.n} שורות`);
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(450);
 await p.click('#detail [data-dmark="handled"]');await p.waitForTimeout(700);
 const ws4=await wcell();
 ok('ו«טופל» ממשיך להוציא מהתור — שני סוגי סימון, שתי התנהגויות',
   ws4.n===ws0.n-1,`${ws0.n} → ${ws4.n} שורות`);
 await p.evaluate(()=>{Object.keys(MARKS).forEach(k=>delete MARKS[k]);saveMarks();apply()});
 await p.waitForTimeout(600);

 /* ============ הכרטיס: המספרים ראשונים, והמעבר בלי לסגור ============
    סעיף 4 במפרט. שני כללים נעולים כאן:
    · רצועת המספרים מציגה את שדות המנוע כפי שהם — לא חישוב מחדש
      בתוך התצוגה, ולא כמות שנשלפת מטקסט ההמלצה.
    · «בהעברה» אומר «לא ידוע» כשאין עמודת CK בדוח. 0 ו«לא יודעים»
      אינם אותו דבר, וזה בדיוק הדפוס של custBlockedKnown. */
 await p.evaluate(()=>{const a=[...document.querySelectorAll('#tabs .tab')]
   .find(t=>t.dataset.m==='today');if(a)a.click()});
 await p.waitForTimeout(600);
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(450);
 const cst=await p.evaluate(()=>{const s=document.querySelector('#detail .cstrip');
   const r=CURRENT_DETAIL,g=l=>{const e=[...s.querySelectorAll('.cs')]
     .find(x=>x.querySelector('i').textContent===l);
     return e?e.querySelector('b').textContent.replace(/[^\d]/g,''):null};
   return {seen:!!s&&s.offsetParent!==null,
     order:[...s.querySelectorAll('.cs i')].map(e=>e.textContent).join('|'),
     eq:{cust:g('לקוח')==String(Math.max(0,r.cust||0)),
         free:g('פנוי')==String(Math.max(0,r.free||0)),
         po:g('רכש פתוח')==String(Math.max(0,r.po||0)),
         miss:g('פער')==String(Math.max(0,r.miss||0))},
     vals:`לקוח ${g('לקוח')} · פנוי ${g('פנוי')} · רכש ${g('רכש פתוח')} · פער ${g('פער')}`,
     unknown:coverStrip({cust:8,free:2,po:0,transfer:0,transferKnown:false,miss:6})}});
 ok('רצועת המספרים פותחת את הכרטיס בסדר שנקבע',
   cst.seen&&cst.order==='לקוח|פנוי|רכש פתוח|בהעברה|פער',cst.order);
 ok('וכל מספר בה הוא שדה המנוע עצמו',
   cst.eq.cust&&cst.eq.free&&cst.eq.po&&cst.eq.miss,cst.vals);
 ok('«בהעברה» אומר «לא ידוע» כשאין עמודת CK — לא 0',
   /לא ידוע/.test(cst.unknown)&&!/>0<\/b>\s*<i>בהעברה/.test(cst.unknown),
   (cst.unknown.match(/בהעברה/)?'נמצא':'חסר'));

 const nav1=await p.evaluate(()=>({pn:document.querySelector('#detail .opnt').textContent.trim(),
   pos:(document.querySelector('#detail .opos')||{}).textContent,
   prevOff:(document.querySelector('#detail [data-dnav="-1"]')||{}).disabled}));
 await p.click('#detail [data-dnav="1"]');await p.waitForTimeout(450);
 const nav2=await p.evaluate(()=>({pn:document.querySelector('#detail .opnt').textContent.trim(),
   pos:(document.querySelector('#detail .opos')||{}).textContent,
   open:document.body.classList.contains('dopen'),
   selPn:(()=>{const t=document.querySelector('#tbl tbody tr.sel');
     return t?t.children[1].textContent.replace('העתק','').replace('✓ טופל','').trim():null})()}));
 ok('בפריט הראשון אין «הקודם» לאן ללחוץ',nav1.prevOff===true,String(nav1.prevOff));
 ok('«הבא» מעביר לפריט אחר בלי לסגור את הכרטיס',
   nav2.pn!==nav1.pn&&nav2.open,`${nav1.pn} → ${nav2.pn}`);
 ok('והשורה המסומנת בטבלה היא הפריט שבכרטיס',nav2.selPn===nav2.pn,
   `שורה ${nav2.selPn} · כרטיס ${nav2.pn}`);
 ok('והמיקום ברשימה מתעדכן',nav2.pos!==nav1.pos,`${nav1.pos} → ${nav2.pos}`);
 await p.evaluate(()=>closeDetail());await p.waitForTimeout(200);

 /* ============ בריאות המלאי על נתוני הדגמה מלאים ============
    ב-routes.xlsx יש שני פריטים עם הון כלוא ושניהם עודף, ולכן הכלל
    «מה שאינו באף דלי נאמר במפורש» עובר שם בלי לבדוק דבר. כאן יש
    900 פריטים ורוב ההון הכלוא אינו באף אחד משלושת הדליים. */
 await p.evaluate(()=>{const a=[...document.querySelectorAll('#tabs .tab')]
   .find(t=>t.dataset.m==='cap');if(a)a.click()});
 await p.waitForTimeout(800);
 const capS=await p.evaluate(()=>{const {list,B}=capBuckets();
   const el=document.getElementById('capTop');
   return {other:B.other.length,n:list.length,
     note:(el.querySelector('.cnote')||{}).textContent||'',
     rows:currentRows().length,
     tiles:[...el.querySelectorAll('.ctile .cn')].map(x=>+x.textContent.replace(/[^\d]/g,'')),
     bars:el.querySelectorAll('.cbar').length,
     panels:el.querySelectorAll('.cpanel').length,
     curs:[...new Set(list.map(r=>r.currency||'—'))].length}});
 ok('מה שאינו עודף, איטי או מת נאמר במפורש ולא נבלע',
   capS.other>0&&capS.note.includes(capS.other.toLocaleString('he-IL'))
   &&capS.note.includes(capS.n.toLocaleString('he-IL')),
   `${capS.other} מתוך ${capS.n} · "${capS.note.slice(0,70)}"`);
 ok('והטבלה מציגה בדיוק את רשימת ההון הכלוא',capS.rows===capS.n&&capS.tiles[0]===capS.n,
   `${capS.rows} שורות · ${capS.n} ברשימה`);
 ok('לוח לכל מטבע, ועמודות לספקים בתוכו',
   capS.panels===capS.curs&&capS.bars>=capS.curs,
   `${capS.panels} לוחות · ${capS.curs} מטבעות · ${capS.bars} עמודות`);
 await p.evaluate(()=>{const a=[...document.querySelectorAll('#tabs .tab')]
   .find(t=>t.dataset.m==='today');if(a)a.click()});
 await p.waitForTimeout(600);

 /* ============ שתי בקרות הגובה: מדדים וצפיפות ============
    סעיף 3ג במפרט: «צפיפות רגילה וצפופה», ויעד של 18-24 שורות קריאות,
    עם איסור מפורש להגיע לשם בהקטנת טקסט. שתיהן נבדקות על אותה
    מידה בדיוק: כמה שורות נכנסו, ובאיזה גופן. */
 const rowsFit=()=>p.evaluate(()=>{const box=document.querySelector('.rows');
   const trs=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
   if(!box||!trs.length)return null;
   const b=box.getBoundingClientRect();
   return {h:Math.round(trs[0].getBoundingClientRect().height),
     n:trs.filter(t=>{const r=t.getBoundingClientRect();
       return r.top>=b.top-1&&r.bottom<=b.bottom+1}).length,
     font:parseFloat(getComputedStyle(trs[0].querySelector('td')).fontSize),
     dense:document.body.classList.contains('dense')}});
 await p.evaluate(()=>{document.body.classList.remove('dense');
   try{localStorage.setItem('planner_dense','0')}catch(_){}});
 await p.waitForTimeout(250);
 const den0=await rowsFit();
 await p.click('#denTog');await p.waitForTimeout(300);
 const den1=await rowsFit();
 ok('מתג הצפיפות קיים ומסומן כשהוא פעיל',
   await p.evaluate(()=>{const b=document.getElementById('denTog');
     return !!b&&b.offsetParent!==null&&b.getAttribute('aria-pressed')==='true'}));
 ok('«צפופה» מכניסה יותר שורות',den1.n>den0.n&&den1.h<den0.h,
   `${den0.n} שורות ב-${den0.h}px → ${den1.n} שורות ב-${den1.h}px`);
 /* הגופן של המסלול הזה הוא 11.5px גם ב«רגילה» — זו בחירה קודמת
    ולא תוצר הצפיפות. מה שנעול כאן: הצפיפות אינה נוגעת בו. */
 ok('והיא עושה זאת בלי להקטין את הטקסט',den1.font===den0.font,
   `${den0.font}px → ${den1.font}px`);
 await p.reload();await p.waitForTimeout(1800);
 ok('הבחירה שורדת רענון',
   await p.evaluate(()=>document.body.classList.contains('dense')));
 await p.evaluate(()=>{const b=document.getElementById('denTog');if(b)b.click()});
 await p.waitForTimeout(250);
 ok('ומתג המדדים מקפל את הרצועה ומחזיר גובה לרשימה',await (async()=>{
   await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(2600);
   const a=await p.evaluate(()=>Math.round(document.querySelector('.rows').getBoundingClientRect().height));
   await p.click('#kpiTog');await p.waitForTimeout(300);
   const b2=await p.evaluate(()=>Math.round(document.querySelector('.rows').getBoundingClientRect().height));
   await p.click('#kpiTog');await p.waitForTimeout(200);
   return b2>a})());

 await b.close();console.log(out.join('\n'));
 process.exit(out.some(l=>l.startsWith('FAIL'))?1:0)})();
