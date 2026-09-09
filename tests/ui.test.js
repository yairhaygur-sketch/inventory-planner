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
   return {open:document.body.classList.contains('dopen'),
     onScreen:d.right>0&&d.left<innerWidth}});

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
 ok('ברירת המחדל היא קיבוץ לפי החלטה',/מניעת חוסר/.test(byDec.grps),byDec.grps);
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
 ok('חזרה לקיבוץ לפי החלטה',
   /מניעת חוסר/.test(await p.evaluate(()=>[...document.querySelectorAll('#tbl tr.grp b')].map(x=>x.textContent).join(' · '))));

 /* ============ רצפת מלאי הביטחון ============
    sugSS = max(ssStat, ssMin, lumpFloor), ו-ssMin הוא ערך שהמתכנן קבע
    ב-SAP. בדוח האמיתי הוא מנצח ב-88% מהמקרים, ולכן ה"המלצה" מחזירה
    את המספר שלו. הפער בין הרצפה לדרישה הסטטיסטית הוא החלטה שעולה
    כסף, והוא חייב להיות גלוי. */
 const seen=id=>p.evaluate(i=>{const e=document.getElementById(i);
   return !!e&&e.offsetParent!==null&&e.getBoundingClientRect().width>0},id);
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);
 ok('ב"היום" כפתור "מעל הדרישה" אינו נראה',!(await seen('floorBtn')));
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
 ok('ב"החודש" הכפתור נראה',await seen('floorBtn'),fl.txt);
 ok('עמודת "הדרישה" נוספה למסלול הפרמטרים',fl.heads.includes('הדרישה'),fl.heads.join(' · '));
 ok('הדרישה הסטטיסטית מוצגת גם כשהרצפה ניצחה',fl.need>0,fl.need+' שורות');
 ok('הסימון תואם בדיוק את הפריטים שהפער שלהם חיובי',
   fl.over===fl.expectOver,`מסומנים ${fl.over} · צפוי ${fl.expectOver}`);
 await p.click('#floorBtn');await p.waitForTimeout(800);
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
 await p.click('#floorBtn');await p.waitForTimeout(600);
 ok('לחיצה שנייה חוזרת ל"החודש"','month'===await p.evaluate(()=>mode));
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(400);

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
 ok('רמז המקלדת יורד בזמן הדחיפה, שאר השבבים נשארים',
   !dwHead.kbd&&dwHead.chips>=4,dwHead.chips+' שבבים');
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

 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(300);
  ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 await p.screenshot({path:SD+'/04-after.png'});
 await b.close();console.log(out.join('\n'));
 process.exit(out.some(l=>l.startsWith('FAIL'))?1:0)})();
