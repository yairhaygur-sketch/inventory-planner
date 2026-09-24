/* ============ הנגשה למעתד שלא בנה את הכלי ============
   מסך הפתיחה, שמות הפעולות, הסברי הניווט, משמעות «טופל», ההיכרות
   ומצב ההדגמה. הטענה המרכזית בכל אחד מהם היא אחת: מי שלא בנה את
   הכלי צריך לדעת איך להתחיל, איפה למצוא, ומה הוא בדיוק סימן. */
const {chromium}=require('playwright'),fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[],ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));
const load=async p=>{await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(2600);
 /* קודם היה כאן ביטול של התדריך המודאלי (`#briefGo`). התדריך אינו
    חוסם יותר — הוא רצועת סיכום מקופלת — ולכן אין מה לבטל. */};

(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});

 /* ---------- א · מסך הפתיחה, בשני הרוחבים שנדרשו ---------- */
 for(const W of [1366,1920]){
  const ctx=await b.newContext({viewport:{width:W,height:W===1366?768:1080}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(700);

  const w0=await p.evaluate(()=>{const w=document.getElementById('welcome');
   const t=(w.innerText||'').replace(/\s+/g,' ');
   const top=document.querySelector('.top').getBoundingClientRect();
   const wb=w.getBoundingClientRect();
   return {seen:!w.hidden&&wb.height>0,txt:t,
    /* חייב לכסות את אזור העבודה ולא את הסרגל — אחרת אי אפשר להעלות */
    belowTop:Math.round(wb.top)>=Math.round(top.bottom)-1,
    upload:[...document.querySelectorAll('label[for="f"]')].filter(e=>e.offsetParent!==null).length,
    search:document.getElementById('qbox').offsetParent!==null,
    hscroll:document.documentElement.scrollWidth>window.innerWidth+1}});
  ok(`${W} · מסך הפתיחה מוצג כשאין דוח`,w0.seen);
  ok(`${W} · והוא אומר מאיפה מתחילים`,
    /מתחילים מדוח המלאי שלך/.test(w0.txt)&&/העלאת דוח ZMRP/.test(w0.txt),w0.txt.slice(0,70));
  ok(`${W} · דוח ה-ETA מוצג כתוספת ולא כדרישה`,
    /תוספת שאפשר לצרף בהמשך/.test(w0.txt));
  ok(`${W} · והפרטיות נאמרת במפורש`,
    /מעובדים בדפדפן ואינם נשלחים לשרת/.test(w0.txt));
  /* «איך מפיקים את הדוח» הוסר לפי החלטת המעתד — הוא ידע להפיק אותו
     ממילא, והמסך תיאר פורמט במקום להוביל לעבודה. */
  ok(`${W} · שתי נקודות הכניסה קיימות`,
    /נתוני דוגמה/.test(w0.txt)&&/היכרות עם הכלי/.test(w0.txt));
  ok(`${W} · והסבר הפקת הדוח אינו שם יותר`,
    !/איך מפיקים/.test(w0.txt)
    &&await p.evaluate(()=>!document.getElementById('welHow')&&!document.getElementById('welHowBox')));
  ok(`${W} · הסרגל העליון נשאר לחיץ מתחת למסך הפתיחה`,
    w0.belowTop&&w0.upload===2&&w0.search,
    `מתחת לסרגל=${w0.belowTop} · תוויות העלאה גלויות=${w0.upload} · חיפוש=${w0.search}`);
  ok(`${W} · אין גלישה אופקית במסך הפתיחה`,!w0.hscroll);


  /* ---------- ב · שמות הפעולות ---------- */
  await p.evaluate(()=>tourEnd());
  await load(p);
  const act=await p.evaluate(()=>{
   const want={filtBtn:'סינון',exportXls:'ייצוא לאקסל',expXls:'ייצוא דוח זירוז',
     stBtn:'גיבוי והעברה',darkToggle:'מצב כהה'};
   const top=document.querySelector('.top');
   const st=Object.entries(want).map(([id,nm])=>{const e=document.getElementById(id);
    if(!e)return {id,ok:false,why:'חסר'};
    const lbl=((e.querySelector('.lbl')||{}).textContent||'').trim();
    const inMenu=!!e.closest('#moreMenu');
    return {id,ok:lbl===nm,lbl,why:inMenu?'בתפריט «עוד»':'בשורה'}});
   return {st,clipped:top.scrollWidth>top.clientWidth+2,
    help:((document.getElementById('helpBtn')||{}).textContent||'').replace(/\s+/g,' ').trim(),
    helpItems:[...document.querySelectorAll('#helpMenu .mi')].map(e=>e.dataset.help)}});
  ok(`${W} · לכל פעולה שם מלא לפי מה שהיא עושה`,act.st.every(x=>x.ok),
    act.st.map(x=>`${x.lbl||x.id}(${x.why})`).join(' · '));
  ok(`${W} · הסרגל אינו נחתך`,!act.clipped);
  ok(`${W} · «עזרה» הוא תפריט עם שתי כניסות`,
    /עזרה/.test(act.help)&&act.helpItems.join(',')==='tour,xp',act.helpItems.join(','));

  /* ---------- ג · הסברי הניווט ---------- */
  const nav=await p.evaluate(()=>{
   const tabs=[...document.querySelectorAll('#tabs .tab.navtab:not(.mini)')];
   return {what:tabs.map(t=>({m:t.dataset.m,
      w:((t.querySelector('.m.what')||{}).textContent||'').trim(),
      sel:t.classList.contains('sel')})),
    /* הכניסות המשניות הן קישורים ולא מסלולים ראשיים: אין להן בלוק
       תיאור, ולכן השם וההסבר שלהן נבדקים ב-title וב-aria-label. */
    mini:[...document.querySelectorAll('#tabs .tab.navtab.mini')]
      .map(t=>({m:t.dataset.m,t:(t.title||''),a:(t.getAttribute('aria-label')||'')})),
    hint:(document.getElementById('navhint').innerText||'').replace(/\s+/g,' '),
    hintSeen:document.getElementById('navhint').offsetParent!==null,
    rdLabel:((document.querySelector('.rdwrap .tl')||{}).textContent||'').trim(),
    rdTitle:(document.getElementById('rd')||{}).title||''}});
  ok(`${W} · לכל מסלול יש הסבר — גם למסלול שאינו הפעיל`,
    nav.what.length>=2&&nav.what.every(x=>x.w.length>5),
    nav.what.map(x=>`${x.m}${x.sel?'*':''}: ${x.w}`).join(' · '));
  ok(`${W} · לכניסות המשניות יש שם והסבר בריחוף ובקורא מסך`,
    nav.mini.length===2&&nav.mini.every(x=>x.t.length>15&&x.a.length>5),
    nav.mini.map(x=>`${x.m}: "${x.t.slice(0,50)}"`).join(' · '));
  ok(`${W} · נאמר במפורש שאלה מסלולי עבודה ולא מסנני תאריך`,
    nav.hintSeen&&/מסלולי עבודה/.test(nav.hint)&&/לא מסנני תאריך/.test(nav.hint));
  ok(`${W} · ונאמר איפה מוצאים פרמטרים, אספקות, טופלו ותנועות`,
    /תיקוני פרמטרים/.test(nav.hint)&&/אספקות צפויות/.test(nav.hint)
    &&/טופלו/.test(nav.hint)&&/תנועות/.test(nav.hint),nav.hint.slice(0,150));
  ok(`${W} · «חודש בסיס» קיבל תווית גלויה והסבר`,
    /חודש בסיס/.test(nav.rdLabel)&&/החודש המלא האחרון/.test(nav.rdTitle),
    `תווית="${nav.rdLabel}"`);
  if(errs.length)ok(`${W} · אין שגיאות JS`,false,errs.join(' | '));
  else ok(`${W} · אין שגיאות JS`,true);
  await ctx.close();
 }

 /* ---------- ד · משמעות «טופל» ---------- */
 {const ctx=await b.newContext({viewport:{width:1512,height:860}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  await load(p);
  await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(400);
  const m0=await p.evaluate(()=>{const d=document.getElementById('detail');

   return (d.innerText||'').replace(/\s+/g,' ')});
  /* ============ קצר ליד הכפתור, מלא לפי דרישה ============
     קודם ישבה כאן פסקה שלמה ליד הכפתור. היא הייתה נכונה ועמוסה.
     עכשיו: שורה אחת גלויה, והפירוט ב-<details> שנגיש במקלדת מעצם
     היותו אלמנט סטנדרטי. מספר הימים נגזר מ-markTTL. */
  const ttl=await p.evaluate(()=>DAYS(markTTL('handled')));
  ok('ליד הכפתור שורה קצרה אחת, עם מספר הימים מההגדרה',
    new RegExp(`סימון מקומי ל-${ttl} ימים`).test(m0)&&/אינו מעדכן SAP/.test(m0)&&/ניתן לבטל/.test(m0),
    (m0.match(/סימון מקומי[^.]*\.[^.]*\./)||[''])[0]);
  ok('וההסבר המלא אינו פתוח כברירת מחדל',
    !/הכלי לא בדק דבר/.test(m0),'הפירוט סגור');
  /* הפירוט — נפתח, ונגיש במקלדת */
  await p.evaluate(()=>{const d=document.querySelector('#detail .dcwhat');if(d)d.open=true});
  await p.waitForTimeout(200);
  const full=await p.evaluate(()=>(document.getElementById('detail').innerText||'').replace(/\s+/g,' '));
  ok('הפירוט מבחין בין סימון ידני לבין שינוי שאומת בדוח הבא',
    /הכלי לא בדק דבר/.test(full)&&/השינוי זוהה בדוח SAP הבא/.test(full)
    &&/הדוח הבא/.test(full),
    (full.match(/אינו[^.]*השינוי זוהה[^.]*\./)||[''])[0].slice(0,120));
  ok('ונאמר בו לכמה זמן',new RegExp(`${ttl} ימים`).test(full));
  ok('הפירוט נגיש במקלדת — הוא <details> תקני',
    await p.evaluate(()=>{const d=document.querySelector('#detail .dcwhat');
      return !!d&&d.tagName==='DETAILS'&&!!d.querySelector('summary')}));
  /* ============ סימון וביטול — דרך הממשק, לא דרך setMark ============
     בכרטיס מסומן היו *שני* כפתורי ביטול עם data-dmark, והחיווט הוא
     querySelector יחיד — כלומר השני מעולם לא היה מחובר. הבדיקה הקודמת
     בדקה שהוא *קיים*, ולכן עברה על כפתור מת. כאן נבדקת הפעולה עצמה:
     סימון דרך הכפתור, מעבר ל«טופלו», ביטול, וחזרה לתור. */
  const pn0=await p.evaluate(()=>CURRENT_DETAIL.pn);
  const cat0=await p.evaluate(()=>ALL.find(x=>x.pn===CURRENT_DETAIL.pn).cat);
  await p.click('#detail [data-dmark]');await p.waitForTimeout(500);
  ok('לחיצה על «סמן כטופל» בממשק מסמנת',
    await p.evaluate(pn=>{const r=ALL.find(x=>x.pn===pn);return !!(r.mark&&r.mark.t==='handled')},pn0));
  const m1=await p.evaluate(()=>(document.getElementById('detail').innerText||'').replace(/\s+/g,' '));
  ok('אחרי הסימון נאמר שסומן, ושאינו מעדכן SAP',
    /סומן כטופל/.test(m1)&&/אינו מעדכן SAP/.test(m1),
    (m1.match(/סומן כטופל[^·]*·[^·]*·/)||[''])[0]);
  ok('יש בדיוק כפתור ביטול אחד בכרטיס — לא שניים',
    await p.evaluate(()=>document.querySelectorAll('#detail [data-dmark]').length)===1,
    'data-dmark = '+await p.evaluate(()=>document.querySelectorAll('#detail [data-dmark]').length));
  ok('והוא הכפתור הראשי «בטל סימון טופל»',
    await p.evaluate(()=>{const b=document.querySelector('#detail [data-dmark]');
      return !!b&&/בטל סימון טופל/.test(b.textContent||'')}));
  /* --- פותחים אותו מתוך «טופלו», דרך לחיצה על רכיב גלוי --- */
  await p.evaluate(()=>closeDetail());await p.waitForTimeout(250);
  const doneTab=await p.evaluate(()=>{const e=document.querySelector('#tabs .tab[data-m="done"]');
    return !!e&&e.offsetParent!==null});
  ok('«טופלו» מופיע בניווט אחרי שיש פריט מסומן',doneTab);
  await p.click('#tabs .tab[data-m="done"]');await p.waitForTimeout(500);
  const inDone=await p.evaluate(pn=>[...document.querySelectorAll('#tbl tbody tr[data-i]')]
    .some(tr=>(tr.textContent||'').includes(pn)),pn0);
  ok('והפריט המסומן נמצא שם',inDone,pn0);
  await p.click(`#tbl tbody tr[data-i]:has-text("${pn0}")`);await p.waitForTimeout(450);
  /* --- הביטול, במקלדת --- */
  await p.evaluate(()=>document.querySelector('#detail [data-dmark]').focus());
  await p.keyboard.press('Enter');await p.waitForTimeout(500);
  const after=await p.evaluate(pn=>{const r=ALL.find(x=>x.pn===pn);
    return {marked:!!r.mark,cat:r.cat}},pn0);
  ok('הפעלת הביטול במקלדת מוחקת את הסימון',!after.marked,
    `mark=${after.marked}`);
  ok('והפריט חוזר לסיווג שלו',after.cat===cat0,`${cat0} → ${after.cat}`);
  await ctx.close()}

 /* ---------- ה · ההיכרות ---------- */
 {const ctx=await b.newContext({viewport:{width:1512,height:860}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();
  await p.goto('file://'+path.join(SD,'..','index.html'));   /* בלי nobrief — כמו משתמש אמיתי */
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(1400);
  const t0=await p.evaluate(()=>{const t=document.getElementById('tour');
   return {seen:!t.hidden,n:(document.getElementById('tourN').textContent||''),
     ttl:(document.getElementById('tourTtl').textContent||'')}});
  ok('משתמש חדש מקבל את ההיכרות מעצמה',t0.seen&&t0.n==='1/5',`${t0.ttl} ${t0.n}`);
  const steps=[];
  for(let i=0;i<5;i++){
   steps.push(await p.evaluate(()=>document.getElementById('tourTtl').textContent));
   await p.click('#tourNext');await p.waitForTimeout(200)}
  ok('חמישה צעדים לאורך מסלול העבודה',
    steps.length===5&&/מתחילים מדוח/.test(steps[0])&&/מסלול/.test(steps[1])
    &&/פריט/.test(steps[2])&&/ההמלצה/.test(steps[3])&&/טיפול/.test(steps[4]),
    steps.join(' → '));
  ok('בסיום היא נסגרת',await p.evaluate(()=>document.getElementById('tour').hidden));
  /* לא חוזרת — לא ברענון, ולא בהעלאת דוח */
  await p.reload();await p.waitForTimeout(1400);
  ok('ואינה חוזרת ברענון',await p.evaluate(()=>document.getElementById('tour').hidden));
  await load(p);
  ok('ואינה חוזרת בהעלאת דוח',await p.evaluate(()=>document.getElementById('tour').hidden));
  /* אבל נפתחת מחדש מתפריט העזרה */
  await p.click('#helpBtn');await p.waitForTimeout(150);
  await p.click('#helpMenu .mi[data-help="tour"]');await p.waitForTimeout(300);
  ok('ונפתחת מחדש מתפריט העזרה',
    await p.evaluate(()=>!document.getElementById('tour').hidden&&document.getElementById('tourN').textContent==='1/5'));
  /* דילוג מוקדם גם הוא סוגר לתמיד */
  await p.click('#tourSkip');await p.waitForTimeout(200);
  ok('ודילוג סוגר אותה',await p.evaluate(()=>document.getElementById('tour').hidden));
  await ctx.close()}

 /* ---------- ו · מצב ההדגמה — מבודד מהעבודה האמיתית ---------- */
 {const ctx=await b.newContext({viewport:{width:1512,height:860}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  /* קודם עבודה אמיתית: דוח, סימון, וזיכרון שנשמר */
  await load(p);
  await p.evaluate(()=>{setMark(ALL[0],'campaign','אמיתי');setMark(ALL[1],'ignore','אמיתי')});
  await p.waitForTimeout(300);
  const real=await p.evaluate(()=>({marks:Object.keys(MARKS).length,
    keys:Object.keys(localStorage).filter(k=>k.startsWith('planner_')).sort(),
    blob:localStorage.getItem('planner_marks_v1')}));
  ok('לפני ההדגמה — יש עבודה אמיתית שמורה',real.marks===2&&!!real.blob,
    `${real.marks} סימונים · ${real.keys.length} מפתחות`);
  /* עכשיו הדגמה */
  await p.evaluate(()=>demoStart());await p.waitForTimeout(3000);
  const dm=await p.evaluate(()=>({demo:DEMO,n:ALL.length,
    bar:!document.getElementById('demobar').hidden,
    barTxt:(document.getElementById('demobar').innerText||'').replace(/\s+/g,' '),
    marks:Object.keys(MARKS).length,hist:HIST.runs.length,par:PARAMS.runs.length,
    ledger:Object.keys(LEDGER.days).length,
    cats:[...new Set(ALL.map(r=>r.cat))].length,
    pns:ALL.slice(0,3).map(r=>r.pn)}));
  ok('מצב ההדגמה טוען נתונים סינתטיים',dm.demo&&dm.n>40&&dm.cats>=5,
    `${dm.n} פריטים · ${dm.cats} סיווגים · ${dm.pns.join(',')}`);
  ok('והוא מסומן בבירור על המסך',dm.bar&&/מצב הדגמה/.test(dm.barTxt)&&/לא המלאי שלך/.test(dm.barTxt),
    dm.barTxt.slice(0,80));
  /* הזיכרונות מתאפסים בכניסה, ואז הרצת ההדגמה עצמה רושמת אחד —
     שלה. מה שחייב להיות אפס הוא *שאריות מהעבודה האמיתית*. */
  ok('הוא מתחיל נקי — בלי סימונים, ועם היסטוריה של ההדגמה בלבד',
    dm.marks===0&&dm.hist<=1&&dm.par<=1,
    `סימונים ${dm.marks} · תור ${dm.hist} · יישום ${dm.par} · פנקס ${dm.ledger}`);
  /* סימון בתוך ההדגמה אינו מגיע לאחסון */
  await p.evaluate(()=>{setMark(ALL[0],'handled','הדגמה')});await p.waitForTimeout(300);
  const after=await p.evaluate(()=>({marks:Object.keys(MARKS).length,
    blob:localStorage.getItem('planner_marks_v1'),
    keys:Object.keys(localStorage).filter(k=>k.startsWith('planner_')).sort().join(',')}));
  ok('סימון בהדגמה נראה על המסך אבל אינו נכתב לאחסון',
    after.marks===1&&after.blob===real.blob,
    `בזיכרון ${after.marks} · האחסון זהה למקור=${after.blob===real.blob}`);
  ok('ואף מפתח אחסון לא נוסף או השתנה בגלל ההדגמה',
    after.keys===real.keys.join(','),`${real.keys.join(',')} → ${after.keys}`);
  /* והיציאה מחזירה את האמת */
  await p.reload();await p.waitForTimeout(400);
  await load(p);
  const back=await p.evaluate(()=>({demo:DEMO,marks:Object.keys(MARKS).length,
    notes:Object.values(MARKS).map(m=>m.note).sort().join(','),
    bar:!document.getElementById('demobar').hidden}));
  ok('יציאה מההדגמה מחזירה את העבודה האמיתית במדויק',
    !back.demo&&!back.bar&&back.marks===2&&back.notes==='אמיתי,אמיתי',
    `DEMO=${back.demo} · ${back.marks} סימונים · הערות=${back.notes}`);
  if(errs.length)ok('אין שגיאות JS בהדגמה',false,errs.join(' | '));
  else ok('אין שגיאות JS בהדגמה',true);
  await ctx.close()}

 /* ---------- ז · בידוד ההדגמה: מחיקות, שחזורים, ו-ETA ---------- */
 {const ctx=await b.newContext({viewport:{width:1512,height:860}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  await load(p);
  await p.setInputFiles('#fe',SD+'/zmrp-demo-eta.xlsx');await p.waitForTimeout(1500);
  await p.evaluate(()=>{setMark(ALL[0],'campaign','אמיתי');setMark(ALL[1],'ignore','אמיתי')});
  await p.waitForTimeout(300);
  const base=await p.evaluate(()=>({eta:!!ETA,etaLS:!!localStorage.getItem('planner_eta_v1'),
    marks:localStorage.getItem('planner_marks_v1')}));
  ok('בסיס: דוח ETA אמיתי וסימונים אמיתיים שמורים',base.eta&&base.etaLS&&!!base.marks);
  await p.evaluate(()=>demoStart());await p.waitForTimeout(3000);
  const iso=await p.evaluate(prev=>{const L={};
   L.etaMem=!!ETA;
   etaClear();                                   /* מחיקה = כתיבה */
   L.etaLSafter=!!localStorage.getItem('planner_eta_v1');
   const fake={keys:{'planner_marks_v1':JSON.stringify({ZZZ:{t:'handled',ts:1}}),
     'planner_history_v1':JSON.stringify({runs:[{sig:'x'}],items:{}})}};
   L.restore=stateRestoreAll(fake);
   L.merged=stateMergeMarks(fake).added;
   L.marksIntact=localStorage.getItem('planner_marks_v1')===prev;
   return L},base.marks);
  ok('בהדגמה — דוח ה-ETA האמיתי יורד מהזיכרון',!iso.etaMem);
  ok('«הסר דוח ETA» בהדגמה אינו מוחק אותו מהאחסון',iso.etaLSafter);
  ok('«שחזור מלא» חסום בהדגמה',iso.restore===false);
  ok('«מיזוג סימונים» חסום בהדגמה',iso.merged===0);
  ok('והסימונים השמורים נשארו בדיוק כפי שהיו',iso.marksIntact);

  /* ---------- ח · העלאת דוח אמיתי מתוך ההדגמה ---------- */
  await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(3000);
  const back=await p.evaluate(()=>{
   setMark(ALL[2],'handled','אחרי יציאה');
   let ls={};try{ls=JSON.parse(localStorage.getItem('planner_marks_v1')||'{}')}catch(_){}
   return {demo:DEMO,bar:!document.getElementById('demobar').hidden,
     marks:Object.keys(MARKS).length,eta:!!ETA,saved:Object.keys(ls).length}});
  ok('העלאת דוח אמיתי מתוך ההדגמה מבטלת את מצב ההדגמה',!back.demo&&!back.bar);
  ok('והעבודה האמיתית — סימונים ו-ETA — חוזרת לפני הטעינה',
    back.marks===3&&back.eta,`${back.marks} סימונים · ETA=${back.eta}`);
  ok('וסימון חדש על הנתונים האמיתיים באמת נשמר',back.saved===3,back.saved+' באחסון');
  if(errs.length)ok('אין שגיאות JS בבידוד',false,errs.join(' | '));
  else ok('אין שגיאות JS בבידוד',true);
  await ctx.close()}

 /* ---------- ט · מקלדת ---------- */
 {const ctx=await b.newContext({viewport:{width:1920,height:1080}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  await load(p);
  /* פריטי התפריט חייבים להיות פקדים, לא div-ים עם onclick */
  const kind=await p.evaluate(()=>[...document.querySelectorAll('#helpMenu .mi')]
    .map(e=>e.tagName+':'+(e.getAttribute('role')||'')));
  ok('פריטי העזרה הם כפתורים עם role=menuitem',
    kind.length===2&&kind.every(x=>x==='BUTTON:menuitem'),kind.join(' · '));
  await p.evaluate(()=>document.getElementById('helpBtn').focus());
  await p.keyboard.press('Enter');await p.waitForTimeout(200);
  const k1=await p.evaluate(()=>({open:document.getElementById('helpMenu').classList.contains('open'),
    exp:document.getElementById('helpBtn').getAttribute('aria-expanded'),
    f:document.activeElement.dataset.help||''}));
  ok('Enter פותח את התפריט ומעביר מיקוד לפריט הראשון',
    k1.open&&k1.exp==='true'&&k1.f==='tour',`open=${k1.open} focus=${k1.f}`);
  await p.keyboard.press('ArrowDown');await p.waitForTimeout(120);
  ok('חצים נעים בין הפריטים',
    'xp'===await p.evaluate(()=>document.activeElement.dataset.help||''));
  await p.keyboard.press('Escape');await p.waitForTimeout(150);
  const k3=await p.evaluate(()=>({open:document.getElementById('helpMenu').classList.contains('open'),
    f:document.activeElement.id}));
  ok('Escape סוגר ומחזיר את המיקוד לכפתור',!k3.open&&k3.f==='helpBtn',
    `open=${k3.open} focus=${k3.f}`);
  await p.keyboard.press('Enter');await p.waitForTimeout(150);
  await p.keyboard.press('ArrowDown');await p.waitForTimeout(100);
  await p.keyboard.press('Enter');await p.waitForTimeout(300);
  ok('Enter על פריט מפעיל אותו',await p.evaluate(()=>document.body.classList.contains('xp')));
  /* ============ הבאג שהתפריט חשף ============
     המנוע הגלובלי עשה preventDefault על Enter/רווח לכל מה שאינו שדה
     קלט, ולכן ביטל את ה-click הסינתטי — אף כפתור בכלי לא היה ניתן
     להפעלה במקלדת. הטענה כאן היא על כפתור ותיק, לא על החדשים. */
  await p.evaluate(()=>{xpSet(false);document.getElementById('filtBtn').focus()});
  const railBefore=await p.evaluate(()=>{const r=document.querySelector('.wfilters');
    return !!r&&r.offsetParent!==null});
  await p.keyboard.press('Enter');await p.waitForTimeout(400);
  const railAfter=await p.evaluate(()=>{const r=document.querySelector('.wfilters');
    return !!r&&r.offsetParent!==null});
  ok('Enter מפעיל גם כפתור ותיק בסרגל (היה חסום גלובלית)',
    railBefore!==railAfter,`מסילה ${railBefore} → ${railAfter}`);
  await ctx.close()}

 await b.close();console.log(out.join('\n'));
 process.exit(out.some(l=>l.startsWith('FAIL'))?1:0)})();
