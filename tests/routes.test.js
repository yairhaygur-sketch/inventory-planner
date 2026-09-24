/* ============ הדרכים למסלולים, והמספרים שמצביעים עליהן ============
   שלושה דברים נבדקים כאן, וכולם דרך הממשק:

   1. המונה בקטלוג, הניסוח שלו והרשימה שהכפתור פותח — אותה אוכלוסייה.
      קודם המונה ספר decisionList('short') (inToday — בלי דרישת קצב)
      והכפתור הוביל ל-setMode('line'), שכל רשימותיו נגזרות מ-lineDry
      (rate>0). נמדד על דוח ההדגמה: מונה 300 מול יעד של 27 כש-LINE_SEL
      היה 'nopo', ו-235 כשהיה 'recur'. גם בברירת המחדל נראו 265 שורות
      מתוך 300, כי קבוצה מקופלת מסתירה שורות מהרינדור.

   3. הכותרת והכרטיסים של רצועת הסיכום מתארים אותו בסיס השוואה.

   4. נגישות המסלולים נבדקת בלחיצות על רכיבים גלויים בלבד. קריאה ישירה
      ל-setMode() אינה מוכיחה שמעתד יכול להגיע לשם. */
const XLSX=require('xlsx'),{chromium}=require('playwright'),fs=require('fs'),path=require('path');
const SD=__dirname, MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const hdr=['מק"ט מוביל','תיאור חומר','תיאור חומר2','שם ספק','סטטוס חומר','תיאור','סוג MRP','ABC','רמת שרות',
 'מלאי בטחון','נק.הז.מחדש','אספ.מתוכנ.','זמ.עב.קבלת','מל.בט.מינ.','מחיר FOB','מטבע FOB','סה"כ מלאי','מלאי פנוי',
 'מלאי מרלוג','הז. רכש','בהעברה','אספקות פת.','כמות בהז.פ','סוג חומר','תיא.קבוצ.חומרים','קב.חו.חיצו','טקסט ארוך',
 'תב.אח.הש.','ת.היר.1 מח','ת.היר.2 מח','היררכייה1','היררכייה1','היררכייה2','היררכייה2','היררכייה3','היררכייה3',
 'צר.השנה','צר.שנה-1','צר.שנה-2','צר.החודש',...MON,"תאר' מכירה","תאר' כניסה"];
const D=n=>{const d=new Date(Date.now()-n*864e5);return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`};
const mk=(pn,o)=>{const m=o.months;const y=m.reduce((a,b)=>a+b,0);
 /* !=null ולא ||: המבחן צריך דווקא אפס ברמת שירות ובזמן אספקה */
 return [pn,'פריט מבחן '+pn,'Test '+pn,o.sup||'ספק',o.status||'01',o.stx||'פעיל',o.mrp||'ND','A',o.srv!=null?o.srv:95,
  o.ss||0,o.rop||0,o.lt!=null?o.lt:30,2,o.ssMin||0,o.price!=null?o.price:100,o.cur||'USD',o.free,o.free,0,o.po||0,0,0,o.cust||0,'Z004','מנוע','ZT','מתכנן','0',
  'שיווק','מערכת','100','ZT','200','דגם','300','מערכת',
  o.y0!=null?o.y0:y,o.y1!=null?o.y1:y,o.y2!=null?o.y2:y,0,...m,D(o.saleAgo||10),D(o.entAgo||60)]};

/* המקרה שהמפרט נקב בו: לקוח ממתין, אפס מלאי, קצב צריכה אפס.
   inToday() סופר אותו (isT1). lineDry() דורש rate>0 ולכן מדלג עליו. */
const rows=[
 mk('ZERO-RATE-CUST',{months:[0,0,0,0,0,0,0,0,0,0,0],y0:0,y1:0,y2:0,free:0,cust:3,price:400,lt:120}),
 /* פריט עם אספקה מתוארכת — נכנס לשתי האוכלוסיות, ונשאר לאחר ETA_COVER */
 mk('DATED-SUPPLY',{months:[10,10,10,10,10,10,10,10,10,10,10],free:0,cust:2,po:40,price:120,lt:60}),
 /* חוסר רגיל עם קצב — בשתיהן */
 mk('PLAIN-SHORT',{months:[20,20,20,20,20,20,20,20,20,20,20],free:2,price:90,lt:60}),
 mk('PLAIN-SHORT-2',{months:[15,15,15,15,15,15,15,15,15,15,15],free:1,price:70,lt:60,sup:'ספק ב'}),
 /* פריטים שדורשים תיקון פרמטרים — כדי שדלת «תיקוני פרמטרים» לא תהיה
    ריקה. רמת שירות 0 וזמן אספקה 0 הם שני המסלולים הישירים ל-paramFix. */
 mk('PARAM-SRV',{months:[30,30,30,30,30,30,30,30,30,30,30],free:900,rop:1,ss:0,price:60,lt:120,srv:0}),
 mk('PARAM-LT',{months:[25,25,25,25,25,25,25,25,25,25,25],free:800,rop:1,ss:0,price:55,lt:0}),
 /* ============ הדלי הבוער ============
    BURN-PARTIAL הוא המקרה שהפיל את הדלת הקיימת: לקוח ממתין ל-19 יח׳,
    על המדף יש 1, ואין שום הזמנת רכש. custWaiting דורש free<=0 ולכן
    אינו רואה אותו — 15 מתוך 16 בדוח האמיתי נראים בדיוק כך. */
 mk('BURN-PARTIAL',{months:[3,3,3,3,3,3,3,3,3,3,3],free:1,cust:19,po:0,price:300,lt:60}),
 /* אותו מצב בדיוק, אבל יש רכש פתוח — זה כבר לא «ללא רכש» */
 mk('COVERED-BY-PO',{months:[3,3,3,3,3,3,3,3,3,3,3],free:1,cust:10,po:4,price:300,lt:60}),
 /* רקע שקט */
 mk('QUIET',{months:[1,1,1,1,1,1,1,1,1,1,1],free:500,price:20})];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['ZMRP'],[],hdr,...rows]),'ZMRP');return wb})(),
 SD+'/routes.xlsx');

const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');

(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const mkctx=async()=>{const ctx=await b.newContext({viewport:{width:1512,height:900}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  return ctx};
 const load=async(p,file)=>{await p.setInputFiles('#f',file);
  await p.waitForSelector('#tbl tbody tr[data-i]',{timeout:120000});await p.waitForTimeout(700)};

 /* ================= 1 · המונה, הניסוח והרשימה ================= */
 {const ctx=await mkctx();const p=await ctx.newPage();
  const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  await load(p,SD+'/routes.xlsx');

  /* --- קודם כול: המקרה עצמו קיים בדוח המבחן --- */
  const z=await p.evaluate(()=>{const r=ALL.find(x=>x.pn==='ZERO-RATE-CUST');
    return r?{rate:+(r.rate||0).toFixed(3),cust:r.cust,free:r.free,
      inShort:decisionList('short').some(x=>x.pn==='ZERO-RATE-CUST'),
      inLine:(()=>{LINE_M=null;const m=lineModel();
        return [...m.dry,...m.stuck,...m.recur,...m.noPo].some(x=>x.pn==='ZERO-RATE-CUST')})()}:null});
  ok('פריט עם לקוח ממתין וללא צריכה נספר בחוסרים',z&&z.inShort,
    z?`rate=${z.rate} cust=${z.cust} free=${z.free}`:'לא נמצא');
  ok('ואינו קיים באף אחת מרשימות הציר — זה היה הפער',z&&!z.inLine,
    z?`inLine=${z.inLine}`:'');

  /* --- מצב קודם ביעד: תחנה/פילוח שנבחרו קודם, וקבוצה מקופלת --- */
  for(const sel of ['stuck','recur','now','nopo']){
   await p.click('#tabs .tab[data-m="catalog"]');await p.waitForTimeout(500);
   await p.evaluate(s=>{LINE_SEL=s;GCOLL.add('t1');GCOLL.add('t2');groupsSave()},sel);
   await p.evaluate(()=>render());await p.waitForTimeout(400);
   const link=await p.evaluate(()=>({vis:!document.getElementById('shortlink').hidden,
     txt:(document.getElementById('shortlinkTx').textContent||'').replace(/\s+/g,' '),
     pns:decisionList('short').map(r=>r.pn).sort()}));
   if(sel==='stuck'){
    ok('הקישור מוצג בקטלוג',link.vis,link.txt);
    ok('והניסוח אומר במפורש מי נספר',
      /לקוח ממתין ללא מלאי/.test(link.txt)&&/לא תסופק החודש/.test(link.txt),link.txt);}
   await p.click('#shortlinkGo');await p.waitForTimeout(700);
   const tg=await p.evaluate(()=>{
     const rows=currentRows();
     const dom=[...document.querySelectorAll('#tbl tbody tr[data-i]')]
       .map(tr=>(rows[+tr.getAttribute('data-i')]||{}).pn).filter(Boolean);
     return {mode,pns:rows.map(r=>r.pn).sort(),dom:dom.sort()}});
   const same=link.pns.length===tg.pns.length&&link.pns.every((x,i)=>x===tg.pns[i]);
   const domSame=link.pns.length===tg.dom.length&&link.pns.every((x,i)=>x===tg.dom[i]);
   ok(`אותה אוכלוסייה למרות בחירה קודמת «${sel}»`,same,
     `מונה=${link.pns.length} יעד=${tg.pns.length} mode=${tg.mode}`);
   ok(`וגם השורות שמוצגות בפועל, אחרי קיפול קבוצות «${sel}»`,domSame,
     `מונה=${link.pns.length} מוצג=${tg.dom.length}`);
   const miss=link.pns.filter(x=>!tg.pns.includes(x));
   ok(`אף מק״ט שנספר אינו נופל בדרך «${sel}»`,miss.length===0,miss.slice(0,4).join(' · '));
  }
  ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
  await ctx.close()}

 /* ================= 5 · «לקוח ממתין ללא רכש» — מסלול משלו =================
    נמדד על דוח ההדגמה לפני: 16 פריטים, ואף מצב שבו currentRows() הוא
    הדלי הזה — כלומר doExport() לא יכול היה לייצא אותם, והדרך היחידה
    הייתה הקטלוג (900) פלוס סינון ידני. הדלת «לקוחות ממתינים» היא
    cust>0 && free<=0 ופספסה 15 מתוך 16. */
 {const ctx=await mkctx();const p=await ctx.newPage();
  const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  await load(p,SD+'/routes.xlsx');

  const f=await p.evaluate(()=>{const g=pn=>ALL.find(x=>x.pn===pn);
    const B=g('BURN-PARTIAL'),C=g('COVERED-BY-PO');
    return {bCat:B.cat,bBurn:!!B.burn,bFree:B.free,bCust:B.cust,bMiss:B.miss,
      cCat:C.cat,cBurn:!!C.burn,cPo:C.po,
      inWaiting:QF.waiting.some(r=>r.pn==='BURN-PARTIAL'),
      inCustDoor:(QF.all||[]).filter(r=>r.cust>0&&r.free<=0).some(r=>r.pn==='BURN-PARTIAL')}});
  ok('לקוח ממתין עם שארית מלאי ובלי רכש מסווג כבוער',
    f.bCat==='לקוח ממתין – אין רכש'&&f.bBurn&&f.inWaiting,
    `cat="${f.bCat}" burn=${f.bBurn} לקוח=${f.bCust} מדף=${f.bFree} חסר=${f.bMiss}`);
  ok('והדלת «לקוחות ממתינים» אינה רואה אותו — היא דורשת מדף אפס',
    !f.inCustDoor,`free=${f.bFree}`);
  ok('פריט עם רכש פתוח אינו נחשב «ללא רכש»',
    !f.cBurn&&f.cCat!=='לקוח ממתין – אין רכש',`cat="${f.cCat}" רכש=${f.cPo}`);

  /* --- הכניסה: ראשונה, גלויה, כפתור --- */
  const nav=await p.evaluate(()=>{const t=document.getElementById('tabs');
    const e=t.querySelector('.tab[data-m="burn"]');
    return {first:t.children[0]&&t.children[0].dataset.m,
      vis:!!e&&e.offsetParent!==null,tag:e?e.tagName:'',
      name:e?((e.querySelector('.t')||{}).textContent||''):'' ,
      bdg:e?((e.querySelector('.bdg')||{}).textContent||''):''}});
  ok('«לקוח ממתין ללא רכש» הוא הכניסה הראשונה בסרגל',nav.first==='burn',nav.first);
  ok('והיא גלויה, כפתור, עם שם ומונה',
    nav.vis&&nav.tag==='BUTTON'&&/לקוח ממתין ללא רכש/.test(nav.name)&&nav.bdg.length>0,
    `${nav.tag} "${nav.name}" מונה=${nav.bdg}`);

  /* --- לחיצה אחת מגיעה בדיוק לדלי, והייצוא הוא בדיוק הוא --- */
  await p.click('#tabs .tab[data-m="burn"]');await p.waitForTimeout(700);
  const r=await p.evaluate(()=>{
    const W=QF.waiting.map(x=>x.pn).sort(),c=currentRows().map(x=>x.pn).sort();
    return {mode,n:c.length,
      exact:W.length===c.length&&W.every((x,i)=>x===c[i]),
      dom:document.querySelectorAll('#tbl tbody tr[data-i]').length,
      exp:exportRows(currentRows()).length-1}});
  ok('לחיצה אחת מהניווט פותחת בדיוק את הדלי הבוער',
    r.mode==='burn'&&r.exact,`mode=${r.mode} n=${r.n}`);
  ok('ומה שמוצג בטבלה הוא אותו דבר',r.dom===r.n,`${r.dom} מתוך ${r.n}`);
  ok('והייצוא משם הוא בדיוק הפריטים האלה — בלי לעבור בקטלוג',
    r.exp===r.n,`${r.exp} שורות בייצוא מול ${r.n} פריטים`);

  /* --- מקלדת --- */
  await p.click('#tabs .tab[data-m="catalog"]');await p.waitForTimeout(450);
  await p.evaluate(()=>document.querySelector('#tabs .tab[data-m="burn"]').focus());
  await p.keyboard.press('Enter');await p.waitForTimeout(600);
  ok('ואפשר להגיע לשם גם במקלדת',await p.evaluate(()=>mode==='burn'));

  /* --- תעדוף: הבוער ראשון גם ברשימות שהוא חלק מהן --- */
  const pri=await p.evaluate(()=>{
    const pos=list=>{const i=list.findIndex(r=>r.burn);return i<0?null:i+1};
    LINE_M=null;const m=lineModel();
    return {short:pos(decisionList('short')),stuck:pos(m.stuck),dry:pos(m.dry)}});
  ok('ובכל רשימה שהוא חלק ממנה הוא בשורה הראשונה',
    [pri.short,pri.stuck,pri.dry].every(v=>v===null||v===1),
    `short=${pri.short} stuck=${pri.stuck} dry=${pri.dry}`);
  ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
  await ctx.close()}

 /* ================= 4 · נגישות המסלולים, בלחיצות בלבד ================= */
 {const ctx=await mkctx();const p=await ctx.newPage();
  const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  await load(p,SD+'/routes.xlsx');
  const clickable=async id=>await p.evaluate(i=>{const e=document.getElementById(i);
    if(!e||e.hidden)return false;const r=e.getBoundingClientRect();
    if(!(r.width>0&&r.height>0))return false;
    return document.elementFromPoint(r.left+r.width/2,r.top+r.height/2)===e
        || e.contains(document.elementFromPoint(r.left+r.width/2,r.top+r.height/2))},id);

  ok('«חוסרים לטיפול» הוא כפתור גלוי ולחיץ',await clickable('shortBtn'));
  ok('«תיקוני פרמטרים ומעקב אספקה» הוא כפתור גלוי ולחיץ',await clickable('fixBtn'));

  await p.click('#shortBtn');await p.waitForTimeout(600);
  const a=await p.evaluate(()=>({mode,track,n:currentRows().length}));
  ok('לחיצה עליו פותחת את רשימת החוסרים',a.mode==='today'&&a.track==='short',
    `mode=${a.mode} track=${a.track} n=${a.n}`);

  await p.click('#fixBtn');await p.waitForTimeout(600);
  const f=await p.evaluate(()=>({mode,track,n:currentRows().length,
    fix:currentRows().filter(r=>r.paramFix).length,
    ttl:document.getElementById('fixBtn').getAttribute('title')||''}));
  ok('ולחיצה על השני פותחת את רשימת תיקוני הפרמטרים המלאה',
    f.mode==='month'&&f.track==='month',`mode=${f.mode} n=${f.n}`);
  ok('הרשימה שם היא הרשימה המלאה ולא פילוח מתוך האוזלים',
    f.n>=f.fix&&f.fix>0,`${f.fix} תיקונים מתוך ${f.n}`);
  ok('והכותרת מפרקת אותה לתיקונים ולמעקב אספקה',
    /לתיקון פרמטרים ב-SAP/.test(f.ttl)&&/במעקב אספקה/.test(f.ttl),f.ttl);

  /* מקלדת — כפתור תקני, בלי preventDefault שחוסם את ה-click */
  await p.click('#tabs .tab[data-m="catalog"]');await p.waitForTimeout(450);
  await p.evaluate(()=>document.getElementById('shortBtn').focus());
  await p.keyboard.press('Enter');await p.waitForTimeout(600);
  ok('אפשר להפעיל את הדלת במקלדת',
    await p.evaluate(()=>mode==='today'));

  /* הסבר הניווט מצביע על המיקום בפועל */
  const hint=await p.evaluate(()=>{const e=document.getElementById('navhint');
    return e?(e.textContent||'').replace(/\s+/g,' '):''});
  /* הרמז ב-#68 הצביע על «רצועה שמעל הטבלה» — renderTrack() מת ולכן
     היא אינה קיימת. הוא מצביע היום על המקום שבו הכניסות באמת יושבות. */
  ok('הסבר הניווט מצביע על המיקום בפועל ולא על רצועה שאינה קיימת',
    /הכניסות הקטנות בסוף הסרגל/.test(hint)&&!/ברצועה שמעל הטבלה/.test(hint),
    hint.slice(0,170));
  /* ומי שקורא את הרמז מוצא שם באמת את שתי הכניסות */
  ok('ושתי הכניסות אכן יושבות שם',
    await p.evaluate(()=>['shortBtn','fixBtn'].every(i=>{
      const e=document.getElementById(i);
      return !!e&&e.closest('#tabs')&&e.offsetParent!==null})));
  ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
  await ctx.close()}

 /* ============ שם, לא סמל, ובלי לגנוב גובה מהרשימה ============
    שתי כניסות נוספות שברו את שורת הניווט לשתיים מתחת ל-1280px — 25px
    שנלקחים בדיוק מהרשימה. fitTabs מוריד קודם את הסכום הכספי שליד
    «הקטלוג» (הוא חוזר בתוך הקטלוג עצמו) ורק אחר כך את השמות. */
 {const ctx=await mkctx();
  for(const W of [1366,1920]){
   const p2=await ctx.newPage();await p2.setViewportSize({width:W,height:W===1366?768:1080});
   await p2.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
   await p2.evaluate(()=>localStorage.clear());await p2.reload();await p2.waitForTimeout(400);
   await load(p2,SD+'/routes.xlsx');
   for(const m of ['line','today','month','catalog']){
    await p2.evaluate(k=>setMode(k),m);await p2.waitForTimeout(350);
    const r=await p2.evaluate(()=>{const t=document.getElementById('tabs');
     const mini=[...t.querySelectorAll('.tab.mini .t')];
     return {lines:new Set([...t.children].map(e=>Math.round(e.getBoundingClientRect().top))).size,
       named:mini.filter(e=>e.offsetParent!==null).length}});
    ok(`${W} · ${m} · שורת ניווט אחת, ושתי הכניסות נושאות שם`,
      r.lines===1&&r.named===2,`שורות=${r.lines} · עם שם=${r.named}`);}
   await p2.close();}
  await ctx.close()}

 /* ================= 3 · העלאה חוזרת — כותרת וכרטיסים ================= */
 {const ctx=await mkctx();const p=await ctx.newPage();
  const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html'));   /* בלי nobrief */
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  const snap=async()=>await p.evaluate(()=>({
    kicker:(document.getElementById('briefKicker').textContent||'').trim(),
    why:(document.getElementById('briefSub').textContent||'').trim(),
    base:document.getElementById('briefBase').hidden?''
      :(document.getElementById('briefBase').textContent||'').trim(),
    cards:[...document.querySelectorAll('#briefCards .bcard')]
      .map(c=>c.querySelector('.bl').textContent.trim()+'='+c.querySelector('.bv').textContent.trim()),
    runs:HIST.runs.length,repeat:!!(HRUN&&HRUN.repeat)}));

  /* --- א׳ --- */
  await load(p,SD+'/routes.xlsx');
  const s1=await snap();
  ok('בהעלאה הראשונה נאמר שאין בסיס להשוואה',
    /אין עדיין דוח קודם/.test(s1.why),s1.why);
  ok('ואין כרטיסי שינוי',
    s1.cards.every(c=>!/^חדשים|^החמירו|^יצאו/.test(c)),s1.cards.join(' · '));

  /* --- ב׳: פריט חדש בתור והחמרה ---
     runSig נגזר מ-ANCHOR, ממספר הפריטים ומסכום הצריכה — לא מהמלאי.
     שינוי מלאי בלבד היה נקרא «אותו דוח שוב», ולכן ב׳ משנה גם צריכה. */
  const rows2=rows.map(r=>r.slice());
  const iFree=hdr.indexOf('מלאי פנוי'), iTot=hdr.indexOf('סה"כ מלאי'), iM1=hdr.indexOf('צר.חודש-1');
  const q=rows2.find(r=>r[0]==='QUIET'); q[iFree]=0; q[iTot]=0; q[iM1]=40;  /* נכנס לתור */
  const ps=rows2.find(r=>r[0]==='PLAIN-SHORT'); ps[iFree]=0; ps[iTot]=0;    /* החמיר */
  XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
   XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['ZMRP'],[],hdr,...rows2]),'ZMRP');return wb})(),
   SD+'/routes-b.xlsx');
  await load(p,SD+'/routes-b.xlsx');
  const s2=await snap();
  ok('בהעלאה השנייה יש בסיס והכרטיסים מתארים אותו',
    s2.runs===2&&!s2.repeat&&!!s2.base,`runs=${s2.runs} base="${s2.base}"`);
  const numOf=(a,l)=>{const h=a.find(c=>c.indexOf(l+'=')===0);return h?+h.split('=')[1]:null};
  const newN=numOf(s2.cards,'חדשים לטיפול'), wN=numOf(s2.cards,'החמירו');

  /* --- ב׳ שוב: טעינה חוזרת ---
     אותו תוכן בדיוק בשם אחר. setInputFiles על אותו נתיב אינו מפעיל
     change שוב, וזה גם המקרה האמיתי: הדוח נשמר מחדש ונטען. הזיהוי
     נשען על תוכן (runSig) ולא על שם הקובץ. */
  fs.copyFileSync(SD+'/routes-b.xlsx',SD+'/routes-b2.xlsx');
  await load(p,SD+'/routes-b2.xlsx');
  const s3=await snap();
  ok('טעינה חוזרת מזוהה ככזו',s3.repeat,`repeat=${s3.repeat} runs=${s3.runs}`);
  ok('והכותרת אומרת שמוצגת ההשוואה האחרונה בין שני דוחות שונים',
    /טעינה חוזרת/.test(s3.why)&&/ההשוואה האחרונה בין שני דוחות שונים/.test(s3.why),s3.why);
  ok('ולא נאמר יותר «אין שינוי להשוות אליו» מעל כרטיסים שמציגים שינוי',
    !/אין שינוי להשוות אליו/.test(s3.why),s3.why);
  ok('הקידומת מתאימה למשפט',s3.kicker==='ההשוואה האחרונה',s3.kicker);
  ok('מועדי הדוחות המושווים מצוינים',
    /דוח \d+\/\d{4}/.test(s3.why)||/נטען/.test(s3.why),s3.why);
  ok('הכותרת והכרטיסים מתארים אותו בסיס',
    !!s3.base&&s3.base===s2.base,`"${s3.base}" מול "${s2.base}"`);
  ok('והמספרים נשמרים כהשוואת א׳ מול ב׳',
    numOf(s3.cards,'חדשים לטיפול')===newN&&numOf(s3.cards,'החמירו')===wN,
    `${s2.cards.join(' · ')}  →  ${s3.cards.join(' · ')}`);
  ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
  await ctx.close()}

 console.log(out.join('\n'));
 await b.close();
 process.exit(bad?1:0)})();
