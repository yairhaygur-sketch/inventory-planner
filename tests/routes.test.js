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
 mk('QUIET',{months:[1,1,1,1,1,1,1,1,1,1,1],free:500,price:20}),
 /* ============ CAP_SPLIT · TWO_CLOCKS — פריטים לארבע הידיות ============
    הקובץ הזה נבנה לפני שהדלת פוצלה, וההון הכלוא שבו היה שני פריטי
    עודף בלבד — כלומר שלוש מארבע הידיות היו ריקות ולא נבדקו כלל,
    ושורת «שני שעונים» לא הייתה מופיעה על אף פריט. הפריטים כאן נבנו
    כל אחד למצב אחר; לאן המנוע מסווג אותם בפועל נמדד בבדיקה עצמה. */
 /* רכש פתוח שעוד נכנס למלאי שכבר עודף */
 mk('CAP-PO',{months:[1,1,1,1,1,1,1,1,1,1,1],free:60,po:10,rop:3,ss:1,ssMin:1,price:200,lt:60,saleAgo:40,entAgo:120}),
 /* נקודת הזמנה גבוהה בהרבה מהנדרש — יש מה להקטין */
 mk('CAP-STOP',{months:[1,1,1,1,1,1,1,1,1,1,1],free:60,rop:50,ss:20,price:200,lt:30,saleAgo:40,entAgo:120}),
 /* מלאי מת: אין צריכה בשלוש שנים, ROP כבוי */
 mk('CAP-DEAD',{months:[0,0,0,0,0,0,0,0,0,0,0],y0:0,y1:0,y2:0,free:40,rop:0,ss:0,price:50,saleAgo:1400,entAgo:1400}),
 /* שני שעונים, צורת «יש קצב מדיד»: 40 חודשי כיסוי על המדף,
    וזמן אספקה של 4 חודשים שדורש ROP גבוה בהרבה מהקיים */
 mk('CLOCKS-RATED',{months:[5,5,5,5,5,5,5,5,5,5,5],free:200,rop:3,ss:1,ssMin:2,price:150,lt:120,saleAgo:30,entAgo:150}),
 /* שני שעונים, צורת «אין קצב מדיד»: הצריכה כולה מחוץ לחלון 11
    החודשים, ולכן ה-ROP המוצע הוא רצפת ה-SS בלבד */
 mk('CLOCKS-FLAT',{months:[0,0,0,0,0,0,0,0,0,0,0],y0:1,y1:1,y2:1,free:40,rop:3,ss:2,ssMin:4,price:80,saleAgo:300,entAgo:400}),
 /* ============ SLOW_PO ============
    מלאי איטי שיש לו רכש פתוח. הענף האיטי לא הזכיר אותו כלל, ובקובץ
    הזה לא היה אף פריט כזה — כלומר חמש בדיקות היו רצות על רשימה ריקה.
    שניהם נבנו למצב אחר: באחד יש יותר בדרך פנימה מההקטנה שתוצע,
    ובשני אין בכלל מה להקטין על המדף. */
 mk('SLOW-PO-OVER',{months:[0,0,0,0,0,0,0,0,0,0,0],y0:1,y1:1,y2:1,free:40,po:60,rop:3,ss:2,ssMin:4,price:80,saleAgo:300,entAgo:400}),
 mk('SLOW-PO-NONE',{months:[0,0,0,0,0,0,0,0,0,0,0],y0:1,y1:1,y2:1,free:3,po:10,rop:3,ss:2,ssMin:4,price:80,saleAgo:300,entAgo:400}),
 /* ============ FOLLOW_ROP ============
    «מעקב אספקה» עם נקודת הזמנה נמוכה מהנדרש — המצב של 6608440675,
    שהמעתד הגיע איתו לחוסר. בקובץ הזה לא היה **אף** פריט «מעקב
    אספקה», כלומר הבדיקה הייתה רצה על רשימה ריקה.
    קצב 2 · LT 120 יום (ltM=4) · מדף 5 (כיסוי 2.5 < 4) · בדרך 10
    (כיסוי כולל 7.5 ≥ 4) → הענף נתפס. ssMin 2 → sugSS 2, ולכן
    sugROP = ⌈2×4 + 2⌉ = 10 מול rop=3: פער 7, ו-2.5 חודשים יבשים.
    והבקרה השלילית: אותו מצב בדיוק עם rop=12, שאין בו פער. */
 mk('FOLLOW-ROP-GAP',{months:[2,2,2,2,2,2,2,2,2,2,2],free:5,po:10,rop:3,ss:2,ssMin:2,price:150,lt:120,saleAgo:20,entAgo:60}),
 mk('FOLLOW-ROP-OK',{months:[2,2,2,2,2,2,2,2,2,2,2],free:5,po:10,rop:12,ss:2,ssMin:2,price:150,lt:120,saleAgo:20,entAgo:60})];
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
 /* הכלי נוחת ב«מרכז עבודה», שהוא סדר יום ולא טבלה. העוזר הזה
    משרת בדיקות טבלה, ולכן הוא נכנס למסך טבלה במפורש. */
 const load=async(p,file)=>{await p.setInputFiles('#f',file);
  await p.evaluate(()=>{try{setMode('today')}catch(_){}});
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

  /* --- הכניסה: ראשונה בתחום החוסרים, גלויה, כפתור ---
     היה: «הכניסה הראשונה בסרגל» — סרגל אופקי של יעדים שווים. במסילה
     האנכית יש שתי רמות, ו«אין רכש כלל» הוא הדלת הראשונה בתוך
     «חוסרים ולקוחות». מה שנשמר הוא מה שההחלטה באמת דרשה: כניסה
     גלויה עם מונה, בלי לעבור דרך הקטלוג ובלי לדעת מראש שהיא קיימת. */
  await p.click('#tabs .tab[data-m="today"]');await p.waitForTimeout(500);
  const nav=await p.evaluate(()=>{const sb=document.getElementById('railsub');
    const e=sb.querySelector('.railsub2[data-m="burn"]');
    const area=document.querySelector('#tabs .tab[data-m="today"]');
    return {first:(SUBNAV.today||[])[0]&&SUBNAV.today[0][0],
      areaFirstVisible:!!area&&area.offsetParent!==null,
      vis:!!e&&e.offsetParent!==null,tag:e?e.tagName:'',
      name:e?((e.querySelector('.t')||{}).textContent||''):'' ,
      bdg:e?((e.querySelector('.bdg')||{}).textContent||''):''}});
  ok('«אין רכש כלל» הוא הדלת הראשונה בתוך «חוסרים ולקוחות»',
    nav.first==='burn'&&nav.areaFirstVisible,nav.first);
  ok('והיא גלויה, כפתור, עם שם ומונה',
    nav.vis&&nav.tag==='BUTTON'&&/אין רכש כלל/.test(nav.name)&&nav.bdg.length>0,
    `${nav.tag} "${nav.name}" מונה=${nav.bdg}`);

  /* --- לחיצה מגיעה בדיוק לדלי, והייצוא הוא בדיוק הוא --- */
  await p.click('#railsub .railsub2[data-m="burn"]');await p.waitForTimeout(700);
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
  await p.evaluate(()=>document.querySelector('#tabs .tab[data-m="today"]').focus());
  await p.keyboard.press('Enter');await p.waitForTimeout(600);
  await p.evaluate(()=>document.querySelector('#railsub .railsub2[data-m="burn"]').focus());
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

  /* היו: shortBtn ו-fixBtn — שתי כניסות משניות (.mini) שנוספו בסוף
     סרגל אופקי, כי «חוסרים לטיפול» ו«תיקוני פרמטרים ומעקב אספקה»
     היו מרונדרים בלי שום דלת. במסילה האנכית שניהם *תחומים ראשיים*
     בזכות עצמם — «חוסרים ולקוחות» ו«תכנון ו-MRP» — ולכן הבדיקה
     מכוונת לשם. היכולת לא ירדה; היא עלתה רמה. */
  const clickableSel=async sel=>await p.evaluate(q=>{const e=document.querySelector(q);
    if(!e||e.hidden)return false;const r=e.getBoundingClientRect();
    if(!(r.width>0&&r.height>0))return false;
    return document.elementFromPoint(r.left+r.width/2,r.top+r.height/2)===e
        || e.contains(document.elementFromPoint(r.left+r.width/2,r.top+r.height/2))},sel);
  ok('«חוסרים ולקוחות» הוא כפתור גלוי ולחיץ',
    await clickableSel('#tabs .tab[data-m="today"]'));
  ok('«תכנון ו-MRP» הוא כפתור גלוי ולחיץ',
    await clickableSel('#tabs .tab[data-m="month"]'));

  await p.click('#tabs .tab[data-m="today"]');await p.waitForTimeout(600);
  const a=await p.evaluate(()=>({mode,track,n:currentRows().length}));
  ok('לחיצה עליו פותחת את רשימת החוסרים',a.mode==='today'&&a.track==='short',
    `mode=${a.mode} track=${a.track} n=${a.n}`);

  await p.click('#tabs .tab[data-m="month"]');await p.waitForTimeout(600);
  const f=await p.evaluate(()=>({mode,track,n:currentRows().length,
    fix:currentRows().filter(r=>r.paramFix).length,
    ttl:document.querySelector('#tabs .tab[data-m="month"]').getAttribute('title')||'',
    sub:(SUBNAV.month||[]).map(x=>x[1]).join(' · ')}));
  ok('ולחיצה על השני פותחת את רשימת תיקוני הפרמטרים המלאה',
    f.mode==='month'&&f.track==='month',`mode=${f.mode} n=${f.n}`);
  ok('הרשימה שם היא הרשימה המלאה ולא פילוח מתוך האוזלים',
    f.n>=f.fix&&f.fix>0,`${f.fix} תיקונים מתוך ${f.n}`);
  /* היה: הכותרת של fixBtn פירקה את הרשימה לשני חלקים בריחוף. במסילה
     הפירוק גלוי בעין — דלתות המשנה של התחום — ולא מוסתר בתוך title. */
  ok('והתחום מפרק את עצמו לדלתות משנה גלויות',
    /תיקוני פרמטרים/.test(f.sub)&&/רצפת SS/.test(f.sub)&&/יושמו ב-SAP/.test(f.sub),f.sub);

  /* מקלדת — כפתור תקני, בלי preventDefault שחוסם את ה-click */
  await p.click('#tabs .tab[data-m="catalog"]');await p.waitForTimeout(450);
  await p.evaluate(()=>document.querySelector('#tabs .tab[data-m="today"]').focus());
  await p.keyboard.press('Enter');await p.waitForTimeout(600);
  ok('אפשר להפעיל את הדלת במקלדת',
    await p.evaluate(()=>mode==='today'));

  /* היה: «הסבר הניווט מצביע על המיקום בפועל» — פסקה שהסבירה איפה
     יושבות שתי הכניסות הקטנות. שתיהן תחומים במסילה, ולכן ההסבר
     הוחלף במה שהוא תיאר: הן שם, בשמן, עם מונה. */
  ok('שני התחומים יושבים במסילה בשמם ועם מונה',
    await p.evaluate(()=>['today','month'].every(k=>{
      const e=document.querySelector(`#tabs .tab[data-m="${k}"]`);
      return !!e&&e.offsetParent!==null&&!!e.querySelector('.t')&&!!e.querySelector('.bdg')})));
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
    /* היה: «שורת ניווט אחת» — סרגל אופקי שנשבר לשתי שורות מתחת
       ל-1280px. המסילה אנכית, ולכן מה שנבדק הוא מה שאותו כלל באמת
       הגן עליו: כל שישה התחומים נושאים שם קריא, אף אחד אינו נחתך
       לסמל בלבד, והמסילה אינה גולשת. */
    const r=await p2.evaluate(()=>{const t=document.getElementById('tabs');
     const tabs=[...t.querySelectorAll('.tab')];
     const rail=document.getElementById('rail');
     return {n:tabs.length,
       named:tabs.filter(e=>{const s=e.querySelector('.t');
         return s&&s.offsetParent!==null&&s.textContent.trim().length>2}).length,
       cols:new Set(tabs.map(e=>Math.round(e.getBoundingClientRect().left))).size,
       ov:rail.scrollWidth-rail.clientWidth}});
    ok(`${W} · ${m} · שבעה תחומים בעמודה אחת, כולם נושאים שם`,
      r.n===7&&r.named===7&&r.cols===1&&r.ov<=1,
      `תחומים=${r.n} · עם שם=${r.named} · עמודות=${r.cols} · גלישה=${r.ov}`);}
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
  /* ============ מה הטענה הזאת באמת בודקת ============
     שהכותרת והכרטיסים מדברים על *אותו זוג דוחות*. היא הושוותה
     כמחרוזת מלאה, וזו כללה את שעת הטעינה — recordRun מרענן את
     last.ts בטעינה חוזרת, ולכן הדקה זזה והבדיקה נפלה כשהריצה חצתה
     גבול דקה. זהו הסייג שתועד ב-PR #71, לא רגרסיה.
     הזהות נבדקת עכשיו לפי הדוחות עצמם (חודש העוגן), שאינם זזים. */
  const pairOf=t=>(String(t).match(/דוח \d+\/\d{4}/g)||[]).join(' מול ');
  ok('הכותרת והכרטיסים מתארים אותו זוג דוחות',
    !!s3.base&&pairOf(s3.base)===pairOf(s2.base)&&pairOf(s3.base)!=='',
    `"${pairOf(s3.base)}" מול "${pairOf(s2.base)}"`);
  ok('והמספרים נשמרים כהשוואת א׳ מול ב׳',
    numOf(s3.cards,'חדשים לטיפול')===newN&&numOf(s3.cards,'החמירו')===wN,
    `${s2.cards.join(' · ')}  →  ${s3.cards.join(' · ')}`);
  ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
  await ctx.close()}

 /* ================= 4 · מרכז העבודה =================
    מסך הנחיתה אינו טבלה חמישית ואינו סיווג חדש: ארבע רשימות שכבר
    קיימות, והקישור אל כל אחת. מה שנעול כאן הוא בדיוק זה — שכל מונה
    בכרטיס הוא אורך הרשימה שהוא מצביע אליה, ולא ספירה משלו. */
 {const ctx=await mkctx();const p=await ctx.newPage();
  const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  await load(p,SD+'/routes.xlsx');
  await p.click('#tabs .tab[data-m="home"]');await p.waitForTimeout(600);
  /* ============ מרכז העבודה הוא רשימה מדורגת, לא ארבעה כרטיסים ============
     היה: ארבעה כרטיסים, כל אחד עם שבע שורות ראשונות מרשימה בת
     817 / 301 / 2,572. פרוסות דקות שלא הסתכמו למונה שבמסילה,
     שניצלו 57% מגובה החלון, ושענו «מה» ולא «באיזה סדר».
     הטענות נשמרו — כל אחת מהן — ונבדקות על המבנה החדש. */
  const h=await p.evaluate(()=>({mode,track,
    rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
    agenda:agendaRows().length,
    badge:navCount('home').n,
    title:(document.getElementById('lt')||{}).textContent||'',
    month:(document.getElementById('rd')||{}).value||'',
    band:!!document.querySelector('.aband'),
    chips:[...document.querySelectorAll('.abchip')].map(c=>({
      go:c.dataset.go,
      n:+(c.querySelector('b').textContent.replace(/[^\d]/g,'')),
      t:c.textContent})),
    eng:{burn:burnRows().length,cap:decisionList('cap').length,
      line:navCount('line').n,qual:navCount('qual').n,
      month:modeRows('month').filter(r=>r.paramFix&&(r.sugROP||0)>(r.rop||0)).length,
      stuck:(QF.all||[]).filter(r=>expStuck(r)>0).length},
    /* הבוערים חייבים לפתוח את הרשימה — ההבטחה ללקוח כבר ניתנה */
    firstAreBurn:(()=>{const b=new Set(burnRows());
      return agendaRows().slice(0,burnRows().length).every(r=>b.has(r))})(),
    gaps:(document.querySelector('.abgap')||{}).textContent||'',
    noPx:(QF.all||[]).filter(r=>r.priceMissing).length,
    hasEta:!!ETA}));
  ok('«מרכז עבודה» הוא רשימת סדר היום, לא תור העבודה הכללי',
    h.mode==='home'&&h.track==='agenda'&&h.rows===h.agenda,
    `mode=${h.mode} track=${h.track} · ${h.rows} שורות מתוך ${h.agenda}`);
  ok('המונה במסילה הוא אורך הרשימה עצמה',
    h.badge===h.agenda,`תג ${h.badge} · רשימה ${h.agenda}`);
  ok('והוא אומר את שמו ואת חודש הדוח',
    /מרכז עבודה/.test(h.title)&&/^\d{4}-\d{2}$/.test(h.month),
    `${h.title} · ${h.month}`);
  /* «לקוח ממתין ללא רכש» ראשון — לא כי הוא גדול, אלא כי הוא
     היחיד שבו ההבטחה כבר ניתנה ואין מולה כלום. */
  ok('הבוערים פותחים את הרשימה',h.firstAreBurn,
    `${h.eng.burn} בוערים בראש ${h.agenda}`);
  ok('רצועה עם ארבעה שבבים, כל אחד דלת לתחום',
    h.band&&h.chips.length===4&&h.chips.map(c=>c.go).join(',')==='line,month,qual,cap',
    h.chips.map(c=>c.go).join(' · '));
  /* הטענה המרכזית שנשמרה מהמבנה הישן: כל מספר הוא אורך הרשימה
     שהוא מצביע אליה, ולא ספירה משלו. */
  ok('כל שבב הוא אורך הרשימה שהוא מצביע אליה',
    h.chips[0].n===h.eng.line&&h.chips[1].n===h.eng.month
    &&h.chips[2].n===h.eng.qual&&h.chips[3].n===h.eng.cap,
    h.chips.map(c=>`${c.go} ${c.n}`).join(' · '));
  ok('מונה הרכש בלי תאריך הוא הספירה האמיתית, גם בלי דוח ETA',
    h.eng.stuck===0||h.gaps.includes(h.eng.stuck.toLocaleString('he-IL'))
      ||h.gaps.includes(String(h.eng.stuck)),
    `ברצועה "${h.gaps.slice(0,80)}" · בפועל ${h.eng.stuck} · ETA=${h.hasEta}`);
  ok('ופערי הנתונים נאמרים מעל הרשימה',
    (!h.noPx||h.gaps.includes(String(h.noPx)))&&(h.hasEta||/דוח ETA/.test(h.gaps)),
    h.gaps.slice(0,110));
  await p.click('.abchip[data-go="cap"]');await p.waitForTimeout(600);
  ok('שבב מגיע לתחום הנכון',
    'cap'===await p.evaluate(()=>mode),await p.evaluate(()=>mode));
  await p.click('#tabs .tab[data-m="home"]');await p.waitForTimeout(600);
  const pn=await p.evaluate(()=>{const tr=document.querySelector('#tbl tbody tr[data-i]');
    if(!tr)return null;const t=tr.children[1].textContent.replace(/העתק|✓ טופל/g,'').trim();
    tr.click();return t});
  await p.waitForTimeout(500);
  ok('לחיצה על שורה פותחת את כרטיס הפריט שלה',
    !pn||pn===await p.evaluate(()=>{const e=document.querySelector('#detail .opnt');
      return e?e.textContent.trim():null}),pn);
  ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
  await ctx.close()}

 /* ================= 5 · בריאות המלאי =================
    התחום היה קיים במסילה, אבל הגוף שלו היה בדיוק הקטלוג: cur='all',
    900 פריטים, ועמודת «הון כלוא» ריקה ברוב השורות הגלויות. */
 {const ctx=await mkctx();const p=await ctx.newPage();
  const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  await load(p,SD+'/routes.xlsx');
  await p.click('#tabs .tab[data-m="cap"]');await p.waitForTimeout(700);
  const c=await p.evaluate(()=>{
   const el=document.getElementById('capTop');
   const {list,B}=capBuckets();
   const sup=capBySupplier(list);
   const money=[...el.querySelectorAll('.cpanel h4 span')].map(x=>x.textContent);
   return {mode,cur,track,
     rows:currentRows().length,eng:list.length,
     catalog:(QF.all||[]).length,
     shown:document.querySelectorAll('#tbl tbody tr[data-i]').length,
     tiles:[...el.querySelectorAll('.ctile')].map(t=>({
       t:t.querySelector('.ct').textContent,n:+t.querySelector('.cn').textContent.replace(/[^\d]/g,'')})),
     eb:{excess:B.excess.length,slow:B.slow.length,dead:B.dead.length,other:B.other.length},
     panels:sup.length,money,
     curs:[...new Set(list.map(r=>r.currency||'—'))].length,
     note:(el.querySelector('.cnote')||{}).textContent||'',
     capCol:(()=>{const i=[...document.querySelectorAll('#tbl thead th')]
        .findIndex(t=>/הון כלוא/.test(t.textContent));
       if(i<0)return null;
       const tds=[...document.querySelectorAll('#tbl tbody tr[data-i]')].slice(0,10)
         .map(tr=>(tr.children[i].textContent||'').trim());
       return tds.filter(x=>x&&x!=='—').length})()}});
  ok('הטבלה היא רשימת ההון הכלוא ולא הקטלוג',
    c.rows===c.eng&&c.eng<c.catalog,`${c.rows} שורות · ${c.eng} עם הון כלוא · ${c.catalog} בקטלוג`);
  ok('ולכל שורה גלויה יש באמת הון כלוא',c.capCol===null||c.capCol>=Math.min(10,c.shown),
    `${c.capCol} מתוך ${Math.min(10,c.shown)} הראשונות`);
  ok('ארבעה אריחים, והמספרים הם הפילוח של המנוע',
    c.tiles.length===4&&c.tiles[0].n===c.eng
    &&c.tiles[1].n===c.eb.excess&&c.tiles[2].n===c.eb.slow&&c.tiles[3].n===c.eb.dead,
    c.tiles.map(t=>`${t.t}=${t.n}`).join(' · '));
  /* «רקע» אינו נעלם: אם רוב הפריטים בעלי ההון הכלוא אינם באף אחד
     משלושת הדליים, המסך אומר זאת ולא מציג שלושה מספרים כאילו הם
     כל הסיפור. */
  ok('ומה שאינו באף דלי נאמר במפורש',
    !c.eb.other||(c.note.includes(String(c.eb.other))&&/אינם עודף/.test(c.note)),
    c.note.slice(0,90)||'(אין «רקע»)');
  ok('לוח לכל מטבע — ואין סכום אחד חוצה מטבעות',
    c.panels===c.curs&&c.money.length===c.curs,
    `${c.panels} לוחות · ${c.curs} מטבעות · ${c.money.join(' | ')}`);
  /* ============ CAP_SPLIT — ארבע הידיות ============
     הדלת הכוללת מסודרת לפי סיווג: מלאי מת · איטי · עודף · רקע. הסיווג
     אומר למה הכסף תקוע, לא מה אפשר לעשות. נמדד על הדוח של 5.10
     (2,700 בדלת): 1,651 מתוך 2,335 פריטי מת/איטי כבר מחזיקים ROP ≤ 1,
     כלומר ההוראה «לאפס ROP/SS» מופנית למי שכבר כבוי; ו-277 פריטים
     מחזיקים ROP > 1 על מלאי עודף אבל sugROP ≥ rop וגם sugSS ≥ ss —
     הון כלוא בלי שום מספר להציע.

     הבדיקות כאן נועלות שלושה דברים: שכל אחת מארבע הדלתות מכילה אך
     ורק את מי שה-`capAct` שלו הוא היא, שארבעתן יחד הן **בדיוק**
     הדלת הכוללת (אף פריט לא נעלם, אף פריט לא נספר פעמיים), ושכל
     דלת אומרת למי שנכנס מה הידית — כולל «אין מה לשנות», שקיימת
     כדי שלא יתחפש לרשימת עבודה. */
  const cs=await p.evaluate(async()=>{
   const tot=capPick(decisionList('cap'),'cap');
   const K=['capPO','capStop','capBiz','capNone'];
   const part={};for(const k of K)part[k]=capPick(decisionList('cap'),k);
   /* חלוקה אמיתית: הסך נשמר, ואף פריט אינו בשתי דלתות */
   const seen=new Set();let dup=0;
   for(const k of K)for(const r of part[k]){if(seen.has(r))dup++;seen.add(r)}
   return {tot:tot.length,n:Object.fromEntries(K.map(k=>[k,part[k].length])),
    sum:K.reduce((a,k)=>a+part[k].length,0),dup,
    nav:Object.fromEntries(['cap',...K].map(k=>[k,navCount(k).n])),
    /* כל דלת והתנאי שהיא מתיימרת לקיים — נבדק מול השדות של המנוע */
    poClean:part.capPO.filter(r=>!(r.po>0)).length,
    stopClean:part.capStop.filter(r=>!(!(r.po>0)&&!r.floorClash&&r.rop>1
       &&((r.sugROP||0)<(r.rop||0)||(r.sugSS||0)<(r.ss||0)))).length,
    bizClean:part.capBiz.filter(r=>!(!(r.po>0)&&(r.floorClash
       ||(!(r.rop>1)&&CAP_DEAD.has(r.cat))))).length,
    /* «אין מה לשנות» חייבת להיות באמת כזו: או ROP כבוי ולא מת,
       או ROP חי שהמנוע אינו מציע להקטין */
    noneHasLever:part.capNone.filter(r=>r.po>0
       ||((r.sugROP||0)<(r.rop||0)||(r.sugSS||0)<(r.ss||0))&&r.rop>1&&!r.floorClash).length,
    /* ההון הכלוא חיובי בכל שורה בכל דלת — ההגדרה של הדלת הכוללת */
    noCap:K.reduce((a,k)=>a+part[k].filter(r=>!(r.expCap>0)).length,0)}});
  ok('ארבע הידיות יחד הן בדיוק הדלת הכוללת — אף פריט לא נעלם',
    cs.sum===cs.tot&&cs.dup===0,
    `${Object.values(cs.n).join(' + ')} = ${cs.sum} מול ${cs.tot} · ${cs.dup} כפולים`);
  ok('והמונה במסילה מסכים עם הרשימה בכל אחת מהן',
    ['capPO','capStop','capBiz','capNone','cap'].every(k=>cs.nav[k]===(k==='cap'?cs.tot:cs.n[k])),
    JSON.stringify(cs.nav));
  ok('«רכש בדרך לעודף» — לכל פריט בה יש רכש פתוח',cs.poClean===0,cs.poClean+' בלי רכש פתוח');
  ok('«להפסיק להזמין» — ROP חי והמנוע עצמו מציע מספר נמוך ממנו',
    cs.stopClean===0,cs.stopClean+' חריגות');
  ok('«החלטה מסחרית» — מת עם ROP כבוי, או התנגשות מדיניות',
    cs.bizClean===0,cs.bizClean+' חריגות');
  ok('«אין מה לשנות» — ואין בה אף פריט שיש לו ידית',
    cs.noneHasLever===0,cs.noneHasLever+' פריטים עם ידית');
  ok('ובכל ארבעתן ההון הכלוא חיובי — זו ההגדרה של הדלת',cs.noCap===0,cs.noCap+' עם אפס');
  /* ============ TWO_CLOCKS ============
     277 פריטים הציגו «הון כלוא» ו«אל תקטין / העלה את ה-ROP» זה לצד
     זה בלי הסבר, כי העודף נמדד בחודשי כיסוי וה-ROP בחודשי אספקה.
     נוספה שורת נימוק בלבד — אף מספר לא השתנה. הבדיקה נועלת שלושה
     דברים: שהיא מופיעה בדיוק על מי שעומד בתנאי, שהיא לעולם אינה
     מופיעה על פריט שיש לו הקטנה להציע (שם היא הייתה שקר), ושהיא
     אומרת את שני השעונים בשתי הצורות הנכונות — ביקוש בזמן אספקה
     כשיש קצב מדיד, רצפת SS כשאין. */
  const tc=await p.evaluate(()=>{
   const all=QF.all||[];
   const TXT=r=>((r.why||[]).find(w=>w&&/שני שעונים/.test(w[1]||''))||[])[1]||null;
   const has=all.filter(TXT);
   const should=all.filter(r=>(r.cat==='עודף מלאי'||r.cat==='מלאי איטי')&&r.rop>ROP_MIN_ACT
     &&(r.sugROP||0)>=(r.rop||0)&&(r.sugSS||0)>=(r.ss||0));
   return {n:has.length,should:should.length,
    /* אף פריט שיש לו הקטנה להציע, או שה-ROP שלו כבוי */
    wrong:has.filter(r=>(r.sugROP||0)<(r.rop||0)||(r.sugSS||0)<(r.ss||0)
       ||!(r.rop>ROP_MIN_ACT)).length,
    /* הצורה חייבת להתאים לסיבה שמחזיקה את ה-ROP */
    /* הנוסח נוקב בזמן האספקה בימים, לא בחודשים: ltM הוא ltD/30 ויצא
       4.066666666666666 על הדוח האמיתי — מכפלה שהקורא יבדוק ולא תסתדר. */
    badRated:has.filter(r=>r.rate>0&&r.cov!=null
      &&!/ביקוש בזמן אספקה [\d.]+ יח׳ \(קצב [\d.]+ לחודש · זמן אספקה \d+ ימים\)/.test(TXT(r))).length,
    badFlat:has.filter(r=>!(r.rate>0&&r.cov!=null)&&!/אין קצב מדיד/.test(TXT(r))).length,
    /* «1 חודשי אספקה» אינו עברית */
    plural:has.filter(r=>/\b1 חודשי/.test(TXT(r))).length,
    /* המספרים בשורה הם של המנוע, לא מחושבים מחדש */
    numOk:has.filter(r=>r.rate>0&&r.cov!=null)
      .every(r=>TXT(r).includes(`ROP מוצע ${r.sugROP} =`)
        &&TXT(r).includes(`מלאי ביטחון ${r.sugSS} `)),
    sample:has.length?TXT(has.sort((a,b)=>(b.expCap||0)-(a.expCap||0))[0]):''}});
  ok('«שני שעונים» מופיעה בדיוק על מי שעומד בתנאי',
    tc.n===tc.should&&tc.n>0,`${tc.n} קיבלו · ${tc.should} עומדים בתנאי`);
  ok('ולעולם לא על פריט שיש לו הקטנה להציע, או ש-ROP שלו כבוי',
    tc.wrong===0,tc.wrong+' חריגות');
  ok('הנוסח תואם לסיבה שמחזיקה את ה-ROP — ביקוש בזמן אספקה או רצפה',
    tc.badRated===0&&tc.badFlat===0,`${tc.badRated} עם קצב · ${tc.badFlat} בלי`);
  ok('והמספרים בשורה הם של המנוע עצמו',tc.numOk,tc.sample.slice(0,150));
  ok('ואין «1 חודשי»',tc.plural===0,tc.plural+' מופעים');
  /* ============ SEVIR — «סביר» הייתה אמירה, לא היעדר אמירה ============
     ענף העודף בדק רק את כיוון ההקטנה: `rop > sugROP ? «להקטין» :
     «נקודת הזמנה סבירה»`. כשהמנוע המליץ דווקא להעלות, ה-else הצהיר
     שהפרמטר תקין — בניגוד למספר שהמנוע חישב שתי שורות קודם.

     נמדד על הדוח של 5.10: 84 פריטי עודף — 66 «נקודת הזמנה סבירה»
     בעוד sugROP > rop (פער 165 יח׳), ו-28 «מלאי ביטחון סביר» בעוד
     sugSS > ss (פער 52 יח׳). 1300-03-01134 החזיק SS 10 מול 20 מוצע
     ואמר «סביר». $101,822 · ₪3,908 · €30 הון כלוא על אותן שורות.

     הבדיקה אינה מחפשת את המילה «סביר» — היא דורשת שהטקסט יסכים עם
     המספרים של המנוע בשני הכיוונים. ניסוח חדש שישקר ייתפס גם הוא. */
  const sv=await p.evaluate(()=>{
   const all=QF.all||[];const A=r=>(r.act||[]).join(' || ');
   const ex=all.filter(r=>r.cat==='עודף מלאי');
   const bad=ex.filter(r=>{const t=A(r);
     return (/נקודת הזמנה תואמת/.test(t)&&r.sugROP!==r.rop)
         || (/מלאי ביטחון תואם/.test(t)&&r.sugSS!==r.ss)
         || (r.sugROP>r.rop&&!/להעלות נקודת הזמנה/.test(t))
         || (r.sugSS>r.ss&&!/להעלות מלאי ביטחון/.test(t))
         || (r.sugROP<r.rop&&!/להקטין נקודת הזמנה/.test(t))
         || (r.sugSS<r.ss&&!/להקטין מלאי ביטחון/.test(t))});
   return {ex:ex.length,bad:bad.length,
    raise:ex.filter(r=>r.sugROP>r.rop||r.sugSS>r.ss).length,
    /* ואף הצהרת «סביר» לא שרדה בשום מקום בקובץ */
    leftover:all.filter(r=>/נקודת הזמנה סבירה|מלאי ביטחון סביר/.test(A(r))).length,
    smp:bad.length?`${bad[0].pn} ROP ${bad[0].rop}→${bad[0].sugROP} SS ${bad[0].ss}→${bad[0].sugSS} :: ${A(bad[0]).slice(0,120)}`:''}});
  ok('ב«עודף מלאי» הטקסט מסכים עם המספרים של המנוע בשני הכיוונים',
    sv.ex>0&&sv.bad===0,`${sv.bad} סתירות מתוך ${sv.ex} · ${sv.smp}`);
  ok('ויש בקובץ פריט עודף שהמנוע ממליץ להעלות בו פרמטר — אחרת הבדיקה ריקה',
    sv.raise>0,sv.raise+' פריטים');
  ok('ואף הצהרת «סביר» לא שרדה',sv.leftover===0,sv.leftover+' מופעים');
  /* ============ SLOW_PO — הענף האיטי לא קרא את הרכש הפתוח ============
     ענף העודף בודק `po` בנימוק ובפעולה; הענף האיטי לא הזכיר אותו כלל.
     נמדד על הדוח של 5.10: 1,208 פריטי «מלאי איטי», 91 עם רכש פתוח,
     847 יח׳ בדרך פנימה ($47,459 · ₪221) — וכל 91 שתקו עליו.
     77 מתוך 91 (85%) מחזיקים po ≥ over: 8892021569 — 190 על המדף,
     143 בדרך, והכלי אמר «לבחון הקטנת מלאי ב-37 יח׳». מי שמבצע בלי
     לטפל ברכש מגדיל את המלאי.

     שלושה מצבים ושלושה משפטים, כי «ביצוע ההמלצה יגדיל את המלאי» נכון
     רק כשיש המלצה לבצע — ב-over=0 הפעולה היא «לעקוב». */
  const sp=await p.evaluate(()=>{
   const all=QF.all||[];
   const A=r=>(r.act||[]).join(' || '),W=r=>(r.why||[]).map(w=>w&&w[1]||'').join(' || ');
   const slow=all.filter(r=>r.cat==='מלאי איטי'),po=slow.filter(r=>r.po>0);
   return {slow:slow.length,po:po.length,
    /* אף פריט איטי עם רכש פתוח אינו שותק עליו */
    silent:po.filter(r=>!/רכש פתוח/.test(A(r)+' || '+W(r))).length,
    /* והפעולה הראשונה היא ביטול הרכש, עם הכמות הנכונה */
    firstAct:po.filter(r=>(r.act||[])[0]!==`לבדוק ביטול רכש פתוח (${r.po} יח׳)`).length,
    /* «ביצוע ההמלצה» לא נאמר כשאין המלצה לבצע */
    badPhrase:po.filter(r=>!(r.excessQty>0)&&/ביצוע ההמלצה/.test(W(r))).length,
    /* וכשיש יותר בדרך מההקטנה — זה נאמר במפורש */
    geSilent:po.filter(r=>r.excessQty>0&&r.po>=r.excessQty
      &&!/יותר מ-\d+ היחידות שמוצע להקטין/.test(W(r))).length,
    /* פריט איטי בלי רכש פתוח אינו מקבל את השורה */
    noPoLeak:slow.filter(r=>!(r.po>0)&&/רכש פתוח/.test(A(r)+' || '+W(r))).length}});
  ok('אף פריט «מלאי איטי» עם רכש פתוח אינו שותק עליו',
    sp.po>0&&sp.silent===0,`${sp.silent} שתקו מתוך ${sp.po} (מתוך ${sp.slow} איטיים)`);
  ok('והפעולה הראשונה היא ביטול הרכש, עם הכמות שבדוח',sp.firstAct===0,sp.firstAct+' חריגות');
  ok('«ביצוע ההמלצה» לא נאמר כשאין המלצה לבצע',sp.badPhrase===0,sp.badPhrase+' מופעים');
  ok('וכשיש יותר יחידות בדרך מההקטנה המוצעת — זה נאמר',sp.geSilent===0,sp.geSilent+' שתקו');
  ok('ופריט איטי בלי רכש פתוח אינו מקבל את השורה',sp.noPoLeak===0,sp.noPoLeak+' דליפות');
  /* ============ FOLLOW_ROP — הרכש מכסה מחזור אחד, ה-ROP מנהל את כולם ====
     שוחזר מ-6608440675, פריט שהמעתד הגיע איתו לחוסר. בספטמבר היו לו
     5 על המדף, 7 בדרך, ROP=3 — ו-sugROP=13. בקצב 2.33 וזמן אספקה
     120 יום, הזמנה שנפתחת ב-3 משאירה 2.7 חודשים יבשים. ספטמבר 5 →
     אוקטובר 0, בדיוק לפי החשבון.

     והכלי ידע: sugROP=13 ישב בדוח של ספטמבר. אבל מה שהוא אמר היה
     «אין צורך ברכש יזום נוסף · לא לפתוח רכש נוסף בשלב זה» — אף מילה
     על נקודת ההזמנה. ענף «מניעת חוסר» אומר «נקודת הזמנה X נמוכה
     מהנדרש Y»; ענף «מעקב אספקה», שלוש שורות מעליו, שתק.

     נמדד על הדוח של 5.10: 65 מתוך 138 פריטי «מעקב אספקה» מחזיקים
     sugROP > rop, ו-64 שתקו. */
  const fr=await p.evaluate(()=>{
   const all=QF.all||[];
   const LN=r=>(r.why||[]).map(w=>w&&w[1]||'').find(t=>/נמוכה מהנדרש/.test(t))||null;
   const fol=all.filter(r=>r.cat==='מעקב אספקה');
   const gap=fol.filter(r=>(r.sugROP||0)>(r.rop||0));
   const nogap=fol.filter(r=>!((r.sugROP||0)>(r.rop||0)));
   return {fol:fol.length,gap:gap.length,nogap:nogap.length,
    silent:gap.filter(r=>!LN(r)).length,
    leak:nogap.filter(r=>LN(r)).length,
    /* המספרים בשורה הם של המנוע, לא מחושבים מחדש */
    nums:gap.every(r=>LN(r).includes(`נקודת הזמנה ${r.rop} נמוכה מהנדרש ${r.sugROP}`)),
    /* זמן היובש נאמר רק כשיש קצב מדיד והוא חיובי */
    dryBad:gap.filter(r=>{const has=/משאירה כ-[\d.]+ חודשים בלי מלאי/.test(LN(r));
      const want=r.rate>0&&r.ltM>0&&(r.ltM-(r.rop/r.rate))>0;return has!==want}).length,
    smp:gap.length?LN(gap[0]):''}});
  ok('יש בקובץ «מעקב אספקה» עם פער ובלי — אחרת הבדיקה ריקה',
    fr.gap>0&&fr.nogap>0,`${fr.gap} עם פער · ${fr.nogap} בלי · מתוך ${fr.fol}`);
  ok('כל פריט «מעקב אספקה» שה-ROP שלו נמוך מהנדרש אומר זאת',
    fr.silent===0,fr.silent+' שתקו');
  ok('ופריט שאין בו פער אינו מקבל את השורה',fr.leak===0,fr.leak+' דליפות');
  ok('המספרים בשורה הם של המנוע עצמו',fr.nums,fr.smp.slice(0,110));
  ok('וזמן היובש נאמר בדיוק כשיש קצב מדיד והוא חיובי',fr.dryBad===0,fr.dryBad+' חריגות');
  /* ============ ROP_BRIDGE — הדלת שנולדה מהחוסר ============
     השאלה: האם נקודת ההזמנה מגשרת על הביקוש בזמן האספקה. ב-6608440675
     היא לא — ROP 3 מול ביקוש 9.33, ולכן 2.7 חודשים יבש בכל מחזור.

     המדידה הראשונה שלי נתנה 25 חשופים. 14 מהם היו «רקע» עם ROP 0→0:
     פריטי 04/14 ו«לפי דרישה» שהמנוע **בכוונה** אינו נותן להם המלצה,
     ושאינם מנוהלים ב-ROP כלל. אחרי ההחרגה: 634 בתחום, **7 חשופים**,
     246 עם כרית דקה. ההחרגה הזאת היא הדבר העיקרי שהבדיקה נועלת —
     בלעדיה הדלת מציגה פי שלושה פריטים שאין בהם מה לעשות. */
  const rb=await p.evaluate(()=>{
   const all=QF.all||[];
   const gap=ropPick(ropAll(),'ropGap'),thin=ropPick(ropAll(),'ropThin');
   const inSet=new Set([...gap,...thin]);
   return {scope:all.filter(ropScope).length,gap:gap.length,thin:thin.length,
    /* חלוקה: אין חפיפה, ושתיהן בתוך התחום */
    overlap:gap.filter(r=>thin.includes(r)).length,
    outOfScope:[...inSet].filter(r=>!ropScope(r)).length,
    /* מי שהמנוע אינו ממליץ לו — לא נכנס. זו הטעות שכמעט שלחתי */
    noRec:[...inSet].filter(r=>r.isOD||r.noStock||isPD(r.mrp)
      ||(!(r.sugROP>0)&&!(r.sugSS>0))).length,
    /* ההגדרה של כל דלת, מול שדות המנוע */
    gapBad:gap.filter(r=>!(ropDry(r)>0)).length,
    thinBad:thin.filter(r=>ropDry(r)>0||!((r.sugROP||0)>(r.rop||0))).length,
    /* ropDry הוא בדיוק הנוסחה, לא קירוב */
    formula:[...inSet].every(r=>Math.abs(ropDry(r)-(r.ltM-(r.rop/r.rate)))<1e-9),
    /* ואין פריט חשוף שנשאר מחוץ לדלת */
    missed:all.filter(r=>ropScope(r)&&ropDry(r)>0&&!gap.includes(r)).length}});
  ok('יש בקובץ פריטים בשתי דלתות ה-ROP — אחרת הבדיקה ריקה',
    rb.gap>0&&rb.thin>0,`${rb.gap} לא מגשר · ${rb.thin} מתחת למוצע · מתוך ${rb.scope} בתחום`);
  ok('שתי הדלתות אינן חופפות ואף פריט חשוף לא נשאר בחוץ',
    rb.overlap===0&&rb.missed===0,`${rb.overlap} חפיפות · ${rb.missed} הוחמצו`);
  ok('כל פריט בדלתות נמצא בתחום — מנוהל מלאי, לא PD, ויש לו המלצה',
    rb.outOfScope===0&&rb.noRec===0,`${rb.outOfScope} מחוץ לתחום · ${rb.noRec} בלי המלצה`);
  ok('«ROP לא מגשר» מכילה אך ורק פריטים שהיובש שלהם חיובי',rb.gapBad===0,rb.gapBad+' חריגות');
  ok('ו«מתחת למוצע» אך ורק פער בלי יובש',rb.thinBad===0,rb.thinBad+' חריגות');
  ok('וחישוב היובש הוא בדיוק זמן האספקה פחות ROP÷קצב',rb.formula);
  /* המסך: הגרוע ביותר קודם, והעמודה שמכריעה מציגה מספר */
  await p.evaluate(()=>setMode('ropGap'));await p.waitForTimeout(700);
  const rv=await p.evaluate(()=>{
   const view=currentRows();
   const tr=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
   const HX=n=>[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim()).indexOf(n);
   const i=HX('בלי מלאי');
   const shown=tr.map(t=>view[+t.dataset.i]);
   return {rows:tr.length,hasCol:i>=0,
    /* ממוין לפי חודשי היובש, הגרוע קודם */
    sorted:shown.every((r,k)=>!k||ropDry(shown[k-1])>=ropDry(r)-1e-9),
    /* ובכל שורה התא מראה מספר, לא «—» */
    cells:i<0?null:tr.filter(t=>/[\d.]+ חו׳/.test(t.children[i].textContent)).length,
    over:(()=>{const t=document.getElementById('tbl');return t.scrollWidth-t.clientWidth>2})(),
    sub:document.getElementById('phdSub').textContent,
    pgR:document.getElementById('pgR').textContent}});
  ok('עמודת «בלי מלאי» קיימת ומציגה מספר בכל שורה',
    rv.hasCol&&rv.cells===rv.rows,`${rv.cells}/${rv.rows}`);
  ok('והדלת ממוינת כך שהגרוע ביותר ראשון',rv.sorted);
  ok('התת-כותרת אומרת מה ההגדרה',/נמוכה מהביקוש בזמן האספקה/.test(rv.sub),rv.sub.slice(0,90));
  ok('והסיכום אומר כמה כבר בלי מלאי',/כבר בלי מלאי/.test(rv.pgR),rv.pgR.slice(0,90));
  ok('ואין גלישה אופקית',!rv.over);
  /* ============ XDOOR — אותו פריט, כמה דלתות ============
     נמדד על הדוח של 5.10: שתים עשרה דלתות העבודה מחזיקות 5,660 שורות
     אבל רק 4,096 פריטים שונים — 1,360 (33%) יושבים בשתיים או יותר,
     ובכלי לא היה שום סימן לכך. 465 מהם נושאים «להקטין» מול רצפה
     מאושרת שמחזיקה למעלה, ו-149 «להקטין פרמטר אחד ולהעלות את השני».

     הבדיקות כאן נועלות את מה שחייב להיות נכון כדי שהסימן לא ישקר:
     שהמפה מסכימה עם הדלתות *בשני הכיוונים*, שהדלת שאתה עומד בה
     אינה נספרת כ«גם ב», שהסימן מופיע בדיוק על השורות שיש להן דלת
     נוספת, ושהקפיצה אכן מגיעה לדלת ומשאירה את הפריט פתוח. */
  const xd=await p.evaluate(()=>{
   xBuild();
   const lists={};for(const [k] of XDOORS)lists[k]=xRows(k);
   /* כיוון א: כל זוג (פריט, דלת) במפה באמת נמצא באותה דלת */
   let ghost=0;
   for(const [r,ks] of XMAP)for(const k of ks)if(!lists[k].includes(r))ghost++;
   /* כיוון ב: כל פריט בכל דלת נמצא במפה תחת אותה דלת */
   let missed=0;
   for(const k in lists)for(const r of lists[k]){
     const a=XMAP.get(r);if(!a||a.indexOf(k)<0)missed++}
   const all=QF.all||[];
   return {ghost,missed,mapped:XMAP.size,
    multi:all.filter(r=>xDoors(r).length>1).length,
    /* אף פריט אינו רשום פעמיים תחת אותה דלת */
    dup:[...XMAP.values()].filter(a=>a.length!==new Set(a).size).length}});
  ok('מפת הדלתות מסכימה עם הדלתות עצמן — בשני הכיוונים',
    xd.ghost===0&&xd.missed===0,`${xd.ghost} רפאים · ${xd.missed} חסרים`);
  ok('ואף פריט אינו רשום פעמיים באותה דלת',xd.dup===0,xd.dup+' כפולים');
  ok('ויש בקובץ פריטים שיושבים ביותר מדלת אחת — אחרת הבדיקה ריקה',
    xd.multi>0,`${xd.multi} מתוך ${xd.mapped}`);
  /* הסימן על המסך — בדיוק על השורות שיש להן דלת נוספת */
  await p.evaluate(()=>setMode('capNone'));await p.waitForTimeout(700);
  const xb=await p.evaluate(()=>{
   const view=currentRows();
   const tr=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
   const shown=tr.map(t=>view[+t.dataset.i]);
   return {rows:tr.length,
    want:shown.filter(r=>xOther(r).length).length,
    got:document.querySelectorAll('#tbl tbody .xd').length,
    /* התאמה שורה-לשורה, לא רק סכום */
    match:tr.every((t,i)=>!!t.querySelector('.xd')===(xOther(shown[i]).length>0)),
    /* המספר על הסימן הוא מספר הדלתות הנוספות. הוא יושב ב-data-n
       ומרונדר דרך CSS, כדי שלא ייכנס ל-textContent של תא המק״ט. */
    num:tr.every((t,i)=>{const e=t.querySelector('.xd');
      return !e||e.dataset.n===String(xOther(shown[i]).length)}),
    /* והמק״ט עצמו נשאר נקי — מי שמעתיק אותו ל-SAP לא גורר סימן */
    clean:tr.every((t,i)=>{const o=t.querySelector('.obj');
      return o&&!/⧉/.test(t.textContent||'')&&o.textContent.trim()===String(shown[i].pn)}),
    /* והדלת שאני עומד בה אינה נספרת */
    self:shown.filter(r=>xOther(r).includes(mode)).length,
    title:(document.querySelector('#tbl tbody .xd')||{}).title||''}});
  ok('הסימן «גם ב» מופיע בדיוק על השורות שיש להן דלת נוספת',
    xb.match&&xb.got===xb.want,`${xb.got}/${xb.want} מתוך ${xb.rows} שורות`);
  ok('והמספר עליו הוא מספר הדלתות הנוספות',xb.num);
  ok('ותא המק״ט נשאר נקי — הסימן אינו נכנס לטקסט שמעתיקים',xb.clean);
  /* ============ והוא אינו עולה רוחב ============
     בגרסה הראשונה הסימן היה inline והוסיף 34px לעמודת המק״ט
     (325→359), גזל 6px מהתיאור, והפך את הכרעת הדחיפה ב-1920 —
     שתי בדיקות ממשק נפלו. מיקום מוחלט הוא *המנגנון* שמחזיק את
     העלות באפס, ולכן הוא נבדק ישירות ולא רק דרך תוצאותיו. */
  const xpos=await p.evaluate(()=>{const e=document.querySelector('#tbl tbody .xd');
    if(!e)return null;const cs=getComputedStyle(e);
    const td=e.closest('td');
    return {pos:cs.position,tdPos:td?getComputedStyle(td).position:null}});
  ok('הסימן ממוקם מוחלט ואינו משתתף ברוחב העמודה',
    xpos&&xpos.pos==='absolute'&&xpos.tdPos==='relative',
    xpos?`${xpos.pos} בתוך td ${xpos.tdPos}`:'(אין סימן)');
  ok('הדלת שאתה עומד בה אינה נספרת כ«גם ב»',xb.self===0,xb.self+' הפניות עצמיות');
  ok('והכיתוב נוקב בשמות הדלתות',/נמצא גם ב: .+/.test(xb.title),xb.title.slice(0,90));
  /* הכרטיס, והקפיצה */
  const xj=await p.evaluate(async()=>{
   const r=(QF.all||[]).find(x=>xOther(x).length>0&&currentRows().includes(x));
   if(!r)return null;detail(r);
   await new Promise(s=>setTimeout(s,400));
   const btns=[...document.querySelectorAll('#detail .xgo')];
   return {pn:r.pn,from:mode,
     labels:btns.map(b=>b.textContent.trim()),
     want:xOther(r).map(k=>XNAME[k]),
     note:!!document.querySelector('#detail .dcxn')}});
  ok('הכרטיס מונה את אותן דלתות בדיוק, בשמותיהן',
    xj&&xj.labels.length>0&&xj.labels.join('|')===xj.want.join('|'),
    xj?`${xj.labels.join(' · ')} מול ${xj.want.join(' · ')}`:'(לא נמצא פריט)');
  ok('ויש סייג שאומר שההוראה אינה בהכרח זהה',!!(xj&&xj.note));
  const xjump=await p.evaluate(async()=>{
   const b=document.querySelector('#detail .xgo');if(!b)return null;
   const want=b.dataset.go;b.click();
   await new Promise(s=>setTimeout(s,600));
   return {want,mode,open:!!document.querySelector('#detail'),
     pn:((document.querySelector('#detail .opnt')||{}).textContent||'').trim()}});
  ok('לחיצה על דלת בכרטיס מגיעה אליה — והפריט נשאר פתוח',
    xjump&&xjump.mode===xjump.want&&xjump.open&&xjump.pn===xj.pn,
    xjump?`${xjump.mode} (ביקשתי ${xjump.want}) · ${xjump.pn}`:'(אין כפתור)');
  /* המפה נבנית מחדש בכל render — אחרת היא תשקר אחרי סינון */
  /* ההרעלה היא הדרך היחידה להוכיח איפוס: render בונה את המפה מחדש
     תוך כדי הרינדור, ולכן «XMAP===null» אחריו תמיד שקרי. */
  const xstale=await p.evaluate(async()=>{
   const pn=[...XMAP].find(([r,k])=>k.length>1)[0].pn;
   XMAP=new Map();                      /* מפה מורעלת — ריקה */
   const poisoned=xDoors((QF.all||[]).find(r=>r.pn===pn)).length;
   render();await new Promise(s=>setTimeout(s,400));
   const after=xDoors((QF.all||[]).find(r=>r.pn===pn)).length;
   const lists={};for(const [k] of XDOORS)lists[k]=xRows(k);
   let ghost=0;for(const [r,ks] of XMAP)for(const k of ks)if(!lists[k].includes(r))ghost++;
   return {pn,poisoned,after,ghost}});
  ok('render זורק מפה מורעלת ובונה אותה מחדש',
    xstale.poisoned===0&&xstale.after>1,
    `${xstale.pn}: ${xstale.poisoned} אחרי הרעלה · ${xstale.after} אחרי render`);
  ok('והמפה החדשה עדיין מסכימה עם הדלתות',xstale.ghost===0,xstale.ghost+' רפאים');
  /* הדלת חייבת לדבר. «אין מה לשנות» היא המקרה המסוכן: רשימה שנראית
     כמו עבודה ואינה. */
  const cd=[];
  for(const k of ['capPO','capStop','capBiz','capNone']){
   await p.evaluate(m=>setMode(m),k);await p.waitForTimeout(500);
   cd.push(await p.evaluate(()=>({m:mode,
     rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
     n:currentRows().length,
     ttl:(document.querySelector('.phd .title')||{}).textContent||'',
     sub:document.getElementById('phdSub').textContent||'',
     pgR:document.getElementById('pgR').textContent||'',
     over:(()=>{const t=document.getElementById('tbl');return t.scrollWidth>t.clientWidth+2})()})))}
  ok('לכל ידית כותרת משלה, תת-כותרת שאומרת מה הידית, ושורות שתואמות את הרשימה',
    cd.every(d=>d.ttl.includes('·')&&d.sub.length>40&&(d.n===0||d.rows>0)),
    cd.map(d=>`${d.m}: ${d.rows}/${d.n} «${d.sub.slice(0,30)}»`).join(' | '));
  ok('«אין מה לשנות» אומרת במפורש שאין מה לשנות',
    (d=>/אין מולו שום מספר להציע|ROP כבר כבוי/.test(d.sub))(cd[3]),cd[3].sub.slice(0,120));
  /* בקובץ ההדגמה הדלת הזאת יכולה לצאת ריקה — הסיכום נבדק כשיש בה מה לסכם */
  ok('«רכש בדרך לעודף» מסכמת את הכמות שעוד לא נכנסה',
    cd[0].n===0||/יח׳ רכש פתוח עוד בדרך פנימה/.test(cd[0].pgR),
    `${cd[0].n} פריטים · ${cd[0].pgR.slice(0,110)}`);
  ok('ואף אחת מהן אינה גולשת אופקית',cd.every(d=>!d.over),
    cd.filter(d=>d.over).map(d=>d.m).join(',')||'אין גלישה');
  ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
  await ctx.close()}

 console.log(out.join('\n'));
 await b.close();
 process.exit(bad?1:0)})();
