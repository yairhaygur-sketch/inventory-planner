/* דוח ה-ETA — טעינה נפרדת, צירוף לפי מק״ט, ופיצול "בדרך" לשניים.
   כל מקרה כאן נבדק מול הדוחות האמיתיים לפני שנכתב. ראה ETA_SPLIT. */
const XLSX=require('xlsx'),{chromium}=require('playwright'),fs=require('fs'),path=require('path');
const SD=__dirname, MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const hdr=['מק"ט מוביל','תיאור חומר','תיאור חומר2','שם ספק','סטטוס חומר','תיאור','סוג MRP','ABC','רמת שרות',
 'מלאי בטחון','נק.הז.מחדש','אספ.מתוכנ.','זמ.עב.קבלת','מל.בט.מינ.','מחיר FOB','מטבע FOB','סה"כ מלאי','מלאי פנוי',
 'מלאי מרלוג','הז. רכש','בהעברה','אספקות פת.','כמות בהז.פ','סוג חומר','תיא.קבוצ.חומרים','קב.חו.חיצו','טקסט ארוך',
 'תב.אח.הש.','ת.היר.1 מח','ת.היר.2 מח','היררכייה1','היררכייה1','היררכייה2','היררכייה2','היררכייה3','היררכייה3',
 'צר.השנה','צר.שנה-1','צר.שנה-2','צר.החודש',...MON,"תאר' מכירה","תאר' כניסה"];
const D=n=>{const d=new Date(Date.now()-n*864e5);
 return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`};
const mk=(pn,o)=>{const m=o.months,y=m.reduce((a,b)=>a+b,0);
 return [pn,'פריט מבחן '+pn,'Test '+pn,'ספק','01','פעיל','ND','A',95,
  0,0,o.lt||30,2,0,o.price!=null?o.price:100,'USD',o.free,o.free,0,o.po||0,0,0,o.cust||0,'Z004','מנוע','ZT','מתכנן','0',
  'שיווק','מערכת','100','ZT','200','דגם','300','מערכת',y,y,y,0,...m,D(10),D(60)]};

/* מק״טים עטופים בסימני כיוון RTL, בדיוק כמו בשני הדוחות האמיתיים.
   בלי הסרתם אף מק״ט לא מתחבר — וזו הייתה הבדיקה הראשונה שנכשלה. */
const RLM=s=>'‎'+s+'‏';
const rows=[
 // רכש 18, מתוכם 5 משובצות — הדגם של 8894765000 בדוח האמיתי
 mk('ETA-PARTIAL',{months:[4,4,4,4,4,4,4,4,4,4,4],free:0,po:18,cust:12,price:1200}),
 // כל הרכש משובץ — "תקוע" חייב להיות אפס
 mk('ETA-FULL',{months:[3,3,3,3,3,3,3,3,3,3,3],free:0,po:10,cust:2,price:300}),
 // מופיע בדוח, אבל התאריך כבר עבר. משלוח שאיחר אינו כיסוי.
 mk('ETA-PAST',{months:[3,3,3,3,3,3,3,3,3,3,3],free:0,po:12,cust:1,price:400}),
 // רכש פתוח שאינו מופיע בדוח ה-ETA כלל — הכול תקוע
 mk('ETA-NONE',{months:[3,3,3,3,3,3,3,3,3,3,3],free:0,po:20,cust:3,price:500}),
 // אספקה גדולה מהרכש (13 מקרים בדוח האמיתי) — תקוע אפס, לא שלילי
 mk('ETA-OVER',{months:[3,3,3,3,3,3,3,3,3,3,3],free:0,po:4,cust:1,price:200}),
 // בלי רכש פתוח בכלל
 mk('ETA-NOPO',{months:[3,3,3,3,3,3,3,3,3,3,3],free:0,po:0,cust:1,price:600}),
 /* ETA_COVER · שכבה 2 — אפס לקוח ממתין, ולכן «כסף שלא יסופק» ולא כשל
    שירות. קצב 10, מדף 2 ⇒ 8 יח׳ לא יסופקו לפי המדף בלבד. */
 mk('T2-COVERED',{months:Array(11).fill(10),free:2,po:20,cust:0,price:700}),
 mk('T2-LATE',{months:Array(11).fill(10),free:2,po:20,cust:0,price:700}),
 mk('T2-PARTIAL',{months:Array(11).fill(10),free:2,po:20,cust:0,price:700}),
 /* שכבה 1 עם אספקה בתוך החודש — חייב להישאר, ובאותו מספר בדיוק */
 mk('T1-WITH-ETA',{months:Array(11).fill(10),free:0,po:20,cust:4,price:700})];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['ZMRP'],[],hdr,...rows]),'ZMRP');return wb})(),
 SD+'/eta-zmrp.xlsx');

const dplus=n=>{const d=new Date(Date.now()+n*864e5);return new Date(d.getFullYear(),d.getMonth(),d.getDate())};
/* חודש הבסיס הוא החודש המלא האחרון, ולכן «סוף החודש» הוא סוף החודש
   הנוכחי. שני התאריכים האלה נכונים בכל יום שבו הבדיקה תרוץ — dplus(7)
   היה חוצה את סוף החודש אם היא תרוץ ב-25 בחודש. */
const nowD=new Date();
const inMonth=new Date(nowD.getFullYear(),nowD.getMonth()+1,0);   // היום האחרון של החודש
const afterMonth=new Date(nowD.getFullYear(),nowD.getMonth()+2,15); // אמצע החודש הבא
const ehdr=['אספקה','פריט','חומר','תיאור','כמות באספקה',"א'",'תארי.אספקה'];
const erows=[
 ['4190000001','000010',RLM('ETA-PARTIAL'),'PARTIAL',2,'EA',dplus(7)],
 ['4190000001','000020',RLM('ETA-PARTIAL'),'PARTIAL',3,'EA',dplus(21)],
 ['4190000002','000010',RLM('ETA-FULL'),'FULL',10,'EA',dplus(14)],
 ['4190000003','000010',RLM('ETA-PAST'),'PAST',12,'EA',dplus(-30)],
 ['4190000004','000010',RLM('ETA-OVER'),'OVER',9,'EA',dplus(10)],
 ['4190000005','000010',RLM('NOT-IN-CATALOGUE'),'GHOST',5,'EA',dplus(10)],
 // שורה בלי תאריך — מוחרגת, ואסור שתפיל את הקריאה
 ['4190000006','000010',RLM('ETA-NONE'),'NODATE',7,'EA',''],
 // ETA_COVER: מכסה במלואו בתוך החודש ⇒ יורד מהרשימה
 ['4190000007','000010',RLM('T2-COVERED'),'T2COV',8,'EA',inMonth],
 // אותה כמות בדיוק, אבל אחרי סוף החודש ⇒ אינו מכסה כלום
 ['4190000008','000010',RLM('T2-LATE'),'T2LATE',8,'EA',afterMonth],
 // מכסה חלקית ⇒ נשאר ברשימה, עם מספר קטן יותר
 ['4190000009','000010',RLM('T2-PARTIAL'),'T2PART',3,'EA',inMonth],
 /* משלוח שני, אחרי סוף החודש: כך «משובץ» תמיד גדול מ«מכוסה» ואפשר
    לנעול שהתא מציג את מה שמכסה החודש ולא את כל העתיד. */
 ['4190000011','000020',RLM('T2-PARTIAL'),'T2PART',4,'EA',afterMonth],
 // שכבה 1 עם אספקה בתוך החודש ⇒ חייב להישאר, בלי שינוי במספר
 ['4190000010','000010',RLM('T1-WITH-ETA'),'T1ETA',20,'EA',inMonth]];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([ehdr,...erows]),'גיליון1');return wb})(),
 SD+'/eta-report.xlsx');

/* ============ LEDGER_SEED — דוח ETA היסטורי ============
   דוח עם «מסמך רכש» ועם שורות שתאריכן כבר עבר. הפנקס הרגיל דורש שני
   ימים שמקיפים את התאריך; כאן הרכש הפתוח של היום משמש כ«אחרי».
   שלושת המקרים בנויים במפורש מול הרכש הפתוח שב-eta-zmrp:
     ETA-FULL    po=10 · עתיד 10 · עבר 25  ->  po==עתיד          -> נחת
     ETA-PAST    po=12 · עתיד  0 · עבר 12  ->  po==עתיד+עבר      -> איחור
     ETA-PARTIAL po=18 · עתיד  5 · עבר  7  ->  לא זה ולא זה      -> לא מוסבר
   ו-T1-WITH-ETA נשאר עתידי בלבד, ולכן אינו נספר כלל. */
const hhdr=['אספקה','פריט','חומר','תיאור','מסמך רכש','כמות באספקה',"א'",'תארי.אספקה'];
const hrows=[
 ['4190100001','000010',RLM('ETA-FULL'),'FULL','4500100001',10,'EA',dplus(14)],
 ['4190100002','000010',RLM('ETA-FULL'),'FULL','4500100002',15,'EA',dplus(-40)],
 ['4190100003','000010',RLM('ETA-FULL'),'FULL','4500100002',10,'EA',dplus(-70)],
 ['4190100004','000010',RLM('ETA-PAST'),'PAST','4500100003',12,'EA',dplus(-30)],
 ['4190100005','000010',RLM('ETA-PARTIAL'),'PARTIAL','4500100004',5,'EA',dplus(21)],
 ['4190100006','000010',RLM('ETA-PARTIAL'),'PARTIAL','4500100005',7,'EA',dplus(-55)],
 ['4190100007','000010',RLM('T1-WITH-ETA'),'T1ETA','4500100006',20,'EA',inMonth]];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([hhdr,...hrows]),'גיליון1');return wb})(),
 SD+'/eta-hist.xlsx');

const snap=()=>{const g=pn=>{const r=ALL.find(x=>x.pn===pn);if(!r)return null;
  return {po:r.po,eta:r.etaQty,past:r.etaPast,first:r.etaFirst,stuck:r.stuck,has:r.etaHas,
    month:r.etaMonth,gross:r.unsGross,uns:r.unsQty,val:Math.round(r.unsVal||0),t1:isT1(r),
    inList:decisionList('short').some(x=>x.pn===r.pn)}};
 return {P:g('ETA-PARTIAL'),F:g('ETA-FULL'),T:g('ETA-PAST'),N:g('ETA-NONE'),
  O:g('ETA-OVER'),Z:g('ETA-NOPO'),
  loaded:!!ETA,parts:ETA?ETA.parts:0,rows:ETA?ETA.rows:0,units:ETA?ETA.units:0,undated:ETA?ETA.undated:0,
  hit:ALL.filter(x=>x.etaHas).length,
  cols:[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim()),
  nSched:document.querySelectorAll('#tbl tbody td.sched').length,
  nOtw:document.querySelectorAll('#tbl tbody td.otw').length,
  chipHidden:document.getElementById('etachip').hidden,
  chip:document.getElementById('etachip').textContent,
  C:g('T2-COVERED'),L:g('T2-LATE'),PA:g('T2-PARTIAL'),W:g('T1-WITH-ETA'),
  eom:typeof etaEom==='function'?etaEom():null,
  covered:typeof etaCoveredRows==='function'?etaCoveredRows().length:0,
  badge:(document.querySelector('.gcov')||{}).textContent||'',
  listN:decisionList('short').length,
  t1N:decisionList('short').filter(isT1).length,
  /* ETA_HONEST — הסריקה שתפסה 30 מקומות אחרי ש"תיקנתי הכול".
     כל פריט שיש לו אספקה משובצת: אף מחרוזת שלו — הסבר או פעולה —
     לא רשאית להכריז שאין תאריך או להורות לבקש אחד. */
  liars:(()=>{const BAD=[/אין תאריך אספקה בדוח/,/אין תאריך הגעה/,/אין תאריך בדוח/,
      /^לבקש ETA$/,/לדרוש ETA/,/אין ETA/,/משלוח בלי תאריך אינו כיסוי/];
    const txt=r=>[...(r.why||[]).map(w=>w[1]),...(r.act||[])];
    return ALL.filter(r=>(r.etaQty||0)>0)
      .filter(r=>txt(r).some(t=>BAD.some(re=>re.test(t))))
      .map(r=>r.pn+': '+txt(r).filter(t=>BAD.some(re=>re.test(t))).join(' | '))})(),
  scheduled:ALL.filter(r=>(r.etaQty||0)>0).length,
  /* והכיוון ההפוך: בלי דוח, אסור שמשהו יטען שיש תאריך. */
  ghosts:(()=>{const txt=r=>[...(r.why||[]).map(w=>w[1]),...(r.act||[])];
    return ETA?[]:ALL.filter(r=>txt(r).some(t=>/שובצו לאספקה|משובצות ל-|מגיעות ב-/.test(t)))
      .map(r=>r.pn)})()}};

(async()=>{
 const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');

 // ── שלב 1: ZMRP בלבד. המסך חייב להיראות בדיוק כמו לפני השינוי ──
 await p.setInputFiles('#f',SD+'/eta-zmrp.xlsx');await p.waitForTimeout(1800);
 const before=await p.evaluate(snap);

 /* ── SHELL_BAND ──
    הבדיקה שהייתה תופסת את הבאג מלכתחילה. הגיבור נבנה כך שהוא מרונדר
    רק כשיש דוח ETA: בלי הדוח renderLine החזיר הודעה בת שורה ויצא,
    והטקסט הגדול במסך היה 19px. שום בדיקה לא הסתכלה על המצב הזה, כי
    כולן טוענות ZMRP בלבד ואף אחת לא בדקה מה יש בראש המסך.
    לכן נמדדים כאן *שני* המצבים, ובמפורש. */
 const bandOf=()=>({band:(()=>{const e=document.querySelector('.lband');
     return e?Math.round(e.getBoundingClientRect().height):0})(),
   hero:Math.max(0,...[...document.querySelectorAll('.lband .hn')]
     .map(e=>parseFloat(getComputedStyle(e).fontSize)||0)),
   heroTxt:(document.querySelector('.lband .hn')||{}).textContent||'',
   segs:document.querySelectorAll('.lband .bseg').length});
 const bDry=await p.evaluate(bandOf);
 const gapDry=await p.evaluate(()=>!!document.querySelector('.lband .hgap'));

 // ── שלב 2: דוח ה-ETA נטען בנפרד, אחרי ZMRP ──
 await p.setInputFiles('#fe',SD+'/eta-report.xlsx');await p.waitForTimeout(1200);
 const after=await p.evaluate(snap);
 const bEta=await p.evaluate(bandOf);
 const gapEta=await p.evaluate(()=>!!document.querySelector('.lband .hgap'));

 // ── שלב 3: ZMRP נטען מחדש. הדוח חייב לשרוד — זו כל הסיבה שהוא נשמר ──
 await p.setInputFiles('#f',SD+'/eta-zmrp.xlsx');await p.waitForTimeout(1800);
 const reload=await p.evaluate(snap);

 // ── שלב 4: ההפך — דפדפן נקי, ETA לפני ZMRP ──
 const p2=await ctx.newPage();p2.on('pageerror',e=>errs.push('p2: '+e.message));
 await p2.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p2.evaluate(()=>localStorage.removeItem('planner_eta_v1'));
 await p2.reload();await p2.waitForTimeout(300);
 await p2.setInputFiles('#fe',SD+'/eta-report.xlsx');await p2.waitForTimeout(700);
 await p2.setInputFiles('#f',SD+'/eta-zmrp.xlsx');await p2.waitForTimeout(1800);
 const reverse=await p2.evaluate(snap);

 // ── שלב 5: קובץ שאינו דוח ETA נדחה בלי לשבור כלום ──
 await p2.setInputFiles('#fe',SD+'/eta-zmrp.xlsx');await p2.waitForTimeout(700);
 const wrong=await p2.evaluate(snap);

 // ── שלב 6: הסרת הדוח — הכול חוזר בדיוק לאיפה שהיה ──
 await p2.evaluate(()=>etaClear());await p2.waitForTimeout(600);
 const cleared=await p2.evaluate(snap);

 /* ── שלב 7: LEDGER_SEED — דוח היסטורי מול הרכש הפתוח ── */
 const p3=await ctx.newPage();p3.on('pageerror',e=>errs.push('p3: '+e.message));
 await p3.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p3.evaluate(()=>{localStorage.removeItem('planner_eta_v1');
   localStorage.removeItem('planner_ledger_v1')});
 await p3.reload();await p3.waitForTimeout(300);
 await p3.setInputFiles('#f',SD+'/eta-zmrp.xlsx');await p3.waitForTimeout(1800);
 const seedBefore=await p3.evaluate(()=>({sd:ledgerSeed(ALL),line:lineLedgerLine()}));
 await p3.setInputFiles('#fe',SD+'/eta-hist.xlsx');await p3.waitForTimeout(1500);
 const seed=await p3.evaluate(()=>{const sd=ledgerSeed(ALL);
   const g=pn=>{const r=ALL.find(x=>x.pn===pn)||{};
     return {po:r.po|0,fut:r.etaQty|0,past:r.etaPast|0}};
   return {hasPo:!!(ETA&&ETA.hasPo),
     sd:sd&&{pns:sd.pns,units:sd.units,landed:sd.landed,landedU:sd.landedU,
       late:sd.late,lateU:sd.lateU,murky:sd.murky,murkyU:sd.murkyU},
     full:g('ETA-FULL'),past:g('ETA-PAST'),part:g('ETA-PARTIAL'),t1:g('T1-WITH-ETA'),
     line:lineLedgerLine(),
     seeded:document.querySelector('.lseed')?document.querySelector('.lseed').textContent:''}});
 /* דוח בלי «מסמך רכש» — העמודה אינה תנאי לקליטה */
 await p3.setInputFiles('#fe',SD+'/eta-report.xlsx');await p3.waitForTimeout(1200);
 const noPoCol=await p3.evaluate(()=>({hasPo:!!(ETA&&ETA.hasPo),rows:ETA&&ETA.rows}));

 /* ============ שורות אספקה שלא שויכו לאף מק״ט ============
    דוח ה-ETA מגיע ממערכת אחרת והשיוך הוא לפי מחרוזת. שורה שלא מתחברת
    נופלת בשקט, והכלי ממשיך לחשב כאילו אין אספקה בדרך. על פריט רגיל זו
    אי-דיוק; על פריט שנמצא בחוסר זו המלצה לפתוח רכש על סחורה שכבר
    בדרך. הקובץ כאן נפרד מדוח ה-ETA המשותף, כדי לא להזיז ספירות
    שבדיקות אחרות נועלות. */
 const oRows=[
  ['4190000101','000010',RLM('ETA-FULL'),'משויך',5,'EA',inMonth],
  ['4190000102','000010','ZZZ-NOT-IN-CATALOG','יתום',9,'EA',inMonth],
  ['4190000103','000010','ZZZ-ALSO-MISSING','יתום',4,'EA',inMonth],
  /* ETA-NOPO בכתיב אחר — אפסים מובילים. הפריט בוער: לקוח ממתין, אפס
     מלאי ואפס רכש. אי-השיוך הוא בדיוק המקרה המסוכן — הכלי ימליץ
     לפתוח רכש על סחורה שכבר בדרך. */
  ['4190000104','000010',RLM('00ETA-NOPO'),'יתום מסוכן',30,'EA',inMonth]];
 XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([ehdr,...oRows]),'גיליון1');return wb})(),
  SD+'/eta-orphan.xlsx');
 const p4=await ctx.newPage();p4.on('pageerror',e=>errs.push('p4: '+e.message));
 await p4.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p4.evaluate(()=>localStorage.clear());await p4.reload();await p4.waitForTimeout(400);
 await p4.setInputFiles('#f',SD+'/eta-zmrp.xlsx');await p4.waitForTimeout(1800);
 await p4.setInputFiles('#fe',SD+'/eta-orphan.xlsx');await p4.waitForTimeout(1200);
 const orph=await p4.evaluate(()=>{const o=ETA_ORPH;
   return {o:o&&{rows:o.rows,pns:o.pns,units:o.units,risky:o.risky.map(x=>x.pn)},
     chip:(document.getElementById('etachip').textContent||''),
     warn:((DIAG.warn||[]).filter(w=>/לא שויכו לאף פריט בקטלוג/.test(w))[0]||'')}});
 await p4.evaluate(()=>etaClear());await p4.waitForTimeout(300);
 const orphCleared=await p4.evaluate(()=>
   !(DIAG.warn||[]).some(w=>/לא שויכו/.test(w))&&!ETA_ORPH);
 await p4.close();

 const out=[],ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));

 /* ── שורות אספקה שלא שויכו ── */
 ok('שורות שלא שויכו נספרות',
   !!orph.o&&orph.o.rows===3&&orph.o.pns===3&&orph.o.units===43,
   orph.o?`${orph.o.rows} שורות · ${orph.o.pns} מק״טים · ${orph.o.units} יח׳`:'לא נמדד');
 ok('והן נאמרות במפורש ולא נופלות בשקט',
   /לא שויכו לאף פריט בקטלוג/.test(orph.warn),orph.warn.slice(0,85));
 ok('השבב בסרגל מסמן את זה',/לא שויכו/.test(orph.chip),orph.chip);
 /* הדגש שהמעתד ביקש: פריטים שנמצאים *עכשיו* בחוסר */
 ok('ומי שנראה כמו מק״ט שבחוסר מסומן בנפרד ובשמו',
   !!orph.o&&orph.o.risky.length===1&&orph.o.risky[0]==='ETA-NOPO',
   orph.o&&orph.o.risky.join(' · '));
 ok('והנוסח אומר מה לעשות לפני שפותחים רכש',
   /בדוק את הכתיב לפני שתפתח עליהם רכש/.test(orph.warn),
   (orph.warn.match(/⚠[^.]*\./)||[''])[0].slice(0,115));
 ok('הסרת דוח ה-ETA מסירה גם את האזהרה',orphCleared);

 /* ── SHELL_BAND — ראה ההערה בשלב 1 ── */
 /* היה: «מספר-גיבור» בגודל 48px ומעלה. המפרט החדש מבקש רצועת מדדים
    *קומפקטית* מעל הטבלה, והוא מחליף העדפת עיצוב קודמת שסותרת אותו.
    מה שנשמר הוא מה שהכלל הזה באמת הגן עליו: יש רצועה, יש בה מספר
    גדול מהטקסט סביבו, ושני המצבים — עם דוח ETA ובלי — באותו סדר
    גודל. מה שנוסף: הרצועה עצמה חייבת להישאר קומפקטית. */
 ok('בלי דוח ETA יש רצועה קומפקטית עם מספר בולט',
   bDry.band>0&&bDry.band<=72&&bDry.hero>=24,
   `רצועה ${bDry.band}px · גיבור ${bDry.hero}px «${bDry.heroTxt}»`);
 ok('ובלי ETA הפס מפצל בין «הוזמן» ל«לא הוזמן»',bDry.segs>0,bDry.segs+' מקטעים');
 ok('והמסך אומר במפורש שאין דוח ETA',gapDry);
 ok('עם דוח ETA הרצועה נשארת והמספר הבולט נשאר',
   bEta.band>0&&bEta.band<=72&&bEta.hero>=24,
   `רצועה ${bEta.band}px · גיבור ${bEta.hero}px «${bEta.heroTxt}»`);
 ok('שתי הרצועות באותו סדר גודל — אף מצב אינו «המצב העני»',
   Math.abs(bEta.band-bDry.band)<=40,`${bDry.band}px בלי ETA · ${bEta.band}px עם`);
 ok('ושורת «אין דוח ETA» נעלמת כשיש דוח',!gapEta);

 /* ── LEDGER_SEED ──
    הפנקס הרגיל דורש שני ימים שמקיפים את תאריך ההבטחה, ולכן ביום
    הראשון הוא מחזיר אפס. דוח ETA היסטורי עוקף את זה: הרכש הפתוח של
    היום משמש כ«אחרי». זו הסקה ולא תצפית, והמסך אומר את זה. */
 ok('«מסמך רכש» נקלט כשהוא קיים',seed.hasPo);
 ok('ודוח בלי העמודה עדיין נטען במלואו',!noPoCol.hasPo&&noPoCol.rows>0,
   `hasPo=${noPoCol.hasPo} · ${noPoCol.rows} שורות`);
 ok('בלי הבטחות שעברו — אין זריעה בכלל',!seedBefore.sd);
 ok('שורת הפנקס אז אומרת שהוא התחיל היום',/התחיל היום/.test(seedBefore.line),seedBefore.line);
 ok('הזריעה מוצאת את שלושת המקרים',
   !!seed.sd&&seed.sd.landed===1&&seed.sd.late===1&&seed.sd.murky===1,
   seed.sd?`נחת ${seed.sd.landed} · איחור ${seed.sd.late} · לא מוסבר ${seed.sd.murky}`:'אין');
 ok('«נחת» — רכש פתוח שווה ליחידות העתידיות',
   seed.full.po===seed.full.fut&&seed.full.past===25&&seed.sd.landedU===25,
   `po=${seed.full.po} עתיד=${seed.full.fut} עבר=${seed.full.past}`);
 ok('«איחור» — רכש פתוח שווה לעבר ועתיד יחד',
   seed.past.po===seed.past.fut+seed.past.past&&seed.sd.lateU===12,
   `po=${seed.past.po} עתיד=${seed.past.fut} עבר=${seed.past.past}`);
 ok('«לא מוסבר» — לא תואם לאף אחד משניהם',
   seed.part.po!==seed.part.fut&&seed.part.po!==seed.part.fut+seed.part.past&&seed.sd.murkyU===7,
   `po=${seed.part.po} עתיד=${seed.part.fut} עבר=${seed.part.past}`);
 ok('פריט שכל הבטחותיו עתידיות אינו נספר',seed.t1.past===0&&seed.sd.pns===3,
   `עבר=${seed.t1.past} · נספרו ${seed.sd.pns} מק״טים`);
 ok('הסכום הוא של היחידות שעברו בלבד',seed.sd.units===44,`${seed.sd.units} יח׳ (25+12+7)`);
 /* ההבחנה שאסור לטשטש: הסקה אינה תצפית */
 ok('המסך אומר «מוסקת» ולא «אמינות תאריכים»',
   /אמינות מוסקת/.test(seed.line)&&!/אמינות תאריכים/.test(seed.line),
   seed.line.replace(/<[^>]+>/g,' ').replace(/\s+/g,' ').trim().slice(0,90));
 ok('ומציין כמה לא מוסברים',/לא מוסבר/.test(seed.seeded),seed.seeded);

 ok('בלי דוח ETA — עמודה אחת "בדרך", בדיוק כמו קודם',
    !before.loaded&&before.chipHidden&&before.nOtw>0&&before.nSched===0
    &&before.cols.includes('בדרך ⌛')&&!before.cols.includes('משובץ'),
    `otw=${before.nOtw} sched=${before.nSched}`);
 ok('בלי דוח ETA — אין שדות ETA על השורות',
    before.P.eta===0&&before.P.stuck===0&&!before.P.has,
    `eta=${before.P.eta} stuck=${before.P.stuck}`);

 ok('הדוח נקלט ומזוהה לפי הכותרת שלו',
    /* 11 שורות ולא 10: נוספה אספקה שנייה ל-T2-PARTIAL אחרי סוף
       החודש, כדי ש«משובץ» יהיה תמיד גדול מ«מכוסה». 9 מק״טים —
       שתי השורות שייכות לאותו פריט. */
    after.loaded&&after.parts===9&&after.rows===11&&after.units===84,
    `מק״טים=${after.parts} שורות=${after.rows} יח׳=${after.units}`);
 ok('שורה בלי תאריך מוחרגת ולא מפילה את הקריאה',after.undated===1,`undated=${after.undated}`);
 ok('סימני כיוון RTL מוסרים — המק״טים מתחברים',after.hit===8,
    `התחברו ${after.hit} מתוך 9 (אחד אינו בקטלוג — בכוונה)`);

 ok('משלוח חלקי: 18 ברכש, 5 משובצות, 13 תקועות',
    after.P.po===18&&after.P.eta===5&&after.P.stuck===13,
    `po=${after.P.po} eta=${after.P.eta} stuck=${after.P.stuck}`);
 /* לפריט יש שתי אספקות, +7 ו-+21. המוצג חייב להיות המוקדמת. */
 const ymd=d=>d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
 ok('התאריך המוצג הוא המוקדם מבין האספקות',
    after.P.first===ymd(dplus(7)),
    `first=${after.P.first} · המוקדמת ${ymd(dplus(7))} · המאוחרת ${ymd(dplus(21))}`);
 ok('משלוח מלא: תקוע אפס',after.F.eta===10&&after.F.stuck===0,
    `eta=${after.F.eta} stuck=${after.F.stuck}`);
 ok('תאריך שעבר אינו נספר כמשובץ — המשלוח איחר, הוא לא כיסוי',
    after.T.eta===0&&after.T.past===12&&after.T.stuck===12,
    `eta=${after.T.eta} past=${after.T.past} stuck=${after.T.stuck}`);
 ok('מק״ט שאינו בדוח — כל הרכש תקוע',
    !after.N.has&&after.N.eta===0&&after.N.stuck===20,
    `stuck=${after.N.stuck}`);
 ok('אספקה גדולה מהרכש — תקוע אפס ולא שלילי',
    after.O.eta===9&&after.O.stuck===0,`eta=${after.O.eta} stuck=${after.O.stuck}`);
 ok('בלי רכש פתוח — תקוע אפס',after.Z.stuck===0,`stuck=${after.Z.stuck}`);

 /* «משובץ» שונה שמו ל«מכוסה» ומציג את מה שנוחת עד סוף החודש
    בלבד — ראה SHORT_ANSWER. הפיצול לשני תאים לא השתנה. */
 ok('עם דוח — "בדרך" מתפצל ל"מכוסה" ו"ללא תאריך"',
    after.cols.includes('מכוסה')&&after.cols.includes('ללא תאריך')
    &&!after.cols.includes('בדרך ⌛')&&after.nSched>0&&after.nOtw===0,
    `sched=${after.nSched} otw=${after.nOtw}`);
 ok('הצ׳יפ מציג את מספר המק״טים בדוח',
    !after.chipHidden&&/\b9\b/.test(after.chip)&&/ETA/.test(after.chip),after.chip.trim());

 ok('העלאת ZMRP מחדש אינה מוחקת את דוח ה-ETA',
    reload.loaded&&reload.P.eta===5&&reload.P.stuck===13&&reload.nSched>0,
    `loaded=${reload.loaded} eta=${reload.P.eta}`);
 ok('סדר הפוך — ETA לפני ZMRP — נותן בדיוק אותה תוצאה',
    reverse.loaded&&reverse.P.eta===5&&reverse.P.stuck===13&&reverse.hit===8,
    `eta=${reverse.P.eta} stuck=${reverse.P.stuck} hit=${reverse.hit}`);
 ok('קובץ ZMRP שהועלה לשדה ה-ETA נדחה, והדוח הקיים נשאר',
    wrong.loaded&&wrong.parts===9&&wrong.P.eta===5,
    `parts=${wrong.parts}`);

 // ── ETA_COVER · אספקה מתוארכת נחשבת כיסוי — בשכבה 2 בלבד ──
 ok('בלי דוח: 8 יח׳ לא יסופקו לפי המדף בלבד',
    before.C.uns===8&&before.C.gross===8&&before.C.inList,
    `uns=${before.C.uns} gross=${before.C.gross}`);
 ok('אספקה בתוך החודש מכסה — הפריט יורד מהרשימה',
    after.C.month===8&&after.C.uns===0&&after.C.val===0&&!after.C.inList,
    `month=${after.C.month} uns=${after.C.uns} ברשימה=${after.C.inList}`);
 ok('אותה כמות אחרי סוף החודש אינה מכסה כלום',
    after.L.month===0&&after.L.uns===8&&after.L.inList,
    `eom=${after.eom} month=${after.L.month} uns=${after.L.uns}`);
 ok('כיסוי חלקי מקטין ולא מוחק',
    after.PA.month===3&&after.PA.uns===5&&after.PA.inList,
    `month=${after.PA.month} uns=${after.PA.uns} (היה ${before.PA.uns})`);
 ok('תאריך שעבר אינו מכסה — ETA-PAST נשאר עם הברוטו',
    after.T.month===0&&after.T.uns===before.T.uns,
    `month=${after.T.month} uns=${after.T.uns} (היה ${before.T.uns})`);
 /* הבדיקה שמגינה על ההבטחה: שכבה 1 היא כשל שירות שכבר קרה, ותאריך
    עתידי אינו מבטל אותו. המספר והכסף חייבים להישאר זהים בול. */
 ok('שכבה 1 עם אספקה בתוך החודש — נשארת, ובאותו מספר בדיוק',
    after.W.t1&&after.W.inList&&after.W.month===20
    &&after.W.uns===before.W.uns&&after.W.val===before.W.val,
    `month=${after.W.month} uns=${after.W.uns} (היה ${before.W.uns}) · כסף ${after.W.val} (היה ${before.W.val})`);
 ok('מספר פריטי שכבה 1 לא השתנה',after.t1N===before.t1N,
    `${before.t1N} → ${after.t1N}`);
 ok('הרשימה התקצרה רק בשכבה 2',
    after.listN===before.listN-1&&after.covered===1,
    `${before.listN} → ${after.listN} · כוסו ${after.covered}`);
 ok('הכותרת אומרת כמה ירדו — הרשימה לא מתקצרת בשקט',
    /1/.test(after.badge)&&/כוסו/.test(after.badge),after.badge.trim()||'(אין תג)');
 ok('הכיסוי שורד קובץ שגוי שהועלה לשדה',wrong.C.uns===0,`uns=${wrong.C.uns}`);
 /* הסרת הדוח חייבת להחזיר את המצב בול — אחרת «כיסוי» הוא דלת חד-כיוונית. */
 ok('הסרת הדוח מחזירה את המספרים במדויק',
    !cleared.loaded&&cleared.C.uns===8&&cleared.C.val===before.C.val
    &&cleared.listN===before.listN&&cleared.covered===0&&cleared.badge==='',
    `uns=${cleared.C.uns} רשימה=${cleared.listN} (היה ${before.listN})`);

 // ── ETA_HONEST · אף מסך לא מכריז «אין תאריך» על פריט שיש לו תאריך ──
 ok('יש פריטים משובצים לסרוק',after.scheduled>0,`${after.scheduled} פריטים עם אספקה משובצת`);
 ok('אף הסבר או פעולה לא מכחישים תאריך שקיים',
    after.liars.length===0,after.liars.slice(0,3).join('  //  '));
 ok('ובכיוון ההפוך — בלי דוח אף מסך לא ממציא תאריך',
    before.ghosts.length===0&&cleared.ghosts.length===0,
    (before.ghosts.concat(cleared.ghosts)).slice(0,3).join(', '));

 /* ============ SHORT_ANSWER · ארבע תשובות בכל שורת חוסר ============
    T2-PARTIAL הוא המקרה היחיד בפיקסצ'ר שבו הכיסוי *כן* מנוכה:
    שכבה 2, כיסוי 3 מתוך 8. שם, ורק שם, מוצגת שרשרת החשבון.
    T1-WITH-ETA הוא ההפך: מכוסה 20, ואינו מנוכה כלל. */
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(700);
 const ans=await p.evaluate(()=>{
  const heads=[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim());
  const rows=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
  const cellOf=pn=>{const tr=rows.find(x=>x.textContent.includes(pn));
    return tr?{cov:tr.querySelector('td.sched, td.otw'),
      chain:tr.querySelector('td .chain'),nod:tr.querySelector('td .nodeduct'),
      act:tr.querySelector('td.act1')}:null};
  const pa=cellOf('T2-PARTIAL'),t1=cellOf('T1-WITH-ETA');
  const m=pn=>{const r=ALL.find(x=>x.pn===pn);return r?{month:r.etaMonth,fut:r.etaQty,
    gross:Math.ceil(r.unsGross!=null?r.unsGross:r.unsQty),left:Math.ceil(r.unsQty||0),t1:isT1(r)}:null};
  const w=document.querySelector('.rows');
  return {heads:heads.join('|'),
    nAct:rows.filter(tr=>{const a=tr.querySelector('td.act1');
      return a&&a.textContent.trim()&&a.textContent.trim()!=='—'}).length,
    nRows:rows.length,
    paChain:pa&&pa.chain?pa.chain.textContent.trim():null,
    paCov:pa&&pa.cov?pa.cov.textContent.trim():null,
    paNod:!!(pa&&pa.nod),
    t1Cov:t1&&t1.cov?t1.cov.textContent.trim():null,
    t1Nod:!!(t1&&t1.nod),t1Chain:!!(t1&&t1.chain),
    PA:m('T2-PARTIAL'),W:m('T1-WITH-ETA'),
    ovf:w?w.scrollWidth-w.clientWidth:0}});
 ok('הכותרת נושאת את ארבע התשובות',
    /מכוסה/.test(ans.heads)&&/חוסר חזוי/.test(ans.heads)&&/הפעולה הבאה/.test(ans.heads)
    /* «דרישת לקוח» מופרדת מ«חוסר חזוי» — עובדה מול תחזית */
    &&/דרישת לקוח/.test(ans.heads),ans.heads);
 ok('לכל שורה יש פעולה — לא תווית סיווג',
    ans.nAct===ans.nRows&&ans.nRows>0,`${ans.nAct} / ${ans.nRows}`);
 /* «מכוסה» הוא מה שנוחת עד סוף החודש, לא כל העתיד המתוארך. */
 ok('«מכוסה» מציג את מה שמכסה החודש ולא את כל העתיד',
    ans.paCov===String(ans.PA.month)&&ans.PA.fut>ans.PA.month,
    `מוצג ${ans.paCov} · מכסה ${ans.PA.month} · משובץ ${ans.PA.fut}`);
 ok('כיסוי שמנוכה — שרשרת החשבון מוצגת בתא «נשאר»',
    ans.paChain===`${ans.PA.gross}−${ans.PA.month}→`&&ans.PA.left<ans.PA.gross,
    `${ans.paChain} (חסר ${ans.PA.gross} · נשאר ${ans.PA.left})`);
 ok('וכיסוי שמנוכה אינו נושא ⊘',!ans.paNod);
 /* שכבה 1: הכיסוי קיים, אינו מנוכה, והמסך אומר זאת. */
 ok('שכבה 1 — הכיסוי מוצג ונושא ⊘',
    ans.t1Nod&&ans.t1Cov===String(ans.W.month)+'⊘',`${ans.t1Cov} · ⊘=${ans.t1Nod}`);
 ok('ובשכבה 1 אין שרשרת — כי לא נוכה דבר',
    !ans.t1Chain&&ans.W.left===ans.W.gross,
    `חסר ${ans.W.gross} · נשאר ${ans.W.left} · מכוסה ${ans.W.month}`);
 ok('מסלול החוסרים אינו גולש אופקית',ans.ovf<=20,ans.ovf+'px');

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 console.log(out.join('\n'));
 const f=out.filter(x=>x.startsWith('FAIL')).length;
 console.log(`\n${out.length-f}/${out.length} עברו`);
 await b.close();
 process.exit(f?1:0);
})();
