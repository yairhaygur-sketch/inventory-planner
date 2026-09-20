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
 // שכבה 1 עם אספקה בתוך החודש ⇒ חייב להישאר, בלי שינוי במספר
 ['4190000010','000010',RLM('T1-WITH-ETA'),'T1ETA',20,'EA',inMonth]];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([ehdr,...erows]),'גיליון1');return wb})(),
 SD+'/eta-report.xlsx');

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

 const out=[],ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));

 /* ── SHELL_BAND — ראה ההערה בשלב 1 ── */
 ok('בלי דוח ETA יש רצועה עם מספר-גיבור',bDry.band>0&&bDry.hero>=48,
   `רצועה ${bDry.band}px · גיבור ${bDry.hero}px «${bDry.heroTxt}»`);
 ok('ובלי ETA הפס מפצל בין «הוזמן» ל«לא הוזמן»',bDry.segs>0,bDry.segs+' מקטעים');
 ok('והמסך אומר במפורש שאין דוח ETA',gapDry);
 ok('עם דוח ETA הרצועה נשארת ומספר-הגיבור נשאר',bEta.band>0&&bEta.hero>=48,
   `רצועה ${bEta.band}px · גיבור ${bEta.hero}px «${bEta.heroTxt}»`);
 ok('שתי הרצועות באותו סדר גודל — אף מצב אינו «המצב העני»',
   Math.abs(bEta.band-bDry.band)<=40,`${bDry.band}px בלי ETA · ${bEta.band}px עם`);
 ok('ושורת «אין דוח ETA» נעלמת כשיש דוח',!gapEta);

 ok('בלי דוח ETA — עמודה אחת "בדרך", בדיוק כמו קודם',
    !before.loaded&&before.chipHidden&&before.nOtw>0&&before.nSched===0
    &&before.cols.includes('בדרך ⌛')&&!before.cols.includes('משובץ'),
    `otw=${before.nOtw} sched=${before.nSched}`);
 ok('בלי דוח ETA — אין שדות ETA על השורות',
    before.P.eta===0&&before.P.stuck===0&&!before.P.has,
    `eta=${before.P.eta} stuck=${before.P.stuck}`);

 ok('הדוח נקלט ומזוהה לפי הכותרת שלו',
    after.loaded&&after.parts===9&&after.rows===10&&after.units===80,
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

 ok('עם דוח — "בדרך" מתפצל ל"משובץ" ו"תקוע"',
    after.cols.includes('משובץ')&&after.cols.includes('תקוע')
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

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 console.log(out.join('\n'));
 const f=out.filter(x=>x.startsWith('FAIL')).length;
 console.log(`\n${out.length-f}/${out.length} עברו`);
 await b.close();
 process.exit(f?1:0);
})();
