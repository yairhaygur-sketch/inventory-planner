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
 mk('ETA-NOPO',{months:[3,3,3,3,3,3,3,3,3,3,3],free:0,po:0,cust:1,price:600})];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['ZMRP'],[],hdr,...rows]),'ZMRP');return wb})(),
 SD+'/eta-zmrp.xlsx');

const dplus=n=>{const d=new Date(Date.now()+n*864e5);return new Date(d.getFullYear(),d.getMonth(),d.getDate())};
const ehdr=['אספקה','פריט','חומר','תיאור','כמות באספקה',"א'",'תארי.אספקה'];
const erows=[
 ['4190000001','000010',RLM('ETA-PARTIAL'),'PARTIAL',2,'EA',dplus(7)],
 ['4190000001','000020',RLM('ETA-PARTIAL'),'PARTIAL',3,'EA',dplus(21)],
 ['4190000002','000010',RLM('ETA-FULL'),'FULL',10,'EA',dplus(14)],
 ['4190000003','000010',RLM('ETA-PAST'),'PAST',12,'EA',dplus(-30)],
 ['4190000004','000010',RLM('ETA-OVER'),'OVER',9,'EA',dplus(10)],
 ['4190000005','000010',RLM('NOT-IN-CATALOGUE'),'GHOST',5,'EA',dplus(10)],
 // שורה בלי תאריך — מוחרגת, ואסור שתפיל את הקריאה
 ['4190000006','000010',RLM('ETA-NONE'),'NODATE',7,'EA','']];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([ehdr,...erows]),'גיליון1');return wb})(),
 SD+'/eta-report.xlsx');

const snap=()=>{const g=pn=>{const r=ALL.find(x=>x.pn===pn);if(!r)return null;
  return {po:r.po,eta:r.etaQty,past:r.etaPast,first:r.etaFirst,stuck:r.stuck,has:r.etaHas}};
 return {P:g('ETA-PARTIAL'),F:g('ETA-FULL'),T:g('ETA-PAST'),N:g('ETA-NONE'),
  O:g('ETA-OVER'),Z:g('ETA-NOPO'),
  loaded:!!ETA,parts:ETA?ETA.parts:0,rows:ETA?ETA.rows:0,units:ETA?ETA.units:0,undated:ETA?ETA.undated:0,
  hit:ALL.filter(x=>x.etaHas).length,
  cols:[...document.querySelectorAll('#tbl thead th')].map(t=>t.textContent.trim()),
  nSched:document.querySelectorAll('#tbl tbody td.sched').length,
  nOtw:document.querySelectorAll('#tbl tbody td.otw').length,
  chipHidden:document.getElementById('etachip').hidden,
  chip:document.getElementById('etachip').textContent}};

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

 // ── שלב 2: דוח ה-ETA נטען בנפרד, אחרי ZMRP ──
 await p.setInputFiles('#fe',SD+'/eta-report.xlsx');await p.waitForTimeout(1200);
 const after=await p.evaluate(snap);

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

 const out=[],ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));

 ok('בלי דוח ETA — עמודה אחת "בדרך", בדיוק כמו קודם',
    !before.loaded&&before.chipHidden&&before.nOtw>0&&before.nSched===0
    &&before.cols.includes('בדרך ⌛')&&!before.cols.includes('משובץ'),
    `otw=${before.nOtw} sched=${before.nSched}`);
 ok('בלי דוח ETA — אין שדות ETA על השורות',
    before.P.eta===0&&before.P.stuck===0&&!before.P.has,
    `eta=${before.P.eta} stuck=${before.P.stuck}`);

 ok('הדוח נקלט ומזוהה לפי הכותרת שלו',
    after.loaded&&after.parts===5&&after.rows===6&&after.units===41,
    `מק״טים=${after.parts} שורות=${after.rows} יח׳=${after.units}`);
 ok('שורה בלי תאריך מוחרגת ולא מפילה את הקריאה',after.undated===1,`undated=${after.undated}`);
 ok('סימני כיוון RTL מוסרים — המק״טים מתחברים',after.hit===4,
    `התחברו ${after.hit} מתוך 5 (אחד אינו בקטלוג — בכוונה)`);

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
    !after.chipHidden&&/310|5/.test(after.chip)&&/ETA/.test(after.chip),after.chip.trim());

 ok('העלאת ZMRP מחדש אינה מוחקת את דוח ה-ETA',
    reload.loaded&&reload.P.eta===5&&reload.P.stuck===13&&reload.nSched>0,
    `loaded=${reload.loaded} eta=${reload.P.eta}`);
 ok('סדר הפוך — ETA לפני ZMRP — נותן בדיוק אותה תוצאה',
    reverse.loaded&&reverse.P.eta===5&&reverse.P.stuck===13&&reverse.hit===4,
    `eta=${reverse.P.eta} stuck=${reverse.P.stuck} hit=${reverse.hit}`);
 ok('קובץ ZMRP שהועלה לשדה ה-ETA נדחה, והדוח הקיים נשאר',
    wrong.loaded&&wrong.parts===5&&wrong.P.eta===5,
    `parts=${wrong.parts}`);

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 console.log(out.join('\n'));
 const f=out.filter(x=>x.startsWith('FAIL')).length;
 console.log(`\n${out.length-f}/${out.length} עברו`);
 await b.close();
 process.exit(f?1:0);
})();
