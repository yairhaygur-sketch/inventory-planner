/* דוח הזירוז — מה לבקש מכל ספק, ומה להוריד מהרשימה לפני שמתקשרים.

   שתי הבדיקות החשובות כאן אינן על מה שנכנס לדוח אלא על מה שיוצא
   ממנו, ועל מה שכמעט יצא ממנו בטעות. שתיהן נמצאו במדידה על הדוח
   האמיתי אחרי שהקוד כבר עבד:
     · 3 פריטים עם לקוח ממתין מול מדף ריק סווגו «נחת» והושמטו,
       דווקא בגלל שמשלוח קודם כן הגיע.
     · פריט אחד נשר כי חישוב האזילה לפי קצב ממוצע אמר שהוא מכוסה,
       בזמן שהזמנת לקוח בשם יושבת מולו. */
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
 return [pn,'פריט '+pn,'Test '+pn,o.sup||'ספק א',  '01','פעיל','ND','A',95,
  0,0,o.lt||30,2,0,o.price!=null?o.price:100,'USD',o.free,o.free,0,o.po||0,0,0,o.cust||0,'Z004','מנוע','ZT','מתכנן','0',
  'שיווק','מערכת','100','ZT','200','דגם','300','מערכת',y,y,y,0,...m,D(10),D(60)]};
const T=Array(11).fill(10);   // קצב 10 לחודש
const NOR=[0,0,2,0,0,0,0,0,0,0,0]; // d12=2 ⇒ קצב אפס, «אוזל בעוד» חסר מובן
const rows=[
 mk('EXP-LATE'     ,{months:T,free:0, po:12,cust:0,price:400}),
 mk('EXP-PARTIAL'  ,{months:T,free:0, po:18,cust:5,price:1200}),
 mk('EXP-NODATE'   ,{months:T,free:0, po:20,cust:0,price:500,sup:'ספק ב'}),
 mk('EXP-LANDED'   ,{months:T,free:10,po:10,cust:0,price:300}),
 mk('EXP-LANDED-CW',{months:T,free:0, po:10,cust:2,price:300}),
 mk('EXP-ONTIME-CW',{months:NOR,free:0,po:6, cust:3,price:700}),
 mk('EXP-PULL'     ,{months:T,free:5, po:20,cust:0,price:250}),
 mk('EXP-NOORDER'  ,{months:T,free:0, po:0, cust:4,price:600}),
 mk('EXP-OK'       ,{months:T,free:50,po:10,cust:0,price:150})];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['ZMRP'],[],hdr,...rows]),'ZMRP');return wb})(),
 SD+'/expedite-zmrp.xlsx');

const RLM=s=>'‎'+s+'‏';
const dplus=n=>{const d=new Date(Date.now()+n*864e5);return new Date(d.getFullYear(),d.getMonth(),d.getDate())};
const ehdr=['אספקה','פריט','חומר','תיאור','מסמך רכש','כמות באספקה',"א'",'תארי.אספקה'];
const erows=[
 // התאריך עבר והכמות כולה עדיין פתוחה ⇒ «איפה זה?»
 ['4190200001','000010',RLM('EXP-LATE'),'LATE','4500200001',12,'EA',dplus(-30)],
 // חלק שובץ לעתיד, חלק עבר, והסכום אינו מסתדר ⇒ «מה הגיע ומה לא?»
 ['4190200002','000010',RLM('EXP-PARTIAL'),'PART','4500200002',5,'EA',dplus(21)],
 ['4190200003','000010',RLM('EXP-PARTIAL'),'PART','4500200003',7,'EA',dplus(-55)],
 // רכש 10, עתיד 10, ועוד 15 שעברו וכבר ירדו ⇒ נחת, יורד מהדוח
 ['4190200004','000010',RLM('EXP-LANDED'),'LAND','4500200004',10,'EA',dplus(14)],
 ['4190200005','000010',RLM('EXP-LANDED'),'LAND','4500200005',15,'EA',dplus(-40)],
 // אותו דפוס בדיוק — אבל לקוח ממתין מול מדף ריק ⇒ אסור שיירד
 ['4190200006','000010',RLM('EXP-LANDED-CW'),'LANDCW','4500200006',10,'EA',dplus(14)],
 ['4190200007','000010',RLM('EXP-LANDED-CW'),'LANDCW','4500200007',15,'EA',dplus(-40)],
 // תאריך קרוב, בלי קצב ⇒ «אוזל בעוד» אינו מחושב. לקוח ממתין ⇒ נכנס
 ['4190200008','000010',RLM('EXP-ONTIME-CW'),'ONTIME','4500200008',6,'EA',dplus(5)],
 // מדף 5 בקצב 10 נגמר בעוד 15 יום, המשלוח בעוד 60 ⇒ «להקדים», דחיפות 3
 ['4190200009','000010',RLM('EXP-PULL'),'PULL','4500200009',20,'EA',dplus(60)],
 // מדף 50 בקצב 10 מחזיק 150 יום, המשלוח בעוד 14 ⇒ אינו בדוח כלל
 ['4190200010','000010',RLM('EXP-OK'),'OK','4500200010',10,'EA',dplus(14)]];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([ehdr,...erows]),'גיליון1');return wb})(),
 SD+'/expedite-eta.xlsx');

const snap=()=>{const {rows,landed}=expediteRows();
 const by={};for(const x of rows)by[x.r.pn]={ask:x.ask.k,urg:x.urg,qty:x.qty,po:expPo(x.r)};
 const sup=expSupplierAoa(rows);
 return {by,n:rows.length,pns:rows.map(x=>x.r.pn),
  landed:landed.map(r=>r.pn),
  order:rows.map(x=>x.r.pn),
  notes:expNotes(rows,landed),
  supHdr:sup[0],supBody:sup.slice(1),
  landedAoa:expLandedAoa(landed),
  cols:EXPCOLS.map(c=>c[0]),
  /* כל פריט עם לקוח ממתין מול מדף ריק חייב להיות בדוח. אין חריג. */
  missedCW:(ALL||[]).filter(r=>r.cust>0&&(r.free||0)<=0)
   .filter(r=>!rows.some(x=>x.r.pn===r.pn)).map(r=>r.pn),
  /* ואף שורה בדוח אינה מחברת מטבעות */
  supMoney:sup.slice(1).map(r=>r[4]).filter(Boolean)}};

(async()=>{
 const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860},acceptDownloads:true});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.evaluate(()=>localStorage.removeItem('planner_eta_v1'));
 await p.reload();await p.waitForTimeout(300);

 // ── שלב 1: ZMRP בלבד. הכלי אינו יודע תאריכים, והוא אומר זאת ──
 await p.setInputFiles('#f',SD+'/expedite-zmrp.xlsx');await p.waitForTimeout(1800);
 const dry=await p.evaluate(snap);
 const btnSeen=await p.evaluate(()=>{const e=document.getElementById('expXls');
  return !!e&&e.offsetParent!==null});

 // ── שלב 2: דוח ה-ETA ──
 await p.setInputFiles('#fe',SD+'/expedite-eta.xlsx');await p.waitForTimeout(1500);
 const wet=await p.evaluate(snap);

 // ── שלב 3: הכפתור עצמו — עד הקובץ שיורד ──
 const dl=await Promise.all([p.waitForEvent('download',{timeout:15000}),
   p.click('#expXls')]).then(a=>a[0]).catch(()=>null);
 let book=null;
 if(dl){const f=await dl.path();
  if(f){const w=XLSX.readFile(f);
   book={names:w.SheetNames,
    lines:XLSX.utils.sheet_to_json(w.Sheets['לזירוז'],{header:1}),
    sup:XLSX.utils.sheet_to_json(w.Sheets['לפי ספק'],{header:1}),
    land:w.Sheets['נחתו — לא לרדוף']
      ?XLSX.utils.sheet_to_json(w.Sheets['נחתו — לא לרדוף'],{header:1}):null,
    fname:dl.suggestedFilename()}}}

 const out=[],ok=(n,c,d)=>out.push(`${c?'PASS':'FAIL'}  ${n}${d?'  ['+d+']':''}`);

 // ── בלי דוח ETA ──
 ok('הכפתור קיים וגלוי',btnSeen,String(btnSeen));
 ok('בלי דוח ETA — כל רכש פתוח הוא «תן תאריך»',
    dry.pns.filter(x=>x!=='EXP-NOORDER').every(x=>dry.by[x].ask==='nodate'),
    JSON.stringify(dry.pns.map(x=>x+':'+dry.by[x].ask)));
 ok('בלי דוח ETA — האזהרה אומרת שזה הכלי ולא הספק',
    dry.notes.some(t=>/לא נטען דוח ETA/.test(t)&&/אינו יודע/.test(t)),
    dry.notes[0]||'(אין)');
 ok('בלי דוח ETA — אין גיליון «נחתו», כי אין ממה להסיק',
    dry.landed.length===0,dry.landed.join(','));

 // ── הסיווג עצמו ──
 const A=wet.by;
 ok('תאריך שעבר והכמות כולה פתוחה ⇒ «איפה זה?»',
    A['EXP-LATE']&&A['EXP-LATE'].ask==='late'&&A['EXP-LATE'].qty===12,
    A['EXP-LATE']?A['EXP-LATE'].ask+' qty='+A['EXP-LATE'].qty:'חסר');
 ok('חלק עבר וחלק שובץ ⇒ «מה הגיע ומה לא?» על ההפרש',
    A['EXP-PARTIAL']&&A['EXP-PARTIAL'].ask==='partial'&&A['EXP-PARTIAL'].qty===13,
    A['EXP-PARTIAL']?A['EXP-PARTIAL'].ask+' qty='+A['EXP-PARTIAL'].qty:'חסר');
 ok('רכש בלי שום תאריך ⇒ «תן תאריך» על כל הכמות',
    A['EXP-NODATE']&&A['EXP-NODATE'].ask==='nodate'&&A['EXP-NODATE'].qty===20,
    A['EXP-NODATE']?A['EXP-NODATE'].ask+' qty='+A['EXP-NODATE'].qty:'חסר');
 ok('לקוח ממתין בלי שום רכש ⇒ «אין הזמנה בכלל»',
    A['EXP-NOORDER']&&A['EXP-NOORDER'].ask==='noorder'&&A['EXP-NOORDER'].qty===4,
    A['EXP-NOORDER']?A['EXP-NOORDER'].ask:'חסר');
 ok('מדף שמחזיק עד ההגעה ⇒ אינו בדוח כלל',
    !A['EXP-OK'],A['EXP-OK']?A['EXP-OK'].ask:'לא בדוח');

 // ── מה יורד מהדוח, ולמה ──
 ok('הבטחה שעברה והכמות ירדה מהרכש ⇒ נחת, יורד מהדוח',
    wet.landed.includes('EXP-LANDED')&&!A['EXP-LANDED'],
    'נחתו: '+wet.landed.join(','));
 ok('הנחיתה אינה נמחקת — היא בגיליון נפרד עם הסיבה',
    wet.landedAoa.length===2&&/נחתה/.test(String(wet.landedAoa[1][6])),
    String(wet.landedAoa[1]&&wet.landedAoa[1][6]||''));

 // ── שתי הרגרסיות שנמצאו במדידה על הדוח האמיתי ──
 ok('רגרסיה · «נחת» אינו בולע לקוח שממתין מול מדף ריק',
    !wet.landed.includes('EXP-LANDED-CW')
    &&A['EXP-LANDED-CW']&&A['EXP-LANDED-CW'].ask==='pull',
    A['EXP-LANDED-CW']?A['EXP-LANDED-CW'].ask:'נבלע');
 ok('רגרסיה · קצב ממוצע אינו מבטל הזמנת לקוח בשם',
    A['EXP-ONTIME-CW']&&A['EXP-ONTIME-CW'].ask==='pull'
    &&A['EXP-ONTIME-CW'].urg===1,
    A['EXP-ONTIME-CW']?A['EXP-ONTIME-CW'].ask+' urg='+A['EXP-ONTIME-CW'].urg:'נשר');
 ok('ואין אף פריט עם לקוח ממתין שאינו בדוח',
    wet.missedCW.length===0,wet.missedCW.join(',')||'0');

 // ── סולם הדחיפות ──
 ok('לקוח ממתין הוא דחיפות 1',
    A['EXP-PARTIAL'].urg===1&&A['EXP-NOORDER'].urg===1,
    'PARTIAL='+A['EXP-PARTIAL'].urg+' NOORDER='+A['EXP-NOORDER'].urg);
 ok('מדף אפס בלי לקוח הוא דחיפות 2',
    A['EXP-LATE'].urg===2&&A['EXP-NODATE'].urg===2,
    'LATE='+A['EXP-LATE'].urg+' NODATE='+A['EXP-NODATE'].urg);
 ok('מדף שנגמר לפני ההגעה, בלי לקוח, הוא דחיפות 3',
    A['EXP-PULL']&&A['EXP-PULL'].urg===3&&A['EXP-PULL'].ask==='pull',
    A['EXP-PULL']?'urg='+A['EXP-PULL'].urg+' '+A['EXP-PULL'].ask:'חסר');
 ok('הדוח ממוין לפי דחיפות — הדחופים בראש',
    wet.order.map(pn=>A[pn].urg).every((u,i,a)=>i===0||a[i-1]<=u),
    wet.order.map(pn=>pn+':'+A[pn].urg).join(' '));

 // ── מספר מסמך הרכש — מה שמצטטים לספק ──
 ok('מספר מסמך הרכש מגיע לשורה',
    /4500200002/.test(A['EXP-PARTIAL'].po)&&/4500200003/.test(A['EXP-PARTIAL'].po),
    A['EXP-PARTIAL'].po);
 ok('ולספק שאינו בדוח ה-ETA אין מה לצטט — והדוח אומר כמה',
    A['EXP-NODATE'].po===''&&wet.notes.some(t=>/בלי מספר מסמך רכש/.test(t)),
    wet.notes.find(t=>/מסמך רכש/.test(t))||'(אין)');

 // ── הסיכום לפי ספק: סדר היום ──
 ok('הסיכום נותן שורה לכל ספק',
    wet.supBody.length===2&&wet.supBody.map(r=>r[0]).sort().join()==='ספק א,ספק ב',
    wet.supBody.map(r=>r[0]).join(','));
 ok('ספק עם לקוחות ממתינים מופיע ראשון',
    wet.supBody[0][0]==='ספק א'&&wet.supBody[0][3]>0,
    wet.supBody[0][0]+' cw='+wet.supBody[0][3]);
 ok('שווי הזמנות הלקוח נכתב עם סימן מטבע ולא כמספר עירום',
    wet.supMoney.length>0&&wet.supMoney.every(s=>/[$€₪]/.test(s)),
    wet.supMoney.join(' | '));

 // ── הקובץ שיורד ──
 ok('לחיצה על הכפתור מורידה קובץ',!!book,dl?(book?book.fname:'ירד בלי נתיב'):'לא ירד');
 if(book){
  ok('שלושת הגיליונות בקובץ',
     book.names.join(',')==='לפי ספק,לזירוז,נחתו — לא לרדוף',book.names.join(','));
  ok('גיליון «לזירוז» מכיל את כל השורות ואת הכותרת',
     book.lines.length===wet.n+1&&book.lines[0][0]==='ספק',
     (book.lines.length-1)+' שורות מול '+wet.n);
  ok('גיליון «לפי ספק» נושא את האזהרות בראשו',
     book.sup.slice(0,8).some(r=>/דוח זירוז/.test(String(r[0])))
     &&book.sup.slice(0,8).some(r=>/הוסרו מהרשימה/.test(String(r[0]))),
     String(book.sup[0]&&book.sup[0][0]||''));
  ok('גיליון הנחיתות קיים בקובץ',!!book.land&&book.land.length===2,
     book.land?(book.land.length-1)+' שורות':'אין');
  ok('שם הקובץ נושא תאריך',/^Expedite_\d{4}-\d{2}-\d{2}\.xlsx$/.test(book.fname),book.fname)}

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 console.log(out.join('\n'));
 const f=out.filter(x=>x.startsWith('FAIL')).length;
 console.log(`\n${out.length-f}/${out.length} עברו`);
 await b.close();
 process.exit(f?1:0);
})();
