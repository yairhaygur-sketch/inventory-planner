/* ============ החרגות סטטוס, וחסר שאינו אפס ============
   חמישה חוקים שהיו מתועדים או מוסכמים, ולא מומשו עד הסוף. כל טענה
   כאן נמדדה על דוח סינתטי לפני שנכתבה, והמספרים בהערות הם מה שהמנוע
   באמת החזיר לפני התיקון. */
const XLSX=require('xlsx'),{chromium}=require('playwright'),fs=require('fs'),path=require('path');
const SD=__dirname, MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const HDR=['מק"ט מוביל','תיאור חומר','תיאור חומר2','שם ספק','סטטוס חומר','תיאור','סוג MRP','ABC','רמת שרות',
 'מלאי בטחון','נק.הז.מחדש','אספ.מתוכנ.','זמ.עב.קבלת','מל.בט.מינ.','מחיר FOB','מטבע FOB','סה"כ מלאי','מלאי פנוי',
 'מלאי מרלוג','הז. רכש','בהעברה','אספקות פת.','כמות בהז.פ','סוג חומר','תיא.קבוצ.חומרים','קב.חו.חיצו','טקסט ארוך',
 'תב.אח.הש.','ת.היר.1 מח','ת.היר.2 מח','היררכייה1','היררכייה1','היררכייה2','היררכייה2','היררכייה3','היררכייה3',
 'צר.השנה','צר.שנה-1','צר.שנה-2','צר.החודש',...MON,"תאר' מכירה","תאר' כניסה"];
const D=n=>{const d=new Date(Date.now()-n*864e5);
 return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`};
const mk=(pn,o={})=>{const m=o.months||Array(11).fill(0),y=m.reduce((a,b)=>a+b,0);
 return [pn,'בדיקה '+pn,'T '+pn,'ספק',o.status||'01',o.stx||'פעיל',o.mrp||'ND','A',
  o.srv!=null?o.srv:95,o.ss||0,o.rop||0,o.lt!=null?o.lt:60,2,o.ssMin||0,
  o.price!=null?o.price:100,o.cur||'USD',o.stock!=null?o.stock:(o.free||0),o.free||0,0,
  o.po||0,o.transfer||0,0,o.cust||0,'Z004','מנוע','ZT','מתכנן',o.warranty!=null?o.warranty:0,
  'שיווק','מערכת','100','ZT','200','דגם','300','מערכת',
  o.y0!=null?o.y0:y,o.y1!=null?o.y1:y,o.y2!=null?o.y2:y,0,...m,
  /* null ⇒ תא ריק. «אין תאריך» אינו «תאריך ישן», וזה בדיוק ההבדל
     שהענף «חדש — אין עדיין היסטוריה» טעה בו. */
  o.saleAgo===null?'':D(o.saleAgo!=null?o.saleAgo:10),
  o.entAgo===null?'':D(o.entAgo!=null?o.entAgo:60)]};
function write(file,rows,opt={}){
 let hdr=HDR.slice(),rr=rows.map(r=>r.slice());
 for(const name of (opt.drop||[])){const i=hdr.indexOf(name);
  if(i>=0){hdr.splice(i,1);rr=rr.map(r=>{const c=r.slice();c.splice(i,1);return c})}}
 const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['ZMRP'],[],hdr,...rr]),'ZMRP');
 XLSX.writeFile(wb,file)}

/* ביקוש גושי ⇒ σ גדולה ⇒ המרכיב הסטטיסטי אינו אפס.
   לפני התיקון קיבלו שלושת הפריטים האלה SS=66 ו-ROP=117 — זהים. */
const LUMPY=[40,0,0,35,0,0,38,0,0,42,0];
write(SD+'/status-a.xlsx',[
 mk('S14-LUMPY',{status:'14',stx:'גמר המלאי',free:2,stock:2,price:300,lt:120,ssMin:7,months:LUMPY}),
 mk('S04-LUMPY',{status:'04',stx:'לא פעיל',  free:2,stock:2,price:300,lt:120,ssMin:7,months:LUMPY}),
 mk('OK-LUMPY', {                            free:2,stock:2,price:300,lt:120,ssMin:7,months:LUMPY}),
 /* נוכחות מינימלית — ROP=1 ב-SAP אינו יורד ל-0 גם בסטטוס 14 */
 mk('S14-MIN1',{status:'14',stx:'גמר המלאי',free:0,stock:0,rop:1,ss:1,price:80,months:[3,3,3,3,3,3,3,3,3,3,3]}),
 /* לקוח ממתין על פריט שמוחרג — יוצא מהתור, לא מהראייה */
 mk('S04-CUST',{status:'04',stx:'לא פעיל',free:0,stock:5,cust:12,price:300,months:[2,2,2,2,2,2,2,2,2,2,2]}),
 mk('S14-DRY', {status:'14',stx:'גמר המלאי',free:1,stock:1,price:300,months:[9,9,9,9,9,9,9,9,9,9,9]}),
 /* PD — חסר מול קיים */
 mk('PD-NOPRICE',{mrp:'PD',srv:50,free:0,price:0,  months:[1,1,1,1,0,0,0,0,0,0,0]}),
 mk('PD-PRICED', {mrp:'PD',srv:50,free:0,price:100,months:[1,1,1,1,0,0,0,0,0,0,0]}),
 mk('BG',{free:400,price:50,months:[2,2,2,2,2,2,2,2,2,2,2]})]);
/* אותם פריטים בלי עמודת האחריות */
write(SD+'/status-b.xlsx',[
 mk('PD-PRICED',{mrp:'PD',srv:50,free:0,price:100,months:[1,1,1,1,0,0,0,0,0,0,0]}),
 mk('BG',{free:400,price:50,months:[2,2,2,2,2,2,2,2,2,2,2]})],{drop:['תב.אח.הש.']});
/* מלאי מת — עם ובלי עמודות הצריכה השנתית */
const DEAD=[mk('DEAD-OLD',{free:9,stock:9,price:400,months:Array(11).fill(0),entAgo:1200,saleAgo:1200}),
            mk('BG',{free:200,price:50,months:[2,2,2,2,2,2,2,2,2,2,2]})];
write(SD+'/status-c.xlsx',DEAD);
write(SD+'/status-d.xlsx',DEAD,{drop:['צר.השנה','צר.שנה-1','צר.שנה-2']});
/* בלי עמודת הזמנות חסומות */
/* ============ שעון הגיל: הביקוש, לא הרכש ============
   `dead3y` דרש `entAge>=1095` כתנאי הכרחי, ו-entAge בא מ«תאר׳ כניסה»
   — כלומר מהרכש שלנו. פריט שלא נצרך שלוש שנים אבל קיבל מכולה לפני
   100 יום יצא 'young' → «חדש — אין עדיין היסטוריה» → Q.passive,
   expCap=0, והכסף שלו יצא מ«כסף תקוע». מכולה שנחתה מחקה מלאי מת.
   נמדד על וריאנט של הפיקסצ׳ר (117 פריטים חסרי-צריכה עם כניסה טרייה):
   ageTier dead 117→0, והכרטיס אמר «הכניסה האחרונה לפני 100 ימים —
   טרם עברו 3 שנים» על פריט שלא נמכר 1,986 יום.
   ארבעת המקרים כאן מכסים את שלושת מצבי השעון ואת הרגרסיה. */
const NOCONS={months:Array(11).fill(0),y0:0,y1:0,y2:0,free:5,stock:5,price:1000};
write(SD+'/status-f.xlsx',[
 /* מכולה טרייה, מכירה לפני 7 שנים ⇒ מת. זה התיקון. */
 mk('CLK-FRESH-ENTRY',{...NOCONS,entAgo:100,saleAgo:2575}),
 /* מכירה לפני 100 יום ⇒ עדיין לא מת, והשעון הוא המכירה ולא הכניסה */
 mk('CLK-RECENT-SALE',{...NOCONS,entAgo:2575,saleAgo:100}),
 /* אין מכירה, כניסה טרייה ⇒ השעון הוא הכניסה */
 mk('CLK-ENTRY-ONLY',{...NOCONS,entAgo:100,saleAgo:null}),
 /* אין שום תאריך ⇒ אין שעון, ואין להמציא «בעוד 37 חודשים» */
 mk('CLK-NO-DATES',{...NOCONS,entAgo:null,saleAgo:null}),
 /* רגרסיה: פריט עם צריכה אינו נוגע בכל זה */
 mk('CLK-ALIVE',{months:[4,4,4,4,4,4,4,4,4,4,4],free:5,stock:5,price:1000,entAgo:100,saleAgo:2575})]);
/* ============ RATE_3Y — קצב מהחלון התלת-שנתי ============
   השער `d12<=3` מחזיר קצב 0 לתשעה מכל עשרה פריטים בדוח האמיתי, גם
   כשעמודות הצריכה השנתיות מראות מאות יחידות. חמש הטענות כאן נמדדו
   על הדוח האמיתי לפני שנכתבו, ושלוש מהן הן *רגרסיות* — דברים שהתיקון
   שבר בגרסה הראשונה שלו והמדידה תפסה:
     · הסרגל: ענף העודף מחזיק 12 חודשים, הענף האיטי 24. ברגע שהקצב
       חדל להיות 0, `cov` חדל להיות null וענף 12 ירה ראשון — 45
       פריטים עברו מ«מלאי איטי» ל«עודף מלאי» בלי שום נתון חדש.
     · שני המעברים המשלימים («SAP יזמין בחזרה» ו«הפרמטר לא יחזיר את
       המלאי») חסומים ב-`rate>0` כפרוקסי ל«רונג הפרמטרים כבר טיפל בו».
       הקצב התלת-שנתי שבר את הפרוקסי ו-10 פריטים נשרו מקובץ העדכון.
   הסף נגזר מ-LOWDEM_MAX ואינו מספר חדש: 3 יח׳ ב-11 חודשים. */
const Y=n=>({y0:n,y1:0,y2:0}), Z11=Array(11).fill(0);
write(SD+'/status-g.xlsx',[
 /* עובר את הסף בגדול — 60 יח׳ בשלוש שנים, אפס בחלון הקצר */
 mk('R3-FAT',  {...Y(60),months:Z11,free:5, stock:5, rop:1,ss:0,lt:60,price:100}),
 /* יחידה אחת בשלוש שנים אינה קצב — היא רעש מחולק ב-34 */
 mk('R3-THIN', {...Y(1), months:Z11,free:5, stock:5, rop:1,ss:0,lt:60,price:100}),
 /* כיסוי 68 חודשים ⇒ הענף האיטי, בסרגל 24 — לא ענף העודף בסרגל 12 */
 mk('R3-SLOW', {...Y(30),months:Z11,free:60,stock:60,rop:1,ss:0,lt:60,price:100}),
 /* כיסוי בין 12 ל-24: לא איטי ולא עודף — נשאר בדיוק איפה שהיה */
 mk('R3-MID',  {...Y(30),months:Z11,free:15,stock:15,rop:1,ss:0,lt:60,price:100}),
 /* ROP ב-SAP מעל כל מה שנמכר בשלוש שנים ⇒ «SAP יזמין בחזרה» */
 mk('R3-BACK', {...Y(30),months:Z11,free:60,stock:60,rop:200,ss:50,lt:60,price:100}),
 /* הצד ההפוך: רצפת ssMin קיימת וה-ROP ב-SAP נמוך ממנה ⇒ SAP לא יחזיר */
 mk('R3-FLOOR',{...Y(30),months:Z11,free:60,stock:60,rop:1,ss:0,ssMin:5,lt:60,price:100}),
 /* מדף ריק: בלי השער, 0.88 יח׳/חודש פחות 0 נותנות «חוסר החודש» של
    0.88 יח׳ — ומכניסות את הפריט לדלת הבוערת */
 mk('R3-DRY',  {...Y(30),months:Z11,free:0, stock:9, rop:2,ss:1,lt:60,price:100}),
 /* הבקרה לשער: אותו מדף ריק בדיוק, אבל הביקוש בחלון הקצר אמיתי —
    כאן «חוסר החודש» חייב להישאר חי, אחרת השער חוסם יותר ממה שצריך */
 mk('R3-ACTDRY',{y0:36,y1:36,y2:36,months:Array(11).fill(3),free:0,stock:5,rop:2,ss:1,lt:60,price:100}),
 /* מגמה על חלון כמעט ריק: שתי יחידות בחודש אחד ואפס מאז.
    tRec=0 מול tPre=2 נותן tRatio=0 — «הביקוש התמוטט» על רעש */
 mk('R3-TREND',{...Y(21),months:[0,0,0,2,0,0,0,0,0,0,0],free:30,stock:30,rop:6,ss:4,lt:120,price:100}),
 /* «שני שעונים» על מלאי חסום: 30 יח׳ על המדף, אפס פנויות. הרונג
    האיטי מדווח כיסוי מהמלאי הכולל (34 ח׳), twoClocks מהפנוי (0 ח׳) —
    שני מספרים נכונים על שתי כמויות, ועד התיקון אף אחד לא אמר על איזו */
 mk('R3-TWOCLK',{...Y(30),months:Z11,free:0,stock:30,rop:2,ss:0,lt:120,price:100}),
 /* הבקרה: ירידה אמיתית בחלון אמיתי חייבת להמשיך להתרסן */
 mk('R3-FALL', {y0:60,y1:60,y2:60,months:[3,3,3,9,9,9,0,0,0,0,0],free:30,stock:30,rop:6,ss:4,lt:120,price:100}),
 /* הבקרה: ביקוש אמיתי בחלון הקצר — ענף העודף בסרגל 12 חייב להמשיך לירות */
 mk('R3-ACTIVE',{y0:36,y1:36,y2:36,months:Array(11).fill(3),free:60,stock:60,rop:1,ss:0,lt:60,price:100}),
 mk('BG',{free:400,price:50,months:[2,2,2,2,2,2,2,2,2,2,2]})]);
write(SD+'/status-e.xlsx',[
 mk('CUST-RAW',{free:0,cust:20,price:150,months:[4,4,4,4,4,4,4,4,4,4,4]}),
 mk('BG',{free:400,price:50,months:[2,2,2,2,2,2,2,2,2,2,2]})]);

const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');

(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const errs=[];
 const load=async f=>{const p=await ctx.newPage();p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.evaluate(()=>localStorage.clear());await p.reload();await p.waitForTimeout(400);
  await p.setInputFiles('#f',f);
  /* לא ממתינים לשורה בטבלה: חלק מהקבצים כאן הם בכוונה בלי שום פריט
     בתור העבודה (בדיוק מה שנבדק). ההמתנה היא לטעינה עצמה. */
  await p.waitForFunction(()=>typeof ALL!=='undefined'&&ALL&&ALL.length>0,{timeout:120000});
  await p.waitForTimeout(600);
  return p};

 /* ================= 1 · 04/14 אינם מקבלים SS/ROP ================= */
 {const p=await load(SD+'/status-a.xlsx');
  const r=await p.evaluate(()=>{const g=pn=>ALL.find(x=>x.pn===pn);
    const pick=pn=>{const x=g(pn);return {ss:x.sugSS,rop:x.sugROP,stat:x.ssStat,src:x.ssSrc,
      noStock:!!x.noStock,rop0:x.rop,ss0:x.ss}};
    return {s14:pick('S14-LUMPY'),s04:pick('S04-LUMPY'),ok:pick('OK-LUMPY'),min1:pick('S14-MIN1'),
      floorDoor:floorRows().map(x=>x.pn),
      trendDoor:trendRows().map(x=>x.pn),riseDoor:riseRows().map(x=>x.pn)}});
  /* לפני: 66/117 לשלושתם. הפעיל הוא הבקרה — הוא לא אמור לזוז. */
  ok('פריט סטטוס 14 אינו מקבל מלאי ביטחון מוצע',r.s14.ss===0,`SS=${r.s14.ss} (היה 66)`);
  ok('ואינו מקבל נקודת הזמנה מוצעת',r.s14.rop===0,`ROP=${r.s14.rop} (היה 117)`);
  ok('גם סטטוס 04',r.s04.ss===0&&r.s04.rop===0,`SS=${r.s04.ss} ROP=${r.s04.rop}`);
  ok('גם המרכיב הסטטיסטי עצמו מתאפס — לא רק הרצפה',r.s14.stat===0,`ssStat=${r.s14.stat} (היה 66)`);
  ok('ומקור ה-SS אומר למה',r.s14.src==='nostock',r.s14.src);
  /* הבקרה: פריט פעיל עם אותם נתונים בדיוק */
  ok('פריט פעיל עם אותם נתונים לא נגע',r.ok.ss===66&&r.ok.rop===117,
    `SS=${r.ok.ss} ROP=${r.ok.rop}`);
  /* ROP_MIN_ACT — מדיניות מאושרת, חלה גם כאן */
  ok('נוכחות מינימלית נשמרת: ROP=1 ב-SAP אינו יורד ל-0',
    r.min1.rop===1&&r.min1.ss===1,`ROP ${r.min1.rop0}→${r.min1.rop} · SS ${r.min1.ss0}→${r.min1.ss}`);
  /* הדלתות שנגזרות מ-sugROP */
  ok('פריט מוחרג אינו נכנס לדלת רצפת הביטחון',!r.floorDoor.includes('S14-LUMPY'),r.floorDoor.join(','));
  ok('ולא לדלתות המגמה והעלייה',
    !r.trendDoor.includes('S14-LUMPY')&&!r.riseDoor.includes('S14-LUMPY'),
    `ירידה=[${r.trendDoor}] עלייה=[${r.riseDoor}]`);

  /* ================= 2 · אותה החרגה בכל מסלול ================= */
  const w=await p.evaluate(()=>{
    const short=new Set(decisionList('short').map(r=>r.pn));
    LINE_M=null;const m=lineModel();
    const dry=new Set(m.dry.map(r=>r.pn));
    return {s04:{short:short.has('S04-CUST'),dry:dry.has('S04-CUST')},
      s14:{short:short.has('S14-DRY'),dry:dry.has('S14-DRY')},
      m1:ALL.filter(x=>isM1(x)).map(x=>x.pn)}});
  ok('פריט 04 עם לקוח ממתין אינו ב«לטיפול היום»',!w.s04.short,`short=${w.s04.short} (היה כן)`);
  ok('ואינו על הציר',!w.s04.dry,`dry=${w.s04.dry} (היה כן)`);
  ok('פריט 14 אינו באף אחד מהם',!w.s14.short&&!w.s14.dry,
    `short=${w.s14.short} dry=${w.s14.dry}`);
  ok('m1Active המשיך להחריג אותם כמו קודם',
    !w.m1.includes('S04-CUST')&&!w.m1.includes('S14-DRY'),w.m1.join(','));

  /* ================= 3 · ההחרגה אינה הסתרה ================= */
  const vis=await p.evaluate(()=>{const r=ALL.find(x=>x.pn==='S04-CUST');
    return {why:(r.why||[]).map(w=>w[1]).join(' | '),act:(r.act||[]).join(' | '),
      inCatalog:(QF.all||[]).some(x=>x.pn==='S04-CUST'),cap:r.expCap}});
  ok('ההזמנה הפתוחה נאמרת בכרטיס',/12 יח׳ בהזמנת לקוח פתוחה/.test(vis.why),
    (vis.why.match(/\d+ יח׳ בהזמנת לקוח[^|]*/)||[''])[0].slice(0,90));
  ok('והפעולה היא טיפול בלקוח, לא רכש',
    /אספקה חלופית או עדכון סטטוס/.test(vis.act)&&!/לפתוח רכש|להעלות נקודת הזמנה/.test(vis.act),
    vis.act.slice(0,90));
  ok('והפריט נשאר בקטלוג ובהון הכלוא',vis.inCatalog&&vis.cap>0,`כסף תקוע=${vis.cap}`);

  /* ================= 4 · חסר אינו אפס ================= */
  const pd=await p.evaluate(()=>{const g=pn=>ALL.find(x=>x.pn===pn);
    const f=pn=>{const r=g(pn);return {why:(r.why||[]).map(w=>w[1]).join(' | '),
      act:(r.act||[]).join(' | '),priceMissing:!!r.priceMissing,warrantyKnown:!!r.warrantyKnown}};
    return {no:f('PD-NOPRICE'),yes:f('PD-PRICED')}});
  /* החוק PD→ZM-C (צריכה מעל 3, מחיר עד 2,000, בלי תביעות אחריות) הוחלף
     בהחלטת המעתד בכלל תדירות אחד לכל הפרופילים — ראה PROFILE. מה שנבדק
     כאן הוא הכלל החדש, ומה שנשאר מעיקרון «חסר אינו אפס»: הכלל החדש אינו
     נשען על מחיר או על אחריות, ולכן הם אינם חוסמים ואינם מכריעים. */
  ok('המשפט של החוק הישן אינו מופיע עוד',
    !/עומד בתנאי המלצה ל-ZM-C/.test(pd.no.why+pd.yes.why)&&!/PD ל-ZM-C/.test(pd.no.act+pd.yes.act),
    pd.no.why.slice(0,60));
  const pr=await p.evaluate(()=>{const g=pn=>ALL.find(x=>x.pn===pn);
    const t=pn=>{const r=g(pn),i=profInfo(r);return {exp:profTgt(profH(r)),got:i?i.tgt:'D',h:profH(r)}};
    const m=n=>Array.from({length:11},(_,i)=>i<n?3:0);
    const base={mrp:'ZM',abc:'C',cust:0,po:0,noCons3y:true,pn:'T-PROF'};
    return {no:t('PD-NOPRICE'),yes:t('PD-PRICED'),
      th:[11,9,8,5,4,2,1,0].map(profTgt).join(''),
      up:profInfo({...base,months:m(10)}).state,
      cust:profInfo({...base,months:m(0),cust:2}).state,
      c3y:profInfo({...base,months:m(0),noCons3y:false}).state,
      po:profInfo({...base,months:m(0),po:4}).state,
      wait:profInfo({...base,months:m(0)}).state,
      same:profInfo({...base,months:m(3)})}});
  ok('מחיר חסר אינו חוסם ואינו מכריע את הפרופיל',pr.no.got===pr.no.exp,JSON.stringify(pr.no));
  ok('ופריט עם מחיר מקבל את אותו כלל בדיוק',pr.yes.got===pr.yes.exp,JSON.stringify(pr.yes));
  ok('ספי התדירות: 11/9→A · 8/5→B · 4/2→C · 1/0→D',pr.th==='AABBCCDD',pr.th);
  ok('העלאה — מיד',pr.up==='up',pr.up);
  ok('הורדה ל-D עם הזמנת לקוח פתוחה — לא',pr.cust==='keep',pr.cust);
  ok('הורדה ל-D כשהייתה צריכה ב-3 שנים — לא',pr.c3y==='keep',pr.c3y);
  ok('הורדה ל-D עם רכש פתוח — «לבדוק»',pr.po==='check',pr.po);
  ok('הורדה בלי חודש קודם מתחת לסף — «בהמתנה»',pr.wait==='wait',pr.wait);
  ok('פרופיל שתואם את התדירות אינו ברשימה',pr.same===null,String(pr.same));
  await p.close()}

 /* עמודת אחריות חסרה — ברמת הדוח */
 {const p=await load(SD+'/status-b.xlsx');
  const r=await p.evaluate(()=>{const x=ALL.find(y=>y.pn==='PD-PRICED');
    return {known:!!x.warrantyKnown,why:(x.why||[]).map(w=>w[1]).join(' | ')}});
  ok('עמודת אחריות שאינה בדוח מסומנת כלא-ידועה',!r.known,`warrantyKnown=${r.known}`);
  /* בכלל הפרופיל החדש אחריות אינה תנאי, ולכן עמודה חסרה אינה חוסמת דבר */
  const pw=await p.evaluate(()=>{const x=ALL.find(y=>y.pn==='PD-PRICED'),i=profInfo(x);
    return {exp:profTgt(profH(x)),got:i?i.tgt:'D'}});
  ok('ואינה משנה את המלצת הפרופיל',pw.got===pw.exp&&!/עומד בתנאי המלצה/.test(r.why),JSON.stringify(pw));
  await p.close()}

 /* ================= 5 · מלאי מת ודאי מול מועמד ================= */
 {const p=await load(SD+'/status-c.xlsx');
  const a=await p.evaluate(()=>{const r=ALL.find(x=>x.pn==='DEAD-OLD');
    return {has3y:!!r.has3y,cat:r.cat,act:(r.act||[])[0],why:(r.why||[]).map(w=>w[1])[0]}});
  ok('עם עמודות הצריכה השנתית — הסיווג נשאר ודאי',
    a.has3y&&a.cat==='מלאי מת'&&/לבחון חיסול/.test(a.act),`${a.cat} · ${a.act}`);
  await p.close();
  const p2=await load(SD+'/status-d.xlsx');
  const c=await p2.evaluate(()=>{const r=ALL.find(x=>x.pn==='DEAD-OLD');
    return {has3y:!!r.has3y,cat:r.cat,tier:r.ageTier,act:(r.act||[])[0],
      why:(r.why||[]).map(w=>w[1])[0],cap:r.expCap}});
  ok('בלי העמודות — מועמד לבדיקה ולא קביעה',c.cat==='מלאי מת — לבדיקה',c.cat);
  ok('והנימוק אינו מצטט מדיניות שלא נבדקה',
    !/לפי מדיניות 3 השנים/.test(c.why)&&/מדיניות 3 השנים לא נבדקה/.test(c.why),c.why.slice(0,110));
  ok('והפעולה היא אימות לפני חיסול',/לאמת מול צריכת 3 השנים/.test(c.act),c.act.slice(0,70));
  ok('ההון הכלוא ממשיך להיספר — זו שאלת ודאות, לא הסתרה',c.cap>0,`כסף תקוע=${c.cap}`);
  await p2.close()}

 /* ================= 6 · הזמנות חסומות ================= */
 {const p=await load(SD+'/status-e.xlsx');
  const r=await p.evaluate(()=>{const x=ALL.find(y=>y.pn==='CUST-RAW');
    const h=XCOLS.map(c=>c[0]);
    return {known:!!x.custBlockedKnown,cust:x.cust,raw:x.custRaw,
      exp:XCOLS[h.indexOf('הזמנות חסומות')][1](x),
      warn:(DIAG.warn||[]).filter(w=>/חסומ/.test(w))[0]||''}});
  ok('היעדר עמודת ההזמנות החסומות מסומן',!r.known,`custBlockedKnown=${r.known}`);
  ok('ויש אזהרה ייעודית',/אינה בדוח/.test(r.warn),r.warn.slice(0,90));
  ok('והאזהרה אומרת שהפער הוא חסם עליון',/חסם עליון/.test(r.warn));
  ok('והייצוא מציג «—» ולא 0',r.exp==='—',`ערך בייצוא="${r.exp}"`);
  await p.close()}

 /* ================= 7 · תוויות חלון הצריכה ================= */
 {const p=await load(SD+'/status-a.xlsx');
  const txt=await p.evaluate(()=>{
    const all=ALL.map(r=>[...(r.why||[]).map(w=>w[1]),...(r.act||[])].join(' ')).join(' ');
    return all});
  ok('אין טקסט שאומר «12 החודשים» על חלון של 11',!/ב-12 החודשים/.test(txt));
  ok('ואין «יח׳/שנה» על אותו חלון',!/יח׳\/שנה/.test(txt));
  await p.close()}

 /* ================= 8 · שעון המלאי המת ================= */
 {const p=await load(SD+'/status-f.xlsx');
  const g=await p.evaluate(()=>{
   const o={};for(const r of ALL)o[r.pn]={tier:r.ageTier,cat:r.cat,
     entAge:r.entAge,saleAge:r.saleAge,expCap:r.expCap||0,
     why:((r.why||[])[0]||[])[1]||'',act:(r.act||[])[1]||''};
   return o});
  ok('מכולה טרייה אינה מוחקת מלאי מת — מכירה לפני 7 שנים מכריעה',
    g['CLK-FRESH-ENTRY'].tier==='dead',
    `tier=${g['CLK-FRESH-ENTRY'].tier} · entAge=${g['CLK-FRESH-ENTRY'].entAge} · saleAge=${g['CLK-FRESH-ENTRY'].saleAge}`);
  ok('והכסף שלו נספר — expCap אינו אפס',
    g['CLK-FRESH-ENTRY'].expCap>0,`expCap=${g['CLK-FRESH-ENTRY'].expCap}`);
  ok('מכירה לפני 100 יום ⇒ טרם מת',
    g['CLK-RECENT-SALE'].tier==='young',g['CLK-RECENT-SALE'].tier);
  ok('והטקסט מצטט את המכירה, לא את הכניסה שלפני 7 שנים',
    /המכירה האחרונה/.test(g['CLK-RECENT-SALE'].why)&&!/הכניסה האחרונה/.test(g['CLK-RECENT-SALE'].why),
    g['CLK-RECENT-SALE'].why.slice(0,80));
  ok('ואינו אומר «בעוד 0 חודשים» על שעון שרק התחיל',
    !/בעוד 0 חודשים/.test(g['CLK-RECENT-SALE'].act),g['CLK-RECENT-SALE'].act);
  ok('בלי מכירה — השעון הוא הכניסה',
    /הכניסה האחרונה/.test(g['CLK-ENTRY-ONLY'].why),g['CLK-ENTRY-ONLY'].why.slice(0,70));
  ok('בלי שום תאריך — הכלי אומר שאין שעון ואינו ממציא מספר',
    !/לפני —/.test(g['CLK-NO-DATES'].why)&&!/בעוד 37 חודשים/.test(g['CLK-NO-DATES'].act)
    &&/אי אפשר לקבוע/.test(g['CLK-NO-DATES'].why),
    `${g['CLK-NO-DATES'].why.slice(0,60)} // ${g['CLK-NO-DATES'].act.slice(0,50)}`);
  ok('ומבקש להשלים את התאריך במקום להמתין',
    /להשלים תאריך/.test(g['CLK-NO-DATES'].act),g['CLK-NO-DATES'].act);
  ok('רגרסיה: פריט עם צריכה אינו נכנס לשום ענף של השעון',
    g['CLK-ALIVE'].tier!=='dead'&&g['CLK-ALIVE'].tier!=='young',g['CLK-ALIVE'].tier);
  await p.close()}

 /* ================= 7 · קצב מהחלון התלת-שנתי ================= */
 {const p=await load(SD+'/status-g.xlsx');
  const g=await p.evaluate(()=>{const o={};for(const r of ALL)o[r.pn]={
    cat:r.cat,rate:+(r.rate||0).toFixed(3),from3y:!!r.rateFrom3y,
    uns:+(r.unsQty||0).toFixed(3),gross:+(r.unsGross||0).toFixed(3),
    inToday:inToday(r),dry:lineDry(r),
    tRatio:r.tRatio,trendF:+(r.trendF||0).toFixed(2),tgap:trendGap(r),
    why2:(r.why||[]).map(w=>w&&w[1]).filter(Boolean).join(' ¦ '),
    cov:r.cov==null?null:+r.cov.toFixed(1),rop:r.rop,sugROP:r.sugROP,
    excess:r.excessQty||0,pf:!!r.paramFix,
    why:(r.why||[]).map(w=>w&&w[1]).filter(Boolean).join(' ¦ '),
    act:(r.act||[]).filter(Boolean).join(' ¦ ')};return o});
  ok('צריכה תלת-שנתית מעל הסף נותנת קצב, גם כשהחלון הקצר אפס',
    g['R3-FAT'].rate>0&&g['R3-FAT'].from3y,`rate=${g['R3-FAT'].rate} from3y=${g['R3-FAT'].from3y}`);
  ok('ונקודת ההזמנה המוצעת כבר אינה נשענת על רצפות בלבד',
    g['R3-FAT'].sugROP>g['R3-THIN'].sugROP,
    `FAT=${g['R3-FAT'].sugROP} THIN=${g['R3-THIN'].sugROP}`);
  ok('יחידה אחת בשלוש שנים אינה קצב — הסף חוסם',
    g['R3-THIN'].rate===0&&!g['R3-THIN'].from3y,`rate=${g['R3-THIN'].rate}`);
  /* הרגרסיה שהמדידה תפסה: 45 פריטים עברו איטי→עודף בגלל הסרגל */
  ok('קצב תלת-שנתי נשפט בסרגל 24 של הענף האיטי, לא בסרגל 12 של העודף',
    g['R3-SLOW'].cat==='מלאי איטי',`cat=${g['R3-SLOW'].cat} cov=${g['R3-SLOW'].cov}`);
  ok('וכיסוי בין 12 ל-24 אינו מייצר סיווג חדש',
    g['R3-MID'].cat!=='עודף מלאי'&&g['R3-MID'].excess===0,
    `cat=${g['R3-MID'].cat} cov=${g['R3-MID'].cov} עודף=${g['R3-MID'].excess}`);
  /* הרגרסיה השנייה: הפרוקסי rate>0 בשני המעברים המשלימים */
  ok('האזהרה «SAP יזמין בחזרה» שורדת את הקצב החדש',
    g['R3-BACK'].pf&&/יזמין בחזרה|אחרת המלאי יחזור/.test(g['R3-BACK'].why+g['R3-BACK'].act),
    `paramFix=${g['R3-BACK'].pf}`);
  ok('וגם הצד ההפוך — פרמטר מתחת לרצפה אינו מאבד את הדגל',
    g['R3-FLOOR'].pf,`paramFix=${g['R3-FLOOR'].pf} ROP ${g['R3-FLOOR'].rop}→${g['R3-FLOOR'].sugROP}`);
  /* ============ קצב תלת-שנתי אינו תחזית לחודש ============
     השלישית מבין הרגרסיות שהמדידה תפסה: `unsQty = קצב − מלאי פנוי`
     הוא משפט על החודש הקרוב, ועם קצב של 0.30 הוא טוען «חוסר» של
     שליש יחידה ומכניס את הפריט לדלת הבוערת. שני השערים נדרשים —
     `unsQty` לבדו מתקן את «היום» ומשאיר את הציר. */
  ok('קצב תלת-שנתי אינו מייצר חוסר לחודש הקרוב, גם על מדף ריק',
    g['R3-DRY'].uns===0&&g['R3-DRY'].gross===0,
    `unsQty=${g['R3-DRY'].uns} unsGross=${g['R3-DRY'].gross} קצב=${g['R3-DRY'].rate}`);
  ok('ולכן אינו נכנס לדלת הבוערת ולא לציר האספקות',
    !g['R3-DRY'].inToday&&!g['R3-DRY'].dry,
    `inToday=${g['R3-DRY'].inToday} lineDry=${g['R3-DRY'].dry}`);
  ok('רגרסיה: ביקוש אמיתי על אותו מדף ריק כן מייצר חוסר לחודש',
    g['R3-ACTDRY'].uns>0&&!g['R3-ACTDRY'].from3y&&g['R3-ACTDRY'].dry,
    `uns=${g['R3-ACTDRY'].uns} from3y=${g['R3-ACTDRY'].from3y} lineDry=${g['R3-ACTDRY'].dry}`);
  /* ============ המגמה נמדדת על החלון שהקצב בא ממנו ============
     הרביעית מבין הרגרסיות. הדלת «ביקוש יורד» כותרתה «הביקוש ירד בתוך
     ששת החודשים שמהם נגזר הקצב», והקצב התלת-שנתי לא נגזר מהם. 14
     פריטים נכנסו לדלת על הדוח האמיתי. */
  ok('מגמה אינה מרסנת קצב שלא בא מחלון המגמה',
    g['R3-TREND'].trendF===1&&g['R3-TREND'].tgap===0,
    `tRatio=${g['R3-TREND'].tRatio} trendF=${g['R3-TREND'].trendF} פער=${g['R3-TREND'].tgap}`);
  ok('רגרסיה: ירידה אמיתית בחלון אמיתי כן מתרסנת',
    g['R3-FALL'].trendF<1&&!g['R3-FALL'].from3y,
    `trendF=${g['R3-FALL'].trendF} tRatio=${g['R3-FALL'].tRatio} from3y=${g['R3-FALL'].from3y}`);
  /* ============ «שני שעונים» נוקב בחלון ובבסיס הכיסוי ============
     החמישית. לפני ג5 נפלו 28 הפריטים האלה לענף השטוח, שנקב במפורש
     ב«כיסוי בקצב 3 השנים». מרגע שיש להם rate ו-cov הם עברו לענף
     המדורג, ששתק על החלון — כלומר התיקון החליף נוסח מדויק בשותק. */
  ok('«שני שעונים» על קצב תלת-שנתי נוקב בחלון',
    /שני שעונים: ROP מוצע/.test(g['R3-TWOCLK'].why2)
    &&/קצב [\d.]+ לחודש מחלון 3 השנים/.test(g['R3-TWOCLK'].why2),
    g['R3-TWOCLK'].why2.slice(0,150));
  /* בלי לקבע את המספר: win3y נגזר מחודש העוגן ולכן נע. מה שנעול הוא
     ההבחנה — שני כיסויים באותו כרטיס, כל אחד מסומן על הכמות שלו. */
  ok('ומסמן שהכיסוי הוא של המלאי הפנוי — לא מול מספר אחר באותו כרטיס',
    (()=>{const t=g['R3-TWOCLK'].why2;
      const slow=t.match(/כיסוי (\d+) חודשים בקצב הזה/);
      const freeCov=/על המלאי הפנוי 0 חודשי כיסוי/.test(t);
      return !!slow&&+slow[1]>0&&freeCov})(),
    g['R3-TWOCLK'].why2.slice(0,300));
  /* הבקרה: פריט עם ביקוש אמיתי לא נגע */
  ok('רגרסיה: ביקוש בחלון הקצר עדיין נשפט בסרגל 12 ⇒ עודף מלאי',
    g['R3-ACTIVE'].cat==='עודף מלאי'&&!g['R3-ACTIVE'].from3y,
    `cat=${g['R3-ACTIVE'].cat} from3y=${g['R3-ACTIVE'].from3y} cov=${g['R3-ACTIVE'].cov}`);
  await p.close()}

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 console.log(out.join('\n'));
 const f=out.filter(x=>x.startsWith('FAIL')).length;
 console.log(`\n${out.length-f}/${out.length} עברו`);
 await b.close();
 process.exit(bad?1:0)})();
