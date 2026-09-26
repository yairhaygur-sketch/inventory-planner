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
  D(o.saleAgo!=null?o.saleAgo:10),D(o.entAgo!=null?o.entAgo:60)]};
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
  ok('והפריט נשאר בקטלוג ובהון הכלוא',vis.inCatalog&&vis.cap>0,`הון כלוא=${vis.cap}`);

  /* ================= 4 · חסר אינו אפס ================= */
  const pd=await p.evaluate(()=>{const g=pn=>ALL.find(x=>x.pn===pn);
    const f=pn=>{const r=g(pn);return {why:(r.why||[]).map(w=>w[1]).join(' | '),
      act:(r.act||[]).join(' | '),priceMissing:!!r.priceMissing,warrantyKnown:!!r.warrantyKnown}};
    return {no:f('PD-NOPRICE'),yes:f('PD-PRICED')}});
  ok('PD בלי מחיר אינו מוצג ככשיר ל-ZM-C',!/עומד בתנאי המלצה ל-ZM-C/.test(pd.no.why),
    pd.no.why.slice(0,80));
  ok('ונאמר במפורש מה חסר',/חסר מחיר FOB/.test(pd.no.why),
    (pd.no.why.match(/לא ניתן להכריע[^|]*/)||[''])[0].slice(0,100));
  ok('וגם אינו נפסל — ההכרעה נחסמה, לא הוכרעה',
    !/נשאר PD עד שיתקיימו/.test(pd.no.why)&&/להשלים מחיר FOB/.test(pd.no.act),pd.no.act.slice(0,70));
  ok('PD עם מחיר ממשיך לעבוד כרגיל',/עומד בתנאי המלצה ל-ZM-C/.test(pd.yes.why),
    pd.yes.why.slice(0,70));
  await p.close()}

 /* עמודת אחריות חסרה — ברמת הדוח */
 {const p=await load(SD+'/status-b.xlsx');
  const r=await p.evaluate(()=>{const x=ALL.find(y=>y.pn==='PD-PRICED');
    return {known:!!x.warrantyKnown,why:(x.why||[]).map(w=>w[1]).join(' | ')}});
  ok('עמודת אחריות שאינה בדוח מסומנת כלא-ידועה',!r.known,`warrantyKnown=${r.known}`);
  ok('וחוסמת את ההכרעה על ZM-C',
    /חסר עמודת תביעות אחריות/.test(r.why)&&!/עומד בתנאי המלצה/.test(r.why),
    (r.why.match(/לא ניתן להכריע[^|]*/)||[''])[0].slice(0,100));
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
  ok('ההון הכלוא ממשיך להיספר — זו שאלת ודאות, לא הסתרה',c.cap>0,`הון כלוא=${c.cap}`);
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

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 console.log(out.join('\n'));
 const f=out.filter(x=>x.startsWith('FAIL')).length;
 console.log(`\n${out.length-f}/${out.length} עברו`);
 await b.close();
 process.exit(bad?1:0)})();
