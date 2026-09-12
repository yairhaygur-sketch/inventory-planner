const XLSX=require('xlsx'),{chromium}=require('playwright'),fs=require('fs'),path=require('path');
const SD=__dirname, MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const hdr=['מק"ט מוביל','תיאור חומר','תיאור חומר2','שם ספק','סטטוס חומר','תיאור','סוג MRP','ABC','רמת שרות',
 'מלאי בטחון','נק.הז.מחדש','אספ.מתוכנ.','זמ.עב.קבלת','מל.בט.מינ.','מחיר FOB','מטבע FOB','סה"כ מלאי','מלאי פנוי',
 'מלאי מרלוג','הז. רכש','בהעברה','אספקות פת.','כמות בהז.פ','סוג חומר','תיא.קבוצ.חומרים','קב.חו.חיצו','טקסט ארוך',
 'תב.אח.הש.','ת.היר.1 מח','ת.היר.2 מח','היררכייה1','היררכייה1','היררכייה2','היררכייה2','היררכייה3','היררכייה3',
 'צר.השנה','צר.שנה-1','צר.שנה-2','צר.החודש',...MON,"תאר' מכירה","תאר' כניסה"];
const D=n=>{const d=new Date(Date.now()-n*864e5);return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`};
// months[0] = חודש-1 (הקרוב ביותר)
const mk=(pn,o)=>{const m=o.months;const y=m.reduce((a,b)=>a+b,0);
 return [pn,'פריט מבחן '+pn,'Test '+pn,'ספק','01','פעיל',o.mrp||'ND','A',o.srv||95,
  o.ss||0,o.rop||0,o.lt||30,2,0,o.price!=null?o.price:100,o.cur||'USD',o.free,o.free,0,o.po||0,0,0,o.cust||0,'Z004','מנוע','ZT','מתכנן','0',
  'שיווק','מערכת','100','ZT','200','דגם','300','מערכת',
  o.y0!=null?o.y0:y, o.y1!=null?o.y1:y, o.y2!=null?o.y2:y, 0, ...m, D(o.saleAgo||10), D(o.entAgo||60)]};

const rows=[
 // A · סטיית תקן: ביקוש שטוח לגמרי, חודש אזילה בודד
 mk('FLAT-CLEAN',{months:[10,10,10,10,10,10,10,10,10,10,10],free:60}),
 mk('FLAT-STOCKOUT',{months:[10,10,10,10,10,0,10,10,10,10,10],free:60}),
 // B · אזילה של 3 חודשים רצופים באמצע, ביקוש בריא סביב
 mk('RUN-STOCKOUT',{months:[20,20,20,0,0,0,20,20,20,20,20],free:100}),
 // B1 · אותה אזילה בדיוק, אבל הפריט חשוף גם היום — כאן זו כן מניעת חוסר
 mk('RUN-STOCKOUT-OPEN',{months:[20,20,20,0,0,0,20,20,20,20,20],free:4,lt:90}),
 // C · ביקוש 2 יח׳/שנה, מעט מלאי, ROP ענק ב-SAP
 mk('LOW-DEMAND',{months:[0,0,1,0,0,0,1,0,0,0,0],y0:2,y1:2,y2:2,free:3,rop:200,ss:80}),
 // C1 · אותו פריט בדיוק, אבל 5 דולר ליחידה — רצפת מדיניות של 1 במקום אפס
 mk('LOW-CHEAP',{months:[0,0,1,0,0,0,1,0,0,0,0],y0:2,y1:2,y2:2,free:3,rop:200,ss:80,price:5,lt:120}),
 // C2 · זול, ביקוש זעום, ובלי שום נוכחות ב-SAP — הצד השני של אותה מדיניות
 mk('LOW-CHEAP-ZERO',{months:[0,0,1,0,0,0,1,0,0,0,0],y0:2,y1:2,y2:2,free:0,rop:0,ss:0,price:5,lt:120}),
 // C4 · מלאי איטי עם נקודת הזמנה שגבוהה מכל מה שנמכר בשלוש שנים —
 //      הקטנת המלאי לבדה לא תחזיק כי SAP יזמין בחזרה
 mk('SLOW-BAD-ROP',{months:[0,0,0,0,0,0,0,0,0,0,0],y0:2,y1:0,y2:0,free:198,rop:6,ss:0,price:0.18}),
 mk('SLOW-OK-ROP',{months:[0,0,0,0,0,0,0,0,0,0,0],y0:2,y1:0,y2:0,free:198,rop:1,ss:0,price:0.18}),
 // C3 · אותו מחיר בדיוק בשני מטבעות — הסף לכל מטבע בנפרד, בלי המרה
 mk('LOW-150-ILS',{months:[0,0,1,0,0,0,1,0,0,0,0],y0:2,y1:2,y2:2,free:3,rop:200,ss:80,price:150,cur:'ILS'}),
 mk('LOW-150-USD',{months:[0,0,1,0,0,0,1,0,0,0,0],y0:2,y1:2,y2:2,free:3,rop:200,ss:80,price:150,cur:'USD'}),
 // D · ביקוש בירידה חדה שבה ה-ROP המוצע דווקא גבוה מהקיים
 mk('DECLINE',{months:[2,2,3,2,3,20,25,22,24,23,25],free:40,rop:1,ss:0,lt:120}),
 // B2 · ספורדי: חודשי האפס הם היעדר ביקוש, לא אזילה
 mk('SPORADIC',{months:[0,0,10,0,0,0,0,10,0,0,10],free:200}),
 mk('EDGE-ONLY',{months:[5,0,0,0,0,0,0,0,0,0,6],free:200}),
 // ETA · רכש פתוח שמכסה את הלקוח, מלאי אפס. הרכש חסר ETA ולכן אינו כיסוי:
 // ההחלטה חייבת להיות הכמות, לא "לבדוק Back Order אצל היצרן"
 mk('PO-NO-ETA',{months:[20,20,20,20,20,20,20,20,20,20,20],free:0,po:500,cust:10}),
 // אותו מבנה בדיוק אבל עם מלאי שנשחק — חומרה 2 ולא 3
 mk('PO-SOME-STOCK',{months:[20,20,20,20,20,20,20,20,20,20,20],free:12,po:500,cust:10}),
 // אותו מבנה בדיוק, אבל הצריכה מתחת לסף — נשאר מעקב אספקה
 mk('FOLLOW-OK',{months:[1,1,1,1,1,1,1,1,1,1,1],free:0,po:100,cust:5}),
 // COVERED · מלאי גדול, בלי הזמנות לקוח — כמו הדוח האמיתי שבו אין חוסרים
 mk('COVERED-1',{months:[8,8,8,8,8,8,8,8,8,8,8],free:900,cust:0,rop:5,ss:2}),
 mk('COVERED-2',{months:[6,6,6,6,6,6,6,6,6,6,6],free:800,cust:0,rop:5,ss:2}),
 // מטבעות — עודף מלאי זהה בשלושה מטבעות שונים
 mk('CUR-USD',{months:[2,2,2,2,2,2,2,2,2,2,2],free:500,cur:'USD'}),
 mk('CUR-EUR',{months:[2,2,2,2,2,2,2,2,2,2,2],free:500,cur:'EUR'}),
 mk('CUR-ILS',{months:[2,2,2,2,2,2,2,2,2,2,2],free:500,cur:'ILS'}),
 // שני פריטי דולר בשווי שונה — לבדיקת סדר יורד בתוך בלוק המטבע
 mk('CUR-USD2',{months:[2,2,2,2,2,2,2,2,2,2,2],free:500,cur:'USD',price:250}),
 // SAP ss=1 ו-rop=3 בלי קצב: הרצפות מרימות SS, והכלל SS≤ROP חייב לרוץ אחריהן
 mk('SSROP-GUARD',{months:[0,0,1,0,0,0,0,0,0,0,0],y0:1,y1:0,y2:0,free:3,rop:3,ss:1,price:5000,cur:'USD'}),
];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['ZMRP'],[],hdr,...rows]),'ZMRP');return wb})(),
 SD+'/defects.xlsx');

(async()=>{
 const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html'));
 await p.setInputFiles('#f',SD+'/defects.xlsx');await p.waitForTimeout(2200);
 const o=await p.evaluate(()=>{const g=pn=>{const r=ALL.find(x=>x.pn===pn);if(!r)return null;
   const q=Object.keys(Q).find(k=>Q[k].includes(r))||'—';
   return {cat:r.cat,sd:+r.A.sd.toFixed(2),drops:r.A.drops.length,sugSS:r.sugSS,sugROP:r.sugROP,rop:r.rop,ss:r.ss,lowFloor:r.lowFloor,
     act:(r.act||[])[0]||'',acts:(r.act||[]).join(' | '),actN:(r.act||[]).length,paramFix:r.paramFix,why:(r.why||[]).map(w=>w[1]).join(' | '),q,sev:r.sev,cov:r.covA,trend:r.A.trend.pct}};
  return {clean:g('FLAT-CLEAN'),so:g('FLAT-STOCKOUT'),run:g('RUN-STOCKOUT'),
          runOpen:g('RUN-STOCKOUT-OPEN'),eta:g('PO-NO-ETA'),etaWet:g('PO-SOME-STOCK'),fok:g('FOLLOW-OK'),low:g('LOW-DEMAND'),slowBad:g('SLOW-BAD-ROP'),slowOk:g('SLOW-OK-ROP'),cheap:g('LOW-CHEAP'),cheap0:g('LOW-CHEAP-ZERO'),ils150:g('LOW-150-ILS'),usd150:g('LOW-150-USD'),
          dec:g('DECLINE'),spor:g('SPORADIC'),edge:g('EDGE-ONLY'),
          guard:g('SSROP-GUARD'),
          inv:{ssGtRop:ALL.filter(x=>x.sugSS>x.sugROP).length,
               header:(COLS.find(c=>c[0]==='risk')||[])[1]},
          rank:(()=>{const ex=Q.excess.filter(x=>x.expCap>0);
            const ord=[...ex].sort((a,b)=>cmpMoneyDesc(a,b,'expCap'))
              .map(x=>({pn:x.pn,cur:curKey(x),cap:x.expCap}));
            const blocks=[];for(const r of ord){if(!blocks.length||blocks[blocks.length-1][0]!==r.cur)blocks.push([r.cur,[]]);
              blocks[blocks.length-1][1].push(r)}
            return {ord,blocks:blocks.map(([c,l])=>[c,l.length]),
              desc:blocks.every(([c,l])=>l.every((r,i)=>i===0||l[i-1].cap>=r.cap)),
              cut:cut80(ex.slice().sort((a,b)=>cmpMoneyDesc(a,b,'expCap')),'cap')}})(),
          cur:{usd:g('CUR-USD'),eur:g('CUR-EUR'),ils:g('CUR-ILS')}}});
 const out=[],ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));
 // A · חודש אזילה אינו יוצר מלאי ביטחון יש מאין
 ok('ביקוש שטוח → סטיית תקן 0',o.clean.sd===0&&o.clean.sugSS===0);
 ok('חודש אזילה אינו מנפח את סטיית התקן',o.so.sd===0&&o.so.sugSS===0,
    `sd=${o.so.sd} SS=${o.so.sugSS} (לפני התיקון: 3.02 / 6)`);
 // B · רצף אזילה מזוהה ואינו נראה כירידת ביקוש
 ok('אזילה של 3 חודשים מזוהה',o.run.drops===3,`drops=${o.run.drops} (לפני: 0)`);
 ok('רצף אזילה אינו מוצג כירידה בביקוש',o.run.trend===0,`מגמה ${o.run.trend}% (לפני: -40%)`);
 // B1 · אזילה בעבר מול חשיפה היום — שני מצבים שונים, שני מסלולים שונים
 ok('אזילה בעבר + מכוסה היום → לא במניעת חוסר',o.run.q!=='prevent',
    `מסלול ${o.run.q} · ${o.run.cat} (לפני: prevent · חשד אזילה)`);
 ok('ומקבל את הלקח הפרמטרי כהחלטה הראשונה',/נקודת הזמנה|מלאי ביטחון|אין פעולה/.test(o.run.act),o.run.act);
 ok('אימות מול המוסך יורד ממקום ההחלטה',!/סגירת מוסך/.test(o.run.act)&&/סגירת מוסך/.test(o.run.acts),o.run.act);
 ok('אינו מדורג כחומרה 2',o.run.sev<2,`sev=${o.run.sev} (לפני: 2)`);
 ok('אותה אזילה כשהפריט חשוף היום → כן חשד אזילה',
    o.runOpen.q==='prevent'&&o.runOpen.cat==='חשד אזילה',`${o.runOpen.q} · ${o.runOpen.cat}`);
 ok('ואז ההחלטה היא השלמת הכמות',/להשלים/.test(o.runOpen.act),o.runOpen.act);
 ok('ההודעה מציינת את אורך הרצף',/3 החודשים/.test(o.runOpen.acts),o.runOpen.acts);
 // B2 · פריט ספורדי אינו אזילה — חודשי אפס הם היעדר ביקוש, לא חוסר מלאי
 ok('פריט ספורדי (3 מתוך 11) אינו מסווג אזילה',o.spor.drops===0,
    `drops=${o.spor.drops} · ${o.spor.cat}`);
 ok('פריט בשני קצוות בלבד אינו מסווג אזילה',o.edge.drops===0,
    `drops=${o.edge.drops} · ${o.edge.cat}`);
 // E · רכש פתוח בלי ETA אינו כיסוי — ההחלטה היא הכמות, לא בירור מול הספק
 ok('מלאי שלא מכסה חודש אינו מסווג מעקב אספקה',o.eta.q!=='follow'&&o.eta.cat!=='מעקב אספקה',
    `${o.eta.q} · ${o.eta.cat} (לפני: follow · מעקב אספקה)`);
 ok('וההחלטה היא הכמות החסרה',/^להשלים \d+ יח׳/.test(o.eta.act),o.eta.act);
 ok('בירור ה-ETA יורד לפעולה נלווית',!/ETA|Back Order/.test(o.eta.act)&&/ETA/.test(o.eta.acts),o.eta.act);
 ok('הלקוח הממתין והרכש הפתוח מוזכרים בהסבר',/500/.test(o.eta.acts),o.eta.acts.slice(0,60));
 // חומרה 3 רק כשאין מה למכור מחר בבוקר
 ok('"אוזל החודש" עם מלאי אפס מקבל חומרה 3',o.eta.sev===3,`sev=${o.eta.sev} (לפני: 2)`);
 ok('אותו מבנה עם מלאי שנשחק נשאר חומרה 2',o.etaWet.sev===2,`sev=${o.etaWet.sev} · ${o.etaWet.cat}`);
 ok('אותו מבנה מתחת לסף הצריכה נשאר מעקב אספקה',o.fok.cat==='מעקב אספקה',
    `${o.fok.q} · ${o.fok.cat}`);

 // C · ביקוש זעום עם פרמטרים גבוהים ב-SAP
 ok('פריט 2 יח׳/שנה עם ROP=200 אינו "תקין"',o.low.cat!=='תקין',o.low.cat);
 ok('ומוצע לו אפס',o.low.sugROP===0&&o.low.sugSS===0&&/ל-0/.test(o.low.act),o.low.act);
 // C1–C3 · רצפת מדיניות לביקוש זעום — ראה LOWDEM_CHEAP
 ok('פריט זול באותה להקה מקבל רצפה של 1 ולא אפס',
    o.cheap.sugROP===1&&o.cheap.sugSS===1,`sugROP=${o.cheap.sugROP} sugSS=${o.cheap.sugSS} (לפני: 0/0)`);
 ok('וההמלצה אומרת את המספר החדש',/ל-1/.test(o.cheap.act),o.cheap.act);
 ok('ההסבר מנמק במחיר וללא המרת מטבע',/\$5/.test(o.cheap.acts+o.cheap.why),o.cheap.why);
 ok('פריט זול בלי נוכחות ב-SAP מקבל שורת עבודה',
    o.cheap0.cat==='עדכון פרמטרים'&&o.cheap0.q==='quality',`${o.cheap0.q} · ${o.cheap0.cat} (לפני: תקין)`);
 ok('וההחלטה היא לקבוע 1',/לקבוע נקודת הזמנה 1/.test(o.cheap0.act),o.cheap0.act);
 ok('150 ש"ח נחשב זול — מתחת לסף ה-ILS',o.ils150.lowFloor===1,`lowFloor=${o.ils150.lowFloor}`);
 ok('150 דולר אינו זול — מעל סף ה-USD',o.usd150.lowFloor===0&&o.usd150.sugROP===0,
    `lowFloor=${o.usd150.lowFloor} sugROP=${o.usd150.sugROP}`);
 ok('אותו מספר, שני מטבעות, שתי תשובות — אין המרה',
    o.ils150.sugROP!==o.usd150.sugROP,`ILS→${o.ils150.sugROP} · USD→${o.usd150.sugROP}`);
 // C4 · הפרמטר שמחזיר את המלאי — ראה הפוסט-פאס ב-classify
 ok('מלאי איטי עם נקודת הזמנה חורגת נשאר באותו תור',o.slowBad.q==='excess',
    `${o.slowBad.q} · ${o.slowBad.cat}`);
 ok('אך מקבל גם את ההחלטה על הפרמטר',/לעדכן נקודת הזמנה/.test(o.slowBad.acts),o.slowBad.acts);
 ok('וההסבר אומר למה הקטנת מלאי לבדה לא תחזיק',/יזמין בחזרה/.test(o.slowBad.why),o.slowBad.why);
 ok('והוא נכנס למעקב היישום',o.slowBad.paramFix===true,`paramFix=${o.slowBad.paramFix}`);
 ok('אותו פריט עם נקודת הזמנה 1 אינו מקבל את ההערה',
    !/לעדכן נקודת הזמנה/.test(o.slowOk.acts)&&o.slowOk.paramFix!==true,
    `paramFix=${o.slowOk.paramFix} · ${o.slowOk.acts}`);
 ok('ההחלטות נשארות עד 4',o.slowBad.actN<=4,`${o.slowBad.actN} החלטות`);
 // D · כיוון ההמלצה נגזר מהמספרים
 ok('ROP מוצע גבוה → "העלאת"',/העלאת/.test(o.dec.act),o.dec.act);
 // מטבע — BZ נקוב במטבע של BE, ואסור להציג אותו כשקלים או לחבר מטבעות
 const cur=await p.evaluate(()=>{
  const cap=ALL.filter(r=>r.expCap>0);
  const one=ALL.find(r=>r.pn==='CUR-USD');
  return {itemCell:(()=>{const i=ALL.indexOf(one);return i<0?'':(one.currency||'')})(),
   mix:moneyMix(cap,r=>r.expCap),
   usdTxt:money(one.expCap,one.currency),
   ilsTxt:money(one.expCap,'ILS'),
   symbols:[...new Set(cap.map(r=>curSym(r.currency)))].sort().join('')}});
 // תקלה 1 · הכלל SS≤ROP רץ אחרי הרצפות
 ok('אין שום פריט שבו מלאי הביטחון המוצע גדול מנקודת ההזמנה המוצעת',
    o.inv.ssGtRop===0,`${o.inv.ssGtRop} חריגים`);
 ok('SAP ss=1 rop=3 בלי קצב → ההמלצה קוהרנטית',
    o.guard.sugSS<=o.guard.sugROP&&o.guard.sugROP>=1,
    `sugSS=${o.guard.sugSS} sugROP=${o.guard.sugROP} (לפני: 1 / 0)`);
 // תקלה 8 · כותרת ההון הכלוא בלי סימן מטבע
 ok('כותרת "הון כלוא" אינה נוקבת מטבע',!/[₪$€]/.test(o.inv.header||''),o.inv.header);
 // תקלה 2 · דירוג כספי בבלוקים לפי מטבע, בלי המרה
 ok('כל מטבע יושב בבלוק רציף אחד',
    o.rank.blocks.length===new Set(o.rank.blocks.map(b=>b[0])).size,
    JSON.stringify(o.rank.blocks));
 ok('בתוך כל בלוק הסדר יורד לפי הסכום',o.rank.desc,
    o.rank.ord.slice(0,4).map(r=>`${r.cur} ${r.cap}`).join(' · '));
 ok('קו ה-80% מחושב לכל מטבע בנפרד',
    !!o.rank.cut.byCur&&Object.keys(o.rank.cut.byCur).length>1,
    JSON.stringify(Object.entries(o.rank.cut.byCur||{}).map(([c,v])=>`${c}: ${v.n}/${v.count}`)));
 ok('פריט דולרי מוצג ב-$ ולא ב-₪',/^\$/.test(cur.usdTxt),`${cur.usdTxt} (היה ${cur.ilsTxt})`);
 ok('סכום מפוצל לפי מטבע ולא מחובר',cur.mix.split(' · ').length>1,cur.mix);
 ok('כל המטבעות בקטלוג מיוצגים',cur.symbols.length>1,cur.symbols);
 // פקיעת סימונים — ראה MARK_TTL
 const mkx=await p.evaluate(()=>{
  const t=ALL.find(x=>x.sev>0)||ALL[0], k=markKey(t), ago=d=>Date.now()-d*864e5;
  const snap=()=>{const r=ALL.find(x=>x.pn===t.pn);
    return {cat:r.cat,q:Object.keys(Q).find(z=>Q[z].includes(r)),stale:!!r.markStale,
      note:(r.why||[]).map(w=>w[1]).find(x=>/פג/.test(x))||''}};
  const run=(type,days)=>{MARKS[k]={t:type,ts:ago(days)};Q=classify(ALL);return snap()};
  const o={base:(()=>{delete MARKS[k];Q=classify(ALL);return snap()})(),
   ttl:{handled:markTTL('handled')/864e5,campaign:markTTL('campaign')/864e5,
        ignore:markTTL('ignore')/864e5,ondemand:markTTL('ondemand')/864e5},
   handledFresh:run('handled',2), handledStale:run('handled',10),
   campFresh:run('campaign',100), campStale:run('campaign',200),
   ignoreStale:run('ignore',200)};
  MARKS[k]={t:'ondemand',ts:ago(200)};Q=classify(ALL);
  o.odStale=(()=>{const r=ALL.find(x=>x.pn===t.pn);return {cat:r.cat,odStale:!!r.odStale}})();
  delete MARKS[k];Q=classify(ALL);o.after=snap();
  return o});
 ok('«טופל» פג אחרי שבוע — נגזר מקצב של שני דוחות ביום',mkx.ttl.handled===7,`${mkx.ttl.handled} ימים`);
 ok('שאר הסוגים פגים אחרי חצי שנה',
    mkx.ttl.campaign===183&&mkx.ttl.ignore===183&&mkx.ttl.ondemand===183,JSON.stringify(mkx.ttl));
 ok('«טופל» בן יומיים עדיין מסתיר',mkx.handledFresh.q==='marked'&&mkx.handledFresh.cat==='טופל',
    `${mkx.handledFresh.q} · ${mkx.handledFresh.cat}`);
 ok('«טופל» בן 10 ימים מחזיר את הפריט לסיווג האמיתי',
    mkx.handledStale.q!=='marked'&&mkx.handledStale.cat===mkx.base.cat,
    `${mkx.handledStale.q} · ${mkx.handledStale.cat} (בסיס: ${mkx.base.cat})`);
 ok('וההסבר אומר שהסימון פג',/פג/.test(mkx.handledStale.note),mkx.handledStale.note);
 ok('«קמפיין» בן 100 יום עדיין מסתיר',mkx.campFresh.q==='marked',mkx.campFresh.q);
 ok('«קמפיין» בן 200 יום פג',mkx.campStale.q!=='marked'&&mkx.campStale.stale,
    `${mkx.campStale.q} · ${mkx.campStale.note}`);
 ok('«מוחרג ידנית» בן 200 יום פג',mkx.ignoreStale.q!=='marked'&&mkx.ignoreStale.stale,
    `${mkx.ignoreStale.q} · ${mkx.ignoreStale.note}`);
 ok('«לפי דרישה» שפג ממשיך למסלול שלו ולא לסיווג רגיל',
    mkx.odStale.odStale===true&&mkx.odStale.cat==='עדכון פרמטרים',JSON.stringify(mkx.odStale));
 ok('הסרת הסימון מחזירה את המצב המקורי',
    mkx.after.cat===mkx.base.cat&&!mkx.after.stale,`${mkx.after.cat} (בסיס ${mkx.base.cat})`);
 // המסך הראשי לא נשאר ריק כשאין חוסרים
 const scr=await p.evaluate(()=>({track,shortN:decisionList('short').length,
   rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
   empty:(document.querySelector('#tbl tbody .empty')||{}).textContent||''}));
 ok('כשאין חוסרים — המסך עובר למסלול שיש בו עבודה',
    scr.shortN===0?(scr.track!=='short'&&scr.rows>0):true,
    `מסלול=${scr.track} · חוסרים=${scr.shortN} · שורות=${scr.rows}`);
 if(errs.length)ok('אין שגיאות JS',false,errs.join(' | '));else ok('אין שגיאות JS',true);
 await b.close();console.log(out.join('\n'));
 process.exit(out.some(l=>l.startsWith('FAIL'))?1:0)})();
