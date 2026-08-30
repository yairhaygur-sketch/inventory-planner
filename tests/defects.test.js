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
  o.ss||0,o.rop||0,o.lt||30,2,0,100,o.cur||'USD',o.free,o.free,0,o.po||0,0,0,o.cust||0,'Z004','מנוע','ZT','מתכנן','0',
  'שיווק','מערכת','100','ZT','200','דגם','300','מערכת',
  o.y0!=null?o.y0:y, o.y1!=null?o.y1:y, o.y2!=null?o.y2:y, 0, ...m, D(o.saleAgo||10), D(o.entAgo||60)]};

const rows=[
 // A · סטיית תקן: ביקוש שטוח לגמרי, חודש אזילה בודד
 mk('FLAT-CLEAN',{months:[10,10,10,10,10,10,10,10,10,10,10],free:60}),
 mk('FLAT-STOCKOUT',{months:[10,10,10,10,10,0,10,10,10,10,10],free:60}),
 // B · אזילה של 3 חודשים רצופים באמצע, ביקוש בריא סביב
 mk('RUN-STOCKOUT',{months:[20,20,20,0,0,0,20,20,20,20,20],free:100}),
 // C · ביקוש 2 יח׳/שנה, מעט מלאי, ROP ענק ב-SAP
 mk('LOW-DEMAND',{months:[0,0,1,0,0,0,1,0,0,0,0],y0:2,y1:2,y2:2,free:3,rop:200,ss:80}),
 // D · ביקוש בירידה חדה שבה ה-ROP המוצע דווקא גבוה מהקיים
 mk('DECLINE',{months:[2,2,3,2,3,20,25,22,24,23,25],free:40,rop:1,ss:0,lt:120}),
 // B2 · ספורדי: חודשי האפס הם היעדר ביקוש, לא אזילה
 mk('SPORADIC',{months:[0,0,10,0,0,0,0,10,0,0,10],free:200}),
 mk('EDGE-ONLY',{months:[5,0,0,0,0,0,0,0,0,0,6],free:200}),
 // COVERED · מלאי גדול, בלי הזמנות לקוח — כמו הדוח האמיתי שבו אין חוסרים
 mk('COVERED-1',{months:[8,8,8,8,8,8,8,8,8,8,8],free:900,cust:0,rop:5,ss:2}),
 mk('COVERED-2',{months:[6,6,6,6,6,6,6,6,6,6,6],free:800,cust:0,rop:5,ss:2}),
 // מטבעות — עודף מלאי זהה בשלושה מטבעות שונים
 mk('CUR-USD',{months:[2,2,2,2,2,2,2,2,2,2,2],free:500,cur:'USD'}),
 mk('CUR-EUR',{months:[2,2,2,2,2,2,2,2,2,2,2],free:500,cur:'EUR'}),
 mk('CUR-ILS',{months:[2,2,2,2,2,2,2,2,2,2,2],free:500,cur:'ILS'}),
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
   return {cat:r.cat,sd:+r.A.sd.toFixed(2),drops:r.A.drops.length,sugSS:r.sugSS,rop:r.rop,
     act:(r.act||[])[0]||'',trend:r.A.trend.pct}};
  return {clean:g('FLAT-CLEAN'),so:g('FLAT-STOCKOUT'),run:g('RUN-STOCKOUT'),low:g('LOW-DEMAND'),
          dec:g('DECLINE'),spor:g('SPORADIC'),edge:g('EDGE-ONLY'),
          cur:{usd:g('CUR-USD'),eur:g('CUR-EUR'),ils:g('CUR-ILS')}}});
 const out=[],ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));
 // A · חודש אזילה אינו יוצר מלאי ביטחון יש מאין
 ok('ביקוש שטוח → סטיית תקן 0',o.clean.sd===0&&o.clean.sugSS===0);
 ok('חודש אזילה אינו מנפח את סטיית התקן',o.so.sd===0&&o.so.sugSS===0,
    `sd=${o.so.sd} SS=${o.so.sugSS} (לפני התיקון: 3.02 / 6)`);
 // B · רצף אזילה מזוהה ואינו נראה כירידת ביקוש
 ok('אזילה של 3 חודשים מזוהה',o.run.drops===3,`drops=${o.run.drops} (לפני: 0)`);
 ok('רצף אזילה מסווג כחשד אזילה',o.run.cat==='חשד אזילה',o.run.cat);
 ok('רצף אזילה אינו מוצג כירידה בביקוש',o.run.trend===0,`מגמה ${o.run.trend}% (לפני: -40%)`);
 ok('ההודעה מציינת את אורך הרצף',/3 החודשים/.test(o.run.act),o.run.act);
 // B2 · פריט ספורדי אינו אזילה — חודשי אפס הם היעדר ביקוש, לא חוסר מלאי
 ok('פריט ספורדי (3 מתוך 11) אינו מסווג אזילה',o.spor.drops===0,
    `drops=${o.spor.drops} · ${o.spor.cat}`);
 ok('פריט בשני קצוות בלבד אינו מסווג אזילה',o.edge.drops===0,
    `drops=${o.edge.drops} · ${o.edge.cat}`);
 // C · ביקוש זעום עם פרמטרים גבוהים ב-SAP
 ok('פריט 2 יח׳/שנה עם ROP=200 אינו "תקין"',o.low.cat!=='תקין',o.low.cat);
 ok('ומוצעת לו פעולת איפוס',/איפוס/.test(o.low.act),o.low.act);
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
 ok('פריט דולרי מוצג ב-$ ולא ב-₪',/^\$/.test(cur.usdTxt),`${cur.usdTxt} (היה ${cur.ilsTxt})`);
 ok('סכום מפוצל לפי מטבע ולא מחובר',cur.mix.split(' · ').length>1,cur.mix);
 ok('כל המטבעות בקטלוג מיוצגים',cur.symbols.length>1,cur.symbols);
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
