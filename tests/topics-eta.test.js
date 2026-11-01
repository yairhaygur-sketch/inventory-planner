/* ============ התצוגה המצומצמת: דוח ETA, «מלאי 0 / צפוי להיגמר» ו«הזמנות לקוח» ============
   החלטות המעתד (1.11): דוח ETA נטען גם בתצוגה המצומצמת — אותו קובץ, אותה קריאה
   ואותו עיבוד של הכלי המלא (#fe → etaParse → etaApply). «צפוי להיגמר החודש»
   מקזז אספקה מתוארכת שתגיע עד סוף החודש, ו«מלאי 0 עכשיו» הוא רשימת «חוסר» של
   הכלי המלא (inToday) — כולל זה שפריט שהאספקה המתוארכת מכסה יורד ממנה.
   הדוחות נבנים כאן בתאריכים יחסיים להיום, כמו ב-eta.test.js: חודש הבסיס הוא
   החודש המלא האחרון, ולכן «סוף החודש» הוא סוף החודש הנוכחי בכל יום שהבדיקה רצה. */
const XLSX=require('xlsx'),{chromium}=require('playwright'),fs=require('fs'),path=require('path'),os=require('os');
const SD=__dirname,TMP=fs.mkdtempSync(path.join(os.tmpdir(),'topics-eta-'));
const MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const hdr=['מק"ט מוביל','תיאור חומר','תיאור חומר2','שם ספק','סטטוס חומר','תיאור','סוג MRP','ABC','רמת שרות',
 'מלאי בטחון','נק.הז.מחדש','אספ.מתוכנ.','זמ.עב.קבלת','מל.בט.מינ.','מחיר FOB','מטבע FOB','סה"כ מלאי','מלאי פנוי',
 'מלאי מרלוג','הז. רכש','בהעברה','אספקות פת.','כמות בהז.פ','סוג חומר','תיא.קבוצ.חומרים','קב.חו.חיצו','טקסט ארוך',
 'תב.אח.הש.','ת.היר.1 מח','ת.היר.2 מח','היררכייה1','היררכייה1','היררכייה2','היררכייה2','היררכייה3','היררכייה3',
 'צר.השנה','צר.שנה-1','צר.שנה-2','צר.החודש',...MON,"תאר' מכירה","תאר' כניסה"];
const D=n=>{const d=new Date(Date.now()-n*864e5);
 return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`};
const mk=(pn,o)=>{const m=o.months||Array(11).fill(10),y=m.reduce((a,b)=>a+b,0);
 return [pn,'פריט מבחן '+pn,'Test '+pn,'ספק','01','פעיל','ND','A',95,
  0,0,30,2,0,100,'USD',o.free,o.free,0,o.po||0,0,0,o.cust||0,'Z004','מנוע','ZT','מתכנן','0',
  'שיווק','מערכת','100','ZT','200','דגם','300','מערכת',y,y,y,0,...m,D(10),D(60)]};
/* קצב 10 בחודש בכל הפריטים */
const rows=[
 mk('M-COVER',{free:2,po:20}),   // צפוי להיגמר; ETA של 8 בתוך החודש מכסה ⇒ יורד
 mk('M-LATE',{free:2,po:20}),    // אותה כמות, אחרי סוף החודש ⇒ נשאר
 mk('M-PART',{free:2,po:20}),    // ETA של 3 בתוך החודש ⇒ נשאר, עם 5 לא יסופקו
 mk('Z-COVER',{free:0,po:20}),   // מדף ריק, בלי לקוח; ETA מכסה את החודש ⇒ יורד (כמו בכלי המלא)
 mk('Z-T1',{free:0,po:20,cust:4}),// מדף ריק ולקוח ממתין ⇒ נשאר ב«מלאי 0» גם עם ETA
 mk('C-NONE',{free:1,po:0,cust:5}),   // הזמנה 5, מלאי 1, אין רכש ⇒ «אין רכש כלל»
 mk('C-PART',{free:2,po:3,cust:10}),  // הזמנה 10, מלאי 2 + רכש 3 ⇒ «כיסוי חלקי»
 mk('C-COVER',{free:2,po:10,cust:5})];// מלאי 2 + רכש 10 ≥ 5 ⇒ מכוסה, אינו בנושא
const ZF=path.join(TMP,'zmrp.xlsx'),EF=path.join(TMP,'eta.xlsx');
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['ZMRP'],[],hdr,...rows]),'ZMRP');return wb})(),ZF);
const nowD=new Date();
const inMonth=new Date(nowD.getFullYear(),nowD.getMonth()+1,0);     // היום האחרון של החודש
const afterMonth=new Date(nowD.getFullYear(),nowD.getMonth()+2,15); // אמצע החודש הבא
const ehdr=['אספקה','פריט','חומר','תיאור','כמות באספקה',"א'",'תארי.אספקה'];
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([ehdr,
 ['4190000001','000010','M-COVER','x',8,'EA',inMonth],
 ['4190000002','000010','M-LATE','x',8,'EA',afterMonth],
 ['4190000003','000010','M-PART','x',3,'EA',inMonth],
 ['4190000004','000010','Z-COVER','x',10,'EA',inMonth],
 ['4190000005','000010','Z-T1','x',20,'EA',inMonth]]),'גיליון1');return wb})(),EF);

let pass=0,fail=0;const ok=(n,c,d)=>{c?pass++:fail++;console.log(`${c?'PASS':'FAIL'} · ${n}${d?'  ['+d+']':''}`)};
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML='file://'+path.join(SD,'..','index.html');
/* מה כל אחד משני הכלים אומר — רץ בעמוד */
const STATE=()=>{const by=Object.fromEntries(ALL.map(r=>[r.pn,r]));
 const sh=decisionList('short').map(r=>r.pn).sort();
 const T=typeof TOPICS!=='undefined'&&document.documentElement.classList.contains('lite')
  ?Object.fromEntries(['cust','short'].map(t=>[t,Object.fromEntries(TOPICS[t].cats.map(c=>[c[0],TOPICS[t].rows().filter(c[3]).map(r=>r.pn).sort()]))])):null;
 return {eta:!!ETA,short:sh,T,
  uns:Object.fromEntries(Object.keys(by).map(pn=>[pn,by[pn].unsQty])),
  etaMonth:Object.fromEntries(Object.keys(by).map(pn=>[pn,by[pn].etaMonth||0]))}};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const open=async q=>{const ctx=await b.newContext({viewport:{width:1440,height:900}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));await p.goto(HTML+q);await p.waitForTimeout(600);return {p,ctx,errs}};
 const loadZ=async p=>{await p.setInputFiles('#f',ZF);
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB,null,{timeout:120000});await p.waitForTimeout(300)};

 /* ---- 1. התצוגה המצומצמת: לפני ETA ---- */
 const L=await open('?topics');
 ok('?topics: כפתור «דוח ETA» בכותרת, גם לפני שנטען דוח',await L.p.evaluate(()=>!!document.querySelector('#lite .tph label[for="fe"]')));
 await loadZ(L.p);await L.p.waitForSelector('#lite table');
 const s0=await L.p.evaluate(STATE);
 /* גם שלושת פריטי C-* צורכים 10 בחודש עם מדף של 1–2, ולכן גם הם צפויים להיגמר — פריט יכול להיות בשני נושאים */
 ok('בלי ETA: «צפוי להיגמר החודש» = כל פריט עם מלאי שקצב 10 גדול ממנו',JSON.stringify(s0.T.short.month)===JSON.stringify(['C-COVER','C-NONE','C-PART','M-COVER','M-LATE','M-PART']),s0.T.short.month.join(','));
 ok('בלי ETA: «מלאי 0 עכשיו» כולל Z-COVER ו-Z-T1',['Z-COVER','Z-T1'].every(pn=>s0.T.short.zero.includes(pn)),s0.T.short.zero.join(','));
 ok('«אין רכש כלל» = C-NONE (מלאי 1, אין רכש, הזמנה 5)',s0.T.cust.none.includes('C-NONE')&&!s0.T.cust.part.includes('C-NONE'),JSON.stringify(s0.T.cust));
 ok('«כיסוי חלקי» = C-PART (מלאי 2 + רכש 3 < הזמנה 10)',s0.T.cust.part.includes('C-PART')&&!s0.T.cust.none.includes('C-PART'));
 ok('הזמנה שמלאי + רכש מכסים (C-COVER, Z-T1) אינה בנושא 1',!s0.T.cust.none.concat(s0.T.cust.part).some(pn=>pn==='C-COVER'||pn==='Z-T1'));
 ok('בלי ETA: מצב הדוח בכותרת — «אין דוח ETA»',await L.p.evaluate(()=>/אין דוח ETA/.test(document.querySelector('#lite .src.eta').textContent)));

 /* ---- 2. העלאת ETA דרך הכפתור — בורר קבצים אמיתי ---- */
 await L.p.click('[data-tp-tab="short"]');
 let ch=null;try{[ch]=await Promise.all([L.p.waitForEvent('filechooser',{timeout:5000}),L.p.click('#lite .tph label[for="fe"]')])}catch(_){}
 ok('לחיצה על «דוח ETA» בתצוגה המצומצמת פותחת בורר קבצים',!!ch);
 if(ch){await ch.setFiles(EF);await L.p.waitForFunction(()=>!!ETA&&ALL.some(r=>r.etaHas),null,{timeout:20000}).catch(()=>{});await L.p.waitForTimeout(500)}
 const s1=await L.p.evaluate(STATE);
 ok('ETA נטען דרך אותו מנגנון של הכלי המלא (ETA, etaApply)',s1.eta&&s1.etaMonth['M-COVER']===8&&s1.etaMonth['M-PART']===3&&s1.etaMonth['M-LATE']===0,JSON.stringify(s1.etaMonth));
 ok('M-COVER: אספקה של 8 עד סוף החודש מכסה את החוסר ⇒ יורד מ«צפוי להיגמר»',!s1.T.short.month.includes('M-COVER')&&s1.uns['M-COVER']===0,`uns=${s1.uns['M-COVER']}`);
 ok('M-LATE: אותה כמות אחרי סוף החודש ⇒ נשאר, 8 לא יסופקו',s1.T.short.month.includes('M-LATE')&&s1.uns['M-LATE']===8,`uns=${s1.uns['M-LATE']}`);
 ok('M-PART: כיסוי חלקי בתוך החודש ⇒ נשאר, 5 לא יסופקו',s1.T.short.month.includes('M-PART')&&s1.uns['M-PART']===5,`uns=${s1.uns['M-PART']}`);
 ok('Z-COVER: מדף ריק בלי לקוח, ETA מכסה את החודש ⇒ יורד מ«מלאי 0» — כמו ברשימת «חוסר» של הכלי המלא',!s1.T.short.zero.includes('Z-COVER'),s1.T.short.zero.join(','));
 ok('Z-T1: מדף ריק ולקוח ממתין ⇒ נשאר ב«מלאי 0» גם עם ETA',s1.T.short.zero.includes('Z-T1'));
 ok('נושא 1 אינו מושפע מ-ETA (מלאי + רכש פתוח)',JSON.stringify(s1.T.cust)===JSON.stringify(s0.T.cust));
 const ui=await L.p.evaluate(()=>({chip:document.querySelector('#lite .src.eta').textContent,
  month:+String((document.querySelector('#lite .tile[data-tp-cat="month"] .v')||{}).textContent||'').replace(/[^\d]/g,''),
  col:[...document.querySelectorAll('#lite th')].some(th=>/אספקה עד סוף החודש/.test(th.textContent))}));
 ok('אחרי הטעינה: הכותרת מציגה את הדוח, האריח מתעדכן, ויש עמודת אספקה עד סוף החודש',/ETA · /.test(ui.chip)&&ui.month===s1.T.short.month.length&&ui.col,JSON.stringify(ui));

 /* ---- 3. אותם קבצים בכלי המלא — אותה תוצאה בדיוק ---- */
 const F=await open('?nobrief=1');await loadZ(F.p);
 await F.p.setInputFiles('#fe',EF);await F.p.waitForFunction(()=>!!ETA&&ALL.some(r=>r.etaHas),null,{timeout:20000}).catch(()=>{});await F.p.waitForTimeout(500);
 const f1=await F.p.evaluate(STATE);
 ok('הכלי המלא: אותה רשימת «חוסר» ואותם «לא יסופק החודש» לכל פריט',JSON.stringify(f1.short)===JSON.stringify(s1.short)&&JSON.stringify(f1.uns)===JSON.stringify(s1.uns),
  `${f1.short.join(',')} · ${s1.short.join(',')}`);
 ok('נושא 3 בתצוגה המצומצמת = רשימת «חוסר» של הכלי המלא',JSON.stringify([...s1.T.short.zero,...s1.T.short.month].sort())===JSON.stringify(f1.short));

 /* ---- 4. ETA לפני ZMRP — נשמר ומצטרף בטעינה ---- */
 const E=await open('?topics');
 await E.p.setInputFiles('#fe',EF);await E.p.waitForTimeout(800);
 const pre=await E.p.evaluate(()=>({eta:!!ETA,chip:(document.querySelector('#lite .src.eta')||{}).textContent||''}));
 ok('ETA לפני ZMRP: נשמר, והכותרת מציגה אותו',pre.eta&&/ETA · /.test(pre.chip),JSON.stringify(pre));
 await loadZ(E.p);await E.p.waitForSelector('#lite table');
 const e1=await E.p.evaluate(STATE);
 ok('ואחרי טעינת ZMRP — אותה תוצאה כמו בטעינה בסדר ההפוך',JSON.stringify(e1.T.short)===JSON.stringify(s1.T.short)&&JSON.stringify(e1.uns)===JSON.stringify(s1.uns));

 ok('אפס שגיאות JS',!L.errs.length&&!F.errs.length&&!E.errs.length,[...L.errs,...F.errs,...E.errs].slice(0,2).join(' | '));
 await b.close();fs.rmSync(TMP,{recursive:true,force:true});
 console.log(`\n${pass} עברו · ${fail} נכשלו`);process.exit(fail?1:0)})();
