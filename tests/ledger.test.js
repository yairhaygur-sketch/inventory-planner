/* הפנקס — הבטחה מול ביצוע לאורך זמן.
   אחסון שלא הוכח על פני יותר מיום אחד הוא הבטחה, לא יכולת. הבדיקה
   מדמה שני ימים: ביום הראשון יש רכש פתוח והבטחת אספקה, וביום השני
   אחד המשלוחים נחת והשני לא. */
const XLSX=require('xlsx'),{chromium}=require('playwright'),fs=require('fs'),path=require('path');
const SD=__dirname, MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const hdr=['מק"ט מוביל','תיאור חומר','תיאור חומר2','שם ספק','סטטוס חומר','תיאור','סוג MRP','ABC','רמת שרות',
 'מלאי בטחון','נק.הז.מחדש','אספ.מתוכנ.','זמ.עב.קבלת','מל.בט.מינ.','מחיר FOB','מטבע FOB','סה"כ מלאי','מלאי פנוי',
 'מלאי מרלוג','הז. רכש','בהעברה','אספקות פת.','כמות בהז.פ','סוג חומר','תיא.קבוצ.חומרים','קב.חו.חיצו','טקסט ארוך',
 'תב.אח.הש.','ת.היר.1 מח','ת.היר.2 מח','היררכייה1','היררכייה1','היררכייה2','היררכייה2','היררכייה3','היררכייה3',
 'צר.השנה','צר.שנה-1','צר.שנה-2','צר.החודש',...MON,"תאר' מכירה","תאר' כניסה"];
const D=n=>{const d=new Date(Date.now()-n*864e5);
 return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`};
const mk=(pn,o)=>{const m=Array(11).fill(6),y=66;
 return [pn,'פריט '+pn,'Part '+pn,o.sup||'ספק א','01','פעיל','ND','A',95,
  0,0,30,2,0,100,'USD',o.free,o.free,0,o.po,0,0,0,'Z004','מנוע','ZT','מתכנן','0',
  'שיווק','מערכת','100','ZT','200','דגם','300','מערכת',y,y,y,0,...m,D(10),D(60)]};
const book=rows=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['ZMRP'],[],hdr,...rows]),'ZMRP');return wb};

/* יום 1: לשניהם רכש פתוח ואפס על המדף */
XLSX.writeFile(book([
 mk('LEDG-KEPT',  {po:20,free:0,sup:'ספק שעומד בזמן'}),
 mk('LEDG-MISSED',{po:15,free:0,sup:'ספק שלא עומד'})]), SD+'/ledger-d1.xlsx');
/* יום 2: KEPT נחת — הרכש ירד ל-0 והמדף עלה ל-20. MISSED לא זז. */
XLSX.writeFile(book([
 mk('LEDG-KEPT',  {po:0, free:20,sup:'ספק שעומד בזמן'}),
 mk('LEDG-MISSED',{po:15,free:0, sup:'ספק שלא עומד'})]), SD+'/ledger-d2.xlsx');

const RLM=s=>'‎'+s+'‏';
const today=new Date(); const ymd=d=>d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
XLSX.writeFile((()=>{const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([
  ['אספקה','פריט','חומר','תיאור','כמות באספקה',"א'",'תארי.אספקה'],
  ['4200000001','000010',RLM('LEDG-KEPT'),'KEPT',20,'EA',new Date(today.getFullYear(),today.getMonth(),today.getDate())],
  ['4200000002','000010',RLM('LEDG-MISSED'),'MISSED',15,'EA',new Date(today.getFullYear(),today.getMonth(),today.getDate())]
 ]),'גיליון1');return wb})(), SD+'/ledger-eta.xlsx');

(async()=>{
 const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.evaluate(()=>{localStorage.removeItem('planner_ledger_v1');localStorage.removeItem('planner_eta_v1')});
 await p.reload();await p.waitForTimeout(300);

 // ── יום 1 ──
 await p.setInputFiles('#fe',SD+'/ledger-eta.xlsx');await p.waitForTimeout(600);
 await p.setInputFiles('#f',SD+'/ledger-d1.xlsx');await p.waitForTimeout(1800);
 const d1=await p.evaluate(()=>({days:ledgerDays().length,n:LEDGER.days[ledgerDays()[0]].n,
   promises:ledgerPromises().length}));

 /* הזזת היום הראשון ליום אחורה — כך שתאריך ההבטחה (היום) נופל בין
    «לפני» ל«אחרי». אי אפשר לחכות יממה בבדיקה. */
 await p.evaluate(()=>{const k=ledgerDays()[0];
   const d=new Date(k);d.setDate(d.getDate()-1);
   const p2=n=>String(n).padStart(2,'0');
   const prev=d.getFullYear()+'-'+p2(d.getMonth()+1)+'-'+p2(d.getDate());
   LEDGER.days[prev]=LEDGER.days[k];delete LEDGER.days[k];ledgerSave()});

 // ── יום 2: KEPT נחת ──
 await p.setInputFiles('#f',SD+'/ledger-d2.xlsx');await p.waitForTimeout(1800);
 const d2=await p.evaluate(()=>{
  const obs=ledgerPromises();
  const g=ledgerBySupplier(ALL);
  return {days:ledgerDays().length,obs,sup:g,
    kept:obs.find(o=>o.pn==='LEDG-KEPT'),missed:obs.find(o=>o.pn==='LEDG-MISSED'),
    kb:+(JSON.stringify(LEDGER).length/1024).toFixed(1),
    /* הפנקס עבר משורה משלו לשורת המקרא של הפסים — הגובה יקר.
       מה שנבדק לא השתנה: שהמסך מציג אמינות ולא «התחלנו היום». */
    line:(document.querySelector('.bledg')||{}).textContent||'',
    chips:document.querySelectorAll('.bledg .lsup').length}});

 // ── שרידות: הפנקס חייב לשרוד רענון וטעינה מחדש ──
 await p.reload();await p.waitForTimeout(400);
 const surv=await p.evaluate(()=>({days:ledgerDays().length,
   inState:STATE_KEYS.includes('planner_ledger_v1')}));

 const out=[],ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));
 ok('יום 1 נרשם, ורק הפריטים שיש להם רכש או הבטחה',d1.days===1&&d1.n===2,
    `ימים ${d1.days} · פריטים ${d1.n}`);
 ok('ביום אחד אין מה למדוד — ואומרים את זה',d1.promises===0,`תצפיות ${d1.promises}`);
 ok('אחרי יומיים יש שתי תצפיות',d2.days===2&&d2.obs.length===2,
    `ימים ${d2.days} · תצפיות ${d2.obs.length}`);
 /* זה הלב: ההבטחה הייתה 20, והרכש הפתוח ירד ב-20 והמדף עלה ב-20. */
 ok('משלוח שנחת נמדד — הרכש ירד והמדף עלה',
    !!d2.kept&&d2.kept.q===20&&d2.kept.dPo===-20&&d2.kept.dFree===20,
    d2.kept?`הובטח ${d2.kept.q} · רכש ${d2.kept.dPo} · מדף +${d2.kept.dFree}`:'לא נמצא');
 ok('משלוח שלא נחת נמדד גם הוא — שום דבר לא זז',
    !!d2.missed&&d2.missed.dPo===0&&d2.missed.dFree===0,
    d2.missed?`רכש ${d2.missed.dPo} · מדף ${d2.missed.dFree}`:'לא נמצא');
 const kept=d2.sup.find(([s])=>/עומד בזמן/.test(s));
 const late=d2.sup.find(([s])=>/לא עומד/.test(s));
 ok('הספק שעמד בזמן מסומן 1/1',!!kept&&kept[1].kept===1&&kept[1].n===1,
    kept?`${kept[1].kept}/${kept[1].n}`:'לא נמצא');
 ok('הספק שלא עמד מסומן 0/1',!!late&&late[1].kept===0&&late[1].n===1,
    late?`${late[1].kept}/${late[1].n}`:'לא נמצא');
 ok('המסך מציג אמינות ולא הודעת "התחלנו היום"',
    /אמינות/.test(d2.line)&&!/התחיל היום/.test(d2.line),d2.line.trim().slice(0,70));
 /* «1/1» ו«0/9» באותו אפור אינם אומרים אותו דבר — הצבע נושא משמעות */
 ok('לכל ספק תג צבעוני לפי העמידה בזמן',d2.chips===2,d2.chips+' תגים');
 ok('הפנקס שורד רענון',surv.days===2,`ימים ${surv.days}`);
 ok('הפנקס נכלל בגיבוי מצב העבודה',surv.inState);
 ok('הגודל סביר',d2.kb<50,d2.kb+'KB לשני ימים');
 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));

 console.log(out.join('\n'));
 const f=out.filter(x=>x.startsWith('FAIL')).length;
 console.log(`\n${out.length-f}/${out.length} עברו`);
 await b.close();
 process.exit(f?1:0);
})();
