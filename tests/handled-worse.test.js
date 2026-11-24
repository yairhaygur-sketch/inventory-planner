/* ============ HANDLED_WORSE — «טופל» חוזר מוקדם אם המצב החמיר (אושר 23.11) ============
   בסימון «טופל» נשמרים המלאי הפנוי והחוסר מול הזמנות הלקוח (snap). דוח מאוחר יותר שבו המלאי הפנוי
   נגמר (היה מעל 0, עכשיו 0) או שהחוסר מול הזמנות הלקוח גדל — הפריט חוזר מיד, בשני הכלים, עם «החמיר מאז הסימון». מצב זהה או טוב
   יותר — נשאר מסומן. סימון ישן בלי snap — כמו קודם. «דוח מאוחר יותר» מדומה כאן בשינוי השדות
   של הפריט ו-apply() — אותו מסלול שעובר דוח חדש (classify על הנתונים החדשים מול MARKS). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');
 await p.evaluate(()=>{try{localStorage.clear()}catch(_){}});
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');

 /* 1 · הסימון שומר תמונת מצב */
 const S=await p.evaluate(()=>{TP.tab='cust';TP.cat={};tpRender();const r=TP.list[0],r2=TP.list[1],r3=TP.list[2];
  const want={free:Math.max(0,r.free),gap:Math.max(0,r.cust-Math.max(0,r.free)-Math.max(0,r.po))};
  tpDoneNext(r,'הוזמן בדחוף');tpDoneNext(r2,'');tpDoneNext(r3,'');
  const m=MARKS[markKey(r)];return {pn:r.pn,pn2:r2.pn,pn3:r3.pn,want,snap:m&&m.snap,gap:want.gap,inList:TP.list.includes(r)}});
 ok('«טופל» שומר את המלאי הפנוי ואת החוסר מול הזמנות הלקוח',JSON.stringify(S.snap)===JSON.stringify(S.want),JSON.stringify(S.snap));
 ok('בנושא «הזמנות לקוח ללא כיסוי» החוסר בתמונה גדול מאפס, והפריט יצא מהרשימה',S.gap>0&&!S.inList,`gap ${S.gap}`);

 /* 2 · אותו דוח שוב — נשאר מסומן */
 const R2=await p.evaluate(pn=>{apply();const r=QF.all.find(x=>x.pn===pn);return {mark:r.mark&&r.mark.t,worse:!!r.markWorse}},S.pn);
 ok('אותו דוח שוב (apply): הפריט נשאר «טופל»',R2.mark==='handled'&&!R2.worse,JSON.stringify(R2));

 /* 3 · מצב טוב יותר — נשאר מסומן */
 const R3=await p.evaluate(pn=>{const r=QF.all.find(x=>x.pn===pn);r.po+=1;apply();const v={mark:r.mark&&r.mark.t,worse:!!r.markWorse};r.po-=1;apply();return v},S.pn);
 ok('דוח עם רכש פתוח גדול יותר (החוסר קטן): נשאר «טופל»',R3.mark==='handled'&&!R3.worse,JSON.stringify(R3));

 /* 4 · הזמנת לקוח גדלה — חוזר, עם «החמיר מאז הסימון» */
 const R4=await p.evaluate(pn=>{const r=QF.all.find(x=>x.pn===pn);const g0=MARKS[markKey(r)].snap.gap;r.cust+=3;apply();tpRender();
  const why=(r.why||[]).map(w=>w[1]).join(' | ');
  return {mark:r.mark,worse:!!r.markWorse,why,g0,g1:Math.max(0,r.cust-Math.max(0,r.free)-Math.max(0,r.po)),inList:TP.list.includes(r),
   stillStored:!!MARKS[markKey(r)],inQ:[...QF.waiting,...QF.immediate].includes(r),inMarked:QF.marked.includes(r),
   prof:profMarkWhy(r)}},S.pn);
 ok('הזמנת לקוח גדלה: הסימון אינו פעיל, והפריט חזר לנושא',R4.mark==null&&R4.worse&&R4.inList,JSON.stringify({w:R4.worse,l:R4.inList}));
 ok('…ובכלי המלא: חזר לתור ויצא מ«מסומנים»',R4.inQ&&!R4.inMarked);
 ok('…ב«למה»: «החמיר מאז הסימון», החוסר לפני ואחרי, וההערה',R4.why.includes('החמיר מאז הסימון')&&R4.why.includes(`החוסר מול הזמנות לקוח גדל מ-${R4.g0} ל-${R4.g1}`)&&R4.why.includes('הוזמן בדחוף'),R4.why.slice(0,200));
 ok('…הסימון עצמו נשמר (תאריך והערה), ואינו מחריג המלצת פרופיל',R4.stillStored&&R4.prof==='');

 /* 5 · כרטיס ושורה בתצוגת 4 הנושאים */
 const R5=await p.evaluate(pn=>{const r=QF.all.find(x=>x.pn===pn);tpdShow(r);const n=document.querySelector('#tpd .tpprev');
  const k=(TP.shown||[]).indexOf(r),tr=document.querySelector(`#lite tr.r[data-tp-row="${k}"]`),ic=tr&&tr.querySelector('.tpnote');
  const v={note:n?n.textContent:'',worse:!!n&&n.classList.contains('worse'),icon:ic?ic.textContent:'',tip:ic?ic.title:''};tpdClose();return v},S.pn);
 ok('בכרטיס: «סומן «טופל» ב-…: ההערה — החמיר מאז הסימון: … — וחזר לרשימה לפני הזמן»',R5.worse&&R5.note.includes('החמיר מאז הסימון')&&R5.note.includes('הוזמן בדחוף')&&R5.note.includes('לפני הזמן'),R5.note);
 ok('בשורה: ⚠, ובבועית ההערה ומה החמיר',R5.icon==='⚠'&&R5.tip.includes('החמיר מאז הסימון')&&R5.tip.includes('הוזמן בדחוף'),R5.icon+' '+R5.tip);

 /* 6 · סימון מחדש לוקח תמונה חדשה */
 const R6=await p.evaluate(pn=>{const r=QF.all.find(x=>x.pn===pn);setMark(r,'handled','שוב');const v={mark:r.mark&&r.mark.t,snapGap:MARKS[markKey(r)].snap.gap,
  gap:Math.max(0,r.cust-Math.max(0,r.free)-Math.max(0,r.po))};return v},S.pn);
 ok('סימון «טופל» מחדש: תמונה חדשה, והפריט מסומן שוב',R6.mark==='handled'&&R6.snapGap===R6.gap,JSON.stringify(R6));

 /* 7 · מלאי פנוי ירד אבל לא נגמר — נשאר; נגמר — חוזר (בחירת המעתד 24.11) */
 const R7=await p.evaluate(pn=>{const r=QF.all.find(x=>x.pn===pn);r.free=5;r.cust=0;setMark(r,'handled','');
  r.free=2;apply();const a={mark:r.mark&&r.mark.t,worse:!!r.markWorse};
  r.free=0;apply();return {a,mark:r.mark,worse:!!r.markWorse,why:(r.why||[]).map(w=>w[1]).join(' | ')}},S.pn2);
 ok('מלאי פנוי ירד מ-5 ל-2 (צריכה רגילה): נשאר «טופל»',R7.a.mark==='handled'&&!R7.a.worse,JSON.stringify(R7.a));
 ok('מלאי פנוי נגמר: חוזר, עם «המלאי הפנוי נגמר (היה 5, עכשיו 0)»',R7.mark==null&&R7.worse&&R7.why.includes('המלאי הפנוי נגמר (היה 5, עכשיו 0)'),R7.why.slice(0,200));

 /* 8 · סימון ישן בלי תמונה — כמו קודם */
 const R8=await p.evaluate(pn=>{const r=QF.all.find(x=>x.pn===pn);const k=markKey(r);const m={...MARKS[k]};delete m.snap;MARKS[k]=m;r.cust+=10;r.free=0;apply();
  return {mark:r.mark&&r.mark.t,worse:!!r.markWorse}},S.pn3);
 ok('סימון ישן בלי תמונה: נשאר «טופל» עד שיפוג, כמו קודם',R8.mark==='handled'&&!R8.worse,JSON.stringify(R8));

 /* 9 · סימון שפג — «פג» קודם ל«החמיר» */
 const R9=await p.evaluate(pn=>{const r=QF.all.find(x=>x.pn===pn);const k=markKey(r);MARKS[k]={...MARKS[k],ts:Date.now()-8*864e5};apply();
  return {stale:!!r.markStale,worse:!!r.markWorse}},S.pn);
 ok('סימון שפג (8 ימים) ושהחמיר: מסומן «פג», כמו קודם',R9.stale&&!R9.worse,JSON.stringify(R9));

 /* 10 · סימון קבוצתי (setMarks) — גם הוא שומר תמונה */
 const R10=await p.evaluate(()=>{const L=QF.all.filter(x=>!x.mark&&!MARKS[markKey(x)]).slice(0,3);setMarks(L,'handled','');
  return L.every(x=>{const m=MARKS[markKey(x)];return m&&m.snap&&m.snap.free===Math.max(0,x.free)})});
 ok('סימון קבוצתי «טופל»: גם הוא שומר תמונת מצב',R10);
 /* סימונים אחרים — בלי תמונה */
 const R11=await p.evaluate(()=>{const x=QF.all.find(y=>!y.mark&&!MARKS[markKey(y)]);setMark(x,'campaign','');return !('snap' in MARKS[markKey(x)])});
 ok('סימון אחר («קמפיין»): בלי תמונה — אינו חוזר מוקדם',R11);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
