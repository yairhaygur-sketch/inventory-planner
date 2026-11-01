/* ============ SIG — «האם זה אותו דוח?» ============
   הבאג: `runSig` בנתה חתימה משישה מספרים — חודש, מספר פריטים, וארבעה
   סכומים. נקודת ההזמנה ומלאי הביטחון לא היו בה כלל. מעתד שעשה בדיוק
   את מה שהכלי ביקש, עדכן ROP ו-SS ב-SAP והעלה את הדוח למחרת, קיבל
   «אותו קובץ שוב»: לא נשמרה תמונה חדשה, `recordParams` החזירה null,
   והיישום שזה עתה בוצע לא נספר.

   הקובץ הזה מפעיל את *מסלול הקליטה האמיתי* — `setInputFiles` על
   הכלי — ולא את פונקציית החתימה לבדה. חתימה שמחזירה ערכים שונים
   ומנגנון שלא שומר תמונה הם שני באגים שונים, ורק השני הוא זה שנראה
   בשטח. הקובץ הזה בונה את הדוחות שלו בעצמו, כולל את הדוח שבו
   הפרמטרים עודכנו *בדיוק להמלצה שהכלי נתן בהרצה הקודמת* — שנקראת
   מהמנוע בזמן ריצה ולא מוקלדת כאן. */
const {chromium}=require('playwright'),fs=require('fs'),path=require('path'),XLSX=require('xlsx');
const SD=__dirname, HTML='file://'+path.join(SD,'..','index.html');
const out=[];const ok=(n,c,d)=>out.push(`${c?'PASS':'FAIL'} · ${n}${d?'  ['+d+']':''}`);

const MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const hdr=['מק"ט מוביל','תיאור חומר','תיאור חומר2','שם ספק','סטטוס חומר','תיאור',
 'סוג MRP','ABC','רמת שרות','מלאי בטחון','נק.הז.מחדש','אספ.מתוכנ.','זמ.עב.קבלת','מל.בט.מינ.',
 'מחיר FOB','מטבע FOB','סה"כ מלאי','מלאי פנוי','מלאי מרלוג','הז. רכש','בהעברה','אספקות פת.','כמות בהז.פ','סוג חומר',
 'תיא.קבוצ.חומרים','קב.חו.חיצו','טקסט ארוך','תב.אח.הש.','ת.היר.1 מח','ת.היר.2 מח',
 'היררכייה1','היררכייה1','היררכייה2','היררכייה2','היררכייה3','היררכייה3',
 'צר.השנה','צר.שנה-1','צר.שנה-2','צר.החודש',...MON,"תאר' מכירה","תאר' כניסה"];
const C={ss:hdr.indexOf('מלאי בטחון'),rop:hdr.indexOf('נק.הז.מחדש'),
 free:hdr.indexOf('מלאי פנוי'),stock:hdr.indexOf('סה"כ מלאי'),
 m1:hdr.indexOf('צר.חודש-1'),m2:hdr.indexOf('צר.חודש-2')};
/* BASE_MONTH: תאריכים קבועים מתאריך המצב של הדוחות הסינתטיים, לא מיום ההרצה */
const {FIXTURE_NOW,FIXTURE_HEAD}=require('./lib/fixture-date.js');
const DAY=864e5,now=FIXTURE_NOW;
const dstr=d=>{const x=new Date(d);
 return `${String(x.getDate()).padStart(2,'0')}.${String(x.getMonth()+1).padStart(2,'0')}.${x.getFullYear()}`};
/* דוח דטרמיניסטי לחלוטין — בלי הגרלה. פריט שמתנהג רגיל ושיש לו
   היסטוריית צריכה יציבה מקבל המלצת פרמטרים מהמנוע, וזה כל מה שדרוש. */
/* ============ למה הפרמטרים מתחילים ב-0 ============
   הגרסה הראשונה של הפיקסטורה נתנה ROP=10 מול המלצה 12. המנוע סיווג
   אותה — נכון — כ«הפרש קטן מכדי לתקן»: 2 יחידות על קצב 10 הן 0.2 חודשי
   ביקוש, מתחת ל-PB_MONTHS. פריט כזה אינו נכנס לתמונה, ולכן לא היה מה
   ליישם ולא היה מה להשוות. הפער כאן גדול מחודש ביקוש אחד בכוונה.
   הצריכה משתנה בין החודשים כדי שסטיית התקן לא תהיה 0 — אחרת גם
   מלאי הביטחון המומלץ יוצא 0 ושדה שלם אינו נבדק. */
const VARY=[8,12,9,11,10,13,7,12,9,11,10];
function row(pn,o){
 const months=o.months||VARY.slice();
 const y=months.reduce((a,b)=>a+b,0);
 return [pn,`חלק ${pn}`,`Part ${pn}`,'ספק א׳','01','פעיל','ND','B',80,
  o.ss==null?0:o.ss, o.rop==null?0:o.rop, 30,5,0,100,'USD',
  (o.free==null?50:o.free)+5, o.free==null?50:o.free, 0, 0,0,0, o.cust||0, 'Z001',
  'מנוע','ZT','מתכנן א׳','12 חודשים','מנוע','מנוע',
  '100','ZT','200','ZT T5','300','מנוע',
  y,y,y,0,...months, dstr(now-20*DAY), dstr(now-60*DAY)]};
function base(){const r=[];
 for(let i=0;i<24;i++)r.push(row('ITEM-'+String(i).padStart(3,'0'),{}));
 r.push(row('TEST-ROP',{rop:0,ss:0}));
 r.push(row('PAIR-A',{rop:10}));
 r.push(row('PAIR-B',{rop:10}));
 r.push(row('SPLIT-1',{months:VARY.slice()}));
 return r}
function write(f,rows){const w=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet(
  [FIXTURE_HEAD(),[],hdr,...rows]),'ZMRP');
 XLSX.writeFile(w,f);return f}
const at=(rows,pn)=>rows.find(r=>r[0]===pn);
const clone=rows=>rows.map(r=>r.slice());
const F=n=>SD+'/sig-'+n+'.xlsx';

(async()=>{
 const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];
 p.on('pageerror',e=>errs.push(e.message.split('\n')[0].slice(0,140)));
 await p.goto(HTML+'?nobrief=1');await p.waitForTimeout(1200);

 const A=base();
 /* syncHistChip מרנדר רק במסך «היום». בדיקה שקוראת את הרצועה ממסך
   אחר מקבלת מחרוזת ריקה ועוברת על «אין כאן את הטקסט הזה» בלי לבדוק
   כלום — בדיוק סוג הבדיקה הריקה שאסור להשאיר. */
 const chip=async()=>{await p.evaluate(()=>setMode('today'));await p.waitForTimeout(500);
  return p.evaluate(()=>{const c=document.getElementById('histChip');
    return {hidden:!c||c.hidden,txt:((c&&c.textContent)||'').trim()}})};
 const load=async f=>{await p.setInputFiles('#f',f);await p.waitForTimeout(2600);
  return p.evaluate(()=>({hist:HIST.runs.length,par:PARAMS.runs.length,
    repeat:!!(HRUN&&HRUN.repeat),migrated:!!(HRUN&&HRUN.migrated),
    sig:HIST.runs[HIST.runs.length-1].sig,
    applied:APPLIED?{total:APPLIED.total,done:APPLIED.done,other:APPLIED.other,
      none:APPLIED.none,migrated:!!APPLIED.migrated}:null,
    test:(ALL.find(x=>x.pn==='TEST-ROP')||{}) &&
      (r=>({rop:r.rop,ss:r.ss,sugROP:r.sugROP,sugSS:r.sugSS,
        apStatus:r.apStatus||null,paramFix:!!r.paramFix}))(ALL.find(x=>x.pn==='TEST-ROP'))}))};

 /* ── 1. דוח הבסיס ─────────────────────────────────────────────── */
 const r1=await load(write(F('a'),A));
 ok('דוח ראשון נקלט ונשמרה תמונה',r1.hist===1&&r1.par===1,
   `${r1.hist} הרצות · ${r1.par} תמונות · ${r1.sig.slice(0,10)}…`);
 ok('ולפריט המבחן יש המלצת פרמטרים מהמנוע',
   r1.test.paramFix&&(r1.test.rop!==r1.test.sugROP||r1.test.ss!==r1.test.sugSS),
   `ROP ${r1.test.rop}→${r1.test.sugROP} · SS ${r1.test.ss}→${r1.test.sugSS}`);
 const SUG={rop:r1.test.sugROP,ss:r1.test.sugSS};

 /* ── 2. אותו קובץ שוב, בסדר שורות הפוך ────────────────────────── */
 const shuffled=clone(A).reverse();
 const r2=await load(write(F('a-shuf'),shuffled));
 ok('אותו דוח בסדר שורות שונה — טעינה חוזרת',
   r2.repeat&&r2.hist===1&&r2.par===1,
   `repeat=${r2.repeat} · ${r2.hist} הרצות · ${r2.par} תמונות`);

 /* ── 3. שינוי ROP בלבד ────────────────────────────────────────── */
 const onlyRop=clone(A);at(onlyRop,'ITEM-000')[C.rop]=99;
 const r3=await load(write(F('rop'),onlyRop));
 ok('שינוי ROP בלבד — דוח חדש',!r3.repeat&&r3.hist===2&&r3.par===2,
   `repeat=${r3.repeat} · ${r3.hist} הרצות · ${r3.par} תמונות`);

 /* ── 4. שינוי SS בלבד ─────────────────────────────────────────── */
 const onlySs=clone(onlyRop);at(onlySs,'ITEM-001')[C.ss]=77;
 const r4=await load(write(F('ss'),onlySs));
 ok('שינוי SS בלבד — דוח חדש',!r4.repeat&&r4.hist===3,
   `repeat=${r4.repeat} · ${r4.hist} הרצות`);

 /* ── 5. שני מק״טים שמתקזזים בסכום ─────────────────────────────── */
 const offset=clone(onlySs);
 at(offset,'PAIR-A')[C.rop]=10+7; at(offset,'PAIR-B')[C.rop]=10-7;
 const sumBefore=(+at(onlySs,'PAIR-A')[C.rop])+(+at(onlySs,'PAIR-B')[C.rop]);
 const sumAfter=(+at(offset,'PAIR-A')[C.rop])+(+at(offset,'PAIR-B')[C.rop]);
 const r5=await load(write(F('offset'),offset));
 ok('שינויים שמתקזזים בסכום — עדיין דוח חדש',
   !r5.repeat&&r5.hist===4&&sumBefore===sumAfter,
   `סכום ROP ${sumBefore}→${sumAfter} (זהה) · repeat=${r5.repeat}`);

 /* ── 6. העברת צריכה בין חודשים, אותו סך ───────────────────────── */
 const split=clone(offset);
 const sr=at(split,'SPLIT-1');
 const m1b=+sr[C.m1],m2b=+sr[C.m2];
 sr[C.m1]=m1b+4; sr[C.m2]=m2b-4;
 const r6=await load(write(F('split'),split));
 ok('העברת צריכה בין חודשים בלי שינוי בסך — דוח חדש',
   !r6.repeat&&r6.hist===5&&(m1b+m2b)===((+sr[C.m1])+(+sr[C.m2])),
   `${m1b}/${m2b} → ${sr[C.m1]}/${sr[C.m2]} · סך זהה · repeat=${r6.repeat}`);

 /* ── 7. סימון ידני בלבד — אינו דוח חדש ────────────────────────── */
 const sigBeforeMark=r6.sig;
 await p.evaluate(()=>{setMark('ITEM-005','handled','');setMark('ITEM-006','checking','')});
 await p.waitForTimeout(400);
 const r7=await load(write(F('split-again'),split));
 ok('סימון ידני בלבד אינו הופך את דוח המקור לחדש',
   r7.repeat&&r7.sig===sigBeforeMark&&r7.hist===5,
   `repeat=${r7.repeat} · אותה חתימה=${r7.sig===sigBeforeMark}`);

 /* ── 8. עדכון שני הפרמטרים בדיוק להמלצה ───────────────────────── */
 const applied=clone(split);
 const tr=at(applied,'TEST-ROP'); tr[C.rop]=SUG.rop; tr[C.ss]=SUG.ss;
 const r8=await load(write(F('applied'),applied));
 ok('עדכון ROP ו-SS בדיוק להמלצה — דוח חדש',
   !r8.repeat&&r8.hist===6,`repeat=${r8.repeat} · ${r8.hist} הרצות`);
 ok('והפריט נספר כ«יושם»',
   r8.test.apStatus==='done'&&!!r8.applied&&r8.applied.done>=1,
   `apStatus=${r8.test.apStatus} · ${r8.applied?r8.applied.done:'—'} יושמו `+
   `מתוך ${r8.applied?r8.applied.total:'—'}`);
 ok('גם כאשר paramFix שלו כבר כבוי',
   r8.test.apStatus==='done'&&r8.test.rop===SUG.rop&&r8.test.ss===SUG.ss,
   `paramFix=${r8.test.paramFix} · ROP ${r8.test.rop} · SS ${r8.test.ss}`);

 /* ── 9. אותו דוח שוב אחרי היישום ──────────────────────────────── */
 const parBefore=r8.par, doneBefore=r8.applied?r8.applied.done:-1;
 const r9=await load(write(F('applied-again'),clone(applied)));
 ok('אותו דוח שוב אחרי עדכון מוצלח — אין תמונה נוספת',
   r9.repeat&&r9.par===parBefore,
   `repeat=${r9.repeat} · תמונות ${parBefore}→${r9.par}`);
 ok('וההשוואה האחרונה בין שני דוחות שונים נשמרת',
   !!r9.applied&&r9.applied.done===doneBefore&&r9.test.apStatus==='done',
   `יושמו ${doneBefore}→${r9.applied?r9.applied.done:'—'} · apStatus=${r9.test.apStatus}`);
 const ttl=await chip();
 ok('והכותרת מתארת את אותו בסיס השוואה, לא «בסיס חדש»',
   !ttl.hidden&&ttl.txt.length>0&&!/בסיס חדש/.test(ttl.txt),
   `מוצגת=${!ttl.hidden} · «${ttl.txt.slice(0,46)}»`);

 /* ── 10. שינוי חלקי לכיוון ההמלצה ─────────────────────────────── */
 const partial=clone(applied);
 const pr=at(partial,'ITEM-010');
 const curRop=+pr[C.rop];
 pr[C.rop]=curRop+1;                 /* זז, אך לא אל המספר שהוצע */
 const r10=await load(write(F('partial'),partial));
 const st10=await p.evaluate(()=>{const r=ALL.find(x=>x.pn==='ITEM-010');
   return r?{ap:r.apStatus,drift:r.apDrift,rop:r.rop,sug:r.sugROP}:null});
 ok('שינוי חלקי לכיוון ההמלצה — הסיווג הקיים נשמר',
   !r10.repeat&&st10&&st10.ap==='other',
   `apStatus=${st10&&st10.ap} · drift=${st10&&st10.drift} · ROP ${st10&&st10.rop} מול הצעה ${st10&&st10.sug}`);

 /* ── 11. היסטוריה בפורמט הישן ─────────────────────────────────── */
 const seeded=await p.evaluate(()=>{
  /* רשומה בפורמט הישן בדיוק: «חודש|מספר|סכומים», בלי v. התמונה
     נשמרת כפי שהיא — היא אינה תלויה בפורמט החתימה. */
  const lastPar=PARAMS.runs[PARAMS.runs.length-1];
  HIST.runs=[{sig:'2026-9|28|1|2|3|4',ts:Date.now()-40*864e5,n:28,q:9}];
  HIST.items={'ITEM-000':{since:Date.now()-40*864e5,seen:3,miss:1,prevMiss:0,cat:'בדיקה'}};
  PARAMS.runs=[{sig:'2026-9|28|1|2|3|4',ts:Date.now()-40*864e5,it:lastPar.it}];
  saveHist();saveParams();
  return {hist:HIST.runs.length,par:PARAMS.runs.length,
    items:Object.keys(HIST.items).length,snap:Object.keys(lastPar.it).length}});
 const marksBefore=await p.evaluate(()=>Object.keys(MARKS).length);
 const r11=await load(write(F('legacy'),clone(partial)));
 const mchip=await chip();
 const after=await p.evaluate(()=>({
   marks:Object.keys(MARKS).length,
   oldKept:HIST.runs.some(r=>r.sig==='2026-9|28|1|2|3|4'),
   oldParKept:PARAMS.runs.some(r=>r.sig==='2026-9|28|1|2|3|4'),
   chip:(document.getElementById('histChip')||{}).textContent||''}));
 ok('מעבר מפורמט ישן — לא נמחקת שום רשומה',
   after.oldKept&&after.oldParKept&&r11.hist===2&&r11.par===2,
   `HIST ${r11.hist} · PARAMS ${r11.par} · הישנות נשמרו=${after.oldKept&&after.oldParKept}`);
 ok('וגם לא סימון של המשתמש',after.marks===marksBefore&&marksBefore>0,
   `${marksBefore}→${after.marks} סימונים`);
 ok('לא נקבע «אותו דוח» על סמך החתימה הישנה',
   !r11.repeat&&r11.migrated,`repeat=${r11.repeat} · migrated=${r11.migrated}`);
 ok('המעבר נאמר במפורש על המסך',
   !mchip.hidden&&/בסיס חדש/.test(mchip.txt),
   `מוצגת=${!mchip.hidden} · «${mchip.txt.slice(0,46)}»`);
 ok('וההשוואה מהתמונה הישנה עדיין מתבצעת',
   !!r11.applied&&r11.applied.total>0,
   `${r11.applied?r11.applied.total:'—'} המלצות הושוו מהתמונה שנשמרה בפורמט הישן`);

 /* ── 12. חוזה רשימת השדות ─────────────────────────────────────── */
 /* הבדיקות שלמעלה מפעילות את המסלול; זו שומרת על *הרשימה*. שדה
    שיימחק מ-DSIG_FIELDS בעריכה עתידית יחזיר בדיוק את הבאג הזה, בלי
    שאף תרחיש ייפול — כי כל תרחיש בודק שדה אחד. */
 const req=['rop','ss','ssMin','ltD','srv','srvMissing','mrp','st','stTxt',
   'stock','free','transfer','transferKnown','cust','custRaw','blockedCust',
   'custBlockedKnown','po','months','curMonth','price','currency','sales','entries'];
 const miss=await p.evaluate(r=>r.filter(f=>DSIG_FIELDS.indexOf(f)<0),req);
 ok('רשימת השדות כוללת את כל מה שהמפרט דורש',miss.length===0,
   miss.length?'חסרים: '+miss.join(', '):`${req.length} שדות נבדקו · ${await p.evaluate(()=>DSIG_FIELDS.length)} ברשימה`);

 const isolated=await p.evaluate(()=>{
  const c=ALL.map(x=>Object.assign({},x)),s0=dataSig(c);
  /* תוצאות חישוב של המנוע — שינוי בהן אינו דוח חדש */
  c[0].sugROP=(c[0].sugROP||0)+500;c[0].sugSS=(c[0].sugSS||0)+500;
  c[0].rate=(c[0].rate||0)+77;c[0].cat='משהו אחר';c[0].paramFix=!c[0].paramFix;
  const sEngine=dataSig(c);
  /* ומפריד שמסתתר בתוך טקסט אינו יכול להזיז את גבול השדות */
  const d=ALL.map(x=>Object.assign({},x));
  d[1].desc='AB';d[2].desc='A';d[2].desc2='B'+(d[2].desc2||'');
  const e=ALL.map(x=>Object.assign({},x));
  e[1].desc='A';e[1].desc2='B'+(e[1].desc2||'');e[2].desc='AB';
  return {same:s0===sEngine,sepDistinct:dataSig(d)!==dataSig(e)}});
 ok('שינוי בתוצאת חישוב של המנוע אינו דוח חדש',isolated.same,
   'sugROP · sugSS · rate · cat · paramFix שונו — החתימה זהה');
 ok('מפריד בתוך טקסט אינו מזיז את גבול השדות',isolated.sepDistinct,
   'שני סידורים שונים של אותם תווים — חתימות שונות');

 /* ── 13. אין שגיאות JS לאורך כל המסלול ────────────────────────── */
 ok('אף שגיאת JS לאורך כל המסלול',errs.length===0,errs.slice(0,2).join(' · ')||'נקי');

 await b.close();
 for(const l of out)console.log(l);
 const bad=out.filter(l=>l.startsWith('FAIL')).length;
 console.log(`\n${out.length-bad}/${out.length} עברו`);
 process.exit(bad?1:0);
})().catch(e=>{console.log('FAIL · הבדיקה קרסה · '+e.message);process.exit(1)});
