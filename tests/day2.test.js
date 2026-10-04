/* ============ DAY2 — היום השני ============
   כל שאר הסוויטה טוענת דוח אחד לתוך דפדפן ריק. הבאג של 4.10 חי בדיוק
   במה שהיא לא רואה: מעתד שחוזר למחרת — סימונים, דוח ETA שמור,
   היסטוריה, פריסת עמודות, סינון ומסך אחרון — ומעלה דוח *אחר*.
   ה-stack שהגיע מהשטח:
     TypeError: Cannot read properties of null (reading 'quality')
       at decisionList -> currentRows -> render -> setMode -> rd.onload
   ארבעה מסכים שמורים הפילו כל העלאה, והסוויטה הירוקה לא ראתה דבר.

   הקובץ הזה בונה מצב צבור על ידי *שימוש* בכלי — לא על ידי כתיבה
   ישירה ל-localStorage — ואז פותח דף חדש באותו אחסון ומעלה את הדוח
   של מחר. אחריו, מטריצת הרשאות: כל מפתח אחסון לבדו. */
const {chromium}=require('playwright'),fs=require('fs'),path=require('path');
const SD=__dirname, HTML='file://'+path.join(SD,'..','index.html');
const D1=SD+'/zmrp-demo.xlsx', D2=SD+'/zmrp-demo-day2.xlsx', ETA=SD+'/zmrp-demo-eta.xlsx';
const out=[];const ok=(n,c,d)=>out.push(`${c?'PASS':'FAIL'} · ${n}${d?'  ['+d+']':''}`);
const MODES=['home','today','line','month','qual','stale','cap','catalog',
  'moves','burn','cust','done','floor','trend','rise','applied'];
const KEYS=['planner_mode_v1','planner_marks_v1','planner_eta_v1','planner_history_v1',
  'planner_params_v1','planner_layout_v2','planner_cols_v1','planner_colw_v1',
  'planner_ledger_v1','planner_dark_v1','planner_groups_v1','planner_band_v1',
  'planner_groupby_v1','planner_recent_q_v1','planner_tour_v1'];

(async()=>{
 const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const stub=c=>c.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 for(const f of [D1,D2,ETA])if(!fs.existsSync(f)){
   console.log('FAIL · חסר קובץ בדיקה '+path.basename(f)+' — הרץ npm run fixture');process.exit(1)}

 /* ── יום ראשון: עבודה אמיתית, דרך הממשק ───────────────────────── */
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await stub(ctx);
 const errs1=[];
 let p=await ctx.newPage();p.on('pageerror',e=>errs1.push(e.message.split('\n')[0].slice(0,120)));
 await p.goto(HTML);await p.waitForTimeout(2000);
 await p.setInputFiles('#f',D1);await p.waitForTimeout(6000);
 await p.setInputFiles('#fe',ETA);await p.waitForTimeout(2500);
 const day1=await p.evaluate(()=>({n:ALL.length,tasks:decisionList('short').length}));
 ok('יום ראשון נטען',day1.n>0,`${day1.n} פריטים`);

 /* סימונים, פריסה, סינון, מיון, צפיפות, כהה — כמו מעתד שעבד */
 const built=await p.evaluate(()=>{
  const pick=decisionList('short').slice(0,6).map(r=>r.pn);
  setMark(pick[0],'handled','');setMark(pick[1],'checking','');
  setMark(pick[2],'ordered','בדיקה');setMark(pick[3],'supplier','');
  try{toggleWidget('kpis')}catch(_){}
  try{colFilters['c:cat']={v:new Set(['אוזל החודש'])}}catch(_){}
  try{sortCol='val';sortDir='asc'}catch(_){}
  try{localStorage.setItem('planner_dark_v1','1')}catch(_){}
  return {marked:Object.keys(MARKS).length,pick}});
 ok('סימונים נשמרו ביום הראשון',built.marked>=4,`${built.marked} סימונים`);

 /* כל מסך נפתח — וה*אחרון* הוא זה שנשמר. qual הוא אחד מארבעת
    המסכים שהפילו את הטעינה בייצור, ולכן הוא נבחר במכוון. */
 for(const m of MODES){await p.evaluate(mm=>setMode(mm),m);await p.waitForTimeout(120)}
 await p.evaluate(()=>setMode('qual'));await p.waitForTimeout(400);
 const stored=await p.evaluate(()=>Object.keys(localStorage).sort());
 ok('המצב נצבר באחסון',stored.length>=5,stored.join(' · '));
 ok('והמסך האחרון הוא qual',
    await p.evaluate(()=>localStorage.getItem('planner_mode_v1'))==='qual');
 ok('אין שגיאות ביום הראשון',errs1.length===0,errs1.join(' | '));
 await p.close();

 /* ── יום שני: דף חדש, אותו אחסון, דוח אחר ─────────────────────── */
 const errs2=[];
 p=await ctx.newPage();p.on('pageerror',e=>errs2.push(e.message.split('\n')[0].slice(0,120)));
 await p.goto(HTML);await p.waitForTimeout(2000);
 const boot=await p.evaluate(()=>({mode:(typeof mode!=='undefined')?mode:'',
   eta:!!ETA,marks:Object.keys(MARKS||{}).length,dark:document.body.classList.contains('dark')}));
 ok('דוח ה-ETA שרד את סגירת הדף',boot.eta);
 ok('והסימונים שרדו',boot.marks>=4,`${boot.marks} סימונים`);
 ok('הפתיחה עצמה לא נפלה',errs2.length===0,errs2.join(' | '));

 await p.setInputFiles('#f',D2);await p.waitForTimeout(7000);
 const d2=await p.evaluate(()=>({n:(typeof ALL!=='undefined'&&ALL)?ALL.length:0,
   upd:(document.getElementById('updated')||{}).textContent||'',
   diag:(document.getElementById('diag')||{}).textContent||'',
   mode:(typeof mode!=='undefined')?mode:'',
   marks:Object.keys(MARKS||{}).length,
   hist:HRUN?{runs:HRUN.runs,newN:HRUN.newN,backN:HRUN.backN,resolved:HRUN.resolved}:null,
   rows:document.querySelectorAll('#tbl tbody tr[data-i]').length}));
 /* זו הטענה שהייתה תופסת את התקלה של 4.10 */
 ok('הדוח של מחר נטען על מצב צבור, במסך השמור',
    d2.n>0&&!/תקלה|שגיאה/.test(d2.upd),`${d2.n} פריטים · «${d2.upd}» · מסך ${d2.mode}`);
 ok('ונפתח באותו מסך שבו עבד אתמול',d2.mode==='qual',d2.mode);
 ok('הסימונים עדיין שם אחרי דוח חדש',d2.marks>=4,`${d2.marks}`);
 ok('הטבלה מציגה שורות',d2.rows>0,`${d2.rows} שורות`);

 /* ההשוואה בין הימים חייבת להיות לא-טריוויאלית — אחרת לא נבדקה */
 ok('ההיסטוריה סופרת שתי העלאות',!!d2.hist&&d2.hist.runs>=2,JSON.stringify(d2.hist));
 ok('ויש בה תנועה אמיתית בין הימים',
    !!d2.hist&&(d2.hist.newN+d2.hist.backN+d2.hist.resolved)>0,JSON.stringify(d2.hist));

 /* סריקת כל המסכים — על המצב הצבור, לא על דפדפן נקי */
 const sweep=[];
 for(const m of MODES){
  const e0=errs2.length;
  await p.evaluate(mm=>setMode(mm),m);await p.waitForTimeout(250);
  const r=await p.evaluate(()=>({rows:document.querySelectorAll('#tbl tbody tr').length,
    head:document.querySelectorAll('#tbl thead th').length}));
  if(errs2.length>e0||r.head===0)sweep.push(`${m}(${r.head} עמודות${errs2.length>e0?' · '+errs2[errs2.length-1]:''})`);
 }
 ok('כל מסך נפתח על המצב הצבור',sweep.length===0,sweep.join(' · ')||`${MODES.length} מסכים`);

 /* כרטיס פריט, ייצוא ובורר עמודות — על היום השני */
 const deep=await p.evaluate(()=>{
  const r=[];
  try{setMode('today');const x=currentRows()[0];if(x)detail(x);
    const d=document.querySelector('.wdetail')||document.querySelector('#detail');
    const t=d?d.innerText:'';
    r.push((/ההחלטה/.test(t)&&/מה מצדיק את זה/.test(t)&&t.length>200)?'':'כרטיס ריק')}
  catch(e){r.push('כרטיס: '+e.message.slice(0,60))}
  try{const v=viewExportRows();r.push((v&&v.aoa&&v.aoa.length>1)?'':'ייצוא ריק')}
  catch(e){r.push('ייצוא: '+e.message.slice(0,60))}
  try{const n=exportRows(currentRows());r.push(n.length>1?'':'גיליון מלא ריק')}
  catch(e){r.push('גיליון מלא: '+e.message.slice(0,60))}
  return r.filter(Boolean)});
 ok('כרטיס הפריט והייצוא עובדים ביום השני',deep.length===0,deep.join(' · '));
 /* ============ OUTCOME ============
    «מה יושם» היה חצי שאלה. החצי השני — «ומה קרה אחר כך» — נמדד
    מתמונת הפרמטרים, שמחזיקה מעכשיו גם את מצב החוסר ברגע הצילום.
    שני הסייגים נבדקים כאן: שהתמונה באמת מדדה (ולא אפס, כפי
    שקרה כש-inShortQ נשען על QF שעוד לא הוצב), ושהמסך אינו מציג
    אחוזים על מדגם קטן מדי. */
 await p.evaluate(()=>setMode('applied'));await p.waitForTimeout(800);
 const oc=await p.evaluate(()=>{
   const A=APPLIED;
   return {has:!!A,o:A&&A.outcome?A.outcome:null,txt:typeof outcomeTxt==='function'?outcomeTxt(A):'',
     sub:(document.getElementById('phdSub')||{}).textContent||''}});
 ok('יש השוואת פרמטרים בין שתי ההעלאות',!!oc.has&&!!oc.o,JSON.stringify(oc.o));
 ok('התמונה הקודמת מדדה מצב חוסר — ולא אפס',
    !!oc.o&&oc.o.measured>0&&(oc.o.doneN+oc.o.restN)>0,JSON.stringify(oc.o));
 ok('התוצאה נספרת בשתי הקבוצות',
    !!oc.o&&(oc.o.doneShort+oc.o.restShort)>=0&&oc.o.restN>0,JSON.stringify(oc.o));
 /* הסייג שאסור לוותר עליו: אחוז מוצג רק כשיש על מה */
 const thin=!!oc.o&&(oc.o.doneN<50||oc.o.restN<50||oc.o.doneShort<10||oc.o.restShort<10);
 ok('מדגם קטן — נאמר «אין מה להשוות» ולא אחוז',
    !thin||(/מעט מדי כדי להשוות|אין מה להשוות/.test(oc.txt)&&!/%/.test(oc.txt)),
    `דק=${thin} · ${oc.txt}`);
 ok('וגם «אפס יושמו» נאמר ולא נבלע',
    (oc.o&&oc.o.doneN>0)||/אף המלצה לא יושמה/.test(oc.txt),oc.txt);
 ok('מדגם מספיק — נאמר גם שזו תצפית ולא ניסוי',
    thin||/תצפית ולא ניסוי/.test(oc.txt),oc.txt);

 ok('אין שגיאות JS ביום השני',errs2.length===0,errs2.join(' | '));
 await ctx.close();

 /* ── מטריצה: כל מפתח אחסון לבדו ────────────────────────────────
    מצב צבור אינו גוש אחד. מעתד יכול להגיע עם מסך שמור בלבד, או עם
    פריסת עמודות בלבד. כל מפתח נבדק כשהוא היחיד שנשאר. */
 const perKey=[];
 for(const k of KEYS){
  const c2=await b.newContext({viewport:{width:1512,height:860}});
  await stub(c2);
  const q=await c2.newPage();const e2=[];
  q.on('pageerror',e=>e2.push(e.message.split('\n')[0].slice(0,90)));
  await q.goto(HTML);await q.waitForTimeout(1200);
  await q.setInputFiles('#f',D1);await q.waitForTimeout(5000);
  await q.evaluate(()=>setMode('qual'));await q.waitForTimeout(300);
  /* משאירים מפתח אחד בלבד */
  await q.evaluate(kk=>{const v=localStorage.getItem(kk);
    localStorage.clear();if(v!=null)localStorage.setItem(kk,v)},k);
  await q.reload();await q.waitForTimeout(1500);
  const before=e2.length;
  await q.setInputFiles('#f',D2);await q.waitForTimeout(5000);
  const st=await q.evaluate(()=>({n:(typeof ALL!=='undefined'&&ALL)?ALL.length:0,
    upd:(document.getElementById('updated')||{}).textContent||''}));
  if(!(st.n>0)||/תקלה|שגיאה/.test(st.upd)||e2.length>before)
   perKey.push(`${k}: ${st.n} · «${st.upd}»${e2.length>before?' · '+e2[e2.length-1]:''}`);
  await c2.close();
 }
 ok('כל מפתח אחסון לבדו שורד טעינת דוח',perKey.length===0,
    perKey.join(' || ')||`${KEYS.length} מפתחות`);

 console.log(out.join('\n'));
 const f=out.filter(x=>x.startsWith('FAIL')).length;
 console.log(`\n${out.length-f}/${out.length} עברו`);
 await b.close();process.exit(f?1:0)})();
