/* ============ NO_CDN — הכלי חייב לעבוד בלי שום רשת ============
   הבדיקה שהייתה תופסת את התקלה שהמעתד דיווח עליה בבוקר 4.10.
   הכלי טען את SheetJS מ-cdn.sheetjs.com בכל פתיחה. ברשת ארגונית
   שחוסמת את הדומיין, ב-VPN, או כש-CDN למטה — XLSX אינו מוגדר,
   הטעינה נופלת, והמסך האשים את הקובץ של המעתד: «ודא שהקובץ תקין
   ואינו מוגן».
   כאן *כל* בקשה שאינה file:// נחסמת. אם הכלי עדיין קורא דוח ZMRP
   מלא — הוא באמת עצמאי, כפי שהמפרט דורש. */
const {chromium}=require('playwright'),fs=require('fs'),path=require('path');
const SD=__dirname, HTML='file://'+path.join(SD,'..','index.html');
const out=[];const ok=(n,c,d)=>out.push(`${c?'PASS':'FAIL'} · ${n}${d?'  ['+d+']':''}`);
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 /* חומת האש של הבדיקה: שום דבר מחוץ לקובץ לא עובר */
 const blocked=[];
 await ctx.route('**/*',r=>{const u=r.request().url();
   if(u.startsWith('file://')||u.startsWith('data:')||u.startsWith('blob:'))return r.continue();
   blocked.push(u.slice(0,80));return r.abort()});
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message.slice(0,200)));
 await p.goto(HTML+'?nobrief=1');await p.waitForTimeout(2500);

 const lib=await p.evaluate(()=>({t:typeof XLSX,v:(typeof XLSX!=='undefined'&&XLSX)?XLSX.version:'',
   read:(typeof XLSX!=='undefined'&&XLSX)?typeof XLSX.read:''}));
 ok('הספרייה קיימת בלי שום בקשת רשת',lib.t==='object'&&lib.read==='function',
    `XLSX ${lib.t} · גרסה ${lib.v}`);
 /* הגרסה ננעלת: החלפה שקטה של build היא בדיוק מה שה-CDN עשה */
 ok('הגרסה היא זו שהכלי נבדק מולה',lib.v==='0.20.2',lib.v);
 ok('אין בקשה ל-cdn.sheetjs.com בכלל',
    !blocked.some(u=>/sheetjs/.test(u)),blocked.join(' · ')||'—');

 /* ודוח אמיתי באמת נקרא — לא רק שהאובייקט קיים */
 /* הדוח של make-fixture, שרץ ראשון ב-npm test. eta-zmrp.xlsx נוצר
    רק בתוך eta.test.js — כלומר הוא אינו קיים בצ'קאאוט נקי, וזה
    בדיוק מה שהפיל את הריצה הראשונה ב-CI תוך ארבע שניות. */
 const FIX=SD+'/zmrp-demo.xlsx';
 if(!fs.existsSync(FIX)){console.log('FAIL · חסר קובץ הבדיקה zmrp-demo.xlsx — הרץ npm run fixture');process.exit(1)}
 await p.setInputFiles('#f',FIX);await p.waitForTimeout(4000);
 const st=await p.evaluate(()=>({n:(typeof ALL!=='undefined'&&ALL)?ALL.length:0,
   diag:(document.getElementById('diag')||{}).textContent||'',
   upd:(document.getElementById('updated')||{}).textContent||''}));
 ok('דוח ZMRP נקרא במלואו בלי רשת',st.n>0,`${st.n} שורות · «${st.upd}»`);
 ok('ואין הודעת שגיאה על המסך',!/⛔/.test(st.diag),st.diag.slice(0,120));

 /* ============ LOAD_ERR_HONEST ============
    ההודעה חייבת להבחין בין «הספרייה לא נטענה» ל«הקובץ פגום».
    היא נבדקת ישירות, כי אי אפשר יותר לייצר את המצב הראשון מבחוץ —
    וזה בדיוק העניין. */
 /* ============ LOAD_ERR_SCOPE ============
    ה-try סביב הטעינה עוטף גם את הסיווג, הרינדור ושחזור המסך — ולכן
    באג שלנו *אחרי* קריאת הקובץ הוצג כ«שגיאה בקריאת הקובץ» והאשים
    את האקסל של המעתד. שלוש משפחות הכשל חייבות להיקרא שונה. */
 const kinds=await p.evaluate(()=>{
   const g=e=>({head:loadErrHead(e),txt:loadErrTxt(e)});
   const file=g(new Error('Corrupted zip: missing End of Central Directory'));
   const app=g(new TypeError("Cannot read properties of null (reading 'length')"));
   const saved=window.XLSX;window.XLSX=undefined;const lib=g(new Error('x'));window.XLSX=saved;
   return {file,app,lib,build:BUILD}});
 ok('קובץ פגום — הכותרת מצביעה על הקובץ',
    /הקובץ לא נקרא/.test(kinds.file.head)&&/Corrupted zip/.test(kinds.file.txt),
    kinds.file.head+' | '+kinds.file.txt.slice(0,90));
 ok('תקלה שלנו אחרי הקריאה — הכותרת אומרת שזו תקלה בכלי',
    /תקלה בכלי/.test(kinds.app.head)&&/לא בקובץ שלך/.test(kinds.app.txt),
    kinds.app.head+' | '+kinds.app.txt.slice(0,90));
 ok('ואינה מבקשת מהמעתד לבדוק את הקובץ שלו',
    !/ודא שהוא תקין|מוגן בסיסמה/.test(kinds.app.txt),kinds.app.txt.slice(0,120));
 ok('ספרייה חסרה — משפחה שלישית, נפרדת',
    /הספרייה לא נטענה/.test(kinds.lib.head),kinds.lib.head);
 ok('כל הודעת שגיאה נושאת את מזהה הבנייה',
    [kinds.file.txt,kinds.app.txt,kinds.lib.txt].every(t=>t.includes(kinds.build)),kinds.build);

 /* ============ RENDER_NEEDS_DATA ============
    המעתד פתח את הכלי במסך שבו עבד אתמול והעלה דוח — והכלי קרס.
    ה-stack ששלח: decisionList -> currentRows -> render -> setMode
    -> rd.onload, עם «Cannot read properties of null (reading
    'quality')». סדר הטעינה קורא ל-setMode לפני apply, ו-QF נבנה רק
    בתוך apply; ה-render הראשון של כל העלאה רץ על null.
    נמדד על הבנייה שהייתה בייצור: 4 מתוך 17 המסכים השמורים — qual,
    catalog, burn, done — הפילו את הטעינה על כל קובץ.
    כאן נבדק *כל* מסך כמצב שמור, בהקשר דפדפן נקי, בדיוק כמו מעתד
    שפותח בבוקר. אם אחד מהם ייפול שוב, זה ייתפס כאן ולא אצלו. */
 const MODES=['home','today','line','month','qual','stale','cap','catalog',
   'moves','burn','cust','done','floor','trend','rise','applied',
  /* דלתות שנוספו אחרי שהרשימה נכתבה ולא נכנסו אליה: הפיצולים.
     מסך שאינו ברשימה הזאת אינו נבדק כמצב שמור בכלל. */
  'over3y','tier','floorNR','capPO','capStop','capBiz','capNone'];
 const bad=[];
 for(const m of MODES){
  const c2=await b.newContext({viewport:{width:1512,height:860}});
  await c2.route('**/*',r=>{const u=r.request().url();
    return (u.startsWith('file://')||u.startsWith('data:')||u.startsWith('blob:'))?r.continue():r.abort()});
  const q=await c2.newPage();const e2=[];
  q.on('pageerror',e=>e2.push(e.message.split('\n')[0].slice(0,100)));
  await q.goto(HTML);await q.waitForTimeout(800);
  await q.evaluate(mm=>{try{localStorage.setItem('planner_mode_v1',mm)}catch(_){}} ,m);
  await q.reload();await q.waitForTimeout(1200);
  await q.setInputFiles('#f',FIX);await q.waitForTimeout(3500);
  const r2=await q.evaluate(()=>({n:(typeof ALL!=='undefined'&&ALL)?ALL.length:0,
    upd:(document.getElementById('updated')||{}).textContent||''}));
  if(!(r2.n>0)||/תקלה|שגיאה/.test(r2.upd)||e2.length)
   bad.push(`${m}: ${r2.n} שורות · «${r2.upd}»${e2.length?' · '+e2[0]:''}`);
  await c2.close();
 }
 ok('כל מסך שמור נפתח ומעלה דוח בלי לקרוס',bad.length===0,
    bad.join(' || ')||`${MODES.length} מסכים`);

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 console.log(out.join('\n'));
 const f=out.filter(x=>x.startsWith('FAIL')).length;
 console.log(`\n${out.length-f}/${out.length} עברו`);
 await b.close();process.exit(f?1:0)})();
