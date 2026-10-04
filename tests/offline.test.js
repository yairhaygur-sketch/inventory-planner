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
 const msgs=await p.evaluate(()=>{
   const real=loadErrTxt(new Error('Corrupted zip: missing End of Central Directory'));
   const saved=window.XLSX;window.XLSX=undefined;
   const noLib=loadErrTxt(new Error('XLSX is not defined'));
   window.XLSX=saved;return {real,noLib}});
 ok('קובץ פגום — ההודעה מצביעה על הקובץ ונושאת את הסיבה',
    /הקובץ|מוגן/.test(msgs.real)&&/Corrupted zip/.test(msgs.real),msgs.real.slice(0,140));
 ok('ספרייה חסרה — ההודעה אומרת שזו תקלה בכלי, לא בקובץ',
    /תקלה בכלי/.test(msgs.noLib)&&!/ודא שהקובץ תקין/.test(msgs.noLib),msgs.noLib.slice(0,140));

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 console.log(out.join('\n'));
 const f=out.filter(x=>x.startsWith('FAIL')).length;
 console.log(`\n${out.length-f}/${out.length} עברו`);
 await b.close();process.exit(f?1:0)})();
