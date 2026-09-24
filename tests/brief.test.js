/* ============ תדריך הפתיחה ============
   הבדיקות האחרות טוענות עם ?nobrief=1 כדי שהתדריך לא יחסום אותן.
   הקובץ הזה הוא היחיד שטוען *בלי* הפרמטר, ולכן הוא היחיד שבודק את
   התכונה עצמה. בלעדיו ?nobrief הופך לדרך לא לבדוק כלום. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
const open=async(ctx,q)=>{
 const p=await ctx.newPage();
 await p.goto('file://'+path.join(SD,'..','index.html')+q);
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');
 await p.waitForSelector('#tbl tbody tr[data-i]',{timeout:120000});
 await p.waitForTimeout(600);
 return p};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));

 /* ============ התדריך אינו חוסם יותר ============
    קודם נבדק כאן שהתדריך *חוסם* (`body.briefing`) ושלחיצה על «התחל
    לעבוד» מסירה אותו. הוא אינו מחליט דבר ואינו מונע שום טעות — הוא
    מסכם, ולכן הפך לרצועה מקופלת. מה שנבדק עכשיו: שאחרי טעינה תקינה
    נכנסים ישר לעבודה, ושהסיכום עדיין שם ונכון. */
 const p=await open(ctx,'');
 const st=await p.evaluate(()=>({
   shown:!document.getElementById('brief').hidden,
   blocking:document.body.classList.contains('briefing'),
   rows:document.querySelectorAll('#tbl tbody tr[data-i]').length,
   why:document.getElementById('briefSub').textContent.trim(),
   openAttr:document.getElementById('briefToggle').getAttribute('aria-expanded'),
   cards:[...document.querySelectorAll('#briefCards .bcard')].map(c=>({
     l:c.querySelector('.bl').textContent.trim(),
     v:c.querySelector('.bv').textContent.trim()})),
   runs:HIST.runs.length}));
 ok('טעינת דוח תקין מובילה ישירות לעבודה',!st.blocking&&st.rows>0,
   `חוסם=${st.blocking} · ${st.rows} שורות`);
 ok('רצועת הסיכום מוצגת',st.shown,st.why);
 /* --- בהעלאה הראשונה אין בסיס להשוואה, ואסור להציג פריטים כשינוי --- */
 ok('בהעלאה הראשונה נאמר שאין דוח קודם להשוואה',
   /אין עדיין דוח קודם|אין בסיס/.test(st.why),st.why);
 ok('ולא מוצג אף פריט כ«שינוי» מול בסיס שאינו קיים',
   st.runs<=1?st.cards.every(c=>!/חדשים|החמירו|יצאו/.test(c.l)):true,
   st.cards.map(c=>c.l+'='+c.v).join(' · '));
 ok('אין כרטיס ריק',st.cards.every(c=>c.v&&c.v!=='undefined'&&c.v!=='NaN'),
   st.cards.map(c=>c.l+'='+c.v).join(' · '));
 /* --- השורה לחיצה בלי שום צעד ביניים --- */
 await p.click('#tbl tbody tr[data-i]');await p.waitForTimeout(400);
 ok('אפשר ללחוץ על שורה בלי לסגור כלום',
   await p.evaluate(()=>document.body.classList.contains('dopen')));
 await p.evaluate(()=>closeDetail());await p.waitForTimeout(250);
 /* --- הרצועה נפתחת ונסגרת, וגם במקלדת --- */
 const before=await p.evaluate(()=>document.getElementById('briefToggle').getAttribute('aria-expanded'));
 await p.evaluate(()=>document.getElementById('briefToggle').focus());
 await p.keyboard.press('Enter');await p.waitForTimeout(250);
 const afterK=await p.evaluate(()=>({
   exp:document.getElementById('briefToggle').getAttribute('aria-expanded'),
   cards:document.getElementById('briefCards').hidden}));
 ok('הרצועה נפתחת ונסגרת במקלדת',afterK.exp!==before&&afterK.cards===(afterK.exp!=='true'),
   `${before} → ${afterK.exp}`);
 await p.close();

 /* --- אותו דוח שוב, מיד: אין שינוי, ולכן אין תדריך --- */
 const p2=await open(ctx,'');
 const again=await p2.evaluate(()=>({
   shown:!document.getElementById('brief').hidden,
   why:document.getElementById('briefSub').textContent.trim(),
   repeat:HRUN&&HRUN.repeat,
   diag:(DIAG&&DIAG.warn)?DIAG.warn.join(' | '):'',
   runs:HIST.runs.length}));
 /* קודם: «טעינה חוזרת אינה מציגה תדריך». הרצועה מוצגת תמיד — היא
    אינה חוסמת — ומה שמשתנה הוא מה שהיא אומרת. */
 ok('טעינה חוזרת של אותו דוח אומרת שאין מה להשוות',
   again.shown&&/אותו דוח נטען שוב|אין מה להשוות|אין עדיין דוח קודם/.test(again.why),
   `repeat=${again.repeat} · סיבה="${again.why}"`);
 ok('הטבלה זמינה ישירות',
   (await p2.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length))>0);
 await p2.close();

 /* --- הפרמטר מדלג --- */
 const p3=await open(ctx,'?nobrief=1');
 ok('?nobrief=1 מדלג על התדריך',
   await p3.evaluate(()=>document.getElementById('brief').hidden));
 const errs=[];p3.on('pageerror',e=>errs.push(e.message));
 await p3.waitForTimeout(300);
 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 await p3.close();

 console.log(out.join('\n'));
 await b.close();
 process.exit(bad?1:0);
})();
