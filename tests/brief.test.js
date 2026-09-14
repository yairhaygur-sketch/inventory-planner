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

 /* --- טעינה ראשונה: אין היסטוריה, ולכן התדריך חייב להופיע --- */
 const p=await open(ctx,'');
 const st=await p.evaluate(()=>({
   shown:!document.getElementById('brief').hidden,
   body:document.body.classList.contains('briefing'),
   why:document.getElementById('briefSub').textContent.trim(),
   cards:[...document.querySelectorAll('#briefCards .bcard')].map(c=>({
     l:c.querySelector('.bl').textContent.trim(),
     v:c.querySelector('.bv').textContent.trim()})),
   sev3:(QF&&QF.all?QF.all:[]).filter(r=>r.sev===3).length,
   items:ALL.length}));
 ok('בטעינה ראשונה התדריך מופיע',st.shown&&st.body,st.why);
 ok('התדריך מנמק למה הוא כאן',st.why.length>0,st.why);
 ok('כרטיס הדוח מציג את מספר המק״טים',
   st.cards.some(c=>c.l==='הדוח'&&c.v===st.items.toLocaleString('he-IL')),
   JSON.stringify(st.cards.find(c=>c.l==='הדוח')));
 /* המספר בכרטיס חייב להיות המספר האמיתי, לא טקסט קבוע */
 ok('כרטיס "בוער עכשיו" תואם את ספירת חומרה 3',
   st.cards.some(c=>c.l==='בוער עכשיו'&&c.v===st.sev3.toLocaleString('he-IL')),
   'בכרטיס '+(st.cards.find(c=>c.l==='בוער עכשיו')||{}).v+' · בפועל '+st.sev3);
 ok('אין כרטיס ריק',st.cards.every(c=>c.v&&c.v!=='undefined'&&c.v!=='NaN'),
   st.cards.map(c=>c.l+'='+c.v).join(' · '));

 /* --- הכפתור מסיר את התדריך ומחזיר את הטבלה לשימוש --- */
 await p.click('#briefGo');await p.waitForTimeout(300);
 const after=await p.evaluate(()=>({
   hidden:document.getElementById('brief').hidden,
   body:document.body.classList.contains('briefing'),
   rows:document.querySelectorAll('#tbl tbody tr[data-i]').length}));
 ok('«התחל לעבוד» מסיר את התדריך',after.hidden&&!after.body);
 ok('הטבלה מלאה מתחתיו',after.rows>0,after.rows+' שורות');
 /* אחרי הסגירה השורה חייבת להיות לחיצה — זה מה שנשבר אם המסך נשאר */
 await p.click('#tbl tbody tr[data-i]');await p.waitForTimeout(400);
 ok('אפשר ללחוץ על שורה אחרי הסגירה',
   await p.evaluate(()=>document.body.classList.contains('dopen')));
 await p.close();

 /* --- אותו דוח שוב, מיד: אין שינוי, ולכן אין תדריך --- */
 const p2=await open(ctx,'');
 const again=await p2.evaluate(()=>({
   shown:!document.getElementById('brief').hidden,
   why:document.getElementById('briefSub').textContent.trim(),
   repeat:HRUN&&HRUN.repeat,
   diag:(DIAG&&DIAG.warn)?DIAG.warn.join(' | '):'',
   runs:HIST.runs.length}));
 ok('טעינה חוזרת של אותו דוח אינה מציגה תדריך',!again.shown,
   `repeat=${again.repeat} · סיבה="${again.why}" · diag="${again.diag}"`);
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
