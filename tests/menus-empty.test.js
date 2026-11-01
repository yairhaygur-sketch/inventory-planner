/* ============ תפריטי הסרגל לפני שנטען דוח ============
   מסך הפתיחה (#welcome, z 420) כיסה את התפריטים הפתוחים של הסרגל (z 95):
   התפריט נפתח, אבל לא נראה ולא נלחץ. כך לא היה אפשר לטעון דוח ETA,
   לתקן חודש בסיס או להגיע לאף פריט ב«עזרה ועוד» לפני העלאת ZMRP.
   נכנס עם D1 (#135), כשכפתור ה-ETA עבר לתוך תפריט «הדוח הטעון».
   הבדיקה עוברת כמשתמש: פותחת את התפריט, בודקת שכל פריט בו הוא האלמנט
   העליון בנקודה שלו, ולוחצת באמת על «דוח ETA» עד שהדוח נטען. */
const {chromium}=require('playwright');const fs=require('fs');const path=require('path');
const SD=__dirname,HTML='file://'+path.join(SD,'..','index.html');
let pass=0,fail=0;const ok=(n,c,d)=>{c?pass++:fail++;console.log(`${c?'PASS':'FAIL'} · ${n}${d?'  ['+d+']':''}`)};
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
/* כל פריט גלוי בתפריט — האם הוא (או ילד שלו) מה שנמצא מעל מרכזו */
const COVER=id=>{const m=document.getElementById(id);
 const items=[...m.querySelectorAll('button,label,a,input[type=month]')].filter(e=>e.getBoundingClientRect().width>0);
 const bad=items.filter(e=>{const r=e.getBoundingClientRect(),t=document.elementFromPoint(r.x+r.width/2,r.y+r.height/2);
  return !(t&&(t===e||e.contains(t)||t.contains(e)))});
 return {n:items.length,bad:bad.map(e=>e.textContent.trim().slice(0,24))}};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 for(const [w,h] of [[1366,768],[1440,900],[1920,1080]]){
  const ctx=await b.newContext({viewport:{width:w,height:h}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto(HTML+'?nobrief=1');await p.waitForTimeout(1200);
  const wel=await p.evaluate(()=>!document.getElementById('welcome').hidden);
  ok(`${w}px · מסך הפתיחה מוצג (זה המצב שנבדק)`,wel);
  for(const [btn,menu] of [['#fchip','fileMenu'],['#helpBtn','helpMenu']]){
   await p.click(btn);await p.waitForTimeout(250);
   const c=await p.evaluate(COVER,menu);
   ok(`${w}px · ${menu}: כל ${c.n} הפריטים גלויים ונלחצים מעל מסך הפתיחה`,c.n>0&&!c.bad.length,c.bad.join(' · '));
   await p.keyboard.press('Escape');await p.mouse.click(5,h-5);await p.waitForTimeout(150)}
  /* המסלול המלא: תפריט → «דוח ETA» → בורר קבצים → הדוח נטען, עוד בלי ZMRP */
  await p.click('#fchip');await p.waitForTimeout(250);
  let ch=null;try{[ch]=await Promise.all([p.waitForEvent('filechooser',{timeout:4000}),p.click('#fileMenu label[for="fe"]',{timeout:4000})])}catch(_){}
  ok(`${w}px · לחיצה על «דוח ETA» לפני ZMRP פותחת בורר קבצים`,!!ch);
  if(ch){await ch.setFiles(path.join(SD,'zmrp-demo-eta.xlsx'));
   await p.waitForFunction(()=>typeof ETA!=='undefined'&&ETA&&ETA.byPn,null,{timeout:20000}).catch(()=>{});
   const e=await p.evaluate(()=>({n:ETA&&ETA.byPn?Object.keys(ETA.byPn).length:0,chip:(document.getElementById('etachip')||{}).textContent||''}));
   ok(`${w}px · דוח ה-ETA נטען, והצ׳יפ מציג אותו`,e.n>0&&/ETA/.test(e.chip),`${e.n} מק״טים · ${e.chip}`)}
  /* ואחרי שהתפריט נסגר — הסרגל חוזר לשכבה הרגילה, כדי לא לכסות מגירות */
  await p.keyboard.press('Escape');await p.mouse.click(5,h-5);await p.waitForTimeout(200);
  const z=await p.evaluate(()=>{const t=document.querySelector('.top');return getComputedStyle(t).zIndex});
  ok(`${w}px · תפריט סגור → הסרגל בשכבה הרגילה (לא 421)`,z!=='421',`z=${z}`);
  ok(`${w}px · אפס שגיאות JS`,!errs.length,errs.slice(0,2).join(' | '));
  await ctx.close()}
 /* אחרי טעינת ZMRP מסך הפתיחה נעלם, והשכבה חוזרת ל-95 גם כשתפריט פתוח */
 {const ctx=await b.newContext({viewport:{width:1440,height:900}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();await p.goto(HTML+'?nobrief=1');
  await p.setInputFiles('#f',path.join(SD,'zmrp-demo.xlsx'));
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&!LOADJOB,null,{timeout:120000});
  await p.click('#fchip');await p.waitForTimeout(250);
  const r=await p.evaluate(id=>({z:getComputedStyle(document.querySelector('.top')).zIndex,wel:!document.getElementById('welcome').hidden}),'');
  const c=await p.evaluate(COVER,'fileMenu');
  ok('אחרי ZMRP: אין מסך פתיחה, והסרגל עם תפריט פתוח נשאר ב-95 כמו קודם',!r.wel&&r.z==='95',JSON.stringify(r));
  ok('אחרי ZMRP: פריטי תפריט הקבצים נלחצים',c.n>0&&!c.bad.length,c.bad.join(' · '));
  await ctx.close()}
 await b.close();
 console.log(`\n${pass} עברו · ${fail} נכשלו`);process.exit(fail?1:0)})();
