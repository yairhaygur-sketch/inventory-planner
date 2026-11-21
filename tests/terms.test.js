/* ============ TERMS — מונחי רכש ותכנון חומרים, וסידור המסילה (אושר 21.11) ============
   ההחלטה (docs/BUSINESS-DECISIONS.md, docs/DECISION-RULES.md 5ת5):
   · שמונה תחומים, לפי הפעולה שבסוף: חריגים לטיפול · חוסרים והזמנות לקוח לא
     מכוסות · מעקב הזמנות רכש · פרמטרי תכנון (MRP) · מלאי עודף, איטי תנועה ומת ·
     נתוני אב חסרים · קטלוג פריטים · בקרת יישום.
   · במסילה שם קצר (המונה חייב להיראות ב-1366); בכותרת המסך ובבועית — השם המלא.
   · «ROP קצר/נמוך» עברו מרכש לפרמטרי תכנון; «אספקות שתאריכן עבר» עברו לרכש;
     «ביקוש עולה» הצטרף ל«ביקוש יורד» ב«בקרה תקופתית», יחד עם מלאי הביטחון
     המינימלי; «טופלו» ו«יושמו ב-SAP» — ב«בקרת יישום».
   · «קטלוג פריטים» בלי מונה ובלי אדום.
   · שמות ומיקום בלבד — הספירות לא זזות (navCount זהה לכל מפתח).
   הבדיקה טוענת את הפיקסצ׳ר, קוראת את המסילה מה-DOM ובודקת כל דלת. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML=process.argv[2]||path.join(SD,'..','index.html');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null?'  ['+x+']':''))};
const AREAS=[['home','חריגים לטיפול','חריגים לטיפול'],['today','חוסרים והזמנות לקוח לא מכוסות','חוסרים ולקוחות'],
 ['line','מעקב הזמנות רכש','מעקב רכש'],['month','פרמטרי תכנון (MRP)','פרמטרי תכנון'],
 ['cap','מלאי עודף, איטי תנועה ומת','מלאי עודף ומת'],['qual','נתוני אב חסרים','נתוני אב חסרים'],
 ['catalog','קטלוג פריטים','קטלוג פריטים'],['applied','בקרת יישום','בקרת יישום']];
/* שלב 2 (21.11): «זירוז ספקים» במעקב רכש; דלתות ה-ROP ו«מעל מכירות 3 שנים» — סינון ברשימה */
const SUBS={today:['burn','cust'],line:['line','follow','stale'],
 month:['month','tier','floor','floorNR','rise','trend','first','prof'],
 cap:['cap','capPO','capStop','capBiz','capNone'],catalog:['catalog','moves'],applied:['applied','done']};
/* מה שהיה לפני TERMS — אסור שיופיע במסילה או בכותרת */
const OLD=['מרכז עבודה','רכש ואספקות','תכנון ו-MRP','איכות נתונים','בריאות המלאי','כסף תקוע',
 'מספר לתקן ב-SAP','נתון שבור ב-SAP','מינימום לפי מחיר','מלאי ביטחון ידני','מינימום בלי דרישה',
 'מתי מה אמור להגיע','רכש בדרך לעודף','החלטה מסחרית','אין מה לשנות','אין רכש כלל','שינוי פרופיל MRP'];
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1366,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+HTML+'?nobrief=1');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});
 const rail=()=>p.evaluate(()=>[...document.querySelectorAll('#tabs .navtab')].map(b=>{const t=b.querySelector('.t'),g=b.querySelector('.bdg');
  const rr=document.querySelector('.rail').getBoundingClientRect(),gr=g&&g.getBoundingClientRect();
  return {k:b.dataset.nav,t:t.textContent,title:b.title,cut:t.scrollWidth>t.clientWidth+1,bdg:g?g.textContent:null,
   bdgIn:!g||(gr.left>=rr.left-1&&gr.right<=rr.right+1&&gr.width>0),alert:b.classList.contains('alert')}}));
 const R0=await rail();
 ok('שמונה תחומים, בסדר שאושר',JSON.stringify(R0.map(x=>x.k))===JSON.stringify(AREAS.map(a=>a[0])),R0.map(x=>x.k).join(','));
 ok('השם הקצר במסילה',AREAS.every(([k,,s])=>(R0.find(x=>x.k===k)||{}).t===s),R0.map(x=>x.t).join(' | '));
 ok('השם המלא בבועית',AREAS.every(([k,f])=>((R0.find(x=>x.k===k)||{}).title||'').startsWith(f+' — ')),'');
 ok('1366: אף שם תחום אינו נחתך',!R0.some(x=>x.cut),R0.filter(x=>x.cut).map(x=>x.t).join(','));
 ok('1366: כל מונה בתוך המסילה',R0.every(x=>x.bdgIn),R0.filter(x=>!x.bdgIn).map(x=>x.k).join(','));
 const cat=R0.find(x=>x.k==='catalog')||{};
 ok('«קטלוג פריטים» — בלי מונה ובלי אדום',cat.bdg===null&&!cat.alert,JSON.stringify(cat));
 ok('תחום עם מונה אפס אינו נעלם (בקרת יישום)',!!R0.find(x=>x.k==='applied'),'');
 /* כל תחום: דלתות המשנה, כותרת המסך עם השם המלא, ואין דלת שנחתכת */
 for(const [area,keys] of Object.entries(SUBS)){
  const S=await p.evaluate(a=>{setMode(a);return {subs:[...document.querySelectorAll('#railsub .railsub2')].map(b=>{const t=b.querySelector('.t');
   return {k:b.dataset.nav,cut:t.scrollWidth>t.clientWidth+1,t:t.textContent}}),lt:document.getElementById('lt').textContent,
   sect:(document.querySelector('#railsub .railsect')||{}).textContent||''}},area);
  ok(`${area}: דלתות המשנה`,JSON.stringify(S.subs.map(x=>x.k))===JSON.stringify(keys),S.subs.map(x=>x.k).join(','));
  ok(`${area}: אף דלת משנה אינה נחתכת ב-1366`,!S.subs.some(x=>x.cut),S.subs.filter(x=>x.cut).map(x=>x.t).join(','));
  ok(`${area}: כותרת המסך = השם המלא`,S.lt===AREAS.find(a=>a[0]===area)[1],S.lt);
 }
 const T=await p.evaluate(()=>{const o={};
  for(const [k,full] of [['tier','חלק זול ללא נקודת הזמנה (ROP)'],['floor','מלאי ביטחון מינימלי (SS) מעל הצורך'],['stale','הזמנות רכש באיחור'],['done','טופלו']]){
   setMode(k);o[k]=[document.getElementById('lt').textContent,full]}
  /* «בקרה תקופתית»: הקבוצה בתוך פרמטרי תכנון, ואחריה כל רשימות העיון */
  setMode('month');const seq=[...document.querySelectorAll('#railsub > *')].map(e=>e.classList.contains('railgrp')?'@':e.dataset.nav||'');
  o.grp=seq.slice(seq.indexOf('@')+1);
  o.txt=document.querySelector('.rail').textContent+' '+document.getElementById('lt').textContent;
  o.counts=['home','today','line','month','qual','cap','burn','cust','follow','tier','floor','floorNR','trend','rise','first','prof','applied','done','stale','capPO','capStop','capBiz','capNone']
   .map(k=>[k,navCount(k).n]);
  return o});
 ok('כותרת דלת = התחום · השם המלא של הדלת',['tier','floor'].every(k=>T[k][0].endsWith(' · '+T[k][1])),JSON.stringify([T.tier,T.floor]));
 ok('«הזמנות רכש באיחור» תחת מעקב הזמנות רכש',T.stale[0]==='מעקב הזמנות רכש · הזמנות רכש באיחור',T.stale[0]);
 ok('«טופלו» תחת בקרת יישום',T.done[0]==='בקרת יישום · טופלו',T.done[0]);
 ok('בקרה תקופתית: מלאי ביטחון מינימלי, מגמות, צריכה ראשונה, סוג MRP',JSON.stringify(T.grp)===JSON.stringify(['floor','floorNR','rise','trend','first','prof']),T.grp.join(','));
 ok('אף שם ישן במסילה או בכותרת',!OLD.some(w=>T.txt.includes(w)),OLD.filter(w=>T.txt.includes(w)).join(','));
 /* הספירות: הבדיקה משווה ל-navCount עצמו בכל דלת — כלומר המסילה מציגה בדיוק
    את מה שהמנוע סופר. ההוכחה שהמספרים לא זזו מול main היא business-guard. */
 const shown=await p.evaluate(()=>{const o={};for(const a of ['home','today','line','month','cap','qual','applied']){setMode(a);
  for(const b of document.querySelectorAll('#tabs .navtab .bdg, #railsub .railsub2 .bdg')){const k=b.closest('[data-nav]').dataset.nav;o[k]=+b.dataset.count}}return o});
 const mism=T.counts.filter(([k,n])=>shown[k]!==undefined&&shown[k]!==n);
 ok('כל מונה במסילה = navCount של אותה דלת',!mism.length&&Object.keys(shown).length>=20,mism.map(x=>x.join('=')).join(',')||Object.keys(shown).length);
 /* 21.11: הדלת בעודף אינה «ללא פעולה» — זה שם סיווג של המנוע שמופיע באותה טבלה */
 const cn=await p.evaluate(()=>{setMode('cap');const e=document.querySelector('#railsub [data-nav="capNone"] .t');return e?e.textContent:''});
 ok('«אין פרמטר לשנות» — ולא «ללא פעולה», שהוא שם סיווג',cn==='אין פרמטר לשנות',cn);
 ok('0 שגיאות JS',!errs.length,errs.join('|'));
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})();
