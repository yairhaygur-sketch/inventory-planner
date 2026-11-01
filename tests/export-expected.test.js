/* ============ PR-B · ההצהרה על שינוי בייצוא — בלי דפדפן ============
   tests/lib/export-expected.js מחליט אם ההבדל בין הייצוא של ה-PR לזה של
   main מוצהר בדיוק. כאן נבדק כל ענף על ספרים סינתטיים קטנים: הבדל לא
   מוצהר, הצהרה שלא קרתה, ערך או סוג שונים, בסיס אחר, תבנית, ומה שאי
   אפשר להצהיר עליו בכלל. ברירת המחדל — אפס הבדלים. */
const XLSX=require('xlsx');
const {cellDiffs,validateDecl,checkDeclared}=require('./lib/export-expected.js');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
const SHA='a'.repeat(40),SHA2='b'.repeat(40);
/* ספר בסיס: כותרת, שורת כותרות עם מק״ט, שלושה פריטים סינתטיים */
const AOA=()=>[['ייצוא בדיקה'],[],['מק״ט','תיאור','ROP מוצע','ביקוש בזמן האספקה'],
 ['T-001','א',7,0.67],['T-002','ב',3,1.5],['T-003','ג',0,0]];
const book=(aoa,mut)=>{const a=aoa||AOA();if(mut)mut(a);const wb=XLSX.utils.book_new();
 const ws=XLSX.utils.aoa_to_sheet(a);ws['!cols']=[{wch:12},{wch:8},{wch:8},{wch:8}];XLSX.utils.book_append_sheet(wb,ws,'כל הנתונים');return wb};
const D=(A,B)=>cellDiffs(A,B);
const chg=(o)=>Object.assign({sheet:'כל הנתונים',field:'ביקוש בזמן האספקה',pn:'T-001',from:0.67,to:0.08},o);
const decl=(changes,base)=>({base:changes.length?(base||SHA):null,changes});
const A=book();
const B1=book(null,a=>{a[3][3]=0.08});                         /* תא אחד זז */
const B2=book(null,a=>{a[3][3]=0.08;a[4][2]=4});               /* שניים */

/* --- ברירת מחדל: אפס הבדלים --- */
{const r=checkDeclared(D(A,book()),decl([]),SHA);ok('ספרים זהים + הצהרה ריקה → עובר',r.ok)}
{const r=checkDeclared(D(A,B1),decl([]),SHA);ok('תא אחד שזז + הצהרה ריקה → נכשל (לא מוצהר)',!r.ok&&r.undeclared.length===1,r.undeclared.map(d=>d.pn).join())}
{const r=checkDeclared(D(A,book()),decl([]),'');ok('הצהרה ריקה אינה דורשת SHA — השוואה מקומית עוברת כשאין הבדל',r.ok)}

/* --- הצהרה מדויקת --- */
{const r=checkDeclared(D(A,B1),decl([chg()]),SHA);ok('ההבדל מוצהר בדיוק → עובר',r.ok&&r.matched===1)}
{const r=checkDeclared(D(A,B2),decl([chg()]),SHA);ok('שני הבדלים, אחד מוצהר → נכשל על השני',!r.ok&&r.undeclared.length===1&&r.undeclared[0].field==='ROP מוצע')}
{const r=checkDeclared(D(A,B1),decl([chg(),chg({pn:'T-002',field:'ROP מוצע',from:3,to:4})]),SHA);
 ok('הצהרה שלא קרתה → נכשל',!r.ok&&r.unused.length===1&&r.unused[0].pn==='T-002')}
{const r=checkDeclared(D(A,book()),decl([chg()]),SHA);ok('הצהרה כשאין שום הבדל → נכשל',!r.ok&&r.unused.length===1)}

/* --- כל רכיב בהצהרה חייב להתאים --- */
for(const [n,o] of [['ערך חדש אחר',{to:0.09}],['ערך קודם אחר',{from:0.66}],['מק״ט אחר',{pn:'T-003'}],
  ['עמודה אחרת',{field:'ROP מוצע'}],['גיליון אחר',{sheet:'התצוגה'}],['אותו ערך כטקסט ("0.08")',{to:'0.08'}]]){
 const r=checkDeclared(D(A,B1),decl([chg(o)]),SHA);
 ok(`הצהרה עם ${n} → נכשל (גם ההבדל וגם ההצהרה)`,!r.ok&&r.undeclared.length===1&&r.unused.length===1)}

/* --- הבסיס --- */
{const r=checkDeclared(D(A,B1),decl([chg()],SHA2),SHA);ok('הצהרה לבסיס אחר → נכשל, עם הוראה לאפס',!r.ok&&r.errs.some(e=>/אפס/.test(e)))}
{const r=checkDeclared(D(A,B1),decl([chg()]),'');ok('הצהרה לא ריקה מול קובץ בלי SHA → נכשל',!r.ok&&r.errs.length===1)}
{const r=checkDeclared(D(A,B1),{base:'abc1234',changes:[chg()]},'abc1234');ok('base מקוצר → נכשל (חייב SHA מלא)',!r.ok&&r.errs.some(e=>/40/.test(e)))}
{const r=checkDeclared(D(A,book()),{base:SHA,changes:[]},SHA);ok('base עם הצהרה ריקה → נכשל (הצהרה ריקה אינה קשורה לבסיס)',!r.ok)}

/* --- אין תבניות ואין שדות חסרים --- */
for(const [n,o] of [['מק״ט «*»',{pn:'*'}],['מק״ט עם «*»',{pn:'T-*'}],['עמודה «*»',{field:'*'}],['גיליון «*»',{sheet:'*'}],['ערך חדש «*»',{to:'*'}]]){
 const r=checkDeclared(D(A,B1),decl([chg(o)]),SHA);ok(`${n} → נדחה`,!r.ok&&r.errs.length>0)}
{const c=chg();delete c.from;const r=checkDeclared(D(A,B1),decl([c]),SHA);ok('שדה חסר (from) → נדחה',!r.ok&&r.errs.some(e=>/חסר/.test(e)))}
{const r=checkDeclared(D(A,B1),decl([chg(),chg()]),SHA);ok('הצהרה כפולה → נדחתה',!r.ok&&r.errs.some(e=>/כפול/.test(e)))}
{const r=checkDeclared(D(A,B1),{base:SHA,changes:[chg()],skip:true},SHA);ok('מפתח לא מוכר (למשל skip) → נדחה',!r.ok&&r.errs.some(e=>/לא מוכרים/.test(e)))}
{const r=checkDeclared(D(A,B1),decl([chg({from:0.08})]),SHA);ok('from = to → נדחה',!r.ok)}

/* --- מה שאי אפשר להצהיר עליו --- */
const cases=[
 ['שורה נוספה',a=>{a.push(['T-004','ד',1,1])}],
 ['סדר המק״טים השתנה',a=>{const t=a[3];a[3]=a[4];a[4]=t}],
 ['כותרת עמודה השתנתה',a=>{a[2][2]='ROP'}],
 ['שורת הכותרת העליונה השתנתה',a=>{a[0][0]='ייצוא אחר'}],
 ['תא בלי מק״ט',a=>{a.push([null,'סה״כ',10,2]);}],
];
for(const [n,mut] of cases){
 const base=n==='תא בלי מק״ט'?book(null,a=>{a.push([null,'סה״כ',9,2])}):A;
 const r=checkDeclared(D(base,book(null,mut)),decl([]),SHA);
 ok(`${n} → נכשל, ואין דרך להצהיר עליו`,!r.ok&&r.other.length>0,r.other[0])}
{const B=book();B.Sheets['כל הנתונים']['!cols'][1]={wch:30};const r=checkDeclared(D(A,B),decl([]),SHA);ok('רוחב עמודה השתנה → נכשל',!r.ok&&r.other.length>0)}
{const B=book();const ws=B.Sheets['כל הנתונים'];ws.C4={t:'s',v:7};const r=checkDeclared(D(A,B),decl([]),SHA);
 ok('מספר שהפך לטקסט → נכשל',!r.ok)}
{const B=XLSX.utils.book_new();XLSX.utils.book_append_sheet(B,XLSX.utils.aoa_to_sheet(AOA()),'אחר');
 const r=checkDeclared(D(A,B),decl([]),SHA);ok('שם גיליון השתנה → נכשל',!r.ok&&r.other.length>0)}
{const dupA=book([['מק״ט','ROP','ROP'],['T-1',1,2]]),dupB=book([['מק״ט','ROP','ROP'],['T-1',1,3]]);
 const r=checkDeclared(D(dupA,dupB),decl([chg({sheet:'כל הנתונים',field:'ROP',pn:'T-1',from:2,to:3})]),SHA);
 ok('עמודה ששמה כפול → אי אפשר להצהיר, נכשל',!r.ok&&r.other.length>0)}

/* --- הקובץ שבריפו --- */
const repo=require('./export-expected.json');
ok('tests/export-expected.json תקין',!validateDecl(repo).length,validateDecl(repo).join(' · ')||`${repo.changes.length} שינויים מוצהרים`);

console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
process.exit(bad?1:0);
