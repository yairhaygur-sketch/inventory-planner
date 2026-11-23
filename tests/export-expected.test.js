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
 ['מק״ט הוחלף במק״ט אחר (קבוצה שונה)',a=>{a[3][0]='T-009'}],
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

/* --- סדר המק״טים: תזוזה מוצהרת במדויק (אושר 11.11) ---
   היה: «סדר המק״טים השתנה → נכשל, ואין דרך להצהיר עליו». עכשיו: נכשל בלי הצהרה,
   ועובר רק כשהפריטים שזזו מוצהרים בדיוק — הרשימה המינימלית, לא כל שורה שהוסטה. */
const mv=(o)=>Object.assign({sheet:'כל הנתונים',pn:'T-003',from:6,to:4},o);
const declM=(moves,changes=[],base)=>({base:(moves.length||changes.length)?(base||SHA):null,changes,moves});
const SW=book(null,a=>{const t=a[3];a[3]=a[4];a[4]=t});                    /* T-001 ↔ T-002 */
const UP=book(null,a=>{const t=a[5];a.splice(5,1);a.splice(3,0,t)});         /* T-003 עולה לראש */
{const r=checkDeclared(D(A,SW),decl([]),SHA);ok('סדר השתנה בלי הצהרה → נכשל (תזוזה לא מוצהרת)',!r.ok&&r.undeclaredMoves.length===1&&!r.other.length,JSON.stringify(r.undeclaredMoves))}
{const d=D(A,UP);ok('פריט אחד עלה שתי שורות → מוצהר פריט אחד בלבד, לא השורות שהוסטו',d.moves.length===1&&d.moves[0].pn==='T-003'&&d.moves[0].from===6&&d.moves[0].to===4,JSON.stringify(d.moves))}
{const r=checkDeclared(D(A,UP),declM([mv()]),SHA);ok('התזוזה מוצהרת בדיוק → עובר',r.ok&&r.movesMatched===1)}
for(const [n,o] of [['שורת יעד אחרת',{to:5}],['שורת מקור אחרת',{from:5}],['מק״ט אחר',{pn:'T-002'}],['גיליון אחר',{sheet:'התצוגה'}]]){
 const r=checkDeclared(D(A,UP),declM([mv(o)]),SHA);ok(`תזוזה מוצהרת עם ${n} → נכשל`,!r.ok&&r.undeclaredMoves.length===1&&r.unusedMoves.length===1)}
{const r=checkDeclared(D(A,book()),declM([mv()]),SHA);ok('תזוזה מוצהרת שלא קרתה → נכשל',!r.ok&&r.unusedMoves.length===1)}
{const B=book(null,a=>{const t=a[5];a.splice(5,1);a.splice(3,0,t);a[3][2]=5});   /* עלה, ו-ROP שלו השתנה */
 const r=checkDeclared(D(A,B),declM([mv()],[chg({pn:'T-003',field:'ROP מוצע',from:0,to:5})]),SHA);
 ok('שינוי תא בשורה שזזה נבדק מול אותו מק״ט → עובר כשמוצהר',r.ok&&r.matched===1&&r.movesMatched===1,JSON.stringify(r.undeclared))}
{const B=book(null,a=>{const t=a[5];a.splice(5,1);a.splice(3,0,t);a[3][2]=5});
 const r=checkDeclared(D(A,B),declM([mv()]),SHA);ok('…ונכשל כשהשינוי בתא לא מוצהר',!r.ok&&r.undeclared.length===1&&r.undeclared[0].pn==='T-003')}
{const dA=book([['מק״ט','ROP'],['T-1',1],['T-1',2],['T-2',3]]),dB=book([['מק״ט','ROP'],['T-2',3],['T-1',1],['T-1',2]]);
 const r=checkDeclared(D(dA,dB),decl([]),SHA);ok('מק״ט כפול בגיליון → סדר אינו ניתן להצהרה, נכשל',!r.ok&&r.other.length>0&&!(r.undeclaredMoves||[]).length)}
for(const [n,m] of [['from לא שלם',{from:4.5}],['to = from',{to:6}],['מק״ט «*»',{pn:'*'}],['שדה חסר',{to:undefined}]]){
 const x=mv(m);if(x.to===undefined)delete x.to;const r=checkDeclared(D(A,UP),declM([x]),SHA);ok(`תזוזה עם ${n} → נדחתה`,!r.ok&&r.errs.length>0)}
{const r=checkDeclared(D(A,UP),declM([mv(),mv()]),SHA);ok('אותה תזוזה פעמיים → נדחתה',!r.ok&&r.errs.some(e=>/פעמיים/.test(e)))}
{const r=checkDeclared(D(A,UP),{base:null,changes:[],moves:[mv()]},SHA);ok('תזוזות בלי base → נדחה',!r.ok&&r.errs.some(e=>/40/.test(e)))}
{const r=checkDeclared(D(A,book()),{base:null,changes:[],moves:[]},SHA);ok('moves ריק → כמו הצהרה ריקה, עובר כשאין הבדל',r.ok)}

/* ============ XSTRUCT — שינוי מבנה מוצהר (אושר 23.11, רפורמת האקסל) ============
   ספר ישן: «כל הנתונים» (מק״ט · תיאור · ROP מוצע · ביקוש בזמן האספקה) + «מכירות».
   ספר חדש: «סיכום» (תוכן חדש) + «נתונים עיקריים» (מק״ט · תיאור · נקודת הזמנה מוצעת)
   + «פירוט מלא» (מק״ט · ביקוש בזמן האספקה · עמודה חדשה). «מכירות» הוסר. */
const xbook=(sheets)=>{const wb=XLSX.utils.book_new();for(const [n,aoa] of sheets)XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),n);return wb};
const XOLD=()=>xbook([['כל הנתונים',[['ייצוא'],[],['מק״ט','תיאור','ROP מוצע','ביקוש בזמן האספקה'],['T-001','א',7,0.67],['T-002','ב',3,1.5],['T-003','ג',0,0]]],
 ['מכירות',[['מק״ט','תאריך'],['T-001','1.1']]]]);
const XNEW=(mut)=>{const main=[['כותרת אחרת — לא נבדקת'],['מק״ט','תיאור','נקודת הזמנה מוצעת'],['T-001','א',7],['T-002','ב',3],['T-003','ג',0]];
 const full=[['מק״ט','ביקוש בזמן האספקה','חדשה'],['T-001',0.67,'x'],['T-002',1.5,'y'],['T-003',0,'z']];
 if(mut)mut(main,full);return xbook([['סיכום',[['סה״כ',3]]],['נתונים עיקריים',main],['פירוט מלא',full]])};
const XS=(o)=>Object.assign({sheets:['סיכום','נתונים עיקריים','פירוט מלא'],
 map:{'סיכום':{from:null},
  'נתונים עיקריים':{from:'כל הנתונים',columns:{'מק״ט':'מק״ט','תיאור':'תיאור','נקודת הזמנה מוצעת':'ROP מוצע'}},
  'פירוט מלא':{from:'כל הנתונים',columns:{'מק״ט':'מק״ט','ביקוש בזמן האספקה':'ביקוש בזמן האספקה','חדשה':null}}},
 droppedSheets:['מכירות']},o||{});
const XD=(S,changes=[],base)=>({base:base===undefined?SHA:base,changes,structure:S});
const XR=(B,S,changes,base,sha)=>checkDeclared(cellDiffs(XOLD(),B,S),XD(S,changes,base),sha===undefined?SHA:sha);
{const r=XR(XNEW(),XS());ok('XSTRUCT · מבנה מוצהר בדיוק, אותם ערכים → עובר',r.ok,JSON.stringify([r.errs,r.other]))}
{const r=XR(XNEW((m)=>{m[3][2]=4}),XS());ok('XSTRUCT · ערך שזז בעמודה ממופה בלי הצהרה → נכשל (הערכים עדיין נבדקים)',!r.ok&&r.undeclared.length===1&&r.undeclared[0].field==='נקודת הזמנה מוצעת'&&r.undeclared[0].pn==='T-002')}
{const r=XR(XNEW((m)=>{m[3][2]=4}),XS(),[{sheet:'נתונים עיקריים',field:'נקודת הזמנה מוצעת',pn:'T-002',from:3,to:4}]);ok('XSTRUCT · …ועובר כשמוצהר בשמות החדשים',r.ok&&r.matched===1)}
{const r=XR(XNEW((m,f)=>{f[2][1]=9}),XS());ok('XSTRUCT · גם בגיליון השני אותו מק״ט נבדק',!r.ok&&r.undeclared.length===1&&r.undeclared[0].sheet==='פירוט מלא')}
{const S=XS();delete S.map['פירוט מלא'].columns['חדשה'];const r=XR(XNEW(),S);ok('XSTRUCT · עמודה חדשה שאינה במיפוי → נכשל',!r.ok&&r.other.some(e=>/אינן מוצהרות/.test(e)))}
{const S=XS();S.map['נתונים עיקריים'].columns['רפאים']=null;const r=XR(XNEW(),S);ok('XSTRUCT · עמודה מוצהרת שאינה בקובץ → נכשל',!r.ok&&r.other.some(e=>/אינן בקובץ/.test(e)))}
{const S=XS();S.map['פירוט מלא'].columns['ביקוש בזמן האספקה']=null;const r=XR(XNEW(),S);ok('XSTRUCT · עמודה ישנה שנעלמה בלי dropped → נכשל',!r.ok&&r.other.some(e=>/נעלמו בלי הצהרה/.test(e)))}
{const S=XS({dropped:{'כל הנתונים':['ביקוש בזמן האספקה']}});S.map['פירוט מלא'].columns['ביקוש בזמן האספקה']=null;const r=XR(XNEW(),S);ok('XSTRUCT · …ועובר כשהיא מוצהרת ב-dropped',r.ok,JSON.stringify(r.other))}
{const S=XS({dropped:{'כל הנתונים':['תיאור']}});const r=XR(XNEW(),S);ok('XSTRUCT · dropped על עמודה שממשיכה → נכשל',!r.ok&&r.other.some(e=>/אינן נעלמות/.test(e)))}
{const S=XS();delete S.droppedSheets;const r=XR(XNEW(),S);ok('XSTRUCT · גיליון ישן שנעלם בלי droppedSheets → נכשל',!r.ok&&r.other.some(e=>/נעלם בלי הצהרה/.test(e)))}
{const S=XS({sheets:['נתונים עיקריים','סיכום','פירוט מלא']});const r=XR(XNEW(),S);ok('XSTRUCT · סדר הגיליונות שונה מהמוצהר → נכשל',!r.ok&&r.other.some(e=>/גיליונות/.test(e)))}
{const B=XNEW((m)=>{const t=m[2];m[2]=m[3];m[3]=t});const r=XR(B,XS());ok('XSTRUCT · סדר שורות השתנה בלי reorder → נכשל',!r.ok&&r.other.some(e=>/סדר השורות/.test(e)))}
{const B=XNEW((m)=>{const t=m[2];m[2]=m[3];m[3]=t});const S=XS();S.map['נתונים עיקריים'].reorder=true;const r=XR(B,S);ok('XSTRUCT · …ועובר עם reorder: true — וערכים עדיין לפי מק״ט',r.ok,JSON.stringify(r.other))}
{const r=XR(XNEW((m)=>{m.pop()}),XS());ok('XSTRUCT · קבוצת המק״טים שונה → נכשל',!r.ok&&r.other.some(e=>/קבוצת המק״טים/.test(e)))}
{const r=XR(XNEW(),XS(),[],null);ok('XSTRUCT · מבנה בלי base → נדחה',!r.ok&&r.errs.some(e=>/40/.test(e)))}
{const r=XR(XNEW(),XS(),[],SHA2);ok('XSTRUCT · מבנה לבסיס אחר → נכשל, עם הוראה לאפס',!r.ok&&r.errs.some(e=>/אפס/.test(e)))}
{const S=XS();S.map['נתונים עיקריים'].columns['תיאור']='*';const r=XR(XNEW(),S);ok('XSTRUCT · «*» במיפוי → נדחה',!r.ok&&r.errs.length>0)}
{const S=XS({extra:1});const r=XR(XNEW(),S);ok('XSTRUCT · מפתח לא מוכר במבנה → נדחה',!r.ok&&r.errs.some(e=>/לא מוכרים/.test(e)))}
{const S=XS();delete S.map['פירוט מלא'];const r=XR(XNEW(),S);ok('XSTRUCT · גיליון חדש שאינו ב-map → נכשל',!r.ok&&r.other.some(e=>/אינו מוצהר/.test(e)))}
/* בלי structure — ההשוואה הרגילה, ושינוי מבנה אינו ניתן להצהרה בכלל (כמו קודם) */
{const r=checkDeclared(cellDiffs(XOLD(),XNEW()),decl([]),SHA);ok('XSTRUCT · בלי structure — שינוי מבנה נכשל כמו קודם',!r.ok&&r.other.length>0)}

/* --- הקובץ שבריפו --- */
const repo=require('./export-expected.json');
ok('tests/export-expected.json תקין',!validateDecl(repo).length,validateDecl(repo).join(' · ')||`${repo.changes.length} שינויים מוצהרים`);

console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
process.exit(bad?1:0);
