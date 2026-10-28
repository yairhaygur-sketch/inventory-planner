/* ============ PR-A · בחירת הבסיס להשוואת הייצוא — בלי דפדפן ============
   export-worker.test משווה את קובץ האקסל ל-main רק אם resolveBase()
   אומר compare. כאן נבדק כל ענף של ההחלטה, עם קבצים ו-git מדומים:
   כל מצב שבו ההשוואה נדרשת ואי אפשר לבצע אותה חייב לצאת error — לא
   skip — כי skip הוא בדיוק איך ההשוואה נעלמה ב-CI עד עכשיו. */
const {resolveBase}=require('./lib/export-base.js');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
const CUR=Buffer.from('<html>pr</html>'),OTHER=Buffer.from('<html>main</html>'),EMPTY=Buffer.alloc(0);
/* מערכת קבצים ו-git מדומים. gitShow/gitRev זורקים כברירת מחדל, כמו ב-CI */
const io=(files={},git=null)=>({
 read:f=>{if(f==='/idx')return CUR;if(!(f in files))throw new Error('ENOENT '+f);return files[f]},
 exists:f=>f in files,
 gitShow:()=>{if(!git)throw Object.assign(new Error('fatal'),{stderr:"fatal: invalid object name 'origin/main'."});return git},
 gitRev:()=>{if(!git)throw new Error('no rev');return 'abc1234def'}});
const R=(env,files,git)=>resolveBase(env,io(files,git),'/idx');
const C=[
 /* [שם, env, קבצים, git, kind צפוי, ביטוי בסיבה] */
 ['required · קובץ בסיס חסר → error, לא skip',{EXPORT_BASE:'required',EXPORT_BASE_SHA:'s1',EXPORT_BASE_HTML:'/b'},{},null,'error',/חסרה/],
 ['required · קובץ בסיס ריק → error',{EXPORT_BASE:'required',EXPORT_BASE_SHA:'s1',EXPORT_BASE_HTML:'/b'},{'/b':EMPTY},null,'error',/ריקה/],
 ['required · בלי EXPORT_BASE_HTML → error',{EXPORT_BASE:'required',EXPORT_BASE_SHA:'s1'},{},null,'error',/EXPORT_BASE_HTML/],
 ['required · בלי SHA → error',{EXPORT_BASE:'required',EXPORT_BASE_HTML:'/b'},{'/b':OTHER},null,'error',/SHA/],
 ['required · זהה ל-index.html → skip «זהה»',{EXPORT_BASE:'required',EXPORT_BASE_SHA:'s1',EXPORT_BASE_HTML:'/b'},{'/b':CUR},null,'skip',/זהה/],
 ['required · שונה → compare, עם ה-SHA',{EXPORT_BASE:'required',EXPORT_BASE_SHA:'s1',EXPORT_BASE_HTML:'/b'},{'/b':OTHER},null,'compare',null],
 ['push → skip «push»',{EXPORT_BASE:'push',CI:'true'},{},null,'skip',/push/],
 ['push · גם כשיש קובץ בסיס → skip (לא נפתח)',{EXPORT_BASE:'push',EXPORT_BASE_HTML:'/b'},{'/b':OTHER},null,'skip',/push/],
 ['CI=true בלי EXPORT_BASE (השלב הוסר) → error',{CI:'true'},{},OTHER,'error',/הוסר/],
 ['EXPORT_BASE לא מוכר → error',{EXPORT_BASE:'yes'},{},null,'error',/לא מוכר/],
 ['מקומי · בלי origin/main → skip «מקומי» עם שגיאת git',{},{},null,'skip',/מקומית.*origin\/main/],
 ['מקומי · origin/main שונה → compare',{},{},OTHER,'compare',null],
 ['מקומי · origin/main זהה → skip «זהה»',{},{},CUR,'skip',/זהה/],
 ['מקומי · origin/main ריק → error',{},{},EMPTY,'error',/ריק/],
 ['מקומי · EXPORT_BASE_HTML נתון → compare מולו',{EXPORT_BASE_HTML:'/b'},{'/b':OTHER},null,'compare',null],
 ['מקומי · EXPORT_BASE_HTML חסר בדיסק → error',{EXPORT_BASE_HTML:'/b'},{},null,'error',/חסרה/]];
for(const [n,env,files,git,kind,re] of C){
 let r;try{r=R(env,files,git)}catch(e){r={kind:'throw',reason:e.message}}
 ok(n,r.kind===kind&&(!re||re.test(r.reason||'')),`${r.kind}${r.reason?' · '+r.reason:''}`)}
/* compare מחזיר את הבסיס עצמו ואת ה-SHA — כדי שהבדיקה תשווה מולו ותדפיס מול מה */
{const r=R({EXPORT_BASE:'required',EXPORT_BASE_SHA:'s1',EXPORT_BASE_HTML:'/b'},{'/b':OTHER});
 ok('compare מחזיר את קובץ הבסיס ואת ה-SHA',r.html&&r.html.equals(OTHER)&&r.sha==='s1')}
/* אף ענף אינו מחזיר skip כשההשוואה נדרשת והבסיס תקול */
{const kinds=['/b missing','/b empty'].map((_,i)=>R({EXPORT_BASE:'required',EXPORT_BASE_SHA:'s',EXPORT_BASE_HTML:'/b'},i?{'/b':EMPTY}:{}).kind);
 ok('required עם בסיס תקול לעולם אינו skip',kinds.every(k=>k==='error'),kinds.join(','))}
console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
process.exit(bad?1:0);
