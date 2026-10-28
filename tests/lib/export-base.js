/* ============ PR-A · מול איזה main משווים את ייצוא האקסל ============
   עד כאן export-worker.test ניסה `git show origin/main:index.html`, וכל
   כשל — כולל זה שב-CI תמיד קרה, כי ה-checkout אינו מוריד את main —
   נבלע ב-catch ונרשם «השוואה דולגה» כ-PASS. כלומר ב-CI ההשוואה לא
   יכלה להיכשל אף פעם: לא על תוכן, לא על גודל, ולא כשהדף של main קרס.

   הפונקציה הזו מחליטה *אם* משווים ומול מה, ומחזירה אחת משלוש:
     compare — יש בסיס, והוא שונה מ-index.html. ההשוואה חייבת לרוץ.
     skip    — רק בשלושה מקרים מוצדקים, עם סיבה שמודפסת:
               index.html זהה לבסיס · ריצת push · הרצה מקומית בלי main.
     error   — ההשוואה נדרשת ואי אפשר לבצע אותה. זה FAIL, לא דילוג.

   EXPORT_BASE נקבע ב-workflow: `required` ב-pull_request, `push` בכל
   השאר. ב-required שלב «הכנת גרסת הבסיס» כותב את index.html של ה-main
   שאליו ה-PR מוזג (HEAD^1 של ה-merge commit) ל-EXPORT_BASE_HTML, ואת
   ה-SHA שלו ל-EXPORT_BASE_SHA. CI=true בלי EXPORT_BASE פירושו שמישהו
   הוריד את השלב מה-workflow — וזה error, כדי שההגנה לא תיעלם בשקט.

   מקומית אפשר לתת EXPORT_BASE_HTML=<קובץ> (תרחישי ההוכחה), ובלעדיו
   הבסיס הוא origin/main. הקלט מוזרק (env, קריאת קבצים, git), ולכן
   export-base.test.js בודק כל ענף בלי דפדפן ובלי git אמיתי. */
const fs=require('fs'),path=require('path');
const ROOT=path.join(__dirname,'..','..');
const realIO={
 read:f=>fs.readFileSync(f),
 exists:f=>fs.existsSync(f),
 gitShow:()=>require('child_process').execSync('git show origin/main:index.html',
  {cwd:ROOT,stdio:['ignore','pipe','pipe'],maxBuffer:64e6}),
 gitRev:()=>String(require('child_process').execSync('git rev-parse origin/main',
  {cwd:ROOT,stdio:['ignore','pipe','pipe']})).trim()};
const short=s=>String(s||'').slice(0,7);
function resolveBase(env=process.env,io=realIO,indexPath=path.join(ROOT,'index.html')){
 const mode=env.EXPORT_BASE||'';
 const cur=io.read(indexPath);
 const fromFile=(f,sha,src)=>{
  if(!f)return {kind:'error',reason:'EXPORT_BASE_HTML חסר — שלב «הכנת גרסת הבסיס» לא כתב קובץ'};
  if(!io.exists(f))return {kind:'error',reason:`גרסת הבסיס חסרה: ${f} לא קיים`};
  const html=io.read(f);
  if(!html.length)return {kind:'error',reason:`גרסת הבסיס ריקה: ${f}`};
  if(html.equals(cur))return {kind:'skip',reason:`index.html זהה לבסיס ${short(sha)||src} — אין מה להשוות`};
  return {kind:'compare',html,sha,src}};
 if(mode==='required'){
  if(!env.EXPORT_BASE_SHA)return {kind:'error',reason:'EXPORT_BASE_SHA חסר — לא ידוע מול איזה main משווים'};
  return fromFile(env.EXPORT_BASE_HTML,env.EXPORT_BASE_SHA,'בסיס ה-PR')}
 if(mode==='push')return {kind:'skip',reason:'ריצת push — ההשוואה מול main רצה ב-PR, לא כאן'};
 if(mode)return {kind:'error',reason:`EXPORT_BASE לא מוכר: «${mode}» (מותר: required · push)`};
 if(env.CI)return {kind:'error',reason:'CI בלי EXPORT_BASE — שלב «הכנת גרסת הבסיס» הוסר מה-workflow?'};
 /* מקומית */
 if(env.EXPORT_BASE_HTML)return fromFile(env.EXPORT_BASE_HTML,'','קובץ מקומי');
 let html,sha='';
 try{html=io.gitShow()}catch(e){
  const why=String((e&&e.stderr)||e&&e.message||e).trim().split('\n')[0];
  return {kind:'skip',reason:`הרצה מקומית בלי origin/main (${why})`}}
 try{sha=io.gitRev()}catch(_){/* ה-SHA הוא תיאור בלבד; הבסיס עצמו כבר נקרא */}
 if(!html.length)return {kind:'error',reason:'origin/main:index.html ריק'};
 if(html.equals(cur))return {kind:'skip',reason:`index.html זהה ל-origin/main ${short(sha)} — אין מה להשוות`};
 return {kind:'compare',html,sha,src:'origin/main'}}
module.exports={resolveBase,short};
