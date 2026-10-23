#!/usr/bin/env node
/* ============ שומר הלוגיקה העסקית — פקודה אחת לכל PR ============
   שימוש:
     node business-guard.js <zmrp.xlsx> [--base <git-ref>] [--html <file>] [--eta <eta.xlsx>] [--out <dir>] [--show-pn]

   מוציא את index.html מ-<git-ref> (ברירת מחדל origin/main) לתיקייה זמנית,
   מצלם אותו ואת הקובץ הנוכחי (--html, ברירת מחדל index.html) באותו שעון
   מקובע ועל אותו דוח, ומשווה ב-business-diff.js. יוצא 0 רק אם הלוגיקה,
   המספרים והייצוא זהים לחלוטין.

   שינוי ממשק (D2) חייב לצאת 0. שינוי עסקי מכוון יוצא 1 — וכל הבדל בפלט
   חייב להיות מוסבר ומאושר לפני מיזוג.

   הצילומים נשמרים ב---out (ברירת מחדל: תיקייה זמנית של מערכת ההפעלה).
   הם מכילים מק״טים אמיתיים — לעולם לא בתוך הריפו. */
const {execFileSync}=require('child_process');
const path=require('path'),fs=require('fs'),os=require('os');
const args=process.argv.slice(2);
const opt=(k,d)=>{const i=args.indexOf(k);if(i<0)return d;const v=args[i+1];args.splice(i,2);return v};
const flag=k=>{const i=args.indexOf(k);if(i<0)return false;args.splice(i,1);return true};
const base=opt('--base','origin/main'),html=opt('--html','index.html'),eta=opt('--eta',null);
const show=flag('--show-pn');let out=opt('--out',null);
const [xlsx]=args;
if(!xlsx){console.error('שימוש: business-guard.js <zmrp.xlsx> [--base <git-ref>] [--html <file>] [--eta <eta.xlsx>] [--out <dir>] [--show-pn]');process.exit(2)}
const root=execFileSync('git',['rev-parse','--show-toplevel'],{encoding:'utf8'}).trim();
/* הבדיקה לפני היצירה: גם תיקייה ריקה בתוך הריפו היא צעד לקראת קומיט של צילום */
const inRepo=d=>{const r=fs.realpathSync(root);let x=path.resolve(d);while(!fs.existsSync(x))x=path.dirname(x);x=fs.realpathSync(x);return x===r||x.startsWith(r+path.sep)};
if(out&&inRepo(out)){console.error('--out בתוך הריפו: הצילומים מכילים מק״טים אמיתיים. בחר תיקייה מחוץ לריפו.');process.exit(2)}
if(!out)out=fs.mkdtempSync(path.join(os.tmpdir(),'guard-'));
fs.mkdirSync(out,{recursive:true});
const sha=execFileSync('git',['rev-parse','--short',base],{cwd:root,encoding:'utf8'}).trim();
const baseHtml=path.join(out,`base-${sha}.html`);
fs.writeFileSync(baseHtml,execFileSync('git',['show',`${base}:index.html`],{cwd:root,maxBuffer:1<<28}));
const now=(()=>{const d=new Date();d.setHours(8,0,0,0);return d.toISOString()})();
const V=__dirname,snap=(h,o)=>execFileSync(process.execPath,[path.join(V,'business-snapshot.js'),h,xlsx,o,'--now',now,...(eta?['--eta',eta]:[])],{stdio:'inherit'});
console.log(`בסיס: ${base} (${sha}) · נבדק: ${html} · דוח: ${path.basename(xlsx)}${eta?' + '+path.basename(eta):''} · שעון: ${now}`);
const A=path.join(out,'before.json'),B=path.join(out,'after.json');
try{snap(baseHtml,A);snap(path.resolve(html),B)}catch(e){console.error('הצילום נכשל — ראה למעלה');process.exit(1)}
try{execFileSync(process.execPath,[path.join(V,'business-diff.js'),A,B,...(show?['--show-pn']:[])],{stdio:'inherit'});console.log(`צילומים: ${out}`);process.exit(0)}
catch(e){console.log(`צילומים: ${out}`);process.exit(1)}
