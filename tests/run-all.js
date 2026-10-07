/* ============ למה יש כאן רץ ולא שרשרת && ============
   `npm test` היה שרשרת של שבעה־עשר `node tests/X && node tests/Y`.
   משמעות ה-&& היא שהכישלון הראשון עוצר הכול: בהרצת ה-CI שנפלה
   ב-7.10 נקראו 686 בדיקות, נפלה אחת ב-`ui.test.js` — והקבצים
   הבאים, שתים־עשרה במספר, **לא הורצו בכלל**. כלומר לוג הכישלון
   הוכיח תקלה אחת והשאיר אלף בדיקות לא נמדדות, ואי אפשר היה לדעת
   אם יש עוד.

   הרץ הזה מריץ את כולם, מדפיס את הפלט של כל אחד, ומסכם בסוף.
   קוד היציאה נשאר 1 אם מישהו נפל — הוא לא מרפה על אמינות
   ה-CI, הוא רק מפסיק להסתיר. */
const {spawnSync}=require('child_process');const path=require('path');
const FILES=['offline.test.js','engine.test.js','defects.test.js','bulk.test.js',
 'ui.test.js','onboarding.test.js','search.test.js','brief.test.js','routes.test.js',
 'statuses.test.js','nlq.test.js','explain.test.js','eta.test.js','ledger.test.js',
 'expedite.test.js','day2.test.js','contrast.js'];
const SD=__dirname;const res=[];
for(const f of FILES){
 process.stdout.write(`\n========== ${f} ==========\n`);
 const t0=Date.now();
 const r=spawnSync(process.execPath,[path.join(SD,f)],{stdio:'inherit',env:process.env});
 const code=r.status===null?1:r.status;
 res.push({f,code,sec:Math.round((Date.now()-t0)/1000),sig:r.signal||''});
}
const bad=res.filter(r=>r.code!==0);
process.stdout.write('\n========== סיכום ==========\n');
for(const r of res)process.stdout.write(
 `${r.code===0?'עבר ':'נפל '} ${r.f.padEnd(20)} ${String(r.sec).padStart(4)}s`
 +(r.sig?` (אות ${r.sig})`:'')+'\n');
process.stdout.write(`\n${res.length-bad.length}/${res.length} קבצים עברו`
 +(bad.length?` · נפלו: ${bad.map(r=>r.f).join(', ')}`:'')+'\n');
process.exit(bad.length?1:0);
