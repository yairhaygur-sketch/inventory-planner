#!/usr/bin/env node
/* ============ הוכחה שהשומר תופס — שינויים מכוונים בעותק ============
   שימוש:  node business-canary.js <zmrp.xlsx> [--html index.html] [--out <dir>] [--jobs 3]

   בדיקה שתמיד עוברת אינה בדיקה. הסקריפט הזה יוצר עותקים של index.html
   בתיקייה זמנית, בכל אחד שינוי קטן ומכוון אחד, מצלם כל עותק, ומוודא
   ש-business-diff מוצא הבדל — ובתחום הנכון. בנוסף: עותק לא משונה חייב
   לצאת 0 הבדלים.

   כל שינוי מוגדר כמחרוזת מדויקת שחייבת להופיע בקוד פעם אחת בדיוק. אם
   הקוד השתנה והמחרוזת כבר לא שם — הסקריפט נכשל ואומר איזו, כדי שיעדכנו
   את השינוי ולא ידלגו עליו בשקט.

   הקובץ index.html עצמו לא משתנה. העותקים והצילומים — מחוץ לריפו. */
const {execFile}=require('child_process');
const path=require('path'),fs=require('fs'),os=require('os');
const {compare}=require('./business-diff.js');
const args=process.argv.slice(2);
const opt=(k,d)=>{const i=args.indexOf(k);if(i<0)return d;const v=args[i+1];args.splice(i,2);return v};
const html=opt('--html','index.html'),JOBS=+opt('--jobs','3');let out=opt('--out',null);
const [xlsx]=args;
if(!xlsx){console.error('שימוש: business-canary.js <zmrp.xlsx> [--html index.html] [--out <dir>] [--jobs 3]');process.exit(2)}
const root=path.resolve(__dirname,'../../../..');
const inRepo=d=>{const r=fs.realpathSync(root);let x=path.resolve(d);while(!fs.existsSync(x))x=path.dirname(x);x=fs.realpathSync(x);return x===r||x.startsWith(r+path.sep)};
if(out&&inRepo(out)){console.error('--out בתוך הריפו — בחר תיקייה מחוץ לריפו');process.exit(2)}
if(!out)out=fs.mkdtempSync(path.join(os.tmpdir(),'canary-'));fs.mkdirSync(out,{recursive:true});

/* כל שינוי: מה הוא מדמה, מה מחליפים, ובאילו תחומים ההבדל חייב להופיע */
const MUT=[
 /* הסף A (9→8) נבדק ונפסל כשינוי בוחן: מאז שהעלאות C→A הוצאו מהדוח, ובדוח
    האמיתי יש פריט ZM-A אחד, לשינוי הזה אין שום תוצאה נראית — אף כלי לא
    יכול לתפוס אותו על הנתונים האלה. הסף C/D מכריע PD↔ZM-C, וזה מורגש. */
 {id:'threshold',what:'סף: פרופיל ZM-C מ-2 חודשי צריכה ל-3',
  from:"h>=2?'C':'D'}",to:"h>=3?'C':'D'}",expect:['items','screens']},
 {id:'ss-calc',what:'חישוב: מלאי ביטחון מוצע גדל ביחידה',
  from:'let sugSS=noRec?0:Math.max(ssStat,ssMinEff,lowFloor);',to:'let sugSS=noRec?0:Math.max(ssStat,ssMinEff,lowFloor)+1;',expect:['items','exports']},
 {id:'money',what:'כסף: מלאי לא פעיל נספר בלי יחידה אחת',
  from:"it.cat='ללא פעולה';it.expCap=it.price*Math.max(0,it.stock);",to:"it.cat='ללא פעולה';it.expCap=it.price*Math.max(0,it.stock-1);",expect:['items','screens']},
 {id:'ranking',what:'דירוג: «לקוח ממתין ללא רכש» כבר לא ראשון',
  from:'(rung(a)-rung(b))',to:'(rung(b)-rung(a))',expect:['screens','exports']},
 {id:'sap-file',what:'קובץ SAP בלבד: ROP בקובץ גדל ביחידה (המנוע לא נוגע)',
  from:"r.sugROP, r.pGroup||''",to:"r.sugROP+1, r.pGroup||''",expect:['exports'],only:true},
 {id:'excel-header',what:'אקסל בלבד: כותרת עמודה «תיאור» שונתה',
  from:"['תיאור',r=>r.desc]",to:"['תאור',r=>r.desc]",expect:['exports'],only:true},
 {id:'mail',what:'טיוטת המייל לספקים: משפט אחד שונה',
  from:'Please confirm the shipping date and quantity for each line.',to:'Please confirm the delivery date and quantity for each line.',expect:['exports'],only:true},
 {id:'expedite',what:'דוח זירוז: שם גיליון שונה',
  from:"XLSX.utils.book_append_sheet(wb,ws,'לזירוז');",to:"XLSX.utils.book_append_sheet(wb,ws,'לזרוז');",expect:['exports'],only:true}];

const src=fs.readFileSync(html,'utf8');
const bad=MUT.filter(m=>src.split(m.from).length!==2);
if(bad.length){console.error('מחרוזות שינוי שאינן מופיעות פעם אחת בדיוק בקוד:\n'+bad.map(m=>`  ${m.id}: ${m.from}`).join('\n')+'\nעדכן את MUT — אל תדלג.');process.exit(1)}
const now=(()=>{const d=new Date();d.setHours(8,0,0,0);return d.toISOString()})();
const snapshot=(h,o)=>new Promise((res,rej)=>execFile(process.execPath,[path.join(__dirname,'business-snapshot.js'),h,xlsx,o,'--now',now],
 {maxBuffer:1<<24,env:process.env},(e,so,se)=>e?rej(new Error((se||so||e.message).trim().split('\n').pop())):res(o)));
const jobs=[{id:'clean',what:'עותק לא משונה (פעמיים)',file:path.join(out,'clean.html'),code:src},
 ...MUT.map(m=>({...m,file:path.join(out,m.id+'.html'),code:src.replace(m.from,m.to)}))];
jobs.push({id:'clean2',what:'',file:path.join(out,'clean2.html'),code:src});
for(const j of jobs)fs.writeFileSync(j.file,j.code);
(async()=>{
 console.log(`${MUT.length} שינויים מכוונים + 2 עותקים נקיים · דוח: ${path.basename(xlsx)} · ${JOBS} במקביל`);
 const q=[...jobs];const run=async()=>{while(q.length){const j=q.shift();j.json=j.file.replace(/\.html$/,'.json');
  try{await snapshot(j.file,j.json)}catch(e){j.err=e.message}}};
 await Promise.all(Array.from({length:JOBS},run));
 const load=f=>JSON.parse(fs.readFileSync(f,'utf8'));
 const clean=jobs.find(j=>j.id==='clean'),clean2=jobs.find(j=>j.id==='clean2');
 if(clean.err||clean2.err){console.error('צילום העותק הנקי נכשל: '+(clean.err||clean2.err));process.exit(1)}
 const C=load(clean.json);let fail=0;
 const self=compare(C,load(clean2.json));
 console.log(`\n${self.total===0?'✓':'✗'} עותק נקי מול עותק נקי: ${self.total} הבדלים${self.total?' — הצילום אינו יציב':''}`);if(self.total)fail++;
 for(const m of MUT){const j=jobs.find(x=>x.id===m.id);
  if(j.err){console.log(`✗ ${m.what}: הצילום נכשל — ${j.err}`);fail++;continue}
  const r=compare(C,load(j.json));const hit=Object.fromEntries(r.sections.map(s=>[s.key,s.n]));
  const miss=m.expect.filter(k=>!hit[k]);const leak=m.only?Object.keys(hit).filter(k=>hit[k]&&!m.expect.includes(k)):[];
  const ok=r.total>0&&!miss.length&&!leak.length;if(!ok)fail++;
  console.log(`${ok?'✓':'✗'} ${m.what}: ${r.total} הבדלים · ${r.sections.filter(s=>s.n).map(s=>`${s.key} ${s.n}`).join(' · ')||'—'}`
   +(miss.length?` · לא נתפס ב: ${miss.join(', ')}`:'')+(leak.length?` · דלף ל: ${leak.join(', ')}`:''))}
 console.log(`\n${fail?`✗ ${fail} בדיקות נכשלו`:`✓ כל ${MUT.length} השינויים נתפסו, בתחום הנכון, והעותק הנקי יצא 0`}\nעותקים וצילומים: ${out}`);
 process.exit(fail?1:0)})();
