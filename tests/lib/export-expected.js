/* ============ PR-B · שינוי מוצהר בייצוא — מדויק, לא היתר ============
   השוואת הייצוא מול main (PR-A) דורשת זהות תא בתא. שינוי עסקי מכוון
   (למשל ROP שמחושב אחרת) חייב לשנות את הייצוא — ועד כאן הדרך היחידה
   לעבור הייתה לא להשוות. כאן ההשוואה נשארת, ומה שמותר לזוז מוצהר
   בקובץ tests/export-expected.json, תא אחר תא:

     { "base": "<SHA מלא של ה-main שמולו משווים>",
       "changes": [ { "sheet", "field", "pn", "from", "to" } ] }

   הכללים, כולם FAIL ולא אזהרה:
     · ברירת המחדל — changes ריק: כל הבדל נכשל.
     · כל הבדל שאינו מוצהר בדיוק (גיליון, שם עמודה, מק״ט, ערך קודם,
       ערך חדש — גם הסוג: 0.67 אינו "0.67") נכשל.
     · כל הצהרה שלא קרתה נכשלת — היתר שאינו בשימוש הוא היתר פתוח.
     · הצהרה לא ריקה תקפה רק מול ה-SHA שבה. בסיס אחר, או השוואה מול
       קובץ מקומי בלי SHA — נכשל. אחרי מיזוג ה-PR שהצהיר, ה-PR הבא
       מאפס את הקובץ (ההודעה אומרת את זה).
     · אין תבניות: "*" בשדה כלשהו נדחה, וכל שדה חובה.
     · מה שאי אפשר להצהיר עליו — אינו מותר בכלל: שינוי בשמות או בסדר
       הגיליונות, בשורות שלפני הכותרת, בכותרת עצמה, במספר השורות או
       בסדר המק״טים, בתא בלי מק״ט, בעמודה ששמה כפול, בסוג תא שערכו לא
       השתנה, או ברוחבי העמודות.

   הפונקציות טהורות (ספר XLSX נכנס, רשימות יוצאות), ולכן
   export-expected.test.js בודק כל ענף בלי דפדפן. */
const XLSX=require('xlsx');
const PN_RE=/^\s*מק[״"']ט/;
const SHA_RE=/^[0-9a-f]{40}$/;
/* כל ההבדלים בין שני ספרים: declarable — תא בשורת מק״ט; other — כל השאר */
function cellDiffs(A,B){
 const declarable=[],other=[];
 if(A.SheetNames.join('\u0001')!==B.SheetNames.join('\u0001')){
  other.push(`גיליונות: ${A.SheetNames.join(', ')} → ${B.SheetNames.join(', ')}`);return {declarable,other}}
 for(const sn of A.SheetNames){
  const wa=A.Sheets[sn],wb=B.Sheets[sn];
  const a=XLSX.utils.sheet_to_json(wa,{header:1,raw:true,defval:null}),b=XLSX.utils.sheet_to_json(wb,{header:1,raw:true,defval:null});
  const hi=a.findIndex(r=>r&&r.some(c=>PN_RE.test(String(c==null?'':c))));
  const hj=b.findIndex(r=>r&&r.some(c=>PN_RE.test(String(c==null?'':c))));
  const same=(x,y)=>JSON.stringify(x)===JSON.stringify(y);
  const wid=w=>(w['!cols']||[]).map(c=>c&&c.wch).join(',');
  if(wid(wa)!==wid(wb))other.push(`${sn}: רוחבי עמודות`);
  if(hi<0||hi!==hj){if(!same(a,b))other.push(`${sn}: אין שורת כותרת עם מק״ט — הגיליון השתנה`);continue}
  if(!same(a.slice(0,hi+1),b.slice(0,hj+1))){other.push(`${sn}: שורות הכותרת השתנו`);continue}
  if(a.length!==b.length){other.push(`${sn}: ${a.length-hi-1} שורות → ${b.length-hj-1}`);continue}
  const H=a[hi],pc=H.findIndex(c=>PN_RE.test(String(c==null?'':c)));
  const dup=new Set(H.filter((h,i)=>h!=null&&H.indexOf(h)!==i));
  for(let r=hi+1;r<a.length;r++){const x=a[r]||[],y=b[r]||[];
   if(x[pc]!==y[pc]){other.push(`${sn} שורה ${r+1}: מק״ט ${x[pc]} → ${y[pc]} (סדר או שורה שונים)`);continue}
   const pn=x[pc];const n=Math.max(x.length,y.length,H.length);
   for(let c=0;c<n;c++){const u=x[c]===undefined?null:x[c],v=y[c]===undefined?null:y[c];
    const ref=XLSX.utils.encode_cell({r,c}),ta=wa[ref]&&wa[ref].t,tb=wb[ref]&&wb[ref].t;
    if(u===v){if(ta!==tb)other.push(`${sn}!${ref}: סוג התא ${ta} → ${tb} בלי שינוי בערך`);continue}
    const field=H[c];
    if(pn==null||pn===''||field==null||field===''||dup.has(field)){
     other.push(`${sn}!${ref}: «${u}» → «${v}» — ${pn==null||pn===''?'שורה בלי מק״ט':'עמודה בלי שם ייחודי'}`);continue}
    declarable.push({sheet:sn,field:String(field),pn:String(pn),from:u,to:v})}}}
 return {declarable,other}}
const keyOf=d=>JSON.stringify([d.sheet,d.field,d.pn,d.from,d.to]);
/* תקינות ההצהרה עצמה — לפני שמשווים אליה */
function validateDecl(decl){
 const errs=[];
 if(!decl||typeof decl!=='object'||Array.isArray(decl))return ['הקובץ אינו אובייקט'];
 const extra=Object.keys(decl).filter(k=>!['base','changes','note'].includes(k));
 if(extra.length)errs.push('מפתחות לא מוכרים: '+extra.join(', '));
 if(!Array.isArray(decl.changes))return errs.concat('changes אינו מערך');
 if(decl.changes.length&&!SHA_RE.test(String(decl.base||'')))errs.push(`base חייב להיות SHA מלא (40 תווים) כשיש שינויים — «${decl.base}»`);
 if(!decl.changes.length&&decl.base!=null)errs.push('base חייב להיות null כשאין שינויים — הצהרה ריקה אינה קשורה לבסיס');
 const seen=new Set();
 decl.changes.forEach((d,i)=>{
  const miss=['sheet','field','pn','from','to'].filter(k=>!(d&&k in d));
  if(miss.length){errs.push(`שינוי ${i+1}: חסר ${miss.join(', ')}`);return}
  const ex=Object.keys(d).filter(k=>!['sheet','field','pn','from','to'].includes(k));
  if(ex.length)errs.push(`שינוי ${i+1}: מפתחות לא מוכרים ${ex.join(', ')}`);
  for(const k of ['sheet','field','pn'])if(typeof d[k]!=='string'||!d[k].trim()||d[k].includes('*'))errs.push(`שינוי ${i+1}: ${k} חייב להיות טקסט מדויק, בלי «*»`);
  for(const k of ['from','to'])if(typeof d[k]==='string'&&d[k].includes('*'))errs.push(`שינוי ${i+1}: ${k} מכיל «*» — אין תבניות`);
  if(JSON.stringify(d.from)===JSON.stringify(d.to))errs.push(`שינוי ${i+1}: from = to — אין כאן שינוי`);
  const k=keyOf(d);if(seen.has(k))errs.push(`שינוי ${i+1}: כפול`);seen.add(k)});
 return errs}
/* ההכרעה: ok, ורשימות של מה שנכשל */
function checkDeclared(diffs,decl,baseSha){
 const errs=validateDecl(decl);
 const res={ok:false,errs,undeclared:[],unused:[],other:diffs.other.slice(),matched:0};
 if(errs.length)return res;
 if(decl.changes.length){
  if(!baseSha)errs.push('ההצהרה קשורה ל-SHA, וההשוואה רצה מול קובץ בלי SHA — אי אפשר לאמת שזה הבסיס הנכון');
  else if(decl.base!==baseSha)errs.push(`ההצהרה שייכת לבסיס ${decl.base.slice(0,7)}, וההשוואה מול ${String(baseSha).slice(0,7)}. אם ה-PR שהצהיר כבר מוזג — אפס את tests/export-expected.json ל-{"base":null,"changes":[]}`)}
 if(errs.length)return res;
 const want=new Map();for(const d of decl.changes)want.set(keyOf(d),d);
 for(const d of diffs.declarable){const k=keyOf(d);if(want.has(k)){want.delete(k);res.matched++}else res.undeclared.push(d)}
 res.unused=[...want.values()];
 res.ok=!res.undeclared.length&&!res.unused.length&&!res.other.length;
 return res}
const fmt=d=>`${d.sheet} · ${d.field} · ${d.pn}: ${JSON.stringify(d.from)} → ${JSON.stringify(d.to)}`;
module.exports={cellDiffs,validateDecl,checkDeclared,keyOf,fmt};
