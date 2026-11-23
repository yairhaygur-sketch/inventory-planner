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
       בקבוצת המק״טים, בתא בלי מק״ט, בעמודה ששמה כפול, בסוג תא שערכו לא
       השתנה, או ברוחבי העמודות.
     · סדר המק״טים (אושר 11.11, OPEN_DEL): כשאותם מק״טים בדיוק — ייחודיים,
       באותו מספר שורות — רק החליפו מקום, מצהירים ב-"moves" על **הפריטים
       שזזו** ולא על כל שורה שהוסטה: [{sheet, pn, from, to}], from/to מספרי
       השורות בגיליון. הרשימה היא המינימלית והדטרמיניסטית (כל מי שאינו
       בתת-הסדרה העולה הארוכה ביותר — LIS), ולכן יש לה צורה אחת בלבד.
       תזוזה לא מוצהרת נכשלת, תזוזה מוצהרת שלא קרתה נכשלת, ושינוי תא
       בשורה שזזה נבדק מול אותו מק״ט, לא מול אותו מספר שורה.

     · שינוי מבנה (XSTRUCT, אושר 23.11 לרפורמת האקסל): "structure" מצהיר על
       המבנה החדש במדויק — רשימת הגיליונות בסדרם, ולכל גיליון חדש מאיזה גיליון
       ישן הוא בא (או null — תוכן חדש, למשל «סיכום»), ולכל עמודה בו מאיזו עמודה
       ישנה (או null — עמודה חדשה). גיליון ישן שלא ממשיך לשום מקום חייב להופיע
       ב-droppedSheets, ועמודה ישנה שלא ממשיכה — ב-dropped. עם מבנה, הערכים
       ממשיכים להיות מושווים מק״ט מול מק״ט בין העמודה הישנה לחדשה, ושינוי ערך
       עדיין חייב להופיע ב-changes (sheet/field — השמות החדשים). מה שאינו נבדק
       במצב הזה הוא מה שהרפורמה משנה בכוונה: שורות שלפני הכותרת ורוחבי עמודות.
       סדר השורות נשמר, אלא אם הגיליון מצהיר "reorder": true.

   הפונקציות טהורות (ספר XLSX נכנס, רשימות יוצאות), ולכן
   export-expected.test.js בודק כל ענף בלי דפדפן. */
const XLSX=require('xlsx');
const PN_RE=/^\s*מק[״"']ט/;
const SHA_RE=/^[0-9a-f]{40}$/;
/* כל ההבדלים בין שני ספרים: declarable — תא בשורת מק״ט; other — כל השאר */
/* הפריטים שזזו: מי שאינו ב-LIS של המיקומים החדשים לפי הסדר הישן. דטרמיניסטי
   (מק״טים ייחודיים — אין שוויונות, ולכן ה-LIS נבנה תמיד באותה דרך). */
function movedRows(sn,oldPns,newPns,o0,n0){
 const pos=new Map(newPns.map((p,i)=>[p,i]));const seq=oldPns.map(p=>pos.get(p));
 const tails=[],tailIdx=[],prev=new Array(seq.length).fill(-1);
 seq.forEach((v,i)=>{let lo=0,hi=tails.length;while(lo<hi){const m=(lo+hi)>>1;if(tails[m]<v)lo=m+1;else hi=m}
  tails[lo]=v;tailIdx[lo]=i;prev[i]=lo>0?tailIdx[lo-1]:-1});
 const keep=new Set();if(tails.length)for(let i=tailIdx[tails.length-1];i>=0;i=prev[i])keep.add(i);
 const out=[];oldPns.forEach((p,i)=>{if(!keep.has(i))out.push({sheet:sn,pn:String(p),from:o0+i+1,to:n0+pos.get(p)+1})});
 return out.sort((a,b)=>a.to-b.to)}
/* XSTRUCT — ההשוואה כשהמבנה מוצהר. ראה ההערה בראש הקובץ. */
function sheetRows(ws){
 const a=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
 const hi=a.findIndex(r=>r&&r.some(c=>PN_RE.test(String(c==null?'':c))));
 if(hi<0)return null;
 const H=a[hi].map(h=>h==null?null:String(h)),pc=H.findIndex(c=>PN_RE.test(String(c==null?'':c)));
 const rows=a.slice(hi+1).filter(r=>r&&r.some(c=>c!=null&&c!==''));
 return {H,pc,rows,hi}}
function structDiffs(A,B,S){
 const declarable=[],other=[],moves=[];
 if(JSON.stringify(B.SheetNames)!==JSON.stringify(S.sheets)){
  other.push(`גיליונות: ${B.SheetNames.join(', ')} — המבנה המוצהר: ${S.sheets.join(', ')}`);return {declarable,other,moves}}
 const used={};/* גיליון ישן → העמודות שממשיכות ממנו */
 for(const sn of S.sheets){const m=S.map[sn];
  if(!m){other.push(`${sn}: הגיליון אינו מוצהר ב-structure.map`);continue}
  if(m.from==null)continue;
  if(!A.SheetNames.includes(m.from)){other.push(`${sn}: הגיליון הישן «${m.from}» אינו בבסיס`);continue}
  const o=sheetRows(A.Sheets[m.from]),n=sheetRows(B.Sheets[sn]);
  if(!o||!n){other.push(`${sn}: אין שורת כותרת עם מק״ט (בבסיס או ב-PR)`);continue}
  const cols=m.columns||{},keys=Object.keys(cols);
  const nh=n.H.filter(h=>h!=null&&h!=='');
  const missing=nh.filter(h=>!(h in cols)),ghost=keys.filter(k=>!nh.includes(k));
  if(missing.length)other.push(`${sn}: עמודות שאינן מוצהרות — ${missing.join(', ')}`);
  if(ghost.length)other.push(`${sn}: עמודות מוצהרות שאינן בקובץ — ${ghost.join(', ')}`);
  if(new Set(nh).size!==nh.length){other.push(`${sn}: שם עמודה כפול`);continue}
  if(missing.length||ghost.length)continue;
  const badOld=keys.filter(k=>cols[k]!=null&&!o.H.includes(cols[k]));
  if(badOld.length){other.push(`${sn}: עמודות ישנות שאינן בבסיס — ${badOld.map(k=>cols[k]).join(', ')}`);continue}
  (used[m.from]=used[m.from]||new Set());keys.forEach(k=>{if(cols[k]!=null)used[m.from].add(cols[k])});
  const oP=o.rows.map(r=>r[o.pc]),nP=n.rows.map(r=>r[n.pc]);
  const uniq=L=>new Set(L).size===L.length,okv=v=>v!=null&&v!=='';
  if(!uniq(oP)||!uniq(nP)||!oP.every(okv)||!nP.every(okv)){other.push(`${sn}: מק״ט כפול או ריק — לא ניתן להשוות`);continue}
  const NS=new Set(nP);
  if(oP.length!==nP.length||!oP.every(p=>NS.has(p))){other.push(`${sn}: קבוצת המק״טים שונה (${oP.length} → ${nP.length})`);continue}
  if(!m.reorder&&oP.some((p,i)=>p!==nP[i])){const i=oP.findIndex((p,k)=>p!==nP[k]);
   other.push(`${sn}: סדר השורות השתנה (שורה ${i+1}: ${oP[i]} → ${nP[i]}) — הצהר "reorder": true אם זה מכוון`);continue}
  const at=new Map(nP.map((p,i)=>[p,i]));
  for(let i=0;i<o.rows.length;i++){const x=o.rows[i],pn=oP[i],y=n.rows[at.get(pn)];
   for(const k of keys){const oc=cols[k];if(oc==null)continue;
    const u=x[o.H.indexOf(oc)],v=y[n.H.indexOf(k)];
    const uu=u===undefined?null:u,vv=v===undefined?null:v;
    if(JSON.stringify(uu)===JSON.stringify(vv))continue;
    declarable.push({sheet:sn,field:k,pn:String(pn),from:uu,to:vv})}}}
 /* שום דבר לא נעלם בשקט: כל גיליון ישן ממשיך או מוצהר כמוסר, וכל עמודה בו ממשיכה או מוצהרת */
 const fromSet=new Set(S.sheets.map(sn=>S.map[sn]&&S.map[sn].from).filter(x=>x!=null));
 for(const os of A.SheetNames){
  if(!fromSet.has(os)){if(!(S.droppedSheets||[]).includes(os))other.push(`הגיליון הישן «${os}» נעלם בלי הצהרה (droppedSheets)`);continue}
  const o=sheetRows(A.Sheets[os]);if(!o)continue;
  const dr=(S.dropped||{})[os]||[];
  const lost=o.H.filter(h=>h!=null&&h!==''&&!(used[os]&&used[os].has(h))&&!dr.includes(h));
  if(lost.length)other.push(`${os}: עמודות שנעלמו בלי הצהרה (dropped) — ${lost.join(', ')}`);
  const fake=dr.filter(h=>!o.H.includes(h)||(used[os]&&used[os].has(h)));
  if(fake.length)other.push(`${os}: dropped מצהיר על עמודות שאינן נעלמות — ${fake.join(', ')}`)}
 for(const ds of S.droppedSheets||[])if(!A.SheetNames.includes(ds)||fromSet.has(ds))other.push(`droppedSheets: «${ds}» אינו גיליון ישן שנעלם`);
 return {declarable,other,moves}}
function cellDiffs(A,B,structure){
 if(structure)return structDiffs(A,B,structure);
 const declarable=[],other=[],moves=[];
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
  /* סדר: אותה קבוצת מק״טים, ייחודיים ולא ריקים — התזוזה ניתנת להצהרה; אחרת לא */
  const oP=a.slice(hi+1).map(x=>(x||[])[pc]),nP=b.slice(hj+1).map(x=>(x||[])[pc]);
  let rowOf=null;
  if(oP.some((p,i)=>p!==nP[i])){
   const okv=v=>v!=null&&v!=='';const uniq=L=>new Set(L).size===L.length;
   const S=new Set(nP);
   const same=uniq(oP)&&uniq(nP)&&oP.every(okv)&&nP.every(okv)&&oP.every(p=>S.has(p));
   if(!same){const i=oP.findIndex((p,k)=>p!==nP[k]);other.push(`${sn} שורה ${hi+2+i}: מק״ט ${oP[i]} → ${nP[i]} (קבוצת המק״טים שונה, או מק״ט כפול/ריק — לא ניתן להצהיר)`);continue}
   moves.push(...movedRows(sn,oP,nP,hi+1,hj+1));
   const at=new Map(nP.map((p,i)=>[p,hj+1+i]));rowOf=r=>at.get(oP[r-hi-1])}
  for(let r=hi+1;r<a.length;r++){const rb=rowOf?rowOf(r):r;const x=a[r]||[],y=b[rb]||[];
   const pn=x[pc];const n=Math.max(x.length,y.length,H.length);
   for(let c=0;c<n;c++){const u=x[c]===undefined?null:x[c],v=y[c]===undefined?null:y[c];
    const ref=XLSX.utils.encode_cell({r,c}),refB=XLSX.utils.encode_cell({r:rb,c}),ta=wa[ref]&&wa[ref].t,tb=wb[refB]&&wb[refB].t;
    if(u===v){if(ta!==tb)other.push(`${sn}!${refB}: סוג התא ${ta} → ${tb} בלי שינוי בערך`);continue}
    const field=H[c];
    if(pn==null||pn===''||field==null||field===''||dup.has(field)){
     other.push(`${sn}!${ref}: «${u}» → «${v}» — ${pn==null||pn===''?'שורה בלי מק״ט':'עמודה בלי שם ייחודי'}`);continue}
    declarable.push({sheet:sn,field:String(field),pn:String(pn),from:u,to:v})}}}
 return {declarable,other,moves}}
const keyOf=d=>JSON.stringify([d.sheet,d.field,d.pn,d.from,d.to]);
const moveKey=d=>JSON.stringify([d.sheet,d.pn,d.from,d.to]);
/* תקינות ההצהרה עצמה — לפני שמשווים אליה */
function validateDecl(decl){
 const errs=[];
 if(!decl||typeof decl!=='object'||Array.isArray(decl))return ['הקובץ אינו אובייקט'];
 const extra=Object.keys(decl).filter(k=>!['base','changes','moves','note','structure'].includes(k));
 if(extra.length)errs.push('מפתחות לא מוכרים: '+extra.join(', '));
 if(!Array.isArray(decl.changes))return errs.concat('changes אינו מערך');
 if('moves' in decl&&!Array.isArray(decl.moves))return errs.concat('moves אינו מערך');
 const mv=decl.moves||[],any=decl.changes.length+mv.length+(decl.structure?1:0);
 if('structure' in decl){const S=decl.structure,isObj=o=>o&&typeof o==='object'&&!Array.isArray(o);
  if(!isObj(S))errs.push('structure חייב להיות אובייקט');
  else{
   const ex=Object.keys(S).filter(k=>!['sheets','map','dropped','droppedSheets'].includes(k));
   if(ex.length)errs.push('structure: מפתחות לא מוכרים '+ex.join(', '));
   if(!Array.isArray(S.sheets)||!S.sheets.length||!S.sheets.every(x=>typeof x==='string'&&x.trim()&&!x.includes('*')))errs.push('structure.sheets חייב להיות רשימת שמות גיליונות מדויקים');
   if(!isObj(S.map))errs.push('structure.map חייב להיות אובייקט');
   else for(const [sn,m] of Object.entries(S.map)){
    if(!isObj(m)){errs.push(`structure.map.${sn}: חייב להיות אובייקט`);continue}
    const mx=Object.keys(m).filter(k=>!['from','columns','reorder'].includes(k));
    if(mx.length)errs.push(`structure.map.${sn}: מפתחות לא מוכרים ${mx.join(', ')}`);
    if(!('from' in m)||!(m.from===null||typeof m.from==='string'))errs.push(`structure.map.${sn}: from חייב להיות שם גיליון או null`);
    if(m.from!=null&&!isObj(m.columns))errs.push(`structure.map.${sn}: columns חייב להיות אובייקט (עמודה חדשה → ישנה או null)`);
    if(isObj(m.columns))for(const [k,v] of Object.entries(m.columns))if(k.includes('*')||!(v===null||(typeof v==='string'&&!v.includes('*'))))errs.push(`structure.map.${sn}.${k}: שם מדויק או null, בלי «*»`);
    if('reorder' in m&&typeof m.reorder!=='boolean')errs.push(`structure.map.${sn}: reorder חייב להיות true/false`)}
   if('dropped' in S&&(!isObj(S.dropped)||!Object.values(S.dropped).every(a=>Array.isArray(a)&&a.every(x=>typeof x==='string'&&!x.includes('*')))))errs.push('structure.dropped: גיליון ישן → רשימת עמודות מדויקות');
   if('droppedSheets' in S&&(!Array.isArray(S.droppedSheets)||!S.droppedSheets.every(x=>typeof x==='string'&&!x.includes('*'))))errs.push('structure.droppedSheets: רשימת שמות מדויקים')}}
 if(any&&!SHA_RE.test(String(decl.base||'')))errs.push(`base חייב להיות SHA מלא (40 תווים) כשיש שינויים — «${decl.base}»`);
 if(!any&&decl.base!=null)errs.push('base חייב להיות null כשאין שינויים — הצהרה ריקה אינה קשורה לבסיס');
 const seenM=new Set();
 mv.forEach((d,i)=>{
  const miss=['sheet','pn','from','to'].filter(k=>!(d&&k in d));
  if(miss.length){errs.push(`תזוזה ${i+1}: חסר ${miss.join(', ')}`);return}
  const ex=Object.keys(d).filter(k=>!['sheet','pn','from','to'].includes(k));
  if(ex.length)errs.push(`תזוזה ${i+1}: מפתחות לא מוכרים ${ex.join(', ')}`);
  for(const k of ['sheet','pn'])if(typeof d[k]!=='string'||!d[k].trim()||d[k].includes('*'))errs.push(`תזוזה ${i+1}: ${k} חייב להיות טקסט מדויק, בלי «*»`);
  for(const k of ['from','to'])if(!Number.isInteger(d[k])||d[k]<1)errs.push(`תזוזה ${i+1}: ${k} חייב להיות מספר שורה שלם`);
  if(d.from===d.to)errs.push(`תזוזה ${i+1}: from = to — אין כאן תזוזה`);
  const k=JSON.stringify([d.sheet,d.pn]);if(seenM.has(k))errs.push(`תזוזה ${i+1}: אותו מק״ט פעמיים`);seenM.add(k)});
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
 const res={ok:false,errs,undeclared:[],unused:[],undeclaredMoves:[],unusedMoves:[],other:diffs.other.slice(),matched:0,movesMatched:0};
 if(errs.length)return res;
 if(decl.changes.length||(decl.moves||[]).length||decl.structure){
  if(!baseSha)errs.push('ההצהרה קשורה ל-SHA, וההשוואה רצה מול קובץ בלי SHA — אי אפשר לאמת שזה הבסיס הנכון');
  else if(decl.base!==baseSha)errs.push(`ההצהרה שייכת לבסיס ${decl.base.slice(0,7)}, וההשוואה מול ${String(baseSha).slice(0,7)}. אם ה-PR שהצהיר כבר מוזג — אפס את tests/export-expected.json ל-{"base":null,"changes":[]}`)}
 if(errs.length)return res;
 const want=new Map();for(const d of decl.changes)want.set(keyOf(d),d);
 for(const d of diffs.declarable){const k=keyOf(d);if(want.has(k)){want.delete(k);res.matched++}else res.undeclared.push(d)}
 res.unused=[...want.values()];
 const wantM=new Map();for(const d of decl.moves||[])wantM.set(moveKey(d),d);
 for(const d of diffs.moves||[]){const k=moveKey(d);if(wantM.has(k)){wantM.delete(k);res.movesMatched++}else res.undeclaredMoves.push(d)}
 res.unusedMoves=[...wantM.values()];
 res.ok=!res.undeclared.length&&!res.unused.length&&!res.undeclaredMoves.length&&!res.unusedMoves.length&&!res.other.length;
 return res}
const fmt=d=>`${d.sheet} · ${d.field} · ${d.pn}: ${JSON.stringify(d.from)} → ${JSON.stringify(d.to)}`;
const fmtMove=d=>`${d.sheet} · ${d.pn}: שורה ${d.from} → ${d.to}`;
module.exports={cellDiffs,structDiffs,validateDecl,checkDeclared,keyOf,moveKey,fmt,fmtMove};
