#!/usr/bin/env node
/* ============ השוואת שני צילומים עסקיים ============
   שימוש:  node business-diff.js <before.json> <after.json> [--show-pn] [--max N]

   משווה כל מה ש-business-snapshot.js צילם, ומדפיס לפי תחום:
     פריטים   — כמה פריטים זזו, ובאיזה שדה (cat, sugSS, why, act, __prof…)
     תורים    — כל תור ודלת: מי נכנס ומי יצא
     מסכים    — מונה במסילה, מספר פריטים, סדר הרשימה, סכומים לפי מטבע
     ייצוא    — לכל מסך: אקסל (כל גיליון, תא בתא), SAP, מלאי איטי;
                דוח זירוז; טיוטת המייל לספקים
   יוצא 0 רק אם הכול זהה. כל הבדל — יוצא 1.

   המק״טים מוסתרים כברירת מחדל (PN-01, PN-02…), כדי שפלט שמודבק ל-PR או
   לשיחה לא יישא מק״ט אמיתי. --show-pn מציג אותם — לעיון מקומי בלבד. */
const fs=require('fs');
function compare(A,B,{SHOW=false,MAX=4}={}){
const alias=new Map();const pn=x=>{if(SHOW)return x;if(!alias.has(x))alias.set(x,'PN-'+String(alias.size+1).padStart(2,'0'));return alias.get(x)};
const J=JSON.stringify;const short=v=>{const s=J(v);return s&&s.length>90?s.slice(0,87)+'…':s};
const sections=[];let total=0;
function section(title,key){const s={title,key,lines:[],n:0};sections.push(s);return s}
function add(s,line){s.n++;total++;if(s.lines.length<MAX*6)s.lines.push(line)}

/* ---- פריטים ---- */
{const s=section('פריטים — כל שדה מחושב','items');
 const ka=Object.keys(A.items),kb=Object.keys(B.items);
 const gone=ka.filter(k=>!(k in B.items)),added=kb.filter(k=>!(k in A.items));
 if(gone.length)add(s,`${gone.length} פריטים נעלמו (למשל ${gone.slice(0,3).map(pn).join(', ')})`);
 if(added.length)add(s,`${added.length} פריטים נוספו (למשל ${added.slice(0,3).map(pn).join(', ')})`);
 const byField={};let moved=0;
 for(const k of ka){if(!(k in B.items))continue;const x=A.items[k],y=B.items[k];if(J(x)===J(y))continue;moved++;
  for(const f of new Set([...Object.keys(x),...Object.keys(y)]))if(J(x[f])!==J(y[f])){(byField[f]=byField[f]||[]).push(k)}}
 if(moved){add(s,`${moved} פריטים זזו מתוך ${ka.length}`);
  for(const [f,L] of Object.entries(byField).sort((a,b)=>b[1].length-a[1].length)){
   add(s,`  שדה ${f}: ${L.length} פריטים`);
   for(const k of L.slice(0,MAX))s.lines.push(`     ${pn(k)}: ${short(A.items[k][f])} → ${short(B.items[k][f])}`)}}}
/* ---- תורים ---- */
{const s=section('תורים ודלתות','queues');
 for(const q of new Set([...Object.keys(A.queues),...Object.keys(B.queues)])){const x=A.queues[q],y=B.queues[q];
  if(J(x)===J(y))continue;
  if(!Array.isArray(x)||!Array.isArray(y)){add(s,`${q}: ${short(x)} → ${short(y)}`);continue}
  const sx=new Set(x),sy=new Set(y);const out=x.filter(k=>!sy.has(k)),inn=y.filter(k=>!sx.has(k));
  add(s,`${q}: ${x.length} → ${y.length}${out.length?` · יצאו ${out.length}`:''}${inn.length?` · נכנסו ${inn.length}`:''}${!out.length&&!inn.length?' · אותם פריטים, סדר אחר':''}`)}}
/* ---- מסכים ---- */
{const s=section('מסכים — מונה, פריטים, סדר וסכומים','screens');
 for(const m of new Set([...Object.keys(A.screens),...Object.keys(B.screens)])){const x=A.screens[m]||{},y=B.screens[m]||{};
  if(J(x)===J(y))continue;
  if(x.err||y.err||!A.screens[m]||!B.screens[m]){add(s,`${m}: ${x.err||(!A.screens[m]?'לא היה':'')} → ${y.err||(!B.screens[m]?'נעלם':'')}`);continue}
  if(x.nav!==y.nav)add(s,`${m}: מונה במסילה ${x.nav} → ${y.nav}`);
  if(x.count!==y.count)add(s,`${m}: ${x.count} פריטים → ${y.count}`);
  else if(J(x.pns)!==J(y.pns)){const sx=new Set(x.pns);const same=y.pns.every(k=>sx.has(k));
   const i=x.pns.findIndex((k,j)=>k!==y.pns[j]);add(s,`${m}: ${same?'אותם פריטים, סדר שונה':'פריטים אחרים'} (שורה ראשונה שונה: ${i+1})`)}
  if(J(x.sums)!==J(y.sums))add(s,`${m}: סכומים ${short(x.sums)} → ${short(y.sums)}`)}}
/* ---- ייצוא ---- */
{const s=section('ייצוא — אקסל, SAP, מלאי איטי, זירוז, מייל','exports');
 const cmp=(where,x,y)=>{if(J(x)===J(y))return;
  if(!x||!y||x.err||y.err){add(s,`${where}: ${x?(x.err||'יש קובץ'):'אין קובץ'} → ${y?(y.err||'יש קובץ'):'אין קובץ'}`);return}
  if(x.length!==y.length){add(s,`${where}: ${x.length} קבצים → ${y.length}`);return}
  x.forEach((f,i)=>{const g=y[i];if(f.name!==g.name)add(s,`${where}: שם קובץ «${f.name}» → «${g.name}»`);
   for(const sn of new Set([...Object.keys(f.sheets),...Object.keys(g.sheets)])){const P=f.sheets[sn],Q=g.sheets[sn];
    if(!P||!Q){add(s,`${where} · גיליון «${sn}» ${P?'נעלם':'נוסף'}`);continue}
    if(J(P)===J(Q))continue;
    if(P.length!==Q.length)add(s,`${where} · «${sn}»: ${P.length-1} שורות → ${Q.length-1}`);
    if(J(P[0])!==J(Q[0]))add(s,`${where} · «${sn}»: כותרות העמודות השתנו`);
    let k=0;for(let r=1;r<Math.min(P.length,Q.length)&&k<2;r++)if(J(P[r])!==J(Q[r])){k++;
     const c=P[r].findIndex((v,j)=>J(v)!==J(Q[r][j]));
     /* תא בעמודת מק״ט הוא מק״ט אמיתי — מוסתר כמו מפתחות הפריטים */
     const isPn=/מק״ט|^חומר/.test(String(P[0][c]??''));
     const cell=v=>isPn&&!SHOW&&v!==''&&v!=null?pn(String(v)):short(v);
     add(s,`${where} · «${sn}» שורה ${r+1}, עמודה «${P[0][c]??c+1}»: ${cell(P[r][c])} → ${cell(Q[r][c])}`)}}})};
 for(const m of new Set([...Object.keys(A.exports),...Object.keys(B.exports)])){if(m.startsWith('__'))continue;
  const x=A.exports[m]||{},y=B.exports[m]||{};
  cmp(`${m} · אקסל`,x.excel,y.excel);cmp(`${m} · SAP`,x.sap,y.sap);cmp(`${m} · מלאי איטי`,x.slow,y.slow)}
 cmp('דוח זירוז',A.exports.__expedite,B.exports.__expedite);
 if(J(A.exports.__mail)!==J(B.exports.__mail)){const a=String(A.exports.__mail||'').split('\n'),b=String(B.exports.__mail||'').split('\n');
  const i=a.findIndex((l,j)=>l!==b[j]);add(s,`טיוטת המייל לספקים: ${a.length} שורות → ${b.length}${i>=0?` · שורה ${i+1} שונה`:''}`)}}

const head=[];
if(A.n!==B.n){head.push(`פריטים בדוח: ${A.n} → ${B.n}`);total++}
if(A.now!==B.now)head.push(`⚠ השעון המקובע שונה בין הצילומים (${A.now} / ${B.now}) — השוואה של חישובים תלויי-תאריך אינה הוגנת`);
if((A.errs||[]).length||(B.errs||[]).length)head.push(`⚠ שגיאות JS: ${(A.errs||[]).length} לפני · ${(B.errs||[]).length} אחרי`);
return {sections,total,head,nI:Object.keys(A.items).length,nS:Object.keys(A.screens).length}}
module.exports={compare};
if(require.main===module){
 const args=process.argv.slice(2);
 const SHOW=args.includes('--show-pn');const mi=args.indexOf('--max');const MAX=mi>=0?+args[mi+1]:4;
 const files=args.filter((a,i)=>!a.startsWith('--')&&!(mi>=0&&i===mi+1));
 if(files.length!==2){console.error('שימוש: business-diff.js <before.json> <after.json> [--show-pn] [--max N]');process.exit(2)}
 const [A,B]=files.map(f=>JSON.parse(fs.readFileSync(f,'utf8')));
 const {sections,total,head,nI,nS}=compare(A,B,{SHOW,MAX});
 for(const h of head)console.log(h);
 for(const s of sections){console.log(`\n${s.n?'✗':'✓'} ${s.title}${s.n?` — ${s.n} הבדלים`:' — זהה'}`);for(const l of s.lines)console.log('  '+l);if(s.n>s.lines.length)console.log(`  … ועוד ${s.n-s.lines.length}`)}
 console.log(`\n${nI.toLocaleString('he-IL')} פריטים · ${nS} מסכים · ${total?`${total} הבדלים`:'0 הבדלים — הלוגיקה העסקית והייצוא זהים'}`);
 process.exit(total?1:0)}
