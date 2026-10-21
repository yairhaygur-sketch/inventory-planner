/* לפני-ואחרי, מק״ט במק״ט, על כל התורים והדלתות.
   אין סכימה בין מטבעות — כל מטבע בשורה משלו (כלל ברזל 6).
   שימוש:  node engine-diff.js <before.json> <after.json> */
const fs=require('fs');
const A=JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
const B=JSON.parse(fs.readFileSync(process.argv[3],'utf8'));
const mb=new Map(B.items.map(r=>[r.pn,r]));
const F=['cat','sev','tier','sugROP','sugSS','orderQty','miss','excessQty','expCap','paramFix','act'];
const trans={},money={};let moved=0,missing=0;
for(const a of A.items){
 const b=mb.get(a.pn);
 if(!b){missing++;continue}
 if(!F.some(k=>String(a[k])!==String(b[k])))continue;
 moved++;
 if(a.cat!==b.cat){const t=a.cat+' → '+b.cat;trans[t]=(trans[t]||0)+1}
 const d=(b.expCap||0)-(a.expCap||0);
 if(d){const c=a.cur||b.cur||'?';money[c]=(money[c]||0)+d}
}
let bad=0;
const cmp=(lbl,oa,ob)=>{
 for(const k of new Set([...Object.keys(oa||{}),...Object.keys(ob||{})])){
  const x=oa&&oa[k],y=ob&&ob[k];
  const lx=Array.isArray(x)?x.length:x, ly=Array.isArray(y)?y.length:y;
  if(String(lx)!==String(ly)){console.log(`  ${lbl} ${k}: ${lx} → ${ly}`);bad++}
 }};
console.log(`פריטים: ${A.n} · זזו: ${moved} · לא נמצאו: ${missing}`);
if(Object.keys(trans).length){console.log('\nמעברי סיווג:');
 for(const k of Object.keys(trans).sort((x,y)=>trans[y]-trans[x]))console.log(`  ${trans[k]}×  ${k}`)}
console.log('\nתורים/דלתות/מצבים שזזו:');
cmp('תור',A.queues,B.queues);cmp('דלת',A.doors,B.doors);cmp('מצב',A.modes,B.modes);
if(!bad)console.log('  אין — כולם זהים');
console.log('\nכסף תקוע (בלי סכימה בין מטבעות):');
const mk=Object.keys(money);
if(!mk.length)console.log('  0 בכל מטבע');
else for(const c of mk)console.log(`  ${c}: ${money[c]>0?'+':''}${Math.round(money[c]).toLocaleString('he-IL')}`);
process.exitCode=0;
