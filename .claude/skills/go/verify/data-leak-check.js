/* האם נתון עסקי אמיתי נכנס לריפו?
   לא מול רשימה שמישהו הרכיב — מול הדוח עצמו, כל המק״טים, כל הקבצים.
   המק״טים בדוח עטופים ב-U+200E/U+200F ו-trim() אינו מסיר אותם; בלי
   הניקוי הזה הבדיקה מחזירה «אפס התאמות» שגוי, וזו הטעות הגרועה כאן.
   שימוש:  node data-leak-check.js <report.xlsx> [repoRoot] */
const XLSX=require('xlsx'),fs=require('fs'),cp=require('child_process'),path=require('path');
const CLEAN=s=>String(s==null?'':s).replace(/[‎‏‪-‮⁦-⁩]/g,'').trim();
const xlsx=process.argv[2],root=process.argv[3]||process.cwd();
if(!xlsx){console.error('שימוש: data-leak-check.js <report.xlsx> [repoRoot]');process.exit(2)}
const wb=XLSX.readFile(xlsx);
const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1});
const hdr=(rows[0]||[]).map(CLEAN);
const cDesc=hdr.findIndex(h=>h==='תיאור חומר'), cSup=hdr.findIndex(h=>/ספק/.test(h));
const pns=[],desc=new Set(),sup=new Set();
for(let i=1;i<rows.length;i++){const r=rows[i]||[];
 const v=CLEAN(r[0]);if(v&&v.length>=6)pns.push(v);
 if(cDesc>=0){const d=CLEAN(r[cDesc]);if(d.length>=12)desc.add(d)}
 if(cSup>=0){const s=CLEAN(r[cSup]);if(s.length>=4)sup.add(s)}}
const files=cp.execSync('git ls-files',{cwd:root}).toString().trim().split('\n')
 .filter(f=>/\.(html|js|md|json|yml|yaml|txt|css)$/.test(f));
let blob='';for(const f of files){try{blob+='\n'+fs.readFileSync(path.join(root,f),'utf8')}catch(e){}}
const hitPn=[...new Set(pns.filter(p=>{const b=p.replace(/^M/,'');return b.length>=6&&(blob.includes(p)||blob.includes(b))}))];
const hitDesc=[...desc].filter(d=>blob.includes(d));
const hitSup=[...sup].filter(s=>blob.includes(s));
console.log(`נסרקו ${files.length} קבצים · ${(blob.length/1e6).toFixed(2)}MB · מול ${pns.length} מק״טים בדוח`);
console.log(`מק״טים אמיתיים בריפו:  ${hitPn.length}`+(hitPn.length?'  → '+hitPn.slice(0,25).join(' · '):'  ✓'));
console.log(`תיאורים אמיתיים:        ${hitDesc.length}`+(hitDesc.length?'  → '+hitDesc.slice(0,10).map(d=>'«'+d+'»').join(' '):'  ✓'));
console.log(`שמות ספקים:             ${hitSup.length}`+(hitSup.length?'  → '+hitSup.join(' · '):'  ✓'));
if(hitPn.length||hitSup.length){console.log('\n❌ נתון עסקי בריפו ציבורי. אין לדחוף.');process.exitCode=1}
else console.log('\n✓ נקי');
