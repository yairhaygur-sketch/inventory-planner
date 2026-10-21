/* תחליף ל-lint/tsc שאינם קיימים בפרויקט: בדיקת תחביר על כל קובץ JS,
   ועל הסקריפט המוטמע ב-index.html. אין כאן טענה שזה lint — זה
   המינימום שאפשר לבדוק, והיעדר lint נאמר במפורש ולא מוסתר.
   שימוש:  node syntax-check.js [repoRoot] */
const fs=require('fs'),cp=require('child_process'),path=require('path'),os=require('os');
const root=process.argv[2]||process.cwd();
let bad=0,n=0;
for(const f of cp.execSync('git ls-files "*.js"',{cwd:root}).toString().trim().split('\n').filter(Boolean)){
 n++;try{cp.execSync(`node --check ${JSON.stringify(path.join(root,f))}`,{stdio:'pipe'})}
 catch(e){console.log('  ❌ '+f+': '+String(e.stderr||e).split('\n').slice(0,3).join(' ').slice(0,200));bad++}}
const html=path.join(root,'index.html');
if(fs.existsSync(html)){
 const src=fs.readFileSync(html,'utf8');
 const blocks=[...src.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m=>m[1]);
 blocks.forEach((b,i)=>{n++;
  const tmp=path.join(os.tmpdir(),`idx-block-${i}.js`);
  fs.writeFileSync(tmp,b);
  try{cp.execSync(`node --check ${JSON.stringify(tmp)}`,{stdio:'pipe'})}
  catch(e){console.log(`  ❌ index.html script#${i}: `+String(e.stderr||e).split('\n').slice(0,3).join(' ').slice(0,200));bad++}
  finally{fs.unlinkSync(tmp)}});
}
console.log(`תחביר: ${n} יחידות נבדקו · ${bad} כשלים`);
console.log('הערה: בפרויקט הזה אין eslint, אין prettier ואין tsconfig. זו בדיקת תחביר, לא lint.');
process.exitCode=bad?1:0;
