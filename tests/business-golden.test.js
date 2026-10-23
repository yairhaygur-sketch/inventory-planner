/* ============ שומר הלוגיקה העסקית — הגרסה שרצה ב-CI ============
   הדוח האמיתי לא נכנס לריפו, ולכן ב-CI אין לו מקום. כאן רץ אותו צילום
   עסקי מלא (business-snapshot.js) על הדוח הסינתטי של הבדיקות (900 פריטים,
   ועוד דוח ETA), ומושווה לטביעת אצבע שמורה: business-golden.json.

   מה בטביעת האצבע:
     · כל פריט — גיבוב של כל השדות המחושבים שלו (סיווג, SS/ROP, כמויות,
       כסף, נימוקים, פעולה, המלצת פרופיל)
     · כל תור ודלת — רשימת המק״טים
     · כל אחד מ-26 המסכים — מונה, מספר פריטים, גיבוב הסדר, סכומים לפי מטבע
     · כל קובץ שהכלי מוציא — אקסל, SAP, מלאי איטי לכל מסך; דוח זירוז;
       טיוטת המייל — לכל גיליון: מספר שורות וגיבוב התוכן

   הדוח הסינתטי נבנה כאן מחדש בשעון מקובע ובאזור זמן קבוע, בתיקייה זמנית
   (לא נוגעים ב-tests/zmrp-demo.xlsx של שאר הבדיקות), כדי שאותה ריצה תצא
   זהה היום, מחר וב-CI.

   שינוי מכוון בלוגיקה או בייצוא יפיל את הבדיקה הזו — זה תפקידה. אחרי
   שהשינוי נבדק מקומית מול הדוח האמיתי (business-guard.js) ואושר:
       node tests/business-golden.test.js --update
   ובגוף ה-PR מסבירים מה זז ולמה.

   תקלות שטביעת האצבע קיבעה ותוקנו: ראה KNOWN-ISSUES.md (גיליון «התצוגה»,
   VIEW_SHEET — הגיבובים של «התצוגה» ב-7 מסכים ושל «תנועות» עודכנו אז). */
const {execFileSync}=require('child_process');
const path=require('path'),fs=require('fs'),os=require('os'),crypto=require('crypto');
const ROOT=path.resolve(__dirname,'..'),V=path.join(ROOT,'.claude/skills/go/verify');
const GOLD=path.join(__dirname,'business-golden.json');
const NOW='2026-10-23T08:00:00+03:00',TZ='Asia/Jerusalem';
const UPDATE=process.argv.includes('--update');
let pass=0,fail=0;const ok=(n,c,d)=>{c?pass++:fail++;console.log(`${c?'PASS':'FAIL'} · ${n}${d?'  ['+d+']':''}`)};
const H=v=>crypto.createHash('sha1').update(JSON.stringify(v)).digest('hex').slice(0,16);

/* ---- 1. הדוח הסינתטי, בשעון מקובע ---- */
const tmp=fs.mkdtempSync(path.join(os.tmpdir(),'golden-'));
fs.copyFileSync(path.join(__dirname,'make-fixture.js'),path.join(tmp,'make-fixture.js'));
fs.writeFileSync(path.join(tmp,'freeze.js'),
 `const F=Date.parse(${JSON.stringify(NOW)}),RD=Date;
  global.Date=class extends RD{constructor(...a){a.length?super(...a):super(F)}static now(){return F}};`);
execFileSync(process.execPath,['-r',path.join(tmp,'freeze.js'),path.join(tmp,'make-fixture.js')],
 {cwd:tmp,env:{...process.env,TZ,NODE_PATH:path.join(ROOT,'node_modules'),FIXTURE_SEED:'1'},stdio:'pipe'});
const Z=path.join(tmp,'zmrp-demo.xlsx'),E=path.join(tmp,'zmrp-demo-eta.xlsx');
ok('הדוח הסינתטי נבנה בשעון מקובע',fs.existsSync(Z)&&fs.existsSync(E));

/* ---- 2. צילום עסקי מלא ---- */
const snapF=path.join(tmp,'snap.json');
let snapOk=true;
/* GOLDEN_HTML — להרצת הבדיקה על עותק משונה, כדי להוכיח שהיא נופלת */
const HTML=process.env.GOLDEN_HTML||path.join(ROOT,'index.html');
try{execFileSync(process.execPath,[path.join(V,'business-snapshot.js'),HTML,Z,snapF,'--eta',E,'--now',NOW,'--tz',TZ],{stdio:'inherit',env:process.env})}
catch(e){snapOk=false}
ok('הצילום רץ בלי שגיאת JS',snapOk);
if(!snapOk){console.log(`\n${pass} עברו · ${fail} נכשלו`);process.exit(1)}
const S=JSON.parse(fs.readFileSync(snapF,'utf8'));

/* ---- 3. טביעת אצבע ---- */
const fileDigest=fs=>!fs?null:fs.err?{err:fs.err}:fs.map(f=>({name:f.name,sheets:Object.fromEntries(Object.entries(f.sheets).map(([s,a])=>[s,{rows:a.length-1,h:H(a)}]))}));
const D={n:S.n,
 items:Object.fromEntries(Object.entries(S.items).map(([k,v])=>[k,H(v)])),
 queues:Object.fromEntries(Object.entries(S.queues).map(([k,v])=>[k,Array.isArray(v)?{n:v.length,h:H(v)}:v])),
 screens:Object.fromEntries(Object.entries(S.screens).map(([k,s])=>[k,s.err?{err:s.err}:{nav:s.nav,count:s.count,order:H(s.pns),sums:s.sums}])),
 exports:Object.fromEntries(Object.entries(S.exports).map(([k,x])=>[k,k==='__mail'?{lines:String(x||'').split('\n').length,h:H(x)}:k==='__expedite'?fileDigest(x):{excel:fileDigest(x.excel),sap:fileDigest(x.sap),slow:fileDigest(x.slow)}]))};
ok('26 מסכים בצילום',Object.keys(D.screens).length===26,Object.keys(D.screens).length);
ok('כל מסך נותן ייצוא לאקסל',Object.values(D.exports).filter(x=>x&&x.excel).length===26);
ok('טיוטת המייל לספקים לא ריקה',D.exports.__mail.lines>3,D.exports.__mail.lines+' שורות');
ok('דוח זירוז נוצר',Array.isArray(D.exports.__expedite)&&D.exports.__expedite.length>0);

if(UPDATE){fs.writeFileSync(GOLD,JSON.stringify(D,null,0)+'\n');console.log(`\nנכתב ${path.relative(ROOT,GOLD)} — הסבר ב-PR מה זז ולמה.`);process.exit(fail?1:0)}
if(!fs.existsSync(GOLD)){ok('קיימת טביעת אצבע שמורה',false,'הרץ עם --update');process.exit(1)}
const G=JSON.parse(fs.readFileSync(GOLD,'utf8'));

/* ---- 4. השוואה, לפי תחום ---- */
const J=JSON.stringify;
const cmpMap=(title,a,b)=>{const keys=[...new Set([...Object.keys(a||{}),...Object.keys(b||{})])];
 const diff=keys.filter(k=>J(a[k])!==J(b[k]));ok(title,!diff.length,diff.length?`${diff.length} שונים: ${diff.slice(0,6).join(', ')}${diff.length>6?'…':''}`:`${keys.length} זהים`)};
ok('מספר הפריטים',G.n===D.n,`${G.n} → ${D.n}`);
cmpMap('כל פריט — כל השדות המחושבים',G.items,D.items);
cmpMap('תורים ודלתות',G.queues,D.queues);
for(const m of Object.keys(G.screens)){const a=G.screens[m],b=D.screens[m]||{};
 ok(`מסך ${m}: מונה, פריטים, סדר וסכומים`,J(a)===J(b),J(a)===J(b)?`${b.count??'—'} פריטים`:['nav','count','order','sums','err'].filter(k=>J(a[k])!==J(b[k])).join(', '))}
for(const m of Object.keys(G.exports)){const a=G.exports[m],b=D.exports[m];
 if(m==='__mail'){ok('טיוטת המייל לספקים',J(a)===J(b),`${a.lines} → ${b.lines} שורות`);continue}
 if(m==='__expedite'){ok('דוח זירוז — כל הגיליונות',J(a)===J(b));continue}
 for(const kind of ['excel','sap','slow'])
  /* צד אחד יכול להיות {err} ולא רשימת קבצים — «תנועות» לפני VIEW_SHEET */
  if(J(a[kind])!==J(b&&b[kind])){const bad=[];(Array.isArray(a[kind])?a[kind]:[]).forEach((f,i)=>{const g=(Array.isArray((b||{})[kind])?b[kind]:[])[i];
   for(const s of Object.keys(f.sheets||{}))if(!g||J(f.sheets[s])!==J(g.sheets[s]))bad.push(s)});
   ok(`מסך ${m} · ${({excel:'אקסל',sap:'SAP',slow:'מלאי איטי'})[kind]}`,false,bad.join(', ')||'קובץ נוסף/נעלם')}
 ok(`מסך ${m} · אקסל, SAP ומלאי איטי`,['excel','sap','slow'].every(k=>J(a[k])===J(b&&b[k])))}
fs.rmSync(tmp,{recursive:true,force:true});
console.log(`\n${pass} עברו · ${fail} נכשלו`);
if(fail)console.log('\nלפרטים מלאים — מקומית, מול הדוח האמיתי:\n  node .claude/skills/go/verify/business-guard.js <zmrp.xlsx>');
process.exit(fail?1:0);
