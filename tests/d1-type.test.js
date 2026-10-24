/* ============ D1 · טיפוגרפיה וצפיפות ============
   D1 החליף גופן (Heebo), פלטה (D2) והעלה כל טקסט ל-12px לפחות. הבדיקה
   הזאת שומרת על שלושת הדברים שהמעתד קבע כתנאי:
     · אף טקסט גלוי מתחת ל-12px — בכל מסך, ברוחב שולחני ובטלפון, בשני המצבים
     · בלי גלילה אופקית של העמוד
     · הגופן הוא Heebo, והוא באמת נטען (לא נפל לגופן חלופי)
   הצפיפות עצמה (שורות במסך, עמודות חתוכות) נמדדה מול הדוח האמיתי לפני
   המיזוג; כאן נבדק שהיא לא נשברת בדוח הסינתטי: אף עמודה לא נחתכת ב-1920. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 for(const [w,h] of [[1366,768],[390,844],[1920,1080]])for(const dark of [false,true]){
  const ctx=await b.newContext({viewport:{width:w,height:h}});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
  await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0,null,{timeout:180000});
  await p.evaluate(d=>applyDark(d),dark);await p.waitForTimeout(400);
  const r=await p.evaluate(()=>{const small={},over=[],cut=[];
   const vis=e=>{const r=e.getBoundingClientRect();const cs=getComputedStyle(e);return r.width>0&&r.height>0&&cs.visibility!=='hidden'&&cs.display!=='none'};
   for(const m of Object.keys(MODESET)){setMode(m);window.scrollTo(0,0);
    for(const e of document.querySelectorAll('body *')){if(!vis(e))continue;
     if(![...e.childNodes].some(n=>n.nodeType===3&&n.textContent.trim()))continue;
     const fs=parseFloat(getComputedStyle(e).fontSize);if(fs<12)small[(e.className||e.tagName)+'@'+m]=fs}
    if(document.documentElement.scrollWidth>innerWidth+1)over.push(m);
    if(innerWidth>=1900){const box=document.querySelector('#w_queue .rows');const br=box.getBoundingClientRect();
     for(const t of document.querySelectorAll('#tbl thead th')){const r=t.getBoundingClientRect();if(r.width&&(r.left<br.left-1||r.right>br.right+1))cut.push(m+':'+t.textContent.trim().slice(0,10))}}}
   const fam=getComputedStyle(document.body).fontFamily;
   return {small,over,cut,fam,heebo:document.fonts.check('12px "Heebo"')&&[...document.fonts].some(f=>/Heebo/.test(f.family)&&f.status==='loaded')}});
  const tag=`${w}${dark?' כהה':''}`;
  const sm=Object.entries(r.small);
  ok(`${tag} · אף טקסט גלוי מתחת ל-12px, ב-26 המסכים`,!sm.length,sm.slice(0,4).map(([k,v])=>k+'='+v).join(' · '));
  ok(`${tag} · בלי גלילה אופקית של העמוד`,!r.over.length,r.over.join(', '));
  if(w===1920)ok(`${tag} · אף עמודה לא נחתכת`,!r.cut.length,r.cut.slice(0,4).join(', '));
  if(w===1366&&!dark)ok('הגופן הוא Heebo, והוא נטען',/^"?Heebo/.test(r.fam)&&r.heebo,r.fam.slice(0,30));
  ok(`${tag} · אפס שגיאות JS`,!errs.length,errs.slice(0,2).join(' | '));
  await ctx.close()}
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
