/* ============ D0 · ידיות DOM יציבות — data-row · data-nav · data-count · data-panel · data-bulk ============
   הבדיקות נשענו עד כאן על סלקטורים של עיצוב (`#tbl tbody tr[data-i]`,
   `#tabs .tab .bdg`, `#railsub .railsub2`, `.sapsel`, `#sapbarGo`). כל
   שינוי במראה שבר אותן, ולכן אף סבב עיצוב לא העז לגעת בשלד.

   D0 מוסיף לצד הסלקטורים הקיימים ידיות שאומרות *מה* האלמנט, לא איך
   הוא נראה. הקובץ הזה מוכיח שהן מצביעות בדיוק על אותם אלמנטים — לא
   «בערך»: אותו סט, באותו סדר, בכל 26 המסכים, במחשב ובטלפון. בדיקה
   שעוברת לידית יורשת את הטענה שלה במלואה.

   כשהעיצוב יתחלף (D1 ואילך) הסלקטורים הישנים ייעלמו והידיות יישארו.
   אז השוואת «ידית = סלקטור ישן» תוחלף בהשוואה למקור האמת עצמו —
   ALL / navCount / SAPSEL — שכבר נבדקת כאן במקביל. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,ROOT=path.join(SD,'..');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const XL=process.env.HOOKS_XLSX||path.join(SD,'zmrp-demo.xlsx');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};

/* לכל מסך: האם כל ידית שווה לסלקטור הישן שלה, ומה חסר */
const PROBE=()=>{
 const same=(a,b)=>a.length===b.length&&a.every((e,i)=>e===b[i]);
 const Q=s=>[...document.querySelectorAll(s)];
 const r={};
 const rowsNew=Q('[data-row]'),rowsOld=Q('#tbl tbody tr[data-i], #movtbl tbody tr[data-mi]');
 r.rows=same(rowsNew,rowsOld);r.nRows=rowsNew.length;
 r.rowsTag=rowsNew.every(e=>e.tagName==='TR'&&(e.hasAttribute('data-i')||e.hasAttribute('data-mi')));
 const mainNew=Q('[data-nav-level="main"]'),mainOld=Q('#tabs .tab[data-m]');
 r.main=same(mainNew,mainOld)&&mainNew.every(e=>e.dataset.nav===e.dataset.m);r.nMain=mainNew.length;
 const subNew=Q('[data-nav-level="sub"]'),subOld=Q('#railsub .railsub2[data-m]');
 r.sub=same(subNew,subOld)&&subNew.every(e=>e.dataset.nav===e.dataset.m);
 r.navAll=same(Q('[data-nav]'),[...mainOld,...subOld].sort((a,b)=>a.compareDocumentPosition(b)&4?-1:1));
 const cntNew=Q('[data-count]'),cntOld=Q('#tabs .tab[data-m] .bdg, #railsub .railsub2[data-m] .bdg');
 r.cnt=same(cntNew,cntOld);
 /* המונה בידית = מה שכתוב על התג = navCount. ▤ («תנועות») אינו מספר */
 r.cntBad=cntNew.filter(e=>{const k=e.closest('[data-nav]').dataset.nav,n=+e.dataset.count,
   t=e.textContent.trim();return n!==navCount(k).n||(t!=='▤'&&+t.replace(/[^\d]/g,'')!==n)}).map(e=>e.closest('[data-nav]').dataset.nav);
 r.panel=same(Q('[data-panel="item"]'),Q('#w_detail'));
 r.bulk=same(Q('[data-bulk="bar"]'),Q('#sapbar'))&&same(Q('[data-bulk="select-all"]'),Q('#sapbarAll'))
  &&same(Q('[data-bulk="sap-file"]'),Q('#sapbarGo'))&&same(Q('[data-bulk="clear"]'),Q('#sapbarClr'));
 const pickNew=Q('[data-bulk="pick"]'),pickOld=Q('input.sapsel');
 r.pick=same(pickNew,pickOld);r.nPick=pickNew.length;
 r.pickRow=pickNew.every(e=>e.closest('[data-row]'));
 return r};

(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 for(const [tag,vp,mob] of [['מחשב',{width:1440,height:900},false],['טלפון',{width:390,height:844},true]]){
  const ctx=await b.newContext({viewport:vp,isMobile:mob,hasTouch:mob,deviceScaleFactor:mob?2:1});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+path.join(ROOT,'index.html')+'?nobrief=1');
  await p.setInputFiles('#f',XL);
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB,null,{timeout:180000});
  const modes=await p.evaluate(()=>Object.keys(MODESET));
  const res={};
  for(const m of modes){res[m]=await p.evaluate(([m,P])=>{setMode(m);return eval('('+P+')')()},[m,PROBE.toString()])}
  const all=k=>modes.filter(m=>!res[m][k]);
  ok(`${tag} · data-row = #tbl tr[data-i] + #movtbl tr[data-mi], אותו סט ואותו סדר, ב-${modes.length} מסכים`,!all('rows').length&&!all('rowsTag').length,all('rows').join(','));
  ok(`${tag} · יש שורות לבדוק (האוכלוסייה לא ריקה)`,modes.reduce((a,m)=>a+res[m].nRows,0)>modes.length*5,modes.reduce((a,m)=>a+res[m].nRows,0)+' שורות');
  ok(`${tag} · data-nav-level=main = #tabs .tab[data-m], ו-data-nav = data-m`,!all('main').length&&res[modes[0]].nMain>=7,all('main').join(','));
  ok(`${tag} · data-nav-level=sub = #railsub .railsub2[data-m]`,!all('sub').length,all('sub').join(','));
  ok(`${tag} · [data-nav] = כל יעדי הניווט, בסדר המסמך`,!all('navAll').length,all('navAll').join(','));
  ok(`${tag} · data-count = תגי .bdg של הניווט`,!all('cnt').length,all('cnt').join(','));
  const cb=modes.filter(m=>res[m].cntBad.length).map(m=>m+':'+res[m].cntBad.join('/'));
  ok(`${tag} · המספר בידית = הכתוב על התג = navCount(k).n, בכל מסך`,!cb.length,cb.join(' '));
  ok(`${tag} · data-panel=item = #w_detail`,!all('panel').length,all('panel').join(','));
  ok(`${tag} · data-bulk = #sapbar · #sapbarAll · #sapbarGo · #sapbarClr`,!all('bulk').length,all('bulk').join(','));
  ok(`${tag} · data-bulk=pick = input.sapsel, וכל אחד בתוך שורה`,!all('pick').length&&!all('pickRow').length,all('pick').join(','));
  ok(`${tag} · יש צ׳קבוקסים לבדוק במסך התכנון`,res.month&&res.month.nPick>0,res.month&&res.month.nPick);
  ok(`${tag} · אפס שגיאות JS`,errs.length===0,errs.join(' | '));
  await ctx.close()}
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length} עברו`);process.exit(bad?1:0)
})().catch(e=>{console.error(e);process.exit(1)});
