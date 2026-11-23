/* ============ TP_PARAM_SPLIT — «עדכון SS ו-ROP» לשלוש (הערת משתמשים 23.11, אושר) ============
   ב«טיוב נתונים» העדכון היה רשימה אחת, מעורבבת. עכשיו שלוש תת-קטגוריות, לפי כיוון השינוי של ROP
   ו-SS מול ההמלצה (אותו זוג שבעמודות): «להעלות ROP/SS» — משהו עולה ושום דבר לא יורד;
   «להוריד ROP/SS» — משהו יורד ושום דבר לא עולה; «מעורב» — כיוונים הפוכים. בלי פער במספרים —
   ב«הכול» בלבד. הנושא, המונה שלו והפעולה של כל פריט — בלי שינוי. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');

 /* 1 · האריחים, הסדר, והחלוקה */
 const A=await p.evaluate(()=>{TP.tab='fix';TP.cat={};TP.sort={};TP.q='';TP.f={};TP.view='';tpRender();
  const tiles=[...document.querySelectorAll('#lite .tiles .tile')].map(t=>({k:t.dataset.tpCat,l:t.querySelector('span').textContent,n:+t.querySelector('.v').textContent.replace(/[^\d]/g,'')}));
  const all=tpRows('fix',true),param=all.filter(r=>tpInFix(r)&&!recNA(r));
  /* מימוש עצמאי של הכיוון — מהזוג שבעמודות */
  const dir=r=>{const q=tpRec(r);if(q.na)return null;const a=q.rop>r.rop?1:q.rop<r.rop?-1:0,b=q.ss>r.ss?1:q.ss<r.ss?-1:0;
   return (a>0&&b<0)||(a<0&&b>0)?'mix':(a>0||b>0)?'up':(a<0||b<0)?'down':'same'};
  const sets={};for(const k of ['pup','pdown','pmix']){TP.cat={fix:k};sets[k]=tpRows('fix',false)}
  TP.cat={};const want={pup:'up',pdown:'down',pmix:'mix'};
  return {tiles,param:param.length,same:param.filter(r=>dir(r)==='same').length,
   right:Object.entries(sets).every(([k,L])=>L.every(r=>dir(r)===want[k])),
   cover:param.every(r=>dir(r)==='same'||Object.values(sets).filter(L=>L.includes(r)).length===1),
   sum:Object.values(sets).reduce((a,L)=>a+L.length,0),n:Object.fromEntries(Object.entries(sets).map(([k,L])=>[k,L.length])),
   tab:+document.querySelector('[data-tp-tab="fix"] .n').textContent.replace(/[^\d]/g,''),total:all.length}});
 ok('האריחים ב«טיוב נתונים»: הכול · שינוי פרופיל MRP · להעלות ROP/SS · להוריד ROP/SS · מעורב · נתונים שבורים',
  JSON.stringify(A.tiles.map(t=>t.l))===JSON.stringify(['הכול','שינוי פרופיל MRP','להעלות ROP/SS','להוריד ROP/SS','מעורב','נתונים שבורים']),A.tiles.map(t=>t.l).join(' · '));
 ok('כל פריט בתת-הקטגוריה הנכונה, לפי הזוג SAP ← מוצע (מימוש עצמאי)',A.right,JSON.stringify(A.n));
 ok('כל פריט עדכון — בדיוק בתת-קטגוריה אחת, וסך שלושתן = כל פריטי העדכון',A.cover&&A.sum+A.same===A.param,`${A.sum}+${A.same}/${A.param}`);
 ok('המונה על האריח = מספר השורות בתת-הקטגוריה',['pup','pdown','pmix'].every(k=>A.tiles.find(t=>t.k===k).n===A.n[k]));
 ok('הנושא כולו לא זז: המונה בלשונית = כל פריטי הנושא',A.tab===A.total,`${A.tab}/${A.total}`);

 /* 2 · בדוח ההדגמה אין «להוריד» ואין «בלי פער» — מזיזים זמנית שני פריטים ובודקים, ומחזירים */
 const B=await p.evaluate(()=>{const param=tpRows('fix',true).filter(r=>tpInFix(r)&&!recNA(r)&&!tpRec(r).na&&!pdUpgrade(r)&&r.rop>0&&r.ss>0);
  const [d,s0]=param,keep=[d,s0].map(r=>({r,rop:r.sugROP,ss:r.sugSS}));
  d.sugROP=d.rop-1;d.sugSS=d.ss-1;s0.sugROP=s0.rop;s0.sugSS=s0.ss;
  const inCat=(k,r)=>{TP.cat={fix:k};return tpRows('fix',false).includes(r)};
  const res={down:inCat('pdown',d)&&!inCat('pup',d)&&!inCat('pmix',d),same:['pup','pdown','pmix'].every(k=>!inCat(k,s0)),sameAll:(TP.cat={},tpRows('fix',false).includes(s0))};
  /* הפעולה בתת-קטגוריה — כמו ב«עדכון SS ו-ROP» */
  TP.cat={fix:'pdown'};res.act=tpTopicAct('fix',d);
  keep.forEach(x=>{x.r.sugROP=x.rop;x.r.sugSS=x.ss});TP.cat={};return res});
 ok('ROP ו-SS שיורדים — «להוריד ROP/SS», ורק שם',B.down);
 ok('בלי פער במספרים — באף אחת משלוש, אבל ב«הכול»',B.same&&B.sameAll);
 ok('הפעולה בתת-קטגוריה: «לעדכן ב-SAP: נק׳ הזמנה ומלאי ביטחון» — כמו ב«עדכון SS ו-ROP»',B.act==='לעדכן ב-SAP: נק׳ הזמנה ומלאי ביטחון',B.act);

 /* 3 · לחיצה על אריח מציגה את תת-הקטגוריה */
 const C=await p.evaluate(()=>{TP.cat={};tpRender();document.querySelector('#lite .tile[data-tp-cat="pup"]').click();
  return {cat:TP.cat.fix,rows:document.querySelectorAll('#lite tbody tr.r').length,n:tpRows('fix',false).length,on:document.querySelector('#lite .tile.on').dataset.tpCat}});
 ok('לחיצה על «להעלות ROP/SS»: האריח מסומן, והטבלה מציגה את שורותיו',C.cat==='pup'&&C.on==='pup'&&C.rows===Math.min(C.n,200)&&C.n>0,`${C.rows}/${C.n}`);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
