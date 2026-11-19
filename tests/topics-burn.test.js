/* ============ BURN + STICKY_HEAD (החלטות המעתד 18.11) ============
   «בוער עכשיו» — פס מעל שורת הנושאים, בכל נושא: הזמנות לקוח ללא כיסוי (שני הסוגים),
   אותה רשימה ואותו סדר כמו נושא 1, בלי «טופל», לפי המתכנן. שורת הנושאים (עם הפס)
   קבועה בראש הדף, וכותרת הטבלה נשארת גלויה בגלילה. תצוגה בלבד. */
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

 /* 1 · התוכן: מונה, «בלי רכש», שלושת הדחופים — מול נושא 1 */
 const A=await p.evaluate(()=>{const exp=tpSorted('cust',TOPICS.cust.rows().filter(r=>!tpDone(r)));
  const res={};for(const t of ['cust','trend','short','fix']){TP.tab=t;TP.cat={};tpRender();
   const el=document.querySelector('#lite .tpstick .burn');
   res[t]={has:!!el,txt:el?el.textContent.replace(/\s+/g,' '):'',chips:[...document.querySelectorAll('#lite .burn [data-tp="burnpn"]')].map(x=>x.dataset.key)}}
  TP.view='arr';tpRender();res.arr=!!document.querySelector('#lite .tpstick .burn');TP.view='';tpRender();
  return {res,n:exp.length,none:exp.filter(r=>r.burn).length,top:exp.slice(0,3).map(markKey),
   nopo:[...document.querySelectorAll('#lite .burn .bi')].map(x=>x.classList.contains('nopo')),topBurn:exp.slice(0,3).map(r=>!!r.burn)}});
 ok('הפס מופיע בכל ארבעת הנושאים ובלוח ההגעות',['cust','trend','short','fix'].every(t=>A.res[t].has)&&A.res.arr);
 ok('מונה = הזמנות לקוח ללא כיסוי (נושא 1, בלי «טופל»), ו«בלי רכש» = burn',A.n>0&&A.res.short.txt.includes(`${A.n.toLocaleString('he-IL')} הזמנות ללא כיסוי`)&&A.res.short.txt.includes(`${A.none.toLocaleString('he-IL')} בלי רכש`),A.res.short.txt.slice(0,90));
 ok('שלושת הדחופים = שלושת הראשונים בסדר של נושא 1, זהים בכל נושא',JSON.stringify(A.res.short.chips)===JSON.stringify(A.top)&&['cust','trend','fix'].every(t=>JSON.stringify(A.res[t].chips)===JSON.stringify(A.top)));
 ok('גבול אדום רק לפריט בלי רכש כלל',JSON.stringify(A.nopo)===JSON.stringify(A.topBurn),JSON.stringify(A.nopo));

 /* 2 · לחיצות */
 const C=await p.evaluate(()=>{TP.tab='fix';TP.cat={};tpRender();const k=document.querySelector('#lite .burn [data-tp="burnpn"]').dataset.key;
  document.querySelector('#lite .burn [data-tp="burnpn"]').click();const r1={tab:TP.tab,open:TP.open?markKey(TP.open):null,k};tpdClose();
  TP.tab='short';TP.cat={short:'zero'};tpRender();document.querySelector('#lite .burn [data-tp="burnall"]').click();
  return {...r1,tab2:TP.tab,cat2:TP.cat.cust||''}});
 ok('לחיצה על פריט בפס: עוברים להזמנות לקוח והכרטיס שלו נפתח',C.tab==='cust'&&C.open===C.k,JSON.stringify(C));
 ok('«לכל ההזמנות»: נושא 1, בלי תת-קטגוריה',C.tab2==='cust'&&C.cat2==='');

 /* 3 · «טופל» ומתכנן */
 const D=await p.evaluate(()=>{const first=tpBurnRows()[0],n0=tpBurnRows().length;setMark(first,'handled','');
  const n1=tpBurnRows().length,gone=!tpBurnRows().includes(first);setMark(first,'','');
  const pls=[...new Set(QF.all.map(r=>r.planner).filter(Boolean))];TP.pl=pls[0]||'';
  const okPl=tpBurnRows().every(r=>!TP.pl||r.planner===TP.pl);TP.pl='';tpRender();return {n0,n1,gone,okPl}});
 ok('פריט שסומן «טופל» יוצא מהפס',D.gone&&D.n1===D.n0-1,JSON.stringify(D));
 ok('המתכנן שנבחר חל על הפס',D.okPl);

 /* 4 · קפוא בגלילה: שורת הנושאים בראש, כותרת הטבלה מתחתיה — 1366×768 ו-1920×1080 */
 for(const [w,h] of [[1366,768],[1920,1080]]){await p.setViewportSize({width:w,height:h});
  await p.evaluate(()=>{TP.tab='short';TP.cat={};TP.sort={};tpRender()});
  const S=await p.evaluate(async()=>{const se=document.scrollingElement;se.scrollTop=1e6;document.body.scrollTop=1e6;
   const tw=document.querySelector('#lite .tw');tw.scrollTop=400;await new Promise(r=>setTimeout(r,80));
   const st=document.querySelector('#lite .tpstick').getBoundingClientRect(),th=document.querySelector('#lite thead th').getBoundingClientRect(),
    nav=document.querySelector('#lite nav').getBoundingClientRect(),row=document.querySelector('#lite tbody tr').getBoundingClientRect();
   return {stTop:Math.round(st.top),stBot:Math.round(st.bottom),navTop:Math.round(nav.top),thTop:Math.round(th.top),thBot:Math.round(th.bottom),
    hscroll:document.documentElement.scrollWidth>innerWidth+1,innerScrolled:tw.scrollTop>0,rowHidden:row.bottom<th.bottom}});
  ok(`${w}px: אחרי גלילה — הפס ושורת הנושאים בראש המסך`,S.stTop===0&&S.navTop>=0&&S.navTop<S.stBot,JSON.stringify(S));
  ok(`${w}px: כותרת הטבלה גלויה מתחת לשורת הנושאים`,S.thTop>=S.stBot-1&&S.thTop<=S.stBot+12&&S.innerScrolled,`ראש ${S.stBot} · כותרת ${S.thTop}`);
  ok(`${w}px: אין גלילה אופקית של הדף`,!S.hscroll);
  await p.evaluate(()=>{document.scrollingElement.scrollTop=0;document.body.scrollTop=0})}

 /* 5 · שורת הנושאים בולטת יותר: גופן מודגש וגדול מהטקסט הרגיל */
 const F=await p.evaluate(()=>{const nb=getComputedStyle(document.querySelector('#lite nav button')),bd=getComputedStyle(document.querySelector('#lite main'));
  return {nav:parseFloat(nb.fontSize),base:parseFloat(bd.fontSize),w:+nb.fontWeight}});
 ok('שורת הנושאים: גופן גדול מהטקסט הרגיל ומודגש',F.nav>F.base&&F.w>=600,JSON.stringify(F));

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0);
})().catch(e=>{console.error(e);process.exit(1)});
