/* ============ ZMC_TOPICS — «לבדוק» גם בתצוגת 4 הנושאים (החלטת המעתד 21.11) ============
   מועמד PD ← ZM-C שהסימולציה מצאה בו סיבה (profSplit → rev) מוצג «לבדוק», עם הסיבה — כמו
   במסך «שינוי סוג MRP» של הכלי המלא (ZMC_PR2). בעמודה «שינוי פרופיל», בפעולה של «טיוב
   נתונים», בכרטיס וב-CSV. המונים, הפריטים בנושא והמצבים — בלי שינוי. */
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

 /* 1 · העמודה «שינוי פרופיל» בנושא «טיוב נתונים», תת-קטגוריה «שינוי פרופיל MRP» */
 const C=await p.evaluate(()=>{TP.tab='fix';TP.cat={fix:'prof'};TP.sort={};TP.q='';TP.f={};tpRender();
  const ths=[...document.querySelectorAll('#lite thead th')],i=ths.findIndex(t=>t.dataset.tpSort==='prof');
  const trs=[...document.querySelectorAll('#lite tbody tr')],L=TP.list||[];
  const rows=trs.map((tr,j)=>({r:L[j],t:i<0?null:tr.children[i].textContent.trim(),ttl:i<0?'':((tr.children[i].querySelector('span')||{}).title||'')}));
  const rev=rows.filter(x=>x.r&&profInfo(x.r)&&profInfo(x.r).state==='up'&&profSplit(x.r).grp==='rev');
  return {i,n:rows.length,all:rows.every(x=>x.r&&x.t===profStatus(x.r).t&&x.ttl===(profStatus(x.r).why||'')),
   rev:rev.length,revOk:rev.every(x=>x.t==='לבדוק'&&x.ttl===profSplit(x.r).reasons.map(w=>PFSIM_WHY[w]).join(' · ')),
   upOk:rows.filter(x=>x.r&&profInfo(x.r).state==='up'&&profSplit(x.r).grp==='ok').every(x=>x.t==='להעלות'),
   key:rev[0]&&markKey(rev[0].r)}});
 ok('העמודה «שינוי פרופיל» = הסטטוס של הכלי המלא (profStatus), עם הסיבה בבועית',C.i>=0&&C.n>0&&C.all,`${C.n} שורות`);
 ok('מועמד שהסימולציה מצאה בו סיבה — «לבדוק», והסיבות שלו',C.rev>0&&C.revOk,`${C.rev} «לבדוק»`);
 ok('מועמד תקין — «להעלות», כמו קודם',C.upOk);

 /* 2 · הפעולה, ה-CSV והכרטיס של פריט «לבדוק» */
 const D=await p.evaluate(k=>{const r=QF.all.find(x=>markKey(x)===k);TP.cat={fix:'prof'};
  const act=tpTopicAct('fix',r),col=TOPICS.fix.cols.find(c=>c.k==='prof'),csv=col.x[0][1](r);
  const ac=TOPICS.fix.cols.find(c=>c.k==='act'),csvAct=((ac&&ac.x||[]).find(x=>x[0]==='פעולה בנושא')||[,()=>null])[1](r);
  tpdShow(r);const card=(document.getElementById('tpd')||{}).textContent||'';tpdClose();
  return {act,csv,csvAct,card:card.replace(/\s+/g,' '),why:profStatus(r).why}},C.key);
 ok('הפעולה: «לבדוק פרופיל: PD ← ZM-C»',D.act==='לבדוק פרופיל: PD ← ZM-C',D.act);
 ok('ה-CSV: «PD ← ZM-C · לבדוק»',D.csv==='PD ← ZM-C · לבדוק',D.csv);
 ok('ה-CSV, «פעולה בנושא»: «לבדוק פרופיל: PD ← ZM-C»',D.csvAct==='לבדוק פרופיל: PD ← ZM-C',D.csvAct);
 ok('הכרטיס: «לבדוק — לא לעדכן פרופיל אוטומטית» וכל הסיבות',D.card.includes('לבדוק — לא לעדכן פרופיל אוטומטית')&&D.card.includes(D.why),D.why);

 /* 3 · מה שלא זז: מספר הפריטים בתת-הקטגוריה = המונה של «שינוי פרופיל» בכלי המלא */
 const N=await p.evaluate(()=>{TP.cat={fix:'prof'};tpRender();return {n:(TP.list||[]).length,eng:navCount('prof').n}});
 ok('«שינוי פרופיל MRP» = המונה במסילת הכלי המלא, כמו קודם',N.n===N.eng,`${N.n}/${N.eng}`);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
