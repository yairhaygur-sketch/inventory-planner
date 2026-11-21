/* ============ ZMC_PR2 — «שינוי פרופיל MRP»: שני טאבים, «לבדוק» בהעלאות, גיליון «התצוגה» (אושר 21.11) ============
   · «העלאה ל-ZM-C»: כל מועמדי PD ← ZM-C. מועמד שהסימולציה מצאה בו סיבה (profSplit → rev) מוצג
     «לבדוק» והסיבה ב«למה» — לא נשאר PD בשקט ולא נעלם. העמודות, בסדר שאושר: מק״ט · דגם · תיאור ·
     המלצה · למה · צריכה נטו 12 ח׳ · חודשי צריכה · ROP אחרי ZM-C · SS אחרי ZM-C · מלאי · בדרך.
   · «הורדות וממתינים»: כל השאר, כמו קודם ובלי עמודות הסימולציה.
   · גיליון «התצוגה»: הטאב הפתוח, אותן עמודות ואותו סדר; ROP/SS אחרי ZM-C עם «(סימולציה)».
   · המצבים עצמם והמונה במסילה לא השתנו. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 const wait=ms=>p.waitForTimeout(ms);
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.evaluate(()=>{try{localStorage.clear()}catch(_){}});
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB,null,{timeout:180000});
 await wait(600);
 await p.evaluate(()=>setMode('prof'));await wait(600);
 const snap=()=>p.evaluate(()=>{const rows=profTabRows(currentRows()),trs=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
  const by=new Map(QF.all.map(r=>[String(r.pn),r]));
  const cell=(tr,k)=>{const td=tr.querySelector(`td[data-ck="${k}"]`);return td?td.textContent.trim():null};
  const R=trs.map(tr=>{const r=by.get((tr.querySelector('.obj')||{}).textContent.trim());return {r,st:cell(tr,'profSt'),why:cell(tr,'profWhy')}});
  const v=viewExportRows();
  return {tab:PROF_TAB,view:colViewKey(),heads:[...document.querySelectorAll('#tbl thead th[data-ck]')].map(t=>t.dataset.ck),
   tabs:[...document.querySelectorAll('#track .proftab')].map(b=>({k:b.dataset.proftab,on:b.classList.contains('on'),sel:b.getAttribute('aria-selected'),n:+b.querySelector('.fn').textContent.replace(/[^\d]/g,'')})),
   n:rows.length,trs:trs.length,
   allUp:rows.every(r=>profInfo(r).state==='up'),noneUp:rows.every(r=>profInfo(r).state!=='up'),
   stOk:R.every(x=>{if(!x.r)return false;const i=profInfo(x.r);if(i.state!=='up')return x.st===PROF_ST[i.state][0]&&x.why===i.why;
    const sp=profSplit(x.r);return sp.grp==='rev'?(x.st==='לבדוק'&&x.why===sp.reasons.map(w=>PFSIM_WHY[w]).join(' · ')):(x.st==='להעלות'&&x.why===i.why)}),
   revAfterOk:(()=>{const g=rows.map(r=>profInfo(r).state==='up'?profSplit(r).grp:'');const i=g.indexOf('rev');return i<0||g.slice(i).every(x=>x==='rev')})(),
   counts:R.reduce((a,x)=>(a[x.st]=(a[x.st]||0)+1,a),{}),
   sheet:v&&{keys:v.keys,head:v.aoa[0],n:v.n,pns:v.aoa.slice(1).map(x=>String(x[0]))},pns:rows.map(r=>String(r.pn))}});

 /* 1 · ברירת המחדל: «העלאה ל-ZM-C» */
 const A=await snap();
 const all=await p.evaluate(()=>{const L=profRows();return {n:L.length,up:L.filter(r=>profInfo(r).state==='up').length,
  rev:L.filter(r=>profInfo(r).state==='up'&&profSplit(r).grp==='rev').length,nav:navCount('prof').n,act:L.filter(profActive).length,
  upChk:L.filter(r=>['up','down','check'].includes(profInfo(r).state)).length}});
 ok('שני טאבים, «העלאה ל-ZM-C» פתוח; המספר בכל טאב = השורות שבו',A.tabs.length===2&&A.tabs[0].k==='up'&&A.tabs[0].on&&A.tabs[0].sel==='true'
  &&A.tabs[0].n===all.up&&A.tabs[1].n===all.n-all.up,JSON.stringify(A.tabs));
 ok('העלאות: כל מועמדי PD ← ZM-C, ורק הם',A.tab==='up'&&A.view==='prof'&&A.allUp&&A.n===all.up&&A.trs===A.n,`${A.n}`);
 ok('העמודות בסדר שאושר',JSON.stringify(A.heads)===JSON.stringify(['pnPlain','model','desc','profSt','profWhy','profNet','profH','profSimRop','profSimSs','free','poPlain']),A.heads.join(','));
 ok('«להעלות» או «לבדוק» לפי הסימולציה; ב«לבדוק» — הסיבות ב«למה»',A.stOk&&(A.counts['לבדוק']||0)===all.rev&&all.rev>0,JSON.stringify(A.counts));
 ok('«להעלות» קודם, «לבדוק» אחריהם',A.revAfterOk);
 ok('המונה במסילה לא השתנה (להעלות + להוריד + לבדוק — רכש פתוח)',all.nav===all.act&&all.act===all.upChk,`${all.nav}`);
 /* הטאב חל על הטבלה ועל «התצוגה» בלבד: רשימת המסך (currentRows), גיליון «כל הנתונים» וקובץ ה-SAP — כל השורות, כמו קודם */
 const M=await p.evaluate(()=>{window.__cap=[];const o=XLSX.writeFile;XLSX.writeFile=(wb,name)=>{const sh={};for(const sn of wb.SheetNames)sh[sn]=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1}).length-1;window.__cap.push(sh)};
  try{doExport()}finally{XLSX.writeFile=o}const sh=window.__cap[0]||{};const main=Object.keys(sh).find(n=>/^כל הנתונים/.test(n));
  return {cur:currentRows().length,all:profRows().length,main:main?sh[main]:null,view:sh['התצוגה']}});
 ok('רשימת המסך וגיליון «כל הנתונים» — כל השורות של שני הטאבים, כמו קודם',M.cur===all.n&&M.main===all.n,`${M.cur}/${M.main}/${all.n}`);
 ok('גיליון «התצוגה» — רק הטאב הפתוח',M.view===all.up,`${M.view}/${all.up}`);
 ok('גיליון «התצוגה»: הטאב הפתוח, אותן עמודות ואותו סדר',A.sheet&&JSON.stringify(A.sheet.keys)===JSON.stringify(A.heads)&&A.sheet.n===A.n&&JSON.stringify(A.sheet.pns)===JSON.stringify(A.pns),A.sheet&&A.sheet.n);
 ok('ROP/SS אחרי ZM-C בגיליון — עם «(סימולציה)» בכותרת',A.sheet.head.includes('ROP אחרי ZM-C (סימולציה)')&&A.sheet.head.includes('SS אחרי ZM-C (סימולציה)'),A.sheet.head.slice(7,9).join(' | '));
 const N=await p.evaluate(()=>{const tr=document.querySelector('#tbl tbody tr[data-i]');const r=QF.all.find(x=>String(x.pn)===tr.querySelector('.obj').textContent.trim());
  return {cell:tr.querySelector('td[data-ck="profNet"]').textContent.trim(),want:r.net12==null?'—':r.net12.toLocaleString('he-IL')}});
 ok('«צריכה נטו 12 ח׳» = net12',N.cell===N.want,`${N.cell}/${N.want}`);

 /* 2 · «הורדות וממתינים» */
 await p.click('#track [data-proftab="down"]');await wait(500);
 const D=await snap();
 ok('לחיצה עוברת ל«הורדות וממתינים»',D.tab==='down'&&D.view==='profDn'&&D.tabs[1].on&&D.tabs[1].sel==='true');
 ok('הורדות: כל השאר, ורק הם',D.noneUp&&D.n===all.n-all.up&&D.trs===D.n,`${D.n}`);
 ok('העמודות כמו קודם, בלי הסימולציה',JSON.stringify(D.heads)===JSON.stringify(['pnPlain','model','desc','profCur','profTgt','profH','profSt','profWhy','free','poPlain']),D.heads.join(','));
 ok('ההמלצה וה«למה» כמו קודם',D.stOk,JSON.stringify(D.counts));
 ok('גיליון «התצוגה»: הטאב הזה, בלי עמודות הסימולציה',JSON.stringify(D.sheet.keys)===JSON.stringify(D.heads)&&D.sheet.n===D.n&&!D.sheet.head.some(h=>/סימולציה/.test(h)));
 await p.click('#track [data-proftab="up"]');await wait(400);
 ok('וחזרה לטאב ההעלאות',(await snap()).tab==='up');

 /* 3 · נוסח הסיבה ss_zero_rate3y */
 ok('«ROP≥2 ו-SS=0» אומר מתי זה קורה',await p.evaluate(()=>/אין צריכה ב-11 החודשים/.test(PFSIM_WHY.ss_zero_rate3y)&&/3 השנים/.test(PFSIM_WHY.ss_zero_rate3y)));

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
