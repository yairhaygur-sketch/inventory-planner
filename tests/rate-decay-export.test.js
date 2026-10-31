/* ============ RATE_DECAY — מהדוח ועד קובץ SAP והאקסל ============
   rate-decay.test.js בודק את הכלל ב-build(). כאן — המסלול שהמשתמש עובר:
   טעינת דוח דרך #f, המנוע, קובץ העדכון ל-SAP (doSapMass, שני הקבצים),
   וייצוא האקסל (doExport) — ומה שבאמת נכתב בהם.

   הדוח: הסינתטי ועוד ארבעה פריטים (tests/lib/decay-fixture.js), שניים
   מהם בדעיכה חדה. נבדק:
     · ROP/SS שהמנוע חישב — כצפוי, וזהה למה שנכתב בקובץ ובאקסל
     · ssMin בקובץ SAP = הערך מהדוח, בלי שינוי
     · בכל שורה של כל קובץ SAP, בכל 26 המסכים: SS ≥ ssMin */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const XLSX=require('xlsx');
const {ITEMS,writeDecayReport}=require('./lib/decay-fixture.js');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{
 const XL=writeDecayReport();
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB,null,{timeout:180000});
 /* כל קובץ שנכתב — נלכד כטבלה, כמו שהוא היה יוצא */
 await p.evaluate(()=>{window.__cap=[];window.__xf=XLSX.writeFile;
  XLSX.writeFile=(wb,name)=>{window.__cap.push({name,sheets:wb.SheetNames.map(n=>[n,XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:null})])})};
  window.Worker=undefined});

 /* 1 · המנוע */
 const E=await p.evaluate(pns=>Object.fromEntries(ALL.filter(x=>pns.includes(x.pn)).map(x=>[x.pn,
  {rop:x.sugROP,ss:x.sugSS,ssMin:x.ssMin,dec:!!x.rateDecay,rate:x.rate,rateROP:x.rateDecay?x.rateROP:x.rate,dLT:x.dLT,ltM:x.ltM,paramFix:!!x.paramFix,cat:x.cat}])),ITEMS.map(i=>i.pn));
 ok('ארבעת הפריטים נטענו',ITEMS.every(i=>E[i.pn]),Object.keys(E).join(', '));
 for(const i of ITEMS){const e=E[i.pn];if(!e)continue;
  if(i.decay)ok(`${i.pn}: דעיכה — ROP ${i.expROP}, SS ${i.expSS}`,e.dec===true&&e.rop===i.expROP&&e.ss===i.expSS,
   `קצב ${e.rate.toFixed(2)} → ל-ROP ${e.rateROP!=null?e.rateROP.toFixed(3):'—'} · ROP ${e.rop} · SS ${e.ss}`);
  else ok(`${i.pn}: אין דעיכה — ROP/SS על הקצב הקיים, כמו קודם`,e.dec===false&&Math.abs(e.dLT-e.rate*e.ltM)<1e-9&&e.rop===Math.ceil(e.dLT+e.ss-1e-9),
   `ROP ${e.rop} · SS ${e.ss}`);
  ok(`${i.pn}: ssMin = הערך מהדוח, ו-SS ≥ ssMin`,e.ssMin===i.ssMin&&e.ss>=e.ssMin,`ssMin ${e.ssMin} · SS ${e.ss}`)}

 /* 2 · קובצי SAP — כל המסכים, שני הקבצים */
 const S=await p.evaluate(()=>{const res=[];
  for(const m of Object.keys(MODESET)){try{setMode(m)}catch(_){continue}
   try{currentRows()}catch(_){continue}
   for(const kind of [undefined,'slow']){window.__cap.length=0;doSapMass(kind);
    for(const f of window.__cap)res.push({m,kind:kind||'params',name:f.name,rows:f.sheets[0][1]})}}
  return res});
 const files=S.filter(f=>f.rows.length>1);
 let rows=0,viol=[],hdrBad=[];const seen={};
 for(const f of S){const h=f.rows[0]||[];const iS=h.indexOf('מלאי בטחון'),iM=h.indexOf('מלאי ביטחון מינימלי'),iR=h.indexOf('נקודת הזמנה חוזרת'),iP=h.indexOf('חומר ');
  if(iS<0||iM<0||iR<0||iP<0){hdrBad.push(f.m+'/'+f.kind);continue}
  for(const r of f.rows.slice(1)){rows++;if(!((r[iS]||0)>=(r[iM]||0)))viol.push(`${f.m}/${f.kind}: ${r[iP]} SS ${r[iS]} < ${r[iM]}`);
   if(String(r[iP]).startsWith('DECAY-'))(seen[r[iP]]=seen[r[iP]]||[]).push({ss:r[iS],min:r[iM],rop:r[iR],where:f.m+'/'+f.kind})}}
 ok('כל קובץ SAP נושא את ארבע העמודות, בשמן המדויק',!hdrBad.length,hdrBad.slice(0,4).join(', '));
 ok('בכל שורה של כל קובץ SAP, בכל המסכים: SS ≥ ssMin',rows>0&&!viol.length,
   viol.length?viol.slice(0,3).join(' · '):`${files.length} קבצים · ${rows.toLocaleString('he-IL')} שורות`);
 /* הפריט בדעיכה שהכלי מציע לתקן ב-SAP חייב להופיע — בערכים שהמנוע חישב */
 const fixItems=ITEMS.filter(i=>i.decay&&E[i.pn]&&E[i.pn].paramFix);
 ok('לפחות פריט אחד בדעיכה נכנס לקובץ SAP',fixItems.length>0&&fixItems.every(i=>seen[i.pn]),fixItems.map(i=>i.pn).join(', ')||'אין');
 for(const i of fixItems){const all=seen[i.pn]||[];const e=E[i.pn];
  ok(`${i.pn} בקובץ SAP: SS ${e.ss}, ssMin ${i.ssMin} (ללא שינוי), ROP ${e.rop}`,
   all.length>0&&all.every(x=>x.ss===e.ss&&x.min===i.ssMin&&x.rop===e.rop),all.slice(0,2).map(x=>`${x.where}: ${x.ss}/${x.min}/${x.rop}`).join(' · '))}
 /* פריט שאינו «לתקן ב-SAP» לא נכנס לקובץ — גם כשה-ROP שלו ירד */
 for(const i of ITEMS.filter(i=>E[i.pn]&&!E[i.pn].paramFix))
  ok(`${i.pn}: אינו «לתקן ב-SAP» (${E[i.pn].cat}) — ולכן אינו בקובץ`,!seen[i.pn]);

 /* 3 · ייצוא האקסל — הגיליון «כל הנתונים · כל הפריטים» */
 const X=await p.evaluate(()=>{setMode('catalog');window.__cap.length=0;doExport();return window.__cap.slice()});
 const sh=X.length&&X[0].sheets.find(s=>/כל הנתונים/.test(s[0]));
 ok('האקסל נכתב, עם גיליון «כל הנתונים»',!!sh,X.length?X[0].sheets.map(s=>s[0]).join(' | '):'לא נכתב');
 if(sh){const t=sh[1];const hi=t.findIndex(r=>r&&r.includes('ROP מוצע'));const h=t[hi]||[];
  const col=k=>h.indexOf(k);const ip=h.findIndex(c=>/מק״ט|מק"ט/.test(String(c)));
  const need=['ROP מוצע','SS מוצע','ביקוש בזמן האספקה','רצפת ssMin מ-SAP'];
  ok('העמודות קיימות באקסל',ip>=0&&need.every(k=>col(k)>=0),need.filter(k=>col(k)<0).join(', '));
  for(const i of ITEMS){const e=E[i.pn];const r=t.find(r=>r&&r[ip]===i.pn);
   ok(`${i.pn} באקסל: ROP ${e&&e.rop} · SS ${e&&e.ss} · ssMin ${i.ssMin} · ביקוש בזמן אספקה ${e&&e.dLT.toFixed(2)}`,
    !!r&&r[col('ROP מוצע')]===e.rop&&r[col('SS מוצע')]===e.ss&&r[col('רצפת ssMin מ-SAP')]===i.ssMin&&r[col('ביקוש בזמן האספקה')]===+e.dLT.toFixed(2),
    r?`${r[col('ROP מוצע')]} · ${r[col('SS מוצע')]} · ${r[col('רצפת ssMin מ-SAP')]} · ${r[col('ביקוש בזמן האספקה')]}`:'לא נמצא')}}
 ok('אין שגיאות JS',!errs.length,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
