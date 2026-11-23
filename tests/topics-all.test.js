/* ============ TP_ALL — לשונית «כל הפריטים» בתצוגת 4 הנושאים (הערת משתמשים 23.11, אושר) ============
   כל פריטי הדוח — גם מה שאינו באף נושא, וגם מה שטופל או נוטרל. חיפוש מק״ט/תיאור, העמודות של «קטלוג
   פריטים» בכלי המלא ועמודת «נושא», מיון וסינון בכותרות, וכרטיס בלחיצה. אינה נושא: לא ב«גם ב…»,
   לא בלוח ההגעות, והמונים של ארבעת הנושאים לא זזים. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx'),ETAX=path.join(SD,'zmrp-demo-eta.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');
 const before=await p.evaluate(()=>TP_REAL.map(k=>tpRows(k,true).length));

 /* 1 · הלשונית */
 const N=await p.evaluate(()=>{const tabs=[...document.querySelectorAll('#lite [data-tp-tab]')].map(b=>b.dataset.tpTab);
  const b=document.querySelector('#lite [data-tp-tab="all"]');return {tabs,txt:b?b.textContent:'',n:b?+b.querySelector('.n').textContent.replace(/[^\d]/g,''):0,all:QF.all.length}});
 ok('לשונית חמישית «כל הפריטים», אחרי ארבעת הנושאים',JSON.stringify(N.tabs)===JSON.stringify(['cust','trend','short','fix','all'])&&/5\. כל הפריטים/.test(N.txt),N.tabs.join(','));
 ok('המונה בלשונית = כל פריטי הדוח',N.n===N.all,`${N.n}/${N.all}`);
 await p.click('#lite [data-tp-tab="all"]');
 const T=await p.evaluate(()=>{const heads=[...document.querySelectorAll('#lite thead th')].map(t=>t.dataset.tpSort);
  return {heads,rows:document.querySelectorAll('#lite tbody tr.r').length,list:TP.list.length,all:QF.all.length,
   outside:QF.all.filter(r=>!tpTopicsOf(r).length).length,done:QF.all.filter(r=>tpDone(r)||isNeutral(r)).length}});
 ok('העמודות: מק״ט · תיאור · סיווג · 11 חודשים · חסר · מלאי פנוי · לקוח · רכש פתוח · כיסוי · שווי עודף · נושא',
  ['pn','desc','cat','spark','miss','free','cust','po','cov','risk','topic'].every(k=>T.heads.includes(k)),T.heads.join(','));
 ok('כל פריטי הדוח ברשימה (מוצגים 200 ראשונים, ועוד בגלילה)',T.list===T.all&&T.rows===Math.min(200,T.all),`${T.rows}/${T.list}/${T.all}`);
 ok('גם פריטים שאינם באף נושא',T.outside>0,String(T.outside));

 /* 2 · חיפוש מוצא פריט שאינו בנושאים */
 const pn=await p.evaluate(()=>QF.all.find(r=>!tpTopicsOf(r).length).pn);
 await p.evaluate(v=>{TP.q=v;tpRender()},pn);
 const S=await p.evaluate(v=>({n:TP.list.length,hit:TP.list.some(r=>r.pn===v),topic:(()=>{const r=TP.list.find(x=>x.pn===v);const i=[...document.querySelectorAll('#lite thead th')].findIndex(t=>t.dataset.tpSort==='topic');const tr=[...document.querySelectorAll('#lite tbody tr.r')][TP.list.indexOf(r)];return tr?tr.children[i].textContent.trim():''})()}),pn);
 ok('חיפוש מק״ט מוצא פריט שאינו באף נושא, ו«נושא» שלו «—»',S.hit&&S.topic==='—',`${S.n} תוצאות · «${S.topic}»`);
 await p.evaluate(()=>{TP.q='';tpRender()});

 /* 3 · עמודת «נושא» = הנושאים שבהם הפריט מופיע */
 const NT=await p.evaluate(()=>{const ok=QF.all.every(r=>{const want=TP_REAL.filter(k=>tpRows(k,true).includes(r)).map(k=>TP_SHORT[k]).join(', ')||null;return TOPICS.all.cols.find(c=>c.k==='topic').v(r)===want});
  return {ok,multi:QF.all.filter(r=>tpTopicsOf(r).length>1).length}});
 ok('«נושא» בכל שורה = בדיוק הנושאים שבהם הפריט מופיע (מימוש עצמאי)',NT.ok,`${NT.multi} פריטים ביותר מנושא אחד`);

 /* 4 · כרטיס */
 const C=await p.evaluate(()=>{const out=QF.all.find(r=>!tpTopicsOf(r).length),inT=QF.all.find(r=>tpTopicsOf(r).length);
  tpdShow(out);const a=document.querySelector('#tpd .db').textContent;tpdClose();
  tpdShow(inT);const b2=document.querySelector('#tpd .q3 .big').textContent,want=tpdProblem(tpTopicsOf(inT)[0],inT);tpdClose();
  return {out:/אינו באף אחד מארבעת הנושאים/.test(a)&&/מצב המלאי/.test(a),inT:b2===want.replace(/<[^>]+>/g,'')}});
 ok('כרטיס של פריט מחוץ לנושאים: «אינו באף אחד מארבעת הנושאים», עם מצב המלאי',C.out);
 ok('כרטיס של פריט שבנושא: כמו בנושא הראשון שלו',C.inT);

 /* 5 · מיון וסינון בכותרות, ו-CSV */
 await p.click('#lite th[data-tp-sort="free"] [data-tp-fil]');await p.fill('#tpcMin','100');await p.click('#tpcOk');
 const F=await p.evaluate(()=>{const L=tpRows('all',false);return {n:L.length,ok:L.every(r=>(r.free||0)>=100),csv:tpSorted('all',L).length}});
 ok('סינון «מלאי פנוי מ-100» בלשונית — כמו בשאר הטבלאות, וה-CSV אותן שורות',F.n>0&&F.ok&&F.csv===F.n,String(F.n));
 await p.evaluate(()=>{TP.cf={};tpRender()});

 /* 6 · מה שלא זז */
 const A=await p.evaluate(()=>({counts:TP_REAL.map(k=>tpRows(k,true).length),also:(TP.tab='cust',tpRender(),[...document.querySelectorAll('#lite .also')].every(e=>!/כל הפריטים/.test(e.title)))}));
 ok('המונים של ארבעת הנושאים לא זזו, ו«גם ב…» לא מזכיר את «כל הפריטים»',JSON.stringify(A.counts)===JSON.stringify(before)&&A.also,A.counts.join(','));
 await p.setInputFiles('#fe',ETAX);await p.waitForTimeout(2500);
 const AR=await p.evaluate(()=>{TP.view='arr';tpRender();const D=tpArrData();const inT=new Set(TP_REAL.flatMap(k=>tpRows(k,true)));TP.view='';return D.every(x=>x.ls.every(e=>inT.has(e.r)&&!e.tk.includes('all')))});
 ok('לוח ההגעות — רק פריטי ארבעת הנושאים, כמו קודם',AR);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
