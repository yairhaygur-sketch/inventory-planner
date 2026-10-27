/* ============ TOPICS — התצוגה המצומצמת אינה מנוע שני ============
   index.html?topics מציג ארבעה נושאים: טיוב נתוני תכנון · חריגת צריכה ·
   הזמנות לקוח ללא כיסוי · מלאי 0 או צפוי להיגמר. היא נולדה מהשוואה מול
   דשבורד ניסויי שענה על אותן ארבע שאלות במנוע משלו, ונמצאו בו כללים
   שונים לאותה שאלה ושתי הוראות הפוכות לאותו פריט.

   לכן הבדיקה הזאת נועלת דבר אחד מעל הכול: כל שורה וכל מונה בתצוגה הם
   בדיוק מה שהמנוע של הכלי המלא כבר בחר — אותה קבוצת מק״טים, לא רק אותו
   מספר. ובנוסף: הכלי המלא אינו משתנה בלי ?topics, כסף מוצג במטבע של
   הפריט, הייצוא הוא התצוגה, ואין גלילה אופקית של העמוד.
   TOPICS_XLSX=<דוח> — על דוח אחר, מקומית. */
const {chromium}=require('playwright');
const fs=require('fs'),path=require('path');
const SD=__dirname,ROOT=path.join(SD,'..');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const XL=process.env.TOPICS_XLSX||path.join(SD,'zmrp-demo.xlsx');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
const URL0='file://'+path.join(ROOT,'index.html');

(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const errs=[];
 const page=async(w,h,q)=>{const ctx=await b.newContext({viewport:{width:w,height:h},acceptDownloads:true});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();p.on('pageerror',e=>errs.push(e.message.split('\n')[0]));
  await p.goto(URL0+q);return p};
 const load=async p=>{await p.setInputFiles('#f',XL);
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0,null,{timeout:180000});
  await p.waitForTimeout(500)};

 /* ---- 1. בלי ?topics — הכלי המלא כמו שהוא ---- */
 {const p=await page(1440,900,'?nobrief=1');await load(p);
  const r=await p.evaluate(()=>({lite:document.documentElement.classList.contains('lite'),
   liteShown:getComputedStyle(document.getElementById('lite')).display!=='none',
   liteHtml:document.getElementById('lite').innerHTML.length,
   shell:getComputedStyle(document.getElementById('shell')).display!=='none'}));
  ok('בלי ?topics: התצוגה המצומצמת כבויה והכלי המלא מוצג',!r.lite&&!r.liteShown&&r.liteHtml===0&&r.shell,JSON.stringify(r));
  await p.context().close()}

 /* ---- 2. ?topics — מצב ריק, ואז טעינה ---- */
 const p=await page(1440,900,'?topics');
 {const r=await p.evaluate(()=>({shell:getComputedStyle(document.getElementById('shell')).display,
   empty:!!document.querySelector('#lite .empty'),load:!!document.querySelector('#lite [data-tp="load"]'),
   brief:!document.getElementById('brief')||getComputedStyle(document.getElementById('brief')).display==='none'}));
  ok('?topics לפני טעינה: מצב ריק עם כפתור העלאה, והכלי המלא מוסתר',r.shell==='none'&&r.empty&&r.load&&r.brief,JSON.stringify(r))}
 await load(p);
 await p.waitForSelector('#lite table');

 /* ---- 3. כל נושא = הבורר של המנוע, מק״ט במק״ט ---- */
 const eng=await p.evaluate(()=>{const ids=l=>l.map(r=>r.pn).sort();
  const uni=(...ls)=>[...new Set(ls.flat())].sort();
  return {
   fix:ids(decisionList('fix')),
   trend:uni(ids(riseRows()),ids(trendRows()),ids(QF.all.filter(r=>r.cat==='ייתכן שאזל ולא נרשם'))),
   cust:ids(QF.all.filter(r=>isT1(r)||r.custGap)),
   short:uni(ids(decisionList('short')),ids(decisionList('m1'))),
   nav:{burn:navCount('burn').n,trend:navCount('trend').n,rise:navCount('rise').n},
   custGap:QF.all.filter(r=>r.custGap).length}});
 const T={};
 for(const t of ['fix','trend','cust','short']){
  await p.click(`[data-tp-tab="${t}"]`);
  T[t]=await p.evaluate(t=>{const gn=s=>+String(s).replace(/[^\d]/g,'');return {ids:TOPICS[t].rows().map(r=>r.pn).sort(),
   nav:gn(document.querySelector(`[data-tp-tab="${t}"] .n`).textContent),
   label:gn(document.querySelector('#lite .bar .mini').textContent),
   trs:document.querySelectorAll('#lite tr.r').length,
   tiles:[...document.querySelectorAll('#lite .tile')].map(x=>({k:x.dataset.tpCat,n:gn(x.querySelector('.v').textContent)}))}},t);
  const same=JSON.stringify(T[t].ids)===JSON.stringify(eng[t]);
  ok(`${t}: השורות הן בדיוק בחירת המנוע (${eng[t].length} מק״טים)`,same&&eng[t].length>0,
   same?'':`תצוגה ${T[t].ids.length} · מנוע ${eng[t].length}`);
  ok(`${t}: המונה בלשונית = אורך הרשימה = «N שורות»`,T[t].nav===eng[t].length&&T[t].label===eng[t].length,
   `${T[t].nav} · ${T[t].label} · ${eng[t].length}`);
  ok(`${t}: מוצגות עד 200 שורות`,T[t].trs===Math.min(200,eng[t].length),String(T[t].trs));
  /* כל אריח, כשלוחצים עליו, מציג בדיוק את המספר שעליו */
  const miss=[];
  for(const tl of T[t].tiles){await p.click(`#lite .tile[data-tp-cat="${tl.k}"]`);
   const n=await p.evaluate(()=>+String(document.querySelector('#lite .bar .mini').textContent).replace(/[^\d]/g,''));
   if(n!==tl.n)miss.push(`${tl.k||'הכול'}: ${tl.n}→${n}`)}
  await p.click('#lite .tile[data-tp-cat=""]');
  ok(`${t}: כל אריח פותח את המספר שעליו`,!miss.length,miss.join(', '))}
 /* המונים של התצוגה ושל המסילה בכלי המלא — אותו מספר */
 {const tl=(t,k)=>(T[t].tiles.find(x=>x.k===k)||{}).n;
  ok('«לקוח ממתין · אין רכש» = «אין רכש כלל» בכלי המלא',tl('cust','burn')===eng.nav.burn,`${tl('cust','burn')} · ${eng.nav.burn}`);
  ok('«ביקוש יורד» = הדלת בכלי המלא',tl('trend','down')===eng.nav.trend,`${tl('trend','down')} · ${eng.nav.trend}`);
  ok('«ביקוש עולה» = הדלת בכלי המלא',tl('trend','up')===eng.nav.rise,`${tl('trend','up')} · ${eng.nav.rise}`);
  ok('אין רכש + ההזמנה גדולה = כל ההזמנות שהמנוע סימן כלא מכוסות',
   tl('cust','burn')+tl('cust','gap')===eng.custGap,`${tl('cust','burn')}+${tl('cust','gap')} · ${eng.custGap}`)}

 /* ---- 4. טיוב נתונים: בכל שורה משהו שונה מ-SAP, וכסף במטבע של הפריט ---- */
 await p.click('[data-tp-tab="fix"]');
 {const r=await p.evaluate(()=>{const rows=TOPICS.fix.rows();
   const cells=[...document.querySelectorAll('#lite tr.r')].map(tr=>tr.children);
   const ki=TP_HEAD.concat(TOPICS.fix.cols).findIndex(c=>c.k==='val');
   const wrong=[];TP.shown.forEach((row,i)=>{const t=cells[i][ki].textContent.trim();
    if(t!=='—'&&!t.includes(curSym(row.currency).trim()))wrong.push(row.pn+':'+t)});
   return {same:rows.filter(r=>r.sugROP===r.rop&&r.sugSS===r.ss).length,wrong}});
  ok('טיוב נתונים: אין שורה שבה ROP ו-SS זהים ל-SAP',r.same===0,String(r.same));
  ok('שווי השינוי מוצג במטבע של הפריט עצמו',!r.wrong.length,r.wrong.slice(0,3).join(', '))}

 /* ---- 5. מיון, פירוט, חיפוש, מתכנן ---- */
 {await p.click('#lite th[data-tp-sort="rop"]');
  const v=await p.evaluate(()=>TP.shown.map(r=>r.sugROP));
  ok('מיון לפי נקודת הזמנה — יורד בלחיצה ראשונה',v.every((x,i)=>!i||v[i-1]>=x),v.slice(0,5).join(','));
  await p.click('#lite th[data-tp-sort="rop"]');
  const u=await p.evaluate(()=>TP.shown.map(r=>r.sugROP));
  ok('לחיצה שנייה — עולה',u.every((x,i)=>!i||u[i-1]<=x),u.slice(0,5).join(','))}
 {await p.click('#lite tr.r');
  const d=await p.evaluate(()=>{const det=document.querySelector('#lite tr.det');const r=TP.shown[0];
   return {open:!!det,txt:det?det.textContent:'',why:(r.why||[]).filter(Boolean).map(w=>Array.isArray(w)?w[1]:w)}});
  ok('לחיצה על שורה פותחת את ההסבר של המנוע',d.open&&d.why.length>0&&d.why.every(w=>d.txt.includes(w)),String(d.why.length));
  await p.click('#lite tr.r');
  ok('לחיצה שנייה סוגרת',await p.evaluate(()=>!document.querySelector('#lite tr.det')))}
 {const pn=await p.evaluate(()=>TP.shown[1].pn);
  await p.fill('#lite [data-tp="q"]',pn);await p.waitForTimeout(600);
  const r=await p.evaluate(()=>({n:TP.shown.length,pn:TP.shown.map(x=>x.pn),focus:document.activeElement&&document.activeElement.dataset.tp}));
  ok('חיפוש מק״ט מוצא אותו, והמיקוד נשאר בתיבה',r.pn.includes(pn)&&r.n>=1&&r.focus==='q',`${r.n} · ${r.focus}`);
  await p.fill('#lite [data-tp="q"]','');await p.waitForTimeout(600)}
 {const pl=await p.evaluate(()=>[...document.querySelectorAll('#lite [data-tp="pl"] option')].map(o=>o.value).filter(Boolean)[0]);
  await p.selectOption('#lite [data-tp="pl"]',pl);
  const r=await p.evaluate(pl=>({all:TOPICS[TP.tab].rows().filter(r=>r.planner===pl).length,
   lab:+String(document.querySelector('#lite .bar .mini').textContent).replace(/[^\d]/g,'')}),pl);
  ok('סינון מתכנן = השורות של אותו מתכנן',r.all===r.lab,`${r.lab} · ${r.all}`);
  await p.selectOption('#lite [data-tp="pl"]','')}

 /* ---- 6. הייצוא הוא התצוגה ---- */
 for(const t of ['fix','cust']){
  await p.click(`[data-tp-tab="${t}"]`);
  const [dl]=await Promise.all([p.waitForEvent('download'),p.click('#lite [data-tp="csv"]')]);
  const txt=fs.readFileSync(await dl.path(),'utf8');const lines=txt.replace(/^﻿/,'').split('\r\n');
  const n=await p.evaluate(()=>+String(document.querySelector('#lite .bar .mini').textContent).replace(/[^\d]/g,''));
  ok(`${t}: קובץ CSV — כותרת + שורה לכל פריט בתצוגה`,lines.length===n+1&&/מק״ט/.test(lines[0]),`${lines.length-1} · ${n}`)}

 /* ---- 7. מקלדת הכלי המלא כבויה, ומצב כהה ---- */
 {const before=await p.evaluate(()=>({mode:typeof mode!=='undefined'?mode:null,tab:TP.tab}));
  await p.click('#lite main .lead');for(const k of ['ArrowLeft','ArrowRight','Enter','d'])await p.keyboard.press(k);
  const after=await p.evaluate(()=>({mode:typeof mode!=='undefined'?mode:null,tab:TP.tab}));
  ok('קיצורי המקלדת של הכלי המלא אינם פועלים על המסך המוסתר',JSON.stringify(before)===JSON.stringify(after),JSON.stringify(after));
  const d0=await p.evaluate(()=>document.body.classList.contains('dark'));
  await p.click('#lite [data-tp="dark"]');
  const d1=await p.evaluate(()=>document.body.classList.contains('dark'));
  await p.click('#lite [data-tp="dark"]');
  ok('מצב כהה / בהיר מתחלף',d0!==d1)}
 await p.context().close();

 /* ---- 8. רוחב: אין גלילה אופקית לעמוד; ב-1366 גם הטבלה נכנסת ---- */
 for(const [w,h] of [[1366,768],[1920,1080],[390,844]]){
  const q=await page(w,h,'?topics');await load(q);await q.waitForSelector('#lite table');
  const bad1=[];
  for(const t of ['fix','trend','cust','short']){await q.click(`[data-tp-tab="${t}"]`);
   const m=await q.evaluate(()=>({page:document.documentElement.scrollWidth-innerWidth,
     tw:(t=>t.scrollWidth-t.clientWidth)(document.querySelector('#lite .tw'))}));
   if(m.page>0||(w>=1366&&m.tw>0))bad1.push(`${t}: עמוד ${m.page} · טבלה ${m.tw}`)}
  ok(`${w}px: אין גלילה אופקית${w>=1366?' ואין טבלה חתוכה':''}`,!bad1.length,bad1.join(' · '));
  await q.context().close()}

 ok('אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
