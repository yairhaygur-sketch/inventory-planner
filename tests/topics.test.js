/* ============ TOPICS — התצוגה המצומצמת אינה מנוע שני ============
   index.html?topics מציג ארבעה נושאים: טיוב נתוני תכנון · חריגת צריכה ·
   הזמנות לקוח ללא כיסוי · מלאי 0 או צפוי להיגמר. היא נולדה מהשוואה מול
   דשבורד ניסויי שענה על אותן ארבע שאלות במנוע משלו, ונמצאו בו כללים
   שונים לאותה שאלה ושתי הוראות הפוכות לאותו פריט.

   לכן הבדיקה הזאת נועלת דבר אחד מעל הכול: כל שורה וכל מונה בתצוגה הם
   בדיוק מה שהמנוע של הכלי המלא כבר בחר — אותה קבוצת מק״טים, לא רק אותו
   מספר. ובנוסף: הכלי המלא אינו משתנה בלי ?topics, כסף מוצג במטבע של
   הפריט, הייצוא הוא התצוגה, ואין גלילה אופקית של העמוד.
   התצוגה מיועדת למחשב בלבד: הבדיקה רצה ב-1366 · 1440 · 1920, בלי טלפון.
   ההשוואה העיקרית רצה בין שני עמודים נפרדים: הכלי המלא (בלי ?topics)
   מחשב את הנושאים מהבוררים של המנוע, והתצוגה (?topics) נטענת בנפרד על
   אותו קובץ. כל שדה שהתצוגה מציגה או מייצאת חייב לצאת זהה, מק״ט במק״ט.
   TOPICS_XLSX=<דוח> — על דוח אחר, מקומית. */
const {chromium}=require('playwright');
const fs=require('fs'),path=require('path');
const SD=__dirname,ROOT=path.join(SD,'..');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const XL=process.env.TOPICS_XLSX||path.join(SD,'zmrp-demo.xlsx');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
const URL0='file://'+path.join(ROOT,'index.html');
const os=require('os'),XLSX=require('xlsx');
/* מה המנוע אומר על כל נושא — רץ בכל אחד משני העמודים. אותם בוררים שהמסכים
   בכלי המלא משתמשים בהם, ואותם שדות שהתצוגה מציגה או מייצאת. */
const ENGINE=()=>{const uni=(...ls)=>[...new Set(ls.flat())];
 const pick=r=>({pn:r.pn,desc:r.desc,planner:r.planner,mrp:r.mrp,rop:r.rop,sugROP:r.sugROP,ss:r.ss,sugSS:r.sugSS,
  ssSrc:r.ssSrc,ltD:r.ltD,srv:r.srvMissing?null:r.srv,rate:r.rate,price:r.priceMissing?null:r.price,currency:r.currency,
  tRec:r.tRec,tPre:r.tPre,tRatio:r.tRatio,sugROPT:r.sugROPT,sugROPU:r.sugROPU,cust:r.cust,free:r.free,po:r.po,
  unsQty:r.unsQty,burn:!!r.burn,custGap:!!r.custGap,cat:r.cat,sev:r.sev,act:(r.act||[])[0]||'',
  why:(r.why||[]).filter(Boolean).map(w=>Array.isArray(w)?w[1]:String(w))});
 const by=l=>Object.fromEntries(l.map(r=>[r.pn,pick(r)]));
 return {fix:by(decisionList('fix')),
  trend:by(uni(riseRows(),trendRows(),QF.all.filter(r=>r.cat==='ייתכן שאזל ולא נרשם'))),
  cust:by(QF.all.filter(r=>isT1(r)||r.custGap)),
  short:by(uni(decisionList('short'),decisionList('m1')))}};
/* יחס ניגודיות WCAG בין צבע הטקסט לרקע האפקטיבי (עולים בעץ עד רקע אטום) */
const CONTRAST=sel=>{const rgb=s=>(s.match(/[\d.]+/g)||[]).map(Number);
 const lum=([r,g,b])=>{const f=c=>{c/=255;return c<=.03928?c/12.92:Math.pow((c+.055)/1.055,2.4)};return .2126*f(r)+.7152*f(g)+.0722*f(b)};
 const bg=el=>{for(let e=el;e;e=e.parentElement){const c=rgb(getComputedStyle(e).backgroundColor);if(c.length>=3&&(c[3]===undefined||c[3]>0.9))return c}
  return rgb(getComputedStyle(document.body).backgroundColor)};
 const out=[];for(const el of document.querySelectorAll(sel)){if(!el.offsetParent||!el.textContent.trim())continue;
  const a=lum(rgb(getComputedStyle(el).color)),c=lum(bg(el));const r=(Math.max(a,c)+.05)/(Math.min(a,c)+.05);
  out.push({sel:el.className||el.tagName,t:el.textContent.trim().slice(0,20),r:Math.round(r*100)/100})}
 return out};

(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const errs=[];let REF=null;
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
  REF=await p.evaluate(ENGINE);
  await p.context().close()}

 /* ---- 1א. הדרך מהכלי המלא: כפתור במסך הפתיחה ופריט בתפריט «⋯» ----
    בלי הקישור, שתי הכתובות היו קיימות והמעתד לא מצא את 4 הנושאים. */
 {const p=await page(1440,900,'?nobrief=1');
  await p.waitForSelector('#welTopics',{state:'visible'});
  await Promise.all([p.waitForURL(/\?topics$/),p.click('#welTopics')]);
  await p.waitForSelector('#lite h1');
  ok('מסך הפתיחה: «תכנון מלאי — 4 נושאים» פותח את ?topics',
   await p.evaluate(()=>document.documentElement.classList.contains('lite')&&/4 נושאים/.test(document.querySelector('#lite h1').textContent)));
  await p.goto(URL0+'?nobrief=1');await load(p);
  await p.click('#helpBtn');
  const [np]=await Promise.all([p.context().waitForEvent('page'),p.click('#topicsOpen')]);
  await np.waitForLoadState();await np.waitForSelector('#lite h1');
  const r=await np.evaluate(()=>({url:location.search,lite:document.documentElement.classList.contains('lite')}));
  const still=await p.evaluate(()=>!!(QF&&QF.all&&QF.all.length)&&!document.documentElement.classList.contains('lite'));
  ok('תפריט «⋯»: «תכנון מלאי — 4 נושאים» נפתח בלשונית חדשה, והכלי המלא נשאר עם הדוח',
   r.url==='?topics'&&r.lite&&still,JSON.stringify({...r,still}));
  await p.context().close()}

 /* ---- 2. ?topics — מצב ריק, ואז טעינה ---- */
 const p=await page(1440,900,'?topics');
 {const r=await p.evaluate(()=>({shell:getComputedStyle(document.getElementById('shell')).display,
   empty:!!document.querySelector('#lite .empty'),load:!!document.querySelector('#lite [data-tp="load"]'),
   brief:!document.getElementById('brief')||getComputedStyle(document.getElementById('brief')).display==='none'}));
  ok('?topics לפני טעינה: מצב ריק עם כפתור העלאה, והכלי המלא מוסתר',r.shell==='none'&&r.empty&&r.load&&r.brief,JSON.stringify(r))}
 await load(p);
 await p.waitForSelector('#lite table');

 /* ---- 3א. העמוד הנפרד = הכלי המלא, שדה בשדה ---- */
 {const MINE=await p.evaluate(()=>Object.fromEntries(Object.keys(TOPICS).map(t=>[t,TOPICS[t].rows().map(r=>r.pn)])));
  const here=await p.evaluate(ENGINE);
  for(const t of ['fix','trend','cust','short']){const a=REF[t],b=here[t],ids=Object.keys(a).sort();
   const sameIds=JSON.stringify(ids)===JSON.stringify(Object.keys(b).sort())&&JSON.stringify(ids)===JSON.stringify([...MINE[t]].sort());
   const diff=ids.filter(pn=>JSON.stringify(a[pn])!==JSON.stringify(b[pn]));
   ok(`${t}: מול הכלי המלא בעמוד נפרד — אותם ${ids.length} מק״טים ואותם ערכים בכל שדה`,sameIds&&!diff.length&&ids.length>0,
    sameIds?diff.slice(0,3).join(', '):'קבוצת המק״טים שונה')}}
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
 const parse=l=>{const o=[];let c='',q=false;for(let i=0;i<l.length;i++){const ch=l[i];
  if(q){if(ch==='"'&&l[i+1]==='"'){c+='"';i++}else if(ch==='"')q=false;else c+=ch}
  else if(ch==='"')q=true;else if(ch===','){o.push(c);c=''}else c+=ch}o.push(c);return o};
 /* עמודות ה-CSV שנבדקות מול הכלי המלא: כותרת → שדה */
 const CSVF={'מק״ט':'pn','נק הזמנה SAP':'rop','נק הזמנה מוצעת':'sugROP','מלאי ביטחון SAP':'ss','מלאי ביטחון מוצע':'sugSS',
  'הזמנת לקוח':'cust','מלאי פנוי':'free','רכש פתוח':'po','3 ח׳ אחרונים':'tRec','3 ח׳ קודמים':'tPre','נק׳ הזמנה מוצעת':'sugROP',
  'מטבע':'currency','פעולה':'act','מתכנן':'planner','סוג MRP':'mrp'};
 for(const t of ['fix','trend','cust','short']){
  await p.click(`[data-tp-tab="${t}"]`);
  const [dl]=await Promise.all([p.waitForEvent('download'),p.click('#lite [data-tp="csv"]')]);
  const raw=fs.readFileSync(await dl.path(),'utf8');const lines=raw.replace(/^﻿/,'').split('\r\n');
  const n=await p.evaluate(()=>+String(document.querySelector('#lite .bar .mini').textContent).replace(/[^\d]/g,''));
  ok(`${t}: CSV — BOM לעברית, כותרת, ושורה לכל פריט בתצוגה`,raw.charCodeAt(0)===0xfeff&&lines.length===n+1&&/מק״ט/.test(lines[0]),`${lines.length-1} · ${n}`);
  const H=parse(lines[0]),bad2=[];
  for(const l of lines.slice(1)){const c=parse(l),pn=c[H.indexOf('מק״ט')],e=REF[t][pn];
   if(!e){bad2.push(pn+': לא בכלי המלא');continue}
   H.forEach((h,i)=>{const f=CSVF[h];if(!f)return;const v=e[f]==null?'':String(e[f]);if(c[i]!==v)bad2.push(`${pn} ${h}: ${c[i]}≠${v}`)})}
  ok(`${t}: כל ערך ב-CSV = הערך בכלי המלא`,!bad2.length,bad2.slice(0,3).join(' · '))}

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
 for(const [w,h] of [[1366,768],[1440,900],[1920,1080]]){
  const q=await page(w,h,'?topics');await load(q);await q.waitForSelector('#lite table');
  const bad1=[];
  for(const t of ['fix','trend','cust','short']){await q.click(`[data-tp-tab="${t}"]`);
   const m=await q.evaluate(()=>({page:document.documentElement.scrollWidth-innerWidth,
     tw:(t=>t.scrollWidth-t.clientWidth)(document.querySelector('#lite .tw'))}));
   if(m.page>0||(w>=1366&&m.tw>0))bad1.push(`${t}: עמוד ${m.page} · טבלה ${m.tw}`)}
  ok(`${w}px: אין גלילה אופקית${w>=1366?' ואין טבלה חתוכה':''}`,!bad1.length,bad1.join(' · '));
  await q.context().close()}

 /* ---- 9. בהיר וכהה: ניגודיות ו-RTL ----
    התצוגה מיועדת למחשב בלבד (החלטת המעתד, 27.10.2026) — אין כאן מדידות טלפון. */
 for(const [w,h] of [[1366,768],[1920,1080]])for(const dark of [false,true]){
  const q=await page(w,h,'?topics');await load(q);await q.waitForSelector('#lite table');
  if(dark!==await q.evaluate(()=>document.body.classList.contains('dark')))await q.click('#lite [data-tp="dark"]');
  /* החלפת ערכת הצבעים מונפשת — מודדים אחרי שהמעבר הסתיים, לא באמצעו */
  await q.waitForTimeout(1500);
  const tag=`${w}px ${dark?'כהה':'בהיר'}`,low=[];
  for(const t of ['fix','trend','cust','short']){await q.click(`[data-tp-tab="${t}"]`);await q.click('#lite tr.r');
   for(const c of await q.evaluate(CONTRAST,'#lite h1,#lite .src,#lite nav button,#lite .lead,#lite th,#lite td,#lite .p,#lite .tile div,#lite .mini,#lite .tph .btn,#lite tr.det li'))
    if(c.r<4.5)low.push(`${t} ${c.sel} «${c.t}» ${c.r}`)}
  ok(`${tag}: ניגודיות טקסט ≥ 4.5 בכל ארבעת הנושאים`,!low.length,[...new Set(low)].slice(0,3).join(' · '));
  const r=await q.evaluate(()=>{const ths=[...document.querySelectorAll('#lite th')];
   return {dir:getComputedStyle(document.getElementById('lite')).direction,
    first:ths[0].textContent.trim(),right:ths[0].getBoundingClientRect().right>=ths[1].getBoundingClientRect().right,
    page:document.documentElement.scrollWidth-innerWidth}});
  ok(`${tag}: RTL — העמודה הראשונה (מק״ט) בצד ימין, ואין גלילה אופקית`,r.dir==='rtl'&&/מק/.test(r.first)&&r.right&&r.page<=0,JSON.stringify(r));
  await q.context().close()}

 /* ---- 10. מק״ט עם גרש, מירכאות ו-HTML — לא שובר את הלחיצה ולא מוזרק ---- */
 {const tmp=fs.mkdtempSync(path.join(os.tmpdir(),'topics-'));
  const wb=XLSX.readFile(XL),ws=wb.Sheets[wb.SheetNames[0]],aoa=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:''});
  const hi=aoa.findIndex(r=>r.includes('מק"ט מוביל')),H=aoa[hi],col=n=>H.indexOf(n);
  const row=aoa[hi+1];const PN=`O'K"<b>X</b>`;
  row[col('מק"ט מוביל')]=PN;row[col('מלאי פנוי')]=0;row[col('כמות בהז.פ')]=7;row[col('הז. רכש')]=0;
  if(col('סה"כ מלאי')>=0)row[col('סה"כ מלאי')]=0;
  const wb2=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb2,XLSX.utils.aoa_to_sheet(aoa),wb.SheetNames[0]);
  const f2=path.join(tmp,'odd.xlsx');XLSX.writeFile(wb2,f2);
  const q=await page(1440,900,'?topics');await q.setInputFiles('#f',f2);await q.waitForSelector('#lite table');
  await q.click('[data-tp-tab="cust"]');await q.fill('#lite [data-tp="q"]',"O'K");await q.waitForTimeout(600);
  const r0=await q.evaluate(PN=>{const tr=[...document.querySelectorAll('#lite tr.r')].find(t=>t.firstElementChild.textContent===PN);
   return {found:!!tr,injected:!!document.querySelector('#lite td b b')}},PN);
  await q.click('#lite tr.r');
  const r1=await q.evaluate(()=>!!document.querySelector('#lite tr.det'));
  ok('מק״ט עם גרש/מירכאות/HTML: מוצג כטקסט, ולחיצה פותחת את ההסבר',r0.found&&!r0.injected&&r1,JSON.stringify({...r0,open:r1}));
  await q.context().close();fs.rmSync(tmp,{recursive:true,force:true})}

 ok('אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
