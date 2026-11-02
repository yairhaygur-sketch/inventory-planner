/* ============ D1 · השלד — הפעולות שעברו לתפריטים, במקלדת ובטלפון ============
   D1 העביר פעולות מהשורה אל תפריטים: שבב הקובץ (אזהרות, חודש בסיס, ETA),
   «⋯ עזרה ועוד» (ייצוא, SAP, זירוז, מצב כהה, גיבוי) ו«▦ תצוגה» (מדדים,
   צפיפות, עמודות). הטענה: כל אחת מהן נגישה במקלדת בלבד (פתיחה, מעבר,
   הפעלה, Escape שמחזיר מיקוד לכפתור שפתח) ובנגיעה בטלפון, בבהיר ובכהה,
   וכל תפריט נפתח בתוך המסך. נכתב אחרי שהמדידה מצאה שלושה פערים: Escape
   השאיר aria-expanded=true ומיקוד אבוד, «מדדים»/«צפיפות» לא היו בטווח
   Tab, ותפריט הקובץ יצא מהמסך ב-390. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML=process.argv[2]||path.join(SD,'..','index.html');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null?'  ['+x+']':''))};
const load=async(ctx)=>{await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+HTML+'?nobrief=1');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});
 await p.evaluate(()=>{window.__x=0;const o=XLSX.writeFile;XLSX.writeFile=()=>{window.__x++}});return [p,errs]};
const act=p=>p.evaluate(()=>{const a=document.activeElement;return a.id||a.dataset.help||a.className||a.tagName});
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 /* ---------- מקלדת · מחשב ---------- */
 {const ctx=await b.newContext({viewport:{width:1440,height:900}});const [p,errs]=await load(ctx);
  await p.evaluate(()=>setMode('today'));
  // שבב הקובץ
  await p.focus('#fchip');await p.keyboard.press('Enter');await p.waitForTimeout(150);
  ok('מקלדת · Enter על שבב הקובץ פותח את התפריט',await p.evaluate(()=>document.getElementById('fileMenu').classList.contains('open')&&document.getElementById('fchip').getAttribute('aria-expanded')==='true'));
  const seen=[];for(let i=0;i<8;i++){await p.keyboard.press('Tab');seen.push(await act(p))}
  ok('מקלדת · Tab נכנס לתפריט הקובץ: אזהרות, חודש בסיס, ETA',['dchip','rd','fe'].every(x=>seen.includes(x)),seen.join(','));
  await p.focus('#rd');await p.keyboard.press('Escape');await p.waitForTimeout(150);
  const esc=await p.evaluate(()=>({open:document.getElementById('fileMenu').classList.contains('open'),exp:document.getElementById('fchip').getAttribute('aria-expanded'),f:document.activeElement.id}));
  ok('מקלדת · Escape בתוך תפריט הקובץ סוגר, מעדכן aria-expanded ומחזיר מיקוד לשבב',!esc.open&&esc.exp==='false'&&esc.f==='fchip',JSON.stringify(esc));
  await p.focus('#fchip');await p.keyboard.press('Enter');await p.waitForTimeout(100);
  await p.focus('#dchip');await p.keyboard.press('Enter');await p.waitForTimeout(300);
  ok('מקלדת · Enter על «אזהרות» פותח את הפירוט',await p.evaluate(()=>document.getElementById('diag').getBoundingClientRect().height>0));
  await p.evaluate(()=>{const d=document.getElementById('dchip');if(document.getElementById('diag').getBoundingClientRect().height>0)d.click();document.querySelectorAll('.top .menu.open').forEach(m=>m.classList.remove('open'))});
  // ⋯
  await p.focus('#helpBtn');await p.keyboard.press('Enter');await p.waitForTimeout(150);
  /* UX2: פעולות הקובץ (ייצוא, SAP, מלאי איטי, זירוז) עברו מ«⋯» לכותרת המסך.
     ב«⋯» נשארו 5: עזרה, תצוגות והגדרות. הפעולות שעברו נבדקות למטה, במקלדת, במקומן החדש. */
  const items=[await act(p)];for(let i=0;i<4;i++){await p.keyboard.press('ArrowDown');items.push(await act(p))}
  ok('מקלדת · «⋯» נפתח במיקוד על הפריט הראשון והחצים עוברים על כל 5',items.join(',')==='tour,helpXp,topicsOpen,darkToggle,stBtn',items.join(','));
  // ייצוא במקלדת
  await p.keyboard.press('Escape');await p.waitForTimeout(100);
  ok('מקלדת · Escape סוגר את «⋯» ומחזיר מיקוד',await act(p)==='helpBtn');
  /* UX2: ייצוא וזירוז — בכותרת המסך. בטווח Tab (לא hidden, tabIndex ≥ 0), ו-Enter מפעיל */
  for(const id of ['exportXls','expXls']){
   const x0=await p.evaluate(()=>window.__x);
   const tab=await p.evaluate(i=>{const e=document.getElementById(i);return !!e&&!e.hidden&&e.tabIndex>=0&&e.getBoundingClientRect().width>0&&!e.closest('.menu')},id);
   await p.focus('#'+id);const at=await act(p);await p.keyboard.press('Enter');await p.waitForTimeout(1200);
   ok(`מקלדת · ${id} בכותרת המסך, בטווח Tab, ומופעל ב-Enter`,tab&&at===id&&await p.evaluate(()=>window.__x)>x0,`${at} tab=${tab}`)}
  for(const [id,chk] of [['darkToggle','dark']]){
   const x0=await p.evaluate(()=>({x:window.__x,d:document.body.classList.contains('dark')}));
   await p.focus('#helpBtn');await p.keyboard.press('Enter');await p.waitForTimeout(120);
   for(let i=0;i<10&&await act(p)!==id;i++)await p.keyboard.press('ArrowDown');
   const at=await act(p);await p.keyboard.press('Enter');await p.waitForTimeout(1200);
   const x1=await p.evaluate(()=>({x:window.__x,d:document.body.classList.contains('dark'),open:document.getElementById('helpMenu').classList.contains('open')}));
   ok(`מקלדת · ${id} מגיע בחצים ומופעל ב-Enter, והתפריט נסגר`,at===id&&(chk==='x'?x1.x>x0.x:x1.d!==x0.d)&&!x1.open,`${at} ${JSON.stringify(x1)}`);
   if(chk==='dark')await p.evaluate(()=>document.getElementById('darkToggle').click())}
  // SAP — פותח תצוגה מקדימה, לא קובץ ישר. UX2: מכותרת המסך
  await p.focus('#sapXls');await p.keyboard.press('Enter');await p.waitForTimeout(800);
  const sap=await p.evaluate(()=>({x:window.__x,modal:!![...document.querySelectorAll('[role=dialog],.modal,.sapprev')].find(e=>e.offsetParent!==null)}));
  ok('מקלדת · «עדכון ל-SAP» מופעל מכותרת המסך (תצוגה מקדימה או קובץ)',sap.modal||sap.x>0,JSON.stringify(sap));
  await p.keyboard.press('Escape');await p.waitForTimeout(200);
  await p.evaluate(()=>document.querySelectorAll('[role=dialog]').forEach(d=>{const c=d.querySelector('[data-close],.x,.close');if(c&&d.offsetParent)c.click()}));
  // stBtn
  await p.focus('#helpBtn');await p.keyboard.press('Enter');await p.waitForTimeout(120);
  for(let i=0;i<10&&await act(p)!=='stBtn';i++)await p.keyboard.press('ArrowDown');
  await p.keyboard.press('Enter');await p.waitForTimeout(500);
  ok('מקלדת · «גיבוי והעברה» נפתח מהתפריט',await p.evaluate(()=>!![...document.querySelectorAll('[role=dialog],.modal,.ovl')].find(e=>e.offsetParent!==null&&/גיבוי|העברה/.test(e.textContent))));
  await p.keyboard.press('Escape');await p.waitForTimeout(200);
  // תצוגה
  await p.evaluate(()=>{setMode('month');document.querySelectorAll('.modal,.ovl').forEach(()=>{})});await p.waitForTimeout(200);
  await p.focus('#colsBtn');await p.keyboard.press('Enter');await p.waitForTimeout(200);
  ok('מקלדת · Enter על «תצוגה» פותח את הבורר',await p.evaluate(()=>document.getElementById('colspick').classList.contains('open')));
  const tv=[await act(p)];for(let i=0;i<3;i++){await p.keyboard.press('Tab');tv.push(await act(p))}
  ok('מקלדת · Tab מ«תצוגה» מגיע ל«מדדים» ול«צפיפות»',tv.includes('kpiTog')&&tv.includes('denTog'),tv.join(','));
  await p.focus('#denTog');const d0=await p.evaluate(()=>document.body.classList.contains('dense'));await p.keyboard.press('Enter');await p.waitForTimeout(250);
  ok('מקלדת · Enter על «צפיפות» מחליף צפיפות, והמיקוד נשאר עליו',d0!==await p.evaluate(()=>document.body.classList.contains('dense'))&&await act(p)==='denTog',await act(p));
  await p.keyboard.press('Enter');await p.waitForTimeout(250);
  await p.focus('#kpiTog');const k0=await p.evaluate(()=>LAYOUT.hidden.includes('kpis'));await p.keyboard.press('Enter');await p.waitForTimeout(250);
  ok('מקלדת · Enter על «מדדים» מקפל/פותח את הרצועה, והמיקוד נשאר עליו',k0!==await p.evaluate(()=>LAYOUT.hidden.includes('kpis'))&&await act(p)==='kpiTog',await act(p));
  await p.keyboard.press('Enter');await p.waitForTimeout(250);
  /* עמודה בבורר במקלדת: רווח על תיבה משנה עמודה, והמיקוד נשאר עליה */
  const ck=await p.evaluate(()=>{const c=[...document.querySelectorAll('#colspick input[data-ck]:not([disabled])')].find(x=>!x.checked);if(!c)return null;c.focus();return c.dataset.ck});
  if(ck){await p.keyboard.press('Space');await p.waitForTimeout(300);
   const cs=await p.evaluate(k=>({on:document.querySelector('#colspick input[data-ck="'+k+'"]').checked,f:document.activeElement.dataset.ck}),ck);
   ok('מקלדת · רווח על עמודה בבורר מוסיף אותה, והמיקוד נשאר על התיבה',cs.on&&cs.f===ck,JSON.stringify(cs));
   await p.keyboard.press('Space');await p.waitForTimeout(300);
   const c2=await p.evaluate(k=>({on:document.querySelector('#colspick input[data-ck="'+k+'"]').checked,f:document.activeElement.dataset.ck||document.activeElement.tagName,open:document.getElementById('colspick').classList.contains('open')}),ck);
   ok('מקלדת · רווח נוסף מסיר אותה, והמיקוד עדיין על התיבה',!c2.on&&c2.f===ck&&c2.open,JSON.stringify(c2))}
  await p.keyboard.press('Escape');await p.waitForTimeout(150);
  const ce=await p.evaluate(()=>({open:document.getElementById('colspick').classList.contains('open'),f:document.activeElement.id}));
  ok('מקלדת · Escape סוגר את «תצוגה» ומחזיר מיקוד ל«תצוגה»',!ce.open&&ce.f==='colsBtn',JSON.stringify(ce));
  // סינון + קיבוץ
  await p.evaluate(()=>setMode('today'));await p.waitForTimeout(200);
  await p.focus('#filtBtn');const r0=await p.evaluate(()=>document.body.classList.contains('nofilters'));await p.keyboard.press('Enter');await p.waitForTimeout(350);
  ok('מקלדת · Enter על «סינון» בסרגל הרשימה פותח/סוגר את המסילה',r0!==await p.evaluate(()=>document.body.classList.contains('nofilters')));
  await p.keyboard.press('Enter');await p.waitForTimeout(300);
  await p.focus('#gseg button.on');const g0=await p.evaluate(()=>groupBy);await p.keyboard.press('ArrowLeft');await p.waitForTimeout(500);
  const g1=await p.evaluate(()=>({g:groupBy,f:document.activeElement.dataset.gb,ch:document.querySelector('#gseg [aria-checked=true]').dataset.gb,ti:[...document.querySelectorAll('#gseg button')].map(x=>x.tabIndex).join('')}));
  ok('מקלדת · קיבוץ: חץ מחליף, המיקוד ו-aria-checked עוברים, Tab אחד בקבוצה',g1.g!==g0&&g1.f===g1.g&&g1.ch===g1.g&&/^(0-1|-10)$/.test(g1.ti),JSON.stringify(g1));
  await p.keyboard.press('ArrowLeft');await p.waitForTimeout(400);
  // סדר Tab בכותרת
  await p.evaluate(()=>{document.activeElement.blur();document.querySelector('.top .brand').setAttribute('tabindex','-1');document.querySelector('.top .brand').focus()});
  const ord=[];for(let i=0;i<6;i++){await p.keyboard.press('Tab');ord.push(await act(p))}
  ok('מקלדת · סדר Tab בכותרת: חיפוש → שבב קובץ → העלאה → ⋯',ord.join(',').includes('q')&&ord.indexOf('fchip')<ord.indexOf('helpBtn'),ord.join(','));
  /* ============ UX2 · פעולות הקובץ לפי מסך ============
     אין פעולה שאינה רלוונטית למסך, ולכל פעולה יש בית: «קטלוג פריטים» מציג את כולן. */
  const vis=await p.evaluate(()=>{const o={};for(const m of ['home','today','line','month','cap','capBiz','catalog','prof','burn']){setMode(m);
    o[m]=['exportXls','sapXls','sapSlowXls','expXls','lmailBtn'].filter(i=>{const e=document.getElementById(i);return e&&!e.hidden&&e.getBoundingClientRect().width>0})}return o});
  const has=(m,i)=>vis[m].includes(i);
  ok('UX2 · ייצוא לאקסל בכל מסך',Object.values(vis).every(v=>v.includes('exportXls')),JSON.stringify(vis));
  ok('UX2 · «קטלוג פריטים» מציג את כל פעולות הקובץ',['sapXls','sapSlowXls','expXls'].every(i=>has('catalog',i)),vis.catalog.join(','));
  ok('UX2 · SAP במסכי הפרמטרים, לא ב«בריאות המלאי» ולא במרכז העבודה',has('month','sapXls')&&!has('cap','sapXls')&&!has('home','sapXls'),JSON.stringify({month:vis.month,cap:vis.cap,home:vis.home}));
  ok('UX2 · מלאי איטי ל-SAP ב«בריאות המלאי», לא במסך החוסרים',has('cap','sapSlowXls')&&!has('today','sapSlowXls'),JSON.stringify({cap:vis.cap,today:vis.today}));
  ok('UX2 · דוח זירוז במסכי החוסרים והרכש, לא ב«בריאות המלאי»',has('today','expXls')&&has('line','expXls')&&has('burn','expXls')&&!has('cap','expXls'));
  ok('UX2 · מייל לספקים רק ב«רכש ואספקות»',has('line','lmailBtn')&&Object.entries(vis).filter(([m])=>m!=='line').every(([,v])=>!v.includes('lmailBtn')));
  ok('UX2 · «החלטה מסחרית»: אין קובץ SAP ריק להציע',!has('capBiz','sapXls')&&!has('capBiz','sapSlowXls'),vis.capBiz.join(','));
  await p.evaluate(()=>setMode('line'));await p.waitForTimeout(250);
  await p.focus('#lmailBtn');await p.keyboard.press('Enter');await p.waitForTimeout(300);
  const ml=await p.evaluate(()=>{const b=document.getElementById('lmail');return {open:!!b&&!b.hidden&&b.getBoundingClientRect().height>0,len:b?b.querySelector('textarea').value.length:0,pressed:document.getElementById('lmailBtn').getAttribute('aria-pressed')}});
  ok('UX2 · «מייל לספקים» פותח את הטיוטה, במקלדת',ml.open&&ml.pressed==='true'&&ml.len>0,JSON.stringify(ml));
  ok('UX2 · ברצועה אין עוד Excel ומייל כפולים',await p.evaluate(()=>!document.querySelector('#line #lxlsBtn')&&document.querySelectorAll('#lmailBtn').length===1));
  /* טקסט מוקלד בתיבת החיפוש — ניגודיות מול הסרגל הכהה (היה כהה על כהה) */
  await p.click('#q');await p.keyboard.type('Part');await p.waitForTimeout(300);
  const cr=await p.evaluate(()=>{const L=c=>{const m=c.match(/[\d.]+/g).map(Number);const f=v=>{v/=255;return v<=0.03928?v/12.92:((v+0.055)/1.055)**2.4};return 0.2126*f(m[0])+0.7152*f(m[1])+0.0722*f(m[2])};
    let e=document.getElementById('qbox'),bg='rgba(0, 0, 0, 0)';while(e&&/rgba\(0, 0, 0, 0\)|transparent/.test(bg)){bg=getComputedStyle(e).backgroundColor;e=e.parentElement}
    const fg=getComputedStyle(document.getElementById('q')).color;const a=L(fg),b=L(bg);return {fg,bg,ratio:+((Math.max(a,b)+0.05)/(Math.min(a,b)+0.05)).toFixed(2)}});
  ok('UX2 · טקסט החיפוש בסרגל העליון קריא (ניגודיות ≥ 4.5)',cr.ratio>=4.5,JSON.stringify(cr));
  await p.keyboard.press('Escape');
  /* ============ UX3 · סיכומים ורצועות מקופלים ============ */
  await p.evaluate(()=>{try{localStorage.removeItem('planner_captop_open_v1')}catch(_){};CAPTOP_OPEN=false;setMode('cap')});await p.waitForTimeout(400);
  const c0=await p.evaluate(()=>{const el=document.getElementById('capTop');const b=el.querySelector('.captog');
   return {h:Math.round(el.getBoundingClientRect().height),mini:el.classList.contains('mini'),tiles:el.querySelectorAll('.ctile').length,
    panels:el.querySelectorAll('.cpanel').length,panelsVis:[...el.querySelectorAll('.cpanel')].some(x=>x.getBoundingClientRect().height>0),
    cut:[...el.querySelectorAll('.ctile .cm,.ctile .ct')].filter(e=>e.scrollWidth>e.clientWidth+1).length,aria:b&&b.getAttribute('aria-expanded')}});
  ok('UX3 · «בריאות המלאי»: הסיכום מקופל כברירת מחדל — 4 אריחים בשורה, בלי לוחות הספקים',c0.mini&&c0.tiles===4&&!c0.panelsVis&&c0.h<=100&&c0.aria==='false',JSON.stringify(c0));
  ok('UX3 · ואף סכום באריח אינו נחתך',c0.cut===0,c0.cut+' נחתכו');
  ok('UX3 · לוחות הספקים נשארים בדף — מוסתרים, לא נמחקים',c0.panels>0,c0.panels+' לוחות');
  await p.click('#capTop .captog');await p.waitForTimeout(300);
  const c1=await p.evaluate(()=>{const el=document.getElementById('capTop');return {mini:el.classList.contains('mini'),panelsVis:[...el.querySelectorAll('.cpanel')].some(x=>x.getBoundingClientRect().height>0),aria:el.querySelector('.captog').getAttribute('aria-expanded'),stored:localStorage.getItem('planner_captop_open_v1')}});
  ok('UX3 · לחיצה פותחת את הפירוט, והבחירה נשמרת',!c1.mini&&c1.panelsVis&&c1.aria==='true'&&c1.stored==='1',JSON.stringify(c1));
  await p.evaluate(()=>setMode('today'));await p.waitForTimeout(150);await p.evaluate(()=>setMode('cap'));await p.waitForTimeout(300);
  ok('UX3 · ונשארת פתוחה כשחוזרים למסך',await p.evaluate(()=>!document.getElementById('capTop').classList.contains('mini')));
  await p.click('#capTop .captog');await p.waitForTimeout(200);
  /* רצועת «רכש ואספקות» מקופלת — שורה אחת, והסייג «בלי מחיר» נשאר גלוי (NOPRICE) */
  await p.setInputFiles('#fe',path.join(SD,'zmrp-demo-eta.xlsx'));await p.waitForTimeout(1000);
  /* צעד קודם בבדיקה הסתיר את המדדים (body.nokpi מסתיר גם את הרצועה) — מודדים כשהם גלויים, ומחזירים אחר כך */
  const hadNokpi=await p.evaluate(()=>{const h=document.body.classList.contains('nokpi');document.body.classList.remove('nokpi');return h});
  await p.evaluate(()=>{bandSet(false);setMode('line')});await p.waitForTimeout(400);
  const ln=await p.evaluate(()=>{const b=document.querySelector('#line .lband'),x=document.getElementById('bandx'),n=document.querySelector('.lband .hn'),px=document.querySelector('.lband .nopx');
   const ov=(a,c)=>{if(!a||!c)return false;const r=a.getBoundingClientRect(),q=c.getBoundingClientRect();return !(r.right<=q.left||q.right<=r.left||r.bottom<=q.top||q.bottom<=r.top)};
   return {h:Math.round(b.getBoundingClientRect().height),mini:b.classList.contains('mini'),nopx:!!px&&px.getBoundingClientRect().width>0,overlap:ov(x,n)}});
  ok('UX3 · רצועת «רכש ואספקות» מקופלת בשורה אחת (≤ 60px)',ln.mini&&ln.h>0&&ln.h<=60,JSON.stringify(ln));
  ok('UX3 · והסייג «בלי מחיר» גלוי גם מקופל',ln.nopx);
  ok('UX3 · כפתור הקיפול אינו עולה על המספר',!ln.overlap);
  if(hadNokpi)await p.evaluate(()=>document.body.classList.add('nokpi'));
  /* שם המסך אינו מתקצר באף מסך (תיקון ל-UX2) */
  const tcut=await p.evaluate(()=>Object.keys(MODESET).filter(m=>m!=='moves').filter(m=>{setMode(m);const t=document.getElementById('lt');return t.scrollWidth>t.clientWidth+1}));
  ok('UX3 · שם המסך שלם בכל 25 המסכים, גם כשיש כפתורי ייצוא',!tcut.length,tcut.join(','));
  ok('מחשב · 0 שגיאות JS',!errs.length,errs.join('|'));await ctx.close()}
 /* ---------- טלפון · מגע ---------- */
 for(const dark of [false,true]){const ctx=await b.newContext({viewport:{width:390,height:844},isMobile:true,hasTouch:true});const [p,errs]=await load(ctx);
  const T=dark?'טלפון כהה':'טלפון';if(dark)await p.evaluate(()=>document.body.classList.add('dark'));
  await p.evaluate(()=>setMode('today'));await p.waitForTimeout(200);
  const inV=sel=>p.evaluate(s=>[...document.querySelectorAll(s)].filter(e=>e.offsetParent!==null).map(e=>{const r=e.getBoundingClientRect();return {id:e.id||e.dataset.help,ok:r.left>=0&&r.right<=innerWidth&&r.top>=0&&r.bottom<=innerHeight,h:Math.round(r.height)}}),sel);
  await p.tap('#fchip');await p.waitForTimeout(250);
  const fm=await inV('#fileMenu #dchip:not([hidden]), #fileMenu #rd, #fileMenu label[for=fe]');
  ok(`${T} · תפריט הקובץ נפתח בנגיעה וכל פריטיו בתוך המסך`,fm.length>=2&&fm.every(x=>x.ok),JSON.stringify(fm));
  await p.tap('#dchip');await p.waitForTimeout(300);
  ok(`${T} · נגיעה ב«אזהרות» פותחת את הפירוט`,await p.evaluate(()=>document.getElementById('diag').getBoundingClientRect().height>0));
  await p.evaluate(()=>{if(document.getElementById('diag').getBoundingClientRect().height>0)document.getElementById('dchip').click()});
  await p.tap('#helpBtn');await p.waitForTimeout(250);
  const hm=await inV('#helpMenu [role=menuitem]');
  ok(`${T} · «⋯» נפתח בנגיעה, 5 פריטים (UX2: הייצוא עבר לכותרת המסך), כולם בתוך המסך`,hm.length===5&&hm.every(x=>x.ok),hm.filter(x=>!x.ok).map(x=>x.id).join(',')||hm.length);
  ok(`${T} · גובה פריטי «⋯» ≥ 32px`,hm.every(x=>x.h>=32),hm.map(x=>x.h).join(','));
  const x0=await p.evaluate(()=>window.__x);await p.tap('#exportXls');await p.waitForTimeout(1200);
  ok(`${T} · «ייצוא לאקסל» עובד בנגיעה`,await p.evaluate(()=>window.__x)>x0);
  await p.evaluate(()=>setMode('month'));await p.waitForTimeout(250);
  const cb=await p.evaluate(()=>['colsBtn','colsBtnM'].map(i=>document.getElementById(i)).filter(e=>e&&e.offsetParent!==null).map(e=>e.id));
  ok(`${T} · «תצוגה» גלוי בטלפון`,cb.length>0,cb.join(','));
  if(cb.length){await p.tap('#'+cb[0]);await p.waitForTimeout(250);
   const tv=await inV('#colspick #kpiTog, #colspick #denTog');
   ok(`${T} · «מדדים» ו«צפיפות» בתוך «תצוגה» ובתוך המסך`,tv.length===2&&tv.every(x=>x.ok),JSON.stringify(tv));
   const d0=await p.evaluate(()=>document.body.classList.contains('dense'));await p.tap('#denTog');await p.waitForTimeout(250);
   ok(`${T} · נגיעה ב«צפיפות» עובדת`,d0!==await p.evaluate(()=>document.body.classList.contains('dense')));
   await p.tap('#denTog');await p.evaluate(()=>colsOpen(false))}
  const fb=await inV('#filtBtn');ok(`${T} · «סינון» גלוי בטלפון`,fb.length===1&&fb[0].ok);
  await p.tap('#filtBtn');await p.waitForTimeout(400);
  ok(`${T} · נגיעה ב«סינון» פותחת את המגירה`,await p.evaluate(()=>{const r=document.querySelector('.wfilters');return !!r&&!document.body.classList.contains('nofilters')&&r.getBoundingClientRect().height>100}));
  ok(`${T} · 0 שגיאות JS`,!errs.length,errs.join('|'));await ctx.close()}
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})();
