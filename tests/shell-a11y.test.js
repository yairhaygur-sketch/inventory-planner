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
  const items=[await act(p)];for(let i=0;i<7;i++){await p.keyboard.press('ArrowDown');items.push(await act(p))}
  ok('מקלדת · «⋯» נפתח במיקוד על הפריט הראשון והחצים עוברים על כל 8',items.join(',')==='tour,xp,btn,sapXls,sapSlowXls,expXls,darkToggle,stBtn'||new Set(items).size===8,items.join(','));
  // ייצוא במקלדת
  await p.keyboard.press('Escape');await p.waitForTimeout(100);
  ok('מקלדת · Escape סוגר את «⋯» ומחזיר מיקוד',await act(p)==='helpBtn');
  for(const [id,chk] of [['exportXls','x'],['expXls','x'],['darkToggle','dark']]){
   const x0=await p.evaluate(()=>({x:window.__x,d:document.body.classList.contains('dark')}));
   await p.focus('#helpBtn');await p.keyboard.press('Enter');await p.waitForTimeout(120);
   for(let i=0;i<10&&await act(p)!==id;i++)await p.keyboard.press('ArrowDown');
   const at=await act(p);await p.keyboard.press('Enter');await p.waitForTimeout(1200);
   const x1=await p.evaluate(()=>({x:window.__x,d:document.body.classList.contains('dark'),open:document.getElementById('helpMenu').classList.contains('open')}));
   ok(`מקלדת · ${id} מגיע בחצים ומופעל ב-Enter, והתפריט נסגר`,at===id&&(chk==='x'?x1.x>x0.x:x1.d!==x0.d)&&!x1.open,`${at} ${JSON.stringify(x1)}`);
   if(chk==='dark')await p.evaluate(()=>document.getElementById('darkToggle').click())}
  // SAP — פותח תצוגה מקדימה, לא קובץ ישר
  await p.focus('#helpBtn');await p.keyboard.press('Enter');await p.waitForTimeout(120);
  for(let i=0;i<10&&await act(p)!=='sapXls';i++)await p.keyboard.press('ArrowDown');
  await p.keyboard.press('Enter');await p.waitForTimeout(800);
  const sap=await p.evaluate(()=>({x:window.__x,modal:!![...document.querySelectorAll('[role=dialog],.modal,.sapprev')].find(e=>e.offsetParent!==null)}));
  ok('מקלדת · «עדכון ל-SAP» מופעל מהתפריט (תצוגה מקדימה או קובץ)',sap.modal||sap.x>0,JSON.stringify(sap));
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
  ok(`${T} · «⋯» נפתח בנגיעה, 8 פריטים, כולם בתוך המסך`,hm.length===8&&hm.every(x=>x.ok),hm.filter(x=>!x.ok).map(x=>x.id).join(',')||hm.length);
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
