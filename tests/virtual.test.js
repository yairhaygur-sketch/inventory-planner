/* ============ D2A · גלילה וירטואלית — כל הפריטים נגישים ============
   עד D2A כל מרנדר צייר את כל השורות בבת אחת, ולכן רשימה גדולה נחתכה:
   הקטלוג ב-500, «שינוי פרופיל MRP» ב-500, מרכז העבודה ב-300, «תנועות»
   ב-500 — והשאר לא היה נגיש בגלילה. עכשיו הרשימה כולה נשמרת (VL) ובמסך
   מצויר רק החלון שבגלילה. הבדיקה הזאת מוודאה, בכל 26 המסכים:

     · הרשימה במסך = כל מה שבסינון (ובמסכי הכסף: 80% כברירת מחדל,
       «הצג את כל הפריטים» פותח את השאר, וקו ה-80% מסומן)
     · גלילה עד הסוף מגיעה לפריט האחרון, בלי חור באף נקודה בדרך
     · מספר השורות ב-DOM חסום — לא גדל עם הרשימה
     · העמודות לא «רוקדות» ברוחב בזמן הגלילה
     · המקלדת (↓ ↑ PageDown Enter Escape) עוברת על כל הרשימה וגוללת אליה
     · כרטיס הפריט: «הבא» עובר את גבול החלון, והמיקום «מתוך» כל הרשימה
     · החיפוש מוצא פריט שמחוץ לחלון, בוחר אותו וגולל אליו

   ובעותק שבו הסף לציור מלא (VFULL) הוא 10 — כדי שגם רשימה קצרה תעבור
   במסלול הווירטואלי: קבוצות מתקפלות ונפתחות, בחירת SAP נשמרת כשהשורה
   יוצאת מהחלון, «בחר הכול» בכותרת עובד על הרשימה ולא על מה שצויר,
   כותרת דביקה ועמודות מוקפאות.

   VIRT_XLSX=<דוח> — על דוח אחר (למשל האמיתי, מקומית). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path'),os=require('os');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const XL=process.env.VIRT_XLSX||path.join(SD,'zmrp-demo.xlsx');
const HTML=path.join(SD,'..','index.html');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
const N=n=>Number(n).toLocaleString('he-IL');
async function open(b,file,W,H){
 const ctx=await b.newContext({viewport:{width:W,height:H}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+file+'?nobrief=1');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0,null,{timeout:180000});
 await p.waitForTimeout(600);
 await p.evaluate(()=>{window.__raf=()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
  window.__kd=k=>document.dispatchEvent(new KeyboardEvent('keydown',{key:k,bubbles:true}))});
 return {ctx,p,errs}}
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});

 /* ---------------- כל המסכים, בשני רוחבים ---------------- */
 for(const [W,H] of [[1440,900],[390,844]]){
  const {ctx,p,errs}=await open(b,HTML,W,H);
  const R=await p.evaluate(async()=>{GCOLL.clear();CAPALL.clear();const res=[];
   for(const m of Object.keys(MODESET)){
    try{setMode(m)}catch(_){continue}
    await __raf();
    const s=VL[cur==='moves'?'movtbl':'tbl'];if(!s||!s.tb.isConnected){res.push({m,none:true});continue}
    /* מסכי הכסף הם חמש הדלתות של «כסף תקוע». track עצמו נשאר 'cap' גם
       אחרי מעבר לקטלוג, ולכן לא ממנו. */
    const k=/^cap/.test(m)&&TRACKKEY[track];let want;
    if(cur==='moves')want=movRows().length;
    else{const rows=currentRows();want=k?capView(rows,cut80(rows,track)).length:rows.length}
    const box=s.box,n=s.items.length,sel=cur==='moves'?'tr[data-mi]':'tr[data-i]';
    const ths=()=>[...s.tb.parentNode.querySelectorAll('thead th')].map(t=>t.getBoundingClientRect().width);
    const cover=()=>{const bt=box.getBoundingClientRect(),th=(s.tb.parentNode.tHead||{}).offsetHeight||0;
     const rs=[...s.tb.children].filter(t=>!t.classList.contains('vsp'));if(!rs.length)return true;
     const f=rs[0].getBoundingClientRect(),l=rs[rs.length-1].getBoundingClientRect();
     return f.top<=bt.top+th+1&&(l.bottom>=bt.bottom-1||rs[rs.length-1]===s.tb.lastElementChild)};
    box.scrollTop=0;await __raf();const w0=ths();
    let maxTr=s.tb.children.length,gaps=0,jit=0;
    for(let i=0;i<5000&&box.clientHeight>0&&box.scrollTop+box.clientHeight<box.scrollHeight-1;i++){
     box.scrollTop+=box.clientHeight*.9;await __raf();
     maxTr=Math.max(maxTr,s.tb.children.length);if(!cover())gaps++;
     const w=ths();for(let j=0;j<w.length;j++)jit=Math.max(jit,Math.abs(w[j]-(w0[j]||w[j])))}
    const trs=[...s.tb.querySelectorAll(sel)];
    const last=trs.length?+(trs[trs.length-1].dataset.i??trs[trs.length-1].dataset.mi):-1;
    /* קפיצה לאמצע — כמו גרירה של פס הגלילה */
    box.scrollTop=box.scrollHeight/2;await __raf();await __raf();const mid=cover();
    res.push({m,n,want,virt:s.ents.length>VFULL,maxTr,gaps,jit:Math.round(jit),last,mid,boxH:box.clientHeight})}
   return res});
  const L=R.filter(r=>!r.none);
  ok(`${W} · נבדקו כל 26 המסכים`,R.length===26,R.length);
  let x=L.filter(r=>r.n!==r.want);
  ok(`${W} · הרשימה במסך = כל מה שבסינון (במסכי הכסף — עד קו ה-80%)`,!x.length,
    x.map(r=>`${r.m} ${r.n}≠${r.want}`).join(' · ')||`${N(L.reduce((t,r)=>t+r.n,0))} פריטים`);
  /* מסך שתיבת הגלילה שלו בגובה 0 (בטלפון, «כסף תקוע» — כך גם ב-main) נמדד בנפרד */
  const H0=L.filter(r=>r.boxH===0&&r.n);
  x=L.filter(r=>r.boxH>0&&r.n&&r.last!==r.n-1);
  ok(`${W} · גלילה עד הסוף מגיעה לפריט האחרון`,!x.length,x.map(r=>`${r.m}: ${r.last+1}/${r.n}`).join(' · ')
    ||`${L.filter(r=>r.virt).length} רשימות וירטואליות`+(H0.length?` · תיבה בגובה 0: ${H0.map(r=>r.m).join(', ')}`:''));
  x=L.filter(r=>r.gaps||!r.mid);
  ok(`${W} · אין חור בגלילה — גם בקפיצה לאמצע`,!x.length,x.map(r=>`${r.m}: ${r.gaps}`).join(' · '));
  const big=L.filter(r=>r.virt);
  const lim=await p.evaluate(()=>Math.max(VFULL,6*VCHUNK));
  x=big.filter(r=>r.maxTr>lim);
  ok(`${W} · מספר השורות ב-DOM חסום (≤${lim}) גם ברשימה של אלפים`,big.length>0&&!x.length,
    x.map(r=>`${r.m} ${r.maxTr}`).join(' · ')||big.map(r=>`${r.m} ${N(r.n)}→${r.maxTr}`).join(' · '));
  if(W>760){x=L.filter(r=>r.jit>24);
   ok(`${W} · העמודות לא קופצות ברוחב בגלילה (≤24px)`,!x.length,
     x.map(r=>`${r.m} ${r.jit}px`).join(' · ')||`מקס׳ ${Math.max(0,...L.map(r=>r.jit))}px`)}

  /* ---------------- מקלדת וכרטיס — בקטלוג ---------------- */
  const K=await p.evaluate(async()=>{setMode('catalog');await __raf();const s=VL.tbl,n=s.items.length;
   s.box.scrollTop=0;s.sel=-1;await __raf();
   __kd('ArrowDown');await __raf();const first=s.sel;
   let pd=0;while(s.sel<n-1&&pd<5000){__kd('PageDown');pd++;if(pd%25===0)await __raf()}
   await __raf();const bt=s.box.getBoundingClientRect(),tr=s.tb.querySelector('tr.sel');
   const vis=!!tr&&tr.getBoundingClientRect().bottom<=bt.bottom+1&&tr.getBoundingClientRect().top>=bt.top-1;
   const end=s.sel,same=CURRENT_DETAIL===s.view[s.sel];
   __kd('ArrowUp');await __raf();const up=s.sel;
   const pos=(document.querySelector('.opos')||{}).textContent||'';
   __kd('Escape');await __raf();const closed=!document.body.classList.contains('dopen');
   s.box.scrollTop=0;await __raf();
   __kd('Enter');await __raf();const back=s.sel===up&&!!s.tb.querySelector('tr.sel');
   __kd('Escape');await __raf();
   return {n,first,end,vis,same,up,pos,closed,back,pd,virt:s.ents.length>VFULL}});
  ok(`${W} · מקלדת: ↓ בוחר את הראשון, PageDown מגיע עד האחרון וגולל אליו`,
    K.first===0&&K.end===K.n-1&&K.vis&&K.same,`${N(K.n)} פריטים · ${K.pd} לחיצות · ${K.virt?'וירטואלי':'מלא'}`);
  ok(`${W} · ↑ חוזר אחד, והכרטיס אומר «מתוך» כל הרשימה`,K.up===K.n-2&&new RegExp(`${N(K.up+1)} מתוך ${N(K.n)}`).test(K.pos),K.pos);
  ok(`${W} · Escape סוגר, ו-Enter חוזר לאותו פריט גם כשהוא מחוץ לחלון`,K.closed&&K.back);
  const C=await p.evaluate(async()=>{setMode('catalog');await __raf();const s=VL.tbl;s.box.scrollTop=0;await __raf();
   s.tb.querySelector('tr[data-i]').click();await __raf();const steps=Math.min(s.items.length-1,s.b+5);let k=0;
   for(;k<steps;k++){const nx=document.querySelector('[data-dnav="1"]');if(!nx||nx.disabled)break;nx.click();if(k%10===0)await __raf()}
   await __raf();const tr=s.tb.querySelector('tr.sel');
   const r={k,sel:s.sel,same:CURRENT_DETAIL===s.view[s.sel],tr:!!tr&&+tr.dataset.i===s.sel};
   __kd('Escape');await __raf();return r});
  ok(`${W} · כרטיס פריט: «הבא» עובר את גבול החלון ומציג את הפריט הנכון`,C.sel===C.k&&C.same&&C.tr,`${C.k} צעדים`);
  const Q=await p.evaluate(async()=>{setMode('catalog');await __raf();const s=VL.tbl;s.box.scrollTop=0;await __raf();
   const r=s.view[s.view.length-1];openSug(r);await __raf();const tr=s.tb.querySelector('tr.sel'),bt=s.box.getBoundingClientRect();
   const o={i:s.sel,n:s.view.length,vis:!!tr&&tr.getBoundingClientRect().bottom<=bt.bottom+1,
    toast:document.getElementById('toast').textContent};closeDetail();return o});
  ok(`${W} · חיפוש: הפריט האחרון ברשימה נמצא, נבחר ומוצג`,Q.i===Q.n-1&&Q.vis&&!/אינו ברשימה/.test(Q.toast),`${N(Q.i+1)}/${N(Q.n)}`);

  /* ---------------- מסכי הכסף: 80% / «הצג את כל הפריטים» ---------------- */
  const M=await p.evaluate(async()=>{const res=[];
   for(const m of Object.keys(MODESET).filter(m=>/^cap/.test(m))){setMode(m);await __raf();
    const s=()=>VL.tbl,rows=currentRows(),btn=document.getElementById('capAll');
    const a={m,n:s().items.length,total:rows.length,btn:!btn.hidden&&getComputedStyle(btn).display!=='none',
     more:s().ents.some(e=>typeof e.h==='string'&&/class="vmore"/.test(e.h)),cuts:s().ents.filter(e=>typeof e.h==='string'&&/cutrow/.test(e.h)).length};
    if(a.n>=a.total){res.push(a);continue}
    const mb=()=>{const e=s().ents.find(e=>typeof e.h==='string'&&/class="vmore"/.test(e.h));return e};
    s().box.scrollTop=s().box.scrollHeight;await __raf();await __raf();
    const row=s().tb.querySelector('tr.vmore .capall');a.rowBtn=!!row;(row||btn).click();await __raf();
    const ents=s().ents,cj=ents.findIndex(e=>typeof e.h==='string'&&/cutrow/.test(e.h));
    const prev=ents[cj-1],next=ents[cj+1];
    a.all={n:s().items.length,cuts:ents.filter(e=>typeof e.h==='string'&&/cutrow/.test(e.h)).length,
     above:prev&&prev.h&&!/class="below"/.test(prev.h()),below:next&&typeof next.h==='function'&&/class="below"/.test(next.h()),
     back:!!mb()&&/רק 80%/.test(mb().h),pg:document.getElementById('pgL').textContent};
    setMode('catalog');setMode(m);await __raf();a.keep=s().items.length;
    CAPALL.delete(m);render();await __raf();a.reset=s().items.length;
    res.push(a)}
   return res});
  const MM=M.filter(a=>a.n<a.total);
  ok(`${W} · מסכי הכסף: ברירת המחדל היא 80% מהשווי, עם «הצג את כל הפריטים»`,M.length===5&&MM.every(a=>a.more&&a.rowBtn&&(W<=760||a.btn)),
    M.map(a=>`${a.m} ${N(a.n)}/${N(a.total)}`).join(' · '));
  ok(`${W} · «הצג את כל הפריטים» מציג את כולם, וקו ה-80% מסומן במקומו`,
    MM.length>0&&MM.every(a=>a.all.n===a.total&&a.all.cuts===a.cuts&&a.all.cuts>0&&a.all.above&&a.all.below&&a.all.back),
    MM.map(a=>`${a.m}: ${N(a.all.n)} · ${a.all.cuts} קווים`).join(' · '));
  ok(`${W} · הבחירה נשמרת לדלת במעבר מסך, וחוזרת ל-80% בלחיצה`,MM.every(a=>a.keep===a.total&&a.reset===a.n));
  ok(`${W} · אפס שגיאות JS`,errs.length===0,errs.slice(0,2).join(' | '));
  await ctx.close()}

 /* ---------------- VFULL=10, חלון של 5: גם רשימה קצרה במסלול הווירטואלי ---------------- */
 const src=fs.readFileSync(HTML,'utf8');
 const tmp=path.join(fs.mkdtempSync(path.join(os.tmpdir(),'vfull-')),'index.html');
 const patched=src.replace('const VL={},VFULL=200,VCHUNK=40,VOVER=900,','const VL={},VFULL=10,VCHUNK=5,VOVER=0,');
 ok('העותק עם VFULL=10 נוצר',patched!==src);
 fs.writeFileSync(tmp,patched);
 {const {ctx,p,errs}=await open(b,tmp,1440,900);
  const G=await p.evaluate(async()=>{GCOLL.clear();const res=[];
   for(const m of Object.keys(MODESET)){setMode(m);await __raf();const s=VL.tbl;
    if(cur==='moves'||!s.tb.isConnected)continue;
    /* שורת מטבע (grp cur, data-g="cur:…") אינה קבוצה מתקפלת — גם ב-main */
    const gs=s.ents.filter(e=>e.k==='g'&&/<tr class="grp(?: sup)?" data-g="(?!cur:)/.test(e.h));
    if(!gs.length||s.ents.length<=VFULL)continue;
    const n0=s.items.length;
    /* הקבוצה האחרונה — בדרך כלל מחוץ לחלון הראשון */
    s.box.scrollTop=s.box.scrollHeight;await __raf();await __raf();
    const tr=[...VL.tbl.tb.querySelectorAll('tr.grp:not(.cur):not([data-g^="cur:"])')].pop();if(!tr)continue;
    const g=tr.dataset.g,inG=+((tr.querySelector('.gn')||{}).textContent||'0').replace(/[^\d]/g,'');tr.click();await __raf();
    const n1=VL.tbl.items.length;VL.tbl.box.scrollTop=VL.tbl.box.scrollHeight;await __raf();await __raf();
    const shown=!!VL.tbl.tb.querySelector(`tr.grp.coll[data-g="${CSS.escape(g)}"]`);
    VL.tbl.tb.querySelector(`tr.grp[data-g="${CSS.escape(g)}"]`).click();await __raf();
    res.push({m,g,n0,n1,inG,shown,n2:VL.tbl.items.length,open:!GCOLL.has(g)})}
   GCOLL.clear();return res});
  ok('VFULL=10 · קבוצות: קיפול הקבוצה האחרונה מוריד בדיוק את שורותיה, ופתיחה מחזירה',
    G.length>0&&G.every(g=>g.n1===g.n0-g.inG&&g.shown&&g.open&&g.n2===g.n0),
    G.map(g=>`${g.m} ${g.n0}→${g.n1}→${g.n2}`).join(' · '));
  const S=await p.evaluate(async()=>{SAPSEL.clear();setMode('month');await __raf();const s=VL.tbl;
   if(s.ents.length<=VFULL)return {skip:true};
   s.box.scrollTop=0;await __raf();const b0=s.tb.querySelector('.sapsel');if(!b0)return {none:true};const pn=b0.dataset.pn;
   b0.checked=true;b0.dispatchEvent(new Event('change'));
   s.box.scrollTop=s.box.scrollHeight;await __raf();await __raf();const gone=!s.tb.querySelector(`.sapsel[data-pn="${CSS.escape(pn)}"]`);
   s.box.scrollTop=0;await __raf();await __raf();const kept=!!(s.tb.querySelector(`.sapsel[data-pn="${CSS.escape(pn)}"]`)||{}).checked;
   const all=document.querySelector('#tbl thead .sapall');const ind1=all.indeterminate;
   all.checked=true;all.dispatchEvent(new Event('change'));
   const pool=sapAllPool().length,n1=SAPSEL.size,h1=all.checked&&!all.indeterminate;
   const used=s.view.slice(0,300).filter(sapUsableBase).length;
   s.box.scrollTop=s.box.scrollHeight;await __raf();await __raf();
   const far=[...s.tb.querySelectorAll('.sapsel')],farAll=far.length>0&&far.every(x=>x.checked);
   const fb=far[far.length-1];fb.checked=false;fb.dispatchEvent(new Event('change'));
   const ind2=all.indeterminate,bar=document.getElementById('sapbarN').textContent;
   SAPSEL.clear();sapBarSync();return {gone,kept,ind1,pool,n1,h1,used,farAll,ind2,bar}});
  if(S.skip||S.none)ok('VFULL=10 · בחירת SAP — אין בתכנון תיבות לבדוק',true,JSON.stringify(S));
  else{
   ok('VFULL=10 · SAP: תיבה שסומנה נשארת מסומנת אחרי שהשורה יצאה מהחלון וחזרה',S.gone&&S.kept);
   ok('VFULL=10 · SAP: «בחר הכול» בוחר את כל מי שיש לו תיבה — כמו קודם (עד 300) — גם מחוץ לחלון',
     S.ind1&&S.h1&&S.n1===S.pool&&S.pool===S.used&&S.farAll,`${S.n1} נבחרו`);
   ok('VFULL=10 · SAP: ביטול אחת רחוקה — הכותרת «חלקית» והמונה יורד',S.ind2&&S.bar===N(S.n1-1),S.bar)}
  const T=await p.evaluate(async()=>{setMode('catalog');await __raf();const s=VL.tbl;s.box.scrollTop=s.box.scrollHeight/2;await __raf();
   return Math.round(s.tb.parentNode.tHead.querySelector('th').getBoundingClientRect().top-s.box.getBoundingClientRect().top)});
  ok('VFULL=10 · הכותרת נשארת דביקה באמצע הרשימה',Math.abs(T)<=1,T+'px');
  await ctx.close();
  ok('VFULL=10 · אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '))}
 {const {ctx,p,errs}=await open(b,tmp,780,800);
  const F=await p.evaluate(async()=>{setMode('today');await __raf();await new Promise(r=>setTimeout(r,250));const s=VL.tbl,box=s.box;
   const fz=document.body.classList.contains('freeze'),sw=box.scrollWidth-box.clientWidth;
   box.scrollLeft=-Math.min(200,sw);box.scrollTop=box.scrollHeight/2;await __raf();await __raf();
   const tr=[...s.tb.querySelectorAll('tr[data-i]')][2],bx=box.getBoundingClientRect();
   return {fz,sw,virt:s.ents.length>VFULL,left:Math.round(box.scrollLeft),
    pin:tr?Math.round(bx.right-tr.children[0].getBoundingClientRect().right):null,
    pos:tr?getComputedStyle(tr.children[1]).position:''}});
  ok('VFULL=10 · 780px: בגלילה אופקית ואנכית יחד, העמודות המוקפאות נשארות בקצה',
    !F.fz||(F.virt&&F.left<0&&Math.abs(F.pin)<=1&&F.pos==='sticky'),JSON.stringify(F));
  await ctx.close();
  ok('VFULL=10 · 780px · אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '))}
 try{fs.rmSync(path.dirname(tmp),{recursive:true,force:true})}catch(_){}
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
