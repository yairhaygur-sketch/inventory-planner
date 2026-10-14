/* ============ UX — שלוש תקלות שדווחו מהשטח אחרי יומיים של עבודה ============
   כולן נראו למעתד כ«עיצוב» ואף אחת מהן לא הייתה:

   1 · «כפתור הסימון לא עובד». המנגנון היה תקין לגמרי. הפאנל ישב
       ב-data-dt="hist" — תחת לשונית «היסטוריה» — והכרטיס נפתח על
       «החלטה». ארבע לחיצות עד הכפתור, ואחת מהן דורשת לנחש ש«סמן
       כטופל» נמצא תחת «היסטוריה». הפקד האח, «מצב טיפול», לא נשא
       data-dt כלל והיה גלוי תמיד.
   2 · «בשביל להגדיל אני מושך לצד השני». הידית ישבה על הקצה
       ה-inline-start, שהוא הימין ב-RTL — כלומר הגבול עם העמודה
       ה*קודמת* — בעוד הקוד הגדיל בגרירה שמאלה. נמדד: גרירה ימינה
       ב-60px הקטינה מ-245 ל-194.
   3 · «הכותרת דורסת את הטקסט». `#tbl thead th{background:transparent}`
       גבר על `thead th{background:var(--panel0)}` בזכות ה-id, וכותרת
       position:sticky עם רקע שקוף מראה את השורות שנגללות מתחתיה.

   הבדיקות כאן נוגעות בעכבר אמיתי ולא דוחפות ערכים ל-DOM: תקלה 2
   היא תקלת *כיוון*, ואי אפשr לתפוס אותה בלי גרירה. הטענות הן
   על כיוון ועל נוכחות, לא על ספי פיקסלים — גדלים מושפעים מסביבת
   הגופנים, ועל זה כבר נכווינו. */
const {chromium}=require('playwright'),fs=require('fs'),path=require('path');
const SD=__dirname, HTML='file://'+path.join(SD,'..','index.html');
const out=[];const ok=(n,c,d)=>out.push(`${c?'PASS':'FAIL'} · ${n}${d?'  ['+d+']':''}`);
const alpha=c=>{const m=/rgba?\(([^)]+)\)/.exec(c||'');
 if(!m)return c==='transparent'?0:1;
 const p=m[1].split(',').map(s=>parseFloat(s));return p.length>3?p[3]:1};

(async()=>{
 const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];
 p.on('pageerror',e=>errs.push(e.message.split('\n')[0].slice(0,120)));
 await p.goto(HTML+'?nobrief=1');await p.waitForTimeout(1200);
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(3000);
 await p.evaluate(()=>setMode('today'));await p.waitForTimeout(1200);

 /* ── 1. הסימון נמצא על הלשונית שהכרטיס נפתח עליה ────────────── */
 /* פתיחת הכרטיס בחץ למטה, כמו בשאר הסוויטה. לחיצת עכבר על נקודה
    מחושבת בשורה החטיאה על הפיקסטורה (העמודות צרות יותר מבדוח
    האמיתי), ובדיקה שנשענת על גאומטריה של פיקסטורה תישבר שוב. */
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(900);
 ok('הכרטיס נפתח',await p.evaluate(()=>document.body.classList.contains('dopen')),'ArrowDown');

 const card=await p.evaluate(()=>{
  const tab=([...document.querySelectorAll('#dtabs .dt')].find(x=>x.classList.contains('on'))||{}).textContent||'';
  const d=[...document.querySelectorAll('#detail details.dsec')]
    .find(x=>/סימון והחרגה/.test(x.querySelector('summary')?.textContent||''));
  if(!d)return {tab,found:false};
  const r=d.getBoundingClientRect();
  return {tab:tab.trim(),found:true,w:Math.round(r.width),h:Math.round(r.height),
    dt:d.getAttribute('data-dt'),parentHidden:!!d.parentElement.hidden}});
 ok('הכרטיס נפתח על «החלטה»',card.tab==='החלטה',card.tab);
 ok('פאנל הסימון גלוי בלי לעבור לשונית',
   card.found&&card.w>0&&card.h>0&&!card.parentHidden,
   card.found?`${card.w}×${card.h} · data-dt=${card.dt===null?'ללא':card.dt}`:'לא נמצא');
 ok('ואינו משויך ללשונית «היסטוריה»',card.found&&card.dt===null,
   `data-dt=${card.dt===null?'ללא':card.dt}`);

 /* שתי לחיצות בלבד: פתיחת המקטע, ואז «טופל» */
 const pre=await p.evaluate(()=>Object.keys(MARKS||{}).length);
 const sp=await p.evaluate(()=>{const d=[...document.querySelectorAll('#detail details.dsec')]
   .find(x=>/סימון והחרגה/.test(x.querySelector('summary')?.textContent||''));
  const s=d.querySelector('summary');s.scrollIntoView({block:'center'});
  const r=s.getBoundingClientRect();return {x:Math.round(r.x+50),y:Math.round(r.y+r.height/2)}});
 await p.waitForTimeout(300);await p.mouse.click(sp.x,sp.y);await p.waitForTimeout(600);
 const bp=await p.evaluate(()=>{const d=[...document.querySelectorAll('#detail details.dsec')]
   .find(x=>/סימון והחרגה/.test(x.querySelector('summary')?.textContent||''));
  const btn=d.querySelector('.marks button[data-t="handled"]');
  if(!btn)return null;btn.scrollIntoView({block:'center'});
  const r=btn.getBoundingClientRect();
  return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2),w:Math.round(r.width)}});
 ok('כפתור «טופל» מרונדר בגודל אמיתי',!!bp&&bp.w>0,bp?bp.w+'px':'0');
 await p.waitForTimeout(300);await p.mouse.click(bp.x,bp.y);await p.waitForTimeout(900);
 const post=await p.evaluate(()=>Object.keys(MARKS||{}).length);
 ok('ושתי לחיצות מהכרטיס מסמנות פריט',post>pre,`${pre} → ${post}`);

 /* ── 2. ידית הרוחב: המיקום והכיוון מסכימים ─────────────────── */
 /* סוגרים את המגירה דרך Escape — לחיצה על סלקטור מנוחש עלולה
    לפגוע באלמנט אחר ולשנות מצב בלי שהבדיקה תדע. */
 await p.keyboard.press('Escape');await p.waitForTimeout(600);
 const hz=await p.evaluate(()=>{
  const ths=[...document.querySelectorAll('#tbl thead th')];
  const handles=ths.map((t,k)=>t.querySelector('.colrz')?k:-1).filter(k=>k>=0);
  if(!handles.length)return {found:false,ths:ths.length,handles:0};
  /* עמודה שיש לה ידית ושאינה האחרונה — האחרונה אינה גובלת בבאה */
  const i=handles.find(k=>k<ths.length-1);
  if(i==null)return {found:false,ths:ths.length,handles:handles.length};
  const th=ths[i],h=th.querySelector('.colrz');
  const t=th.getBoundingClientRect(),r=h.getBoundingClientRect();
  const next=ths[i+1]?ths[i+1].getBoundingClientRect():null;
  const rtl=getComputedStyle(document.documentElement).direction==='rtl';
  /* ב-RTL העמודה הבאה נמצאת משמאל, ולכן הגבול איתה הוא הקצה השמאלי */
  const edge=rtl?t.left:t.right;
  return {found:true,rtl,i,onBoundary:Math.abs((r.left+r.right)/2-edge)<=6,
    nextIsOnThatSide:next?(rtl?next.right<=t.left+2:next.left>=t.right-2):null,
    col:(th.textContent||'').trim().slice(0,12)}});
 ok('הידית יושבת על הגבול עם העמודה הבאה',
   hz.found&&hz.onBoundary&&hz.nextIsOnThatSide!==false,
   hz.found?`עמודה «${hz.col}» · RTL=${hz.rtl} · על הגבול=${hz.onBoundary}`
          :`אין ידית · ${hz.ths} כותרות · ${hz.handles} ידיות`);
 if(!hz.found){
  ok('גרירת הקצה החופשי החוצה מגדילה את העמודה',false,'לא נבדק — אין ידית');
  ok('וגרירה פנימה מקטינה אותה',false,'לא נבדק — אין ידית')}

 const drag=async dx=>{
  const s=await p.evaluate(i=>{const th=[...document.querySelectorAll('#tbl thead th')][i];
    const h=th.querySelector('.colrz');const r=h.getBoundingClientRect();
    return {x:Math.round(r.x+r.width/2),y:Math.round(r.y+r.height/2),
      w:Math.round(th.getBoundingClientRect().width)}},hz.i);
  await p.mouse.move(s.x,s.y);await p.mouse.down();
  await p.mouse.move(s.x+dx,s.y,{steps:8});await p.waitForTimeout(140);
  await p.mouse.up();await p.waitForTimeout(400);
  const w=await p.evaluate(i=>Math.round(
    [...document.querySelectorAll('#tbl thead th')][i].getBoundingClientRect().width),hz.i);
  return {before:s.w,after:w}};
 /* «החוצה» = הרחק מגוף העמודה. ב-RTL הקצה החופשי הוא השמאלי,
    ולכן גרירה שמאלה היא המשיכה שמרחיבה. */
 if(hz.found){
  const outward=hz.rtl?-70:70;
  const g1=await drag(outward);
  ok('גרירת הקצה החופשי החוצה מגדילה את העמודה',g1.after>g1.before,
    `${g1.before} → ${g1.after}`);
  const g2=await drag(-outward);
  ok('וגרירה פנימה מקטינה אותה',g2.after<g2.before,`${g2.before} → ${g2.after}`)}

 /* ── 3. כותרת דביקה אטומה ───────────────────────────────────── */
 const head=await p.evaluate(()=>{
  const th=document.querySelector('#tbl thead th');
  const cs=getComputedStyle(th);
  const w=document.querySelector('#tbl').closest('.twrap')||document.querySelector('#tbl').parentElement;
  w.scrollTop=300;
  return {pos:cs.position,bg:cs.backgroundColor,z:cs.zIndex}});
 await p.waitForTimeout(400);
 const cover=await p.evaluate(()=>{
  const th=document.querySelector('#tbl thead th');const hr=th.getBoundingClientRect();
  let n=0;document.querySelectorAll('#tbl tbody tr').forEach(tr=>{
   const r=tr.getBoundingClientRect();if(r.bottom>hr.top+2&&r.top<hr.bottom-2)n++});
  return n});
 ok('כותרת הטבלה היא sticky',head.pos==='sticky',head.pos);
 ok('ויש לה רקע אטום — שורות אינן נראות דרכה',alpha(head.bg)>=1,
   `${head.bg} · אלפא ${alpha(head.bg)}`);
 ok('הבדיקה אכן גללה שורות אל מתחת לכותרת',cover>0,
   `${cover} שורות חופפות את רצועת הכותרת`);

 ok('אין שגיאת JS לאורך כל המסלול',errs.length===0,errs.slice(0,2).join(' · ')||'נקי');
 await b.close();
 for(const l of out)console.log(l);
 const bad=out.filter(l=>l.startsWith('FAIL')).length;
 console.log(`\n${out.length-bad}/${out.length} עברו`);
 process.exit(bad?1:0);
})().catch(e=>{console.log('FAIL · הבדיקה קרסה · '+e.message);process.exit(1)});
