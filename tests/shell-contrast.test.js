/* ============ D1 · השלד — ניגודיות בכותרת, בתפריטים ובפוקוס ============
   מסגרת הגרפיט ותפריטים בהירים עליה: כל טקסט גלוי בכותרת, בסרגל הניווט,
   בסרגל הרשימה ובשלושת התפריטים — ≥4.5:1 (גדול ≥3:1) מול הרקע שמתחתיו.
   מחוון הפוקוס של כל פקד בהם — ≥3:1 מול הרקע (הקו או ההילה). צהוב לבד על
   לבן הוא 1.6:1, ולכן על משטח בהיר הפוקוס הוא טבעת דיו עם הילה צהובה;
   הבדיקה נוסתה על עותק שבו הטבעת צהובה בלבד והטקסט המשני כהה — ונפלה. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML=process.argv[2]||path.join(SD,'..','index.html');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null?'  ['+x+']':''))};
const SCAN=(sel)=>{
 const P=c=>{const m=c.match(/[\d.]+/g);return m?m.map(Number):null};
 const lum=([r,g,b])=>{const f=v=>{v/=255;return v<=.03928?v/12.92:((v+.055)/1.055)**2.4};return .2126*f(r)+.7152*f(g)+.0722*f(b)};
 const cr=(a,b)=>{const x=lum(a),y=lum(b);return (Math.max(x,y)+.05)/(Math.min(x,y)+.05)};
 const bgOf=e=>{for(let x=e;x;x=x.parentElement){const c=P(getComputedStyle(x).backgroundColor);if(c&&(c.length<4||c[3]>.9))return c}return [255,255,255]};
 const res=[];
 for(const root of document.querySelectorAll(sel))for(const e of [root,...root.querySelectorAll('*')]){
  const r=e.getBoundingClientRect();if(!r.width||!r.height)continue;const cs=getComputedStyle(e);
  if(cs.visibility==='hidden'||+cs.opacity<1)continue;
  if(![...e.childNodes].some(n=>n.nodeType===3&&n.textContent.trim()))continue;
  let op=1;for(let x=e;x;x=x.parentElement)op*=+getComputedStyle(x).opacity;if(op<1)continue;
  if(e.closest('[disabled],.off'))continue;
  const fs=parseFloat(cs.fontSize),big=fs>=24||(fs>=18.66&&+cs.fontWeight>=700);
  const v=cr(P(cs.color),bgOf(e));res.push({t:e.textContent.trim().slice(0,18),v:+v.toFixed(2),need:big?3:4.5})}
 return res};
const FOCUS=()=>{const P=c=>{const m=c.match(/[\d.]+/g);return m?m.map(Number):null};
 const lum=([r,g,b])=>{const f=v=>{v/=255;return v<=.03928?v/12.92:((v+.055)/1.055)**2.4};return .2126*f(r)+.7152*f(g)+.0722*f(b)};
 const cr=(a,b)=>{const x=lum(a),y=lum(b);return (Math.max(x,y)+.05)/(Math.min(x,y)+.05)};
 const bgOf=e=>{for(let x=e;x;x=x.parentElement){const c=P(getComputedStyle(x).backgroundColor);if(c&&(c.length<4||c[3]>.9))return c}return [255,255,255]};
 const e=document.activeElement;const cs=getComputedStyle(e);const around=bgOf(e.parentElement||e);
 const ow=parseFloat(cs.outlineWidth)||0;const oc=cs.outlineStyle!=='none'&&ow>=2?P(cs.outlineColor):null;
 const sh=cs.boxShadow&&cs.boxShadow!=='none'?P(cs.boxShadow):null;
 const vo=oc?cr(oc,around):0,vs=sh?cr(sh,around):0;
 return {id:e.id||e.dataset.help||e.dataset.ck||e.className.toString().slice(0,12),v:+Math.max(vo,vs).toFixed(2),o:+vo.toFixed(2),s:+vs.toFixed(2)}};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 for(const [tag,vp,mob] of [['מחשב',{width:1440,height:900},false],['טלפון',{width:390,height:844},true]])for(const dark of [false,true]){
  const T=tag+(dark?' כהה':' בהיר');
  const ctx=await b.newContext({viewport:vp,isMobile:mob,hasTouch:mob});await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));const p=await ctx.newPage();
  await p.goto('file://'+HTML+'?nobrief=1');await p.setInputFiles('#f',XL);
  await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});
  if(dark)await p.evaluate(()=>{document.body.classList.add('dark')});
  await p.evaluate(()=>setMode('month'));await p.waitForTimeout(200);
  const scan=async(n,sel)=>{const r=await p.evaluate(`(${SCAN})(${JSON.stringify(sel)})`);const low=r.filter(x=>x.v<x.need);
   ok(`${T} · ניגודיות טקסט · ${n} (${r.length} טקסטים)`,r.length>0&&!low.length,low.length?low.map(x=>x.t+' '+x.v).join(' | '):'מינ׳ '+Math.min(...r.map(x=>x.v)).toFixed(1))};
  await scan('כותרת','.top');await scan('סרגל ניווט','.rail');await scan('סרגל הרשימה','.phd');
  await p.click('#fchip');await p.waitForTimeout(200);await scan('תפריט הקובץ','#fileMenu');await p.keyboard.press('Escape');
  await p.click('#helpBtn');await p.waitForTimeout(200);await scan('תפריט ⋯','#helpMenu');await p.keyboard.press('Escape');
  await p.click('#colsBtn');await p.waitForTimeout(200);await scan('בורר תצוגה','#colspick');await p.evaluate(()=>colsOpen(false));
  // פוקוס
  const foc=[];const fset=async(sel,open)=>{if(open){await p.focus(open);await p.keyboard.press('Enter');await p.waitForTimeout(150)}
   const ids=await p.evaluate(s=>[...document.querySelectorAll(s)].filter(e=>e.getBoundingClientRect().width>0).map((e,i)=>{e.dataset.__f=s+i;return s+i}),sel);
   for(const id of ids){await p.evaluate(i=>{const e=document.querySelector(`[data-__f="${i}"]`);e.focus()},id);
    /* :focus-visible דורש מקלדת: Shift מקיש בלי לשנות דבר */
    await p.keyboard.press('Shift');foc.push(await p.evaluate(`(${FOCUS})()`))}
   if(open)await p.keyboard.press('Escape')};
  await fset('.top .lc>label.btn input, #q, #fchip, #helpBtn');
  await fset('#fileMenu button:not([hidden]), #fileMenu input:not([type=file])','#fchip');
  await fset('#helpMenu [role=menuitem]','#helpBtn');
  await fset('#filtBtn, #colsBtn, #bulkBtn');
  await fset('#colspick button, #colspick input','#colsBtn');
  await fset('#tabs .tab');
  const lowF=foc.filter(x=>x.v<3);
  ok(`${T} · מחוון פוקוס ≥3:1 מול הרקע (${foc.length} פקדים)`,foc.length>15&&!lowF.length,lowF.length?lowF.map(x=>`${x.id} ${x.o}/${x.s}`).join(' | '):'מינ׳ '+Math.min(...foc.map(x=>x.v)).toFixed(1));
  await ctx.close()}
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})();
