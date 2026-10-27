#!/usr/bin/env node
/* ============ visual-snapshot — לפני/אחרי חזותי, על נתונים סינתטיים בלבד ============
   שימוש:
     node visual-snapshot.js [--base <git-ref|קובץ.html>] [--html index.html]
                             [--xlsx tests/zmrp-demo.xlsx] [--out <dir>] [--now <ISO>]
                             [--expect-same] [--max-diff 0.1]

   מצלם את אותם 8 מצבים בשתי גרסאות (ברירת מחדל: origin/main מול index.html):
     7 משפחות המסכים — home · today · line · month · cap · qual · catalog
     ועוד מגירת פריט פתוחה (today, השורה הראשונה)
   כל מצב: מחשב 1440×900 וטלפון 390×844 · בהיר וכהה — 32 צילומים לגרסה.

   לכל צילום: אחוז הפיקסלים ששונו, ומדדים — שורות במסך, גובה שורה,
   היכן מתחילה השורה הראשונה, פקדים מעל הרשימה, גדלי גופן וצבעי טקסט
   שונים, צבעי טקסט מתחת ל-4.5:1, ופקדים מתחת ל-44px בטלפון.
   הפלט: report.html + report.json + PNG (לפני · אחרי · הבדל) ב---out.

   --expect-same: יוצא 1 אם צילום כלשהו שונה ביותר מ---max-diff אחוז
   (ברירת מחדל 0.1%). זה תנאי המעבר של שלב שאין בו שינוי חזותי (D0).

   ============ נתונים סינתטיים בלבד ============
   הריפו ציבורי והצילומים עשויים להיות מצורפים ל-PR. לכן הכלי מסרב לכל
   דוח שאינו בתוך tests/ של הריפו — שם יושבים רק דוחות שהמחולל בנה.
   דוח אמיתי מקומי נבדק ב-business-guard.js, שאינו מצלם. */
const {execFileSync}=require('child_process');
const path=require('path'),fs=require('fs'),os=require('os');
const args=process.argv.slice(2);
const opt=(k,d)=>{const i=args.indexOf(k);return i<0?d:args[i+1]};
const flag=k=>args.includes(k);
const root=execFileSync('git',['rev-parse','--show-toplevel'],{encoding:'utf8'}).trim();
const PW=require(path.join(root,'node_modules','playwright'));
const base=opt('--base','origin/main'),html=path.resolve(opt('--html',path.join(root,'index.html')));
const xlsx=path.resolve(opt('--xlsx',path.join(root,'tests','zmrp-demo.xlsx')));
const out=path.resolve(opt('--out',fs.mkdtempSync(path.join(os.tmpdir(),'visual-'))));
const MAXD=+opt('--max-diff','0.1');
const d0=new Date();d0.setHours(8,0,0,0);const NOW=opt('--now',d0.toISOString()),TZ='Asia/Jerusalem';

const testsDir=path.join(root,'tests')+path.sep;
if(!xlsx.startsWith(testsDir)){console.error(`סירוב: ${xlsx} אינו בתוך tests/. הכלי מצלם רק נתונים סינתטיים — דוח אמיתי בודקים ב-business-guard.js.`);process.exit(2)}
if(!fs.existsSync(xlsx)){console.error(`${xlsx} לא קיים — הרץ קודם: npm run fixture`);process.exit(2)}
fs.mkdirSync(out,{recursive:true});

/* הגרסה הבסיסית: קובץ, או index.html מתוך git */
let baseFile=base;
if(!fs.existsSync(base)){baseFile=path.join(out,'base.html');
 fs.writeFileSync(baseFile,execFileSync('git',['show',`${base}:index.html`],{cwd:root,maxBuffer:64e6}))}

const STATES=[['home'],['today'],['line'],['month'],['cap'],['qual'],['catalog'],['today','drawer']];
const VIEWS=[['d',{width:1440,height:900},false],['m',{width:390,height:844},true]];
const THEMES=['light','dark'];

/* מדדים — רצים בדף. השורות: [data-row] כשקיים, אחרת הסלקטורים הישנים,
   כדי שאותה מדידה תעבוד גם על main שלפני D0. */
const METRICS=()=>{
 const vis=e=>{const b=e.getBoundingClientRect();return b.width>0&&b.height>0&&b.top>=0&&b.bottom<=innerHeight&&b.left>=0&&b.right<=innerWidth};
 const rowsAll=[...document.querySelectorAll('[data-row]')].length?[...document.querySelectorAll('[data-row]')]:[...document.querySelectorAll('#tbl tbody tr[data-i], #movtbl tbody tr[data-mi]')];
 const rows=rowsAll.filter(vis);
 const top=rows.length?rows[0].getBoundingClientRect().top:null;
 const ctrls=[...document.querySelectorAll('button,a[href],input,select,textarea,[role=button]')].filter(e=>vis(e)&&getComputedStyle(e).visibility!=='hidden'&&(top==null||e.getBoundingClientRect().bottom<=top+1)&&!e.closest('#tabs,#railsub,nav'));
 const lum=c=>{const m=c.match(/[\d.]+/g);if(!m)return null;const [r,g,b]=m.slice(0,3).map(v=>{v/=255;return v<=.03928?v/12.92:((v+.055)/1.055)**2.4});return .2126*r+.7152*g+.0722*b};
 const bgOf=e=>{for(let x=e;x;x=x.parentElement){const c=getComputedStyle(x).backgroundColor;const m=c.match(/[\d.]+/g);if(m&&(m.length<4||+m[3]>.9))return c}return 'rgb(255,255,255)'};
 const texts=[...document.querySelectorAll('body *')].filter(e=>vis(e)&&[...e.childNodes].some(n=>n.nodeType===3&&n.textContent.trim()));
 const sizes=new Set(),colors=new Set();let low=0;
 for(const e of texts){const cs=getComputedStyle(e);sizes.add(cs.fontSize);colors.add(cs.color);
  if(+cs.opacity<1)continue;const a=lum(cs.color),b=lum(bgOf(e));if(a==null||b==null)continue;
  const r=(Math.max(a,b)+.05)/(Math.min(a,b)+.05),big=parseFloat(cs.fontSize)>=24||(parseFloat(cs.fontSize)>=18.66&&+cs.fontWeight>=700);
  if(r<(big?3:4.5))low++}
 const small=matchMedia('(pointer:coarse)').matches?[...document.querySelectorAll('button,a[href],input,select,[role=button]')].filter(e=>{if(!vis(e))return false;const b=e.getBoundingClientRect();return b.width<44||b.height<44}).length:null;
 return {rows:rows.length,rowH:rows.length?Math.round(rows.reduce((a,e)=>a+e.getBoundingClientRect().height,0)/rows.length):0,
  firstRowTop:top==null?null:Math.round(top),controlsAbove:ctrls.length,fontSizes:sizes.size,textColors:colors.size,lowContrast:low,smallTargets:small}};

async function shoot(b,file,tag){
 const res={};
 for(const [v,vp,mob] of VIEWS)for(const th of THEMES){
  const ctx=await b.newContext({viewport:vp,isMobile:mob,hasTouch:mob,deviceScaleFactor:1,timezoneId:TZ});
  const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));p.on('dialog',d=>d.accept());
  await p.clock.setFixedTime(new Date(NOW));
  await p.goto('file://'+file+'?nobrief=1');
  await p.addStyleTag({content:'*,*::before,*::after{transition:none!important;animation:none!important;caret-color:transparent!important}'});
  await p.setInputFiles('#f',xlsx);
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&(typeof LOADJOB==='undefined'||!LOADJOB),null,{timeout:180000});
  await p.click('#tourSkip').catch(()=>{});
  await p.evaluate(dk=>{if(typeof applyDark==='function')applyDark(dk);else document.body.classList.toggle('dark',dk)},th==='dark');
  for(const [m,extra] of STATES){
   await p.evaluate(m=>{setMode(m);scrollTo(0,0);document.querySelectorAll('.rows').forEach(x=>x.scrollTop=0)},m);
   if(extra==='drawer'){await p.evaluate(()=>{const tr=document.querySelector('#tbl tbody tr[data-i]');tr&&tr.click()})}
   await p.waitForTimeout(350);
   const key=`${m}${extra?'-'+extra:''}-${v}-${th}`,png=path.join(out,`${tag}-${key}.png`);
   await p.screenshot({path:png});
   res[key]={png,...await p.evaluate(METRICS),errs:errs.length};
   if(extra==='drawer')await p.keyboard.press('Escape')}
  await ctx.close()}
 return res}

/* הבדל פיקסלים — בתוך הדפדפן, על canvas; ערוץ שזז ביותר מ-8 נחשב שינוי */
async function diff(b,a,c,outPng){
 const p=await b.newPage();
 const r=await p.evaluate(async([A,C])=>{
  const load=src=>new Promise(res=>{const i=new Image();i.onload=()=>res(i);i.src=src});
  const [ia,ic]=await Promise.all([load(A),load(C)]);
  const w=Math.max(ia.width,ic.width),h=Math.max(ia.height,ic.height);
  const cv=n=>{const c=document.createElement('canvas');c.width=w;c.height=h;const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,w,h);x.drawImage(n,0,0);return [c,x.getImageData(0,0,w,h)]};
  const [,da]=cv(ia),[cc,dc]=cv(ic);let n=0;const o=new ImageData(w,h);
  for(let i=0;i<da.data.length;i+=4){const d=Math.abs(da.data[i]-dc.data[i])>8||Math.abs(da.data[i+1]-dc.data[i+1])>8||Math.abs(da.data[i+2]-dc.data[i+2])>8;
   if(d){n++;o.data[i]=230;o.data[i+1]=30;o.data[i+2]=30;o.data[i+3]=255}else{const g=(dc.data[i]+dc.data[i+1]+dc.data[i+2])/3;o.data[i]=o.data[i+1]=o.data[i+2]=200+g/5;o.data[i+3]=255}}
  cc.getContext('2d').putImageData(o,0,0);return {pct:n/(w*h)*100,size:ia.width===ic.width&&ia.height===ic.height,png:cc.toDataURL('image/png')}},
  ['data:image/png;base64,'+fs.readFileSync(a).toString('base64'),'data:image/png;base64,'+fs.readFileSync(c).toString('base64')]);
 fs.writeFileSync(outPng,Buffer.from(r.png.split(',')[1],'base64'));await p.close();return {pct:r.pct,sameSize:r.size}}

(async()=>{
 const b=await PW.chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 console.log(`בסיס: ${base} · נבדק: ${path.relative(root,html)} · דוח: ${path.relative(root,xlsx)} · שעון: ${NOW}`);
 const A=await shoot(b,path.resolve(baseFile),'before'),C=await shoot(b,html,'after');
 const rows=[];let worst=0;
 for(const k of Object.keys(A)){const d=await diff(b,A[k].png,C[k].png,path.join(out,`diff-${k}.png`));
  worst=Math.max(worst,d.pct);const {png:_a,...ma}=A[k],{png:_c,...mc}=C[k];rows.push({k,diff:+d.pct.toFixed(3),sameSize:d.sameSize,before:ma,after:mc})}
 await b.close();
 fs.writeFileSync(path.join(out,'report.json'),JSON.stringify({base,html,now:NOW,rows},null,1));
 const M=['rows','rowH','firstRowTop','controlsAbove','fontSizes','textColors','lowContrast','smallTargets'];
 const cell=(a,c)=>a===c?`${c??'—'}`:`<b>${a??'—'} → ${c??'—'}</b>`;
 fs.writeFileSync(path.join(out,'report.html'),`<!doctype html><html dir="rtl" lang="he"><meta charset="utf-8"><title>visual-snapshot</title>
<style>body{font:13px system-ui;margin:16px}table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:3px 6px;text-align:center}b{color:#b3261e}img{width:32%;border:1px solid #ccc;margin:2px}</style>
<h1>לפני / אחרי — ${base} מול ${path.basename(html)}</h1><p>הנתונים: ${path.basename(xlsx)} (סינתטי) · ההבדל הגדול ביותר: ${worst.toFixed(3)}%</p>
<table><tr><th>מצב</th><th>% פיקסלים</th>${M.map(m=>`<th>${m}</th>`).join('')}</tr>
${rows.map(r=>`<tr><td>${r.k}</td><td>${r.diff}</td>${M.map(m=>`<td>${cell(r.before[m],r.after[m])}</td>`).join('')}</tr>`).join('')}</table>
${rows.map(r=>`<h3>${r.k} · ${r.diff}%</h3><img src="before-${r.k}.png"><img src="after-${r.k}.png"><img src="diff-${r.k}.png">`).join('')}</html>`);
 const w=s=>String(s).padEnd(26);
 console.log(w('מצב')+'% פיקסלים   שורות   מעל הרשימה   גופנים   ניגודיות<4.5   <44px');
 for(const r of rows)console.log(w(r.k)+String(r.diff).padEnd(12)+`${cell(r.before.rows,r.after.rows)}`.replace(/<\/?b>/g,'').padEnd(8)
  +`${cell(r.before.controlsAbove,r.after.controlsAbove)}`.replace(/<\/?b>/g,'').padEnd(13)+`${cell(r.before.fontSizes,r.after.fontSizes)}`.replace(/<\/?b>/g,'').padEnd(9)
  +`${cell(r.before.lowContrast,r.after.lowContrast)}`.replace(/<\/?b>/g,'').padEnd(15)+`${cell(r.before.smallTargets,r.after.smallTargets)}`.replace(/<\/?b>/g,''));
 const errs=rows.filter(r=>r.after.errs);if(errs.length)console.log('שגיאות JS:',errs.map(r=>r.k).join(', '));
 console.log(`\n${rows.length} צילומים · ההבדל הגדול ביותר ${worst.toFixed(3)}% · דוח: ${path.join(out,'report.html')}`);
 if(flag('--expect-same')){const over=rows.filter(r=>r.diff>MAXD||!r.sameSize);
  if(over.length||errs.length){console.log(`✗ ${over.length} צילומים שונים מעל ${MAXD}%: ${over.map(r=>r.k).join(', ')}`);process.exit(1)}
  console.log(`✓ אין שינוי חזותי (כל הצילומים ≤ ${MAXD}%)`)}
})().catch(e=>{console.error(e);process.exit(1)});
