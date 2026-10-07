const {chromium}=require('playwright');const fs=require('fs');
const path=require('path');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const hex=h=>{const m=h.match(/\d+/g);return m?[+m[0],+m[1],+m[2]]:[0,0,0]};
const L=c=>{const s=c.map(v=>{v/=255;return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4)});
 return 0.2126*s[0]+0.7152*s[1]+0.0722*s[2]};
const R=(a,b)=>{const x=L(hex(a)),y=L(hex(b));return (Math.max(x,y)+0.05)/(Math.min(x,y)+0.05)};
let total=0;
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1600,height:1000}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();
 await p.goto('file://'+path.join(__dirname,'..','index.html')+'?nobrief=1');
 await p.setInputFiles('#f',path.join(__dirname,'zmrp-demo.xlsx'));
 /* הכלי נוחת ב«מרכז עבודה» ולא בטבלה. הבדיקה סורקת גם מסכי טבלה
    וגם את מסך הנחיתה (ראה הלולאה למטה), ולכן היא נכנסת לטבלה
    במפורש כדי להמתין לשורות. */
 await p.evaluate(()=>{try{setMode('today')}catch(_){}});
 await p.waitForSelector('#tbl tbody tr[data-i]',{timeout:120000});
 /* גם דוח ה-ETA — אחרת הצ׳יפ בכותרת ועמודות «משובץ»/«תקוע» אינם
    מרונדרים כלל, וצבע שלא מרונדר אינו נמדד. */
 await p.setInputFiles('#fe',path.join(__dirname,'zmrp-demo-eta.xlsx'));
 await p.waitForTimeout(1500);
 /* ============ גם המסכים שאינם טבלה ============
    הבדיקה מדדה את מסך ברירת המחדל בלבד. «מרכז עבודה» הוא מסך שלם
    עם כרטיסים, תגי צבע ורצועת אזהרה שאינם מרונדרים שם — וצבע שלא
    מרונדר אינו נמדד. כל מסך נמדד בשני המצבים. */
 for(const [mode,scr] of [['light','line'],['light','home'],['light','cap'],
                          ['dark','line'],['dark','home'],['dark','cap']]){
  await p.evaluate(m=>document.body.classList.toggle('dark',m==='dark'),mode);
  await p.evaluate(k=>{const t=[...document.querySelectorAll('#tabs .tab')]
    .find(x=>x.dataset.m===k);if(t)t.click()},scr);
  await p.waitForTimeout(400);
  // כל אלמנט גלוי עם טקסט: מודדים צבע מול הרקע האפקטיבי
  const bad=await p.evaluate(()=>{
   const eff=el=>{let n=el;while(n&&n!==document.documentElement){
     const bg=getComputedStyle(n).backgroundColor;
     if(bg&&!/rgba\(0, 0, 0, 0\)|transparent/.test(bg))return bg;n=n.parentElement}
    return getComputedStyle(document.body).backgroundColor};
   const out=[];
   for(const el of document.querySelectorAll('body *')){
    const cs=getComputedStyle(el);
    if(cs.display==='none'||cs.visibility==='hidden'||+cs.opacity===0)continue;
    const r=el.getBoundingClientRect(); if(r.width<4||r.height<4)continue;
    const t=[...el.childNodes].filter(n=>n.nodeType===3&&n.textContent.trim()).map(n=>n.textContent.trim()).join(' ');
    if(!t)continue;
    out.push({sel:el.tagName.toLowerCase()+(el.className&&typeof el.className==='string'?'.'+el.className.trim().split(/\s+/).join('.'):''),
      txt:t.slice(0,22),fg:cs.color,bg:eff(el),size:parseFloat(cs.fontSize),weight:cs.fontWeight});
   }
   return out;
  });
  const fails=[];const seen=new Set();
  for(const e of bad){
   const r=R(e.fg,e.bg);
   const large=e.size>=24||(e.size>=18.66&&+e.weight>=700);
   const need=large?3:4.5;
   if(r<need){const k=e.sel+'|'+e.fg+'|'+e.bg;if(seen.has(k))continue;seen.add(k);
    fails.push(`${r.toFixed(2)} (דרוש ${need}) ${e.sel}  «${e.txt}»  ${e.fg} על ${e.bg}`)}
  }
  console.log(`\n=== ${mode} · ${scr} · ${bad.length} אלמנטים · ${fails.length} כשלים ייחודיים ===`);
  total+=fails.length;
  fails.sort((a,b)=>parseFloat(a)-parseFloat(b)).slice(0,25).forEach(f=>console.log('  '+f));
 }
 /* ============ CANVAS_CONTRAST — מה שהסריקה הזאת לא ראתה ============
    הלולאה שלמעלה סורקת **אלמנטים ב-DOM**. גרפי הפריט מצוירים בקנבס,
    ולכן מעולם לא נבדקו — ושני פגמים שרדו שם עד שמעתד הסתכל בהם:

    הלוח כהה תמיד (#23262e), אבל הקוד קרא את טוקני *הנושא הפעיל*.
    במצב בהיר, מול אותו משטח: תוויות החודש 2.68:1, מילוי «צניחה»
    2.13:1, כיתוב החציון 2.93:1 — כולם נכשלים.
    ובמקביל הקנבס צויר ב-clientWidth*2 בלי גובה CSS, ולכן הדפדפן גזר
    את הגובה מיחס התכונות והציג הכל בחצי — «13px» הגיעו כ-6.5px.

    שלוש בדיקות, אחת לכל מצב כשל, ועוד אחת שמוודאת שהקבועים מתארים
    את המציאות ולא רק את הכוונה. */
 {
  const HX=h=>{h=h.replace('#','');return [0,2,4].map(i=>parseInt(h.slice(i,i+2),16))};
  const CR=(a,c)=>{const x=L(HX(a)),y=L(HX(c));return (Math.max(x,y)+0.05)/(Math.min(x,y)+0.05)};
  await p.evaluate(()=>document.body.classList.remove('dark'));
  await p.evaluate(()=>{const r=(QF.all||[]).find(x=>x.A&&x.months.some(v=>v>0))||(QF.all||[])[0];detail(r)});
  await p.waitForTimeout(700);
  const cc=await p.evaluate(()=>{
   const cv=document.getElementById('c1');if(!cv||!cv.width)return null;
   const g=cv.getContext('2d');
   const px=(x,y)=>{const d=g.getImageData(x,y,1,1).data;
     return '#'+[d[0],d[1],d[2]].map(v=>v.toString(16).padStart(2,'0')).join('')};
   const b=cv.getBoundingClientRect();
   const dpr=Math.min(window.devicePixelRatio||1,3);
   return {surf:CHART_SURF,col:Object.assign({},CHART_C),
     /* פיקסל מתוך הלוח עצמו, הרחק מכל סימן — מוכיח שהקבוע הוא המציאות */
     panelPx:px(Math.round(cv.width*0.5),Math.round(cv.height*0.14)),
     cssH:cv.style.height,
     attrW:cv.width,attrH:cv.height,
     wantW:Math.round(b.width*dpr),wantH:Math.round(b.height*dpr),
     data:[...g.getImageData(0,0,cv.width,cv.height).data].join(',').length,
     hash:(()=>{const d=g.getImageData(0,0,cv.width,cv.height).data;
       let h=0;for(let i=0;i<d.length;i+=97)h=(h*31+d[i])>>>0;return h})()}});
  const cfails=[];
  if(!cc)cfails.push('הקנבס לא רונדר כלל');
  else{
   /* טקסט דורש 4.5, סימן גרפי דורש 3 */
   const NEED={ink:4.5,ink2:4.5,ink3:4.5,base:3,spike:3,drop:3,cur:3,med:3};
   /* צבע חסר או שאינו hex הוא כשל, לא קריסה: ערך שמגיע מטוקן של
      הנושא (rgb(...) או ריק) הוא בדיוק הרגרסיה שהבדיקה נועדה לתפוס. */
   for(const k in NEED){const v=cc.col[k];
    if(typeof v!=='string'||!/^#[0-9a-f]{6}$/i.test(v)){
     cfails.push(`CHART_C.${k} אינו צבע קבוע (${JSON.stringify(v)}) — צבע שנגזר מהנושא חוזר להתהפך מול לוח קבוע`);continue}
    const r=CR(v,cc.surf);
    if(r<NEED[k])cfails.push(`${r.toFixed(2)} (דרוש ${NEED[k]}) CHART_C.${k} ${v} על ${cc.surf}`)}
   for(const k in cc.col)if(!(k in NEED))cfails.push(`CHART_C.${k} אינו מכוסה בבדיקה — הוסף אותו ל-NEED`);
   if(cc.panelPx.toLowerCase()!==cc.surf.toLowerCase())
    cfails.push(`הלוח שצויר ${cc.panelPx} אינו CHART_SURF ${cc.surf} — הקבוע אינו מתאר את המציאות`);
   /* הגיאומטריה: גובה CSS מפורש, ומאגר פיקסלים css×dpr בשני הצירים.
      בלי זה הגופן חוזר להיות חצי ממה שנכתב. */
   if(!cc.cssH)cfails.push('אין גובה CSS על הקנבס — הדפדפן יגזור אותו מיחס התכונות ויקטין הכל');
   if(Math.abs(cc.attrW-cc.wantW)>2||Math.abs(cc.attrH-cc.wantH)>2)
    cfails.push(`מאגר הפיקסלים ${cc.attrW}×${cc.attrH} אינו css×dpr (${cc.wantW}×${cc.wantH})`);
  }
  /* הגרף אינו מתהפך עם הנושא — אותו ציור בדיוק בבהיר ובכהה */
  let darkHash=null;
  if(cc){
   await p.evaluate(()=>{document.body.classList.add('dark');
     const r=(typeof CURRENT_DETAIL!=='undefined'&&CURRENT_DETAIL)?CURRENT_DETAIL:null;if(r)detail(r)});
   await p.waitForTimeout(700);
   darkHash=await p.evaluate(()=>{const cv=document.getElementById('c1');
     if(!cv||!cv.width)return null;const d=cv.getContext('2d').getImageData(0,0,cv.width,cv.height).data;
     let h=0;for(let i=0;i<d.length;i+=97)h=(h*31+d[i])>>>0;return h});
   await p.evaluate(()=>document.body.classList.remove('dark'));
   if(darkHash===null)cfails.push('הקנבס לא רונדר במצב כהה');
   else if(darkHash!==cc.hash)
    cfails.push(`הגרף משתנה עם הנושא — bright ${cc.hash} מול dark ${darkHash}. משטח קבוע מחייב צעדים קבועים`);
  }
  console.log(`\n=== canvas · גרף הפריט · ${cfails.length} כשלים ===`);
  cfails.forEach(f=>console.log('  '+f));
  total+=cfails.length;
 }
 const bad=total>0;
 await b.close();
 process.exit(total>0?1:0);
})();
