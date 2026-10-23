/* ============ VIEW_SHEET · גיליון «התצוגה» = כל הרשימה, בסדר שעל המסך ============
   הכלל, בכל 26 המסכים: גיליון «התצוגה» בייצוא לאקסל מכיל את כל מה
   שעומד בסינון הנוכחי, בסדר שעל המסך — בלי קשר לכמה שורות צוירו,
   לאילו קבוצות פתוחות, ולאיזה מסך היה פתוח לפני.

   נמדד על הדוח האמיתי לפני התיקון: «כיסוי חלקי» (5) יצא עם 164 שורות של
   «ביקוש יורד», «יושמו ב-SAP» (0) עם 225, הקטלוג (7,990) עם 182; מסכי
   הכסף התקוע — רק השורות שצוירו (417 מתוך 2,777); וקיפול קבוצה הוריד
   אותה מהגיליון. ובמסך «תנועות אחרונות» הייצוא וכפתורי ה-SAP נפלו על
   שגיאת JS.

   לכל מסך פריטים (25):
     · אותו גיליון, בין אם הגענו מהקטלוג ובין אם מ«מרכז עבודה»
     · אותם פריטים בדיוק כמו currentRows() — בלי כפילות, בלי חסר
     · השורות שעל המסך מופיעות בגיליון באותו סדר
     · קיפול כל הקבוצות לא משנה את הגיליון — והמסך כן מתקפל
   ובנוסף: אחרי שינוי סינון הגיליון מתעדכן; «תנועות אחרונות» — ייצוא
   בלחיצה, כפתורי SAP לא זמינים עם הסבר, אפס שגיאות.
   VIEW_XLSX=<דוח> — על דוח אחר, מקומית. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const XL=process.env.VIEW_XLSX||path.join(SD,'zmrp-demo.xlsx');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0,null,{timeout:180000});
 await p.waitForTimeout(800);
 await p.evaluate(()=>{window.__cap=[];XLSX.writeFile=(wb,name)=>{const sheets={};
  for(const sn of wb.SheetNames){const a=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1});
   const i=(a[0]||[]).findIndex(c=>/מק״ט/.test(c));sheets[sn]={rows:a.length-1,pns:i<0?null:a.slice(1).map(x=>String(x[i]))}}
  window.__cap.push({name,sheets})}});

 const R=await p.evaluate(()=>{
  const grab=()=>{__cap.length=0;try{doExport()}catch(e){return {err:e.message}}
   const f=__cap[0];return f&&f.sheets['התצוגה']?f.sheets['התצוגה'].pns:null};
  const dom=()=>[...document.querySelectorAll('#tbl tbody tr[data-i] .obj')].map(e=>e.textContent.trim());
  const subseq=(small,big)=>{let j=0;for(const x of big)if(x===small[j])j++;return j===small.length};
  const res={};
  for(const m of Object.keys(MODESET).filter(m=>m!=='moves')){
   setMode('catalog');setMode(m);const a=grab();
   setMode('home');setMode(m);const a2=grab();
   const want=currentRows().map(r=>String(r.pn));const shown=dom();
   const gs=[...document.querySelectorAll('#tbl tbody tr.grp[data-g]')].map(t=>t.dataset.g);
   gs.forEach(g=>GCOLL.add(g));render();
   const shownColl=dom().length,c=grab();
   gs.forEach(g=>GCOLL.delete(g));render();
   const S=x=>Array.isArray(x)?[...x].sort().join('\n'):'';
   res[m]={n:want.length,sheet:Array.isArray(a)?a.length:a,err:a&&a.err,
    prev:JSON.stringify(a)===JSON.stringify(a2),
    set:Array.isArray(a)&&a.length===want.length&&new Set(a).size===a.length&&S(a)===S(want),
    order:Array.isArray(a)&&subseq(shown,a),shown:shown.length,
    groups:gs.length,coll:JSON.stringify(c)===JSON.stringify(a),shownColl}}
  return res});
 const ms=Object.keys(R);
 ok('נבדקו 25 מסכי פריטים',ms.length===25,ms.length);
 const f=(k,pred)=>ms.filter(m=>!pred(R[m]));
 let x=f('err',v=>!v.err);ok('הייצוא רץ בלי שגיאה בכל מסך',!x.length,x.join(', '));
 x=f('prev',v=>v.prev);ok('הגיליון לא תלוי במסך שהיה פתוח קודם',!x.length,x.join(', ')||'25 מסכים');
 x=f('set',v=>v.set);ok('«התצוגה» = כל הפריטים שבסינון — בלי חסר ובלי כפילות',!x.length,
   x.map(m=>`${m}: ${R[m].sheet}≠${R[m].n}`).join(' · ')||`${ms.reduce((t,m)=>t+R[m].n,0).toLocaleString('en-US')} שורות`);
 x=f('order',v=>v.order);ok('השורות שעל המסך מופיעות בגיליון באותו סדר',!x.length,x.join(', '));
 const cut=ms.filter(m=>R[m].shown<R[m].n);
 ok('גם במסך שמצייר רק חלק — הגיליון מלא',cut.every(m=>R[m].set),cut.map(m=>`${m} ${R[m].shown}/${R[m].n}`).join(' · ')||'אין מסך חתוך בדוח הזה');
 const gr=ms.filter(m=>R[m].groups);
 ok('קיפול כל הקבוצות לא משנה את הגיליון',gr.every(m=>R[m].coll),gr.filter(m=>!R[m].coll).join(', ')||`${gr.length} מסכים מקובצים`);
 ok('ובמסך — הקבוצות באמת התקפלו',gr.some(m=>R[m].shownColl<R[m].shown),gr.map(m=>`${m} ${R[m].shown}→${R[m].shownColl}`).slice(0,4).join(' · '));

 /* ---- שינוי סינון: הגיליון מתעדכן ---- */
 const F=await p.evaluate(()=>{const k=FDEF[0][0];const c={};for(const r of ALL)c[r[k]]=(c[r[k]]||0)+1;
  const v=Object.entries(c).sort((a,b)=>b[1]-a[1])[0][0];
  setMode('catalog');__cap.length=0;doExport();const before=__cap[0].sheets['התצוגה'].rows;
  SEL[k]=new Set([v]);apply();setMode('catalog');__cap.length=0;doExport();const after=__cap[0].sheets['התצוגה'].rows;
  const want=currentRows().length;SEL[k]=new Set();apply();return {before,after,want,n:c[v]}});
 ok('אחרי סינון — «התצוגה» בקטלוג היא בדיוק מה שנשאר',F.after===F.want&&F.after===F.n&&F.after<F.before,`${F.before} → ${F.after}`);

 /* ---- «תנועות אחרונות» ---- */
 await p.evaluate(()=>{setMode('moves')});await p.waitForTimeout(300);
 const e0=errs.length;
 const M=await p.evaluate(()=>{__cap.length=0;document.getElementById('exportXls').click();
  const f=__cap[0];const t=document.getElementById('toast').textContent;
  return {file:f?f.name:null,sheets:f?Object.fromEntries(Object.entries(f.sheets).map(([k,v])=>[k,v.rows])):null,
   mov:movRows().length,all:QF.all.length,toast:t}});
 ok('תנועות · «ייצוא לאקסל» מוריד קובץ',!!M.file&&errs.length===e0,M.file||errs.slice(e0).join(' | '));
 ok('תנועות · «התצוגה» = טבלת התנועות, «כל הנתונים» = כל הפריטים',
   M.sheets&&M.sheets['התצוגה']===M.mov&&Object.entries(M.sheets).some(([k,v])=>/^כל הנתונים/.test(k)&&v===M.all),JSON.stringify(M.sheets));
 const SB=await p.evaluate(()=>['sapXls','sapSlowXls'].map(id=>{const b=document.getElementById(id);
  __cap.length=0;b.click();const t=document.getElementById('toast').textContent;
  return {id,aria:b.getAttribute('aria-disabled'),na:b.classList.contains('na'),title:b.title,toast:t,
   file:__cap.length,modal:document.getElementById('modalwrap').classList.contains('open')}}));
 ok('תנועות · כפתורי SAP מסומנים כלא זמינים, עם הסבר',SB.every(s=>s.aria==='true'&&s.na&&/אינו רשימת פריטים/.test(s.title)));
 ok('תנועות · לחיצה על SAP: הסבר בהודעה, בלי קובץ, בלי חלון, בלי שגיאה',
   SB.every(s=>/אינו רשימת פריטים/.test(s.toast)&&!s.file&&!s.modal)&&errs.length===e0,SB.map(s=>s.id+':'+s.toast.slice(0,30)).join(' | '));
 const back=await p.evaluate(()=>{setMode('month');const b=document.getElementById('sapXls');
  return {aria:b.getAttribute('aria-disabled'),na:b.classList.contains('na'),title:b.title}});
 ok('ובמסך פריטים — הכפתורים חוזרים לפעולה, עם התיאור המקורי',back.aria==='false'&&!back.na&&!/אינו רשימת פריטים/.test(back.title),back.title.slice(0,40));

 ok('אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
