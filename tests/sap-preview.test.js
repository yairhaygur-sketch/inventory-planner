/* ============ S2 · לפני קובץ SAP — המספרים הם הקובץ ============
   לפני ההורדה מוצג כמה פריטים ייכללו, כמה לא, ולמה. הבדיקה הזאת
   מוודאה שאין פער בין מה שמוצג למה שיוצא:

     · בכל אחד מ-26 המסכים, לשני הקבצים (עדכון ל-SAP · מלאי איטי):
         «ייכללו» = מספר השורות בקובץ ש-doSapMass כותב בפועל
         «ייכללו» + «לא ייכללו» = כל הפריטים בתצוגה
         אין אף פריט בסיבה «לא מזוהה» — כל פסילה מוסברת
     · בחירה ידנית (פס «קובץ SAP לנבחרים»): אותו דבר מול
       doSapMass(undefined,sapSelRows())
     · ממשק: לחיצה על «עדכון ל-SAP» פותחת תצוגה ולא מורידה כלום;
       «ביטול» לא מוריד; «הורד» מוריד קובץ עם בדיוק המספר שהוצג

   SAP_XLSX=<דוח> מריץ את אותה בדיקה על דוח אחר (למשל האמיתי, מקומית).
   ברירת המחדל — הדוח הסינתטי. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const XL=process.env.SAP_XLSX||path.join(SD,'zmrp-demo.xlsx');
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
 await p.evaluate(()=>{window.__cap=[];XLSX.writeFile=(wb,name)=>{
  const ws=wb.Sheets[wb.SheetNames[0]];
  window.__cap.push({name,rows:XLSX.utils.sheet_to_json(ws,{header:1}).length-1})}});

 /* ---- כל המסכים, שני הקבצים ---- */
 const R=await p.evaluate(()=>{const res=[];
  for(const m of Object.keys(MODESET)){try{setMode(m)}catch(_){continue}
   for(const kind of [undefined,'slow']){
    /* «תנועות» אינו רשימת פריטים — currentRows זורק שם. עד SAP_NO_LIST
       גם doSapMass זרק; עכשיו הוא אומר למה ולא כותב קובץ. */
    let view;try{view=currentRows()}catch(e){window.__cap.length=0;let t=false;try{doSapMass(kind)}catch(_){t=true}
     res.push({m,kind:kind||'params',throws:true,both:!t&&window.__cap.length===0});continue}
    const P=sapPlan(kind,view,view);
    window.__cap.length=0;doSapMass(kind);
    const file=window.__cap.length?window.__cap[0].rows:0;
    res.push({m,kind:kind||'params',shown:P.rows,file,total:P.total,out:P.out,
     sumWhy:Object.values(P.why).reduce((a,b)=>a+b,0),other:P.why.other||0,view:view.length})}}
  return res});
 const screens=new Set(R.map(r=>r.m)).size;
 ok('נבדקו כל 26 המסכים, שני הקבצים בכל אחד',screens===26&&R.length===52,`${screens} מסכים · ${R.length} מקרים`);
 const thr=R.filter(r=>r.throws);
 ok('מסך שאינו רשימת פריטים: אין קובץ SAP, ואין שגיאה',thr.every(r=>r.both),thr.map(r=>r.m+'/'+r.kind).join(', ')||'אין');
 const RR=R.filter(r=>!r.throws);
 const gap=RR.filter(r=>r.shown!==r.file);
 ok('«ייכללו» = מספר השורות בקובץ שיוצא, בכל מקרה',!gap.length,
   gap.length?gap.slice(0,4).map(r=>`${r.m}/${r.kind}: ${r.shown}≠${r.file}`).join(' · ')
     :`${RR.length} מקרים · ${RR.filter(r=>r.file).length} קבצים לא ריקים · ${RR.reduce((a,r)=>a+r.file,0).toLocaleString('he-IL')} שורות`);
 const sum=RR.filter(r=>r.shown+r.out!==r.view||r.total!==r.view||r.sumWhy!==r.out);
 ok('«ייכללו» + «לא ייכללו» = כל הפריטים בתצוגה',!sum.length,
   sum.slice(0,3).map(r=>`${r.m}/${r.kind}: ${r.shown}+${r.out}≠${r.view}`).join(' · '));
 const oth=RR.filter(r=>r.other);
 ok('כל פריט שלא ייכלל — עם סיבה מזוהה',!oth.length,oth.slice(0,3).map(r=>`${r.m}/${r.kind}: ${r.other}`).join(' · '));
 const kinds=await p.evaluate(()=>{const s=new Set();for(const m of Object.keys(MODESET)){try{setMode(m)}catch(_){continue}
  let v;try{v=currentRows()}catch(_){continue}
  for(const kind of [undefined,'slow'])Object.keys(sapPlan(kind,v,v).why).forEach(k=>s.add(k))}
  return [...s].sort()});
 ok('סיבות שהופיעו',kinds.length>=3,kinds.join(', '));

 /* ---- בחירה ידנית ---- */
 const S=await p.evaluate(()=>{setMode('catalog');
  const all=QF.all;const pick=[...all.filter(sapMassUsable).slice(0,3),...all.filter(r=>sapUsableBase(r)&&SAP_CAP_CATS.has(r.cat)).slice(0,2),
   ...all.filter(r=>!sapUsableBase(r)).slice(0,2)];
  SAPSEL.clear();pick.forEach(r=>SAPSEL.add(r.pn));
  const pool=all.filter(r=>SAPSEL.has(r.pn));const P=sapPlan(undefined,sapSelRows(),pool);
  window.__cap.length=0;doSapMass(undefined,sapSelRows());
  const file=window.__cap.length?window.__cap[0].rows:0;
  SAPSEL.clear();return {n:pick.length,shown:P.rows,file,out:P.out,why:P.why}});
 ok('בחירה ידנית: «ייכללו» = הקובץ, וכל הנבחרים מוסברים',S.shown===S.file&&S.shown+S.out===S.n&&!S.why.other,
   `${S.n} נבחרו · ${S.shown} בקובץ · ${S.out} לא · ${JSON.stringify(S.why)}`);

 /* ---- ממשק ---- */
 await p.evaluate(()=>{setMode('catalog');window.__cap.length=0});await p.waitForTimeout(400);
 const exp=await p.evaluate(()=>sapPlan(undefined,currentRows(),currentRows()));
 await p.evaluate(()=>document.getElementById('sapXls').click());await p.waitForTimeout(300);
 const ui=await p.evaluate(()=>({open:document.getElementById('modalwrap').classList.contains('open'),
  title:document.getElementById('mtitle').textContent,
  inn:document.getElementById('sapIn')?.textContent,outn:document.getElementById('sapOut')?.textContent,
  why:document.querySelectorAll('#sapWhy tr').length,go:!!document.getElementById('sapGo'),
  focus:document.activeElement&&document.activeElement.id,cap:window.__cap.length}));
 const N=n=>n.toLocaleString('he-IL');
 ok('«עדכון ל-SAP» פותח תצוגה לפני הייצוא — ולא מוריד',ui.open&&/לפני הייצוא/.test(ui.title)&&ui.cap===0,ui.title);
 ok('התצוגה מראה את מספר הפריטים שייכללו ושלא',ui.inn===N(exp.rows)&&ui.outn===N(exp.out),`ייכללו ${ui.inn} · לא ${ui.outn}`);
 ok('ושורה לכל סיבה',ui.why===Object.keys(exp.why).length,ui.why+' סיבות');
 ok('המיקוד על «הורד»',ui.focus==='sapGo');
 await p.click('#sapCancel');await p.waitForTimeout(200);
 ok('«ביטול» סוגר ולא מוריד',await p.evaluate(()=>!document.getElementById('modalwrap').classList.contains('open')&&window.__cap.length===0));
 await p.evaluate(()=>document.getElementById('sapXls').click());await p.waitForTimeout(250);
 await p.keyboard.press('Escape');await p.waitForTimeout(200);
 ok('Escape סוגר ולא מוריד',await p.evaluate(()=>!document.getElementById('modalwrap').classList.contains('open')&&window.__cap.length===0));
 await p.evaluate(()=>document.getElementById('sapXls').click());await p.waitForTimeout(250);
 await p.click('#sapGo');await p.waitForTimeout(300);
 const got=await p.evaluate(()=>window.__cap.slice());
 ok('«הורד» מוריד קובץ אחד עם בדיוק המספר שהוצג',got.length===1&&got[0].rows===exp.rows&&/^SAP_Params_/.test(got[0].name),
   got.map(g=>g.name+' · '+g.rows).join(' | '));
 /* מלאי איטי */
 const expS=await p.evaluate(()=>sapPlan('slow',currentRows(),currentRows()));
 await p.evaluate(()=>{window.__cap.length=0;document.getElementById('sapSlowXls').click()});await p.waitForTimeout(250);
 const us=await p.evaluate(()=>({t:document.getElementById('mtitle').textContent,inn:document.getElementById('sapIn')?.textContent,go:!!document.getElementById('sapGo')}));
 ok('«מלאי איטי ל-SAP» — אותה תצוגה, עם המספר שלו',/מלאי איטי/.test(us.t)&&us.inn===N(expS.rows)&&us.go===(expS.rows>0),`ייכללו ${us.inn}`);
 if(us.go){await p.click('#sapGo');await p.waitForTimeout(300);
  const g=await p.evaluate(()=>window.__cap.slice());
  ok('ו«הורד» מוריד את קובץ המלאי האיטי עם אותו מספר',g.length===1&&g[0].rows===expS.rows&&/^SAP_SlowStock_/.test(g[0].name))}
 else await p.click('#sapCancel');
 /* תצוגה בלי אף שורה: אין כפתור הורדה */
 const empty=await p.evaluate(()=>{for(const m of Object.keys(MODESET)){try{setMode(m)}catch(_){continue}
   let v;try{v=currentRows()}catch(_){continue}
   if(v.length&&sapPlan(undefined,v,v).rows===0)return m}return null});
 if(empty){await p.evaluate(()=>{window.__cap.length=0;document.getElementById('sapXls').click()});await p.waitForTimeout(250);
  const e=await p.evaluate(()=>({go:!!document.getElementById('sapGo'),inn:document.getElementById('sapIn')?.textContent}));
  ok(`מסך בלי שורה לקובץ («${empty}»): מוצג 0, ואין כפתור הורדה`,!e.go&&e.inn==='0');
  await p.click('#sapCancel')}
 ok('אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
