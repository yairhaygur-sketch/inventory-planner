/* ============ PD_ONE_MONTH — מועמד PD → ZM-C עם צריכה בחודש אחד בלבד → «לבדוק» (אושר 25.11) ============
   «שינוי סוג MRP» (כלי מלא) ו«טיוב נתונים › שינוי פרופיל MRP» (4 הנושאים): מועמד שכל הצריכה שלו
   ב-11 החודשים ובחודש הנוכחי באה מחודש אחד, ושלא נמכר בשנים הקודמות (השנה שעברה בלי חודשיה שבחלון,
   ולפני שנתיים), עובר ל«לבדוק» עם הסיבה. שני חודשים ומעלה, או מכירה בשנים הקודמות — נשאר «להעלות».
   הבדיקה מחשבת את התנאי בעצמה מהשורה הגולמית של הדוח ומשווה לכלי. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.evaluate(()=>{try{localStorage.clear()}catch(_){}});
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});

 const R=await p.evaluate(()=>{const m=mapCols(ALL_RAW.hdr),An=ANCHOR.getMonth()+1;
  const raw=new Map();for(const row of ALL_RAW.rows)raw.set(CLEAN(row[m.pn]),row);
  const n=v=>Math.max(0,+v||0);
  /* מימוש עצמאי: חודשים פעילים (11 + הנוכחי), ומכירה בשנים הקודמות מחוץ לחלון */
  const mine=r=>{const row=raw.get(r.pn);const ms=(r.months||[]).map(n);const act=ms.filter(v=>v>0).length+(n(r.curMonth)>0?1:0);
   const has3=[m.y0,m.y1,m.y2].every(i=>i>=0);const y1=has3?n(row[m.y1]):0,y2=has3?n(row[m.y2]):0;
   const inWin=ms.slice(Math.min(An,11),11).reduce((a,v)=>a+v,0);const prior=has3?Math.max(0,y1-inWin)+y2:0;
   return {act,prior,one:act===1&&prior<=0}};
  const up=ALL.filter(r=>{const i=profInfo(r);return i&&i.state==='up'});
  const rows=up.map(r=>{const s=profSplit(r),x=mine(r),st=profStatus(r);
   return {pn:r.pn,one:x.one,act:x.act,prior:x.prior,flag:!!r.oneMonth,has:s.reasons.includes('one_month'),last:s.reasons[s.reasons.length-1]==='one_month',
    only:s.reasons.length===1&&s.reasons[0]==='one_month',grp:s.grp,st:st&&st.k,why:st&&st.why||''}});
  const notPD=ALL.filter(r=>r.oneMonth&&!isPD(r.mrp)).length;
  return {n:up.length,rows,notPD}});
 const one=R.rows.filter(x=>x.one),two=R.rows.filter(x=>x.act>=2),prior=R.rows.filter(x=>x.act===1&&x.prior>0);
 ok('בדוח הסינתטי יש מועמדי PD → ZM-C: בחודש אחד, בשני חודשים ומעלה, ובחודש אחד עם מכירה בשנים הקודמות',
  R.n>0&&one.length>0&&two.length>0,`${R.n} מועמדים · חודש אחד ${one.length} · 2+ ${two.length} · חודש אחד + שנים קודמות ${prior.length}`);
 const mism=R.rows.filter(x=>x.one!==x.has||x.one!==x.flag);
 ok('הסיבה «חודש אחד» — בדיוק במועמדים שהמימוש העצמאי מוצא (חודש פעיל אחד, בלי מכירה בשנים הקודמות)',!mism.length,mism.slice(0,3).map(x=>`${x.pn} act=${x.act} prior=${x.prior}`).join(', '));
 ok('מועמד בחודש אחד — «לבדוק», והסיבה בבועית',one.every(x=>x.grp==='rev'&&x.st==='rev'&&x.why.includes('צריכה בחודש אחד בלבד')),
  one.filter(x=>x.grp!=='rev'||!x.why.includes('צריכה בחודש אחד')).map(x=>x.pn).slice(0,3).join(', '));
 ok('הסיבה אחרונה בסדר — בפריט שכבר «לבדוק» הסיבה הראשית לא מתחלפת',one.every(x=>x.last));
 ok('יש מועמד שעבר ל«לבדוק» רק בגלל הכלל הזה',R.rows.some(x=>x.only),R.rows.filter(x=>x.only).length);
 ok('שני חודשים ומעלה — אין הסיבה',two.every(x=>!x.has));
 ok('חודש אחד אבל מכירה בשנים הקודמות — אין הסיבה (הביקוש חוזר)',prior.every(x=>!x.has));
 ok('השדה oneMonth רק בפריט PD',R.notPD===0,R.notPD);

 /* השנים הקודמות — על השורה הגולמית של מועמד בחודש אחד, דרך build() עצמו */
 const Y=await p.evaluate(pn=>{const m=mapCols(ALL_RAW.hdr),row=ALL_RAW.rows.find(x=>CLEAN(x[m.pn])===pn);
  const run=f=>{const c=row.slice();f(c);const rows0=Object.assign({},ROWS),d0=DIAG;let it;try{it=build([c],ALL_RAW.hdr)[0]}finally{Object.assign(ROWS,rows0);DIAG=d0}return !!(it&&it.oneMonth)};
  const An=ANCHOR.getMonth()+1,ms=ALL.find(x=>x.pn===pn).months.map(v=>Math.max(0,+v||0)),inWin=ms.slice(Math.min(An,11),11).reduce((a,v)=>a+v,0);
  return {base:run(()=>{}),y2:run(c=>{c[m.y2]=5}),y1out:run(c=>{c[m.y1]=inWin+4}),y1in:run(c=>{c[m.y1]=inWin})}},one[0].pn);
 ok('חודש אחד + מכירה לפני שנתיים — אין הסיבה (הביקוש חוזר)',Y.base&&!Y.y2,JSON.stringify(Y));
 ok('חודש אחד + מכירה בשנה שעברה מחוץ לחלון — אין הסיבה',Y.base&&!Y.y1out);
 ok('השנה שעברה רק עם החודשים שבתוך החלון — עדיין «חודש אחד» (לא נספר פעמיים)',Y.y1in);

 /* 4 הנושאים: אותו סטטוס בעמודה «שינוי פרופיל» */
 const T=await p.evaluate(()=>{const L=ALL.filter(r=>r.oneMonth&&profSplit(r)&&profSplit(r).reasons.length===1);const col=TOPICS.fix.cols.find(c=>c.k==='prof');
  return L.map(r=>({v:col.v(r),act:tpTopicAct('fix',r)}))});
 ok('4 הנושאים: «שינוי פרופיל» — «לבדוק», והפעולה «לבדוק פרופיל»',T.length>0&&T.every(x=>/לבדוק/.test(x.v)&&/^לבדוק פרופיל/.test(x.act)),JSON.stringify(T.slice(0,2)));

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
