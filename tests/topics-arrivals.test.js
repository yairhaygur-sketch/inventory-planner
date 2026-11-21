/* ============ ARRIVALS + FIRST_USE (החלטות המעתד 18.11) ============
   ARRIVALS — לוח הגעות לפי תאריך (רק פריטים שבארבעת הנושאים), וקיבוץ המשלוחים
   בכרטיס לפי תאריך. FIRST_USE — «צריכה ראשונה»: פריט שלא נצרך מעולם (11 חודשים
   ו-3 שנים) ויש לו צריכה בחודש המלא האחרון או בחודש הנוכחי. תצוגה בלבד. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx'),ETAX=path.join(SD,'zmrp-demo-eta.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');

 /* 1 · FIRST_USE — הכלל מול מימוש עצמאי, ורשומות גבול */
 const F=await p.evaluate(()=>{
  const ref=r=>r.has3y&&r.months.length&&(Math.max(0,r.months[0])+Math.max(0,r.curMonth||0))>0&&r.months.slice(1).every(v=>!(v>0))
   &&r.cons3y-Math.max(0,r.months[0])-Math.max(0,r.curMonth||0)<=0;
  const diff=QF.all.filter(r=>!!ref(r)!==tpFirst(r)).map(r=>r.pn);
  const base={has3y:true,months:Array(11).fill(0),curMonth:0,cons3y:0};
  const t=o=>tpFirst({...base,...o});
  return {n:QF.all.filter(tpFirst).length,diff,
   last:t({months:[4,...Array(10).fill(0)],cons3y:4}),cur:t({curMonth:2}),
   older:t({months:[4,0,0,1,...Array(7).fill(0)],cons3y:5}),hist:t({curMonth:2,cons3y:7}),none:t({}),no3y:t({curMonth:2,has3y:false}),
   cats:TOPICS.trend.cats.map(c=>c[1]),inRows:QF.all.filter(tpFirst).every(r=>TOPICS.trend.rows().includes(r))}});
 ok('צריכה ראשונה = מימוש עצמאי, בכל פריטי הדוח',F.n>0&&!F.diff.length,`${F.n} פריטים · ${F.diff.slice(0,3).join(',')}`);
 ok('צריכה בחודש המלא האחרון בלבד / בחודש הנוכחי בלבד → צריכה ראשונה',F.last&&F.cur);
 ok('צריכה מוקדמת יותר ב-11 החודשים, או ב-3 השנים → לא',!F.older&&!F.hist);
 ok('בלי צריכה, או בלי עמודות 3 השנים → לא',!F.none&&!F.no3y);
 ok('תת-קטגוריה «צריכה ראשונה» בחריגת צריכה, וכל הפריטים בה',F.cats.join('|')==='ביקוש עולה|ביקוש יורד|צריכה ראשונה'&&F.inRows,F.cats.join(' / '));
 await p.evaluate(()=>{TP.tab='trend';TP.cat={trend:'first'};TP.sort={};tpRender()});
 const T=await p.evaluate(()=>{const ths=[...document.querySelectorAll('#lite thead th')].map(t=>t.dataset.tpSort),ia=ths.indexOf('act');
  const trs=[...document.querySelectorAll('#lite tbody tr')];
  return {n:trs.length,bad:trs.filter((tr,j)=>{const r=TP.shown[j];return !tpFirst(r)||(!(riseGap(r)>0)&&!(trendGap(r)>0)&&tr.children[ia].textContent.trim()!=='לבדוק אם להגדיר תכנון לפריט')}).length}});
 ok('«צריכה ראשונה»: כל השורות, והפעולה «לבדוק אם להגדיר תכנון לפריט» בלי מספר',T.n>0&&T.bad===0,`${T.n} שורות`);
 const C=await p.evaluate(()=>{const r=TP.shown.find(r=>!(riseGap(r)>0)&&!(trendGap(r)>0));tpdShow(r);
  const t=document.querySelector('#tpd .qa .big').textContent;tpdClose();return t});
 ok('כרטיס: «צריכה ראשונה: …» ו«לא נרשמה צריכה קודמת»',/^צריכה ראשונה: /.test(C)&&/לא נרשמה צריכה קודמת/.test(C),C.slice(0,90));

 /* 2 · לוח הגעות בלי דוח ETA */
 await p.evaluate(()=>{TP.cat={};TP.view='arr';tpRender()});
 ok('בלי דוח ETA — הודעה וכפתור העלאה, בלי טבלה',await p.evaluate(()=>!!document.querySelector('#lite .etab')&&!document.querySelector('#lite details.arr')));

 /* 3 · עם דוח ETA */
 await p.setInputFiles('#fe',ETAX);await p.waitForTimeout(2500);
 await p.evaluate(()=>{TP.view='arr';tpRender()});
 const A=await p.evaluate(()=>{const D=tpArrData(),inTopic=new Set(Object.keys(TOPICS).flatMap(k=>tpRows(k,true)));
  /* מימוש עצמאי: לכל פריט בנושא — כל שורות ה-ETA שלו, מקובצות לפי תאריך */
  const exp={};for(const r of inTopic){for(const [q,d] of (ETA.byPn[etaKey(r.pn)]||[])){if(!d)continue;(exp[d]=exp[d]||{q:0,s:new Set()}).q+=q;exp[d].s.add(r)}}
  const det=[...document.querySelectorAll('#lite details.arr')];
  return {days:D.length,expDays:Object.keys(exp).length,sorted:D.every((x,i)=>!i||D[i-1].d<=x.d),
   match:D.every(x=>exp[x.d]&&exp[x.d].q===x.q&&exp[x.d].s.size===x.n),onlyTopics:D.every(x=>x.ls.every(e=>inTopic.has(e.r))),
   pastFirst:(()=>{const f=det.findIndex(d=>!d.classList.contains('past'));return f<0||det.slice(f).every(d=>!d.classList.contains('past'))})(),
   dom:det.length,firstQ:det[0]?det[0].querySelector('summary').textContent:''}});
 ok('לוח הגעות: תאריך לכל תאריך בדוח ה-ETA של פריטי הנושאים, ממוין',A.days>0&&A.days===A.expDays&&A.sorted&&A.dom===A.days,`${A.days} תאריכים`);
 ok('כל תאריך: סכום היחידות ומספר הפריטים = מימוש עצמאי',A.match);
 ok('רק פריטים שבארבעת הנושאים, ותאריכים שעברו בראש',A.onlyTopics&&A.pastFirst);
 /* לחיצה על מק״ט — הכרטיס בנושא הראשון */
 const K=await p.evaluate(()=>{const d=document.querySelector('#lite details.arr');d.open=true;const btn=d.querySelector('[data-tp="arrpn"]');
  const key=btn.dataset.key,t=btn.dataset.t;btn.click();return {key,t,tab:TP.tab,view:TP.view,open:TP.open?markKey(TP.open):null}});
 ok('לחיצה על מק״ט: חזרה לנושא הראשון שלו והכרטיס נפתח',K.view===''&&K.tab===K.t&&K.open===K.key,JSON.stringify(K));
 /* סינון מתכנן חל גם על הלוח */
 const P=await p.evaluate(()=>{tpdClose();const pl=[...new Set(QF.all.map(r=>r.planner).filter(Boolean))][0];TP.view='arr';TP.pl=pl||'';tpRender();
  const ok=tpArrData().every(x=>x.ls.every(e=>!pl||e.r.planner===pl));TP.pl='';return ok});
 ok('המתכנן שנבחר חל גם על לוח ההגעות',P);

 /* 4 · כרטיס: משלוחים מקובצים לפי תאריך */
 const G=await p.evaluate(()=>{TP.view='';tpRender();const r=QF.all.find(r=>(ETA.byPn[etaKey(r.pn)]||[]).length);if(!r)return null;
  /* בדוח ההדגמה אין פריט עם כמה שורות — מוסיפים זמנית: אותו תאריך פעמיים (שני משלוחים) ועוד תאריך */
  /* d1 — היום שאחרי d0. היה «ה-28 באותו חודש», ונפל כש-d0 עצמו יצא ה-28 (תאריכי דוח ה-ETA
     בהדגמה נגזרים מיום יצירת הקובץ): שני התאריכים התלכדו לשורה אחת. נמדד ב-21.11 */
  const k=etaKey(r.pn),orig=ETA.byPn[k],[q0,d0]=orig[0],t1=new Date(d0+'T00:00:00Z');t1.setUTCDate(t1.getUTCDate()+1);const d1=t1.toISOString().slice(0,10);
  ETA.byPn[k]=[[q0,d0,'D-1',''],[3,d0,'D-2',''],[2,d1,'D-3','']];
  tpdShow(r);const li=[...document.querySelectorAll('#tpd ul.eta li')].map(x=>x.textContent);tpdClose();ETA.byPn[k]=orig;
  const same=li.find(t=>t.includes(tpdDate(d0)))||'';
  return {li:li.length,same,sumOk:same.includes(tpN(q0+3)+' יח׳')&&/2 שורות/.test(same)&&/D-1/.test(same)&&/D-2/.test(same)}});
 ok('כרטיס: שורה אחת לכל תאריך — סכום היחידות, מספר השורות ומספרי המשלוחים',G&&G.li===2&&G.sumOk,G?G.same.slice(0,80):'אין פריט עם שורות ETA');

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0);
})().catch(e=>{console.error(e);process.exit(1)});
