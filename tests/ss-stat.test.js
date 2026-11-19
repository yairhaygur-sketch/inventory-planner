/* ============ SS_STAT — מלאי ביטחון סטטיסטי בפריט עם צריכה (החלטת המעתד 19.11) ============
   פריט עם צריכה עדכנית (נתיב 6 החודשים, קצב >0): SS = max(סטטיסטי, רצפת מדיניות) — רצפת
   ssMin מ-SAP אינה קובעת. פריט בלי צריכה ב-6 החודשים (קצב 0, או קצב מ-3 השנים) — רצפת
   SAP נשארת. בקובץ העדכון ל-SAP: «מלאי ביטחון מינימלי» = ה-SS המוצע כשהוא נמוך מ-ssMin,
   ושורה נכנסת גם כשרק המינימום משתנה. ROP_MIN_ACT ו-SS≤ROP — כמו קודם. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.setInputFiles('#f',path.join(SD,'zmrp-demo.xlsx'));
 await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB,null,{timeout:180000});

 /* 1 · דרך build(): פריטים מאותה כותרת של הדוח שנטען */
 const B=await p.evaluate(()=>{
  const hdr=ALL_RAW.hdr,m=mapCols(hdr),tpl=ALL_RAW.rows[0];
  const mk=(pn,o)=>{const r=tpl.slice();const set=(i,v)=>{if(i>=0)r[i]=v};
   set(m.pn,pn);set(m.status,o.st||'01');set(m.statusTxt,'פעיל');set(m.mrp,'ND');set(m.srv,80);
   set(m.ss,o.ss||0);set(m.rop,o.rop||0);set(m.ssMin,o.ssMin||0);set(m.lt,60);set(m.ltAdd,0);
   set(m.price,o.price||100);set(m.currency,'USD');set(m.free,o.free??200);set(m.stock,o.free??200);set(m.po,0);
   if(m.openDel>=0)set(m.openDel,0);if(m.transfer>=0)set(m.transfer,0);if(m.cust>=0)set(m.cust,0);
   m.months.forEach((ix,k)=>set(ix,o.m[k]||0));
   [m.y0,m.y1,m.y2].forEach((ix,k)=>set(ix,o.y[k]));
   if(m.curMonth>=0)set(m.curMonth,0);
   return r};
  const one=o=>build([mk('S-'+Math.random().toString(36).slice(2,7),o)],hdr)[0];
  const sum=a=>a.reduce((x,y)=>x+y,0),An=Math.min(ANCHOR.getMonth()+1,11);
  const yrs=(mo,y1,y2)=>[sum(mo.slice(0,An)),y1,y2];
  const reg=Array(11).fill(6),silent=[0,0,0,0,0,0,6,6,6,6,6],sparse=[0,0,0,0,1,0,0,0,0,1,0];
  const pick=x=>({rate:x.rate,from3y:!!x.rateFrom3y,ssStat:x.ssStat||0,ssMin:x.ssMin,ssMinEff:x.ssMinEff,lowFloor:x.lowFloor||0,
   sugSS:x.sugSS,sugROP:x.sugROP,dLT:x.dLT,ssSrc:x.ssSrc,statOnly:!!x.ssStatOnly,fileMin:sapSsMin(x)});
  return {
   reg:pick(one({m:reg,y:yrs(reg,72,72),ssMin:40,ss:40,rop:60,free:30})),     /* צריכה סדירה, רצפה 40 */
   regLow:pick(one({m:reg,y:yrs(reg,72,72),ssMin:1,ss:1,rop:20,free:30})),      /* רצפה נמוכה מהסטטיסטי */
   silent:pick(one({m:silent,y:yrs(silent,72,72),ssMin:40,ss:40,rop:60})),      /* שקט 6 חודשים */
   from3y:pick(one({m:sparse,y:yrs(sparse,60,60),ssMin:9,ss:9,rop:12})),        /* קצב מ-3 השנים */
   st04:pick(one({m:reg,y:yrs(reg,72,72),ssMin:40,st:'04'})),
   cheap:pick(one({m:[0,1,0,1,0,1,0,1,0,1,0],y:yrs([0,1,0,1,0,1,0,1,0,1,0],6,6),ssMin:8,ss:8,rop:10,price:3,free:4}))}});
 const {reg,regLow,silent,from3y,st04,cheap}=B;
 ok('צריכה סדירה, ssMin=40: SS = הסטטיסטי, הרצפה מ-SAP אינה קובעת',reg.rate>0&&reg.statOnly&&reg.ssMinEff===0&&reg.sugSS===Math.max(reg.ssStat,reg.lowFloor)&&reg.sugSS<40&&reg.ssSrc!=='sap',
   `SS ${reg.sugSS} · סטטיסטי ${reg.ssStat} · מקור ${reg.ssSrc}`);
 ok('ROP = ⌈ביקוש_LT + SS⌉ על ה-SS הסטטיסטי',reg.sugROP===Math.max(reg.sugSS,Math.ceil(reg.dLT+reg.sugSS)),`ROP ${reg.sugROP}`);
 ok('ssMin עצמו נשאר הערך מהדוח (40)',reg.ssMin===40);
 ok('קובץ SAP: המינימום המוצע = ה-SS המוצע כשהוא נמוך מ-ssMin',reg.fileMin===reg.sugSS,`${reg.fileMin}`);
 ok('רצפה נמוכה מהסטטיסטי: אותו SS, והמינימום בקובץ נשאר הערך מ-SAP',regLow.sugSS===regLow.ssStat&&regLow.fileMin===1,`SS ${regLow.sugSS} · מינימום ${regLow.fileMin}`);
 ok('שקט 6 חודשים (קצב 0): רצפת SAP נשארת — SS ≥ 40',silent.rate===0&&!silent.statOnly&&silent.ssMinEff===40&&silent.sugSS>=40&&silent.ssSrc==='sap',`SS ${silent.sugSS} · ${silent.ssSrc}`);
 ok('קצב מהיסטוריית 3 השנים: רצפת SAP נשארת',from3y.from3y&&!from3y.statOnly&&from3y.ssMinEff===9&&from3y.sugSS>=9,`SS ${from3y.sugSS}`);
 ok('סטטוס 04: אין המלצה',st04.sugSS===0&&st04.sugROP===0&&!st04.statOnly);
 ok('פריט זול בביקוש זעום עם צריכה: רצפת המדיניות נשארת',cheap.statOnly&&cheap.sugSS===Math.max(cheap.ssStat,cheap.lowFloor)&&cheap.sugSS<8,
   `SS ${cheap.sugSS} · סטטיסטי ${cheap.ssStat} · רצפת מדיניות ${cheap.lowFloor}`);

 /* 2 · sapUsableBase: שורה נכנסת גם כשרק המינימום משתנה */
 const U=await p.evaluate(()=>{const base={paramFix:true,ltBad:false,ltD:60,sugSS:10,sugROP:20,ss:10,rop:20};
  return {minOnly:sapUsableBase({...base,ssMin:40}),same:sapUsableBase({...base,ssMin:5}),row:sapRowsBy([{...base,ssMin:40,pn:'X',site:'1'}],()=>true)[1]}});
 ok('SS ו-ROP כבר כמו בהמלצה, אבל ssMin=40 מעליהם — השורה נכנסת',U.minOnly===true);
 ok('בלי שינוי בשום שדה — אין שורה',U.same===false);
 ok('השורה: «מלאי בטחון» 10 · «מלאי ביטחון מינימלי» 10 (לא 40) · ROP 20',U.row[2]===10&&U.row[3]===10&&U.row[4]===20,JSON.stringify(U.row));

 /* 3 · הדוח הסינתטי: הכלל חל בדיוק על פריטים עם צריכה עדכנית, וכל שורת SAP: SS ≥ המינימום שבה */
 const D=await p.evaluate(()=>{const live=ALL.filter(r=>!r.isOD&&!r.noStock&&!r.ltBad);
  const bad=live.filter(r=>{const so=r.rate>0&&!r.rateFrom3y;return so?r.ssMinEff!==0:r.ssMinEff!==r.ssMin}).map(r=>r.pn);
  const rows=[...sapMassRows(QF.all),...sapSlowRows(QF.all)].filter(r=>r[0]!=='חומר ');
  const byPn=new Map(QF.all.map(r=>[String(r.pn),r]));
  return {n:live.length,so:live.filter(r=>r.rate>0&&!r.rateFrom3y&&r.ssMin>r.sugSS).length,bad,rows:rows.length,
   viol:rows.filter(x=>!(x[2]>=x[3])).length,minOk:rows.every(x=>{const r=byPn.get(String(x[0]));return r&&x[3]===Math.min(r.ssMin||0,r.sugSS)}),
   lowered:rows.filter(x=>{const r=byPn.get(String(x[0]));return r&&x[3]<(r.ssMin||0)}).length}});
 ok('בדוח הסינתטי: רצפת SAP 0 בדיוק לפריטים עם צריכה עדכנית, וכמו קודם לכל השאר',D.bad.length===0,`${D.n} פריטים · ${D.so} עם SS מתחת ל-ssMin · ${D.bad.slice(0,3).join(',')}`);
 ok('קובצי SAP: בכל שורה SS ≥ המינימום שבה, והמינימום = min(ssMin, SS המוצע)',D.rows>0&&D.viol===0&&D.minOk,`${D.rows} שורות · ${D.lowered} עם מינימום מוקטן`);

 /* 4 · «מאחורי החישוב»: אומר מה הערך ב-SAP ולמה אינו רצפה */
 const M=await p.evaluate(()=>{const r=ALL.find(x=>x.ssStatOnly&&x.sugSS<x.ssMin&&!x.ltBad);if(!r)return null;explain(r);
  const t=document.getElementById('mbody').innerText;document.getElementById('modalwrap').classList.remove('open');
  return {ok:t.includes(`${r.ssMin} ב-SAP → 0`)&&t.includes('הרצפה מ-SAP אינה קובעת')&&t.includes(`יוצע מינימום ${r.sugSS}`),pn:r.pn}});
 ok('«מאחורי החישוב»: «X ב-SAP → 0 · … אינה קובעת · יוצע מינימום Y»',M&&M.ok,M?M.pn:'אין פריט כזה בדוח');

 /* 5 · כרטיס ארבעת הנושאים: «מלאי ביטחון מינימלי» — הערך ב-SAP, למה אינו קובע, והמינימום שיוצע בקובץ */
 const T=await p.evaluate(()=>{const r=ALL.find(x=>x.ssStatOnly&&x.sugSS<x.ssMin&&!x.ltBad&&!recNA(x));if(!r)return null;
  const h=tpdParams(r),d=document.createElement('div');d.innerHTML=h;const tr=[...d.querySelectorAll('tr')].find(t=>/מלאי ביטחון מינימלי/.test(t.textContent));
  return {txt:tr?tr.textContent.replace(/\s+/g,' '):'',want:sapSsMin(r),ssMin:r.ssMin}});
 ok('כרטיס: «מלאי ביטחון מינימלי» — SAP, «לא קובע», והמינימום בקובץ SAP',T&&T.txt.includes(String(T.ssMin))&&/לא קובע/.test(T.txt)&&T.txt.includes(`${T.want} בקובץ SAP`),T?T.txt:'אין פריט');
 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0);
})().catch(e=>{console.error(e);process.exit(1)});
