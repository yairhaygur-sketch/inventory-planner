/* ============ RATE_W — קצב משוקלל לפריט גושי או ספורדי (אישור המעתד 18.11) ============
   70% קצב המנוע (6 החודשים המנוקים) · 20% יתרת השנה הקודמת (בלי החודשים שבחלון)
   · 10% לפני שנתיים. רק גושי/ספורדי, רק בנתיב 6 החודשים, רק עם שלוש עמודות
   שנתיות, לא בירידה חדה ולא בלי המלצה. σ ורצפות ה-SS — כמו קודם; rate לא זז. */
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

 const B=await p.evaluate(()=>{
  const hdr=ALL_RAW.hdr,m=mapCols(hdr),tpl=ALL_RAW.rows[0];
  const mk=(pn,o)=>{const r=tpl.slice();const set=(i,v)=>{if(i>=0)r[i]=v};
   set(m.pn,pn);set(m.status,o.st||'01');set(m.statusTxt,'פעיל');set(m.mrp,'ND');set(m.srv,80);
   set(m.ss,o.ss||0);set(m.rop,o.rop||0);set(m.ssMin,o.ssMin||0);set(m.lt,90);set(m.ltAdd,0);
   set(m.price,100);set(m.currency,'USD');set(m.free,o.free||200);set(m.stock,o.free||200);set(m.po,0);
   if(m.openDel>=0)set(m.openDel,0);if(m.transfer>=0)set(m.transfer,0);if(m.cust>=0)set(m.cust,0);
   m.months.forEach((ix,k)=>set(ix,o.m[k]||0));
   [m.y0,m.y1,m.y2].forEach((ix,k)=>set(ix,o.y[k]));
   if(m.curMonth>=0)set(m.curMonth,0);
   return r};
  const one=(pn,o,h)=>build([mk(pn,o)],h||hdr)[0];
  const An=ANCHOR.getMonth()+1;
  const exp=x=>{const raw=x.months.map(v=>Math.max(0,v)),lastIn=Math.max(0,11-An);
   const ry1=Math.max(0,Y[1]-raw.slice(Math.min(An,11),11).reduce((a,b)=>a+b,0))/(12-lastIn),ry2=Y[2]/12;
   return 0.7*x.rate+0.2*ry1+0.1*ry2};
  const lumpyM=[0,12,0,0,3,0,0,15,0,1,0],sporM=[0,4,0,0,5,0,0,4,0,5,0],smoothM=Array(11).fill(6);
  const sum=a=>a.reduce((x,y)=>x+y,0);
  /* השנה הקודמת ולפני שנתיים — גבוהות מהשנה, כדי שהשקלול יזיז */
  const Y=[sum(lumpyM.slice(0,Math.min(An,11))),120,96];
  const L=one('W-LUMPY',{m:lumpyM,y:Y}),Sp=one('W-SPOR',{m:sporM,y:[sum(sporM.slice(0,Math.min(An,11))),120,96]}),
   Sm=one('W-SMOOTH',{m:smoothM,y:[sum(smoothM.slice(0,Math.min(An,11))),120,96]});
  const hdrNo=hdr.slice();hdrNo[m.y2]='עמודה אחרת';
  const No3=one('W-NO3Y',{m:lumpyM,y:Y},hdrNo);
  const St04=one('W-04',{m:lumpyM,y:Y,st:'04'});
  const Fl=one('W-FLOOR',{m:lumpyM,y:Y,ssMin:30});
  const pick=x=>({statOnly:!!x.ssStatOnly,pat:x.A.pattern,rate:x.rate,rateW:x.rateW??null,rateROP:x.rateROP??null,dLT:x.dLT,ltM:x.ltM,sugSS:x.sugSS,sugROP:x.sugROP,
   ssMin:x.ssMin,ssStat:x.ssStat,lowFloor:x.lowFloor||0,rateSS:x.rateSS,sdSup:x.sdSup,ltvP:x.ltvP??LTV_P_DEF,ropRate:ropRate(x),months:x.months,
   has3y:x.has3y,noW:!('rateW' in x)});
  const sp=pick(Sp);sp.exp=0.7*Sp.rate+0.2*(Math.max(0,120-Sp.months.slice(Math.min(An,11),11).reduce((a,v)=>a+Math.max(0,v),0))/(12-Math.max(0,11-An)))+0.1*96/12;
  return {L:{...pick(L),exp:exp(L)},Sp:sp,Sm:pick(Sm),No3:pick(No3),St04:pick(St04),Fl:pick(Fl),K:{w:RATE_W,pat:RATE_W_PAT}};
 });
 const near=(a,b)=>Math.abs(a-b)<1e-9;
 const {L,Sp,Sm,No3,St04,Fl,K}=B;
 ok('הקבועים: 70/20/10, גושי וספורדי',JSON.stringify(K.w)==='[0.7,0.2,0.1]'&&K.pat.join()==='גושי,ספורדי');
 ok('פריט גושי: rateW = 0.7·קצב + 0.2·יתרת השנה הקודמת + 0.1·(לפני שנתיים÷12)',L.pat==='גושי'&&L.rateW!=null&&near(L.rateW,L.exp),
   `${L.pat} · rate ${L.rate.toFixed(3)} · rateW ${L.rateW&&L.rateW.toFixed(3)} · ציפייה ${L.exp.toFixed(3)}`);
 ok('ביקוש בזמן אספקה = rateW × LT, ו-ROP קצר בודק על אותו קצב',near(L.dLT,L.rateW*L.ltM)&&near(L.rateROP,L.rateW)&&near(L.ropRate,L.rateW));
 ok('סיכון הספק על rateW',near(L.rateSS,L.rateW)&&near(L.sdSup,L.rateW*L.ltvP*L.ltM));
 ok('SS = max(סטטיסטי, ssMin, רצפה) · ROP = ⌈ביקוש_LT + SS⌉',L.sugSS===Math.max(L.ssStat,L.ssMin,L.lowFloor)&&L.sugROP===Math.max(L.sugSS,Math.ceil(L.dLT+L.sugSS)),`SS ${L.sugSS} · ROP ${L.sugROP}`);
 ok('rate עצמו לא זז (ממוצע 6 החודשים המנוקים)',L.rate>0&&!near(L.rate,L.rateW));
 ok('פריט ספורדי: אותו כלל',Sp.pat==='ספורדי'&&Sp.rateW!=null&&near(Sp.rateW,Sp.exp),`${Sp.pat} · ${Sp.rateW&&Sp.rateW.toFixed(3)}`);
 ok('פריט חלק — הכלל אינו חל, הנוסחה כמו קודם',Sm.pat==='חלק'&&Sm.noW&&near(Sm.dLT,Sm.rate*Sm.ltM),Sm.pat);
 ok('בלי שלוש עמודות שנתיות — הכלל אינו חל',No3.has3y===false&&No3.noW&&near(No3.dLT,No3.rate*No3.ltM));
 ok('סטטוס 04 — אין המלצה ואין שדה',St04.noW&&St04.sugROP===0);
 /* SS_STAT (19.11): לפני — «רצפת ssMin נשמרת (SS = 30)». עכשיו פריט עם צריכה: SS סטטיסטי */
 ok('ssMin=30 אינו רצפה לפריט עם צריכה: SS = max(סטטיסטי, רצפת מדיניות)',Fl.rateW!=null&&Fl.statOnly&&Fl.sugSS===Math.max(Fl.ssStat,Fl.lowFloor)&&Fl.sugSS<30&&Fl.sugROP===Math.max(Fl.sugSS,Math.ceil(Fl.dLT+Fl.sugSS)),`SS ${Fl.sugSS} · ROP ${Fl.sugROP}`);

 /* בדוח הסינתטי: כל פריט בלי rateW / דעיכה / ירידה חדה — dLT = rate × LT; לכל פריט עם rateW — הדפוס גושי/ספורדי */
 const D=await p.evaluate(()=>({n:ALL.length,w:ALL.filter(x=>x.rateW!=null).length,
  patOk:ALL.filter(x=>x.rateW!=null).every(x=>RATE_W_PAT.includes(x.A.pattern)&&x.has3y&&!x.sharpDrop&&!x.rateFrom3y),
  diff:ALL.filter(x=>x.rateW==null&&!x.rateDecay&&!x.sharpDrop&&Math.abs(x.dLT-x.rate*x.ltM)>1e-9).length}));
 ok('בדוח הסינתטי: rateW רק בגושי/ספורדי עם 3 שנים; בכל השאר dLT = rate × LT',D.w>0&&D.patOk&&D.diff===0,`${D.n} פריטים · ${D.w} עם קצב משוקלל`);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0);
})().catch(e=>{console.error(e);process.exit(1)});
