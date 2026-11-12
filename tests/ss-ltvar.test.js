/* ============ SS_LTVAR — סיכון הספק ורמת שירות לפריט פעיל (אושר 12.11) ============
   ההחלטה (docs/BUSINESS-DECISIONS.md, docs/ROP-SS-POLICY.md):
   · המרכיב הסטטיסטי של מלאי הביטחון מחבר את תנודת הביקוש בזמן האספקה ואת
     סיכון הספק: ⌈ Z × √( LT·σ² + (קצב · p · LT)² ) ⌉, p = 25% בייבוא, 10% ברכש
     מקומי, וסוג חומר לא מוכר — כייבוא. חריגה: פריט בלי צריכה ב-11 החודשים
     (קצב מהיסטוריית 3 השנים בלבד) — בלי סיכון ספק, כמו קודם. ופריט שזוהתה בו
     דעיכה — סיכון הספק וסף «פריט פעיל» על הקצב העדכני.
   · פריט פעיל (5 יח׳ בחודש ומעלה) מקבל לפחות 95%; רמה גבוהה יותר מ-SAP נשמרת.
   · הרצפות לא זזות, ולכן מלאי ביטחון מוצע אינו יורד באף פריט.
   הבדיקה מחשבת את הנוסחה מחדש על כל פריט בפיקסצ׳ר — לא על דוגמה — ומוודאת
   שיש פריטים בכל ענף (אחרת היא אינה בודקת דבר). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML=process.argv[2]||path.join(SD,'..','index.html');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+HTML+'?nobrief=1');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});
 const R=await p.evaluate(()=>{
  const P={'יבוא':0.25,'רכש מקומי':0.10},Z95=zFor(95);
  const recOK=r=>!(r.isODStatic||r.noStock);
  const live=ALL.filter(r=>recOK(r)&&!r.ltBad);
  const want=r=>{const p=P[r.matType]??0.25,ltM=Math.max(r.ltM,0);
   const q=r.rateRecent!=null&&r.rateRecent<RATE_DECAY*r.rate3y?r.rateRecent:r.rate;
   const sd=r.A.sd*Math.sqrt(ltM),sup=r.d12>0?q*p*ltM:0,z=q>=5?Math.max(r.z,Z95):r.z;
   return r.rate>0?Math.ceil(z*Math.sqrt(sd*sd+sup*sup)):0};
  const old=r=>r.rate>0?Math.ceil(r.z*r.A.sd*Math.sqrt(Math.max(r.ltM,0))):0;
  /* ענפים שאין להם פריט בפיקסצ׳ר (רמות שירות 50/80/95 בלבד; Z001/Z004 בלבד) —
     דרך build() על שורה אמיתית של פריט פעיל, עם תא אחד שונה */
  const hdr=ALL_RAW.hdr,m=mapCols(hdr);
  const src=live.find(r=>r.rate>=5&&r.ltM>0);
  const row=ALL_RAW.rows.find(x=>CLEAN(x[m.pn])===src.pn);
  const mk=(srv,mt)=>{const c=row.slice();c[m.srv]=srv;if(mt!=null)c[m.matType]=mt;return build([c],hdr)[0]};
  const s99=mk(99),s80=mk(80),unk=mk(80,'Z999');
  const br={srcPn:src.pn,z99:s99.zEff,act99:s99.srvAct,z80:s80.zEff,act80:s80.srvAct,pUnk:unk.ltvP,mtUnk:unk.matType};
  const pick=live.find(r=>r.srvAct&&r.ssSrc==='stat'&&r.sdSup>0);
  let card=null;
  if(pick){explain(pick);const t=document.getElementById('mbody').textContent;
   card={pn:pick.pn,supply:/סיכון הספק/.test(t),total:/סיכון כולל/.test(t),srv:t.includes(`${SS_ACT_SRV}%`),
    val:t.includes(`= ${pick.ssStat}`),formula:/√\(LT·σ² \+ \(קצב·σ_LT\)²\)/.test(document.querySelector('#mbody .formula')?.textContent||'')};
   document.getElementById('modalwrap').classList.remove('open')}
  return {n:live.length,
   formula:live.filter(r=>r.ssStat!==want(r)).map(r=>`${r.pn}: ${r.ssStat}≠${want(r)}`),
   imp:live.filter(r=>r.matType==='יבוא').length,loc:live.filter(r=>r.matType==='רכש מקומי').length,
   pOK:live.filter(r=>r.ltvP!==(P[r.matType]??0.25)).map(r=>r.pn),
   act:live.filter(r=>r.srvAct).length,
   actOK:live.filter(r=>r.srvAct!==(r.rateSS>=5&&Z95>r.z)).map(r=>r.pn),
   rssOK:live.filter(r=>r.rateSS!==(r.rateRecent!=null&&r.rateRecent<RATE_DECAY*r.rate3y?r.rateRecent:r.rate)).map(r=>r.pn),
   zEffOK:live.filter(r=>r.zEff!==(r.srvAct?Z95:r.z)).map(r=>r.pn),
   keepHigher:live.filter(r=>r.rateSS>=5&&r.z>Z95).filter(r=>r.zEff!==r.z).map(r=>r.pn),
   slowSap:live.filter(r=>r.rateSS>0&&r.rateSS<5&&r.zEff!==r.z).map(r=>r.pn),
   raised:live.filter(r=>r.ssStat>old(r)).length,
   hist:live.filter(r=>r.rate>0&&!(r.d12>0)).length,
   histOK:live.filter(r=>r.rate>0&&!(r.d12>0)).filter(r=>r.sdSup!==0||r.ssStat!==old(r)).map(r=>r.pn),
   noDrop:live.filter(r=>r.sugSS<Math.max(old(r),r.ssMinEff||0,r.lowFloor||0)).map(r=>r.pn),
   floors:live.filter(r=>r.sugSS<Math.max(r.ssStat,r.ssMinEff||0,r.lowFloor||0)&&r.sugSS!==ROP_MIN_ACT).map(r=>r.pn),
   noRecZero:ALL.filter(r=>!recOK(r)&&r.ssStat!==0).map(r=>r.pn),
   card,br}});
 ok('יש פריטים מנוהלים בפיקסצ׳ר',R.n>0,R.n);
 ok('כל פריט: SS סטטיסטי = ⌈Z × √(LT·σ² + (קצב·p·LT)²)⌉',!R.formula.length,R.formula.slice(0,5).join(' | '));
 ok('שני סוגי החומר קיימים בפיקסצ׳ר (25% ו-10% נבדקים שניהם)',R.imp>0&&R.loc>0,`יבוא ${R.imp} · מקומי ${R.loc}`);
 ok('p נקבע לפי סוג החומר: 25% יבוא, 10% מקומי',!R.pOK.length,R.pOK.join(','));
 ok('יש פריטים פעילים שרמת השירות שלהם הועלתה ל-95%',R.act>0,R.act);
 ok('הועלתה בדיוק כשהקצב ≥5 ו-SAP מתחת ל-95%',!R.actOK.length&&!R.zEffOK.length,[...R.actOK,...R.zEffOK].join(','));
 ok('הקצב לסיכון הספק ולסף: הקצב העדכני כשזוהתה דעיכה, אחרת rate',!R.rssOK.length,R.rssOK.join(','));
 ok('רמת SAP גבוהה מ-95% נשמרת',!R.keepHigher.length,R.keepHigher.join(','));
 ok('מתחת ל-5 יח׳ בחודש — רמת השירות מ-SAP כמו קודם',!R.slowSap.length,R.slowSap.join(','));
 ok('השינוי מעלה מרכיב סטטיסטי בפריטים בפיקסצ׳ר',R.raised>0,R.raised);
 ok('חריגה: יש בפיקסצ׳ר פריטים עם קצב בלי צריכה ב-11 חודשים (קצב 3 שנים)',R.hist>0,R.hist);
 ok('חריגה: בלי צריכה ב-11 חודשים — אין סיכון ספק, והמרכיב הסטטיסטי כמו קודם',!R.histOK.length,R.histOK.join(','));
 ok('אף מלאי ביטחון מוצע אינו יורד מתחת לנוסחה הקודמת והרצפות',!R.noDrop.length,R.noDrop.join(','));
 ok('הרצפות לא זזו: SS מוצע ≥ max(סטטיסטי, ssMin, רצפת מדיניות)',!R.floors.length,R.floors.join(','));
 ok('פריט בלי המלצה (OD/PD/04-14) — המרכיב הסטטיסטי 0',!R.noRecZero.length,R.noRecZero.join(','));
 ok('build: רמת שירות 99% ב-SAP נשמרת (Z 2.33, בלי העלאה)',R.br.z99===2.33&&!R.br.act99,JSON.stringify(R.br));
 ok('build: אותו פריט ב-80% — מועלה ל-95% (Z 1.65)',R.br.z80===1.65&&R.br.act80,JSON.stringify(R.br));
 ok('build: סוג חומר לא מוכר — p = 25%, כמו ייבוא',R.br.pUnk===0.25,JSON.stringify(R.br));
 ok('הכרטיס «מאחורי החישוב» מציג סיכון ספק, סיכון כולל, 95% והתוצאה',!!R.card&&R.card.supply&&R.card.total&&R.card.srv&&R.card.val&&R.card.formula,JSON.stringify(R.card));
 ok('0 שגיאות JS',!errs.length,errs.join('|'));
 await b.close();console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})();
