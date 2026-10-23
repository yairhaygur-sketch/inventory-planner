#!/usr/bin/env node
/* ============ צילום עסקי מלא — כל מה שהכלי מחשב ומוציא ============
   שימוש:
     node business-snapshot.js <index.html> <zmrp.xlsx> <out.json> [--eta <eta.xlsx>] [--now <ISO>] [--tz Asia/Jerusalem]

   מה נכנס לצילום:
     items    כל שדה מחושב של כל פריט ב-ALL, מיד אחרי הטעינה ולפני שנפתח
              מסך כלשהו — סיווג, חומרה, SS/ROP מוצעים, כמות להזמנה, חוסר,
              כסף תקוע, הנימוקים (why) והפעולה (act), והמלצת הפרופיל.
              לא רשימה סגורה של שדות: כל ערך פשוט (מספר, מחרוזת, בוליאני)
              וכל מערך של ערכים פשוטים. שדה חדש שנוסף או נעלם — נתפס.
     queues   כל תור ש-classify בונה (QF), וכל דלת (floor, rop, trend…).
     screens  לכל אחד מ-26 המסכים: המונה במסילה, רשימת המק״טים בסדר שלה,
              וסכומי כסף תקוע ויחידות שלא יסופקו — לפי מטבע, בלי סכימה.
     exports  לכל מסך: ייצוא לאקסל (כל הגיליונות, תא בתא), עדכון ל-SAP,
              מלאי איטי ל-SAP. ופעם אחת: דוח זירוז, וטיוטת המייל לספקים.

   השעון בדפדפן מקובע (--now, ברירת מחדל היום ב-08:00) — חלק מהחישוב
   תלוי בתאריך (ימים מאז, תאריך שעבר), ושתי ריצות באותו יום חייבות
   לצאת זהות. XLSX.writeFile מוחלף בלוכד: שום קובץ לא נכתב לדיסק.
   נכשל על שגיאת JS אחת.

   הצילום מכיל מק״טים אמיתיים כשמזינים לו את הדוח האמיתי — הוא נשמר
   בזיכרון הזמני בלבד, לעולם לא בריפו. */
const path=require('path'),fs=require('fs');
let PW;try{PW=require('playwright')}catch(_){PW=require(path.resolve(__dirname,'../../../../node_modules/playwright'))}
const args=process.argv.slice(2);
const opt=k=>{const i=args.indexOf(k);if(i<0)return null;const v=args[i+1];args.splice(i,2);return v};
const ETA=opt('--eta'),NOW=opt('--now'),TZ=opt('--tz')||'Asia/Jerusalem';
const [html,xlsx,out]=args;
if(!html||!xlsx||!out){console.error('שימוש: business-snapshot.js <index.html> <zmrp.xlsx> <out.json> [--eta <eta.xlsx>] [--now <ISO>]');process.exit(2)}
const now=NOW?new Date(NOW):(()=>{const d=new Date();d.setHours(8,0,0,0);return d})();

(async()=>{
 const b=await PW.chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 /* אזור זמן קבוע: תאריכים בכלי מחושבים לפי השעון המקומי, ובלי זה אותה
    ריצה במחשב וב-CI (UTC) הייתה נותנת תאריכים שונים בייצוא. */
 const ctx=await b.newContext({viewport:{width:1512,height:900},timezoneId:TZ});
 const p=await ctx.newPage();
 const errs=[];p.on('pageerror',e=>errs.push(e.message.split('\n')[0].slice(0,200)));
 p.on('dialog',d=>d.accept());
 await p.clock.setFixedTime(now);
 await p.goto('file://'+path.resolve(html)+'?nobrief=1');
 await p.setInputFiles('#f',path.resolve(xlsx));
 await p.waitForFunction(()=>typeof ALL!=='undefined'&&ALL&&ALL.length>0,null,{timeout:900000});
 await p.waitForTimeout(2500);
 if(ETA){await p.setInputFiles('#fe',path.resolve(ETA));await p.waitForTimeout(3000)}
 await p.click('#tourSkip').catch(()=>{});

 const snap=await p.evaluate(()=>{
  const R=v=>typeof v==='number'?(Number.isFinite(v)?Math.round(v*1e6)/1e6:String(v)):v;
  const prim=v=>v===null||['string','number','boolean'].includes(typeof v);
  /* ---- items: מיד אחרי הטעינה, לפני שמסך כלשהו נפתח ---- */
  const items={};
  for(const r of ALL){const o={};
   for(const k of Object.keys(r).sort()){const v=r[k];
    if(prim(v))o[k]=R(v);
    else if(Array.isArray(v)&&v.every(x=>prim(x)||(Array.isArray(x)&&x.every(prim))))
     o[k]=v.map(x=>Array.isArray(x)?x.map(R):R(x))}
   try{if(typeof profInfo==='function'){const pi=profInfo(r);o.__prof=pi?{cur:pi.cur,tgt:pi.tgt,h:pi.h,state:pi.state,why:pi.why}:null}}catch(e){o.__prof='ERR:'+e.message}
   items[String(r.pn)]=o}
  /* ---- queues + doors ---- */
  const queues={};
  for(const k of Object.keys(QF||{}).sort())if(Array.isArray(QF[k])&&k!=='all')queues['Q.'+k]=QF[k].map(x=>String(x.pn));
  const D={floor:()=>floorPick(floorAll(),'floor'),floorNR:()=>floorPick(floorAll(),'floorNR'),
   ropGap:()=>ropPick(ropAll(),'ropGap'),ropThin:()=>ropPick(ropAll(),'ropThin'),
   trend:()=>trendRows(),rise:()=>riseRows(),burn:()=>burnRows(),agenda:()=>agendaRows(),
   m1:()=>m1Rows(QF.all||[])};
  for(const k in D){try{queues['D.'+k]=(D[k]()||[]).map(x=>String(x.pn))}catch(e){queues['D.'+k]='ERR:'+e.message.slice(0,80)}}
  /* ---- screens + exports ---- */
  const cap=[];const orig=XLSX.writeFile;
  XLSX.writeFile=(wb,name)=>{const sheets={};
   for(const sn of wb.SheetNames)sheets[sn]=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:''}).map(row=>row.map(R));
   cap.push({name:String(name).replace(/_?\d{4}-\d{2}-\d{2}/g,'').replace(/_?\d{8}/g,''),sheets})};
  const grab=fn=>{cap.length=0;try{fn()}catch(e){return {err:e.message.slice(0,120)}}return cap.length?cap.slice():null};
  const sumBy=(L,f)=>{const m={};for(const r of L){const v=+f(r)||0;if(!v)continue;const c=r.currency||'—';m[c]=R((m[c]||0)+v)}return m};
  const screens={},exports={};
  for(const m of Object.keys(MODESET)){
   const s={};
   try{setMode(m)}catch(e){screens[m]={err:'setMode: '+e.message.slice(0,120)};continue}
   /* «תנועות אחרונות» אינו רשימת פריטים של המנוע אלא אריח משלו (movRows) */
   try{const L=m==='moves'&&typeof movRows==='function'?movRows():currentRows();s.count=L.length;s.pns=L.map(r=>String(r.pn));
    s.sums={expCap:sumBy(L,r=>r.expCap),unsVal:sumBy(L,r=>r.unsVal),unsQty:sumBy(L,r=>r.unsQty)}}
   catch(e){s.err='currentRows: '+e.message.slice(0,120)}
   try{if(typeof navCount==='function'){const nc=navCount(m);s.nav=nc&&typeof nc==='object'?nc.n:nc}}catch(e){s.nav='ERR'}
   screens[m]=s;
   exports[m]={excel:grab(()=>doExport()),sap:grab(()=>doSapMass()),slow:grab(()=>doSapMass('slow'))}}
  exports.__expedite=grab(()=>doExpedite());
  try{setMode('line');exports.__mail=lineMailText()||''}catch(e){exports.__mail={err:e.message}}
  XLSX.writeFile=orig;
  return {n:ALL.length,items,queues,screens,exports}});
 snap.now=now.toISOString();snap.errs=errs;snap.eta=!!ETA;
 fs.writeFileSync(out,JSON.stringify(snap));
 const nsc=Object.keys(snap.screens).length,nx=Object.values(snap.exports).filter(x=>x&&x.excel).length;
 console.log(`${path.basename(out)}: ${snap.n} פריטים · ${nsc} מסכים · ${nx} ייצואי אקסל · שגיאות JS: ${errs.length}`);
 await b.close();process.exit(errs.length?1:0)})().catch(e=>{console.error(e.message);process.exit(1)});
