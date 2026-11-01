/* ============ שלב 2 — PD ← ZM-C וזמן אספקה חסר, בשני הכלים ============
   החלטות המעתד (שלב 2):
   · PD → ZM-C כשהצריכה נטו ב-12 החודשים הקלנדריים האחרונים (11 מלאים +
     החודש הנוכחי) ≥ 3 יח׳. נטו: החזרות מקטינות. 3 יח׳ בחודש אחד מספיקות.
     04/14 מוחרגים לפי קוד וטקסט. עמודה חודשית חסרה — אין המלצה.
   · אין הורדה ל-PD מאף פרופיל. ZM-A/B/C יורדים לפי H, והנמוך ביותר ZM-C,
     עם אותה המתנה ואותן הגנות.
   · «זמן אספקה חסר» = השדה המקורי «אספ.מתוכנ.» ריק או 0 — גם כשזמן עיבוד
     הקבלה > 0. אין אז המלצת SS/ROP: לא מספר, לא 0, ולא ערך SAP בשדה המלצה.
   · קובץ ה-SAP: אין שורה לפריט בלי זמן אספקה; «זמן אספקה מתוכנן» = השדה
     המקורי (בלי עיבוד הקבלה). ssMin||0 לא שונה (ממתין לאימות).
   הבדיקה בונה דוח קטן משלה — בפיקסצ׳ר אין זמן אספקה חסר ואין החזרות —
   ובודקת את הכלי המלא ואת ?topics על אותו קובץ, שדה בשדה. */
const {chromium}=require('playwright');
const XLSX=require('xlsx'),fs=require('fs'),path=require('path'),os=require('os');
const SD=__dirname,ROOT=path.join(SD,'..');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
const MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const HDR=['מק"ט מוביל','תיאור חומר','תיאור חומר2','שם ספק','סטטוס חומר','תיאור','סוג MRP','ABC','רמת שרות',
 'מלאי בטחון','נק.הז.מחדש','אספ.מתוכנ.','זמ.עב.קבלת','מל.בט.מינ.','מחיר FOB','מטבע FOB','סה"כ מלאי','מלאי פנוי',
 'מלאי מרלוג','הז. רכש','בהעברה','אספקות פת.','כמות בהז.פ','סוג חומר','תיא.קבוצ.חומרים','קב.חו.חיצו','טקסט ארוך',
 'תב.אח.הש.','ת.היר.1 מח','ת.היר.2 מח','היררכייה1','היררכייה1','היררכייה2','היררכייה2','היררכייה3','היררכייה3',
 'צר.השנה','צר.שנה-1','צר.שנה-2','צר.החודש',...MON,"תאר' מכירה","תאר' כניסה"];
const D=n=>{const d=new Date(Date.now()-n*864e5);return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`};
/* שורה לפי שם עמודה. lt: undefined → 30; '' → ריק; gr — זמן עיבוד קבלה (ברירת מחדל 2) */
function mk(pn,o){const m=(o.m||[]).concat(Array(11).fill(0)).slice(0,11),y=m.reduce((a,b)=>a+Math.max(0,b),0);
 const v={'מק"ט מוביל':pn,'תיאור חומר':'פריט '+pn,'תיאור חומר2':'T '+pn,'שם ספק':'ספק','סטטוס חומר':o.st||'01','תיאור':o.stx||'פעיל',
  'סוג MRP':o.mrp||'ZM','ABC':o.abc||'C','רמת שרות':o.srv!=null?o.srv:80,'מלאי בטחון':o.ss||0,'נק.הז.מחדש':o.rop||0,
  'אספ.מתוכנ.':o.lt===undefined?30:o.lt,'זמ.עב.קבלת':o.gr===undefined?2:o.gr,'מל.בט.מינ.':o.ssMin||0,'מחיר FOB':o.price||50,'מטבע FOB':'USD',
  'סה"כ מלאי':o.stock!=null?o.stock:(o.free||0),'מלאי פנוי':o.free||0,'מלאי מרלוג':0,'הז. רכש':o.po||0,'בהעברה':0,'אספקות פת.':0,
  'כמות בהז.פ':o.cust||0,'סוג חומר':'Z004','תיא.קבוצ.חומרים':'מנוע','קב.חו.חיצו':'ZT','טקסט ארוך':'מתכנן','תב.אח.הש.':'0',
  'ת.היר.1 מח':o.mkt1||'שיווק','ת.היר.2 מח':'מערכת','צר.השנה':o.y0!=null?o.y0:y,'צר.שנה-1':o.y1!=null?o.y1:y,'צר.שנה-2':o.y2!=null?o.y2:y,
  'צר.החודש':o.cur||0,"תאר' מכירה":D(o.saleAgo||20),"תאר' כניסה":D(o.entAgo||200)};
 MON.forEach((h,i)=>v[h]=m[i]);
 return HDR.map((h,i)=>h==='היררכייה1'?(i===30?'100':'ZT'):h==='היררכייה2'?(i===32?'200':'דגם'):h==='היררכייה3'?(i===34?'300':'מערכת'):v[h])}
const ROWS=[
 /* צריכה נטו 0–4, ושלוש יחידות בעסקה אחת */
 mk('PD-N0',{mrp:'PD',m:[]}),mk('PD-N1',{mrp:'PD',m:[1]}),mk('PD-N2',{mrp:'PD',m:[1,1]}),
 mk('PD-N3',{mrp:'PD',m:[1,1,1]}),mk('PD-N4',{mrp:'PD',m:[2,0,2]}),mk('PD-ONE3',{mrp:'PD',m:[0,0,0,0,3]}),
 /* החזרות: 5−2=3 ↑, 5−3=2 ✗ */
 mk('PD-RET3',{mrp:'PD',m:[5,-2]}),mk('PD-RET2',{mrp:'PD',m:[5,0,-3]}),
 /* החודש החלקי נספר: 1+1+1(חלקי)=3 ↑ · 3+(−1)=2 ✗ */
 mk('PD-CUR',{mrp:'PD',m:[1,1],cur:1}),mk('PD-CURNEG',{mrp:'PD',m:[3],cur:-1}),
 /* מחוץ לחלון: צריכה שנתית גבוהה בלי צריכה ב-12 החודשים — לא נספרת */
 mk('PD-OLD',{mrp:'PD',m:[],y0:0,y1:40,y2:40}),
 /* 04/14 — לפי קוד ולפי טקסט */
 mk('PD-S04',{mrp:'PD',m:[5],st:'04'}),mk('PD-S14',{mrp:'PD',m:[5],st:'14'}),
 mk('PD-TXT',{mrp:'PD',m:[5],stx:'לא פעיל'}),mk('PD-TXT2',{mrp:'PD',m:[5],stx:'גמר המלאי'}),
 /* PD עם 3+ — זמן אספקה תקין / ריק / 0 / 0 עם עיבוד קבלה, ו-PD בלי צריכה בלי זמן אספקה */
 mk('PD-VALID',{mrp:'PD',m:[2,2,2,2,2,2],lt:40}),mk('PD-NOLT',{mrp:'PD',m:[4],lt:''}),
 mk('PD-LT0',{mrp:'PD',m:[4],lt:0,gr:0}),mk('PD-LT0GR',{mrp:'PD',m:[4],lt:0,gr:2}),
 mk('PD-ZERO-NOLT',{mrp:'PD',m:[],lt:''}),
 /* כלי עבודה — המלצה עם אזהרה, ולא 0 בסימולציה */
 mk('PD-TOOL',{mrp:'PD',m:[5],mkt1:'כלי עבודה'}),
 /* מנוהל־מלאי בלי זמן אספקה: SS בלבד / ROP בלבד / שניהם / אף אחד; ואחד עם לקוח ממתין */
 mk('ZM-NOLT-SS',{m:[6,6,6,6,6,6,6,6,6,6,6],lt:'',ss:5,rop:0,free:50}),
 mk('ZM-NOLT-ROP',{m:[6,6,6,6,6,6,6,6,6,6,6],lt:0,ss:0,rop:4,free:50}),
 mk('ZM-NOLT-BOTH',{m:[6,6,6,6,6,6,6,6,6,6,6],lt:0,gr:2,ss:3,rop:9,free:50}),
 mk('ZM-NOLT-NONE',{m:[6,6,6,6,6,6,6,6,6,6,6],lt:'',ss:0,rop:0,free:50}),
 mk('ZM-NOLT-CUST',{m:[6,6,6,6,6,6,6,6,6,6,6],lt:'',ss:0,rop:0,free:0,cust:3}),
 /* מלאי איטי בלי זמן אספקה, עם ROP שהיה יוצא «להפסיק להזמין» */
 mk('ZM-SLOW-NOLT',{m:[],y0:1,y1:1,y2:1,lt:'',stock:300,free:300,rop:10,ss:5,ssMin:2,saleAgo:300,entAgo:400}),
 /* ביקורת: מנוהל עם זמן אספקה תקין ופער פרמטר — שורה בקובץ ה-SAP, LT מקורי 30 (לא 32) */
 mk('ZM-VALID',{m:[10,10,10,10,10,10,10,10,10,10,10],lt:30,gr:2,ss:0,rop:1,free:20}),
 /* הורדות: A עם H=1 → C (לא PD) · C עם H=0 → אין · B עם H=3 → C */
 mk('ZM-A-H1',{abc:'A',m:[1],free:20}),mk('ZM-C-H0',{abc:'C',m:[],free:20}),mk('ZM-B-H3',{abc:'B',m:[2,2,2],free:20})];
const tmp=fs.mkdtempSync(path.join(os.tmpdir(),'pdlt-'));
const write=(rows,hdr,name)=>{const ws=XLSX.utils.aoa_to_sheet([['דוח ZMRP — בדיקה'],[],hdr,...rows]);
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'ZMRP');const f=path.join(tmp,name);XLSX.writeFile(wb,f);return f};
const F1=write(ROWS,HDR,'pdlt.xlsx');
/* קובץ שני: עמודת «צר.חודש-11» חסרה — net12 = null, אין המלצה */
const cut=HDR.indexOf('צר.חודש-11');
const F2=write([mk('PD-NOCOL',{mrp:'PD',m:[5]}),mk('ZM-X',{m:[3,3,3]})].map(r=>r.filter((_,i)=>i!==cut)),HDR.filter((_,i)=>i!==cut),'nocol.xlsx');

const UP=['PD-N3','PD-N4','PD-ONE3','PD-RET3','PD-CUR','PD-VALID','PD-NOLT','PD-LT0','PD-LT0GR','PD-TOOL'];
const NOUP=['PD-N0','PD-N1','PD-N2','PD-RET2','PD-CURNEG','PD-OLD','PD-S04','PD-S14','PD-TXT','PD-TXT2','PD-ZERO-NOLT'];
const LTBAD=['PD-NOLT','PD-LT0','PD-LT0GR','PD-ZERO-NOLT','ZM-NOLT-SS','ZM-NOLT-ROP','ZM-NOLT-BOTH','ZM-NOLT-NONE','ZM-NOLT-CUST','ZM-SLOW-NOLT'];
const ZMNOLT=['ZM-NOLT-SS','ZM-NOLT-ROP','ZM-NOLT-BOTH','ZM-NOLT-NONE','ZM-NOLT-CUST','ZM-SLOW-NOLT'];

(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const errs=[];
 const page=async q=>{const ctx=await b.newContext({viewport:{width:1440,height:900},acceptDownloads:true});
  await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  const p=await ctx.newPage();p.on('pageerror',e=>errs.push(e.message.split('\n')[0]));
  await p.goto('file://'+path.join(ROOT,'index.html')+q);return p};
 const load=async(p,f)=>{await p.setInputFiles('#f',f);
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0,null,{timeout:120000});await p.waitForTimeout(400)};

 /* ================= הכלי המלא ================= */
 const p=await page('?nobrief=1');await load(p,F1);
 const F=await p.evaluate(()=>{const by=Object.fromEntries(ALL.map(r=>[r.pn,r]));
  const info=pn=>{const i=profInfo(by[pn]);return i?{state:i.state,cur:i.cur,tgt:i.tgt,why:i.why}:null};
  const pns=Object.keys(by);
  const X=r=>Object.fromEntries(XCOLS.map(([h,f])=>[h,f(r)]));
  const sapAll=[...sapMassRows(QF.all).slice(1),...sapSlowRows(QF.all).slice(1)];
  const snap=paramSnap(QF.all);
  return {
   up:pns.filter(pn=>{const i=profInfo(by[pn]);return i&&i.state==='up'}).sort(),
   pdUp:pns.filter(pn=>pdUpgrade(by[pn])).sort(),
   net:Object.fromEntries(pns.map(pn=>[pn,by[pn].net12])),
   ltBad:pns.filter(pn=>by[pn].ltBad).sort(),
   ltRaw:Object.fromEntries(pns.map(pn=>[pn,ltRawOf(by[pn])])),
   /* השדות החדשים קיימים רק כשהם רלוונטיים (כמו RATE_DECAY) — פריט תקין נשאר כמו שהיה */
   sparse:(x=>({ltBad:'ltBad' in x,ltRaw:'ltRaw' in x,sugSSCalc:'sugSSCalc' in x,net12:'net12' in x,slowEst:'slowEst' in x,plan:ltPlan(x),raw:ltRawOf(x),ltPlanned:x.ltPlanned}))(by['ZM-VALID']),
   sparseBad:(x=>({ltBad:x.ltBad,ltRaw:x.ltRaw,ltPlanned:x.ltPlanned,plan:ltPlan(x)}))(by['PD-LT0GR']),ltD:Object.fromEntries(pns.map(pn=>[pn,by[pn].ltD])),
   info:Object.fromEntries(['ZM-A-H1','ZM-C-H0','ZM-B-H3','PD-TOOL','PD-NOLT'].map(pn=>[pn,info(pn)])),
   anyToPD:pns.some(pn=>{const i=profInfo(by[pn]);return i&&i.tgt==='D'}),
   cat:Object.fromEntries(pns.map(pn=>[pn,by[pn].cat])),
   why:Object.fromEntries(pns.map(pn=>[pn,(by[pn].why||[]).filter(Boolean).map(w=>w[1]).join(' | ')])),
   act:Object.fromEntries(pns.map(pn=>[pn,(by[pn].act||[]).join(' | ')])),
   recNA:pns.filter(pn=>recNA(by[pn])).sort(),
   sap:sapAll.map(r=>({pn:r[0],ss:r[2],ssMin:r[3],rop:r[4],lt:r[7]})),
   sapHdr:SAP_COLS,
   snap:Object.keys(snap),
   x:Object.fromEntries(pns.map(pn=>[pn,X(by[pn])])),
   cap:capAct(by['ZM-SLOW-NOLT']),slowEst:by['ZM-SLOW-NOLT'].slowEst,slowExp:by['ZM-SLOW-NOLT'].expCap,
   capCell:SCOL.capVal.cell(by['ZM-SLOW-NOLT']),
   floorHas:floorAll().map(r=>r.pn).filter(pn=>by[pn].ltBad),
   ropHas:ropAll().map(r=>r.pn).filter(pn=>by[pn].ltBad),
   sim:Object.fromEntries(['PD-VALID','PD-NOLT','PD-TOOL'].map(pn=>[pn,{rop:profSimVal(by[pn],'rop'),ss:profSimVal(by[pn],'ss')}])),
   simCell:SCOL.profSimRop.cell(by['PD-NOLT']),
   ropCell:SCOL.rop.cell(by['ZM-NOLT-BOTH']),ssCell:SCOL.ss.cell(by['ZM-NOLT-BOTH']),
   LT_MSG,LT_NA,LT_EXIST,LT_NODATA,SLOW_EST,PD_LT_MSG}});
 const C=F;
 ok('PD ← ZM-C: בדיוק המועמדים לפי צריכה נטו ≥ 3 (0–4, עסקה אחת, החזרות, חודש חלקי, 04/14 בקוד ובטקסט)',
  JSON.stringify(F.up)===JSON.stringify([...UP].sort())&&JSON.stringify(F.pdUp)===JSON.stringify([...UP].sort()),
  `עלו ${F.up.join(',')}`);
 ok('צריכה נטו: 0/1/2/3/4, החזרות 3 ו-2, חודש חלקי 3 ו-2, צריכה שנתית מחוץ לחלון 0',
  F.net['PD-N0']===0&&F.net['PD-N1']===1&&F.net['PD-N2']===2&&F.net['PD-N3']===3&&F.net['PD-N4']===4&&F.net['PD-ONE3']===3
  &&F.net['PD-RET3']===3&&F.net['PD-RET2']===2&&F.net['PD-CUR']===3&&F.net['PD-CURNEG']===2&&F.net['PD-OLD']===0,JSON.stringify(F.net));
 ok('אף פרופיל אינו מומלץ ל-PD',!F.anyToPD);
 ok('הורדות ZM: A עם H=1 → ZM-C (לא PD) · C עם H=0 → אין המלצה · B עם H=3 → ZM-C',
  F.info['ZM-A-H1']&&F.info['ZM-A-H1'].tgt==='C'&&F.info['ZM-C-H0']===null&&F.info['ZM-B-H3']&&F.info['ZM-B-H3'].tgt==='C',JSON.stringify(F.info));
 ok('כלי עבודה: המלצה ל-ZM-C עם אזהרת בדיקה ידנית',F.info['PD-TOOL']&&F.info['PD-TOOL'].state==='up'&&/לבדיקה ידנית/.test(F.info['PD-TOOL'].why),F.info['PD-TOOL']&&F.info['PD-TOOL'].why);
 ok('זמן אספקה חסר לפי השדה המקורי — ריק, 0, ו-0 עם עיבוד קבלה 2 (ltD=2)',
  JSON.stringify(F.ltBad)===JSON.stringify([...LTBAD].sort())&&F.ltRaw['PD-NOLT']===''&&F.ltRaw['PD-LT0']===0&&F.ltD['PD-LT0GR']===2&&F.ltRaw['ZM-VALID']===30,
  `ltBad: ${F.ltBad.join(',')}`);
 ok('אין המלצת SS/ROP (recNA) לכל מנוהל־מלאי בלי זמן אספקה',JSON.stringify(F.recNA)===JSON.stringify([...ZMNOLT].sort()),F.recNA.join(','));
 ok('הנוסח שאושר: «חסר זמן אספקה — יש לעדכן ב-SAP» על כל פריט בלי זמן אספקה',
  ZMNOLT.concat(['PD-NOLT','PD-LT0','PD-LT0GR']).every(pn=>F.why[pn].includes(F.LT_MSG)),ZMNOLT.filter(pn=>!F.why[pn].includes(F.LT_MSG)).join(','));
 ok('SS/ROP קיימים ב-SAP: רק SS, רק ROP, שניהם → «קיימים…» · אף אחד → «חסרים נתונים לחישוב SS/ROP»',
  ['ZM-NOLT-SS','ZM-NOLT-ROP','ZM-NOLT-BOTH'].every(pn=>F.why[pn].includes(F.LT_EXIST))&&F.why['ZM-NOLT-NONE'].includes(F.LT_NODATA)&&!F.why['ZM-NOLT-NONE'].includes(F.LT_EXIST));
 ok('PD 3+ בלי זמן אספקה: המלצת פרופיל + הודעת המקרה המשולב',['PD-NOLT','PD-LT0','PD-LT0GR'].every(pn=>F.why[pn].includes(F.PD_LT_MSG)));
 ok('אף פעולה אינה מציעה מספר SS/ROP לפריט בלי זמן אספקה',
  ZMNOLT.every(pn=>!/מ-\d+ ל-\d+/.test(F.act[pn])&&!/תקינ/.test(F.act[pn])),ZMNOLT.map(pn=>pn+': '+F.act[pn]).filter(s=>/מ-\d+ ל-\d+|תקינ/.test(s)).join(' ; '));
 /* ---- קובץ ה-SAP, שדה בשדה ---- */
 ok('קובץ SAP: החוזה (8 עמודות) לא השתנה',JSON.stringify(F.sapHdr)===JSON.stringify(['חומר ','אתר','מלאי בטחון','מלאי ביטחון מינימלי','נקודת הזמנה חוזרת','קבוצת רכש','פרופיל MRP','זמן אספקה מתוכנן']));
 ok('קובץ SAP: אין שורה לאף פריט בלי זמן אספקה (גם 0 עם עיבוד קבלה)',!F.sap.some(r=>LTBAD.includes(r.pn)),F.sap.map(r=>r.pn).join(','));
 const v=F.sap.find(r=>r.pn==='ZM-VALID');
 ok('קובץ SAP: «זמן אספקה מתוכנן» = השדה המקורי 30, לא 32 (עם עיבוד הקבלה)',!!v&&v.lt===30,JSON.stringify(v));
 ok('קובץ SAP: כל שדה מספרי הוא מספר — אין טקסט ואין ריק בשדות SS/ROP/LT',F.sap.every(r=>[r.ss,r.rop,r.lt].every(x=>typeof x==='number'&&isFinite(x))),JSON.stringify(F.sap.slice(0,3)));
 ok('קובץ SAP: ssMin נשאר ssMin||0 (לא שונה — ממתין לאימות)',!!v&&v.ssMin===0);
 ok('«יושמו ב-SAP»: אין תמונה לפריט בלי זמן אספקה',!F.snap.some(pn=>LTBAD.includes(pn)),F.snap.join(','));
 /* ---- ייצוא לאקסל ---- */
 const xbad=[];
 for(const pn of ZMNOLT){const x=F.x[pn];
  if(x['ROP מוצע']!==''||x['SS מוצע']!==''||x['SS סטטיסטי']!==''||x['ביקוש בזמן האספקה']!=='')xbad.push(pn+': מספר בהמלצה');
  if(!String(x['סטטוס המלצה']).includes(F.LT_NA))xbad.push(pn+': אין סטטוס');
  if(typeof x['ROP בSAP']!=='number'||typeof x['SS בSAP']!=='number')xbad.push(pn+': ערכי SAP');}
 ok('אקסל: «ROP/SS מוצע», «SS סטטיסטי» ו«ביקוש בזמן אספקה» ריקים; הסבר בשדה «סטטוס המלצה»; ערכי SAP בעמודות שלהם',!xbad.length,xbad.join(' · '));
 ok('אקסל: «זמן אספקה מתוכנן ב-SAP (מקור)» — ריק נשאר ריק, 0 נשאר 0',F.x['PD-NOLT']['זמן אספקה מתוכנן ב-SAP (מקור)']===''&&F.x['PD-LT0']['זמן אספקה מתוכנן ב-SAP (מקור)']===0&&F.x['ZM-VALID']['זמן אספקה מתוכנן ב-SAP (מקור)']===30);
 ok('אקסל: לפריט תקין — ההמלצה מספר, והסטטוס ריק',typeof F.x['ZM-VALID']['ROP מוצע']==='number'&&F.x['ZM-VALID']['סטטוס המלצה']==='');
 ok('אקסל: PD ← ZM-C בלי זמן אספקה — המלצת הפרופיל והאזהרה בסטטוס',/PD ← ZM-C/.test(F.x['PD-NOLT']['סטטוס המלצה'])&&F.x['PD-NOLT']['סטטוס המלצה'].includes(F.LT_NA));
 /* ---- מלאי איטי, כסף תקוע, דלתות ---- */
 ok('שדות שלב 2 רק בפריט שהם חלים עליו: פריט ZM תקין בלי ltBad/ltRaw/sugSSCalc/net12/slowEst, וזמן אספקה מתוכנן = 30 (בלי 2 ימי קבלה)',
  !F.sparse.ltBad&&!F.sparse.ltRaw&&!F.sparse.sugSSCalc&&!F.sparse.net12&&!F.sparse.slowEst&&F.sparse.plan===30&&F.sparse.raw===30&&F.sparse.ltPlanned===30,JSON.stringify(F.sparse));
 ok('LT=0 עם 2 ימי קבלה: ltBad, ltRaw=0, ltPlanned=0 שמור במפורש (ltD=2)',F.sparseBad.ltBad===true&&F.sparseBad.ltRaw===0&&F.sparseBad.ltPlanned===0&&F.sparseBad.plan===0,JSON.stringify(F.sparseBad));
 ok('מלאי איטי בלי זמן אספקה: החישוב נשאר, מסומן «אומדן» — בנימוק, בעמודה ובאקסל',
  F.slowEst===true&&F.slowExp>0&&F.why['ZM-SLOW-NOLT'].includes(F.SLOW_EST)&&/אומדן/.test(F.capCell)&&F.x['ZM-SLOW-NOLT']['סטטוס המלצה'].includes(F.SLOW_EST),
  `cat=${F.cat['ZM-SLOW-NOLT']} exp=${F.slowExp}`);
 ok('כסף תקוע: בלי זמן אספקה → «אין מה לשנות», לא «להפסיק להזמין»',F.cap==='capNone',F.cap);
 ok('«מלאי ביטחון ידני» ודלתות ה-ROP אינן כוללות פריט בלי זמן אספקה',!F.floorHas.length&&!F.ropHas.length,`${F.floorHas} · ${F.ropHas}`);
 ok('טבלה: ROP/SS — ערך SAP, ובמקום ההמלצה «LT חסר» (הנוסח המלא בכותרת)',/LT חסר/.test(F.ropCell)&&F.ropCell.includes(F.LT_NA)&&/LT חסר/.test(F.ssCell));
 /* ---- מסך הפרופיל: הסימולציה ---- */
 ok('מסך הפרופיל: ROP/SS אחרי ZM-C — מספר כשזמן האספקה תקין, «לא ניתן לחשב — LT חסר» כשחסר, ולא 0 לכלי עבודה',
  typeof F.sim['PD-VALID'].rop==='number'&&typeof F.sim['PD-VALID'].ss==='number'&&F.sim['PD-NOLT'].rop===F.LT_NA&&F.sim['PD-TOOL'].rop==='—'&&/LT חסר/.test(F.simCell),
  JSON.stringify(F.sim));
 ok('סימולציית PD ← ZM-C אינה נכנסת לקובץ ה-SAP',!F.sap.some(r=>r.pn.startsWith('PD-')));
 /* ---- כרטיס הפריט ו«מאחורי החישוב» ---- */
 {const d=await p.evaluate(()=>{const r=ALL.find(x=>x.pn==='ZM-NOLT-BOTH');detail(r);const t=document.getElementById('detail').textContent;
   const boxes=[...document.querySelectorAll('#detail .fixbox .new')].map(e=>e.textContent.trim());
   explain(r);const m=document.getElementById('mbody').textContent;document.getElementById('modalwrap').classList.remove('open');
   return {t,boxes,m}});
  ok('כרטיס הכלי המלא: אזהרה, «קיימים…», ערכי SAP «קיים», ובתיבות ההמלצה «לא ניתן לחשב — LT חסר»',
   d.t.includes(C.LT_MSG)&&d.t.includes(C.LT_EXIST)&&/ב-SAP \(קיים\)/.test(d.t)&&d.boxes.length>=2&&d.boxes.slice(0,2).every(x=>x===C.LT_NA),JSON.stringify(d.boxes));
  ok('«מאחורי החישוב» בלי זמן אספקה: הודעה, בלי טבלת מרכיבים',d.m.includes(C.LT_MSG)&&!/⌈Z/.test(d.m))}
 await p.context().close();

 /* ================= ?topics — אותן החלטות ================= */
 const q=await page('?topics');await load(q,F1);await q.waitForSelector('#lite table');
 const L=await q.evaluate(()=>{const T=TOPICS.fix,rows=T.rows(),cat=k=>rows.filter(T.cats.find(c=>c[0]===k)[3]).map(r=>r.pn).sort();
  return {pd:cat('pd'),lt:cat('lt'),both:rows.filter(r=>pdUpgrade(r)&&tpLtMissing(r)).map(r=>r.pn).sort()}});
 ok('?topics: «PD ← ZM-C» = אותם פריטים כמו בכלי המלא',JSON.stringify(L.pd)===JSON.stringify(F.up),L.pd.join(','));
 ok('?topics: «זמן אספקה חסר» = כל ltBad בלי 04/14, כולל PD בלי צריכה',
  JSON.stringify(L.lt)===JSON.stringify([...LTBAD].sort())&&L.lt.includes('PD-ZERO-NOLT'),L.lt.join(','));
 ok('?topics: פריט יכול להיות בשתי הקטגוריות',JSON.stringify(L.both)===JSON.stringify(['PD-LT0','PD-LT0GR','PD-NOLT']),L.both.join(','));
 await q.click('[data-tp-tab="fix"]');await q.click('#lite .tile[data-tp-cat="lt"]');
 const ord=await q.evaluate(()=>TP.shown.map(r=>[r.pn,tpLtPrio(r)]));
 ok('?topics: סדר «זמן אספקה חסר» — לקוח ממתין, אחר כך חוסר, אחר כך השאר',ord.length>0&&ord[0][0]==='ZM-NOLT-CUST'&&ord.every((x,i)=>!i||ord[i-1][1]<=x[1]),JSON.stringify(ord));
 /* CSV של נושא 1, שדה בשדה */
 await q.click('#lite .tile[data-tp-cat=""]');
 const [dl]=await Promise.all([q.waitForEvent('download'),q.click('#lite [data-tp="csv"]')]);
 const raw=fs.readFileSync(await dl.path(),'utf8').replace(/^﻿/,'').split('\r\n');
 const parse=l=>{const o=[];let c='',s=false;for(let i=0;i<l.length;i++){const ch=l[i];if(s){if(ch==='"'&&l[i+1]==='"'){c+='"';i++}else if(ch==='"')s=false;else c+=ch}else if(ch==='"')s=true;else if(ch===','){o.push(c);c=''}else c+=ch}o.push(c);return o};
 const H=parse(raw[0]),row=pn=>{const r=raw.slice(1).map(parse).find(c=>c[H.indexOf('מק״ט')]===pn);return r?Object.fromEntries(H.map((h,i)=>[h,r[i]])):null};
 const cb=[];
 for(const pn of ZMNOLT.filter(pn=>row(pn))){const r=row(pn);
  if(r['נק הזמנה מוצעת']!==''||r['מלאי ביטחון מוצע']!==''||r['שווי שינוי']!=='')cb.push(pn+': מספר');
  if(!r['סטטוס המלצה'].includes(C.LT_NA))cb.push(pn+': סטטוס');
  if(r['נק הזמנה SAP']===''||r['מלאי ביטחון SAP']==='')cb.push(pn+': ערכי SAP')}
 ok('?topics CSV: אין מספר בעמודות ההמלצה, הסבר ב«סטטוס המלצה», ערכי SAP בעמודות שלהם',!cb.length&&ZMNOLT.some(pn=>row(pn)),cb.join(' · '));
 const rv=row('PD-VALID'),rn=row('PD-NOLT');
 ok('?topics CSV: PD ← ZM-C — ערכי הסימולציה כשזמן האספקה תקין, ריק כשחסר',!!rv&&rv['נק הזמנה מוצעת']!==''&&+rv['נק הזמנה מוצעת']===F.sim['PD-VALID'].rop&&!!rn&&rn['נק הזמנה מוצעת']===''&&/PD ← ZM-C/.test(rn['סטטוס המלצה']));
 ok('?topics CSV: זמן אספקה מתוכנן מקורי — ריק ריק, 0 אפס',row('PD-NOLT')['זמן אספקה מתוכנן ב-SAP (מקור)']===''&&row('PD-LT0')['זמן אספקה מתוכנן ב-SAP (מקור)']==='0');
 /* כרטיס הפריט — שתי הבעיות, בנפרד */
 {await q.evaluate(()=>tpdShow(TOPICS.fix.rows().find(r=>r.pn==='PD-NOLT')));
  const d=await q.evaluate(()=>{const el=document.getElementById('tpd');
   return {t:el.textContent,sec:[...el.querySelectorAll('section h3')].map(h=>h.textContent),
    recCells:[...el.querySelectorAll('table')[0].querySelectorAll('tbody tr')].slice(0,2).map(tr=>tr.children[2].textContent.trim())}});
  ok('?topics כרטיס: «המלצת פרופיל» ו«אזהרת זמן אספקה» בשני מקטעים נפרדים, ובשדות ההמלצה «לא ניתן לחשב — LT חסר»',
   d.sec.some(h=>/המלצת פרופיל/.test(h))&&d.sec.some(h=>/אזהרת זמן אספקה/.test(h))&&d.t.includes(C.PD_LT_MSG)&&d.recCells.every(x=>x===C.LT_NA),JSON.stringify(d.recCells))}
 await q.context().close();

 /* ================= עמודה חודשית חסרה ================= */
 {const r=await page('?nobrief=1');await load(r,F2);
  const m=await r.evaluate(()=>{const x=ALL.find(a=>a.pn==='PD-NOCOL');return {net:x.net12,up:pdUpgrade(x),info:profInfo(x),why:(x.why||[]).map(w=>w[1]).join(' | ')}});
  ok('עמודה חודשית חסרה: net12=null, אין המלצה, ונאמר למה',m.net===null&&!m.up&&m.info===null&&/חסרה עמודת צריכה/.test(m.why),JSON.stringify(m));
  await r.context().close()}

 ok('אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();fs.rmSync(tmp,{recursive:true,force:true});
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
