/* ============ PROF_SIM — ROP/SS אחרי מעבר PD → ZM-C ============
   profSimAll() מריץ את build() הקיים על עותק של שורות ההעלאה, עם
   סוג MRP=ZM, רמת שירות 80 ו-ABC=C. profSplit() מחלק לתקין / לבדיקה
   עם כל הסיבות שחלות. אין לו קורא במסך — ולכן הבדיקה היא גם הוכחה
   שהקריאה אינה משנה כלום: לא את ALL, לא את התורים, לא את ROWS ו-DIAG,
   לא את ALL_RAW, לא את המסך ולא את האחסון.

   הדוח כאן סינתטי לגמרי. על הדוח האמיתי (מקומית בלבד, כי יש בו
   מק״טים) נמדד: 34 העלאות — 19 תקינות, 15 לבדיקה, זהה מק״ט במק״ט
   לסימולציה שקדמה לקוד. */
const XLSX=require('xlsx'),{chromium}=require('playwright');
const fs=require('fs'),path=require('path'),os=require('os');
const SD=__dirname,ROOT=path.join(SD,'..');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null?'  ['+x+']':''))};
const J=JSON.stringify;

const MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const HDR=['מק"ט מוביל','תיאור חומר','תיאור חומר2','שם ספק','סטטוס חומר','תיאור','סוג MRP','ABC','רמת שרות',
 'מלאי בטחון','נק.הז.מחדש','אספ.מתוכנ.','זמ.עב.קבלת','מל.בט.מינ.','מחיר FOB','מטבע FOB','סה"כ מלאי','מלאי פנוי',
 'מלאי מרלוג','הז. רכש','בהעברה','אספקות פת.','כמות בהז.פ','סוג חומר','תיא.קבוצ.חומרים','קב.חו.חיצו','טקסט ארוך',
 'תב.אח.הש.','ת.היר.1 מח','ת.היר.2 מח','היררכייה1','היררכייה1','היררכייה2','היררכייה2','היררכייה3','היררכייה3',
 'צר.השנה','צר.שנה-1','צר.שנה-2','צר.החודש',...MON,"תאר' מכירה","תאר' כניסה"];
const D=n=>{const d=new Date(Date.now()-n*864e5);
 return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`};
/* ברירת המחדל: פריט PD (רמת שירות 50) — המועמד להעלאה */
const mk=(pn,o={})=>{const m=o.months||Array(11).fill(0),y=m.reduce((a,b)=>a+b,0);
 return [pn,'בדיקה '+pn,'T '+pn,'ספק',o.status||'01',o.stx||'פעיל',o.mrp||'PD',o.abc||'',
  o.srv!=null?o.srv:50,o.ss||0,o.rop||0,o.lt!=null?o.lt:60,o.ltAdd!=null?o.ltAdd:2,o.ssMin||0,
  o.price!=null?o.price:100,o.cur||'USD',o.free||0,o.free||0,0,o.po||0,0,0,o.cust||0,'Z004','מנוע','ZT','מתכנן',0,
  o.mkt1||'שיווק','מערכת','100','ZT','200','דגם','300','מערכת',
  o.y0!=null?o.y0:y,o.y1!=null?o.y1:y,o.y2!=null?o.y2:y,0,...m,D(10),D(60)]};
function write(file,rows,opt={}){let hdr=HDR.slice(),rr=rows.map(r=>r.slice());
 for(const name of (opt.drop||[])){const i=hdr.indexOf(name);
  if(i>=0){hdr.splice(i,1);rr=rr.map(r=>{const c=r.slice();c.splice(i,1);return c})}}
 const wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['ZMRP'],[],hdr,...rr]),'ZMRP');
 XLSX.writeFile(wb,file)}

/* צריכה סדירה בלי פיק ובלי «חודש אזילה» — כדי שהבדיקה הידנית תוכל
   לחשב קצב ו-σ בלי לשחזר את ניקוי החודשים של analyze */
const REG=[6,7,5,6,8,5,7,6,5,7,6];
const LOW=[1,1,0,0,0,0,0,0,0,0,0];              /* 2 יח׳ ב-11 חודשים, 6 ב-3 שנים — מתחת לרף */
const CASES=[
 ['PF-REG',     {months:REG},                                   'ok',  []],
 ['PF-TOOL',    {months:REG,mkt1:'כלי עבודה'},                  'rev', ['od_group']],
 ['PF-ACC',     {months:REG,mkt1:'אביזרים'},                    'rev', ['od_group']],
 ['PF-S04',     {months:REG,status:'04',stx:'לא פעיל'},          'rev', ['inactive']],
 ['PF-S14',     {months:REG,status:'14',stx:'גמר המלאי'},        'rev', ['inactive']],
 ['PF-ACC04',   {months:REG,mkt1:'אביזרים',status:'04',stx:'לא פעיל'},'rev',['od_group','inactive']],
 ['PF-MARK',    {months:REG},                                   'ok',  []],
 ['PF-LOWHI',   {months:LOW,price:500},                         'rev', ['zero_low']],
 ['PF-LOWCHEAP',{months:LOW,price:5},                           'ok',  []],
 ['PF-NOPRICE', {months:LOW,price:0},                           'rev', ['zero_noprice']],
 ['PF-NOLT',    {months:REG,lt:0,ltAdd:0},                      'rev', ['no_lt']],
 ['PF-SSMIN7',  {months:[3,3,3,4,3,3,3,4,3,3,3],ssMin:7},       'ok',  []],
 ['PF-ROP1',    {months:LOW,price:500,rop:1,ss:0},              'rev', ['ss_zero']],
 ['PF-FLAT',    {months:Array(11).fill(1),price:500},           'rev', ['ss_zero_rate3y']],
 ['PF-SILENT',  {months:[0,0,0,0,0,0,3,4,2,3,1]},               'rev', ['zero_silent']],
 ['PF-DUP',     {months:REG},                                   'rev', ['no_row']]];
const MANAGED=[mk('BG-ZM',{mrp:'ZM',abc:'C',srv:80,months:REG,free:30,rop:20,ss:8}),
 mk('BG-ND',{mrp:'ND',abc:'A',srv:95,months:[2,2,2,2,2,2,2,2,2,2,2],free:400,price:50}),
 /* PD שאינו מועמד: צריכה בחודש אחד בלבד */
 mk('BG-PD1',{months:[2,0,0,0,0,0,0,0,0,0,0]})];
const tmp=fs.mkdtempSync(path.join(os.tmpdir(),'profsim-'));
const XA=path.join(tmp,'a.xlsx'),XB=path.join(tmp,'b.xlsx');
write(XA,[...CASES.map(([pn,o])=>mk(pn,o)),mk('PF-DUP',{months:REG}),...MANAGED]);
/* אותם פריטים בלי עמודת רמת השירות */
write(XB,[...CASES.filter(c=>c[0]!=='PF-DUP').map(([pn,o])=>mk(pn,o)),...MANAGED],{drop:['רמת שרות']});

/* כל מה שהמסך והמנוע מחזיקים, מצולם כמחרוזת אחת */
const SNAP=()=>{
 const pick=it=>{const o={};for(const k in it)if(typeof it[k]!=='function')o[k]=it[k];return o};
 const ids=v=>Array.isArray(v)?v.map(x=>x&&x.pn):v;
 const qo=q=>q?Object.fromEntries(Object.entries(q).map(([k,v])=>[k,ids(v)])):null;
 let h=0;for(const row of ALL_RAW.rows){const s=JSON.stringify(row);for(let i=0;i<s.length;i++)h=(h*31+s.charCodeAt(i))|0}
 const T=s=>(document.querySelector(s)||{}).textContent||'';
 return JSON.stringify({ROWS,DIAG,all:ALL.map(pick),Q:qo(Q),QF:qo(QF),raw:h,hdr:ALL_RAW.hdr,
  ls:Object.keys(localStorage).sort().map(k=>[k,localStorage.getItem(k)]),
  rail:T('#railsub .railsub2[data-m="prof"] .bdg'),pgL:T('#pgL'),pgR:T('#pgR'),
  body:document.getElementById('tbl')?document.getElementById('tbl').innerHTML:'',
  prof:profRows().map(r=>r.pn),act:profRows().filter(profActive).length,mode:typeof MODE!=='undefined'?MODE:null})};

(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(ROOT,'index.html')+'?nobrief=1');
 const load=async f=>{await p.evaluate(()=>{window.__prev=ALL});await p.setInputFiles('#f',f);
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&ALL!==window.__prev&&ALL.length>0&&!LOADJOB,null,{timeout:180000,polling:200})};
 await load(XA);
 await p.evaluate(()=>setMode('prof'));

 /* ---- 1. בידוד: לפני ואחרי הקריאה הראשונה — הכול זהה ---- */
 const iso=await p.evaluate(SNAP=>{const snap=eval('('+SNAP+')');
  const s0=snap(),d0=DIAG,pf0=PFSIM;const S=profSimAll();const s1=snap();
  const sp=ALL.map(profSplit);const s2=snap();
  return {same1:s0===s1,same2:s0===s2,diagId:DIAG===d0,pf0:pf0===null,map:[...S.map.keys()].sort(),srvCol:S.srvCol}},SNAP.toString());
 ok('לפני הקריאה המטמון ריק — שום מסלול טעינה או רינדור לא הפעיל את הסימולציה',iso.pf0);
 ok('profSimAll() לא שינה דבר: ALL, Q, QF, ROWS, DIAG, ALL_RAW, המסך, המונה, localStorage',iso.same1);
 ok('profSplit() על כל הפריטים לא שינה דבר',iso.same2);
 ok('DIAG הוא אותו אובייקט (mapCols יוצר חדש — הוחזר המקורי)',iso.diagId);
 const upPn=CASES.filter(c=>c[0]!=='PF-DUP').map(c=>c[0]).sort();
 ok('הסימולציה רצה רק על פריטי «להעלות» — אף פריט מנוהל, אף PD שאינו מועמד',J(iso.map)===J(upPn),iso.map.join(','));

 /* ---- 2. החלוקה וקודי הסיבה, פריט אחר פריט ---- */
 const sp=await p.evaluate(()=>Object.fromEntries(ALL.map(r=>[r.pn+(r===ALL.find(x=>x.pn===r.pn)?'':'#2'),profSplit(r)])));
 for(const [pn,,grp,why] of CASES){const x=sp[pn];
  ok(`${pn} → ${grp} ${why.join('+')||'(ללא סיבה)'}`,!!x&&x.grp===grp&&J(x.reasons)===J(why)&&x.primary===(why[0]||null),
   x?`${x.grp} ${x.reasons.join('+')} ROP=${x.sim&&x.sim.rop} SS=${x.sim&&x.sim.ss}`:'null')}
 ok('PF-DUP — גם השורה הכפולה השנייה → no_row, בלי לבחור שורה',sp['PF-DUP#2']&&sp['PF-DUP#2'].grp==='rev'&&J(sp['PF-DUP#2'].reasons)===J(['no_row'])&&sp['PF-DUP#2'].sim===null);
 ok('חפיפה: PF-ACC04 נושא את שני הקודים, primary=od_group',J(sp['PF-ACC04'].reasons)===J(['od_group','inactive'])&&sp['PF-ACC04'].primary==='od_group');
 ok('פריט מנוהל ו-PD שאינו מועמד → null (אין להם חלוקה)',sp['BG-ZM']===null&&sp['BG-ND']===null&&sp['BG-PD1']===null);
 for(const pn of ['PF-TOOL','PF-ACC','PF-S04','PF-S14','PF-ACC04'])
  ok(`${pn} — מוחרג, ולכן ROP=0 ו-SS=0`,sp[pn].sim.rop===0&&sp[pn].sim.ss===0);
 ok('PF-LOWCHEAP — רצפת מדרגה A: SS=3, ROP=3',sp['PF-LOWCHEAP'].sim.ss===3&&sp['PF-LOWCHEAP'].sim.rop===3,`${sp['PF-LOWCHEAP'].sim.ss}/${sp['PF-LOWCHEAP'].sim.rop}`);
 ok('PF-SSMIN7 — רצפת SAP נשמרת: SS=7',sp['PF-SSMIN7'].sim.ss===7&&sp['PF-SSMIN7'].sim.ssSrc==='sap',`${sp['PF-SSMIN7'].sim.ss} ${sp['PF-SSMIN7'].sim.ssSrc}`);
 ok('PF-ROP1 — נוכחות מינימלית: ROP=1, SS=0',sp['PF-ROP1'].sim.rop===1&&sp['PF-ROP1'].sim.ss===0);
 ok('PF-FLAT — ROP≥2 ו-SS=0',sp['PF-FLAT'].sim.rop>=2&&sp['PF-FLAT'].sim.ss===0,`${sp['PF-FLAT'].sim.rop}/${sp['PF-FLAT'].sim.ss}`);
 ok('כל הסימולציות על רמת שירות 80 ו-z=0.84',Object.values(sp).filter(x=>x&&x.sim).every(x=>x.sim.srv===80&&x.sim.z===0.84));

 /* ---- 3. אורקל ידני: הנוסחה מחושבת כאן, לא רק משווים את הקוד לעצמו ---- */
 const reg=sp['PF-REG'].sim,ltM=62/30,mu=REG.reduce((a,b)=>a+b,0)/11;
 const sd=Math.sqrt(REG.reduce((a,b)=>a+(b-mu)**2,0)/10),rate=REG.slice(0,6).reduce((a,b)=>a+b,0)/6;
 const ssX=Math.ceil(0.84*sd*Math.sqrt(ltM)),ropX=Math.ceil(rate*ltM+ssX);
 ok('PF-REG — SS = ⌈z·σ·√LT⌉ ו-ROP = ⌈קצב·LT + SS⌉, מחושב ידנית',reg.ss===ssX&&reg.rop===ropX&&ssX>0,`${reg.ss}/${reg.rop} מול ${ssX}/${ropX}`);
 /* ואותו build() על אותה שורה, כשמשנים את שלושת התאים ביד */
 const same=await p.evaluate(()=>{const m=mapCols(ALL_RAW.hdr);const r0=Object.assign({},ROWS),d0=DIAG;
  try{const rows=ALL_RAW.rows.filter(row=>/^PF-/.test(CLEAN(row[m.pn]))&&CLEAN(row[m.pn])!=='PF-DUP').map(row=>{const c=row.slice();c[m.mrp]='ZM';c[m.srv]=80;c[m.abc]='C';return c});
   return build(rows,ALL_RAW.hdr).every(it=>{const x=PFSIM.map.get(it.pn);return x&&x.rop===it.sugROP&&x.ss===it.sugSS})}
  finally{Object.assign(ROWS,r0);DIAG=d0}});
 ok('לכל פריט: ROP/SS זהים ל-build() על השורה כ-ZM/80/C',same);

 /* ---- 4. חזרתיות ומטמון ---- */
 const rep=await p.evaluate(()=>{const a=J=>JSON.stringify(ALL.map(profSplit));const r1=a(),S=PFSIM;
  const r2=a(),r3=a();const hit=profSimAll()===S;PFSIM=null;const r4=a();return {eq:r1===r2&&r2===r3&&r3===r4,hit,fresh:PFSIM!==S}});
 ok('3 קריאות, ושוב אחרי ביטול המטמון — אותה תוצאה בדיוק',rep.eq);
 ok('המטמון נשמר בין קריאות ומחושב מחדש כשמבטלים אותו',rep.hit&&rep.fresh);

 /* ---- 5. סימון ידני «לפי דרישה» — חי, בלי טעינה ---- */
 const mark=await p.evaluate(()=>{const it=ALL.find(r=>r.pn==='PF-MARK');const S=PFSIM;
  setMark('PF-MARK','ondemand');const a=profSplit(ALL.find(r=>r.pn==='PF-MARK'));
  setMark('PF-MARK',null);const b=profSplit(ALL.find(r=>r.pn==='PF-MARK'));
  return {a,b,cache:PFSIM===S}});
 ok('סימון «לפי דרישה» → rev + od_mark',mark.a.grp==='rev'&&J(mark.a.reasons)===J(['od_mark']));
 ok('ביטול הסימון → חוזר ל-ok, בלי חישוב מחדש',mark.b.grp==='ok'&&!mark.b.reasons.length&&mark.cache);

 /* ---- 6. build() נכשל באמצע — ROWS ו-DIAG חוזרים, אין מטמון חלקי ---- */
 const fail=await p.evaluate(SNAP=>{const snap=eval('('+SNAP+')');
  PFSIM=null;const s0=snap(),r0=JSON.stringify(ROWS),d0=DIAG,a0=analyze;let n=0,threw=false;
  analyze=function(m){if(++n===3)throw new Error('כשל מכוון');return a0(m)};
  try{profSimAll()}catch(e){threw=e.message==='כשל מכוון'}finally{analyze=a0}
  const mid={rows:JSON.stringify(ROWS)===r0,diag:DIAG===d0,pf:PFSIM===null,same:snap()===s0,n};
  const sp=ALL.map(profSplit).filter(Boolean);
  return {threw,...mid,after:sp.length===17&&profSplit(ALL.find(r=>r.pn==='PF-REG')).grp==='ok'}},SNAP.toString());
 ok('כשל בתוך build() — השגיאה עולה למי שקרא',fail.threw,`נכשל בקריאה ${fail.n}`);
 ok('אחרי הכשל: ROWS זהה, DIAG אותו אובייקט',fail.rows&&fail.diag);
 ok('אחרי הכשל: אין מטמון חלקי, וכל שאר המצב זהה',fail.pf&&fail.same);
 ok('ואחרי הכשל הקריאה הבאה עובדת',fail.after);

 /* ---- 7. שורת מקור חסרה ---- */
 const miss=await p.evaluate(()=>{const R=ALL_RAW,m=mapCols(R.hdr);
  const r0=Object.assign({},ROWS),d0=DIAG;
  ALL_RAW={rows:R.rows.filter(row=>CLEAN(row[m.pn])!=='PF-REG'),hdr:R.hdr};
  try{return profSplit(ALL.find(r=>r.pn==='PF-REG'))}finally{ALL_RAW=R;PFSIM=null;Object.assign(ROWS,r0);DIAG=d0}});
 ok('שורת מקור חסרה → rev + no_row',miss&&miss.grp==='rev'&&J(miss.reasons)===J(['no_row'])&&miss.sim===null);

 /* ---- 8. שינוי חודש בסיס — ALL חדש, המטמון מתבטל ---- */
 const anc=await p.evaluate(()=>{profSimAll();const S=PFSIM,a0=ANCHOR;
  const d=new Date(ANCHOR);d.setMonth(d.getMonth()-1);
  document.getElementById('rd').onchange({target:{value:`${d.getFullYear()}-${d.getMonth()+1}`}});
  const r=profSplit(ALL.find(x=>x.pn==='PF-REG'));const fresh=PFSIM!==S&&PFSIM.all===ALL;
  ANCHOR=a0;rebuildLabels();ALL=build(ALL_RAW.rows,ALL_RAW.hdr);apply();
  return {fresh,grp:r&&r.grp}});
 ok('שינוי חודש בסיס → המטמון מחושב מחדש על ALL החדש',anc.fresh&&anc.grp==='ok');

 /* ---- 9. בלי עמודת רמת שירות — לא מחשבים 80% בהשערה ---- */
 await load(XB);
 const ns=await p.evaluate(()=>{const S=profSimAll();
  return {col:S.srvCol,map:S.map.size,sp:Object.fromEntries(ALL.map(r=>[r.pn,profSplit(r)]).filter(x=>x[1]))}});
 ok('אין עמודת רמת שירות → לא רצה סימולציה בכלל',ns.col===false&&ns.map===0);
 ok('כל פריטי «להעלות» → rev + no_srv, וסיבות הפריט עצמו נשמרות',
  Object.keys(ns.sp).length===upPn.length&&Object.values(ns.sp).every(x=>x.grp==='rev'&&x.reasons.includes('no_srv')&&x.sim===null)
  &&J(ns.sp['PF-ACC04'].reasons)===J(['od_group','inactive','no_srv'])&&J(ns.sp['PF-REG'].reasons)===J(['no_srv']));

 ok('אפס שגיאות JS',errs.length===0,errs.join(' | '));
 await b.close();fs.rmSync(tmp,{recursive:true,force:true});
 console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length} עברו`);process.exit(bad?1:0)
})().catch(e=>{console.error(e);process.exit(1)});
