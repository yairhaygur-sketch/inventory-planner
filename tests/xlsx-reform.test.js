/* ============ XLS_REFORM — קובץ האקסל של הכלי (אושר 23.11) ============
   ההחלטות (docs/BUSINESS-DECISIONS.md):
   · גיליונות: «סיכום» · «התצוגה» · («ללא שינוי פרופיל») · «נתונים עיקריים» · «פירוט מלא» אחרון.
     גיליון התנועות — רק ממסך «תנועות אחרונות» (שם הוא «התצוגה»).
   · «נתונים עיקריים»: העמודות שאושרו, בסדר שאושר, בקבוצות. כל תא = התא של
     XCOLS לאותו מק״ט (שום ערך לא חושב מחדש), «הפעולה הבאה» = הפעולה הראשונה.
   · «פירוט מלא»: מק״ט, תיאור, וכל שאר עמודות XCOLS — שום עמודה לא אבדה.
   · «כסף תקוע (₪)» → «שווי עודף», ולצידו «מטבע».
   · «סיכום»: מסך, סינון, מספר פריטים, תאריכים, שווי עודף לכל מטבע בנפרד.
   · שם הקובץ: שם המסך בעברית ותאריך, «מסונן» כשיש סינון.
   · עיצוב — נבדק בבתים של הקובץ שירד (Worker ובדף): מימין לשמאל, כותרת קפואה
     וסינון אוטומטי, מפריד אלפים; מק״ט בלי עיצוב.
   · דוח הזירוז — אותו עיצוב. קובצי SAP — ללא שינוי (בלי מימין לשמאל, בלי הקפאה). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path'),os=require('os');
const XLSX=require('xlsx');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML=process.argv[2]||path.join(SD,'..','index.html');
const TMP=fs.mkdtempSync(path.join(os.tmpdir(),'xr-'));
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
const MAIN=['מק״ט','תיאור','דגם','ספק','סיווג','הפעולה הבאה','מלאי פנוי','הזמנות לקוח','רכש פתוח','אספקה קרובה (תאריך)',
 'זמן אספקה (ימים)','צריכה חודשית','חודשי כיסוי','נקודת הזמנה SAP','נקודת הזמנה מוצעת','מלאי ביטחון SAP','מלאי ביטחון מוצע',
 'כמות להזמנה','כמות חסרה','שווי עודף','מטבע'];
/* העמודה החדשה → העמודה ב-XCOLS שהיא מציגה */
const SRC={'סיווג':'סטטוס','אספקה קרובה (תאריך)':'תאריך האספקה הקרוב','צריכה חודשית':'קצב חודשי','נקודת הזמנה SAP':'ROP בSAP',
 'נקודת הזמנה מוצעת':'ROP מוצע','מלאי ביטחון SAP':'SS בSAP','מלאי ביטחון מוצע':'SS מוצע','כמות חסרה':'כמות חסרה (יח׳)','מטבע':'מטבע FOB'};
const GROUPS=[['פריט',0],['מלאי',6],['רכש',8],['פרמטרי תכנון',11],['כמויות',17],['שווי',19]];
/* קובץ שירד: התאים (SheetJS) וה-XML של כל גיליון (CFB) */
function inspect(f){const buf=fs.readFileSync(f);const wb=XLSX.read(buf,{cellNF:true});
 const z=XLSX.CFB.read(buf,{type:'buffer'});
 const xml=wb.SheetNames.map((_,i)=>{const p=z.FullPaths.find(q=>q.endsWith('/xl/worksheets/sheet'+(i+1)+'.xml'));
  return p?Buffer.from(XLSX.CFB.find(z,p).content).toString('utf8'):''});
 const pane=x=>{const m=x.match(/<pane [^>]*ySplit="(\d+)"[^>]*state="frozen"/);return m?+m[1]:0};
 const af=x=>{const m=x.match(/<autoFilter ref="([^"]+)"/);return m?m[1]:null};
 return {wb,names:wb.SheetNames,rtl:xml.map(x=>/<sheetView [^>]*rightToLeft="1"/.test(x)),pane:xml.map(pane),af:xml.map(af)}}
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900},acceptDownloads:true});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+HTML+'?nobrief=1');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});
 /* ---- 1 · הספר (XLSX.writeFile מוחלף, כמו בשומר העסקי) ---- */
 const cap=fn=>p.evaluate(fn=>{const o=XLSX.writeFile,res=[];XLSX.writeFile=(wb,name)=>res.push({name,
   sheets:wb.SheetNames.map(n=>[n,XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:null}),(wb.Sheets[n]['!merges']||[]).map(m=>[m.s.r,m.s.c,m.e.r,m.e.c])])});
  try{(new Function(fn))()}finally{XLSX.writeFile=o}return res},fn);
 const C=await p.evaluate(()=>{setMode('catalog');for(const k of Object.keys(colFilters))delete colFilters[k];apply();
  const rows=currentRows(),X=exportRows(rows);
  const cur={};for(const r of rows){const v=r.expCap||0;if(!(v>0))continue;const k=String(r.currency||'').trim().toUpperCase()||'לא צוין';cur[k]=(cur[k]||0)+v}
  return {X,pns:rows.map(r=>r.pn),act:rows.map(r=>String((r.act||[])[0]||'')),scr:screenName(),
   cur:Object.fromEntries(Object.entries(cur).map(([k,v])=>[k,Math.round(v)])),n:rows.length}});
 const [E]=await cap('doExport()');
 const S=Object.fromEntries((E?E.sheets:[]).map(s=>[s[0],s]));
 ok('קטלוג: הגיליונות בסדר שאושר, «פירוט מלא» אחרון, בלי גיליון תנועות',
  !!E&&JSON.stringify(E.sheets.map(s=>s[0]))===JSON.stringify(['סיכום','התצוגה','נתונים עיקריים','פירוט מלא']),E&&E.sheets.map(s=>s[0]).join(' | '));
 const XH=C.X[0],xcol=h=>XH.indexOf(h);
 const M=S['נתונים עיקריים']&&S['נתונים עיקריים'][1],F=S['פירוט מלא']&&S['פירוט מלא'][1];
 ok('«נתונים עיקריים»: העמודות שאושרו, בסדר שאושר',!!M&&JSON.stringify(M[1])===JSON.stringify(MAIN),M&&M[1].join(' | '));
 ok('שורת הקבוצות מעל הכותרת, וכל קבוצה ממוזגת על העמודות שלה',!!M&&GROUPS.every(([g,c])=>M[0][c]===g)
   &&M[0].filter(x=>x!=null&&x!=='').length===GROUPS.length
   &&GROUPS.every(([,c],i)=>{const e=(GROUPS[i+1]||[,MAIN.length])[1]-1;return e===c||S['נתונים עיקריים'][2].some(m=>m[0]===0&&m[1]===c&&m[2]===0&&m[3]===e)}),M&&JSON.stringify(M[0]));
 /* TERMS שלב 3 (23.11): גם במסך «שווי עודף», ולכן גם «התצוגה» (שמעתיק את כותרות המסך) — בלי «כסף תקוע» באף גיליון */
 ok('בלי «כסף תקוע» באף גיליון; «שווי עודף» ולצידו «מטבע»',!!E&&!E.sheets.some(s=>s[1].some(r=>r&&r.some(c=>/כסף תקוע/.test(String(c)))))
   &&MAIN.indexOf('מטבע')===MAIN.indexOf('שווי עודף')+1&&xcol('שווי עודף')>=0&&xcol('כסף תקוע (₪)')<0);
 /* כל תא = התא של XCOLS לאותו מק״ט */
 {const bad=[];if(M)for(let k=0;k<C.n;k++){const row=M[k+2]||[];
   if(row[0]!==C.pns[k]){bad.push(`שורה ${k+3}: ${row[0]} ≠ ${C.pns[k]}`);continue}
   MAIN.forEach((h,c)=>{const want=h==='הפעולה הבאה'?C.act[k]:C.X[k+1][xcol(SRC[h]||h)];const got=row[c];
    if(JSON.stringify(got==null?'':got)!==JSON.stringify(want==null?'':want))bad.push(`${C.pns[k]} · ${h}: ${got} ≠ ${want}`)})}
  ok('«נתונים עיקריים»: כל תא = הערך של XCOLS לאותו מק״ט, באותו סדר שורות',!!M&&M.length===C.n+2&&!bad.length,bad.slice(0,3).join(' · ')||`${C.n} פריטים × ${MAIN.length}`)}
 {const used=new Set(MAIN.filter(h=>h!=='הפעולה הבאה').map(h=>SRC[h]||h));
  const want=['מק״ט','תיאור'].concat(XH.filter(h=>!used.has(h)));
  ok('«פירוט מלא»: מק״ט, תיאור, וכל שאר עמודות XCOLS בסדרן',!!F&&JSON.stringify(F[0])===JSON.stringify(want),F&&F[0].length+' / '+want.length);
  const lost=XH.filter(h=>!used.has(h)&&!(F&&F[0].includes(h)));
  ok('שום עמודה של XCOLS לא אבדה — כל אחת בעיקריים או בפירוט',!lost.length&&used.size+want.length-2===XH.length,lost.join(', '));
  const bad=[];if(F)for(let k=0;k<C.n;k++)F[0].forEach((h,c)=>{const v=(F[k+1]||[])[c],w=C.X[k+1][xcol(h)];
   if(JSON.stringify(v==null?'':v)!==JSON.stringify(w==null?'':w))bad.push(`${C.pns[k]} · ${h}`)});
  ok('«פירוט מלא»: כל תא = XCOLS',!!F&&F.length===C.n+1&&!bad.length,bad.slice(0,3).join(' · '))}
 /* סיכום */
 {const T=S['סיכום']&&S['סיכום'][1]||[],v=l=>(T.find(r=>r&&r[0]===l)||[])[1];
  const hi=T.findIndex(r=>r&&r[0]==='מטבע'&&r[2]==='שווי עודף'),cur={};
  if(hi>=0)for(let i=hi+1;i<T.length&&T[i]&&T[i].length&&T[i][0]!=null;i++)if(typeof T[i][2]==='number')cur[T[i][0]]=T[i][2];
  ok('סיכום: המסך, ללא סינון, ומספר הפריטים',v('מסך')===C.scr&&/^ללא סינון/.test(v('סינון')||'')&&v('פריטים')===C.n,`${v('מסך')} · ${v('פריטים')}`);
  ok('סיכום: שווי עודף לכל מטבע בנפרד = סכום הפריטים באותו מטבע, בלי שורת סך הכול',
   Object.keys(C.cur).length>1&&JSON.stringify(cur)===JSON.stringify(Object.fromEntries(Object.entries(C.cur).sort((a,b)=>b[1]-a[1])))
   &&!T.some(r=>r&&/סה״כ|סך הכול/.test(String(r[0]))),JSON.stringify(cur));
  ok('סיכום: תאריך הדוח ותאריך הייצוא',!!v('תאריך הדוח')&&!!v('תאריך הייצוא'))}
 ok('שם הקובץ: שם המסך בעברית ותאריך',!!E&&C.scr==='קטלוג פריטים'&&/^קטלוג פריטים \d{4}-\d{2}-\d{2}\.xlsx$/.test(E.name),E&&E.name);
 /* סינון → «מסונן»; פרמטרי תכנון → שם עם סוגריים, בלי תו אסור */
 const F2=await p.evaluate(()=>{const k=FDEF[0][0];SEL[k]=new Set([ALL[0][k]]);apply();return currentRows().length});
 const [EF]=await cap('doExport()');
 const fs2=EF?Object.fromEntries(EF.sheets.map(s=>[s[0],s[1]])):{};
 ok('מסונן: «מסונן» בשם הקובץ ובסיכום, והגיליונות = הסינון',!!EF&&/^קטלוג פריטים · מסונן \d{4}-\d{2}-\d{2}\.xlsx$/.test(EF.name)
   &&/^מסונן/.test(((fs2['סיכום']||[]).find(r=>r&&r[0]==='סינון')||[])[1]||'')&&(fs2['פירוט מלא']||[]).length-1===F2&&(fs2['נתונים עיקריים']||[]).length-2===F2,EF&&EF.name);
 await p.evaluate(()=>{for(const [k] of FDEF)SEL[k].clear();apply()});
 const [EM]=await cap("setMode('month');doExport()");
 ok('פרמטרי תכנון: שם המסך המלא בשם הקובץ, בלי תווים אסורים',!!EM&&/^פרמטרי תכנון \(MRP\) \d{4}-\d{2}-\d{2}\.xlsx$/.test(EM.name)&&!/[\\/:*?"<>|]/.test(EM.name),EM&&EM.name);
 ok('פרמטרי תכנון: אותם גיליונות, בלי גיליון תנועות',!!EM&&EM.sheets.map(s=>s[0]).join('|')==='סיכום|התצוגה|נתונים עיקריים|פירוט מלא');
 /* תנועות אחרונות: «התצוגה» = טבלת התנועות, והסיכום סופר אותן */
 const MV=await p.evaluate(()=>{setMode('moves');return movRows().length});
 const [EV]=await cap('doExport()');
 const sv=EV?Object.fromEntries(EV.sheets.map(s=>[s[0],s[1]])):{};
 ok('תנועות אחרונות: «התצוגה» = טבלת התנועות, והסיכום סופר אותן',!!EV&&MV>0&&(sv['התצוגה']||[]).length-1===MV
   &&((sv['סיכום']||[]).find(r=>r&&r[0]==='תנועות')||[])[1]===MV&&/^תנועות אחרונות /.test(EV.name),EV&&`${(sv['התצוגה']||[]).length-1} / ${MV}`);
 await p.evaluate(()=>setMode('catalog'));
 /* ---- 2 · הקובץ שירד — הבתים (Worker, ואז בדף) ---- */
 const dl=async(name,fn)=>{const w=p.waitForEvent('download',{timeout:120000});await p.evaluate(fn);const d=await w;
  await p.waitForFunction(()=>typeof XJOB==='undefined'||!XJOB,null,{timeout:120000});const f=path.join(TMP,name+'.xlsx');await d.saveAs(f);return inspect(f)};
 const chk=(tag,I)=>{const ix=n=>I.names.indexOf(n);
  ok(`${tag}: כל הגיליונות מימין לשמאל`,I.rtl.length===I.names.length&&I.rtl.every(Boolean),I.rtl.join(','));
  ok(`${tag}: כותרת קפואה — 2 שורות בעיקריים, 1 בתצוגה ובפירוט, בלי הקפאה בסיכום`,
   I.pane[ix('נתונים עיקריים')]===2&&I.pane[ix('התצוגה')]===1&&I.pane[ix('פירוט מלא')]===1&&I.pane[ix('סיכום')]===0,I.names.map((n,i)=>n+':'+I.pane[i]).join(' '));
  const ref=n=>{const ws=I.wb.Sheets[n];return ws&&ws['!ref']?XLSX.utils.decode_range(ws['!ref']):null};
  const afOk=(n,h)=>{const r=ref(n),a=I.af[ix(n)];return !!r&&a===XLSX.utils.encode_range({s:{r:h,c:0},e:{r:r.e.r,c:r.e.c}})};
  ok(`${tag}: סינון אוטומטי על שורת הכותרת, עד השורה האחרונה`,afOk('נתונים עיקריים',1)&&afOk('פירוט מלא',0)&&afOk('התצוגה',0)&&!I.af[ix('סיכום')],I.af.join(' | '));
  const ws=I.wb.Sheets['נתונים עיקריים'],r=ref('נתונים עיקריים');const z={int:0,dec:0,pn:0,pnBad:0,money:0,moneyBad:0};
  for(let R=2;R<=r.e.r;R++)for(let c=0;c<=r.e.c;c++){const x=ws[XLSX.utils.encode_cell({r:R,c})];if(!x)continue;
   if(c===0){z.pn++;if(x.z&&x.z!=='General')z.pnBad++;continue}
   if(x.t!=='n')continue;
   if(c===MAIN.indexOf('שווי עודף')){z.money++;if(x.z!=='#,##0')z.moneyBad++;continue}
   if(Number.isInteger(x.v)?x.z==='#,##0':x.z==='#,##0.0#')Number.isInteger(x.v)?z.int++:z.dec++;else z.bad=(z.bad||0)+1}
  ok(`${tag}: מפריד אלפים — שלם #,##0, שבר עד שתי ספרות, שווי תמיד שלם; מק״ט בלי עיצוב`,
   z.int>0&&z.dec>0&&!z.bad&&z.money>0&&!z.moneyBad&&z.pn>0&&!z.pnBad,JSON.stringify(z))};
 chk('Worker',await dl('worker',()=>doExport()));
 chk('בדף (בלי Worker)',await dl('main',()=>{window.__W=window.Worker;window.Worker=undefined;try{doExport()}finally{window.Worker=window.__W}}));
 /* דוח הזירוז — אותו עיצוב, אותו תוכן */
 const X1=await dl('exp',()=>doExpedite());
 const ix=n=>X1.names.indexOf(n);
 ok('דוח זירוז: מימין לשמאל, «לזירוז» עם כותרת קפואה וסינון',X1.rtl.every(Boolean)&&X1.pane[ix('לזירוז')]===1&&/^A1:/.test(X1.af[ix('לזירוז')]||''),X1.names.map((n,i)=>`${n}:${X1.pane[i]}:${X1.af[i]}`).join(' '));
 {const a=XLSX.utils.sheet_to_json(X1.wb.Sheets['לפי ספק'],{header:1,defval:null});const hi=X1.af[ix('לפי ספק')]?XLSX.utils.decode_range(X1.af[ix('לפי ספק')]).s.r:-1;
  ok('דוח זירוז: «לפי ספק» — סינון על כותרת הטבלה שמתחת לאזהרות, בלי הקפאה',hi>2&&X1.pane[ix('לפי ספק')]===0&&a[hi].filter(x=>x!=null&&x!=='').length>3,`כותרת בשורה ${hi+1}`)}
 /* קובצי SAP — ללא שינוי: בלי מימין לשמאל ובלי הקפאה */
 const SP=await dl('sap',()=>doSapMass());
 ok('קובץ SAP: ללא שינוי — גיליון1, בלי מימין לשמאל, בלי הקפאה ובלי סינון',SP.names.join()==='גיליון1'&&!SP.rtl[0]&&!SP.pane[0]&&!SP.af[0]);
 ok('0 שגיאות JS',!errs.length,errs.join(' | '));
 await b.close();fs.rmSync(TMP,{recursive:true,force:true});
 console.log(out.join('\n'));console.log(`\n${out.length-bad}/${out.length}`);process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
