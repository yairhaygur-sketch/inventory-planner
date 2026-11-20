/* ============ COL_ORDER + FOB_COL — סידור עמודות ומחיר FOB בכלי המלא (החלטת המעתד 20.11) ============
   כמו בתצוגת 4 הנושאים: חצים ↑↓ בחלון «עמודות» וגרירת כותרת בטבלה. הסדר נשמר לכל תצוגה
   בנפרד (planner_cols_v1). המק״ט — ונקודת התנועה ותיבת הבחירה — נשארים ראשונים ואינם זזים.
   רוחב שנקבע ביד זז יחד עם העמודה, וגרירת ידית הרוחב אינה גרירת עמודה.
   «מחיר FOB»: עמודה לבחירה, במסך בלבד — לא בברירת המחדל ולא באקסל (noX, החלטה 1.11). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 const wait=ms=>p.waitForTimeout(ms);
 const load=async()=>{await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');await p.setInputFiles('#f',XL);
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length>0&&!LOADJOB,null,{timeout:180000});
  await wait(600);await p.evaluate(()=>setMode('today'));await wait(500)};
 await load();
 await p.evaluate(()=>{try{localStorage.removeItem('planner_cols_v1');localStorage.removeItem('planner_colw_v1')}catch(_){}
  WSORT={col:null,dir:'desc'};render()});await wait(400);
 const heads=(s='#tbl')=>p.evaluate(s=>[...document.querySelectorAll(s+' thead th[data-ck]')].map(t=>t.dataset.ck),s);
 const wOf=(s='#tbl')=>p.evaluate(s=>Object.fromEntries([...document.querySelectorAll(s+' thead th[data-ck]')]
  .map(t=>[t.dataset.ck,Math.round(t.getBoundingClientRect().width)])),s);
 const isOpen=()=>p.evaluate(()=>document.getElementById('colspick').classList.contains('open'));
 const openPick=async()=>{if(!await isOpen()){await p.click(await p.evaluate(()=>'#'+colsBtnEl().id));await wait(300)}};
 const closePick=async()=>{await p.evaluate(()=>colsOpen(false));await wait(200)};
 const focusCmv=()=>p.evaluate(()=>document.activeElement&&(document.activeElement.dataset.cmv||document.activeElement.dataset.ck||document.activeElement.tagName));

 /* 1 · ברירת המחדל לא זזה; «מחיר FOB» אינו בברירת המחדל של אף תצוגה */
 const D=await p.evaluate(()=>({v:colViewKey(),exp:SCOL_DEF.dec.filter(colUsable),
  got:[...document.querySelectorAll('#tbl thead th[data-ck]')].map(t=>t.dataset.ck),
  noFob:Object.values(SCOL_DEF).every(d=>!d.includes('fob')),noX:!!(SCOL.fob&&SCOL.fob.noX)}));
 ok('ברירת המחדל של «לפי החלטה» לא השתנתה',D.v==='dec'&&JSON.stringify(D.exp)===JSON.stringify(D.got),D.got.join(','));
 ok('«מחיר FOB» אינו בברירת המחדל של אף תצוגה, ומסומן «במסך בלבד» (noX)',D.noFob&&D.noX);
 const G=await p.evaluate(()=>[...document.querySelectorAll('#tbl thead th[data-ck]')].map(t=>t.dataset.ck+':'+(t.getAttribute('draggable')||'-')));
 ok('כל כותרת נגררת — פרט ל«מק״ט»',G[0]==='pn:-'&&G.slice(1).every(x=>/:true$/.test(x)),G.slice(0,4).join(' '));

 /* 2 · חלון «עמודות»: המוצגות קודם ובסדר שלהן, עם חצים; «מק״ט» בלי חצים */
 await openPick();
 const P=await p.evaluate(()=>{const vis=visCols(),rows=[...document.querySelectorAll('#colspick .cprow')];
  const R=rows.map(r=>{const c=r.querySelector('input[data-ck]'),up=r.querySelector('[data-cmv$=":-1"]'),dn=r.querySelector('[data-cmv$=":1"]');
   return {k:c.dataset.ck,on:c.checked,up:up?(up.disabled?'x':'o'):'',dn:dn?(dn.disabled?'x':'o'):'',lbl:dn?dn.getAttribute('aria-label'):''}});
  const by=k=>R.find(r=>r.k===k)||{};
  return {vis,order:R.filter(r=>r.on).map(r=>r.k),firstOff:R.findIndex(r=>!r.on),
   pn:by('pn'),second:by(vis[1]),mid:by(vis[3]),last:by(vis[vis.length-1]),
   offNoArrows:R.filter(r=>!r.on).every(r=>!r.up&&!r.dn),fob:!!by('fob').k,
   hint:(document.querySelector('#colspick .cpn')||{}).textContent||''}});
 ok('בחלון: העמודות המוצגות קודם, בסדר שעל המסך',JSON.stringify(P.order)===JSON.stringify(P.vis)&&P.firstOff===P.vis.length,P.order.slice(0,5).join(','));
 ok('«מק״ט» בלי חצים; ↑ כבוי בעמודה שאחריו; ↓ כבוי באחרונה',!P.pn.up&&!P.pn.dn&&P.second.up==='x'&&P.second.dn==='o'&&P.last.dn==='x'&&P.last.up==='o',
  `${P.vis[1]} ${P.second.up}${P.second.dn} · ${P.vis[P.vis.length-1]} ${P.last.up}${P.last.dn}`);
 ok('חצים רק לעמודות מוצגות, ולכל חץ שם נגיש',P.offNoArrows&&P.mid.up==='o'&&P.mid.dn==='o'&&/מאוחר יותר/.test(P.mid.lbl),P.mid.lbl);
 ok('«מחיר FOB» מוצע בחלון, וההסבר מזכיר חצים וגרירה',P.fob&&/החצים/.test(P.hint)&&/לגרור/.test(P.hint)&&/החלטה/.test(P.hint),P.hint.trim().slice(0,60));

 /* 3 · ↓ ו-↑: צעד אחד בין המוצגות, המיקוד נשאר על החץ, ובקצה עובר לחץ השני */
 const h0=await heads(),k1=h0[1];
 await p.click(`#colspick [data-cmv="${k1}:1"]`);await wait(400);
 let H=await heads();
 ok('↓ מחליף עמודה עם זו שאחריה',H[2]===k1&&H[1]===h0[2]&&H[0]==='pn',H.slice(0,4).join(','));
 ok('המיקוד נשאר על אותו חץ אחרי הבנייה מחדש',await focusCmv()===`${k1}:1`,await focusCmv());
 await p.click(`#colspick [data-cmv="${k1}:-1"]`);await wait(400);
 H=await heads();
 ok('↑ מחזיר אותה',JSON.stringify(H)===JSON.stringify(h0),H.slice(0,4).join(','));
 ok('בקצה (↑ כבוי עכשיו) המיקוד עובר ל-↓ של אותה עמודה',await focusCmv()===`${k1}:1`,await focusCmv());
 /* מקלדת: Enter על החץ */
 const k3=h0[3];
 await p.focus(`#colspick [data-cmv="${k3}:1"]`);await p.keyboard.press('Enter');await wait(400);
 H=await heads();
 ok('מקלדת · Enter על ↓ מזיז, והמיקוד נשאר',H[4]===k3&&await focusCmv()===`${k3}:1`,`${H.slice(0,6).join(',')} · ${await focusCmv()}`);
 await p.keyboard.press('Shift+Tab');await p.keyboard.press('Enter');await wait(400);
 H=await heads();
 ok('מקלדת · Shift+Tab ל-↑ ו-Enter מחזיר',JSON.stringify(H)===JSON.stringify(h0),H.slice(0,6).join(','));
 const S1=await p.evaluate(()=>({fix:colStep('fix','model',-1),pin:colMove('dec','pn','desc'),onPin:colMove('dec','desc','pn'),
  sel:colMove('fix','model','sel'),sav:colSaved().fix}));
 ok('עמודה מוצמדת אינה זזה ואי אפשר לדחוף לפניה (גם תיבת הבחירה ב«תיקוני פרמטרים»)',!S1.fix&&!S1.pin&&!S1.onPin&&!S1.sel&&S1.sav===undefined,JSON.stringify(S1));

 /* 4 · «מחיר FOB»: סימון — נכנס אחרי «שווי»; תא = מחיר במטבע הפריט, או מקף */
 await p.click('#colspick input[data-ck="fob"]');await wait(450);
 H=await heads();
 ok('«מחיר FOB» נכנס מיד אחרי «שווי» (ולא מיד אחרי המק״ט)',H.indexOf('fob')===H.indexOf('val')+1&&H[1]===h0[1],H.join(','));
 const F=await p.evaluate(()=>{const trs=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
  const fmt=v=>Number(v).toLocaleString('he-IL',{maximumFractionDigits:Math.abs(v)<100?2:0});
  let n=0,miss=0,good=0;const badL=[];
  for(const tr of trs){const pn=(tr.querySelector('.obj')||{}).textContent;const r=pn&&QF.all.find(x=>String(x.pn)===pn.trim());
   const td=tr.querySelector('td[data-ck="fob"]');if(!r||!td)continue;n++;const t=td.textContent.trim();
   if(r.priceMissing){miss++;if(t==='—')good++;else badL.push(pn+':'+t)}
   else{const sym=CURSYM[String(r.currency||'').trim().toUpperCase()];
    if(t.includes(fmt(r.price))&&(!sym||t.startsWith(sym)))good++;else badL.push(pn+':'+t+' ≠ '+fmt(r.price))}}
  const th=document.querySelector('#tbl thead th[data-ck="fob"]');
  return {n,miss,good,bad:badL.slice(0,3),ttl:th?th.querySelector('.thc').title:''}});
 ok('תא FOB: מחיר ליחידה במטבע הפריט; פריט בלי מחיר — מקף',F.n>20&&F.good===F.n,`${F.good}/${F.n} · בלי מחיר ${F.miss} ${F.bad.join(' | ')}`);
 ok('בדוח הדמו יש גם פריט בלי מחיר (המקף נבדק באמת)',F.miss>0,String(F.miss));
 ok('הכותרת מסבירה: במסך בלבד, אינו יוצא לאקסל',/במסך בלבד/.test(F.ttl)&&/לאקסל/.test(F.ttl),F.ttl.slice(0,80));
 const X=await p.evaluate(()=>{const v=viewExportRows();const exp=visCols().filter(k=>SCOL[k]&&(SCOL[k].xval||SCOL[k].val)&&k!=='spark'&&k!=='fob');
  return {keys:v.keys,exp,head:v.aoa[0]}});
 ok('גיליון «התצוגה» בייצוא — בלי «מחיר FOB», ושאר העמודות כמו על המסך',!X.keys.includes('fob')&&!X.head.includes('מחיר FOB')&&JSON.stringify(X.keys)===JSON.stringify(X.exp),X.keys.join(','));
 await closePick();
 await p.evaluate(()=>{document.querySelector('#tbl thead th[data-ck="fob"] .thc').click()});await wait(400);
 const FS=await p.evaluate(()=>({ws:{...WSORT},miss:SCOL.fob.val({priceMissing:true,price:5}),v:SCOL.fob.val({price:12.5})}));
 ok('לחיצה על «מחיר FOB» ממיינת לפי המחיר (בלי מחיר — אחרון)',FS.ws.col==='fob'&&FS.ws.dir==='desc'&&FS.miss===-1&&FS.v===12.5,JSON.stringify(FS));
 await p.evaluate(()=>{WSORT={col:null,dir:'desc'};render()});await wait(300);

 /* 5 · גרירת כותרת */
 const g0=await heads(),from=g0[g0.length-1],to=g0[3];
 await p.dragAndDrop(`#tbl thead th[data-ck="${from}"]`,`#tbl thead th[data-ck="${to}"]`);await wait(450);
 H=await heads();
 ok('גרירת הכותרת האחרונה אל המקום הרביעי',H[3]===from&&H[4]===to&&H[0]==='pn'&&H.length===g0.length,`${from} → ${H.slice(0,5).join(',')}`);
 const g1=await heads();
 await p.dragAndDrop('#tbl thead th[data-ck="pn"]',`#tbl thead th[data-ck="${g1[5]}"]`);await wait(350);
 await p.dragAndDrop(`#tbl thead th[data-ck="${g1[5]}"]`,'#tbl thead th[data-ck="pn"]');await wait(350);
 H=await heads();
 ok('«מק״ט» אינו נגרר, ואי אפשר לשחרר עליו',JSON.stringify(H)===JSON.stringify(g1),H.slice(0,4).join(','));
 ok('אחרי גרירה לא נשאר סימון גרירה על הכותרות',await p.evaluate(()=>!document.querySelector('#tbl thead th.dragging,#tbl thead th.dropto')));
 await p.evaluate(()=>{const th=document.querySelector('#tbl thead th.wsrt[data-ck="cust"]');th.querySelector('.thc').click()});await wait(350);
 ok('לחיצה על כותרת עדיין ממיינת',await p.evaluate(()=>WSORT.col==='cust'));
 await p.evaluate(()=>{WSORT={col:null,dir:'desc'};render()});await wait(300);
 const X2=await p.evaluate(()=>viewExportRows().keys);
 ok('גיליון «התצוגה» הולך אחרי הסדר שעל המסך (כמו בהסתרה/הוספה היום)',!X2.includes('fob')&&X2.length>5&&JSON.stringify(X2)===JSON.stringify(H.filter(k=>X2.includes(k))),X2.slice(0,6).join(','));

 /* 6 · רוחב: ידית הרוחב אינה גוררת עמודה, ורוחב שנקבע ביד זז עם העמודה */
 const r0=await heads(),rk=r0[2],w0=await wOf();
 const hb=await p.evaluate(k=>{const h=document.querySelector(`#tbl thead th[data-ck="${k}"] .colrz`).getBoundingClientRect();
  return {x:h.left+h.width/2,y:h.top+h.height/2}},rk);
 await p.mouse.move(hb.x,hb.y);await p.mouse.down();await p.mouse.move(hb.x-80,hb.y,{steps:8});await p.mouse.up();await wait(400);
 const w1=await wOf();H=await heads();
 ok('גרירת ידית הרוחב משנה רוחב ולא מזיזה עמודה',w1[rk]>w0[rk]+40&&JSON.stringify(H)===JSON.stringify(r0),`${rk}: ${w0[rk]}→${w1[rk]}px`);
 await openPick();
 await p.click(`#colspick [data-cmv="${rk}:1"]`);await wait(450);
 const w2=await wOf();H=await heads();
 ok('הרוחב שנקבע ביד זז יחד עם העמודה',H[3]===rk&&Math.abs(w2[rk]-w1[rk])<=3&&Math.abs(w2[r0[3]]-w1[r0[3]])<=3,
  `${rk} ${w1[rk]}→${w2[rk]} · ${r0[3]} ${w1[r0[3]]}→${w2[r0[3]]}`);
 await p.dragAndDrop(`#tbl thead th[data-ck="${rk}"]`,`#tbl thead th[data-ck="${H[H.length-1]}"]`);await wait(450);
 const w3=await wOf();H=await heads();
 ok('גם בגרירה — הרוחב הולך עם העמודה',H[H.length-1]===rk&&Math.abs(w3[rk]-w1[rk])<=3,`${rk} ${w1[rk]}→${w3[rk]}`);

 /* 7 · שמירה: לכל תצוגה בנפרד, ושורדת טעינה מחדש */
 const dec=await heads();
 const SUP=await p.evaluate(()=>{groupBy='supplier';render();const v=colViewKey(),got=visCols();groupBy='decision';render();
  return {v,def:JSON.stringify(got)===JSON.stringify(SCOL_DEF.sup.filter(colUsable))}});
 ok('תצוגה אחרת («לפי ספק») לא הושפעה',SUP.v==='sup'&&SUP.def,SUP.v);
 await closePick();await load();
 H=await heads();
 ok('אחרי טעינה מחדש — אותו סדר',JSON.stringify(H)===JSON.stringify(dec),H.join(','));

 /* 8 · טבלת התנועות: נקודת התנועה והמק״ט נשארים ראשונים; FOB לא יוצא לגיליון */
 await p.evaluate(()=>{const a=[...document.querySelectorAll('#tabs .tab')].find(t=>t.dataset.m==='catalog');if(a)a.click()});await wait(400);
 await p.evaluate(()=>{const s2=document.querySelector('#railsub .railsub2[data-m="moves"]');if(s2)s2.click()});await wait(700);
 const m0=await heads('#movtbl');
 ok('טבלת התנועות: נקודה ומק״ט מוצמדים, השאר נגררים',
  await p.evaluate(()=>[...document.querySelectorAll('#movtbl thead th[data-ck]')].map(t=>!!t.getAttribute('draggable')).join(','))
   ===m0.map((k,i)=>i>1).join(','),m0.slice(0,4).join(','));
 await p.dragAndDrop(`#movtbl thead th[data-ck="${m0[m0.length-1]}"]`,`#movtbl thead th[data-ck="${m0[2]}"]`);await wait(500);
 H=await heads('#movtbl');
 ok('תנועות · גרירה מזיזה, והנקודה והמק״ט נשארים ראשונים',H[2]===m0[m0.length-1]&&H[0]===m0[0]&&H[1]===m0[1],H.slice(0,4).join(','));
 await openPick();
 await p.click('#colspick input[data-ck="fob"]');await wait(500);
 const MX=await p.evaluate(()=>({head:movExportRows()[0],th:[...document.querySelectorAll('#movtbl thead th[data-ck]')].map(t=>t.dataset.ck),
  exp:visCols('mv_'+movTab).filter(k=>SCOL[k]&&(SCOL[k].xval||SCOL[k].val)&&k!=='fob').map(k=>SCOL[k].t)}));
 ok('תנועות · «מחיר FOB» במסך ולא בגיליון',MX.th.includes('fob')&&!MX.head.includes('מחיר FOB')&&JSON.stringify(MX.head)===JSON.stringify(MX.exp),MX.head.join('|'));
 await p.evaluate(()=>{colsOpen(false);colReset(colViewKey());renderMoves()});await wait(400);
 ok('תנועות · חזרה לברירת המחדל',JSON.stringify(await heads('#movtbl'))===JSON.stringify(m0));

 /* 9 · «קטלוג פריטים» (טבלה כללית, בלי עמודת #): אותו מנגנון, והמק״ט נשאר ראשון */
 await p.evaluate(()=>setMode('catalog'));await wait(600);
 const c0=await heads();
 ok('קטלוג: «מק״ט» אינו נגרר, השאר נגררים',c0[0]==='cPn'&&await p.evaluate(()=>[...document.querySelectorAll('#tbl thead th[data-ck]')]
  .every((t,i)=>(i===0)!==(t.getAttribute('draggable')==='true'))),c0.slice(0,3).join(','));
 await p.dragAndDrop(`#tbl thead th[data-ck="${c0[c0.length-1]}"]`,`#tbl thead th[data-ck="${c0[1]}"]`);await wait(500);
 H=await heads();
 ok('קטלוג · גרירה מזיזה, והמק״ט נשאר ראשון',H[1]===c0[c0.length-1]&&H[0]==='cPn'&&H.length===c0.length,H.slice(0,3).join(','));
 await p.evaluate(()=>{colReset('catalog');render()});await wait(400);
 ok('קטלוג · חזרה לברירת המחדל',JSON.stringify(await heads())===JSON.stringify(c0));

 /* 10 · חזרה לברירת המחדל; אחסון חסום אינו שובר */
 await p.evaluate(()=>setMode('today'));await wait(500);
 await openPick();await p.click('#colspick [data-cpa="reset"]');await wait(450);
 H=await heads();
 ok('«חזרה לברירת המחדל» מחזירה בדיוק את הסדר המקורי',JSON.stringify(H)===JSON.stringify(D.got),H.join(','));
 await closePick();
 const blocked=await p.evaluate(()=>{const o=Storage.prototype.setItem,g=Storage.prototype.getItem;
  Storage.prototype.setItem=()=>{throw new Error('blocked')};Storage.prototype.getItem=()=>{throw new Error('blocked')};
  let okk=true;try{colStep('dec','desc',1);colMove('dec','age','desc');render()}catch(e){okk=false}
  Storage.prototype.setItem=o;Storage.prototype.getItem=g;return okk&&!!document.querySelector('#tbl tbody tr[data-i]')});
 ok('אחסון חסום — הטבלה ממשיכה לעבוד',blocked);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await p.evaluate(()=>{try{localStorage.removeItem('planner_cols_v1');localStorage.removeItem('planner_colw_v1')}catch(_){}});
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
