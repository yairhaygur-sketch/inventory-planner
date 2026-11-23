/* ============ TP_STOCK_TOP — מצב המלאי ב-SAP בראש הכרטיס (הערת משתמשים 23.11, אושר) ============
   בכרטיס של תצוגת 4 הנושאים «מצב המלאי» היה בתחתית, וצריך היה לגלול כדי לראות אותו.
   עכשיו הוא הדבר הראשון בכרטיס: השורה (לקוח · פנוי · רכש פתוח · בהעברה · פער) וסה״כ המלאי.
   «תנועה אחרונה» (מכירה וכניסה אחרונות) נשארת למטה, אחרי ההחלטה. במסך בלבד. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');

 /* בכל נושא: עשר השורות הראשונות — הכרטיס של כל אחת */
 const R=await p.evaluate(()=>{const res=[];const n=v=>Math.max(0,v||0).toLocaleString('he-IL');
  for(const t of Object.keys(TOPICS)){TP.tab=t;TP.cat={};TP.sort={};TP.q='';TP.f={};TP.view='';tpRender();
   for(const r of (TP.list||[]).slice(0,10)){tpdShow(r);const db=document.querySelector('#tpd .db'),secs=[...db.children];
    const first=secs[0],h3=first&&first.querySelector('h3'),cells=first?[...first.querySelectorAll('.cstrip .cs')].map(c=>[c.querySelector('i').textContent,c.querySelector('b').textContent]):[];
    const want=[['לקוח',n(r.cust)],['פנוי',n(r.free)],['רכש פתוח',n(r.po)],['בהעברה',r.transferKnown?n(r.transfer):'לא ידוע'],['פער',n(r.miss)]];
    const iQ3=secs.findIndex(s=>s.classList.contains('q3')),iMove=secs.findIndex(s=>s.classList.contains('tpmove'));
    const stockHeads=[...db.querySelectorAll('h3')].filter(h=>/^מצב המלאי/.test(h.textContent.trim())).length;
    res.push({t,pn:r.pn,first:!!first&&first.classList.contains('tpstock')&&!!h3&&/^מצב המלאי/.test(h3.textContent.trim()),
     cells:JSON.stringify(cells)===JSON.stringify(want),note:!!first&&first.querySelector('.note').textContent.includes(`סה״כ מלאי ${tpN(r.stock)}`),
     q3next:iQ3===1,once:stockHeads===1,
     move:movePanel(r)?(iMove>iQ3&&/תנועה אחרונה/.test(secs[iMove].textContent)):iMove<0});
    tpdClose()}}
  return res});
 const bad1=k=>R.filter(x=>!x[k]).map(x=>`${x.t}:${x.pn}`).slice(0,3).join(', ');
 ok('בכל כרטיס (עשרה בכל נושא): הקטע הראשון הוא «מצב המלאי (SAP)»',R.length>=30&&R.every(x=>x.first),`${R.length} כרטיסים ${bad1('first')}`);
 ok('השורה: לקוח · פנוי · רכש פתוח · בהעברה · פער — הערכים של הפריט',R.every(x=>x.cells),bad1('cells'));
 ok('מתחת לשורה: «סה״כ מלאי» של הפריט',R.every(x=>x.note),bad1('note'));
 ok('מיד אחריו — ההחלטה («מה הבעיה / למה זה חשוב / מה לעשות»)',R.every(x=>x.q3next),bad1('q3next'));
 ok('«מצב המלאי» מופיע פעם אחת בכרטיס — לא נשאר עותק למטה',R.every(x=>x.once),bad1('once'));
 ok('«תנועה אחרונה» נשארת למטה, אחרי ההחלטה (כשיש לפריט תנועות)',R.every(x=>x.move),bad1('move'));

 /* בלי גלילה: מצב המלאי וההחלטה בתוך המסך גם ב-1366×768 */
 await p.setViewportSize({width:1366,height:768});
 const V=await p.evaluate(()=>{TP.tab='cust';TP.cat={};tpRender();tpdShow(TP.list[0]);
  const st=document.querySelector('#tpd .tpstock').getBoundingClientRect(),act=document.querySelector('#tpd .qa.act').getBoundingClientRect();
  const v={st:Math.round(st.bottom),act:Math.round(act.top),h:innerHeight};tpdClose();return v});
 ok('ב-1366×768: מצב המלאי כולו בתוך המסך, ו«מה לעשות» מתחיל בתוך המסך',V.st<V.h&&V.act<V.h,JSON.stringify(V));

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
