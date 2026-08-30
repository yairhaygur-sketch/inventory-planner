const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];const ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html'));
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(2500);

 // 1. ניווט מקלדת ↓
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(250);
 let pn1=await p.evaluate(()=>document.querySelector('#detail .opn')?.textContent.replace('העתק','').trim());
 ok('חץ למטה בוחר פריט ומעדכן את הכרטיס',!!pn1,pn1);
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(250);
 let pn2=await p.evaluate(()=>document.querySelector('#detail .opn')?.textContent.replace('העתק','').trim());
 ok('חץ נוסף מתקדם לפריט הבא',pn2&&pn2!==pn1,pn1+' → '+pn2);
 await p.keyboard.press('ArrowUp');await p.waitForTimeout(250);
 ok('חץ למעלה חוזר אחורה',
   (await p.evaluate(()=>document.querySelector('#detail .opn')?.textContent.replace('העתק','').trim()))===pn1);

 // 2. Enter מסמן כטופל וממשיך
 const before=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
 await p.keyboard.press('Enter');await p.waitForTimeout(400);
 const after=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
 const stillSel=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr.sel').length);
 ok('Enter מסמן כטופל ומסיר מהתור',after<before,before+' → '+after);
 ok('הבחירה עוברת אוטומטית לפריט הבא',stillSel===1);

 // 3. / ממקד חיפוש
 await p.keyboard.press('/');await p.waitForTimeout(150);
 ok('"/" ממקד את שדה החיפוש',await p.evaluate(()=>document.activeElement.id==='q'));
 await p.keyboard.press('Escape');await p.waitForTimeout(150);
 ok('Escape יוצא מהשדה',await p.evaluate(()=>document.activeElement.id!=='q'));

 // 4. טאב תנועות
 await p.evaluate(()=>[...document.querySelectorAll('#tabs .tab')].find(t=>t.dataset.k==='moves')?.click());
 await p.waitForTimeout(500);
 const mv=await p.evaluate(()=>({rows:document.querySelectorAll('#movtbl tbody tr').length,
   panelShown:getComputedStyle(document.getElementById('movesPanel')).display!=='none',
   queueHidden:getComputedStyle(document.querySelector('#w_queue .qpanel')).display==='none',
   h:document.querySelector('#movesPanel .rows')?.clientHeight}));
 ok('טאב תנועות מציג את הטבלה במקום תור העבודה',mv.panelShown&&mv.queueHidden&&mv.rows>0,
    mv.rows+' שורות, גובה '+mv.h+'px');
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(300);
 ok('ניווט מקלדת עובד גם בטאב תנועות',
   await p.evaluate(()=>document.querySelectorAll('#movtbl tbody tr.sel').length===1));
 // חזרה
 await p.evaluate(()=>[...document.querySelectorAll('#tabs .tab')].find(t=>t.dataset.k==='today')?.click());
 await p.waitForTimeout(400);
 ok('חזרה לתור העבודה משחזרת את הרשימה',
   await p.evaluate(()=>getComputedStyle(document.querySelector('#w_queue .qpanel')).display!=='none'
     &&document.querySelectorAll('#tbl tbody tr[data-i]').length>0));

 // 5. גרפים מתקפלים
 await p.locator('#tbl tbody tr[data-i]').first().click();await p.waitForTimeout(400);
 const d0=await p.evaluate(()=>{const o=document.querySelector('#detail .opad');return {h:o.clientHeight,sh:o.scrollHeight,ovf:getComputedStyle(o).overflow}});
 await p.evaluate(()=>document.querySelector('#detail details[data-chart="c1"]').open=true);
 await p.waitForTimeout(400);
 const drawn=await p.evaluate(()=>{const c=document.getElementById('c1');return c&&c.width>0&&
   c.getContext('2d').getImageData(0,0,c.width,c.height).data.some(v=>v!==0)});
 ok('גרף הצריכה מצויר בפתיחת המקטע',drawn);
 ok('כרטיס הפריט גולל ולא נחתך',d0.ovf==='auto',d0.ovf);
 ok('כרטיס הפריט: פחות גלילה מהבסיס',d0.sh<950,'תוכן '+d0.sh+'px בחלון '+d0.h+'px (בסיס: 1066/328)');

 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 await p.screenshot({path:SD+'/04-after.png'});
 await b.close();console.log(out.join('\n'));
 process.exit(out.some(l=>l.startsWith('FAIL'))?1:0)})();
