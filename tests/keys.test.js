/* ============ מקלדת ו«טופל» — S1 ============
   Enter ורווח סימנו «טופל»: פעולה שמוציאה פריט מהתור, על המקש הכי קל
   ללחוץ בטעות, בלי ביטול. עכשיו:
     Enter  פותח את הפריט שעומדים עליו (גם אחרי Escape — אותה שורה)
     רווח   לא עושה כלום
     D      מסמן «טופל», גם במקלדת עברית (אותו מקש = «ג»)
     «בטל»  בהודעה, 6 שניות, מחזיר את הסימון בדיוק למה שהיה
   ההודעה מוכרזת לקורא מסך (role=status, aria-live=polite). */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''))};
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?nobrief=1');
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');
 await p.evaluate(()=>{try{setMode('today')}catch(_){}});
 await p.waitForSelector('#tbl tbody tr[data-i] [data-done]',{timeout:120000});
 await p.waitForTimeout(500);
 await p.evaluate(()=>document.activeElement&&document.activeElement.blur());
 const st=()=>p.evaluate(()=>{const rows=[...document.querySelectorAll('#tbl tbody tr[data-i]')];
  const i=rows.findIndex(r=>r.classList.contains('sel'));
  const t=document.getElementById('toast');
  return {i,n:rows.length,open:document.body.classList.contains('dopen'),
   pn:CURRENT_DETAIL?CURRENT_DETAIL.pn:null,marks:JSON.stringify(MARKS),
   toast:t.classList.contains('show')?t.textContent.trim():'',undo:t.classList.contains('show')&&!!t.querySelector('.tundo')}});
 const key=k=>p.keyboard.press(k).then(()=>p.waitForTimeout(250));

 /* ---- Enter פותח ---- */
 await key('ArrowDown');await key('ArrowDown');
 const s0=await st();
 ok('חץ בוחר שורה ופותח את הכרטיס',s0.i===1&&s0.open,`שורה ${s0.i+1}`);
 await key('Escape');
 const s1=await st();ok('Escape סוגר את הכרטיס',!s1.open&&s1.i<0);
 await key('Enter');
 const s2=await st();
 ok('Enter פותח שוב את אותו פריט',s2.open&&s2.i===s0.i&&s2.pn===s0.pn,`שורה ${s2.i+1}`);
 ok('Enter לא מסמן «טופל»',s2.marks===s0.marks&&s2.n===s0.n,`${s0.n} → ${s2.n} שורות`);

 /* ---- רווח לא עושה כלום ---- */
 const y0=await p.evaluate(()=>scrollY);
 await key('Space');
 const s3=await st();
 ok('רווח לא מסמן «טופל»',s3.marks===s0.marks&&s3.n===s0.n&&s3.pn===s0.pn,`${s0.n} → ${s3.n} שורות`);
 ok('ורווח לא מגלגל את העמוד',await p.evaluate(()=>scrollY)===y0);

 /* ---- D מסמן, ו«בטל» מחזיר בדיוק ---- */
 const before=s3;
 await key('KeyD');
 const s4=await st();
 ok('D מסמן «טופל» — הפריט יצא מהתור',s4.n===before.n-1&&JSON.parse(s4.marks)&&
   Object.values(JSON.parse(s4.marks)).some(m=>m.t==='handled'),`${before.n} → ${s4.n}`);
 ok('והבחירה עברה לפריט הבא באותו מקום',s4.i===before.i&&s4.pn!==before.pn);
 ok('ההודעה אומרת מה סומן ומציעה «בטל»',s4.undo&&s4.toast.includes(before.pn)&&/בטל/.test(s4.toast),s4.toast);
 const aria=await p.evaluate(()=>{const t=document.getElementById('toast');return t.getAttribute('role')+'/'+t.getAttribute('aria-live')});
 ok('ההודעה מוכרזת לקורא מסך',aria==='status/polite',aria);
 await p.waitForTimeout(4500);                 /* 4.75 שנ׳ אחרי הסימון — עדיין בתוך החלון */
 ok('«בטל» עדיין זמין אחרי 4.5 שניות',(await st()).undo);
 await p.click('#toast .tundo');await p.waitForTimeout(300);
 const s5=await st();
 ok('«בטל» מחזיר את הסימון בדיוק למה שהיה',s5.marks===before.marks);
 ok('והפריט חזר לתור, לאותו מקום, והוא הנבחר',s5.n===before.n&&s5.i===before.i&&s5.pn===before.pn,`שורה ${s5.i+1}`);

 /* ---- מקלדת עברית: אותו מקש פיזי נותן «ג» ---- */
 await p.evaluate(()=>document.activeElement&&document.activeElement.blur());
 await p.evaluate(()=>document.dispatchEvent(new KeyboardEvent('keydown',{key:'ג',code:'KeyD',bubbles:true,cancelable:true})));
 await p.waitForTimeout(250);
 const s6=await st();
 ok('D במקלדת עברית («ג») מסמן גם הוא',s6.n===before.n-1&&s6.undo);
 /* ---- אחרי 6 שניות ההודעה נעלמת והסימון נשאר ---- */
 await p.waitForTimeout(6300);
 const s7=await st();
 ok('אחרי 6 שניות «בטל» נעלם — והסימון נשאר',!s7.toast&&!s7.undo&&s7.n===before.n-1&&s7.marks!==before.marks);
 const ms=await p.evaluate(()=>UNDO_MS);ok('חלון הביטול הוא 6 שניות',ms===6000,ms+'ms');

 /* ---- D בתוך שדה חיפוש הוא אות, לא פעולה ---- */
 const m0=(await st()).marks;
 await p.focus('#q');await p.keyboard.type('d');await p.waitForTimeout(250);
 ok('D בשדה החיפוש נכתב ולא מסמן',(await st()).marks===m0&&await p.evaluate(()=>document.getElementById('q').value)==='d');
 await p.evaluate(()=>{const q=document.getElementById('q');q.value='';q.blur()});
 /* ---- Shift+D אינו D ---- */
 await key('ArrowDown');const m1=(await st()).marks;
 await key('Shift+KeyD');ok('Shift+D לא מסמן',(await st()).marks===m1);

 /* ---- כפתור «✓ טופל» בשורה — אותה פעולה, אותו ביטול ---- */
 await p.waitForTimeout(1800);
 const b0=await st();
 const pnBtn=await p.evaluate(()=>{const tr=document.querySelector('#tbl tbody tr[data-i]');
  const k=tr.querySelector('[data-done]').dataset.done;tr.querySelector('[data-done]').click();return k});
 await p.waitForTimeout(250);
 const b1=await st();
 ok('כפתור «✓ טופל» בשורה מציע גם הוא «בטל»',b1.undo&&b1.n===b0.n-1);
 await p.click('#toast .tundo');await p.waitForTimeout(300);
 ok('ו«בטל» מחזיר אותו',(await st()).marks===b0.marks,pnBtn.split('|').pop());

 ok('אפס שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
