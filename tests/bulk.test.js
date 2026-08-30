const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];const ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH});
 const ctx=await b.newContext({viewport:{width:1512,height:860}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 p.on('dialog',d=>d.accept());
 await p.goto('file://'+path.join(SD,'..','index.html'));
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(2600);
 // עוברים לטאב "הון לשחרור" ומסננים
 await p.evaluate(()=>{cur='excess';render()});await p.waitForTimeout(400);
 const before=await p.evaluate(()=>({n:currentRows().length,marked:Object.keys(MARKS).length}));
 ok('כפתור הסימון הקבוצתי מופיע עם מספר',
    /סמן [\d,]+ פריטים/.test(await p.locator('#bulkBtn').innerText()),
    await p.locator('#bulkBtn').innerText());
 await p.click('#bulkBtn');await p.waitForTimeout(150);
 ok('התפריט נפתח ומציג את מספר הפריטים',
    await p.locator('#bulkMenu.open').count()>0&&/מסמן את/.test(await p.locator('#bulkMenu .bh').innerText()),
    (await p.locator('#bulkMenu .bh').innerText()).slice(0,60));
 await p.click('#bulkMenu button[data-bm="campaign"]');await p.waitForTimeout(700);
 const after=await p.evaluate(()=>({marked:Object.keys(MARKS).length,
   camp:Object.values(MARKS).filter(m=>m.t==='campaign').length,
   excess:currentRows().length}));
 ok('כל הרשימה סומנה כקמפיין',after.camp>=before.n&&after.marked>before.marked,
    `${before.marked} → ${after.marked} סימונים · ${after.camp} קמפיין`);
 ok('הפריטים יצאו מהמסלול',after.excess<before.n,`${before.n} → ${after.excess}`);
 ok('הטוסט מציע ביטול',await p.locator('.toast .tundo').count()>0);
 await p.click('.toast .tundo');await p.waitForTimeout(700);
 const undone=await p.evaluate(()=>({marked:Object.keys(MARKS).length,n:currentRows().length}));
 ok('הביטול מחזיר את המצב במדויק',undone.marked===before.marked&&undone.n===before.n,
    `סימונים ${undone.marked} (היה ${before.marked}) · רשימה ${undone.n} (היה ${before.n})`);
 ok('אין שגיאות JS',errs.length===0,errs.join(' | '));
 await b.close();console.log(out.join('\n'));
 process.exit(out.some(l=>l.startsWith('FAIL'))?1:0)})();
