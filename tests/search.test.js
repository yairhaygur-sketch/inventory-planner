const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[],ok=(n,c,x)=>out.push((c?'PASS':'FAIL')+' · '+n+(x?'  ['+x+']':''));
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH});
 const ctx=await b.newContext({viewport:{width:1512,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html'));
 await p.setInputFiles('#f',SD+'/zmrp-demo.xlsx');await p.waitForTimeout(2600);

 // נורמליזציה
 const nz=await p.evaluate(()=>({
   dash:norm('1001-2233',true), quote:norm('מק"ט',true), fin:norm('םןץ'),
   nikud:norm('שָׁלוֹם')}));
 ok('מקפים נופלים בנרמול',nz.dash==='10012233',nz.dash);
 ok('גרשיים נופלים',nz.quote==='מקט',nz.quote);
 ok('אותיות סופיות מנורמלות',nz.fin==='מנצ',nz.fin);
 ok('ניקוד נופל',nz.nikud==='שלומ',nz.nikud);  /* ם→מ בנרמול אותיות סופיות */

 // חיפוש מק״ט עם מקף
 const r=await p.evaluate(()=>{const pn=ALL[0].pn;
   const dashed=pn.slice(0,3)+'-'+pn.slice(3);
   const q=parseQ(dashed);
   return {pn,dashed,hit:!!(q&&scoreRow(ALL[0],q))}});
 ok('מק״ט עם מקף מוצא את הפריט',r.hit,`${r.dashed} → ${r.pn}`);

 // סדר אסימונים
 const tk=await p.evaluate(()=>{const w=ALL.find(x=>x.desc.trim().split(/\s+/).length>=2);
   if(!w)return null;const p2=w.desc.trim().split(/\s+/).slice(0,2);
   const a=scoreRow(w,parseQ(p2.join(' '))),b=scoreRow(w,parseQ(p2.slice().reverse().join(' ')));
   return {q:p2.join(' '),a,b}});
 ok('סדר האסימונים אינו משנה',tk&&tk.a>0&&tk.b>0,tk?`${tk.q} → ${tk.a} / ${tk.b}`:'—');

 // קידומת שדה + החרגה
 const fx=await p.evaluate(()=>{const s=ALL.find(x=>x.supplier);
   const one=scoreRow(s,parseQ('ספק:'+s.supplier));
   const other=ALL.find(x=>x.supplier&&x.supplier!==s.supplier);
   const zero=other?scoreRow(other,parseQ('ספק:'+s.supplier)):-1;
   const neg=scoreRow(s,parseQ('-'+s.supplier));
   return {sup:s.supplier,one,zero,neg}});
 ok('קידומת "ספק:" מסננת לספק הנכון',fx.one>0&&fx.zero===0,`${fx.sup} → ${fx.one} / ${fx.zero}`);
 ok('"-מילה" מחריגה',fx.neg===0,'ציון '+fx.neg);

 // דירוג: מק״ט מלא מנצח התאמה חלקית בתיאור
 const rk=await p.evaluate(()=>{const t=ALL[5];
   const res=searchAll(parseQ(t.pn),8);return {first:res.rows[0]&&res.rows[0].pn,want:t.pn,n:res.n}});
 ok('מק״ט מלא עולה ראשון בדירוג',rk.first===rk.want,`${rk.first} מול ${rk.want}`);

 // ממשק: חלונית, הדגשה, ניווט מקלדת
 await p.evaluate(()=>{cur='all';render()});await p.waitForTimeout(300);
 await p.keyboard.press('/');await p.waitForTimeout(200);
 const focused=await p.evaluate(()=>document.activeElement&&document.activeElement.id);
 ok('«/» ממקד את תיבת החיפוש',focused==='q',focused);
 /* מונח שמופיע בכמה פריטים — כדי לבדוק דירוג, הדגשה וניווט */
 const term=await p.evaluate(()=>{const c={};for(const r of ALL){const w=(r.desc||'').trim().split(/\s+/)[0];if(w&&w.length>2)c[w]=(c[w]||0)+1}
   return Object.entries(c).sort((a,b)=>b[1]-a[1])[0][0]});
 await p.fill('#q',term);await p.waitForTimeout(400);
 const sug=await p.evaluate(()=>{const el=document.getElementById('qsug');
   const b=el.getBoundingClientRect();
   return {open:el.classList.contains('open'),items:el.querySelectorAll('.sgi').length,
     inView:b.left>=0&&b.right<=window.innerWidth&&b.top>=0,
     marks:document.querySelectorAll('#tbl mark').length,
     stat:document.getElementById('qstat').textContent}});
 ok('חלונית התוצאות נפתחת',sug.open&&sug.items>0,sug.items+' תוצאות');
 ok('החלונית אינה חורגת מהמסך',sug.inView);
 ok('ההתאמות מודגשות בטבלה',sug.marks>0,sug.marks+' סימונים');
 ok('שורת המצב מציגה ספירה',/\d/.test(sug.stat),sug.stat);
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(120);
 const onIdx=await p.evaluate(()=>[...document.querySelectorAll('#qsug .sgi')].findIndex(e=>e.classList.contains('on')));
 ok('↓ מסמן את התוצאה הראשונה',onIdx===0,'index '+onIdx);
 // קפיצה לכל התוצאות כשהן מחוץ לטאב הפעיל
 const jump=await p.evaluate(async()=>{cur='today';render();
   document.getElementById('q').focus();renderSug();
   const el=document.getElementById('sgall');if(!el)return {shown:false};
   el.click();return {shown:true,cur}});
 ok('קיימת קפיצה ל"כל הפריטים" כשהתוצאות מחוץ לטאב',jump.shown&&jump.cur==='all',JSON.stringify(jump));
 await p.evaluate(()=>{cur='all';render();document.getElementById('q').focus();renderSug()});
 await p.waitForTimeout(250);
 await p.keyboard.press('ArrowDown');await p.waitForTimeout(120);
 await p.keyboard.press('Enter');await p.waitForTimeout(400);
 const opened=await p.evaluate(()=>({closed:!document.getElementById('qsug').classList.contains('open'),
   detail:document.getElementById('detail').innerText.slice(0,40)}));
 ok('↵ פותח כרטיס פריט וסוגר את החלונית',opened.closed&&opened.detail.length>3,opened.detail.replace(/\n/g,' '));
 await p.click('#q');await p.keyboard.press('Escape');await p.waitForTimeout(300);
 const cleared=await p.evaluate(()=>({v:document.getElementById('q').value,
   open:document.getElementById('qsug').classList.contains('open')}));
 ok('Esc מנקה וסוגר',cleared.v===''&&!cleared.open,`"${cleared.v}"`);
 const restored=await p.evaluate(()=>document.querySelectorAll('#tbl tbody tr[data-i]').length);
 ok('הרשימה חוזרת מלאה אחרי ניקוי',restored>1,restored+' שורות');

 // מנוע ההחלטות לא נגע
 const eng=await p.evaluate(()=>({prevent:Q.prevent.length,quality:Q.quality.length,
   ss:ALL.filter(r=>r.sugSS>0).length}));
 ok('הסיווג לא הושפע מהחיפוש',eng.prevent>0&&eng.ss>0,`מניעת חוסר ${eng.prevent} · SS ${eng.ss}`);
 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 console.log(out.join('\n'));
 await p.screenshot({path:SD+'/07-search.png'});
 await b.close();process.exit(out.some(x=>x.startsWith('FAIL'))?1:0)})();
