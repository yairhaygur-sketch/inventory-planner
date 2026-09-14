const {chromium}=require('playwright');const fs=require('fs');
const path=require('path');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const hex=h=>{const m=h.match(/\d+/g);return m?[+m[0],+m[1],+m[2]]:[0,0,0]};
const L=c=>{const s=c.map(v=>{v/=255;return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4)});
 return 0.2126*s[0]+0.7152*s[1]+0.0722*s[2]};
const R=(a,b)=>{const x=L(hex(a)),y=L(hex(b));return (Math.max(x,y)+0.05)/(Math.min(x,y)+0.05)};
let total=0;
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1600,height:1000}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();
 await p.goto('file://'+path.join(__dirname,'..','index.html'));
 await p.setInputFiles('#f',path.join(__dirname,'zmrp-demo.xlsx'));
 await p.waitForSelector('#tbl tbody tr[data-i]',{timeout:120000});
 await p.waitForTimeout(1200);
 for(const mode of ['light','dark']){
  await p.evaluate(m=>document.body.classList.toggle('dark',m==='dark'),mode);
  await p.waitForTimeout(300);
  // כל אלמנט גלוי עם טקסט: מודדים צבע מול הרקע האפקטיבי
  const bad=await p.evaluate(()=>{
   const eff=el=>{let n=el;while(n&&n!==document.documentElement){
     const bg=getComputedStyle(n).backgroundColor;
     if(bg&&!/rgba\(0, 0, 0, 0\)|transparent/.test(bg))return bg;n=n.parentElement}
    return getComputedStyle(document.body).backgroundColor};
   const out=[];
   for(const el of document.querySelectorAll('body *')){
    const cs=getComputedStyle(el);
    if(cs.display==='none'||cs.visibility==='hidden'||+cs.opacity===0)continue;
    const r=el.getBoundingClientRect(); if(r.width<4||r.height<4)continue;
    const t=[...el.childNodes].filter(n=>n.nodeType===3&&n.textContent.trim()).map(n=>n.textContent.trim()).join(' ');
    if(!t)continue;
    out.push({sel:el.tagName.toLowerCase()+(el.className&&typeof el.className==='string'?'.'+el.className.trim().split(/\s+/).join('.'):''),
      txt:t.slice(0,22),fg:cs.color,bg:eff(el),size:parseFloat(cs.fontSize),weight:cs.fontWeight});
   }
   return out;
  });
  const fails=[];const seen=new Set();
  for(const e of bad){
   const r=R(e.fg,e.bg);
   const large=e.size>=24||(e.size>=18.66&&+e.weight>=700);
   const need=large?3:4.5;
   if(r<need){const k=e.sel+'|'+e.fg+'|'+e.bg;if(seen.has(k))continue;seen.add(k);
    fails.push(`${r.toFixed(2)} (דרוש ${need}) ${e.sel}  «${e.txt}»  ${e.fg} על ${e.bg}`)}
  }
  console.log(`\n=== ${mode} · ${bad.length} אלמנטים · ${fails.length} כשלים ייחודיים ===`);
  total+=fails.length;
  fails.sort((a,b)=>parseFloat(a)-parseFloat(b)).slice(0,25).forEach(f=>console.log('  '+f));
 }
 const bad=total>0;
 await b.close();
 process.exit(total>0?1:0);
})();
