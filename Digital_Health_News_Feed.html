<!DOCTYPE html><html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Chat with the Board Pack — AI Q&A</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Lora:wght@600;700&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box}
:root{--bg:#0f1216;--panel:#161a21;--panel2:#1b2028;--line:#272d38;--ink:#e9edf3;--muted:#97a2b2;--soft:#c4ccd8;
 --accent:#ff6719;--teal:#2dd4bf;--serif:'Lora',Georgia,serif;--sans:'Inter',-apple-system,Segoe UI,Roboto,sans-serif;color-scheme:dark}
body{margin:0;font-family:var(--sans);background:var(--bg);color:var(--ink);-webkit-font-smoothing:antialiased}
a{color:var(--teal);text-decoration:none}
.top{display:flex;align-items:center;justify-content:space-between;padding:14px 22px;border-bottom:1px solid var(--line);background:#0b0d11;position:sticky;top:0;z-index:10}
.top .t{font-family:var(--serif);font-weight:700;font-size:17px}
.top .t span{color:var(--teal)}
.top a{font-size:12.5px;font-weight:600;color:var(--muted)}
.wrap{max-width:1080px;margin:0 auto;padding:20px 22px 60px;display:grid;grid-template-columns:300px 1fr;gap:20px}
@media(max-width:820px){.wrap{grid-template-columns:1fr}}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:16px 18px}
.kick{font-size:10.5px;letter-spacing:.1em;text-transform:uppercase;color:var(--accent);font-weight:800;margin:0 0 10px}
.bp h2{font-family:var(--serif);font-size:17px;margin:0 0 2px}
.bp .sub{font-size:11.5px;color:var(--muted);margin:0 0 12px}
.bp .row{display:flex;justify-content:space-between;font-size:12.5px;padding:6px 0;border-top:1px solid var(--line)}
.bp .row b{font-weight:700}.bp .row .up{color:#34d399}.bp .row .down{color:#fb7185}
.note{font-size:11px;color:var(--muted);margin-top:12px;line-height:1.5}
/* chat */
.chat{display:flex;flex-direction:column;height:560px;background:var(--panel);border:1px solid var(--line);border-radius:14px;overflow:hidden}
.msgs{flex:1;overflow-y:auto;padding:18px;display:flex;flex-direction:column;gap:14px}
.m{max-width:84%;font-size:13.5px;line-height:1.55;padding:11px 14px;border-radius:13px}
.m.u{align-self:flex-end;background:var(--accent);color:#fff;border-bottom-right-radius:4px}
.m.a{align-self:flex-start;background:var(--panel2);border:1px solid var(--line);border-bottom-left-radius:4px}
.m.a .fig{font-family:var(--serif);font-size:20px;font-weight:700;color:var(--teal);display:block;margin:2px 0 4px}
.m.a .src{display:inline-block;margin-top:9px;font-size:10.5px;font-weight:700;letter-spacing:.03em;text-transform:uppercase;color:#7fbdb6;background:rgba(45,212,191,.1);border:1px solid rgba(45,212,191,.25);padding:3px 9px;border-radius:999px}
.typing{align-self:flex-start;color:var(--muted);font-size:12px;padding:6px 14px}
.sugg{display:flex;flex-wrap:wrap;gap:7px;padding:0 18px 12px}
.chip{font-size:11.5px;font-weight:600;color:var(--soft);background:var(--panel2);border:1px solid var(--line);padding:6px 11px;border-radius:999px;cursor:pointer;transition:.15s}
.chip:hover{border-color:var(--teal);color:#fff}
.inbar{display:flex;gap:8px;padding:12px 14px;border-top:1px solid var(--line);background:#0e1218}
.inbar input{flex:1;background:var(--bg);border:1px solid var(--line);border-radius:10px;color:var(--ink);font-size:13.5px;padding:11px 13px;font-family:inherit}
.inbar input:focus{outline:none;border-color:var(--teal)}
.inbar button{background:var(--teal);color:#08201d;border:none;border-radius:10px;font-weight:800;font-size:13.5px;padding:0 18px;cursor:pointer}
.foot{max-width:1080px;margin:0 auto;padding:0 22px 40px;font-size:11px;color:var(--muted);text-align:center}
</style></head><body>
<div class="top"><div class="t">Chat with the <span>Board Pack</span></div><a href="index.html">← Back to portfolio</a></div>
<div class="wrap">
  <div class="panel bp">
    <p class="kick">The board pack</p>
    <h2>Cleardose Health</h2>
    <p class="sub">Series B · Q2 2026 · powered by my synthetic GL</p>
    <div class="row"><span>Revenue (Q2)</span><b>$3.85M</b></div>
    <div class="row"><span>ARR</span><b>$5.28M <span class="up">+75%</span></b></div>
    <div class="row"><span>Gross margin</span><b>50%</b></div>
    <div class="row"><span>Net income (Q2)</span><b class="down">−$2.4M</b></div>
    <div class="row"><span>Cash / runway</span><b class="down">$2.4M · ~3 mo</b></div>
    <div class="row"><span>VBC members</span><b>8,555 <span class="up">+2.5x</span></b></div>
    <div class="row"><span>Patient collections</span><b class="down">62% · DSO 16.9d</b></div>
    <div class="row"><span>SaaS customers</span><b>138</b></div>
    <p class="note">Grounded in the dimensioned <b>Cleardose GL</b> I built (the same data behind my FP&amp;A dashboard). This demo answers from the Q2 pack via retrieval; free-form questions route to an LLM over the full GL when deployed.</p>
  </div>

  <div>
    <div class="chat">
      <div class="msgs" id="msgs"></div>
      <div class="sugg" id="sugg"></div>
      <div class="inbar"><input id="q" placeholder="Ask the board pack…  e.g. why did gross margin fall?" autocomplete="off"><button onclick="send()">Ask</button></div>
    </div>
  </div>
</div>
<p class="foot">Demo · synthetic data. Built by Mark Kempa as an example of AI-native finance tooling — retrieval now, full LLM Q&A on deploy.</p>
<script>
const CHUNKS=[
 {k:["revenue","top line","growth","arr","sales","mix","service line"],fig:"Q2 rev $3.85M · ARR $5.28M",
  a:"Q2-2026 revenue was $3.85M. The mix: Value-Based Care capitation ~$1.35M (fastest-growing), SaaS subscription ~$1.29M (ARR $5.28M, up 75% since Jan-25), telehealth ~$0.81M, and CCM ~$0.41M.",src:"GL_Detail + KPIs, Q2-26"},
 {k:["gross margin","margin","gm","cogs","cost of goods"],fig:"Gross margin 50%",
  a:"Gross margin is ~50% — COGS of $1.91M on $3.85M revenue. The capitated VBC line carries medical cost, which dilutes the SaaS-level margins as VBC mixes up.",src:"GL_Detail, Q2-26"},
 {k:["runway","cash","burn","how long","months","out of money"],fig:"$2.4M cash · ~3 months",
  a:"This is the headline risk: cash is down to ~$2.4M against a burn of roughly $0.8M/month, leaving only ~3 months of runway. A raise or bridge is needed now — not in H2.",src:"Bank_Activity, Jun-26"},
 {k:["net","ebitda","loss","profit","profitability","bottom line"],fig:"Net −$2.4M (Q2)",
  a:"Q2 net loss was $2.4M (operating loss ~$2.5M before $0.1M interest income), on $4.45M of opex. Losses are roughly flat — the business is scaling revenue but hasn't bent the cost curve yet.",src:"GL_Detail, Q2-26"},
 {k:["collection","collections","dso","ar","receivable","patient","cash collection"],fig:"Collections 62% · DSO 16.9d",
  a:"A real concern: patient collection rate has slipped from 74% to ~62% over 18 months, and DSO has more than doubled to 16.9 days. Patient AR has built to ~$514K — a working-capital drag worsening the cash crunch.",src:"Drivers_KPIs, Jun-26"},
 {k:["vbc","value based","capitation","members","risk"],fig:"8,555 VBC members",
  a:"VBC is the growth engine — 8,555 capitated members (2.5x since Jan-25) generating ~$484K/month. It's now the largest revenue line, but it also brings medical-cost risk that pressures gross margin.",src:"Drivers_KPIs, Q2-26"},
 {k:["saas","subscription","mrr","customers","churn","logo"],fig:"ARR $5.28M · 138 customers",
  a:"SaaS subscription ARR is $5.28M across 138 customers, growing steadily with logo churn running ~1.5–2%/month. New logos are 4–8 per month.",src:"Drivers_KPIs, Q2-26"},
 {k:["opex","operating expense","spend","headcount","hiring","fte"],fig:"Opex $4.45M · 83 FTEs",
  a:"Q2 opex was $4.45M against headcount of 83 (up from 51 in Jan-25). Sales & marketing and G&A payroll are the largest cash outflows — the main levers if runway needs extending.",src:"GL_Detail + Bank, Q2-26"},
 {k:["marketing","cac","acquisition","spend marketing"],fig:"Mktg ~$150K/mo",
  a:"Marketing spend runs ~$130–155K/month driving 4–8 new SaaS logos plus patient/member acquisition. Given the cash position, marketing efficiency is a key lever to scrutinize.",src:"Drivers_KPIs, Q2-26"},
 {k:["raise","series","fundraise","capital","bridge","when raising","next round"],fig:"Raise needed now",
  a:"With ~3 months of runway, the priority is an immediate raise or bridge. The story to investors is strong VBC growth and $5.3M ARR — but it must be paired with a credible plan to fix collections and bend opex.",src:"Board narrative, Q2-26"}
];
const SUGG=["How much runway do we have?","What's driving revenue?","Why are collections down?","How is VBC performing?","What's our gross margin?","When do we need to raise?"];
const msgs=document.getElementById('msgs');
function add(cls,html){const d=document.createElement('div');d.className='m '+cls;d.innerHTML=html;msgs.appendChild(d);msgs.scrollTop=msgs.scrollHeight;return d;}
function localAnswer(q){const t=q.toLowerCase();let best=null,bs=0;
 for(const c of CHUNKS){let sc=0;for(const kw of c.k){if(t.includes(kw))sc+=kw.length;}if(sc>bs){bs=sc;best=c;}}
 if(!best||bs<3)return {a:"I couldn't find that in the Q3 pack. Try a suggested question — on the deployed version, free-form questions route to an LLM that reads the full board pack.",src:null,fig:null};
 return best;}
function answer(q){
 const typ=document.createElement('div');typ.className='typing';typ.textContent='Reading the board pack…';msgs.appendChild(typ);msgs.scrollTop=msgs.scrollHeight;
 // on deploy this calls the LLM; locally it falls back to retrieval
 const done=(r)=>{typ.remove();let h='';if(r.fig)h+='<span class="fig">'+r.fig+'</span>';h+=r.a;if(r.src)h+='<span class="src">Source · '+r.src+'</span>';add('a',h);};
 fetch('/api/ask',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({q})})
  .then(x=>x.ok?x.json():Promise.reject()).then(d=>done({a:d.answer,src:d.source||'Board pack (LLM)',fig:d.fig||null}))
  .catch(()=>setTimeout(()=>done(localAnswer(q)),550));
}
function send(){const i=document.getElementById('q');const q=i.value.trim();if(!q)return;add('u',q);i.value='';answer(q);}
document.getElementById('q').addEventListener('keydown',e=>{if(e.key==='Enter')send();});
document.getElementById('sugg').innerHTML=SUGG.map(s=>`<span class="chip" onclick="document.getElementById('q').value='${s}';send()">${s}</span>`).join('');
add('a','<span class="fig">Ask me about Cleardose Q2 👋</span>Ask me about Cleardose Health&#39;s Q2 — revenue &amp; mix, gross margin, runway, collections/DSO, VBC, SaaS, opex, and the raise. Every answer cites its source in the GL. Try a suggestion below.');
</script>
</body></html>
