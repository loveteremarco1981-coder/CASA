const $=(q,r)=> (r||document).querySelector(q);
function get(){ return fetch(CONFIG.DOGET_URL,{cache:'no-store'}).then(r=>r.json()); }
function post(event,payload){ return fetch(CONFIG.DOGET_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(Object.assign({event},payload||{}))}).then(r=>r.text()); }
function render(m){ $('#banner').textContent=(m.state||'—').replace('_',' '); $('#temp').textContent = (m.weather&&m.weather.tempC!=null)?Math.round(m.weather.tempC):'—'; $('#hum').textContent  = (m.weather&&m.weather.humidity!=null)?Math.round(m.weather.humidity):'—'; document.getElementById('pill').textContent=((m.weather&&m.weather.tempC!=null)?Math.round(m.weather.tempC)+'°C':'—°C')+' · '+((m.weather&&m.weather.windKmh!=null)?Math.round(m.weather.windKmh):'—')+' km/h'; $('#ppl').textContent = Array.isArray(m.people)?m.people.filter(p=>p.onlineSmart||p.onlineRaw).length:'—'; $('#day').textContent = (m.notte?'Notte':'Giorno');
  const fav=$('#fav'); fav.innerHTML='';
  CONFIG.FAVORITES.forEach(f=>{ const tile=document.createElement('article'); let isOn=(f.kind==='toggle'&&f.stateKey&&!!m[f.stateKey]); if(f.id==='tapparelle'){ isOn=(localStorage.getItem('ui_sh_up')==='1'); }
    tile.className='tile'+(isOn?' on':''); tile.innerHTML=`<div class="left"><div class="glyph"></div><div><div class="title">${f.label}</div><div class="sub">${f.subtitle||''}</div><div class="state">${f.id==='tapparelle'?(isOn?'Aperte':'Chiuse'):(f.kind==='toggle'?(isOn?'On':'Off'):'')}</div></div></div><button class="btn" aria-label="toggle">⏻</button>`; tile.querySelector('.btn').addEventListener('click',()=>{ if(f.id==='tapparelle'){ const next=!isOn; post(next?'alza_tutto':'abbassa_tutto',{}).then(()=>{ localStorage.setItem('ui_sh_up', next?'1':'0'); setTimeout(load,400); }); } else if(f.kind==='toggle'&&f.toggleEvent){ const next=!isOn; post(f.toggleEvent,{value:String(next)}).then(()=> setTimeout(load,300)); } else if(f.kind==='action'&&f.event){ post(f.event,{}); }
    }); fav.appendChild(tile); });
  const chips=$('#chips'); chips.innerHTML='';
  (m.people||[]).forEach(p=>{ const inout=!!(p.onlineSmart||p.onlineRaw); const d=document.createElement('div'); d.className='chip'; d.innerHTML=`<span class="dot ${inout?'in':'out'}"></span><span>${p.name||'—'}</span><span class="meta ${inout?'in':'out'}">${inout?'IN':'OUT'}</span>`; chips.appendChild(d); });
}
function load(){ get().then(render).catch(e=>console.error('model',e)); }
document.addEventListener('DOMContentLoaded',()=>{ load(); setInterval(load,60000); });
