let LAST=null; const fmt=n=> (n==null||!isFinite(+n))?'—':(Math.round(+n));
function glyphClass(id){ if(id==='tapparelle') return 'g-shutter'; if(id==='piante') return 'g-plant'; if(id==='vacanza') return 'g-sun'; if(id==='override') return 'g-moon'; return 'g-generic'; }
function glyphIcon(id){ const m={tapparelle:'icons/glyph-shutter.svg',piante:'icons/glyph-plant.svg',vacanza:'icons/glyph-sun.svg',override:'icons/glyph-moon.svg'}; return m[id]||'icons/glyph-generic.svg'; }
function btnIcon(kind){ return (kind==='action')?'icons/ui-play.svg':'icons/ui-power.svg'; }

function render(m){ LAST=m; // state banner
 const sb=$('#stateBanner'); if(sb) sb.textContent=(m.state||'—').replace('_',' ');
 // weather
 if(m.weather){ $('#tempValue').textContent=fmt(m.weather.tempC); $('#humValue').textContent=fmt(m.weather.humidity); $('#weather-pill').textContent=fmt(m.weather.tempC)+'°C · '+fmt(m.weather.windKmh)+' km/h'; }
 // people
 const ppl=Array.isArray(m.people)?m.people:[]; $('#peopleOnline').textContent=ppl.filter(p=>p.onlineSmart||p.onlineRaw).length; $('#houseStatus').textContent=(m.notte?'Notte':'Giorno');
 // energy
 if(m.energy) $('#energyKwh').textContent = (m.energy.kwh==null?'—':m.energy.kwh);
 // favorites
 const grid=$('#favoritesGrid'); grid.innerHTML=''; (CONFIG.FAVORITES||[]).forEach(f=>{ let isOn=(f.kind==='toggle'&&f.stateKey&&!!m[f.stateKey]); if(f.id==='tapparelle'){ isOn=(localStorage.getItem('ui_sh_up')==='1'); }
 const el=document.createElement('article'); el.className='fav'+(isOn?' active':''); el.innerHTML=`<div class="glyph ${glyphClass(f.id)}"><img src="${glyphIcon(f.id)}" alt></div><div class="title">${f.label}</div><div class="subtitle">${f.subtitle||''}</div><div class="status">${f.kind==='toggle'?(isOn?'On':'Off'):''}${f.id==='tapparelle'?(isOn?' · Aperte':' · Chiuse'):''}</div><button class="btn" aria-label="toggle"><img src="${btnIcon(f.kind)}" alt></button>`; el.querySelector('.btn').addEventListener('click',(ev)=>{ ev.stopPropagation(); if(f.id==='tapparelle'){ const next=!isOn; callAdmin(next?'alza_tutto':'abbassa_tutto',{}).then(()=>{ localStorage.setItem('ui_sh_up', next?'1':'0'); setTimeout(load,500); }); return; } if(f.kind==='toggle' && f.toggleEvent){ const next=!isOn; callAdmin(f.toggleEvent,{value:String(next)}).then(()=> setTimeout(load,400)); } if(f.kind==='action' && f.event){ callAdmin(f.event,{}); }
 }); grid.appendChild(el); });
 // people chips using peopleLast
 const host=$('#peopleChips'); host.innerHTML=''; const idx={}; (m.people||[]).forEach(p=> idx[String(p.name||'').toLowerCase()]=p); (m.peopleLast||[]).forEach(x=>{ const k=String(x.name||'').toLowerCase(); (idx[k]=idx[k]||{name:x.name,onlineSmart:false}).lastInOut={event:x.lastEvent,day:x.lastDay,time:x.lastTime}; }); const arr=Object.values(idx); arr.forEach(p=>{ const isIn=!!p.onlineSmart || (p.lastInOut && p.lastInOut.event==='ARRIVO'); const st=isIn?'in':'out'; const time=(p.lastInOut&&(p.lastInOut.time&&p.lastInOut.day))?(p.lastInOut.time+' • '+p.lastInOut.day):'—'; const d=document.createElement('div'); d.className='chip'; d.innerHTML=`<span class="dot ${st}"></span><span>${p.name}</span><span class="meta ${st}">${st.toUpperCase()}</span><span class="time">${time}</span>`; host.appendChild(d); });
}
function load(){ fetchModel().then(render).catch(e=>{ console.error('Model error',e); }); }
document.addEventListener('DOMContentLoaded',()=>{ load(); setInterval(load,60000); });
