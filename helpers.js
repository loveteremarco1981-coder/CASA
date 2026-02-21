const $=(q,r)=> (r||document).querySelector(q); const $$=(q,r)=> Array.from((r||document).querySelectorAll(q));
function fetchModel(){ const u=CONFIG.DOGET_URL; return fetch(u,{cache:'no-store'}).then(r=>r.json()); }
function callAdmin(event,payload){ const u=CONFIG.DOGET_URL; return fetch(u,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(Object.assign({event},payload||{}))}).then(r=>r.text()); }
