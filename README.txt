# Casa — UI completa (GitHub Pages)

1) Carica tutti i file nella root del repo.
2) Settings → Pages → Deploy from a branch (main / root).
3) Hard Refresh.
4) Test in Console: fetch(window.CONFIG.DOGET_URL,{cache:'no-store'}).then(r=>r.status); JSONP.fetch(window.CONFIG.DOGET_URL).then(console.log).
