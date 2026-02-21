/**
 * Automazione Casa v4.6 — COMPLETO (patch PIN + VACANZA/OVERRIDE + Trend 24h)
 * Data: 2026-02-13
 * Note:
 *  - Include: presenza debounced, alba/tramonto con retry, trigger scheduler,
 *    API webhook persone, API admin (vacanza/override) con PIN server-side (Script Property ADMIN_PIN),
 *    doGet JSON/JSONP per dashboard + endpoint trend 24h, esport diagnostica, IFTTT bridge.
 *  - Adatta i nomi dei fogli: Config, Stato, Persone, Geo, Log (già usati nel tuo file).
 */

/***************************************
 * HELPER DIAGNOSTICA
 ***************************************/
function forcePresenceEffectiveTrueNow(){
  const sp = PropertiesService.getScriptProperties();
  sp.setProperty('presenceReported','true');
  sp.setProperty('presenceLastChangeMs', String(Date.now()));
  SpreadsheetApp.getActive().getSheetByName('Config').getRange('B6').setValue(true);
  Logger.log('Forzata PRESENZA_EFFETTIVA=TRUE.');
}
function __writeB6TrueTest(){
  const cfg = SpreadsheetApp.getActive().getSheetByName('Config');
  const before = cfg.getRange('B6').getValue();
  cfg.getRange('B6').setValue(true);
  Logger.log('B6 before='+before+' after='+cfg.getRange('B6').getValue());
}
function _debugPresenceNow(){
  try{
    const P = sh('Persone');
    const last = P.getLastRow();
    let rawIn = false;
    let vals = [];
    if (last >= 2){
      vals = P.getRange(2,6,last-1,1).getValues().flat().map(x => String(x).trim());
      rawIn = vals.some(x => x.toUpperCase()==='IN');
    }
    const cfg = sh('Config');
    const b5 = cfg.getRange('B5').getValue();
    const b6 = cfg.getRange('B6').getValue();

    const sp = PropertiesService.getScriptProperties();
    const repStr = sp.getProperty('presenceReported');
    const lastMs = Number(sp.getProperty('presenceLastChangeMs'));
    const nowMs = Date.now();
    const elapsedMin = isFinite(lastMs) ? Math.round((nowMs-lastMs)/60000) : 'NaN';

    Logger.log('--- DEBUG PRESENZA ---');
    Logger.log('Persone!F: '+JSON.stringify(vals));
    Logger.log('calcPresenceAuto = '+rawIn);
    Logger.log('B5 (PRESENZA_AUTO) = '+b5);
    Logger.log('B6 (PRESENZA_EFFETTIVA) = '+b6);
    Logger.log('presenceReported = '+repStr);
    Logger.log('presenceLastChangeMs = '+lastMs+' ('+elapsedMin+' min)');
  }catch(e){
    Logger.log('DEBUG ERROR: '+e);
  }
}
function resetPresenceDebounce(){
  const sp = PropertiesService.getScriptProperties();
  sp.deleteProperty('presenceReported');
  sp.deleteProperty('presenceLastChangeMs');
  Logger.log('Debounce RESET.');
}

/********************************************
 * CONFIG
 ********************************************/
const IFTTT_KEY = getProp_('IFTTT_KEY',''); // opzionale: imposta in Script Properties

const SHEET_CONFIG  = 'Config';
const SHEET_STATO   = 'Stato';
const SHEET_PERSONE = 'Persone';
const SHEET_GEO     = 'Geo';
const SHEET_LOG     = 'Log';

// Parametri PIANTE
const DELAY_ALBA_MIN = 10;
const DELAY_POST_CHIUSURA_MIN = 5;

// Presenza: debounce
const DEBOUNCE_IN_MIN  = 1;
const DEBOUNCE_OUT_MIN = 5;

// Presenza: hold LIFE (B)
const LIFE_HOLD_IN_MIN = 10;

// Grace diurno prima di chiudere (A)
const EMPTY_GRACE_MIN = 8;

// AUTO-OUT per assenza LIFE
const LIFE_TIMEOUT_MIN = 45;

// AUTO-IN quarantena
const AUTO_IN_MIN = 3;

// Retry alba/tramonto
const SUN_API_RETRY_MAX = 3;
const SUN_API_RETRY_SLEEP_MS = 5000;

/********************************************
 * BASIC UTIL
 ********************************************/
function ss(){ return SpreadsheetApp.getActive(); }
function sh(n){ return ss().getSheetByName(n) || ss().insertSheet(n); }
function v(n,a){ return sh(n).getRange(a).getValue(); }
function s(n,a,val){ sh(n).getRange(a).setValue(val); }
function appendRow(n,arr){ sh(n).appendRow(arr); }

function getProp_(k, def){
  try{ const v = PropertiesService.getScriptProperties().getProperty(k); return (v==null?def:v); }catch(e){ return def; }
}
function setProp_(k, val){
  try{ PropertiesService.getScriptProperties().setProperty(k,String(val)); }catch(e){}
}
function sleepMs_(ms){ Utilities.sleep(Math.max(0,ms|0)); }
function _setInternalWriteFlag_(on){ try{ PropertiesService.getScriptProperties().setProperty('internalWrite', on?'1':'0'); }catch(e){} }
function _isInternalWrite_(){ try{ return PropertiesService.getScriptProperties().getProperty('internalWrite')==='1'; }catch(e){ return false; } }

/********************************************
 * TEST WEBHOOK
 ********************************************/
function __testWebhook(){
  logEvent('TEST_WEBHOOK','Eseguito testWebhook → alza_tutto','');
  callIFTTT('alza_tutto',{test:true});
}

/********************************************
 * LOG
 ********************************************/
function logEvent(code,desc,note){
  const st = v(SHEET_CONFIG,'B1');
  appendRow(SHEET_LOG,[new Date(),st,code,desc,note||'']);
}

/********************************************
 * CALL IFTTT
 ********************************************/
function callIFTTT(eventName,payload){
  const key = IFTTT_KEY || getProp_('IFTTT_KEY','');
  if(!key){ logEvent('IFTTT_SKIP','IFTTT key non impostata',''); return; }
  const url = 'https://maker.ifttt.com/trigger/'+eventName+'/with/key/'+key;
  try{
    UrlFetchApp.fetch(url,{
      method:'post',
      contentType:'application/json',
      payload:JSON.stringify(payload||{}),
      muteHttpExceptions:true
    });
    logEvent('IFTTT_OK','Chiamato '+eventName,'');
  }catch(e){
    logEvent('IFTTT_ERR','Errore '+eventName,e);
  }
}

/********************************************
 * MENU
 ********************************************/
function onOpen(){
  SpreadsheetApp.getUi().createMenu('🔧 Diagnostica')
    .addItem('Esporta Diagnostica','exportDiagnosticsPDFForToday')
    .addSeparator()
    .addItem('Valuta stato adesso','evaluateStateNow')
    .addItem('Reinstalla Trigger','installAutomation')
    .addItem('Test Webhook','__testWebhook')
    .addToUi();
}

/********************************************
 * EXPORT PDF
 ********************************************/
function exportDiagnosticsPDFForToday(){
  const tmpName='Diagnostica_Oggi';
  const tmp = sh(tmpName);
  try{
    tmp.clear();
    const logSh = sh(SHEET_LOG);
    const last = logSh.getLastRow();
    if(last<2) throw new Error('Log vuoto');

    const data = logSh.getRange(1,1,last,5).getValues();
    const header = data[0];

    const tzId = Session.getScriptTimeZone() || 'Europe/Rome';
    const today = Utilities.formatDate(new Date(),tzId,'yyyy-MM-dd');

    const out=[header];
    for(let i=1;i<data.length;i++){
      const d=data[i];
      const day = Utilities.formatDate(new Date(d[0]),tzId,'yyyy-MM-dd');
      if(day===today) out.push(d);
    }
    tmp.getRange(1,1,out.length,out[0].length).setValues(out);

    const url = 'https://docs.google.com/spreadsheets/d/'+ss().getId()+'/export?format=pdf&size=letter&portrait=true'
    + '&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=true&fzr=false'
    + '&gid='+tmp.getSheetId();


    const blob = UrlFetchApp.fetch(url,{headers:{'Authorization':'Bearer '+ScriptApp.getOAuthToken()}})
      .getBlob().setName('Diagnostica_'+today+'.pdf');

    const saved = DriveApp.createFile(blob);
    logEvent('EXPORT_PDF','Creato PDF',saved.getUrl());
    SpreadsheetApp.getUi().alert('PDF: '+saved.getUrl());
  }catch(e){
    logEvent('ERROR_EXPORT_PDF','Errore: '+e,'');
  }
}

/********************************************
 * SUN TIMES
 ********************************************/
function ALBA(lat,lon,tzid,off){ return _sunTime_(lat,lon,tzid,off,'sunrise'); }

function TRAMONTO(lat,lon,tzid,off){ return _sunTime_(lat,lon,tzid,off,'sunset'); }

function _sunTime_(lat,lon,tzid,offMinutes,kind){
  const zone = String(tzid||Session.getScriptTimeZone()||'Europe/Rome');
  const todayLocal = Utilities.formatDate(new Date(),zone,'yyyy-MM-dd');
  const url = 'https://api.sunrise-sunset.org/json?lat='+lat+'&lng='+lon+'&date='+todayLocal+'&formatted=0';
  const resp = UrlFetchApp.fetch(url,{muteHttpExceptions:true});
  const js = JSON.parse(resp.getContentText());
  if(!js || js.status!=='OK') throw new Error('API alba/tramonto non disponibile');
  const tUTC = new Date(js.results[kind]);
  const tLocal = new Date(Utilities.formatDate(tUTC,zone,"yyyy-MM-dd'T'HH:mm:ss"));
  return new Date(tLocal.getTime()+(offMinutes||0)*60000);
}


function _getSunTimesWithRetry_(){
  const g = sh(SHEET_GEO);
  const lat = g.getRange('A2').getValue();
  const lon = g.getRange('B2').getValue();
  const tzid = g.getRange('C2').getValue() || 'Europe/Rome';
  const off = Number(g.getRange('D2').getValue()||0);
  let lastErr=null;
  for(let i=1;i<=SUN_API_RETRY_MAX;i++){
    try{ return { alba: ALBA(lat,lon,tzid,off), tram: TRAMONTO(lat,lon,tzid,off) }; }
    catch(e){ lastErr=e; logEvent('SUN_API_RETRY','Tentativo '+i+' fallito: '+e,''); sleepMs_(SUN_API_RETRY_SLEEP_MS); }
  }
  throw lastErr;
}

/********************************************
 * TRIGGERS
 ********************************************/
function ensureTriggers(){
  ScriptApp.getProjectTriggers().forEach(t=>{
    const f = t.getHandlerFunction();
    if(f==='onSunrise' || f==='onSunset') ScriptApp.deleteTrigger(t);
  });
  if(!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='evaluateStateNow'))
    ScriptApp.newTrigger('evaluateStateNow').timeBased().everyMinutes(5).create();

  if(!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='runMidnightRoutine'))
    ScriptApp.newTrigger('runMidnightRoutine').timeBased().atHour(0).nearMinute(0).everyDays(1).create();

  if(!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='scheduleSunEventsForToday'))
    ScriptApp.newTrigger('scheduleSunEventsForToday').timeBased().atHour(0).nearMinute(5).everyDays(1).create();

  scheduleSunEventsForToday();
}
function scheduleSunEventsForToday(){
  try{
    const { alba, tram } = _getSunTimesWithRetry_();
    s(SHEET_STATO,'B3',alba);
    s(SHEET_STATO,'B4',tram);
    ScriptApp.newTrigger('onSunrise').timeBased().at(alba).create();
    ScriptApp.newTrigger('onSunset').timeBased().at(tram).create();
    logEvent('SCHEDULE_OK','Pianificati alba/tramonto','');
    updateTriggerDashboard();
  }catch(e){ logEvent('ERROR_SCHEDULE','Errore: '+e,''); }
}
function scheduleOnce_(handlerName,minutes){
  try{
    const ms = Math.max(60000,Math.floor((Number(minutes)||0)*60000));
    ScriptApp.newTrigger(handlerName).timeBased().after(ms).create();
    logEvent('SCHEDULE_ONCE','Pianificato '+handlerName+' tra '+minutes+' min','');
  }catch(e){ logEvent('ERROR_SCHEDULE_ONCE','Errore scheduleOnce_: '+e,''); }
}
function clearTriggers_(handlerName){
  ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction && t.getHandlerFunction()===handlerName) ScriptApp.deleteTrigger(t); });
}

/********************************************
 * PRESENZA
 ********************************************/
function calcPresenceAuto(){
  const p = sh(SHEET_PERSONE);
  const last = p.getLastRow();
  if(last<2) return false;
  const now = Date.now();
  const rows = p.getRange(2,1,last-1,6).getValues();
  return rows.some(r=>{
    const online = String(r[5]||'').trim().toUpperCase();
    const lastLifeDt = (r[4] instanceof Date) ? r[4].getTime() : null;
    const recentLife = lastLifeDt && (now-lastLifeDt)<=LIFE_HOLD_IN_MIN*60000;
    return (online==='IN') || recentLife;
  });
}
function applyPresenceDebounce_(rawEff){
  const now = Date.now();
  const prevStr = getProp_('presenceReported',null);
  const hasPrev = (prevStr!==null);
  const prev = (prevStr==='true');
  let lastChange = Number(getProp_('presenceLastChangeMs',String(now)));
  if(!isFinite(lastChange) || lastChange>now) lastChange=now;
  if(!hasPrev && rawEff===true){ setProp_('presenceReported','true'); setProp_('presenceLastChangeMs',String(now)); return { reported:true, lastChange:now }; }
  if(hasPrev && rawEff===prev){ return { reported:prev, lastChange }; }
  const mins = (rawEff ? DEBOUNCE_IN_MIN : DEBOUNCE_OUT_MIN);
  const mustWaitMs = mins*60000;
  const elapsed = now-lastChange;
  if(elapsed>=mustWaitMs){ setProp_('presenceReported',String(rawEff)); setProp_('presenceLastChangeMs',String(now)); return { reported:rawEff, lastChange:now }; }
  return { reported:(hasPrev?prev:false), lastChange };
}
function autoOutByLifeTimeout_(){
  try{
    const P = sh(SHEET_PERSONE);
    const last = P.getLastRow();
    if(last<2) return;
    const now = Date.now();
    const rows = P.getRange(2,1,last-1,6).getValues();
    for(let i=0;i<rows.length;i++){
      const persona = String(rows[i][0]||'').toLowerCase();
      const lastLife = rows[i][2];
      const online = String(rows[i][5]||'').toUpperCase();
      if(online!=='IN') continue;
      if(!(lastLife instanceof Date)) continue;
      const diff = (now-lastLife.getTime())/60000;
      if(diff>=LIFE_TIMEOUT_MIN){
        const row = 2+i;
        _setInternalWriteFlag_(true);
        try{
          P.getRange(row,4).setValue('AUTO_OUT');
          P.getRange(row,5).setValue(new Date());
          P.getRange(row,6).setValue('OUT');
        }finally{ _setInternalWriteFlag_(false); }
        logEvent('AUTO_OUT','Assenza LIFE per '+persona,'');
      }
    }
  }catch(e){ logEvent('ERROR_AUTO_OUT','Errore: '+e,''); }
}
function autoInOnLifeRecovery_(personaLower){
  try{
    const P = sh(SHEET_PERSONE);
    const last = P.getLastRow();
    if(last<2) return;
    const names = P.getRange(2,1,last-1,1).getValues().map(r=>String(r[0]||'').toLowerCase());
    const idx = names.indexOf(personaLower);
    if(idx<0) return;
    const row = 2+idx;
    const online = String(P.getRange(row,6).getValue()||'').toUpperCase();
    if(online==='IN') return;
    const lastEvt = String(P.getRange(row,4).getValue()||'').toUpperCase();
    const lastLifeDt = P.getRange(row,5).getValue();
    const lastLifeRaw = P.getRange(row,3).getValue();
    const now = Date.now();
    let ok = true;
    if(lastEvt==='USCITA' || lastEvt==='AUTO_OUT'){
      const ref = (lastLifeDt instanceof Date ? lastLifeDt.getTime() : (lastLifeRaw instanceof Date ? lastLifeRaw.getTime() : now));
      const diff = (now-ref)/60000; if(diff<AUTO_IN_MIN) ok=false;
    }
    if(ok){
      _setInternalWriteFlag_(true);
      try{
        P.getRange(row,4).setValue('ARRIVO');
        P.getRange(row,5).setValue(new Date());
        P.getRange(row,6).setValue('IN');
      }finally{ _setInternalWriteFlag_(false); }
      logEvent('AUTO_IN','LIFE recovery → IN: '+personaLower,'');
    }else{ logEvent('AUTO_IN_SKIP','LIFE ignorato (quarantena): '+personaLower,''); }
  }catch(e){ logEvent('ERROR_AUTO_IN','Errore: '+e,''); }
}

/********************************************
 * NOTTE?
 ********************************************/
function _fallbackIsNight_(){ const h=new Date().getHours(); return (h>=22 || h<6); }
function isNight(){
  try{
    const now=new Date();
    const tram=v(SHEET_STATO,'B4');
    const alba=v(SHEET_STATO,'B3');
    if(tram instanceof Date && alba instanceof Date){ return (now>=tram) || (now<alba); }
  }catch(e){}
  return _fallbackIsNight_();
}

/********************************************
 * setState (FIX ReferenceError)
 ********************************************/
function setState(val){
  try{ s(SHEET_CONFIG,'B1', val); }
  catch(e){ logEvent('ERROR_SETSTATE','Errore in setState: '+e,''); }
}

/********************************************
 * VALUTAZIONE STATO (CUORE)
 ********************************************/
function evaluateStateNow(){
  try{
    autoOutByLifeTimeout_();

    const stato = v(SHEET_CONFIG,'B1');
    const presenzaMan = !!v(SHEET_CONFIG,'B2');
    const vacanza     = !!v(SHEET_CONFIG,'B3');
    const override    = !!v(SHEET_CONFIG,'B4');

    const presenzaAutoRaw = calcPresenceAuto();
    s(SHEET_CONFIG,'B5',presenzaAutoRaw);

    const rawEff = !!(presenzaMan || presenzaAutoRaw);
    const deb = applyPresenceDebounce_(rawEff);
    const presenzaEff = !!deb.reported;
    s(SHEET_CONFIG,'B6',presenzaEff);

    const prevPresence = !!v(SHEET_CONFIG,'B8');

    const notte = isNight();
    const giorno = !notte;

    logEvent('EVAL_TRACE','now='+new Date().toISOString()+' notte='+notte+' presenzaAutoRaw='+presenzaAutoRaw+' presenzaEff='+presenzaEff+' stato='+stato,'');

    // ARRIVO
    if(presenzaEff && !prevPresence){ logEvent('ARRIVO','Arrivo → APRI TUTTO',''); callIFTTT('alza_tutto'); }

    // Decisione stato
    let desired = stato;
    if(override){ desired = 'SECURITY_NIGHT'; }
    else if(vacanza){ desired = notte? 'SECURITY_NIGHT':'SECURITY_DAY'; }
    else{
      desired = presenzaEff ? (notte?'COMFY_NIGHT':'COMFY_DAY') : (notte?'SECURITY_NIGHT':'SECURITY_DAY');
    }

    const becameSecurityDay = (desired==='SECURITY_DAY' && stato!=='SECURITY_DAY');

    if(desired!==stato){
      setState(desired);
      logEvent('STATE_DECISION','Transizione → '+desired,'');
      if(desired==='SECURITY_NIGHT') applySecurityNight();
      if(desired==='SECURITY_DAY')   applySecurityDay();
      if(desired==='COMFY_DAY')       applyComfyDay();
      if(desired==='COMFY_NIGHT')     applyComfyNight();
    }else{
      logEvent('STATE_DECISION','Nessuna transizione','');
    }

    // CASA VUOTA DI GIORNO — MA CON GRACE
    if(!presenzaEff && giorno && becameSecurityDay){
      logEvent('DAY_SEC_PENDING','Casa forse vuota → verifica tra '+EMPTY_GRACE_MIN+' min','');
      const planned = new Date(Date.now()+EMPTY_GRACE_MIN*60000);
      s(SHEET_STATO,'B11',planned);
      scheduleOnce_('verifyHouseEmptyThenClose',EMPTY_GRACE_MIN);
    }

    s(SHEET_CONFIG,'B8',presenzaEff);
    s(SHEET_STATO,'B5',new Date());
    s(SHEET_STATO,'B6',notte?'NOTTE':'GIORNO');
  }catch(e){ logEvent('ERROR_EVAL','Errore: '+e,''); }
  finally{ updateTriggerDashboard(); }
}

/********************************************
 * VERIFICA → CHIUSURA (GRACE)
 ********************************************/
function verifyHouseEmptyThenClose(){
  try{
    const presenzaMan = !!v(SHEET_CONFIG,'B2');
    const debRep = (getProp_('presenceReported','false')==='true');
    const presenzaEff = !!(presenzaMan || debRep);
    const giorno = !isNight();

    if(!presenzaEff && giorno){
      logEvent('DAY_SEC_CLOSE','Confermata casa vuota → CHIUDI TUTTO','');
      callIFTTT('abbassa_tutto');
      scheduleOnce_('runPianteAfterCloseOnce',DELAY_POST_CHIUSURA_MIN);
    }else{ logEvent('DAY_SEC_CANCEL','Ripristino presenza → niente chiusura',''); }
  }catch(e){ logEvent('ERROR_DAY_SEC_VERIFY','Errore: '+e,''); }
  finally{
    clearTriggers_('verifyHouseEmptyThenClose');
    s(SHEET_STATO,'B11','—');
    updateTriggerDashboard();
  }
}

/********************************************
 * EVENTI TEMPO
 ********************************************/
function onSunrise(){
  try{
    // niente alza_tutto all’alba
    const planned = new Date(Date.now()+DELAY_ALBA_MIN*60000);
    s(SHEET_STATO,'B10',planned);
    scheduleOnce_('runPianteOnce',DELAY_ALBA_MIN);
    logEvent('ALBA','ALBA → PIANTE +'+DELAY_ALBA_MIN,'');
  }catch(e){ logEvent('ERROR_SUNRISE','Errore: '+e,''); }
  evaluateStateNow();
}
function runPianteOnce(){
  try{ callIFTTT('piante'); logEvent('PIANTE_OK','PIANTE (post ALBA)',''); }
  catch(e){ logEvent('PIANTE_ERR','Errore: '+e,''); }
  finally{ clearTriggers_('runPianteOnce'); s(SHEET_STATO,'B10','—'); updateTriggerDashboard(); }
}
function runPianteAfterCloseOnce(){
  try{ callIFTTT('piante'); logEvent('PIANTE_OK','PIANTE (post chiusura diurna)',''); }
  catch(e){ logEvent('PIANTE_ERR','Errore: '+e,''); }
  finally{ clearTriggers_('runPianteAfterCloseOnce'); s(SHEET_STATO,'B11','—'); updateTriggerDashboard(); }
}
function onSunset(){
  try{
    const presenzaMan = !!v(SHEET_CONFIG,'B2');
    const debRep = (getProp_('presenceReported','false')==='true');
    const presenzaEff = !!(presenzaMan || debRep);
    if(!presenzaEff){ logEvent('TRAMONTO','Casa vuota → CHIUDI TUTTO',''); callIFTTT('abbassa_tutto'); }
    else{ logEvent('TRAMONTO','Casa occupata → nessuna tapparella',''); }
  }catch(e){ logEvent('ERROR_SUNSET','Errore: '+e,''); }
  evaluateStateNow();
}
function runMidnightRoutine(){ logEvent('MEZZANOTTE','Routine mezzanotte',''); evaluateStateNow(); }

/********************************************
 * PROFILI AZIONE
 ********************************************/
function applySecurityNight(){ callIFTTT('ezviz_interne_on'); callIFTTT('ezviz_esterne_on'); }
function applySecurityDay(){  callIFTTT('ezviz_interne_on'); callIFTTT('ezviz_esterne_on'); }
function applyComfyDay(){     callIFTTT('ezviz_interne_off'); callIFTTT('ezviz_esterne_off'); }
function applyComfyNight(){   callIFTTT('ezviz_interne_off'); callIFTTT('ezviz_esterne_on'); }

/********************************************
 * LATE CLOSE
 ********************************************/
function _presenceEffective_(){
  try{ const presenzaMan = !!v(SHEET_CONFIG,'B2'); const deb = (getProp_('presenceReported','false')==='true'); return (presenzaMan || deb); }
  catch(e){ return false; }
}
function lateCloseTimeFor_(baseDate){
  const base=new Date(baseDate.getFullYear(),baseDate.getMonth(),baseDate.getDate());
  const hour=isHolidayOrWeekend_(base)?1:23;
  const when=new Date(base.getFullYear(),base.getMonth(),baseDate.getDate(),hour,0,0);
  if(when.getTime()<=Date.now()){
    const tomorrow=new Date(base.getFullYear(),base.getMonth(),baseDate.getDate()+1);
    return lateCloseTimeFor_(tomorrow);
  }
  return when;
}
function runLateClose(){
  try{
    const chiudi = !!v(SHEET_CONFIG,'B7'); // CHIUDI_NOTTE_OCCUPATA
    if(!chiudi && _presenceEffective_()){ logEvent('LATE_CLOSE_SKIP','Flag OFF + casa occupata → salto chiusura',''); return; }
    logEvent('LATE_CLOSE','Chiusura notturna',''); callIFTTT('abbassa_tutto');
  }catch(e){ logEvent('ERROR_LATE_CLOSE','Errore: '+e,''); }
}
function scheduleLateCloseForToday(){
  try{
    ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction()==='runLateClose') ScriptApp.deleteTrigger(t); });
    const when=lateCloseTimeFor_(new Date());
    ScriptApp.newTrigger('runLateClose').timeBased().at(when).create();
    s(SHEET_STATO,'B12',when);
    logEvent('SCHEDULE_LATECLOSE','Chiusura extra alle '+when,'');
  }catch(e){ logEvent('ERROR_SCHEDULE_LATECLOSE','Errore: '+e,''); }
}
function installAutomation(){
  ScriptApp.getProjectTriggers().forEach(t=>{ ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('evaluateStateNow').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('runMidnightRoutine').timeBased().atHour(0).nearMinute(0).everyDays(1).create();
  ScriptApp.newTrigger('scheduleSunEventsForToday').timeBased().atHour(0).nearMinute(5).everyDays(1).create();
  ScriptApp.newTrigger('scheduleLateCloseForToday').timeBased().atHour(0).nearMinute(10).everyDays(1).create();
  scheduleLateCloseForToday();
  scheduleSunEventsForToday();
  updateTriggerDashboard();
}

/********************************************
 * FESTIVI
 ********************************************/
function _easterDate_(y){ var a=y%19,b=Math.floor(y/100),c=y%100,d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25), g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30,i=Math.floor(c/4),k=c%4, l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451), month=Math.floor((h+l-7*m+114)/31), day=((h+l-7*m+114)%31)+1; return new Date(y,month-1,day);} 
function _sameYMD_(d1,d2){ return d1.getFullYear()===d2.getFullYear() && d1.getMonth()===d2.getMonth() && d1.getDate()===d2.getDate(); }
function _addDays_(d,n){ var x=new Date(d.getTime()); x.setDate(x.getDate()+n); return x; }
function _isItalianHoliday_(d){ var y=d.getFullYear(); var fixed=[[0,1],[0,6],[3,25],[4,1],[5,2],[7,15],[10,1],[11,8],[11,25],[11,26]]; for(var i=0;i<fixed.length;i++){ if(_sameYMD_(d,new Date(y,fixed[i][0],fixed[i][1]))) return true; } var easter=_easterDate_(y); if(_sameYMD_(d,_addDays_(easter,1))) return true; return false; }
function isHolidayOrWeekend_(d){ var dw=d.getDay(); if(dw===0||dw===6) return true; if(_isItalianHoliday_(d)) return true; try{ var shF=sh('Festivi'); var last=shF.getLastRow(); if(last>=1){ var vals=shF.getRange(1,1,last,1).getValues(); for(var i=0;i<vals.length;i++){ var v=vals[i][0]; if(v){ if(v instanceof Date){ if(_sameYMD_(d,v)) return true; }else{ var s=String(v).trim(); var m=s.match(/^(\d{4})-(\d{2})-(\d{2})$/); if(m){ if(_sameYMD_(d,new Date(+m[1],+m[2]-1,+m[3]))) return true; } } } } } }catch(e){} return false; }

/********************************************
 * DASHBOARD TRIGGER
 ********************************************/
function updateTriggerDashboard(){
  try{ const triggers = ScriptApp.getProjectTriggers(); s(SHEET_STATO,'B13',(triggers?triggers.length:0)+' attivi'); s(SHEET_STATO,'B14',new Date()); }
  catch(e){ logEvent('ERROR_DASHBOARD','Errore: '+e,''); }
}

/********************************************
 * WEBHOOK (POST) — persone + admin (PIN)
 ********************************************/
function doPost(e){
  try{
    let body={};
    if(e && e.postData && e.postData.contents){
      try{ body=JSON.parse(e.postData.contents); }
      catch(err){ body={ event:e.parameter.event, persona:e.parameter.persona, value:e.parameter.value, pin:e.parameter.pin }; }
    }
    const evt = String(body.event||'').toLowerCase();
    const persona = String(body.persona||'').toLowerCase();
    const note = body.note||'';
    const value = (body.value===true || String(body.value).toLowerCase()==='true');
    const pin = String(body.pin||'');

       // Admin senza PIN (NO PIN)
const isAdminEvt = (evt==='set_vacanza' || evt==='set_override');
if (isAdminEvt){
  if (evt==='set_vacanza'){
    s(SHEET_CONFIG,'B3', value===true);
    logEvent('API_VACANZA','Vacanza→'+value,'(no pin)');
    evaluateStateNow();
    return ContentService.createTextOutput('OK');
  }
  if (evt==='set_override'){
    s(SHEET_CONFIG,'B4', value===true);
    logEvent('API_OVERRIDE','Override→'+value,'(no pin)');
    evaluateStateNow();
    return ContentService.createTextOutput('OK');
  }
}


    // Persone
    if(!evt){ logEvent('WEBHOOK_BAD','Evento mancante',JSON.stringify(body)); return ContentService.createTextOutput('KO'); }

    if(evt==='persona_arriva'){
      if(!persona) { logEvent('WEBHOOK_BAD','Persona mancante',JSON.stringify(body)); return ContentService.createTextOutput('KO'); }
      updatePersonStatus_(persona,'ARRIVO');
      logEvent('WEBHOOK_IN','ARRIVO: '+persona,note);

    }else if(evt==='persona_esce'){
      if(!persona) { logEvent('WEBHOOK_BAD','Persona mancante',JSON.stringify(body)); return ContentService.createTextOutput('KO'); }
      updatePersonStatus_(persona,'USCITA');
      logEvent('WEBHOOK_IN','USCITA: '+persona,note);

    }else if(evt==='persona_life'){
      if(!persona) { logEvent('WEBHOOK_BAD','Persona mancante',JSON.stringify(body)); return ContentService.createTextOutput('KO'); }
      updatePersonStatus_(persona,'LIFE');
      autoInOnLifeRecovery_(persona);
      logEvent('WEBHOOK_IN','LIFE: '+persona,note);

    }else{
      logEvent('WEBHOOK_UNKNOWN','Evento sconosciuto: '+evt,note);
    }

    evaluateStateNow();
    return ContentService.createTextOutput('OK');
  }catch(err){ logEvent('WEBHOOK_ERR','Errore: '+err,''); return ContentService.createTextOutput('ERR'); }
}

/********************************************
 * PERSO
 ********************************************/
function updatePersonStatus_(personaLower,kind){
  try{
    const P = sh(SHEET_PERSONE);
    const last = P.getLastRow();
    if(last<2) return;
    const names = P.getRange(2,1,last-1,1).getValues().map(r=>String(r[0]||'').toLowerCase());
    const idx = names.indexOf(personaLower);
    if(idx<0) return;
    const row = 2+idx;
    const now = new Date();
    _setInternalWriteFlag_(true);
    try{
      if(kind==='ARRIVO'){
        P.getRange(row,4).setValue('ARRIVO');
        P.getRange(row,5).setValue(now);
        P.getRange(row,6).setValue('IN');
        P.getRange(row,3).setValue(now);
      }else if(kind==='USCITA'){
        P.getRange(row,4).setValue('USCITA');
        P.getRange(row,5).setValue(now);
        P.getRange(row,6).setValue('OUT');
        P.getRange(row,3).setValue(now);
      }else if(kind==='LIFE'){
        P.getRange(row,4).setValue('LIFE');
        P.getRange(row,5).setValue(now);
        P.getRange(row,3).setValue(now);
      }
    }finally{ _setInternalWriteFlag_(false); }
  }catch(e){ logEvent('ERROR_UPDATE_PERSON','Errore: '+e,''); }
}

/********************************************
 * doGet — JSON/JSONP + trend 24h
 ********************************************/
function _out(cb, obj){
  var text = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb+'('+text+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function buildTrend24h_(){
  var tz = Session.getScriptTimeZone() || 'Europe/Rome';
  var now = new Date();
  var start = new Date(now.getTime() - 24*60*60000);
  var eff = !!v('Config','B6');
  var out = [];
  for (var i=0;i<288;i++){
    var t = new Date(start.getTime() + i*5*60000);
    out.push({ t: Utilities.formatDate(t, tz, "yyyy-MM-dd'T'HH:mm:ss"), present: eff?1:0 });
  }
  return out;
}


function aggiornaAlbaTramonto() {
  try {
    const geo = sh('Geo');
    const alba = v('Stato','B3');     // già calcolato da scheduler
    const tram = v('Stato','B4');

    if (alba instanceof Date) s('Stato','B3', alba);
    if (tram instanceof Date) s('Stato','B4', tram);

    logEvent('SUN_CACHE','Aggiornati alba/tramonto (cached)','');
  } catch(err){
    logEvent('SUN_API_FAIL','Impossibile aggiornare alba/tramonto: '+err,'');
  }
}

/** =========================
 *  METEO — Open‑Meteo (URL con & reali)
 *  ========================= */

/** Ritorna { tempC:Number, icon:String, provider:String } */
function fetchOpenMeteo_(lat, lon, tz) {
  var url = 'https://api.open-meteo.com/v1/forecast'
    + '?latitude=' + encodeURIComponent(lat)
    + '&longitude=' + encodeURIComponent(lon)
    + '&current=temperature_2m,weather_code'
    + '&timezone=' + encodeURIComponent(tz || (Session.getScriptTimeZone() || 'Europe/Rome'));

  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, timeout: 15000 });
  if (resp.getResponseCode() !== 200) throw new Error('Open‑Meteo HTTP ' + resp.getResponseCode());

  var js = JSON.parse(resp.getContentText());
  var t  = (js && js.current && typeof js.current.temperature_2m === 'number') ? js.current.temperature_2m : null;
  var wc = (js && js.current && js.current.weather_code != null) ? String(js.current.weather_code) : '';

  if (t == null) throw new Error('Open‑Meteo: temperatura assente');
  return { tempC: t, icon: wc, provider: 'Open‑Meteo' };
}

/** Cache su CRUSCOTTO */
function updateWeatherCache_() {
  var geo = sh('Geo');
  var lat = geo.getRange('A2').getValue();
  var lon = geo.getRange('B2').getValue();
  var tz  = geo.getRange('C2').getValue() || (Session.getScriptTimeZone() || 'Europe/Rome');

  if (!(typeof lat === 'number' && typeof lon === 'number'))
    throw new Error('Geo!A2/B2 non valorizzati');

  var data = fetchOpenMeteo_(lat, lon, tz);
  var c = sh('CRUSCOTTO');

  function setKV_(k, v){
    var last = Math.max(1, c.getLastRow());
    var rows = c.getRange(1,1,last,2).getValues();
    for (var i=0;i<rows.length;i++){
      if (String(rows[i][0]).trim().toUpperCase() === k.toUpperCase()){
        c.getRange(i+1,2).setValue(v);
        return;
      }
    }
    c.getRange(last+1,1,1,2).setValues([[k,v]]);
  }

  setKV_('WEATHER_TEMP_C', data.tempC);
  setKV_('WEATHER_ICON',   data.icon);
  setKV_('WEATHER_PROVIDER','Open‑Meteo');
  setKV_('WEATHER_TS',     new Date());

  logEvent('WEATHER_OK','Aggiornato da Open‑Meteo','');
  return data;
}

/** Lettura cache + eventuale refresh */
function getWeather_(maxAgeMin){
  maxAgeMin = Math.max(1, maxAgeMin || 10);

  var c = sh('CRUSCOTTO');
  var last = Math.max(1, c.getLastRow());
  var kv = c.getRange(1,1,last,2).getValues();

  function getKV_(k){
    for (var i=0;i<kv.length;i++){
      if (String(kv[i][0]).trim().toUpperCase() === k.toUpperCase())
        return kv[i][1];
    }
    return null;
  }

  var ts = getKV_('WEATHER_TS');
  var fresh = ts instanceof Date && (Date.now() - ts.getTime()) <= maxAgeMin*60000;

  if (!fresh){
    try { return updateWeatherCache_(); }
    catch(_){ /* fallback cache */ }
  }

  var t  = getKV_('WEATHER_TEMP_C');
  if (t == null) t = v('Stato','B15'); // fallback manuale da foglio
  var ic = getKV_('WEATHER_ICON') || '';
  var pr = getKV_('WEATHER_PROVIDER') || 'Open‑Meteo';

  var num;
  if (typeof t === 'number') num = t;
  else {
    var s = String(t||'').replace(/[^\d,.\-]/g,'').replace(',', '.').trim();
    var n = Number(s);
    num = isFinite(n) ? n : null;
  }

  return { tempC: num, icon: String(ic), provider: String(pr) };
}

/** WMO → icona/testo */
function mapOpenMeteoCode_(wc){
  var m = {
    '0':{icon:'☀️',text:'Sereno'}, '1':{icon:'🌤️',text:'Poco nuv.'}, '2':{icon:'⛅',text:'Parz. nuv.'}, '3':{icon:'☁️',text:'Nuvoloso'},
    '45':{icon:'🌫️',text:'Nebbia'}, '48':{icon:'🌫️',text:'Nebbia ghiacciata'},
    '51':{icon:'🌦️',text:'Pioviggine'}, '53':{icon:'🌦️',text:'Pioviggine'}, '55':{icon:'🌦️',text:'Pioviggine'},
    '61':{icon:'🌧️',text:'Pioggia'}, '63':{icon:'🌧️',text:'Pioggia'}, '65':{icon:'🌧️',text:'Pioggia forte'},
    '80':{icon:'🌦️',text:'Rovesci'}, '81':{icon:'🌧️',text:'Rovesci'}, '82':{icon:'⛈️',text:'Temporali'},
    '95':{icon:'⛈️',text:'Temporale'}, '96':{icon:'⛈️',text:'Grandine'}, '99':{icon:'⛈️',text:'Grandine forte'}
  };
  return m[String(wc)] || {icon:'',text:''};
}

/** ============================================================
 *                       doGet — JSON / JSONP
 * ============================================================ */
function doGet(e){

  // ===== TEST METEO DIRETTO =====
  if (e && e.parameter && e.parameter.test === 'weather'){
    try{
      var w = getWeather_(10);
      return ContentService.createTextOutput(
        JSON.stringify({ok:true, weather:w})
      ).setMimeType(ContentService.MimeType.JSON);
    }catch(err){
      return ContentService.createTextOutput(
        JSON.stringify({ok:false, error:String(err)})
      ).setMimeType(ContentService.MimeType.JSON);
    }
  }

  const cb  = (e && e.parameter && e.parameter.callback) || '';
  const ss_ = SpreadsheetApp.getActive();

  /* ------------------------------
   *  ADMIN COMMANDS (NO PIN)
   * ------------------------------ */
  if (e && e.parameter && e.parameter.admin === '1'){
    try{
      const evt = String(e.parameter.event||'').toLowerCase();
      const asBool = v => String(v||'').toLowerCase()==='true';

      if (evt === 'piante'){ callIFTTT('piante'); return _out(cb,{ok:true}); }
      if (evt === 'set_vacanza'){ s('Config','B3',asBool(e.parameter.value)); evaluateStateNow(); return _out(cb,{ok:true}); }
      if (evt === 'set_override'){ s('Config','B4',asBool(e.parameter.value)); evaluateStateNow(); return _out(cb,{ok:true}); }
      if (evt === 'alza_tutto'){ callIFTTT('alza_tutto'); return _out(cb,{ok:true}); }
      if (evt === 'abbassa_tutto'){ callIFTTT('abbassa_tutto'); return _out(cb,{ok:true}); }

      // (se vuoi rimettere i comandi VIMAR, li aggiungiamo dopo)
      return _out(cb,{ok:false,error:'unknown_event'});
    }catch(err){
      logEvent('ADMIN_ERR','admin: '+err,'');
      return _out(cb,{ok:false,error:String(err)});
    }
  }

  /* ------------------------------
   *  TREND 24h
   * ------------------------------ */
  if (e && e.parameter && e.parameter.trend === '24h'){
    return _out(cb,{trend24h: buildTrend24h_()});
  }

  /* ------------------------------
   *  LOGS
   * ------------------------------ */
  if (e && e.parameter && e.parameter.logs === '1'){
    try{
      const shL = ss_.getSheetByName('Log');
      const last = shL ? shL.getLastRow() : 0;
      const take = Math.min(200, Math.max(0, last-1));
      const vals = (take>0) ? shL.getRange(last - take + 1, 1, take, 5).getValues() : [];
      const out = vals.map(r => ({ ts:r[0], stato:r[1], code:r[2], desc:r[3], note:r[4] })).reverse();
      return _out(cb,{ logs: out });
    }catch(err){
      logEvent('LOGS_ERR','doGet logs: '+err,'');
      return _out(cb,{ logs: [] });
    }
  }

  /* ------------------------------
   *  MODELLO UI COMPLETO
   * ------------------------------ */
  const model = {};

  // Stato base
  model.state              = v('Config','B1');
  model.vacanza            = !!v('Config','B3');
  model.override           = !!v('Config','B4');
  model.presenzaAuto       = v('Config','B5');
  model.presenzaEffettiva  = !!v('Config','B6');
  model.notte              = v('Stato','B6') === 'NOTTE';

  // Persone
  model.people = [];
  const shP = ss_.getSheetByName('Persone');
  if (shP){
    const last = Math.max(2, shP.getLastRow());
    if (last > 2){
      const rows = shP.getRange(2,1,last-1,6).getValues();
      const nowMs = Date.now();
      rows.forEach(r=>{
        const nameRaw = String(r[0]||'').trim(); if(!nameRaw) return;
        const raw = String(r[5]||'').trim().toUpperCase();
        const onlineRaw = ['IN','ONLINE','TRUE','1','SI','SÌ'].includes(raw);
        const lifeDt = (r[4] instanceof Date) ? r[4].getTime() : null;
        const recentLife = lifeDt && (nowMs - lifeDt) <= 600000; // 10'
        const display = nameRaw.charAt(0).toUpperCase()+nameRaw.slice(1).toLowerCase();
        model.people.push({ name:display, onlineRaw, onlineSmart:(onlineRaw||recentLife), lastLifeMinAgo: lifeDt? Math.round((nowMs-lifeDt)/60000): null });
      });
    }
  }

  // Meteo
  try{
    var w = getWeather_(10);
    var mapped = mapOpenMeteoCode_(w.icon);
    model.weather = { tempC:w.tempC, icon:w.icon, iconEmoji:mapped.icon, text:mapped.text, provider:w.provider };
  }catch(err){
    var t = v('Stato','B15');
    var n = (typeof t === 'number') ? t : null;
    model.weather = { tempC:n, icon:'', provider:'N/A' };
    logEvent('WEATHER_FALLBACK','fallback '+err,'');
  }

  // Energy (se usi queste celle)
  (function(){
    function __num(x){ var n=Number(x); return isFinite(n)?n:null; }
    model.energy = { kwh: __num(v('Stato','B16')) };
    var off = __num(v('Stato','B17')); model.devicesOfflineCount = off||0;
  })();

  // Meta / Next
  model.meta = { tz: Session.getScriptTimeZone() || 'Europe/Rome', nowIso: new Date().toISOString(), version: 'automazione-4.6' };
  model.next = { pianteAlba: v('Stato','B10') || '—', piantePostClose: v('Stato','B11') || '—', lateClose: v('Stato','B12') || '—' };

  // Alert badge (errori recenti)
  try{
    const shL = ss_.getSheetByName('Log');
    const last = shL ? shL.getLastRow() : 0;
    const take = Math.min(200, Math.max(0, last-1));
    const vals = (take>0) ? shL.getRange(last - take + 1, 1, take, 5).getValues() : [];
    const errCount = vals.reduce((acc,r)=>{
      const c = String(r[2]||'').toUpperCase();
      const d = String(r[3]||'').toUpperCase();
      return acc + ((/ERR|ERROR|FAIL/.test(c)||/ERR|ERROR|FAIL/.test(d))?1:0);
    },0);
    model.alerts = { logErrors: errCount };
  }catch(_){ model.alerts = { logErrors: 0 }; }

  return _out(cb, model);
}
