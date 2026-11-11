/**
 * TOURisME — приём заявок, статусы менеджера, события в Meta (CAPI).
 * Вставить в Google Apps Script, привязанный к таблице. Инструкция в SETUP.md.
 *
 * Script Properties (Проект → Настройки → Свойства скрипта):
 *   MANAGER_KEY  — пароль менеджера (любой, например 6–8 символов)
 *   PIXEL_ID     — ID пикселя Meta
 *   CAPI_TOKEN   — токен Conversions API из Events Manager
 *   TEST_CODE    — (необязательно) код тестовых событий, пока проверяете
 *   SITE_URL     — (необязательно) адрес сайта, для события Purchase/Qualified
 */

var SHEET = 'Leads';
var COLS = ['id','created','lang','name','phone','dests','vibe','when','nights','adults','kids','ages','budget','passport',
  'tg_click','status','amount','note','fbp','fbc','event_id','updated','source'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(COLS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold');
  }
  return sh;
}
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sha_(s) {
  s = String(s || '').trim().toLowerCase(); if (!s) return '';
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function digits_(s) { return String(s || '').replace(/\D/g, ''); }

function rows_() {
  var sh = sheet_(), n = sh.getLastRow();
  if (n < 2) return [];
  var vals = sh.getRange(2, 1, n - 1, COLS.length).getValues();
  return vals.map(function (r, i) {
    var o = { _row: i + 2 };
    COLS.forEach(function (c, j) { o[c] = r[j]; });
    return o;
  });
}
function findRow_(id) {
  var sh = sheet_(), n = sh.getLastRow();
  if (n < 2) return 0;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (ids[i][0] === id) return i + 2;
  return 0;
}
function setCell_(row, col, val) { sheet_().getRange(row, COLS.indexOf(col) + 1).setValue(val); }
function getRowObj_(row) {
  var v = sheet_().getRange(row, 1, 1, COLS.length).getValues()[0], o = {};
  COLS.forEach(function (c, j) { o[c] = v[j]; });
  return o;
}

/* ---------- Meta Conversions API ---------- */
function capi_(name, eventId, lead, custom) {
  var pixel = prop_('PIXEL_ID'), token = prop_('CAPI_TOKEN');
  if (!pixel || !token) return 'skip';
  var ud = {};
  if (lead.fbp) ud.fbp = lead.fbp;
  if (lead.fbc) ud.fbc = lead.fbc;
  if (lead.ua) ud.client_user_agent = lead.ua;
  var ph = digits_(lead.phone); if (ph) ud.ph = [sha_(ph)];
  var fn = String(lead.name || '').split(' ')[0]; if (fn) ud.fn = [sha_(fn)];
  if (lead.id) ud.external_id = [sha_(lead.id)];
  var t = lead.ts ? Math.floor(new Date(lead.ts).getTime() / 1000) : Math.floor(Date.now() / 1000);
  // Meta принимает веб-события не старше 7 дней
  var now = Math.floor(Date.now() / 1000);
  if (now - t > 6.5 * 86400) t = now;
  var ev = {
    event_name: name, event_time: t, event_id: eventId, action_source: 'website',
    event_source_url: lead.url || prop_('SITE_URL') || undefined, user_data: ud, custom_data: custom || {}
  };
  var body = { data: [ev] };
  if (prop_('TEST_CODE')) body.test_event_code = prop_('TEST_CODE');
  try {
    var r = UrlFetchApp.fetch('https://graph.facebook.com/v19.0/' + pixel + '/events?access_token=' + token, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true
    });
    return r.getResponseCode() + '';
  } catch (e) { return 'err'; }
}

/* ---------- приём от сайта ---------- */
function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad json' }); }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    switch (d.a) {
      case 'lead':   return out_(addLead_(d));
      case 'touch':  return out_(touch_(d));
      case 'event':  return out_(forward_(d));
      case 'list':   return out_(list_(d));
      case 'status': return out_(status_(d));
    }
    return out_({ ok: false, error: 'unknown action' });
  } finally { lock.releaseLock(); }
}
function doGet(e) {
  var p = e.parameter || {};
  if (p.a === 'list') return out_(list_({ key: p.key, day: p.day }));
  return out_({ ok: true, service: 'tourisme-leads' });
}

function addLead_(d) {
  var L = d.lead || {};
  if (!L.id) return { ok: false, error: 'no id' };
  if (findRow_(L.id)) return { ok: true, dup: true };
  var row = COLS.map(function (c) {
    switch (c) {
      case 'created': return Utilities.formatDate(new Date(), 'Asia/Tashkent', "yyyy-MM-dd'T'HH:mm:ss");
      case 'updated': return '';
      case 'status': return 'new';
      case 'tg_click': return '';
      case 'event_id': return d.event_id || '';
      case 'source': return L.source || '';
      default: return L[c] !== undefined ? L[c] : '';
    }
  });
  sheet_().appendRow(row);
  var r = capi_('Lead', d.event_id, { id: L.id, fbp: L.fbp, fbc: L.fbc, ua: d.ua, url: d.url, phone: L.phone, name: L.name },
    d.custom_data);
  return { ok: true, capi: r };
}
function touch_(d) {
  var row = findRow_(d.id); if (!row) return { ok: false, error: 'not found' };
  var allowed = ['lang','name','phone','dests','vibe','when','nights','adults','kids','ages','budget','passport'];
  var f = d.fields || {};
  allowed.forEach(function (k) { if (f[k] !== undefined && f[k] !== '') setCell_(row, k, f[k]); });
  if (d.tg) setCell_(row, 'tg_click', 'да');
  setCell_(row, 'updated', Utilities.formatDate(new Date(), 'Asia/Tashkent', "yyyy-MM-dd'T'HH:mm:ss"));
  return { ok: true };
}
function forward_(d) { // QuizStart, Contact и т.п. из браузера
  var r = capi_(d.event_name, d.event_id, { fbp: d.fbp, fbc: d.fbc, ua: d.ua, url: d.event_source_url, id: d.lead_id, phone: d.phone },
    d.custom_data);
  return { ok: true, capi: r };
}

/* ---------- панель менеджера ---------- */
function auth_(key) { var k = prop_('MANAGER_KEY'); return k && key === k; }

function list_(d) {
  if (!auth_(d.key)) return { ok: false, error: 'auth' };
  var day = d.day || Utilities.formatDate(new Date(), 'Asia/Tashkent', 'yyyy-MM-dd');
  var items = rows_().filter(function (o) { return String(o.created).slice(0, 10) === day; })
    .map(function (o) { delete o._row; delete o.fbp; delete o.fbc; delete o.event_id; return o; })
    .reverse();
  var all = rows_(), stats = { new: 0, qual: 0, sale: 0, trash: 0, amount: 0 };
  items.forEach(function (o) { stats[o.status] = (stats[o.status] || 0) + 1; if (o.status === 'sale') stats.amount += Number(o.amount) || 0; });
  return { ok: true, day: day, stats: stats, items: items, total: all.length };
}

function status_(d) {
  if (!auth_(d.key)) return { ok: false, error: 'auth' };
  var row = findRow_(d.id); if (!row) return { ok: false, error: 'not found' };
  var allowed = { new: 1, trash: 1, qual: 1, sale: 1 };
  if (!allowed[d.status]) return { ok: false, error: 'bad status' };
  var before = getRowObj_(row), res = {};
  setCell_(row, 'status', d.status);
  if (d.status === 'sale') setCell_(row, 'amount', Number(d.amount) || 0);
  if (d.note !== undefined) setCell_(row, 'note', d.note);
  setCell_(row, 'updated', Utilities.formatDate(new Date(), 'Asia/Tashkent', "yyyy-MM-dd'T'HH:mm:ss"));
  var L = { id: before.id, fbp: before.fbp, fbc: before.fbc, phone: before.phone, name: before.name, ts: before.created };
  // в Meta отправляем только при переходе вперёд, один раз
  if (d.status === 'qual' && before.status !== 'qual' && before.status !== 'sale') {
    res.capi = capi_('QualifiedLead', 'q_' + before.id, L, { content_category: before.dests, currency: 'USD' });
  }
  if (d.status === 'sale' && before.status !== 'sale') {
    if (before.status !== 'qual') capi_('QualifiedLead', 'q_' + before.id, L, { content_category: before.dests, currency: 'USD' });
    res.capi = capi_('Purchase', 'p_' + before.id, L, { value: Number(d.amount) || 0, currency: 'USD', content_category: before.dests });
  }
  return { ok: true, res: res };
}
