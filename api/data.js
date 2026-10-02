'use strict';
/* ================================================================
 * GET /api/data[?refresh=1]
 *
 * Lee "Cuadro Control 2026.xlsx" directo desde Microsoft Graph con
 * credenciales de aplicación (client credentials) — no depende de
 * que cada persona tenga su propia cuenta de Microsoft 365 conectada.
 *
 * MODO A (sin credenciales de Azure): EXCEL_SHARE_URL
 *   Enlace de SharePoint/OneDrive con "Cualquier persona con el enlace"
 *   (solo lectura). Se descarga sin autenticación.
 *
 * MODO B (Microsoft Graph, credenciales de aplicación). Si EXCEL_SHARE_URL
 * está definida se usa el modo A y estas no hacen falta:
 *   AZURE_TENANT_ID
 *   AZURE_CLIENT_ID
 *   AZURE_CLIENT_SECRET
 *   GRAPH_DRIVE_ID     (id del drive de SharePoint/OneDrive)
 *   GRAPH_ITEM_ID      (id del archivo .xlsx dentro de ese drive)
 * Opcional:
 *   BUDGET_YEAR        (por defecto 2026)
 *   DATA_CACHE_TTL_MS  (por defecto 45000 = 45s)
 *
 * Ver README.md para instrucciones de cómo obtener estos valores.
 * ================================================================ */

var XLSX = require('xlsx');
var parse = require('../lib/parse');

var GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
var cache = { at: 0, payload: null };

function required(name){
  var v = process.env[name];
  if (!v) { var e = new Error('missing_env:' + name); e.status = 500; throw e; }
  return v;
}

function ttl(){
  var n = Number(process.env.DATA_CACHE_TTL_MS);
  return isFinite(n) && n >= 0 ? n : 45000;
}

async function getToken(){
  var tenant = required('AZURE_TENANT_ID');
  var body = new URLSearchParams({
    client_id: required('AZURE_CLIENT_ID'),
    client_secret: required('AZURE_CLIENT_SECRET'),
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  var r = await fetch('https://login.microsoftonline.com/' + tenant + '/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body
  });
  if (!r.ok){
    var t = await r.text().catch(function(){ return ''; });
    var e = new Error('auth_failed:' + r.status + (t ? ':' + t.slice(0,200) : ''));
    e.status = 502; throw e;
  }
  var j = await r.json();
  return j.access_token;
}


function looksLikeXlsx(buf){ return buf && buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4B; }

// Descarga con redirecciones manuales conservando cookies (SharePoint suele
// entregar una cookie en el primer salto y exigirla en el siguiente).
async function fetchFollow(url, maxHops){
  var jar = {}, cur = url;
  for (var hop = 0; hop <= (maxHops || 6); hop++){
    var headers = { 'User-Agent': 'Mozilla/5.0 (compatible; MonitorBudget/1.0)', 'Accept': '*/*' };
    var cookie = Object.keys(jar).map(function(k){ return k + '=' + jar[k]; }).join('; ');
    if (cookie) headers.Cookie = cookie;
    var r = await fetch(cur, { redirect: 'manual', headers: headers });
    var set = typeof r.headers.getSetCookie === 'function' ? r.headers.getSetCookie() : [];
    set.forEach(function(c){ var kv = c.split(';')[0], i = kv.indexOf('='); if (i > 0) jar[kv.slice(0, i).trim()] = kv.slice(i + 1); });
    if (r.status >= 300 && r.status < 400 && r.headers.get('location')){
      cur = new URL(r.headers.get('location'), cur).toString();
      continue;
    }
    return r;
  }
  var e = new Error('share_link_failed:too_many_redirects'); e.status = 502; throw e;
}

function shareCandidates(link){
  var u = link.trim();
  var withDl = u + (u.indexOf('?') > -1 ? '&' : '?') + 'download=1';
  var enc = 'u!' + Buffer.from(u).toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  var origin = new URL(u).origin;
  return [
    withDl,
    origin + '/_api/v2.0/shares/' + enc + '/driveitem/content',
    'https://api.onedrive.com/v1.0/shares/' + enc + '/root/content'
  ];
}

async function downloadFromShareLink(){
  var link = required('EXCEL_SHARE_URL');
  var tried = [];
  var urls = shareCandidates(link);
  for (var i = 0; i < urls.length; i++){
    try{
      var r = await fetchFollow(urls[i]);
      if (!r.ok){ tried.push(i + ':http_' + r.status); continue; }
      var buf = Buffer.from(await r.arrayBuffer());
      if (!looksLikeXlsx(buf)){ tried.push(i + ':not_xlsx'); continue; }
      return XLSX.read(buf, { type: 'buffer', cellDates: true });
    } catch (err){ tried.push(i + ':' + String((err && err.message) || err).slice(0, 60)); }
  }
  var e = new Error('share_link_failed:' + tried.join(','));
  e.status = 502; throw e;
}

async function downloadWorkbook(token){
  var driveId = required('GRAPH_DRIVE_ID'), itemId = required('GRAPH_ITEM_ID');
  var r = await fetch(GRAPH_BASE + '/drives/' + encodeURIComponent(driveId) + '/items/' + encodeURIComponent(itemId) + '/content', {
    headers: { Authorization: 'Bearer ' + token }
  });
  if (!r.ok){
    var e = new Error('graph_download_failed:' + r.status);
    e.status = r.status === 404 ? 502 : 502; throw e;
  }
  var buf = Buffer.from(await r.arrayBuffer());
  return XLSX.read(buf, { type: 'buffer', cellDates: true });
}

function sheetsFromWorkbook(wb){
  var out = {};
  wb.SheetNames.forEach(function(name){
    out[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' });
  });
  return out;
}

function pickSheets(sheets){
  var names = Object.keys(sheets);
  var bName = names.filter(function(n){ return parse.norm(n).indexOf('budget') > -1; })[0];
  var dName = names.filter(function(n){
    return (sheets[n] || []).slice(0, 10).some(function(row){
      return row.some(function(c){ return parse.norm(c) === 'requerimiento'; });
    });
  })[0];
  return { bName: bName, dName: dName };
}

async function buildPayload(){
  var wb;
  if (process.env.EXCEL_SHARE_URL){
    wb = await downloadFromShareLink();
  } else {
    var token = await getToken();
    wb = await downloadWorkbook(token);
  }
  var sheets = sheetsFromWorkbook(wb);
  var picked = pickSheets(sheets);
  if (!picked.bName){ var e1 = new Error('budget_sheet_not_found'); e1.status = 502; throw e1; }

  var year = Number(process.env.BUDGET_YEAR) || 2026;
  var budget = parse.parseBudget(sheets[picked.bName]);
  var notes = [];
  var records = [];
  var dataError = '';
  if (picked.dName){
    var d = parse.parseData(sheets[picked.dName], year);
    if (d.error) dataError = d.error; else { records = d.records; notes = notes.concat(d.notes || []); }
  } else {
    notes.push('No encontré la pestaña con los registros de gasto (columna REQUERIMIENTO).');
  }
  if (!budget.budgets.length){
    var e2 = new Error('budget_sheet_empty:' + picked.bName); e2.status = 502; throw e2;
  }

  return {
    updatedAt: new Date().toISOString(),
    year: year,
    budgets: budget.budgets,
    records: records,
    notes: notes,
    dataError: dataError
  };
}

module.exports = async function handler(req, res){
  try{
    var force = req.query && (req.query.refresh === '1' || req.query.refresh === 'true');
    var fresh = cache.payload && (Date.now() - cache.at) < ttl();
    if (!force && fresh){
      res.setHeader('Cache-Control', 'no-store');
      res.setHeader('X-Data-Cache', 'hit');
      return res.status(200).json(cache.payload);
    }
    var payload = await buildPayload();
    cache = { at: Date.now(), payload: payload };
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('X-Data-Cache', 'miss');
    return res.status(200).json(payload);
  } catch (err){
    // Si falla una actualización pero ya hay datos en caché, se sirven esos datos
    // (con una marca de advertencia) en vez de dejar el dashboard en blanco.
    var code = String((err && err.message) || err);
    if (cache.payload){
      res.setHeader('Cache-Control', 'no-store');
      res.setHeader('X-Data-Cache', 'stale');
      return res.status(200).json(Object.assign({}, cache.payload, { staleError: code }));
    }
    var status = (err && err.status) || 502;
    return res.status(status).json({ error: true, code: code });
  }
};
