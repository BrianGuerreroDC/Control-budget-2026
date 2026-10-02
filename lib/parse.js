'use strict';
/* ================================================================
 * Lógica de parseo del Excel "Cuadro Control 2026".
 * Idéntica a la usada en el dashboard (ya probada): países/budget en
 * "Budget por Mes", registros de gasto en la pestaña con columna
 * REQUERIMIENTO. Sin dependencias de navegador: se usa tanto en el
 * backend (api/data.js) como se podría reusar en el frontend.
 * ================================================================ */

var EPS = 1e-6;

function norm(s){ return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toLowerCase(); }

function itemKey(s){
  return norm(s).replace(/\s*\|\s*/g,' | ').replace(/ammend/g,'amend').replace(/renewals\b/g,'renewal').replace(/amendments\b/g,'amendment');
}

function prettyCountry(n){
  var k = norm(n);
  if (k === 'mexico') return 'México';
  if (k === 'peru') return 'Perú';
  return String(n).trim();
}

function parseMoney(v){
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v == null ? '' : v).trim();
  if (!s) return 0;
  var neg = /^\(.*\)$/.test(s) || /^[^\d]*-\s*\d/.test(s);
  s = s.replace(/[^\d.,]/g,'');
  if (!s) return 0;
  var ld = s.lastIndexOf('.'), lc = s.lastIndexOf(','), sep = null;
  if (ld > -1 && lc > -1) sep = ld > lc ? '.' : ',';
  else if (ld > -1 || lc > -1){
    var c = ld > -1 ? '.' : ',';
    var count = s.split(c).length - 1;
    var tail = s.length - s.lastIndexOf(c) - 1;
    if (count === 1 && (tail !== 3 || /^0[.,]/.test(s))) sep = c;
  }
  var n;
  if (sep){ var i = s.lastIndexOf(sep); n = parseFloat(s.slice(0,i).replace(/[.,]/g,'') + '.' + s.slice(i+1)); }
  else n = parseFloat(s.replace(/[.,]/g,''));
  if (!isFinite(n)) return 0;
  return neg ? -n : n;
}

function parseDate(v){
  if (v == null || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : { y:v.getFullYear(), m:v.getMonth(), d:v.getDate() };
  var t = String(v).trim();
  if (!t) return null;
  var m;
  if (typeof v === 'number' || /^\d{5}(\.\d+)?$/.test(t)){
    var n = Number(t);
    if (n > 20000 && n < 80000){ var dt = new Date(Math.round((n - 25569) * 86400000)); return { y:dt.getUTCFullYear(), m:dt.getUTCMonth(), d:dt.getUTCDate() }; }
    return null;
  }
  m = t.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m && +m[2] >= 1 && +m[2] <= 12) return { y:+m[1], m:+m[2]-1, d:+m[3] };
  m = t.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})/);
  if (m){
    var a = +m[1], b = +m[2], y = +m[3]; if (y < 100) y += 2000;
    var d = a, mo = b; if (b > 12 && a <= 12){ d = b; mo = a; }
    if (mo >= 1 && mo <= 12) return { y:y, m:mo-1, d:d };
    return null;
  }
  var MN = { ene:0,jan:0,feb:1,mar:2,abr:3,apr:3,may:4,jun:5,jul:6,ago:7,aug:7,sep:8,set:8,oct:9,nov:10,dic:11,dec:11 };
  m = norm(t).match(/(\d{1,2})\s*(?:de\s+)?([a-z]{3})[a-z]*\.?[\s,\-\/]*(?:de\s+|del\s+)?(\d{2,4})/);
  if (m && MN[m[2]] !== undefined){ var yy = +m[3]; if (yy < 100) yy += 2000; return { y:yy, m:MN[m[2]], d:+m[1] }; }
  var dd = new Date(t);
  return isNaN(dd) ? null : { y:dd.getFullYear(), m:dd.getMonth(), d:dd.getDate() };
}

function groupFromBudgetText(t){
  var n = norm(t);
  if (n.indexOf('certif') > -1) return { key:'COC', label:'COC' };
  if (n.indexOf('trade') > -1) return { key:'TRADE', label:'TRADE' };
  if (n.indexOf('regist') > -1 || n.indexOf('hbc') > -1 || n.indexOf('food') > -1) return { key:'REG', label:'HBC + FOOD' };
  return { key:'G:' + n, label:String(t).trim() };
}
function groupFromEquipo(t){
  var n = norm(t);
  if (!n) return { key:'G:-', label:'Sin equipo' };
  if (n === 'coc' || n.indexOf('certif') > -1) return { key:'COC', label:'COC' };
  if (n.indexOf('trade') > -1) return { key:'TRADE', label:'TRADE' };
  if (n === 'food' || n.indexOf('food') > -1) return { key:'FOOD', label:'FOOD' };
  if (n === 'hbc' || n.indexOf('hbc') > -1 || n.indexOf('regist') > -1) return { key:'HBC', label:'HBC' };
  return { key:'G:' + n, label:String(t).trim() };
}
// budget "Registration HBC & Food" sin dividir -> se compara igual contra FOOD y HBC.
function budgetGroupOf(displayGroup){ return (displayGroup === 'FOOD' || displayGroup === 'HBC') ? 'REG' : displayGroup; }
var GROUP_ORDER = ['COC','FOOD','HBC','TRADE'];

// Pestaña "Budget por Mes": países en la fila 1 (etiqueta | valor), bloques por equipo
function parseBudget(rows){
  var out = { budgets:[] };
  var r0 = -1;
  for (var i = 0; i < rows.length; i++){ if (rows[i].some(function(c){ return String(c == null ? '' : c).trim() !== ''; })){ r0 = i; break; } }
  if (r0 < 0) return out;
  var hdr = rows[r0], cols = [];
  for (var c = 0; c < hdr.length; c++){
    var nm = String(hdr[c] == null ? '' : hdr[c]).trim();
    if (nm) cols.push({ idx:c, name:nm, key:norm(nm) });
  }
  var group = null;
  for (var r = r0 + 1; r < rows.length; r++){
    var row = rows[r];
    var cell = function(k){ var v = row[k]; return v == null ? '' : v; };
    var labels = cols.map(function(cc){ return String(cell(cc.idx)).trim(); });
    if (!labels.some(Boolean)) continue;
    var vals = cols.map(function(cc){ return cell(cc.idx + 1); });
    var allEmpty = vals.every(function(v){ return String(v).trim() === ''; });
    if (allEmpty){ group = groupFromBudgetText(labels.filter(Boolean)[0]); continue; }
    if (!group) continue;
    cols.forEach(function(cc, k){
      if (!labels[k]) return;
      out.budgets.push({ ck:cc.key, cname:prettyCountry(cc.name), group:group.key, glabel:group.label, ik:itemKey(labels[k]), label:labels[k], monthly:parseMoney(vals[k]) });
    });
  }
  return out;
}

// Pestaña de gasto: solo PAIS, EQUIPO, REQUERIMIENTO, Monto en Dolares (+ fecha para ubicar el mes)
function parseData(rows, year){
  year = year || 2026;
  var res = { records:[], notes:[], error:'' };
  var hi = -1;
  for (var i = 0; i < rows.length; i++){ if (rows[i].some(function(c){ return norm(c) === 'requerimiento'; })){ hi = i; break; } }
  if (hi < 0){ res.error = 'No encontré la columna REQUERIMIENTO.'; return res; }
  var hdr = rows[hi].map(norm);
  var col = function(){ for (var a = 0; a < arguments.length; a++){ var k = hdr.indexOf(arguments[a]); if (k >= 0) return k; } return -1; };
  var cP = col('pais'), cE = col('equipo'), cR = col('requerimiento'), cU = col('monto en dolares','monto en dolares (usd)','monto usd'),
      cF1 = col('fecha de solicitud'), cF2 = col('fecha de correo con orden');
  if (cP < 0 || cR < 0 || cU < 0){ res.error = 'Faltan columnas en la pestaña de gasto (PAIS, REQUERIMIENTO o Monto en Dolares).'; return res; }
  var noDate = 0, otherYear = 0, noReq = 0;
  for (var r = hi + 1; r < rows.length; r++){
    var row = rows[r];
    var g = function(k){ return k < 0 ? '' : (row[k] == null ? '' : row[k]); };
    var req = String(g(cR)).trim(), pais = String(g(cP)).trim();
    if (!req && !pais && String(g(cU)).trim() === '') continue;
    if (!req){ noReq++; continue; }
    var d = parseDate(g(cF1)) || parseDate(g(cF2));
    if (!d){ noDate++; continue; }
    if (d.y !== year){ otherYear++; continue; }
    var grp = groupFromEquipo(g(cE));
    var paisTxt = pais || 'Sin país';
    res.records.push({ ck:norm(paisTxt), cname:prettyCountry(paisTxt), group:grp.key, glabel:grp.label, ik:itemKey(req), label:req, m:d.m, usd:parseMoney(g(cU)) });
  }
  if (noDate) res.notes.push(noDate + ' registro(s) sin fecha válida no se incluyeron en el monitor.');
  if (otherYear) res.notes.push(otherYear + ' registro(s) con fecha fuera de ' + year + ' no se incluyeron.');
  if (noReq) res.notes.push(noReq + ' registro(s) sin requerimiento no se incluyeron.');
  return res;
}

module.exports = {
  EPS: EPS, norm: norm, itemKey: itemKey, prettyCountry: prettyCountry,
  parseMoney: parseMoney, parseDate: parseDate,
  groupFromBudgetText: groupFromBudgetText, groupFromEquipo: groupFromEquipo,
  budgetGroupOf: budgetGroupOf, GROUP_ORDER: GROUP_ORDER,
  parseBudget: parseBudget, parseData: parseData
};
