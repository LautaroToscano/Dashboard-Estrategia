/******************************************************************************
 * MONITOR DE MERCADO — Soberanos USD · LECAPs/BONCAPs · Carry trade
 *
 * Fuentes:
 *   - BYMA  (open.bymadata.com.ar)  → precios: public-bonds (bonos/BONCAPs) y lebacs (LECAPs)
 *   - A3    (api.marketdata.mae.com.ar) → flujos de fondos USD (validación / fuente)
 *   - FRED  (DGS10 csv)              → UST 10 años
 *   - ArgentinaDatos                  → riesgo país (si falla, se carga a mano en Config)
 *
 * Estructura (pensada para portar a la web):
 *   1) FETCH       → traer*_()        : solo llaman APIs y devuelven datos crudos
 *   2) CÁLCULO     → funciones puras  : tir_, pv_, retornoHorizonte_, calcLecaps_, ...
 *   3) ESCRITURA   → hoja*_()         : lo único que toca SpreadsheetApp
 *
 * Uso:
 *   1. configurarInicial()   → crea Config, Feriados, Lecaps_VF y Flujos (una sola vez)
 *   2. actualizar()          → trae precios y arma Soberanos, Lecaps y Carry
 *   3. instalarTrigger()     → actualiza cada 15 min en horario de mercado
 ******************************************************************************/

const TZ = 'America/Argentina/Buenos_Aires';
const BYMA_URL = 'https://open.bymadata.com.ar/vanoms-be-core/rest/api/bymadata/free/';
const A3_URL = 'https://api.marketdata.mae.com.ar/api/';

const C = {
  header: '#0B2A5B', headerTxt: '#FFFFFF', sub: '#1F3F77',
  red: '#E67C73', white: '#FFFFFF', green: '#57BB8A', yellow: '#FFE599',
  stale: '#999999', verdeFuerte: '#2E9D52', rojoFuerte: '#E0443A'
};

/* =============================================================================
 * TÉRMINOS DE BONOS USD (reestructuración 2020). GD (ley NY) y AL/AE (ley AR)
 * comparten cronograma. Cupón 30/360 semestral (9/1 y 9/7) sobre valor residual.
 * Se usan solo si A3 no trae los flujos del bono. VALIDAR contra la hoja Flujos.
 * ========================================================================== */
const TERMINOS_USD = {
  '29': { amort: amortIguales_('2025-01-09', 10),
          cupon: [['2020-09-04', 1.0]] },
  '30': { amort: [['2024-07-09', 4]].concat(amortIguales_('2025-01-09', 12, 8)),
          cupon: [['2020-09-04', 0.125], ['2021-07-09', 0.5], ['2023-07-09', 0.75], ['2027-07-09', 1.75]] },
  '35': { amort: amortIguales_('2031-01-09', 10),
          cupon: [['2020-09-04', 0.125], ['2021-07-09', 1.125], ['2022-07-09', 1.5], ['2023-07-09', 3.625],
                  ['2024-07-09', 4.125], ['2027-07-09', 4.75], ['2028-07-09', 5.0]] },
  '38': { amort: amortIguales_('2027-07-09', 22),
          cupon: [['2020-09-04', 0.125], ['2021-07-09', 2.0], ['2022-07-09', 3.875], ['2023-07-09', 4.25],
                  ['2024-07-09', 5.0]] },
  '41': { amort: amortIguales_('2028-01-09', 28),
          cupon: [['2020-09-04', 0.125], ['2021-07-09', 2.5], ['2022-07-09', 3.5], ['2029-07-09', 4.875]] }
};
/* Bullets con cupón fijo, según ficha técnica BYMA. Fechas de pago = vencimiento hacia atrás.
 * AO28/AO29: 6% anual, pago mensual (fin de mes). AN29: 6,5% semestral 30/5 y 30/11.
 * TY30P (ARS): 29,5% semestral 30/5 y 30/11. */
const BULLETS = [
  { ticker: 'AO28', venc: '2028-10-31', tasa: 6.0, meses: 1 },
  { ticker: 'AO29', venc: '2029-10-31', tasa: 6.0, meses: 1 },
  { ticker: 'AN29', venc: '2029-11-30', tasa: 6.5, meses: 6 },
  { ticker: 'TY30P', venc: '2030-05-30', tasa: 29.5, meses: 6 }
];

const TICKER_A_TERMINO = {
  GD29: '29', AL29: '29', GD30: '30', AL30: '30', GD35: '35', AL35: '35',
  GD38: '38', AE38: '38', GD41: '41', AL41: '41'
};

/* =============================================================================
 * 0. SETUP
 * ========================================================================== */
function configurarInicial() {
  const ss = SpreadsheetApp.getActive();

  if (!ss.getSheetByName('Config')) {
    const sh = ss.insertSheet('Config');
    const rows = [
      ['Parámetro', 'Valor', 'Nota'],
      ['UST 10y (%)', '', 'Automático (FRED DGS10). Si falla, queda el último valor.'],
      ['Riesgo país (pbs)', '', 'Automático (ArgentinaDatos). Si falla, cargalo a mano.'],
      ['Escenarios RP (pbs)', '300; 350; 400; 450; 700', 'Separados por ";". El riesgo país actual se agrega solo.'],
      ['Horizonte 1', '', 'Ya no se usa: Riesgo País toma sola los próximos dos cierres de semestre'],
      ['Horizonte 2', '', 'Ya no se usa: la hoja Soberanos toma sola los próximos dos cierres de semestre'],
      ['Grilla TIR salida (%)', '5; 6; 7; 7,5; 8; 9; 10; 12; 14; 16; 20', 'Separados por ";"'],
      ['Tickers GD', 'GD29; GD30; GD35; GD38; GD41', ''],
      ['Tickers AL', 'AL29; AL30; AL35; AE38; AL41', ''],
      ['Tickers matriz', 'AO28; AL29; GD29; AO29; AN29; AL30; GD30; AL35; GD35; AE38; GD38; AL41; GD41', 'Se valúan con precio D (USD MEP)'],
      ['Liquidación bonos USD', '24hs', 'CI o 24hs'],
      ['Base TNA', 365, '365 (mercado) o 360 (Finanzas)'],
      ['Grilla dólar', '1450; 1500; 1520; 1530; 1540; 1550; 1560; 1570; 1580; 1600; 1650; 1700', 'Montos redondos, separados por ";" (el script los ordena)'],
      ['Dólar MEP', '', 'Calculado: AL30 / AL30D (24hs)'],
      ['Última actualización', '', '']
    ];
    sh.getRange(1, 1, rows.length, 3).setValues(rows);
    sh.getRange('B5:B6').setNumberFormat('dd/mm/yyyy');
    estiloHeader_(sh.getRange(1, 1, 1, 3));
    sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 420); sh.setColumnWidth(3, 380);
  }

  if (!ss.getSheetByName('Feriados')) {
    const sh = ss.insertSheet('Feriados');
    const fer = [
      ['2026-10-12', 'Diversidad Cultural'], ['2026-11-09', 'Feriado (agregado)'], ['2026-11-23', 'Soberanía Nacional (trasladado)'],
      ['2026-12-08', 'Inmaculada Concepción'], ['2026-12-25', 'Navidad'],
      ['2027-01-01', 'Año Nuevo'], ['2027-02-08', 'Carnaval'], ['2027-02-09', 'Carnaval'],
      ['2027-03-24', 'Memoria'], ['2027-03-25', 'Jueves Santo'], ['2027-03-26', 'Viernes Santo'],
      ['2027-04-02', 'Malvinas'], ['2027-05-25', 'Revolución de Mayo'], ['2027-06-21', 'Güemes (trasladado)'],
      ['2027-07-09', 'Independencia'], ['2027-08-16', 'San Martín (trasladado)'],
      ['2027-10-11', 'Diversidad Cultural (trasladado)'], ['2027-12-08', 'Inmaculada Concepción']
    ].map(r => [fecha_(r[0]), r[1]]);
    sh.getRange(1, 1, 1, 2).setValues([['Fecha', 'Feriado (VERIFICAR con calendario oficial / BYMA)']]);
    sh.getRange(2, 1, fer.length, 2).setValues(fer);
    sh.getRange(2, 1, fer.length, 1).setNumberFormat('dd/mm/yyyy');
    estiloHeader_(sh.getRange(1, 1, 1, 2));
    sh.setColumnWidth(2, 380);
  }

  if (!ss.getSheetByName('Lecaps_VF')) {
    const sh = ss.insertSheet('Lecaps_VF');
    // VF = 100 × (1 + TEM)^n, n = meses enteros desde emisión + días restantes/30 (ficha técnica BYMA).
    // Si hay Emisión y TEM, el script calcula el VF solo; si no, usa el VF cargado a mano.
    // [ticker, vencimiento, VF manual, tipo, nota, emisión, TEM %]
    const lec = [
      ['S16O6', '2026-10-16', '', 'LECAP', 'Ficha BYMA', '2026-07-31', 2.05],
      ['S30O6', '2026-10-30', '', 'LECAP', 'Ficha BYMA', '2025-10-31', 2.55],
      ['S13N6', '2026-11-13', 109.65, 'LECAP', 'Ficha sin TEM: VF aproximado (TEM implícita ~2,08%)', '2026-06-30', ''],
      ['S30N6', '2026-11-30', '', 'LECAP', 'Ficha BYMA', '2025-12-15', 2.30],
      ['T15E7', '2027-01-15', '', 'BONCAP', 'Ficha BYMA', '2025-01-31', 2.05],
      ['S29E7', '2027-01-29', 111.69, 'LECAP', 'Ficha sin TEM: VF aproximado (TEM implícita ~2,25%)', '2026-08-31', ''],
      ['T30A7', '2027-04-30', '', 'BONCAP', 'Ficha BYMA', '2025-10-31', 2.55],
      ['T31Y7', '2027-05-31', '', 'BONCAP', 'Ficha BYMA', '2025-12-15', 2.40],
      ['T30J7', '2027-06-30', '', 'BONCAP', 'Ficha BYMA', '2026-01-16', 2.58],
      ['TY30P', '2030-05-30', '', 'CUPON', 'Usa hoja Flujos', '', '']
    ].map(r => {
      const vf = (r[5] && r[6]) ? round_(vfLecap_(r[5], r[1], r[6]), 4) : r[2];
      return [r[0], fecha_(r[1]), vf, r[3], r[4], r[5] ? fecha_(r[5]) : '', r[6]];
    });
    sh.getRange(1, 1, 1, 7).setValues([['Ticker', 'Vencimiento', 'VF (por 100 VN)', 'Tipo', 'Nota', 'Emisión', 'TEM (%)']]);
    sh.getRange(2, 1, lec.length, 7).setValues(lec);
    sh.getRange(2, 2, lec.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 6, lec.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 3, lec.length, 1).setNumberFormat('#,##0.0000');
    estiloHeader_(sh.getRange(1, 1, 1, 7));
    sh.setColumnWidth(5, 340);
  }

  if (!ss.getSheetByName('Flujos')) regenerarFlujos();
  SpreadsheetApp.getActive().toast('Setup listo. Revisá Lecaps_VF y Flujos, después corré actualizar().');
}

/** Rearma la hoja Flujos: A3 cuando trae el bono, términos propios si no. Pisa la hoja. */
function regenerarFlujos() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Flujos') || ss.insertSheet('Flujos');
  const hoy = hoyISO_();
  const a3 = traerFlujosA3_();                         // {TICKER: [{fecha, renta, amort, total}]}
  const out = [['Ticker', 'Fecha', 'Renta', 'Amortización', 'Total', 'Fuente', 'Dif. vs términos']];

  Object.keys(TICKER_A_TERMINO).forEach(tk => {
    const gen = flujosDesdeTerminos_(TERMINOS_USD[TICKER_A_TERMINO[tk]]).filter(f => f.fecha > hoy);
    const genMap = {}; gen.forEach(f => genMap[f.fecha] = f.total);
    const src = a3[tk] && a3[tk].length ? a3[tk] : null;
    (src || gen).forEach(f => {
      const dif = src ? (genMap[f.fecha] === undefined ? 'fecha no está en términos' : round_(f.total - genMap[f.fecha], 4)) : '';
      out.push([tk, fecha_(f.fecha), f.renta, f.amort, f.total, src ? 'A3' : 'Términos', dif]);
    });
  });

  // Bullets con cupón fijo (condiciones de la ficha técnica BYMA). Cupón = tasa × meses/12 (30/360).
  BULLETS.forEach(b => {
    flujosBullet_(b.venc, b.tasa, b.meses).filter(f => f.fecha > hoy).forEach(f => {
      out.push([b.ticker, fecha_(f.fecha), f.renta, f.amort, f.total, 'Ficha BYMA', '']);
    });
  });

  sh.clear();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(2, 2, out.length - 1, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, 3, out.length - 1, 3).setNumberFormat('#,##0.0000');
  estiloHeader_(sh.getRange(1, 1, 1, out[0].length));
  sh.setFrozenRows(1);
}

/** Manual: busca LECAPs/BONCAPs nuevas en BYMA y las agrega a Lecaps_VF. (También corre dentro de actualizar.) */
function detectarNuevasLetras() {
  const n = detectarNuevasLetras_(traerPreciosByma_(), hoyISO_());
  if (n.length) SpreadsheetApp.getActive().toast('Agregadas: ' + n.join(', '));
}

/**
 * Compara las letras tasa fija que cotizan en BYMA (S16O6, T30J7…) contra Lecaps_VF.
 * Para cada nueva baja la ficha técnica, confirma que capitaliza tasa fija, saca emisión,
 * vencimiento y TEM, y agrega la fila. Si la ficha no trae TEM, la marca en amarillo.
 */
function detectarNuevasLetras_(precios, hoy) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Lecaps_VF');
  if (!sh) return [];
  const existentes = new Set(sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()) : []);
  const patron = /^[ST]\d{2}[A-Z]\d$/;                         // S16O6, T30J7 (excluye TX26, TZX26, TTM26, TY30P)
  const candidatos = Object.keys(precios).map(k => k.split('|')[0])
    .filter((s, i, arr) => patron.test(s) && arr.indexOf(s) === i && !existentes.has(s));

  const cache = CacheService.getScriptCache();             // descartes recientes: no re-consultar su ficha por 6 h (máximo de CacheService)
  const nuevas = [];
  candidatos.filter(tk => !cache.get('descartada_' + tk)).forEach(tk => {
    const descartar = () => cache.put('descartada_' + tk, '1', 21600);
    let f;
    try { f = (JSON.parse(fetchByma_('bnown/fichatecnica/especies/general', { symbol: tk })).data || [])[0]; }
    catch (e) { return; }
    Utilities.sleep(300);
    if (!f) { descartar(); return; }
    const venc = String(f.fechaVencimiento || '').slice(0, 10), emision = String(f.fechaEmision || '').slice(0, 10);
    if (!venc || venc <= hoy) { descartar(); return; }
    const texto = String(f.interes || f.intereses || '');
    const soloTasa = /^\s*\d+[.,]\d+\s*%\s*$/.test(texto);
    if (!/capitaliz/i.test(texto) && !soloTasa) { descartar(); return; }   // no es tasa fija capitalizable (CER, dual, etc.)
    const m = texto.match(/efectiva mensual[^0-9]*(\d+[.,]\d+)/i) || (soloTasa ? texto.match(/(\d+[.,]\d+)/) : null);
    const tem = m ? num_(m[1]) : '';
    const tipo = tk[0] === 'S' ? 'LECAP' : 'BONCAP';
    const vf = (tem && emision) ? round_(vfLecap_(emision, venc, tem), 4) : '';
    const nota = tem ? 'Alta automática (ficha BYMA) ' + fmt_(hoy) : 'FALTA TEM: cargala del resultado de la licitación';
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 7).setValues([[tk, fecha_(venc), vf, tipo, nota, emision ? fecha_(emision) : '', tem]]);
    sh.getRange(row, 2).setNumberFormat('dd/mm/yyyy');
    sh.getRange(row, 6).setNumberFormat('dd/mm/yyyy');
    sh.getRange(row, 3).setNumberFormat('#,##0.0000');
    if (!tem) sh.getRange(row, 1, 1, 7).setBackground('#FFF2CC');
    nuevas.push(tk);
  });
  return nuevas;
}

/** Compara nuestra TIR (flujos de la hoja Flujos, al precio de A3) contra la TIR que publica A3. */
function validarTIR() {
  const r = UrlFetchApp.fetch(A3_URL + 'emisiones/flujofondoscotiz/H', { muteHttpExceptions: true });
  const flujos = leerFlujos_();
  const hoy = hoyISO_(), liq24 = proximoHabil_(hoy, leerFeriados_());
  const out = [['Ticker', 'Precio A3 (x100)', 'TIR A3', 'TIR propia (CI)', 'TIR propia (24hs)', 'Dif. vs A3 (pbs, mejor caso)']];
  JSON.parse(r.getContentText()).forEach(b => {
    const tk = String(b.especie).trim(), px = Number(b.precio) * 100, fl = flujos[tk];
    const tA3 = Number(b.tir) / 100;
    const t0 = fl ? tir_(fl, hoy, px) : null, t1 = fl ? tir_(fl, liq24, px) : null;
    const dif = (t0 === null) ? '' : Math.round(Math.min(Math.abs(t0 - tA3), Math.abs(t1 - tA3)) * 10000);
    out.push([tk, px, tA3, nz_(t0), nz_(t1), dif]);
  });
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Validación') || ss.insertSheet('Validación');
  sh.clear();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(2, 3, out.length - 1, 3).setNumberFormat('0.00%');
  estiloHeader_(sh.getRange(1, 1, 1, out[0].length));
}

/** Bono bullet con cupón fijo: fechas hacia atrás desde el vencimiento cada `meses`. */
function flujosBullet_(venc, tasa, meses) {
  const out = [];
  for (let k = 0; ; k++) {
    const d = sumarMeses_(venc, -meses * k);
    if (d < '2025-01-01') break;
    const renta = round_(tasa * meses / 12, 6);
    out.unshift({ fecha: d, renta: renta, amort: k === 0 ? 100 : 0, total: round_(renta + (k === 0 ? 100 : 0), 6) });
  }
  return out;
}

/** Vuelca la ficha técnica de BYMA de los tickers que falten, para cargar flujos/VF a mano. */
function volcarFichas() {
  const tks = ['AO28', 'AO29', 'AN29', 'TY30P', 'S16O6', 'S30O6', 'S13N6', 'S30N6', 'S29E7', 'T15E7', 'T30A7', 'T31Y7', 'T30J7'];
  const out = [['Ticker', 'Emisión', 'Vencimiento', 'Moneda', 'Amortización', 'Intereses']];
  tks.forEach(tk => {
    try {
      const j = JSON.parse(fetchByma_('bnown/fichatecnica/especies/general', { symbol: tk }));
      const f = (j.data || [])[0] || {};
      out.push([tk, f.fechaEmision || '', f.fechaVencimiento || '', f.moneda || '', f.formaAmortizacion || '', f.interes || f.intereses || JSON.stringify(f).slice(0, 500)]);
    } catch (e) { out.push([tk, '', '', '', 'ERROR', String(e)]); }
    Utilities.sleep(300);
  });
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Fichas') || ss.insertSheet('Fichas');
  sh.clear();
  sh.getRange(1, 1, out.length, out[0].length).setValues(out).setWrap(true);
  estiloHeader_(sh.getRange(1, 1, 1, out[0].length));
  sh.setColumnWidths(5, 2, 450);
}

/* =============================================================================
 * 1. MAIN
 * ========================================================================== */
function actualizar() {
  const ss = SpreadsheetApp.getActive();
  const cfg = leerConfig_();
  const feriados = leerFeriados_();
  const hoy = hoyISO_();
  const liq = { '1': hoy, '2': proximoHabil_(hoy, feriados) };  // 1 = CI, 2 = 24hs

  // --- Fetch
  const precios = traerPreciosByma_();
  try { detectarNuevasLetras_(precios, hoy); } catch (e) { Logger.log('Alta automática de letras: ' + e); }
  const ust = traerUST10y_();
  const rp = traerRiesgoPais_();
  if (ust !== null) setConfig_('UST 10y (%)', ust);
  if (rp !== null) setConfig_('Riesgo país (pbs)', rp);
  const cfg2 = leerConfig_();

  const p = (sym, st) => precios[sym + '|' + st] || null;
  const al30 = p('AL30', '2') || p('AL30', '1'), al30d = p('AL30D', '2') || p('AL30D', '1');
  if (!al30 || !al30d) throw new Error('No hay precio de AL30/AL30D para calcular el MEP.');
  const mep = al30.precio / al30d.precio;
  setConfig_('Dólar MEP', mep);

  const ctx = {
    cfg: cfg2, hoy: hoy, liq: liq, precios: precios, p: p, mep: mep,
    ust: Number(cfg2['UST 10y (%)']), rp: Number(cfg2['Riesgo país (pbs)']),
    flujos: leerFlujos_(), lecaps: leerLecapsVF_(), base: Number(cfg2['Base TNA']) || 365
  };

  // --- Escritura
  const viejaLecaps = ss.getSheetByName('Lecaps');
  if (viejaLecaps) ss.deleteSheet(viejaLecaps);                 // la hoja Lecaps ya no se usa
  hojaRiesgoPais_(ss, ctx);
  hojaSoberanos_(ss, ctx);
  hojaCarry_(ss, ctx);
  hojaPrecios_(ss, ctx);

  configurarRecalculo_(ss);
  setConfig_('Última actualización', Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'));
}

function actualizarProgramado() {
  const now = new Date();
  const dow = Number(Utilities.formatDate(now, TZ, 'u'));     // 1 = lunes … 7 = domingo
  const hhmm = Number(Utilities.formatDate(now, TZ, 'HHmm'));
  if (dow > 5 || hhmm < 1100 || hhmm > 1800) return;     // 11:00–18:00
  if (leerFeriados_().has(hoyISO_())) return;
  actualizar();
}

function instalarTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'actualizarProgramado')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('actualizarProgramado').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getActive().toast('Trigger instalado: cada 15 min, días hábiles 11:00–18:00.');
}

/* =============================================================================
 * 2. FETCH
 * ========================================================================== */
function fetchByma_(endpoint, body) {
  const r = UrlFetchApp.fetch(BYMA_URL + endpoint, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(body),
    muteHttpExceptions: true, validateHttpsCertificates: false
  });
  if (r.getResponseCode() !== 200) throw new Error('BYMA ' + endpoint + ' HTTP ' + r.getResponseCode());
  return r.getContentText();
}

/** Devuelve {'SIMBOLO|settlementType': {precio, var, stale}} de public-bonds + lebacs. */
function traerPreciosByma_() {
  const out = {};
  ['public-bonds', 'lebacs'].forEach(ep => {
    const j = JSON.parse(fetchByma_(ep, { page_size: 5000 }));
    const items = Array.isArray(j) ? j : (j.data || []);
    items.forEach(x => {
      if (!x.symbol) return;
      const key = x.symbol + '|' + x.settlementType;
      const trade = Number(x.trade) || 0, close = Number(x.closingPrice) || 0, prev = Number(x.previousClosingPrice) || 0;
      const precio = trade > 0 ? trade : (close > 0 ? close : prev);
      if (!(precio > 0)) return;
      const stale = !(trade > 0 || close > 0);
      const vr = stale ? 0 : (prev > 0 ? precio / prev - 1 : (Number(x.imbalance) || 0));
      if (out[key] && !out[key].stale) return;              // priorizar el que tiene operación hoy
      out[key] = { precio: precio, var: vr, stale: stale };
    });
    Utilities.sleep(300);
  });
  return out;
}

/** {TICKER: [{fecha, renta, amort, total}]} desde A3 (letra H = hard dollar). */
function traerFlujosA3_() {
  const out = {};
  try {
    const r = UrlFetchApp.fetch(A3_URL + 'emisiones/flujofondoscotiz/H', { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return out;
    JSON.parse(r.getContentText()).forEach(b => {
      const tk = String(b.especie).trim();
      out[tk] = (b.detalle || []).map(d => ({
        fecha: String(d.fechaPago).slice(0, 10),
        renta: Number(d.renta) || 0, amort: Number(d.amortizacion) || 0,
        total: Number(d.cashFlow) || ((Number(d.renta) || 0) + (Number(d.amortizacion) || 0))
      }));
    });
  } catch (e) { Logger.log('A3 flujos: ' + e); }
  return out;
}

function traerUST10y_() {
  try {
    const desde = Utilities.formatDate(new Date(Date.now() - 30 * 864e5), TZ, 'yyyy-MM-dd');
    const r = UrlFetchApp.fetch('https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10&cosd=' + desde, { muteHttpExceptions: true });
    const vals = r.getContentText().trim().split('\n').slice(1)
      .map(l => l.split(',')[1]).filter(v => v && v !== '.' && !isNaN(Number(v)));
    return vals.length ? Number(vals[vals.length - 1]) : null;
  } catch (e) { return null; }
}

function traerRiesgoPais_() {
  try {
    const r = UrlFetchApp.fetch('https://api.argentinadatos.com/v1/finanzas/indices/riesgo-pais/ultimo', { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return null;
    const v = Number(JSON.parse(r.getContentText()).valor);
    return v > 0 ? v : null;
  } catch (e) { return null; }
}

/* =============================================================================
 * 3. CÁLCULO (funciones puras → se portan tal cual a JS en el backend)
 * ========================================================================== */
function amortIguales_(desde, n, pct) {
  const p = pct || 100 / n, out = [];
  for (let i = 0; i < n; i++) out.push([sumarMeses_(desde, 6 * i), p]);
  return out;
}

/** Genera flujos por 100 VN original a partir de cronograma de amortización y step-up. */
function flujosDesdeTerminos_(t) {
  const ultimo = t.amort[t.amort.length - 1][0];
  const out = [];
  let prev = '2020-09-04', d = '2021-01-09', residual = 100;
  while (d <= ultimo) {
    const tasa = t.cupon.filter(s => s[0] <= prev).slice(-1)[0][1];
    const renta = residual * tasa / 100 / 2;
    const am = t.amort.filter(a => a[0] === d).reduce((s, a) => s + a[1], 0);
    out.push({ fecha: d, renta: round_(renta, 6), amort: am, total: round_(renta + am, 6) });
    residual -= am; prev = d; d = sumarMeses_(d, 6);
  }
  return out;
}

function yf_(a, b) { return (diaNum_(b) - diaNum_(a)) / 365; }

/** Valor presente (efectiva anual, ACT/365) de los flujos posteriores a `desde`. */
function pv_(flujos, desde, y) {
  return flujos.filter(f => f.fecha > desde)
    .reduce((s, f) => s + f.total / Math.pow(1 + y, yf_(desde, f.fecha)), 0);
}

/** TIR efectiva anual por bisección. Precio sucio (así cotiza BYMA). */
function tir_(flujos, liq, precio) {
  const fut = flujos.filter(f => f.fecha > liq);
  if (!fut.length || !(precio > 0)) return null;
  let lo = -0.5, hi = 5;
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    if (pv_(fut, liq, mid) > precio) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

/** Retorno de comprar hoy y vender en H a TIR y (incluye flujos cobrados hasta H, sin reinversión). */
function retornoHorizonte_(flujos, liq, precio, H, y) {
  if (H <= liq || !(precio > 0)) return null;
  const cobrado = flujos.filter(f => f.fecha > liq && f.fecha <= H).reduce((s, f) => s + f.total, 0);
  return (pv_(flujos, H, y) + cobrado) / precio - 1;
}

/** Filas de LECAPs/BONCAPs para una liquidación ('1' CI, '2' 24hs). */
function calcLecaps_(ctx, st) {
  const liq = ctx.liq[st], rows = [];
  ctx.lecaps.forEach(l => {
    const px = ctx.p(l.ticker, st);
    if (!px) return;
    const dias = diaNum_(l.venc) - diaNum_(liq);
    if (dias <= 0) return;
    const r = { ticker: l.ticker, tipo: l.tipo, precio: px.precio, var: px.var, stale: px.stale, venc: l.venc, dias: dias };
    if (l.tipo === 'CUPON') {
      const fl = ctx.flujos[l.ticker];
      const t = fl ? tir_(fl, liq, px.precio) : null;
      r.rend = t; r.tna = null; r.tem = t === null ? null : Math.pow(1 + t, 30 / 365) - 1; r.fxbe = null;
    } else {
      if (!(l.vf > 0)) return;
      r.rend = l.vf / px.precio - 1;
      r.tna = r.rend * ctx.base / dias;
      r.tem = Math.pow(1 + r.rend, 30 / dias) - 1;
      r.fxbe = ctx.mep * (1 + r.rend);
    }
    rows.push(r);
  });
  return rows.sort((a, b) => a.venc < b.venc ? -1 : 1);
}

function escenariosRP_(ctx) {
  const list = lista_(ctx.cfg['Escenarios RP (pbs)']).map(Number).filter(n => !isNaN(n));
  const out = list.filter((n, i) => list.indexOf(n) === i).map(n => ({ pbs: n, label: n + ' pbs', actual: n === ctx.rp }));
  if (ctx.rp > 0 && list.indexOf(ctx.rp) < 0) out.push({ pbs: ctx.rp, label: Math.round(ctx.rp) + ' pbs', actual: true });
  out.forEach(e => { if (e.actual) e.label += ' (actual)'; });
  return out.sort((a, b) => a.pbs - b.pbs);
}

/* =============================================================================
 * 4. ESCRITURA
 * ========================================================================== */
/**
 * Hoja "Riesgo País": tablas GD (ley NY) y AL (ley AR) con escenarios de riesgo país,
 * a los próximos dos cierres de semestre (30/06 y 31/12). Un bloque por fecha, uno debajo del otro.
 *
 * Criterio: cada escenario desplaza la TIR de CADA bono en la misma magnitud que cambia el
 * riesgo país:  TIR salida = TIR actual del bono + (RP escenario − RP actual).
 * Así la columna "actual" es puro devengamiento y se respeta la pendiente de la curva
 * y el spread de ley (cada bono parte de su propia TIR).
 */
function hojaRiesgoPais_(ss, ctx) {
  const sh = hoja_(ss, 'Riesgo País');
  sh.getDataRange().breakApart(); sh.clear(); sh.clearConditionalFormatRules();
  const rules = [];
  const st = stBonos_(ctx), liq = ctx.liq[st];
  if (!(ctx.rp > 0)) {
    sh.getRange(1, 1).setValue('Falta el riesgo país actual: cargalo en Config ("Riesgo país (pbs)").');
    return;
  }
  const esc = escenariosRP_(ctx);
  const deltas = esc.map(e => (e.pbs - ctx.rp) / 10000);
  const W = esc.length + 2;                                   // Ticker + TIR actual + escenarios
  let r0 = 1;

  proximosSemestres_(liq, 2).forEach(H => {
    let alto = 0;
    [['Tickers GD', 'Ley NY'], ['Tickers AL', 'Ley AR']].forEach((t, k) => {
      const c0 = 1 + k * (W + 1);
      const tks = lista_(ctx.cfg[t[0]]);
      sh.getRange(r0, c0, 1, W).merge().setValue('RETORNO AL ' + fmt_(H) + ' — ' + t[1]);
      estiloHeader_(sh.getRange(r0, c0, 1, W));
      const filas = [['Riesgo País', 'TIR actual'].concat(esc.map(e => e.label))];
      tks.forEach(tk => {
        const px = ctx.p(tk + 'D', st), fl = ctx.flujos[tk];
        if (!fl || !px) { filas.push([tk, fl ? 'sin precio' : 'sin flujos'].concat(esc.map(() => ''))); return; }
        const t0 = tir_(fl, liq, px.precio);
        filas.push([tk, nz_(t0)].concat(deltas.map(d => t0 === null ? '' : nz_(retornoHorizonte_(fl, liq, px.precio, H, t0 + d)))));
      });
      sh.getRange(r0 + 1, c0, filas.length, W).setValues(filas).setHorizontalAlignment('center');
      estiloHeader_(sh.getRange(r0 + 1, c0, 1, W), C.sub);
      sh.getRange(r0 + 2, c0 + 1, tks.length, 1).setNumberFormat('0.00%').setFontWeight('bold');
      sh.getRange(r0 + 2, c0 + 2, tks.length, esc.length).setNumberFormat('0.00%');
      pintar_(sh.getRange(r0 + 2, c0 + 2, tks.length, esc.length));
      bordes_(sh.getRange(r0, c0, filas.length + 1, W));
      alto = Math.max(alto, filas.length + 1);
    });
    r0 += alto + 2;
  });
  sh.getRange(r0 - 1, 1).setValue('Rendimientos en USD.')
    .setFontStyle('italic');

  const cRef = 2 * (W + 1) + 1;
  sh.getRange(2, cRef, 3, 2).setValues([
    ['Tasa 10y USA', ctx.ust / 100], ['Riesgo País', ctx.rp],
    ['Liquidación', st === '1' ? 'CI' : '24hs']
  ]);
  sh.getRange(2, cRef + 1).setNumberFormat('0.00%');
  sh.getRange(3, cRef + 1).setNumberFormat('0 "pbs"');
  estiloHeader_(sh.getRange(2, cRef, 3, 1), C.sub);
  sh.getRange(2, cRef + 1, 3, 1).setHorizontalAlignment('center');
  bordes_(sh.getRange(2, cRef, 3, 2));

  sh.setConditionalFormatRules(rules);
  sh.setColumnWidths(1, cRef + 1, 90);
  const iAct = esc.findIndex(e => e.actual);                 // columna del riesgo país actual, más ancha
  if (iAct >= 0) [0, 1].forEach(k => sh.setColumnWidth(1 + k * (W + 1) + 2 + iAct, 130));
  sh.setColumnWidth(cRef, 110);
  bloqueEstado_(sh, 5, cRef);
}

/** Hoja "Soberanos": matriz de TIR de salida a los próximos dos cierres de semestre (30/06 y 31/12). */
function hojaSoberanos_(ss, ctx) {
  const sh = hoja_(ss, 'Soberanos');
  sh.getDataRange().breakApart().clearNote(); sh.clear(); sh.clearConditionalFormatRules();
  const rules = [];
  let r = 1;
  proximosSemestres_(ctx.liq[stBonos_(ctx)], 2).forEach(H => {
    r = matrizTirSalida_(sh, ctx, r, H, rules) + 2;
  });
  sh.getRange(r - 1, 1).setValue('Rendimientos en USD, incluyen cupones y amortizaciones cobrados (sin reinversión).')
    .setFontStyle('italic');
  sh.setConditionalFormatRules(rules);
  const nG = lista_(ctx.cfg['Grilla TIR salida (%)']).length;
  sh.setColumnWidths(1, 3 + nG, 90);
  sh.setColumnWidth(3 + nG + 2, 110);
  bloqueEstado_(sh, 1, 3 + nG + 2);
}

/** Escribe una matriz desde la fila r0 y devuelve la última fila usada. */
function matrizTirSalida_(sh, ctx, r0, H, rules) {
  const st = stBonos_(ctx), liq = ctx.liq[st];
  const grid = lista_(ctx.cfg['Grilla TIR salida (%)']).map(num_).map(v => v / 100);
  const tks = lista_(ctx.cfg['Tickers matriz']);
  const W2 = 3 + grid.length;
  sh.getRange(r0, 4, 1, grid.length).merge().setValue('TIR AL ' + fmt_(H));
  estiloHeader_(sh.getRange(r0, 4, 1, grid.length));
  const filas = [['Ticker', 'Vencimiento', 'TIR Actual'].concat(grid)];
  const stale = [];
  tks.forEach(tk => {
    const px = ctx.p(tk + 'D', st), fl = ctx.flujos[tk];
    if (!fl) { filas.push([tk, '', 'sin flujos'].concat(grid.map(() => ''))); return; }
    const venc = fecha_(fl[fl.length - 1].fecha);
    if (!px) { filas.push([tk + 'D', venc, 'sin precio'].concat(grid.map(() => ''))); return; }
    if (px.stale) stale.push(filas.length);
    filas.push([tk + 'D', venc, nz_(tir_(fl, liq, px.precio))]
      .concat(grid.map(y => nz_(retornoHorizonte_(fl, liq, px.precio, H, y)))));
  });
  sh.getRange(r0 + 1, 1, filas.length, W2).setValues(filas).setHorizontalAlignment('center');
  estiloHeader_(sh.getRange(r0 + 1, 1, 1, W2), C.sub);
  sh.getRange(r0 + 1, 4, 1, grid.length).setNumberFormat('0.00%');
  sh.getRange(r0 + 2, 2, tks.length, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(r0 + 2, 3, tks.length, 1).setNumberFormat('0.00%').setFontWeight('bold');
  sh.getRange(r0 + 2, 4, tks.length, grid.length).setNumberFormat('0.00%');
  pintar_(sh.getRange(r0 + 2, 4, tks.length, grid.length));
  bordes_(sh.getRange(r0, 4, 1, grid.length));
  bordes_(sh.getRange(r0 + 1, 1, tks.length + 1, W2));
  return r0 + 1 + tks.length;
}

function stBonos_(ctx) { return String(ctx.cfg['Liquidación bonos USD']).toUpperCase() === 'CI' ? '1' : '2'; }

/** Próximos n cierres de semestre (30/06 y 31/12) posteriores a la fecha dada (la de liquidación). */
function proximosSemestres_(hoy, n) {
  const y = Number(hoy.slice(0, 4)), out = [];
  for (let k = 0; out.length < n; k++) {
    [(y + k) + '-06-30', (y + k) + '-12-31'].forEach(d => { if (d > hoy && out.length < n) out.push(d); });
  }
  return out;
}

function hojaCarry_(ss, ctx) {
  const sh = hoja_(ss, 'Carry');
  sh.getRange(1, 1, 60, 26).breakApart().clear(); sh.clearConditionalFormatRules();
  const grid = lista_(ctx.cfg['Grilla dólar']).map(num_).filter((v, i, a) => !isNaN(v) && a.indexOf(v) === i).sort((a, b) => a - b);
  const rows = calcLecaps_(ctx, '2').filter(x => x.tipo !== 'CUPON');

  sh.getRange(1, 1, 1, 2).setValues([['Dólar MEP', ctx.mep]]);
  estiloHeader_(sh.getRange(1, 1));
  sh.getRange(1, 2).setNumberFormat('$ #,##0.00').setFontWeight('bold');

  const head = ['Vencimiento', 'Ticker', 'Rend. al vto (24hs)'].concat(grid).concat(['Fx BE']);
  sh.getRange(3, 1, 1, head.length).setValues([head]);
  estiloHeader_(sh.getRange(3, 1, 1, head.length));
  sh.getRange(3, 4, 1, grid.length).setNumberFormat('$ #,##0');

  if (rows.length) {
    const vals = rows.map(x => [fecha_(x.venc), x.ticker, x.rend]
      .concat(grid.map(fx => (1 + x.rend) * ctx.mep / fx - 1)).concat([x.fxbe]));
    sh.getRange(4, 1, vals.length, head.length).setValues(vals).setHorizontalAlignment('center');
    sh.getRange(4, 1, vals.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(4, 3, vals.length, 1 + grid.length).setNumberFormat('0.00%');
    sh.getRange(4, 3, vals.length, 1).setFontWeight('bold');
    sh.getRange(4, head.length, vals.length, 1).setNumberFormat('$ #,##0.00').setFontWeight('bold');
    pintar_(sh.getRange(4, 4, vals.length, grid.length));
  }
  bordes_(sh.getRange(1, 1, 1, 2));
  bordes_(sh.getRange(3, 1, rows.length + 1, head.length));
  sh.setColumnWidth(head.length + 2, 110);
  bloqueEstado_(sh, 1, head.length + 2);

  // --- Datos del gráfico: X=Vencimiento, Fx BE, MEP hoy  (columnas V:X)
  sh.getRange(1, 22, 41, 3).clearContent();
  sh.getRange(1, 22, 1, 3).setValues([['Vencimiento', 'Fx BE', 'MEP hoy']]);
  if (rows.length) {
    sh.getRange(2, 22, rows.length, 3).setValues(rows.map(x => [fecha_(x.venc), x.fxbe, ctx.mep]));
    sh.getRange(2, 22, rows.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 23, rows.length, 2).setNumberFormat('$ #,##0.00');
  }
  if (!sh.getCharts().length) {
    const ch = sh.newChart().setChartType(Charts.ChartType.LINE)
      .addRange(sh.getRange('V1:X41')).setNumHeaders(1)
      .setPosition(rows.length + 8, 1, 0, 0)
      .setOption('title', 'Carry trade — Dólar breakeven por vencimiento')
      .setOption('pointSize', 6)
      .setOption('vAxis', { format: '$ #,##0' })
      .setOption('series', { 1: { lineDashStyle: [4, 4] } })
      .setOption('width', 900).setOption('height', 380)
      .build();
    sh.insertChart(ch);
  }
}

function hojaPrecios_(ss, ctx) {
  const sh = hoja_(ss, 'Precios');
  const tks = []
    .concat(lista_(ctx.cfg['Tickers GD']), lista_(ctx.cfg['Tickers AL']), lista_(ctx.cfg['Tickers matriz'])).map(t => t + 'D')
    .concat(['AL30', 'AL30D'], ctx.lecaps.map(l => l.ticker));
  const uniq = tks.filter((t, i) => tks.indexOf(t) === i);
  const out = [['Símbolo', 'Liquidación', 'Precio', 'Variación', 'Estado']];
  uniq.forEach(t => ['1', '2'].forEach(st => {
    const px = ctx.p(t, st);
    out.push([t, st === '1' ? 'CI' : '24hs', px ? px.precio : '', px ? px.var : '', px ? 'ok' : 'sin dato']);
  }));
  sh.clear();
  sh.getRange(1, 1, out.length, 5).setValues(out);
  sh.getRange(2, 4, out.length - 1, 1).setNumberFormat('0.00%');
  estiloHeader_(sh.getRange(1, 1, 1, 5));
  sh.setFrozenRows(1);
}

/* =============================================================================
 * 5. LECTURA DE HOJAS DE SOPORTE
 * ========================================================================== */
function leerConfig_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Config');
  if (!sh) throw new Error('Falta la hoja Config: corré configurarInicial().');
  const out = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) out[r[0]] = r[1]; });
  return out;
}

function setConfig_(key, val) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Config');
  const keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => r[0]);
  const i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 1, 2).setValue(val);
}

function leerFeriados_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Feriados');
  const s = new Set();
  if (!sh || sh.getLastRow() < 2) return s;
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => { if (r[0]) s.add(iso_(r[0])); });
  return s;
}

function leerFlujos_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Flujos');
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(r => {
    if (!r[0] || !r[1]) return;
    const tk = String(r[0]).trim();
    (out[tk] = out[tk] || []).push({ fecha: iso_(r[1]), renta: Number(r[2]) || 0, amort: Number(r[3]) || 0, total: Number(r[4]) || 0 });
  });
  Object.keys(out).forEach(k => out[k].sort((a, b) => a.fecha < b.fecha ? -1 : 1));
  return out;
}

function leerLecapsVF_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Lecaps_VF');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues()
    .filter(r => r[0] && r[1])
    .map(r => {
      const venc = iso_(r[1]), tem = Number(r[6]) || 0;
      const vf = (r[5] && tem > 0) ? vfLecap_(iso_(r[5]), venc, tem) : (Number(r[2]) || 0);
      return { ticker: String(r[0]).trim(), venc: venc, vf: vf, tipo: String(r[3] || 'LECAP').trim().toUpperCase() };
    });
}

/** VF de LECAP/BONCAP por 100 VN: 100 × (1+TEM)^n, n = meses enteros desde emisión + días restantes/30. */
function vfLecap_(emision, venc, temPct) {
  let k = 0;
  while (sumarMeses_(emision, k + 1) <= venc) k++;
  const n = k + (diaNum_(venc) - diaNum_(sumarMeses_(emision, k))) / 30;
  return 100 * Math.pow(1 + temPct / 100, n);
}

/* =============================================================================
 * 6. UTILIDADES
 * ========================================================================== */
function hoja_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

function estiloHeader_(rg, bg) {
  rg.setBackground(bg || C.header).setFontColor(C.headerTxt).setFontWeight('bold').setHorizontalAlignment('center');
}

/**
 * Zona horaria argentina (para NOW()) y recálculo cada minuto (para que "Mercado abierto/cerrado"
 * cambie a horario). Va por la API de Sheets; si falla, no corta la actualización.
 */
function configurarRecalculo_(ss) {
  try {
    if (ss.getSpreadsheetTimeZone() !== TZ) ss.setSpreadsheetTimeZone(TZ);
    Sheets.Spreadsheets.batchUpdate({ requests: [{ updateSpreadsheetProperties: {
      properties: { autoRecalc: 'MINUTE' }, fields: 'autoRecalc' } }] }, ss.getId());
  } catch (e) { Logger.log('Recálculo por minuto: ' + e); }
}

/**
 * Bloque de 2×2: Fuente (BYMA · demora 20 min) y estado del mercado. El estado es una fórmula
 * con NOW(), así cambia sola aunque el script no corra (rueda BYMA 11:00–17:00, días hábiles
 * sin feriados de la hoja Feriados).
 */
function bloqueEstado_(sh, row, col) {
  sh.getRange(row, col, 2, 2).setValues([['Fuente', 'BYMA · demora 20 min'], ['Mercado', '']]);
  const cel = sh.getRange(row + 1, col + 1);
  cel.setFormula('=IF(AND(WEEKDAY(NOW(),2)<=5,SUMPRODUCT(--(INT(Feriados!A2:A300)=INT(NOW())))=0,' +
    'MOD(NOW(),1)>=TIME(11,0,0),MOD(NOW(),1)<TIME(17,0,0)),"Mercado abierto","Mercado cerrado")');
  estiloHeader_(sh.getRange(row, col, 2, 1), C.sub);
  sh.getRange(row, col + 1, 2, 1).setHorizontalAlignment('center').setFontWeight('bold');
  bordes_(sh.getRange(row, col, 2, 2));
  sh.setColumnWidth(col + 1, 175);
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Mercado abierto').setFontColor('#1E8E3E').setRanges([cel]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Mercado cerrado').setFontColor('#D93025').setRanges([cel]).build()
  ]));
}

/** Todos los bordes (finos, negros) a cada celda del rango. */
function bordes_(rg) {
  rg.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Colorea por valor: verde si > 0, rojo si < 0, más intenso cuanto más grande en relación al
 * máximo/mínimo del rango. Escala con raíz cuadrada y piso del 30% para que hasta un valor chico
 * se note de qué lado está.
 */
function pintar_(rg) {
  const vals = rg.getValues();
  let maxP = 0, minN = 0;
  vals.forEach(r => r.forEach(v => { if (typeof v === 'number') { maxP = Math.max(maxP, v); minN = Math.min(minN, v); } }));
  rg.setBackgrounds(vals.map(r => r.map(v => {
    if (typeof v !== 'number' || v === 0) return '#FFFFFF';
    const t = v > 0 ? Math.sqrt(v / maxP) : Math.sqrt(v / minN);
    return mezcla_('#FFFFFF', v > 0 ? C.verdeFuerte : C.rojoFuerte, 0.30 + 0.70 * t);
  })));
}

function mezcla_(c1, c2, k) {
  const h = c => [1, 3, 5].map(i => parseInt(c.substr(i, 2), 16));
  const a = h(c1), b = h(c2);
  return '#' + a.map((v, i) => ('0' + Math.round(v + (b[i] - v) * k).toString(16)).slice(-2)).join('');
}

function gradiente_(rg, cMin, cMid, cMax) {
  return SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(cMin)
    .setGradientMidpointWithValue(cMid, SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpoint(cMax)
    .setRanges([rg]).build();
}

function lista_(v) { return String(v || '').split(';').map(s => s.trim()).filter(Boolean); }
function num_(s) { return parseFloat(String(s).replace(',', '.')); }
function nz_(v) { return v === null || v === undefined || isNaN(v) ? '' : v; }
function round_(v, d) { const k = Math.pow(10, d); return Math.round(v * k) / k; }

function hoyISO_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

/** Date o string → 'yyyy-MM-dd'. */
function iso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s.slice(0, 10);
}

/** 'yyyy-MM-dd' → Date al mediodía (evita corrimientos de zona horaria al escribir). */
function fecha_(isoStr) { const p = isoStr.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2], 12); }
function fmt_(isoStr) { const p = isoStr.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }
function diaNum_(isoStr) { const p = isoStr.split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]) / 864e5; }

function sumarMeses_(isoStr, n) {
  const p = isoStr.split('-').map(Number);
  const t = p[0] * 12 + (p[1] - 1) + n;
  const y = Math.floor(t / 12), m = t % 12 + 1;
  const ult = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + Math.min(p[2], ult)).slice(-2);
}

function proximoHabil_(isoStr, feriados) {
  let n = diaNum_(isoStr);
  for (;;) {
    n += 1;
    const d = new Date(n * 864e5);
    const s = d.toISOString().slice(0, 10);
    const dow = d.getUTCDay();
    if (dow !== 0 && dow !== 6 && !feriados.has(s)) return s;
  }
}
