/**
 * MDV COACH – Backend v9 (permisos + informes estables)
 *
 * Endpoints soportados (action):
 * - ping, diag
 * - get_swimmers, get_swimmer_profile, get_swimmer_marks_with_context
 * - get_swimmer_permissions, set_swimmer_permissions
 * - add_mark (usa permisos cuando actor_role = "swimmer")
 *
 * El script es tolerante a variaciones de encabezados y siempre agrega
 * las columnas canónicas que necesita.
 */

const SHEETS_DEFAULT = {
  swimmers: 'swimmers',
  marks: 'Marks',
  config: 'config',
};

const DEFAULT_CONV_SCM_TO_LCM = 1.0103;
const DEFAULT_PERMISSIONS = { allow_marks_edit: false, allow_marks_delete: false };

const CANONICAL_SWIMMER_COLS = [
  'coach_id', 'swimmer_id', 'nombre', 'fecha_nac', 'genero',
  'altura_cm', 'peso_kg', 'fc_reposo', 'created_at', 'updated_at',
  'allow_marks_edit', 'allow_marks_delete',
];

const CANONICAL_MARK_COLS = [
  'coach_id', 'swimmer_id', 'fecha', 'season_year', 'age_chip', 'tipo_toma',
  'curso', 'estilo', 'distancia_m', 'carril', 'tiempo_str', 'tiempo_s',
  'created_at', 'lugar_evento', 'tiempo_raw', 'updated_at', 'mark_id',
  'edited_by', 'deleted_at', 'source',
];

const HEADER_ALIASES = {
  coach_id: ['coach_id', 'coach', 'coachid', 'id_entrenador', 'idcoach'],
  swimmer_id: ['swimmer_id', 'swimmerid', 'id_nadador', 'nadador_id', 'idnadador'],
  nombre: ['nombre', 'name', 'nadador', 'swimmer_name'],
  fecha_nac: ['fecha_nac', 'fecha_de_nacimiento', 'nacimiento', 'birthdate', 'nac_date'],
  genero: ['genero', 'género', 'sexo', 'gender'],
  altura_cm: ['altura_cm', 'altura', 'height_cm'],
  peso_kg: ['peso_kg', 'peso', 'weight_kg'],
  fc_reposo: ['fc_reposo', 'fc', 'fc_rest', 'rest_hr', 'fc_resting'],
  created_at: ['created_at', 'creado_en', 'timestamp'],
  updated_at: ['updated_at', 'actualizado_en', 'last_update'],
  allow_marks_edit: ['allow_marks_edit', 'can_edit_marks', 'editar_marcas'],
  allow_marks_delete: ['allow_marks_delete', 'can_delete_marks', 'borrar_marcas'],

  fecha: ['fecha', 'fecha_evento', 'fecha_de_toma', 'date'],
  season_year: ['season_year', 'year', 'anio', 'año', 'ano'],
  age_chip: ['age_chip', 'edad_chip', 'age', 'edad_marca'],
  tipo_toma: ['tipo_toma', 'tipo', 'modalidad', 'tipo_de_toma'],
  curso: ['curso', 'pool', 'piscina', 'tipo_pool'],
  estilo: ['estilo', 'trazo', 'stroke'],
  distancia_m: ['distancia_m', 'distancia', 'distancia_mts', 'distancia_metros'],
  carril: ['carril', 'lane'],
  tiempo_raw: ['tiempo_raw', 'tiempo', 'time', 'marca', 'raw_time'],
  tiempo_str: ['tiempo_str', 'time_str', 'time_text', 'time_string'],
  tiempo_s: ['tiempo_s', 'time_s', 'segundos', 'seconds'],
  lugar_evento: ['lugar_evento', 'lugar', 'evento', 'ubicacion', 'ubicación'],
  mark_id: ['mark_id', 'id_marca', 'idmark'],
  edited_by: ['edited_by', 'editado_por', 'editor'],
  deleted_at: ['deleted_at', 'borrado_en', 'eliminado_en'],
  source: ['source', 'fuente', 'origen'],
};

function doGet(e) { return handleRequest_(e); }
function doPost(e) { return handleRequest_(e); }

function handleRequest_(e) {
  try {
    const params = getParams_(e);
    const action = String(params.action || params.accion || '').trim().toLowerCase();
    if (!action) return json_({ status: 'error', error: 'Missing action' });

    switch (action) {
      case 'ping':
        return json_({ status: 'ok', ts: new Date().toISOString() });
      case 'diag':
        return json_(handleDiag_(params));
      case 'get_swimmers':
        return json_(handleGetSwimmers_(params));
      case 'get_swimmer_profile':
        return json_(handleGetSwimmerProfile_(params));
      case 'get_swimmer_marks_with_context':
        return json_(handleGetSwimmerMarksWithContext_(params));
      case 'get_swimmer_permissions':
      case 'get_permissions':
        return json_(handleGetSwimmerPermissions_(params));
      case 'set_swimmer_permissions':
      case 'set_permissions':
        return json_(handleSetSwimmerPermissions_(params));
      case 'add_mark':
        return json_(handleAddMark_(params));
      default:
        return json_({ status: 'error', error: `Unknown action: ${action}` });
    }
  } catch (err) {
    return json_({ status: 'error', error: String(err && err.message ? err.message : err) });
  }
}

function getParams_(e) {
  const out = {};
  if (!e) return out;
  if (e.parameter) Object.keys(e.parameter).forEach(k => out[k] = e.parameter[k]);
  const postData = e.postData;
  if (postData && postData.contents) {
    const c = postData.contents;
    const ct = String(postData.type || '').toLowerCase();
    if (ct.indexOf('application/json') >= 0) {
      try { Object.assign(out, JSON.parse(c)); } catch (_) {}
    } else {
      c.split('&').forEach(pair => {
        const [k, v] = pair.split('=');
        if (!k) return;
        out[decodeURIComponent(k)] = v ? decodeURIComponent(v) : '';
      });
    }
  }
  return out;
}

function json_(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  try {
    output.setHeader('Access-Control-Allow-Origin', '*');
    output.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    output.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  } catch (_) {}
  return output;
}

function getActiveSS_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getConfig_(ss) {
  const cfg = {
    swimmers_sheet: SHEETS_DEFAULT.swimmers,
    marks_sheet: SHEETS_DEFAULT.marks,
    config_sheet: SHEETS_DEFAULT.config,
    conv_scm_to_lcm: DEFAULT_CONV_SCM_TO_LCM,
  };
  const props = PropertiesService.getScriptProperties().getProperties() || {};
  if (props.SWIMMERS_SHEET) cfg.swimmers_sheet = props.SWIMMERS_SHEET;
  if (props.MARKS_SHEET) cfg.marks_sheet = props.MARKS_SHEET;
  if (props.CONV_SCM_TO_LCM) cfg.conv_scm_to_lcm = toNumber_(props.CONV_SCM_TO_LCM) || cfg.conv_scm_to_lcm;

  const cfgSheet = ss.getSheetByName(cfg.config_sheet);
  if (cfgSheet) {
    const values = cfgSheet.getDataRange().getValues();
    values.forEach(row => {
      const key = normalizeHeader_(row[0]);
      const val = row[1];
      if (key === 'swimmers_sheet' || key === 'hoja_nadadores') cfg.swimmers_sheet = String(val || '').trim() || cfg.swimmers_sheet;
      if (key === 'marks_sheet' || key === 'hoja_marcas') cfg.marks_sheet = String(val || '').trim() || cfg.marks_sheet;
      if (key === 'conv_scm_to_lcm' || key === 'factor_conversion') {
        const n = toNumber_(val); if (n) cfg.conv_scm_to_lcm = n;
      }
    });
  }
  return cfg;
}

function normalizeHeader_(h) {
  return String(h == null ? '' : h)
    .replace(/["']/g, '')
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function buildAliasLookup_(aliasesObj) {
  const lookup = {};
  Object.keys(aliasesObj).forEach(canon => {
    (aliasesObj[canon] || []).forEach(alias => lookup[normalizeHeader_(alias)] = canon);
  });
  return lookup;
}

function buildHeaderMap_(headers, aliasLookup) {
  const map = {};
  headers.forEach((h, idx) => {
    const canon = aliasLookup[normalizeHeader_(h)];
    if (canon != null && map[canon] == null) map[canon] = idx;
  });
  return map;
}

function ensureColumns_(sheet, canonicalCols) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  const headers = headerRange.getValues()[0].map(h => String(h || '').trim());
  const normSet = new Set(headers.map(normalizeHeader_));
  let changed = false;
  canonicalCols.forEach(col => {
    if (!normSet.has(normalizeHeader_(col))) {
      headers.push(col);
      normSet.add(normalizeHeader_(col));
      changed = true;
    }
  });
  if (changed) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return headers;
}

function toNumber_(v) { const n = Number(v); return Number.isFinite(n) ? n : null; }
function parseBool_(v) {
  if (v == null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  const s = String(v).trim().toLowerCase();
  return ['true','1','si','sí','yes','y','on'].indexOf(s) >= 0;
}

function parseISODate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function pad2_(n) { return String(n).padStart(2, '0'); }
function formatSeconds_(sec) {
  if (!Number.isFinite(sec)) return '';
  const m = Math.floor(sec / 60);
  const s = sec - m * 60;
  const intS = Math.floor(s);
  const hundredths = Math.round((s - intS) * 100);
  return `${m}:${pad2_(intS)}.${String(hundredths).padStart(2, '0')}`;
}

function parseTimeToSeconds_(t) {
  if (t == null || t === '') return null;
  if (typeof t === 'number') return Number.isFinite(t) ? t : null;
  const s = String(t).trim();
  const m1 = s.match(/^(\d+):(\d{1,2})(?:\.(\d{1,2}))?$/);
  if (m1) {
    const mm = Number(m1[1]);
    const ss = Number(m1[2]);
    const hh = m1[3] ? Number(String(m1[3]).padEnd(2, '0')) : 0;
    if ([mm, ss, hh].some(x => !Number.isFinite(x))) return null;
    return mm * 60 + ss + hh / 100;
  }
  const m2 = s.match(/^(\d+)(?:\.(\d{1,2}))?$/);
  if (m2) {
    const ss = Number(m2[1]);
    const hh = m2[2] ? Number(String(m2[2]).padEnd(2, '0')) : 0;
    if ([ss, hh].some(x => !Number.isFinite(x))) return null;
    return ss + hh / 100;
  }
  return null;
}

function calcAgeAt_(birthDate, atDate) {
  const b = parseISODate_(birthDate);
  const a = parseISODate_(atDate) || new Date();
  if (!b) return null;
  let age = a.getFullYear() - b.getFullYear();
  const m = a.getMonth() - b.getMonth();
  if (m < 0 || (m === 0 && a.getDate() < b.getDate())) age--;
  return age;
}

function seasonFromDate_(d) { const dt = parseISODate_(d); return dt ? dt.getFullYear() : null; }

function readSheetAsObjects_(sheet, canonicalCols, aliases) {
  const headers = ensureColumns_(sheet, canonicalCols);
  const aliasLookup = buildAliasLookup_(aliases);
  const headerMap = buildHeaderMap_(headers, aliasLookup);
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), headers.length).getValues();
  const rows = values.map(r => {
    const obj = {};
    Object.keys(headerMap).forEach(key => obj[key] = r[headerMap[key]]);
    return obj;
  });
  return { headers, headerMap, rows };
}

function handleDiag_() {
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);
  return { status: 'ok', spreadsheet: ss.getId(), sheets: cfg };
}

function handleGetSwimmers_(params) {
  const coachId = String(params.coach_id || params.coach || '').trim();
  if (!coachId) return { status: 'error', error: 'coach_id requerido' };
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);
  const sheet = ss.getSheetByName(cfg.swimmers_sheet);
  if (!sheet) return { status: 'error', error: `No se encontró la hoja ${cfg.swimmers_sheet}` };
  const { rows } = readSheetAsObjects_(sheet, CANONICAL_SWIMMER_COLS, HEADER_ALIASES);
  const swimmers = rows.filter(r => String(r.coach_id || '').trim() === coachId)
    .map(r => ({
      ...DEFAULT_PERMISSIONS,
      ...r,
      allow_marks_edit: parseBool_(r.allow_marks_edit),
      allow_marks_delete: parseBool_(r.allow_marks_delete),
    }));
  return { status: 'ok', swimmers };
}

function findSwimmerRow_(sheet, headerMap, coachId, swimmerId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const c = headerMap.coach_id != null ? String(row[headerMap.coach_id]).trim() : '';
    const s = headerMap.swimmer_id != null ? String(row[headerMap.swimmer_id]).trim() : '';
    if (c === coachId && s === swimmerId) return { rowIndex: i + 1, rowValues: row };
  }
  return null;
}

function handleGetSwimmerProfile_(params) {
  const coachId = String(params.coach_id || '').trim();
  const swimmerId = String(params.swimmer_id || '').trim();
  if (!coachId || !swimmerId) return { status: 'error', error: 'coach_id y swimmer_id son requeridos' };
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);
  const sheet = ss.getSheetByName(cfg.swimmers_sheet);
  if (!sheet) return { status: 'error', error: `No se encontró la hoja ${cfg.swimmers_sheet}` };
  const { headers, headerMap, rows } = readSheetAsObjects_(sheet, CANONICAL_SWIMMER_COLS, HEADER_ALIASES);
  const found = rows.find(r => String(r.coach_id || '').trim() === coachId && String(r.swimmer_id || '').trim() === swimmerId);
  if (!found) return { status: 'error', error: 'Nadador no encontrado' };
  const profile = {
    ...DEFAULT_PERMISSIONS,
    ...found,
    allow_marks_edit: parseBool_(found.allow_marks_edit),
    allow_marks_delete: parseBool_(found.allow_marks_delete),
  };
  const at = new Date();
  return { status: 'ok', profile, headers, last_checked: at.toISOString() };
}

function handleGetSwimmerPermissions_(params) {
  const res = handleGetSwimmerProfile_(params);
  if (res.status !== 'ok') return res;
  return { status: 'ok', permissions: {
    allow_marks_edit: !!res.profile.allow_marks_edit,
    allow_marks_delete: !!res.profile.allow_marks_delete,
  }};
}

function handleSetSwimmerPermissions_(params) {
  const coachId = String(params.coach_id || '').trim();
  const swimmerId = String(params.swimmer_id || '').trim();
  const allowEdit = params.allow_marks_edit;
  const allowDelete = params.allow_marks_delete;
  if (!coachId || !swimmerId) return { status: 'error', error: 'coach_id y swimmer_id son requeridos' };
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);
  const sheet = ss.getSheetByName(cfg.swimmers_sheet);
  if (!sheet) return { status: 'error', error: `No se encontró la hoja ${cfg.swimmers_sheet}` };
  const headers = ensureColumns_(sheet, CANONICAL_SWIMMER_COLS);
  const aliasLookup = buildAliasLookup_(HEADER_ALIASES);
  const headerMap = buildHeaderMap_(headers, aliasLookup);
  const rowInfo = findSwimmerRow_(sheet, headerMap, coachId, swimmerId);
  if (!rowInfo) return { status: 'error', error: 'Nadador no encontrado' };

  const rowVals = rowInfo.rowValues;
  const nowIso = new Date().toISOString();
  if (headerMap.allow_marks_edit != null && allowEdit != null) rowVals[headerMap.allow_marks_edit] = parseBool_(allowEdit);
  if (headerMap.allow_marks_delete != null && allowDelete != null) rowVals[headerMap.allow_marks_delete] = parseBool_(allowDelete);
  if (headerMap.updated_at != null) rowVals[headerMap.updated_at] = nowIso;
  sheet.getRange(rowInfo.rowIndex, 1, 1, headers.length).setValues([rowVals]);

  return { status: 'ok', permissions: {
    allow_marks_edit: parseBool_(allowEdit),
    allow_marks_delete: parseBool_(allowDelete),
  }, updated_at: nowIso };
}

function handleGetSwimmerMarksWithContext_(params) {
  const coachId = String(params.coach_id || '').trim();
  const swimmerId = String(params.swimmer_id || '').trim();
  if (!coachId || !swimmerId) return { status: 'error', error: 'coach_id y swimmer_id son requeridos' };
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);
  const swimmersSheet = ss.getSheetByName(cfg.swimmers_sheet);
  const marksSheet = ss.getSheetByName(cfg.marks_sheet);
  if (!swimmersSheet || !marksSheet) return { status: 'error', error: 'Faltan hojas swimmers o marks' };

  const { rows: swimmers } = readSheetAsObjects_(swimmersSheet, CANONICAL_SWIMMER_COLS, HEADER_ALIASES);
  const swimmer = swimmers.find(r => String(r.coach_id || '').trim() === coachId && String(r.swimmer_id || '').trim() === swimmerId);
  if (!swimmer) return { status: 'error', error: 'Nadador no encontrado' };

  const birth = swimmer.fecha_nac;
  const { rows: marks, headerMap } = readSheetAsObjects_(marksSheet, CANONICAL_MARK_COLS, HEADER_ALIASES);
  const filtered = marks.filter(m => String(m.coach_id || '').trim() === coachId && String(m.swimmer_id || '').trim() === swimmerId)
    .map(m => {
      const markDate = m.fecha || m.created_at;
      const season = m.season_year || seasonFromDate_(markDate);
      const ageChip = m.age_chip || calcAgeAt_(birth, markDate);
      const tiempoS = m.tiempo_s || parseTimeToSeconds_(m.tiempo_raw || m.tiempo_str);
      const tiempoStr = m.tiempo_str || formatSeconds_(tiempoS);
      return { ...m, season_year: season, age_chip: ageChip, tiempo_s: tiempoS, tiempo_str: tiempoStr };
    });

  const permissions = {
    allow_marks_edit: parseBool_(swimmer.allow_marks_edit),
    allow_marks_delete: parseBool_(swimmer.allow_marks_delete),
  };
  const summary = {
    total_marks: filtered.length,
    last_mark_at: filtered.length ? filtered.map(m => m.fecha || m.created_at).sort().slice(-1)[0] : null,
  };

  return { status: 'ok', profile: swimmer, permissions, marks: filtered, summary, meta: { marks_sheet: cfg.marks_sheet } };
}

function handleAddMark_(params) {
  const coachId = String(params.coach_id || '').trim();
  const swimmerId = String(params.swimmer_id || '').trim();
  if (!coachId || !swimmerId) return { status: 'error', error: 'coach_id y swimmer_id son requeridos' };
  const actorRole = String(params.actor_role || '').toLowerCase();

  const ss = getActiveSS_();
  const cfg = getConfig_(ss);
  const swimmersSheet = ss.getSheetByName(cfg.swimmers_sheet);
  const marksSheet = ss.getSheetByName(cfg.marks_sheet);
  if (!swimmersSheet || !marksSheet) return { status: 'error', error: 'Faltan hojas swimmers o marks' };

  const swimmersData = readSheetAsObjects_(swimmersSheet, CANONICAL_SWIMMER_COLS, HEADER_ALIASES);
  const swimmer = swimmersData.rows.find(r => String(r.coach_id || '').trim() === coachId && String(r.swimmer_id || '').trim() === swimmerId);
  if (!swimmer) return { status: 'error', error: 'Nadador no encontrado' };
  const perms = {
    allow_marks_edit: parseBool_(swimmer.allow_marks_edit),
    allow_marks_delete: parseBool_(swimmer.allow_marks_delete),
  };
  if (actorRole === 'swimmer' && !perms.allow_marks_edit) return { status: 'error', error: 'El nadador no puede cargar marcas (permiso denegado)' };

  const headers = ensureColumns_(marksSheet, CANONICAL_MARK_COLS);
  const aliasLookup = buildAliasLookup_(HEADER_ALIASES);
  const headerMap = buildHeaderMap_(headers, aliasLookup);

  const fecha = params.fecha || params.fecha_evento || new Date();
  const season = params.season_year || seasonFromDate_(fecha);
  const ageChip = params.age_chip || calcAgeAt_(swimmer.fecha_nac, fecha);
  const tiempoS = parseTimeToSeconds_(params.tiempo_s || params.tiempo_str || params.tiempo_raw);
  const tiempoStr = params.tiempo_str || formatSeconds_(tiempoS);
  const markId = params.mark_id || `mark_${Date.now()}`;
  const nowIso = new Date().toISOString();

  const row = new Array(headers.length).fill('');
  const set = (key, val) => { if (headerMap[key] != null) row[headerMap[key]] = val; };
  set('coach_id', coachId);
  set('swimmer_id', swimmerId);
  set('fecha', fecha);
  set('season_year', season);
  set('age_chip', ageChip);
  set('tipo_toma', params.tipo_toma || params.tipo || 'Oficial');
  set('curso', params.curso || params.pool || 'LCM');
  set('estilo', params.estilo || 'Libre');
  set('distancia_m', params.distancia_m || params.distancia || '');
  set('carril', params.carril || '');
  set('tiempo_raw', params.tiempo_raw || params.tiempo_str || params.tiempo_s || '');
  set('tiempo_s', tiempoS);
  set('tiempo_str', tiempoStr);
  set('created_at', nowIso);
  set('updated_at', nowIso);
  set('lugar_evento', params.lugar_evento || '');
  set('mark_id', markId);
  set('edited_by', actorRole || 'coach');
  set('source', params.source || 'dashboard');

  marksSheet.appendRow(row);
  return { status: 'ok', mark_id: markId, tiempo_str: tiempoStr, age_chip: ageChip };
}
