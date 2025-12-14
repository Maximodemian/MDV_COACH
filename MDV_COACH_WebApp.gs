// MDV COACH WebApp (Planilla COACH)
// Versión optimizada para evitar timeouts en el dashboard del entrenador.
// Referencia: MDV_Nadador_Dashboard_v6.html, MDV_Dashboard_Entrenador_v5.html,
// y planilla "MDV_COACH_MAURO_(REFERENCIA PARA CODEX).xlsx".


const SHEETS_DEFAULT = {
  swimmers: "nadadores",
  marks: "marcas",
  config: "config",
};


const DEFAULT_CONV_SCM_TO_LCM = 1.02;
const CACHE_TTL_SECONDS = 300; // 5 minutos: reduce lecturas repetidas de la planilla


const POOL_COURSE_SET = new Set(["SCM", "LCM", "SCY"]);
const TAKE_TYPE_SET = new Set(["COMPETENCIA", "ENTRENAMIENTO", "TEST", "CONTROL"]);


const CANONICAL_MARK_COLS = [
  "coach_id",
  "swimmer_id",
  "fecha",
  "lugar_evento",
  "tipo_toma",
  "curso",
  "estilo",
  "distancia_m",
  "carril",
  "tiempo_raw",
  "tiempo_str",
  "tiempo_s",
  "created_at",
  "client_mark_id",
  "edad_ref",
  "categoria_ref",
];


const CANONICAL_SWIMMER_COLS = [
  "coach_id",
  "swimmer_id",
  "nombre",
  "fecha_nac",
  "genero",
  "altura_cm",
  "peso_kg",
  "fc_reposo",
];


const HEADER_ALIASES = {
  coach_id: ["coach_id", "coach", "coachid", "id_entrenador", "idcoach", "coach_id "],
  swimmer_id: ["swimmer_id", "swimmerid", "id_nadador", "nadador_id", "idnadador", "swimmer_id "],


  fecha: ["fecha", "fecha_evento", "fecha_de_toma", "date", "dia", "día"],
  lugar_evento: ["lugar_evento", "lugar", "evento", "ubicacion_del_evento", "ubicación_del_evento", "ubicacion", "ubicación"],


  tipo_toma: ["tipo_toma", "tipo", "modalidad", "tipo_de_toma", "tipo_de_toque"],
  curso: ["curso", "pool", "piscina", "tipo_pool", "piscina_curso"],


  estilo: ["estilo", "trazo", "stroke"],
  distancia_m: ["distancia_m", "distancia", "distancia_mts", "distancia_metros"],
  carril: ["carril", "lane"],


  tiempo_raw: ["tiempo_raw", "tiempo", "time", "marca", "raw_time"],
  tiempo_str: ["tiempo_str", "time_str", "time_text", "time_string"],
  tiempo_s: ["tiempo_s", "time_s", "segundos", "seconds"],


  created_at: ["created_at", "creado_en", "creado", "timestamp", "marca_de_tiempo"],
  client_mark_id: ["client_mark_id", "mark_id", "id_marca_cliente"],


  edad_ref: ["edad_ref", "edad_en_marca", "edad_marca"],
  categoria_ref: ["categoria_ref", "categoria_en_marca", "cat_ref"],
};


const SWIMMER_HEADER_ALIASES = {
  coach_id: HEADER_ALIASES.coach_id,
  swimmer_id: HEADER_ALIASES.swimmer_id,
  nombre: ["nombre", "name", "nadador", "swimmer_name"],
  fecha_nac: ["fecha_nac", "fecha_de_nacimiento", "nacimiento", "birthdate", "nac_date"],
  genero: ["genero", "género", "sexo", "gender"],
  altura_cm: ["altura_cm", "altura", "height_cm"],
  peso_kg: ["peso_kg", "peso", "weight_kg"],
  fc_reposo: ["fc_reposo", "fc", "fc_rest", "rest_hr", "hora_descanso_bpm", "fc_resting"],
};


function doGet(e) { return handleRequest_(e); }
function doPost(e) { return handleRequest_(e); }


function handleRequest_(e) {
  try {
    const params = extractParams_(e);
    const actionRaw = (params.action || params.accion || "").toString().trim();
    const action = actionRaw.toLowerCase();


    switch (action) {
      case "ping":
        return jsonResponse({ status: "ok", message: "alive" });
      case "get_swimmers":
      case "list_swimmers":
      case "lista_nadadores":
        return jsonResponse(handleGetSwimmers_(params));
      case "get_swimmer_profile":
      case "get_profile":
        return jsonResponse(handleGetSwimmerProfile_(params));
      case "get_swimmer_marks_with_context":
      case "get_marks_with_context":
      case "get_marks":
        return jsonResponse(handleGetSwimmerMarksWithContext_(params));
      case "add_mark":
      case "addmark":
        return jsonResponse(handleAddMark_(params));
      case "get_swimmer_permissions":
        return jsonResponse(handleGetSwimmerPermissions_(params));
      case "set_swimmer_permissions":
        return jsonResponse(handleSetSwimmerPermissions_(params));
      case "diag":
        return jsonResponse(handleDiag_(params));
      default:
        return jsonResponse({ ok: false, error: { code: "UNKNOWN_ACTION", message: `Unknown action: ${actionRaw}` } });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: { code: "UNEXPECTED_ERROR", message: (err && err.message) ? err.message : String(err) } });
  }
}


function extractParams_(e) {
  const out = {};
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(k => out[k] = e.parameter[k]);
  }
  if (e && e.postData && e.postData.contents) {
    const ct = (e.postData.type || "").toLowerCase();
    const body = e.postData.contents;
    if (ct.indexOf("application/json") >= 0) {
      try {
        const obj = JSON.parse(body);
        if (obj && typeof obj === "object") Object.keys(obj).forEach(k => out[k] = obj[k]);
      } catch (_) {}
    } else {
      body.split("&").forEach(pair => {
        const idx = pair.indexOf("=");
        if (idx < 0) return;
        const k = decodeURIComponent(pair.slice(0, idx));
        const v = decodeURIComponent(pair.slice(idx + 1));
        if (k) out[k] = v;
      });
    }
  }
  return out;
}


function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


function jsonResponse(obj) {
  return json_(obj);
}


/** ---------------- CONFIG / SHEETS ---------------- */


function getSpreadsheet_(params) {
  if (params && params.spreadsheet_id) {
    try {
      return SpreadsheetApp.openById(params.spreadsheet_id);
    } catch (err) {
      throw new Error(`No se pudo abrir el Spreadsheet ${params.spreadsheet_id}: ${err.message}`);
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}


function getSheetByNameStrict(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`No se encontró la hoja "${name}".`);
  return sh;
}


function getConfig_(ss) {
  const cfg = {
    swimmers_sheet: SHEETS_DEFAULT.swimmers,
    marks_sheet: SHEETS_DEFAULT.marks,
    config_sheet: SHEETS_DEFAULT.config,
    conv_scm_to_lcm: DEFAULT_CONV_SCM_TO_LCM,
    standards_webapp_url: "",
    usa_cycle: "2024-2028",
    cadda_year: "2025",
  };


  const props = PropertiesService.getScriptProperties().getProperties() || {};
  if (props.CONV_SCM_TO_LCM) cfg.conv_scm_to_lcm = toNumber_(props.CONV_SCM_TO_LCM) || cfg.conv_scm_to_lcm;
  if (props.STANDARDS_WEBAPP_URL) cfg.standards_webapp_url = String(props.STANDARDS_WEBAPP_URL).trim();
  if (props.USA_CYCLE) cfg.usa_cycle = String(props.USA_CYCLE).trim();
  if (props.CADDA_YEAR) cfg.cadda_year = String(props.CADDA_YEAR).trim();


  const sheet = ss.getSheetByName(cfg.config_sheet);
  if (sheet) {
    const values = sheet.getDataRange().getValues();
    for (let i = 0; i < values.length; i++) {
      const key = normalizeHeader_(values[i][0]);
      const val = values[i][1];
      if (!key) continue;


      if (key === "marks_sheet" || key === "hoja_marcas") cfg.marks_sheet = String(val).trim() || cfg.marks_sheet;
      if (key === "swimmers_sheet" || key === "hoja_nadadores") cfg.swimmers_sheet = String(val).trim() || cfg.swimmers_sheet;


      if (key === "conv_scm_to_lcm" || key === "conversion_factor" || key === "factor_conversion") {
        const num = toNumber_(val);
        if (num) cfg.conv_scm_to_lcm = num;
      }


      if (key === "standards_webapp_url" || key === "standards_url") {
        if (val) cfg.standards_webapp_url = String(val).trim();
      }


      if (key === "usa_cycle" || key === "ciclo_usa") {
        if (val) cfg.usa_cycle = String(val).trim();
      }


      if (key === "cadda_year" || key === "anio_cadda" || key === "año_cadda") {
        if (val) cfg.cadda_year = String(val).trim();
      }
    }
  }
  return cfg;
}


function getSheetOrThrow_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`No se encontró la hoja "${name}".`);
  return sh;
}


/** ---------------- HEADER MAPS ---------------- */


function normalizeHeader_(h) {
  const raw = String(h == null ? "" : h).replace(/["']/g, "");
  return raw
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}


function buildAliasLookup_(aliasesObj) {
  const lookup = {};
  Object.keys(aliasesObj).forEach(canonical => {
    (aliasesObj[canonical] || []).forEach(alias => {
      lookup[normalizeHeader_(alias)] = canonical;
    });
  });
  return lookup;
}


function buildHeaderMap_(headers, aliasLookup) {
  const map = {};
  const normalized = headers.map(normalizeHeader_);
  for (let i = 0; i < normalized.length; i++) {
    const can = aliasLookup[normalized[i]];
    if (can && map[can] == null) map[can] = i;
  }
  return map;
}


function getHeaderMap(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const map = {};
  headers.forEach((h, idx) => {
    const norm = normalizeHeader_(h);
    if (norm && map[norm] == null) map[norm] = idx;
  });
  return { headers, map };
}


function ensureColumns_(sheet, requiredCanonicalCols) {
  const lastCol = sheet.getLastColumn();
  const headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const norm = headers.map(normalizeHeader_);
  const missing = [];


  requiredCanonicalCols.forEach(c => {
    if (norm.indexOf(normalizeHeader_(c)) === -1) missing.push(c);
  });


  if (headers.length === 0 && missing.length) {
    sheet.getRange(1, 1, 1, missing.length).setValues([missing]);
    return missing;
  }


  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
  return missing;
}


function ensureColumns(sheet, requiredCanonicalCols) {
  const added = ensureColumns_(sheet, requiredCanonicalCols);
  if (added.length) {
    const { map } = getHeaderMap(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      added.forEach(col => {
        const idx = map[normalizeHeader_(col)];
        if (idx != null) {
          sheet.getRange(2, idx + 1, lastRow - 1, 1).setValue(false);
        }
      });
    }
  }
  return added;
}


function findRowByValue(sheet, colName, value, precomputed) {
  const { headers, map } = precomputed || getHeaderMap(sheet);
  const norm = normalizeHeader_(colName);
  const idx = map[norm];
  if (idx == null) return -1;


  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;


  const colValues = sheet.getRange(2, idx + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < colValues.length; i++) {
    if (String(colValues[i][0]).trim() === String(value).trim()) {
      return i + 2;
    }
  }
  return -1;
}


/** ---------------- TIME / PARSE ---------------- */


function toNumber_(v) {
  if (v == null || v === "") return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}


function parseBoolean_(v) {
  if (v === true || v === false) return v;
  if (typeof v === "number") return v !== 0;
  if (v == null) return false;
  const s = String(v).trim().toLowerCase();
  if (s === "true" || s === "1" || s === "si" || s === "sí" || s === "yes") return true;
  if (s === "false" || s === "0" || s === "no") return false;
  return false;
}


function parseTimeToSeconds_(input) {
  if (input == null || input === "") return null;


  if (Object.prototype.toString.call(input) === "[object Date]") {
    const d = input;
    const sec = d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds() + (d.getMilliseconds() / 1000);
    return isFinite(sec) ? sec : null;
  }


  if (typeof input === "number") {
    if (!isFinite(input)) return null;
    if (input > 0 && input < 1) return input * 86400;
    return input;
  }


  const s = String(input).trim();
  if (!s) return null;


  if (/^\w{3}\s\w{3}\s\d{2}\s\d{4}\s/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      return d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds() + (d.getMilliseconds() / 1000);
    }
  }


  const normalized = s.replace(",", ".");
  if (/^\d+(\.\d+)?$/.test(normalized)) {
    const n = parseFloat(normalized);
    return isNaN(n) ? null : n;
  }


  const parts = normalized.split(":").map(p => p.trim());
  if (parts.length === 2) {
    const mm = parseInt(parts[0], 10);
    const ss = parseFloat(parts[1]);
    if (isNaN(mm) || isNaN(ss)) return null;
    return mm * 60 + ss;
  }
  if (parts.length === 3) {
    const hh = parseInt(parts[0], 10);
    const mm = parseInt(parts[1], 10);
    const ss = parseFloat(parts[2]);
    if (isNaN(hh) || isNaN(mm) || isNaN(ss)) return null;
    return hh * 3600 + mm * 60 + ss;
  }


  return null;
}


function formatSeconds_(sec) {
  if (sec == null || !isFinite(sec)) return "";
  const total = Number(sec);
  if (total >= 60) {
    const mm = Math.floor(total / 60);
    const ss = total - mm * 60;
    const ssStr = ss.toFixed(2).padStart(5, "0");
    return `${mm}:${ssStr}`;
  }
  return total.toFixed(2);
}


/** ---------------- AGE / CATEGORY ---------------- */


function parseISODate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]") return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}


function calcAgeAt_(birthDate, atDate) {
  const b = parseISODate_(birthDate);
  const a = parseISODate_(atDate);
  if (!b || !a) return null;


  let age = a.getFullYear() - b.getFullYear();
  const m = a.getMonth() - b.getMonth();
  if (m < 0 || (m === 0 && a.getDate() < b.getDate())) age--;
  return age;
}


function categoryFromAge_(age) {
  if (age == null) return "";
  if (age <= 9) return "Preinfantil";
  if (age === 10) return "Infantil 1";
  if (age === 11) return "Infantil 2";
  if (age === 12) return "Cadete 1";
  if (age === 13) return "Cadete 2";
  if (age === 14) return "Juvenil 1";
  if (age === 15) return "Juvenil 2";
  return "Mayor";
}


/** ---------------- NORMALIZE SWIMMER / MARK ---------------- */


function normalizeSwimmerRow_(row, headers, map) {
  const get = (key) => {
    const idx = map[key];
    return (idx == null) ? "" : row[idx];
  };


  const swimmer = {
    coach_id: String(get("coach_id") || "").trim(),
    swimmer_id: String(get("swimmer_id") || "").trim(),
    nombre: String(get("nombre") || "").trim(),
    fecha_nac: get("fecha_nac") || "",
    genero: String(get("genero") || "").trim(),
    altura_cm: toNumber_(get("altura_cm")),
    peso_kg: toNumber_(get("peso_kg")),
    fc_reposo: toNumber_(get("fc_reposo")),
  };


  const age = calcAgeAt_(swimmer.fecha_nac, new Date());
  if (age != null) swimmer.edad = age;
  swimmer.categoria = categoryFromAge_(age);


  return swimmer;
}


function normalizeMarkRow_(row, headers, map) {
  const get = (key) => {
    const idx = map[key];
    return (idx == null) ? "" : row[idx];
  };


  let coach_id = String(get("coach_id") || "").trim();
  let swimmer_id = String(get("swimmer_id") || "").trim();


  let fecha = get("fecha");
  let lugar_evento = String(get("lugar_evento") || "").trim();


  let tipo_toma = String(get("tipo_toma") || "").trim();
  let curso = String(get("curso") || "").trim();


  let estilo = String(get("estilo") || "").trim();
  let distancia_m = toNumber_(get("distancia_m"));
  let carril = String(get("carril") || "").trim();


  let tiempo_raw = get("tiempo_raw");
  let tiempo_str = String(get("tiempo_str") || "").trim();
  let tiempo_s = toNumber_(get("tiempo_s"));


  const created_at = get("created_at") || "";
  const client_mark_id = String(get("client_mark_id") || "").trim();
  const edad_ref = toNumber_(get("edad_ref"));
  const categoria_ref = String(get("categoria_ref") || "").trim();


  if (tiempo_s == null) {
    tiempo_s = parseTimeToSeconds_(tiempo_str) ?? parseTimeToSeconds_(tiempo_raw);
  }
  if (!tiempo_str) {
    tiempo_str = formatSeconds_(tiempo_s);
  }
  if (tiempo_raw == null || tiempo_raw === "") {
    tiempo_raw = tiempo_str;
  }


  const fixed = fixCourseSwap_({ tipo_toma, curso });
  tipo_toma = fixed.tipo_toma;
  curso = fixed.curso;


  return {
    coach_id,
    swimmer_id,
    fecha,
    lugar_evento,
    tipo_toma,
    curso,
    estilo,
    distancia_m: (distancia_m == null ? null : distancia_m),
    carril,
    tiempo_raw: (tiempo_raw == null ? "" : String(tiempo_raw)),
    tiempo_s: (tiempo_s == null ? null : Number(tiempo_s)),
    tiempo_str,
    created_at,
    client_mark_id,
    edad_ref: (edad_ref == null ? null : Number(edad_ref)),
    categoria_ref,
  };
}


function fixCourseSwap_(obj) {
  let tipo_toma = String(obj.tipo_toma || "").trim();
  let curso = String(obj.curso || "").trim();


  const tipoUp = tipo_toma.toUpperCase();
  const cursoUp = curso.toUpperCase();


  const tipoIsPool = POOL_COURSE_SET.has(tipoUp);
  const cursoIsPool = POOL_COURSE_SET.has(cursoUp);


  const tipoIsTake = TAKE_TYPE_SET.has(tipoUp);
  const cursoIsTake = TAKE_TYPE_SET.has(cursoUp);


  if (tipoIsPool && cursoIsTake) {
    const tmp = tipo_toma;
    tipo_toma = curso;
    curso = tmp;
    return { tipo_toma: tipo_toma.toUpperCase(), curso: curso.toUpperCase() };
  }


  if (tipoIsTake) tipo_toma = tipoUp;
  if (cursoIsPool) curso = cursoUp;


  if (cursoIsTake && !tipoIsPool) {
    tipo_toma = cursoUp;
  }
  if (tipoIsPool && !cursoIsPool) {
    curso = tipoUp;
  }


  return { tipo_toma, curso };
}


/** ---------------- STANDARDS (GLOBAL) ---------------- */


function standardsCacheGet_(key) {
  const cache = CacheService.getScriptCache();
  const v = cache.get(key);
  if (!v) return null;
  try { return JSON.parse(v); } catch (_) { return null; }
}


function standardsCachePut_(key, obj, ttlSeconds) {
  const cache = CacheService.getScriptCache();
  cache.put(key, JSON.stringify(obj), ttlSeconds || 6 * 3600);
}


function buildUsaKey_(cycle, sex, age, style, dist, course, level) {
  return `${cycle}|${sex}|${age}|${style}|${dist}|${course}|${level}`;
}


function buildCaddaKey_(year, sex, age, style, dist, course, label) {
  return `${year}|${sex}|${age}|${style}|${dist}|${course}|${label}`;
}


function fetchStandardsBatch_(cfg, requests) {
  if (!requests.length) return {};


  const results = {};
  const toFetch = [];


  requests.forEach(r => {
    const cached = standardsCacheGet_(r.cacheKey);
    if (cached != null) {
      results[r.cacheKey] = cached;
    } else {
      toFetch.push(r);
    }
  });


  if (!toFetch.length) return results;
  if (!cfg.standards_webapp_url) return results;


  const fetchReqs = toFetch.map(r => ({
    url: r.url,
    muteHttpExceptions: true,
    followRedirects: true,
    method: "get",
    headers: { "Cache-Control": "no-cache" },
  }));


  const responses = UrlFetchApp.fetchAll(fetchReqs);


  for (let i = 0; i < responses.length; i++) {
    const r = toFetch[i];
    const resp = responses[i];
    const code = resp.getResponseCode();
    let parsed = null;


    if (code >= 200 && code < 300) {
      const txt = resp.getContentText() || "{}";
      try { parsed = JSON.parse(txt); } catch (_) { parsed = null; }
    }


    results[r.cacheKey] = parsed;
    if (parsed != null) standardsCachePut_(r.cacheKey, parsed, 6 * 3600);
  }


  return results;
}


function enrichMarksWithStandards_(cfg, swimmer, marks, opts) {
  if (!cfg.standards_webapp_url) return marks;


  const options = opts || {};
  const sex = (swimmer && swimmer.genero) ? String(swimmer.genero).trim().toUpperCase() : "";
  const birth = swimmer ? swimmer.fecha_nac : null;


  const reqs = [];
  const seen = new Set();


  const uniqueRequestsLimit = 90; // evita que solicitudes masivas causen timeouts


  marks.forEach(m => {
    const style = m.estilo;
    const dist = m.distancia_m;
    if (!style || !dist) return;


    const markDate = parseISODate_(m.fecha) || parseISODate_(m.created_at) || new Date();
    const ageRef = (m.edad_ref != null) ? Number(m.edad_ref) : calcAgeAt_(birth, markDate);


    const course = "LCM";
    const cycle = cfg.usa_cycle;
    const year = cfg.cadda_year;


    ["A", "AA"].forEach(level => {
      const key = "USA|" + buildUsaKey_(cycle, sex, ageRef, style, dist, course, level);
      if (seen.size < uniqueRequestsLimit && !seen.has(key)) {
        seen.add(key);
        const url = `${cfg.standards_webapp_url}?action=get_usa&cycle=${encodeURIComponent(cycle)}&sexo=${encodeURIComponent(sex)}&edad=${encodeURIComponent(ageRef)}&estilo=${encodeURIComponent(style)}&distancia_m=${encodeURIComponent(dist)}&curso=${encodeURIComponent(course)}&nivel=${encodeURIComponent(level)}`;
        reqs.push({ cacheKey: key, url });
      }
    });


    const cKey = "CADDA|" + buildCaddaKey_(year, sex, ageRef, style, dist, course, "Minima Nacional");
    if (seen.size < uniqueRequestsLimit && !seen.has(cKey)) {
      seen.add(cKey);
      const url = `${cfg.standards_webapp_url}?action=get_cadda&year=${encodeURIComponent(year)}&sexo=${encodeURIComponent(sex)}&edad=${encodeURIComponent(ageRef)}&estilo=${encodeURIComponent(style)}&distancia_m=${encodeURIComponent(dist)}&curso=${encodeURIComponent(course)}&label=${encodeURIComponent("Minima Nacional")}`;
      reqs.push({ cacheKey: cKey, url });
    }
  });


  const fetched = options.skipStandards ? {} : fetchStandardsBatch_(cfg, reqs);
  const conv = cfg.conv_scm_to_lcm || DEFAULT_CONV_SCM_TO_LCM;


  return marks.map(m => {
    const out = { ...m };


    const style = out.estilo;
    const dist = out.distancia_m;


    const markDate = parseISODate_(out.fecha) || parseISODate_(out.created_at) || new Date();
    const ageRef = (out.edad_ref != null) ? Number(out.edad_ref) : calcAgeAt_(birth, markDate);
    out.edad_ref = (ageRef == null ? out.edad_ref : ageRef);
    out.categoria_ref = out.categoria_ref || categoryFromAge_(ageRef);


    const courseTyped = String(out.curso || "").toUpperCase();
    out.comparacion_curso = "LCM";


    let equivLCM = null;
    if (out.tiempo_s != null) {
      if (courseTyped === "SCM") {
        equivLCM = Number(out.tiempo_s) * conv;
        out.origen = "SCM (conv→LCM)";
      } else {
        equivLCM = Number(out.tiempo_s);
        out.origen = courseTyped || "LCM";
      }
    }
    out.equiv_lcm_s = equivLCM;
    out.equiv_lcm_str = (equivLCM == null ? "" : formatSeconds_(equivLCM));


    const sex = (swimmer && swimmer.genero) ? String(swimmer.genero).trim().toUpperCase() : "";
    const course = "LCM";
    const cycle = cfg.usa_cycle;
    const year = cfg.cadda_year;


    const usaAKey = "USA|" + buildUsaKey_(cycle, sex, ageRef, style, dist, course, "A");
    const usaAAKey = "USA|" + buildUsaKey_(cycle, sex, ageRef, style, dist, course, "AA");
    const cKey = "CADDA|" + buildCaddaKey_(year, sex, ageRef, style, dist, course, "Minima Nacional");


    const usaA = fetched[usaAKey];
    const usaAA = fetched[usaAAKey];
    const cadda = fetched[cKey];


    const usa = {
      A_s: null, A_str: "", A_brecha_s: null, A_brecha_pct: null,
      AA_s: null, AA_str: "", AA_brecha_s: null, AA_brecha_pct: null,
    };


    if (usaA && usaA.status === "ok" && usaA.result && usaA.result.found) {
      usa.A_s = toNumber_(usaA.result.tiempo_s);
      usa.A_str = usaA.result.tiempo_str || (usa.A_s != null ? formatSeconds_(usa.A_s) : "");
    }
    if (usaAA && usaAA.status === "ok" && usaAA.result && usaAA.result.found) {
      usa.AA_s = toNumber_(usaAA.result.tiempo_s);
      usa.AA_str = usaAA.result.tiempo_str || (usa.AA_s != null ? formatSeconds_(usa.AA_s) : "");
    }


    let caddaObj = null;
    if (cadda && cadda.status === "ok" && cadda.result && cadda.result.found) {
      const t = toNumber_(cadda.result.tiempo_s);
      caddaObj = {
        tiempo_s: (t == null ? null : t),
        tiempo_str: cadda.result.tiempo_str || (t != null ? formatSeconds_(t) : ""),
        brecha_s: null,
        brecha_pct: null,
      };
    }


    if (equivLCM != null) {
      if (usa.A_s != null) {
        usa.A_brecha_s = equivLCM - usa.A_s;
        usa.A_brecha_pct = (usa.A_brecha_s / usa.A_s) * 100;
      }
      if (usa.AA_s != null) {
        usa.AA_brecha_s = equivLCM - usa.AA_s;
        usa.AA_brecha_pct = (usa.AA_brecha_s / usa.AA_s) * 100;
      }
      if (caddaObj && caddaObj.tiempo_s != null) {
        caddaObj.brecha_s = equivLCM - caddaObj.tiempo_s;
        caddaObj.brecha_pct = (caddaObj.brecha_s / caddaObj.tiempo_s) * 100;
      }
    }


    out.usa = usa;
    out.cadda = caddaObj;


    out.nivel = "—";
    if (equivLCM != null) {
      if (usa.AA_s != null && equivLCM <= usa.AA_s) out.nivel = "AA";
      else if (usa.A_s != null && equivLCM <= usa.A_s) out.nivel = "A";
      else if (caddaObj && caddaObj.tiempo_s != null && equivLCM <= caddaObj.tiempo_s) out.nivel = "AR";
    }


    return out;
  });
}


/** ---------------- CACHE HELPERS ---------------- */


function cacheKey_(prefix, params) {
  const parts = [prefix];
  Object.keys(params).sort().forEach(k => parts.push(`${k}:${params[k]}`));
  return parts.join("|");
}


function getCached_(key) {
  const cache = CacheService.getScriptCache();
  const val = cache.get(key);
  if (!val) return null;
  try { return JSON.parse(val); } catch (_) { return null; }
}


function putCached_(key, value, ttl) {
  const cache = CacheService.getScriptCache();
  cache.put(key, JSON.stringify(value), ttl || CACHE_TTL_SECONDS);
}


/** ---------------- HANDLERS ---------------- */


function handleGetSwimmers_(params) {
  const ss = getSpreadsheet_(params);
  const cfg = getConfig_(ss);


  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();


  const cacheKey = cacheKey_("swimmers", { ss: ss.getId(), coachId });
  const cached = getCached_(cacheKey);
  if (cached) return cached;


  const sheet = getSheetOrThrow_(ss, cfg.swimmers_sheet);
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || lastCol < 1) return { status: "ok", swimmers: [], nadadores: [] };


  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const aliasLookup = buildAliasLookup_(SWIMMER_HEADER_ALIASES);
  const map = buildHeaderMap_(headers, aliasLookup);


  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let swimmers = values.map(r => normalizeSwimmerRow_(r, headers, map));


  if (coachId) swimmers = swimmers.filter(s => String(s.coach_id || "") === coachId);


  const payload = { status: "ok", swimmers, nadadores: swimmers };
  putCached_(cacheKey, payload);
  return payload;
}


function handleGetSwimmerProfile_(params) {
  const ss = getSpreadsheet_(params);
  const cfg = getConfig_(ss);


  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();
  const swimmerId = String(params.swimmer_id || params.swimmerId || params.id_nadador || "").trim();
  if (!swimmerId) return { status: "error", error: "swimmer_id requerido" };


  const sheet = getSheetOrThrow_(ss, cfg.swimmers_sheet);
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || lastCol < 1) return { status: "ok", swimmer: null };


  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const aliasLookup = buildAliasLookup_(SWIMMER_HEADER_ALIASES);
  const map = buildHeaderMap_(headers, aliasLookup);


  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const swimmers = values.map(r => normalizeSwimmerRow_(r, headers, map));


  const swimmer = swimmers.find(s =>
    String(s.swimmer_id || "") === swimmerId && (!coachId || String(s.coach_id || "") === coachId)
  ) || null;


  return { status: "ok", swimmer };
}


function handleGetSwimmerMarksWithContext_(params) {
  const ss = getSpreadsheet_(params);
  const cfg = getConfig_(ss);


  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();
  const swimmerId = String(params.swimmer_id || params.swimmerId || params.id_nadador || "").trim();
  if (!swimmerId) return { status: "error", error: "swimmer_id requerido" };


  const cacheKey = cacheKey_("marks", { ss: ss.getId(), coachId, swimmerId });
  const cached = getCached_(cacheKey);
  if (cached) return cached;


  const prof = handleGetSwimmerProfile_({ coach_id: coachId, swimmer_id: swimmerId, spreadsheet_id: ss.getId() });
  const swimmer = (prof && prof.status === "ok") ? prof.swimmer : null;


  const marksSheet = getSheetOrThrow_(ss, cfg.marks_sheet);
  ensureColumns_(marksSheet, CANONICAL_MARK_COLS);


  const lastCol = marksSheet.getLastColumn();
  const lastRow = marksSheet.getLastRow();
  if (lastRow < 2 || lastCol < 1) {
    const payload = { status: "ok", swimmer, marks: [] };
    putCached_(cacheKey, payload);
    return payload;
  }


  const headers = marksSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const aliasLookup = buildAliasLookup_(HEADER_ALIASES);
  const map = buildHeaderMap_(headers, aliasLookup);


  const values = marksSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let marks = values.map(r => normalizeMarkRow_(r, headers, map));


  marks = marks.filter(m => {
    const okCoach = !coachId || String(m.coach_id || "") === coachId;
    const okSw = String(m.swimmer_id || "") === swimmerId;
    return okCoach && okSw;
  });


  const enriched = enrichMarksWithStandards_(cfg, swimmer, marks, { skipStandards: params.skip_standards === "1" });
  const payload = { status: "ok", swimmer, marks: enriched };
  putCached_(cacheKey, payload);
  return payload;
}


function handleAddMark_(params) {
  const ss = getSpreadsheet_(params);
  const cfg = getConfig_(ss);


  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();
  const swimmerId = String(params.swimmer_id || params.swimmerId || params.id_nadador || "").trim();
  if (!coachId || !swimmerId) return { status: "error", error: "coach_id y swimmer_id requeridos" };


  const sheet = getSheetOrThrow_(ss, cfg.marks_sheet);
  ensureColumns_(sheet, CANONICAL_MARK_COLS);


  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];


  const mark = {
    coach_id: coachId,
    swimmer_id: swimmerId,
    fecha: params.fecha || params.date || "",
    lugar_evento: params.lugar_evento || params.lugar || params.evento || "",
    tipo_toma: params.tipo_toma || params.tipo || params.modalidad || "",
    curso: params.curso || params.pool || params.piscina || "",
    estilo: params.estilo || params.trazo || "",
    distancia_m: toNumber_(params.distancia_m),
    carril: params.carril || "",
    tiempo_raw: params.tiempo_raw ?? params.tiempo ?? params.time ?? "",
    tiempo_str: params.tiempo_str ?? params.time_str ?? "",
    tiempo_s: toNumber_(params.tiempo_s),
    created_at: params.created_at || new Date().toISOString(),
    client_mark_id: params.client_mark_id || params.mark_id || "",
    edad_ref: null,
    categoria_ref: "",
  };


  const fixed = fixCourseSwap_({ tipo_toma: mark.tipo_toma, curso: mark.curso });
  mark.tipo_toma = fixed.tipo_toma;
  mark.curso = fixed.curso;


  if (mark.tiempo_s == null) {
    mark.tiempo_s = parseTimeToSeconds_(mark.tiempo_str) ?? parseTimeToSeconds_(mark.tiempo_raw);
  }
  if (!mark.tiempo_str) mark.tiempo_str = formatSeconds_(mark.tiempo_s);


  const prof = handleGetSwimmerProfile_({ coach_id: coachId, swimmer_id: swimmerId, spreadsheet_id: ss.getId() });
  const swimmer = (prof && prof.status === "ok") ? prof.swimmer : null;
  if (swimmer) {
    const ageRef = calcAgeAt_(swimmer.fecha_nac, mark.fecha || mark.created_at || new Date());
    if (ageRef != null) {
      mark.edad_ref = ageRef;
      mark.categoria_ref = categoryFromAge_(ageRef);
    }
  }


  const aliasLookup = buildAliasLookup_(HEADER_ALIASES);


  const row = headers.map(h => {
    const can = aliasLookup[normalizeHeader_(h)] || normalizeHeader_(h);
    if (mark.hasOwnProperty(can)) return mark[can];
    return "";
  });


  sheet.appendRow(row);


  const saved = normalizeMarkRow_(row, headers, buildHeaderMap_(headers, aliasLookup));
  const enriched = swimmer ? enrichMarksWithStandards_(cfg, swimmer, [saved]) : [saved];


  CacheService.getScriptCache().remove(cacheKey_("marks", { ss: ss.getId(), coachId, swimmerId }));
  CacheService.getScriptCache().remove(cacheKey_("swimmers", { ss: ss.getId(), coachId }));


  return { status: "ok", mark: enriched[0] || saved };
}


function handleGetSwimmerPermissions_(params) {
  const ss = getSpreadsheet_(params);
  const cfg = getConfig_(ss);


  const swimmerId = String(params.swimmer_id || params.swimmerId || "").trim();
  if (!swimmerId) {
    return { ok: false, error: { code: "MISSING_SWIMMER_ID", message: "swimmer_id requerido" } };
  }


  const sheet = getSheetOrThrow_(ss, cfg.swimmers_sheet);
  ensureColumns(sheet, ["allow_marks_edit", "allow_marks_delete"]);


  const headerInfo = getHeaderMap(sheet);
  const rowIndex = findRowByValue(sheet, "swimmer_id", swimmerId, headerInfo);
  if (rowIndex < 2) {
    return { ok: false, error: { code: "SWIMMER_NOT_FOUND", message: "Nadador no encontrado" } };
  }


  const lastCol = sheet.getLastColumn();
  const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  const map = headerInfo.map;


  const allowEditIdx = map[normalizeHeader_("allow_marks_edit")];
  const allowDeleteIdx = map[normalizeHeader_("allow_marks_delete")];


  const allowMarksEdit = parseBoolean_(allowEditIdx == null ? false : rowValues[allowEditIdx]);
  const allowMarksDelete = parseBoolean_(allowDeleteIdx == null ? false : rowValues[allowDeleteIdx]);


  if ((allowEditIdx != null && rowValues[allowEditIdx] === "") || (allowDeleteIdx != null && rowValues[allowDeleteIdx] === "")) {
    const updates = [];
    if (allowEditIdx != null && rowValues[allowEditIdx] === "") updates.push({ col: allowEditIdx + 1, val: false });
    if (allowDeleteIdx != null && rowValues[allowDeleteIdx] === "") updates.push({ col: allowDeleteIdx + 1, val: false });
    updates.forEach(u => sheet.getRange(rowIndex, u.col).setValue(u.val));
  }


  return {
    ok: true,
    data: {
      swimmer_id: swimmerId,
      allow_marks_edit: allowMarksEdit,
      allow_marks_delete: allowMarksDelete,
    },
  };
}


function handleSetSwimmerPermissions_(params) {
  const ss = getSpreadsheet_(params);
  const cfg = getConfig_(ss);


  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();
  const swimmerId = String(params.swimmer_id || params.swimmerId || "").trim();
  if (!coachId || !swimmerId) {
    return { ok: false, error: { code: "MISSING_PARAMS", message: "coach_id y swimmer_id requeridos" } };
  }


  const sheet = getSheetOrThrow_(ss, cfg.swimmers_sheet);
  ensureColumns(sheet, ["allow_marks_edit", "allow_marks_delete"]);
  const headerInfo = getHeaderMap(sheet);
  const map = headerInfo.map;


  const rowIndex = findRowByValue(sheet, "swimmer_id", swimmerId, headerInfo);
  if (rowIndex < 2) {
    return { ok: false, error: { code: "SWIMMER_NOT_FOUND", message: "Nadador no encontrado" } };
  }


  const lastCol = sheet.getLastColumn();
  const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];


  const coachIdx = map[normalizeHeader_("coach_id")];
  const rowCoach = coachIdx == null ? "" : String(rowValues[coachIdx] || "").trim();
  if (rowCoach && rowCoach !== coachId) {
    return { ok: false, error: { code: "FORBIDDEN", message: "El nadador no pertenece al coach" } };
  }


  const allowMarksEdit = parseBoolean_(params.allow_marks_edit);
  const allowMarksDelete = parseBoolean_(params.allow_marks_delete);


  const editIdx = map[normalizeHeader_("allow_marks_edit")];
  const deleteIdx = map[normalizeHeader_("allow_marks_delete")];
  if (editIdx != null) sheet.getRange(rowIndex, editIdx + 1).setValue(allowMarksEdit);
  if (deleteIdx != null) sheet.getRange(rowIndex, deleteIdx + 1).setValue(allowMarksDelete);


  const updatedAtIdx = map[normalizeHeader_("updated_at")];
  if (updatedAtIdx != null) {
    sheet.getRange(rowIndex, updatedAtIdx + 1).setValue(new Date());
  }


  CacheService.getScriptCache().remove(cacheKey_("swimmers", { ss: ss.getId(), coachId }));


  return {
    ok: true,
    data: {
      swimmer_id: swimmerId,
      allow_marks_edit: allowMarksEdit,
      allow_marks_delete: allowMarksDelete,
    },
  };
}


function handleDiag_(params) {
  const ss = getSpreadsheet_(params);
  const cfg = getConfig_(ss);


  const swimmersSh = ss.getSheetByName(cfg.swimmers_sheet);
  const marksSh = ss.getSheetByName(cfg.marks_sheet);


  const sheets = ss.getSheets().map(s => s.getName());
  const out = {
    status: "ok",
    sheets,
    config: {
      swimmers_sheet: cfg.swimmers_sheet,
      marks_sheet: cfg.marks_sheet,
      config_sheet: cfg.config_sheet,
      conv_scm_to_lcm: cfg.conv_scm_to_lcm,
      standards_webapp_url: cfg.standards_webapp_url,
      usa_cycle: cfg.usa_cycle,
      cadda_year: cfg.cadda_year,
    },
    rows: {},
  };


  if (swimmersSh) out.rows[cfg.swimmers_sheet] = Math.max(0, swimmersSh.getLastRow() - 1);
  if (marksSh) out.rows[cfg.marks_sheet] = Math.max(0, marksSh.getLastRow() - 1);


  return out;
}
