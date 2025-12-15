/**
 * MDV COACH – Apps Script backend (stable rewrite + permissions)
 * - Mantiene actions existentes: ping, diag, get_swimmers, get_swimmer_profile,
 *   get_swimmer_marks_with_context, add_mark
 * - Agrega: get_swimmer_permissions, set_swimmer_permissions
 * - Prepara (no rompe) para etapas siguientes: edit_mark, delete_mark
 *
 * NOTA: Este backend está diseñado para ser tolerante a cambios de headers.
 */

const SHEETS_DEFAULT = {
  swimmers: "swimmers",
  marks: "Marks",
  config: "config",
};

const DEFAULT_CONV_SCM_TO_LCM = 1.0103;

// Columnas canónicas (si faltan, se agregan al final)
const CANONICAL_SWIMMER_COLS = [
  "coach_id",
  "swimmer_id",
  "nombre",
  "fecha_nac",
  "genero",
  "altura_cm",
  "peso_kg",
  "fc_reposo",
  "created_at",
  "updated_at",
  "allow_marks_edit",
  "allow_marks_delete",
];

const CANONICAL_MARK_COLS = [
  "coach_id",
  "swimmer_id",
  "fecha",
  "season_year",
  "age_chip",
  "tipo_toma",
  "curso",
  "estilo",
  "distancia_m",
  "carril",
  "tiempo_str",
  "tiempo_s",
  "created_at",
  "lugar_evento",
  "tiempo_raw",
  "updated_at",
  "mark_id",
  "edited_by",
  "deleted_at",
  "source",
  // Back-compat
  "client_mark_id",
  "edad_ref",
  "categoria_ref",
];

const HEADER_ALIASES = {
  coach_id: ["coach_id", "coach", "coachid", "id_entrenador", "idcoach"],
  swimmer_id: ["swimmer_id", "swimmerid", "id_nadador", "nadador_id", "idnadador"],

  fecha: ["fecha", "fecha_evento", "fecha_de_toma", "date"],
  season_year: ["season_year", "year", "anio", "año", "ano", "año_calendario", "anio_calendario"],
  age_chip: ["age_chip", "edad_chip", "age", "edad_marca", "edad_en_marca"],

  tipo_toma: ["tipo_toma", "tipo", "modalidad", "tipo_de_toma"],
  curso: ["curso", "pool", "piscina", "tipo_pool"],
  estilo: ["estilo", "trazo", "stroke"],
  distancia_m: ["distancia_m", "distancia", "distancia_mts", "distancia_metros"],
  carril: ["carril", "lane"],

  tiempo_raw: ["tiempo_raw", "tiempo", "time", "marca", "raw_time"],
  tiempo_str: ["tiempo_str", "time_str", "time_text", "time_string"],
  tiempo_s: ["tiempo_s", "time_s", "segundos", "seconds"],

  lugar_evento: ["lugar_evento", "lugar", "evento", "ubicacion", "ubicación"],
  created_at: ["created_at", "creado_en", "creado", "timestamp"],
  updated_at: ["updated_at", "actualizado_en", "updated", "last_update", "last_updated"],
  mark_id: ["mark_id", "id_marca", "id_mark"],
  edited_by: ["edited_by", "editado_por", "editor"],
  deleted_at: ["deleted_at", "borrado_en", "eliminado_en", "fecha_borrado"],
  source: ["source", "fuente", "origen"],

  // Back-compat
  client_mark_id: ["client_mark_id", "id_marca_cliente"],
  edad_ref: ["edad_ref", "edad_ref_marca"],
  categoria_ref: ["categoria_ref", "cat_ref"],
};

const SWIMMER_HEADER_ALIASES = {
  coach_id: HEADER_ALIASES.coach_id,
  swimmer_id: HEADER_ALIASES.swimmer_id,
  nombre: ["nombre", "name", "nadador", "swimmer_name"],
  fecha_nac: ["fecha_nac", "fecha_de_nacimiento", "nacimiento", "birthdate", "nac_date"],
  genero: ["genero", "género", "sexo", "gender"],
  altura_cm: ["altura_cm", "altura", "height_cm"],
  peso_kg: ["peso_kg", "peso", "weight_kg"],
  fc_reposo: ["fc_reposo", "fc", "fc_rest", "rest_hr", "fc_resting"],
  created_at: HEADER_ALIASES.created_at,
  updated_at: HEADER_ALIASES.updated_at,
  allow_marks_edit: ["allow_marks_edit", "can_edit_marks", "perm_edit_marks", "editar_marcas"],
  allow_marks_delete: ["allow_marks_delete", "can_delete_marks", "perm_delete_marks", "borrar_marcas"],
};

function doGet(e) { return handleRequest_(e); }
function doPost(e) { return handleRequest_(e); }

function handleRequest_(e) {
  try {
    const params = getParams_(e);
    const actionRaw = String(params.action || params.accion || "").trim();
    const action = actionRaw.toLowerCase();

    if (!action) return json_({ status: "error", error: "Missing action" });

    switch (action) {
      case "ping":
        return json_({ status: "ok", ts: new Date().toISOString() });

      case "diag":
        return json_(handleDiag_(params));

      case "get_swimmers":
      case "getswimmers":
        return json_(handleGetSwimmers_(params));

      case "get_swimmer_profile":
      case "getswimmerprofile":
        return json_(handleGetSwimmerProfile_(params));

      case "get_swimmer_marks_with_context":
      case "getswimmermarkswithcontext":
        return json_(handleGetSwimmerMarksWithContext_(params));

      case "add_mark":
      case "addmark":
        return json_(handleAddMark_(params));

      // Permisos
      case "get_swimmer_permissions":
      case "get_permissions":
      case "get_swimmer_perms":
        return json_(handleGetSwimmerPermissions_(params));

      case "set_swimmer_permissions":
      case "set_permissions":
      case "set_swimmer_perms":
        return json_(handleSetSwimmerPermissions_(params));

      // Etapa 2 (listo, por ahora no usado por dashboards)
      case "edit_mark":
      case "updatemark":
        return json_(handleEditMark_(params));

      case "delete_mark":
      case "deletemark":
        return json_(handleDeleteMark_(params));

      default:
        return json_({ status: "error", error: `Unknown action: ${actionRaw}` });
    }
  } catch (err) {
    return json_({ status: "error", error: String(err && err.message ? err.message : err) });
  }
}

function getParams_(e) {
  const out = {};
  if (!e) return out;

  // querystring
  if (e.parameter) Object.keys(e.parameter).forEach(k => out[k] = e.parameter[k]);

  // body (JSON o form-encoded)
  const postData = e.postData;
  if (postData && postData.contents) {
    const c = postData.contents;
    const ct = String(postData.type || "").toLowerCase();

    if (ct.includes("application/json")) {
      try { Object.assign(out, JSON.parse(c)); } catch (_) {}
    } else {
      // intentamos parsear como querystring key=value&...
      c.split("&").forEach(pair => {
        const [k, v] = pair.split("=");
        if (!k) return;
        out[decodeURIComponent(k)] = v ? decodeURIComponent(v) : "";
      });
    }
  }
  return out;
}

function json_(obj) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  // CORS básico (útil para dashes en Netlify/Wix)
  try {
    output.setHeader("Access-Control-Allow-Origin", "*");
    output.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
    output.setHeader("Access-Control-Allow-Headers", "Content-Type");
  } catch (_) {}
  return output;
}

function getActiveSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getConfig_(ss) {
  const cfg = {
    swimmers_sheet: SHEETS_DEFAULT.swimmers,
    marks_sheet: SHEETS_DEFAULT.marks,
    config_sheet: SHEETS_DEFAULT.config,

    conv_scm_to_lcm: DEFAULT_CONV_SCM_TO_LCM,

    // standards_global webapp (opcional)
    standards_webapp_url: "",
    usa_cycle: "2024-2028",
    cadda_year: "2025",
  };

  // Script Properties override
  const props = PropertiesService.getScriptProperties().getProperties() || {};
  if (props.CONV_SCM_TO_LCM) cfg.conv_scm_to_lcm = toNumber_(props.CONV_SCM_TO_LCM) || cfg.conv_scm_to_lcm;
  if (props.STANDARDS_WEBAPP_URL) cfg.standards_webapp_url = String(props.STANDARDS_WEBAPP_URL).trim();
  if (props.USA_CYCLE) cfg.usa_cycle = String(props.USA_CYCLE).trim();
  if (props.CADDA_YEAR) cfg.cadda_year = String(props.CADDA_YEAR).trim();

  // Sheet config override (key/value)
  const sheet = ss.getSheetByName(cfg.config_sheet);
  if (sheet) {
    const values = sheet.getDataRange().getValues();
    for (let i = 0; i < values.length; i++) {
      const key = normalizeHeader_(values[i][0]);
      const val = values[i][1];
      if (!key) continue;

      if (key === "marks_sheet" || key === "hoja_marcas") cfg.marks_sheet = String(val || "").trim() || cfg.marks_sheet;
      if (key === "swimmers_sheet" || key === "hoja_nadadores") cfg.swimmers_sheet = String(val || "").trim() || cfg.swimmers_sheet;

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
  for (let c = 0; c < headers.length; c++) {
    const canon = aliasLookup[normalizeHeader_(headers[c])];
    if (canon && map[canon] == null) map[canon] = c;
  }
  return map;
}

function ensureColumns_(sheet, canonicalCols) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  const headers = headerRange.getValues()[0].map(h => String(h || "").trim());
  const normSet = new Set(headers.map(normalizeHeader_));

  let changed = false;
  canonicalCols.forEach(col => {
    if (!normSet.has(normalizeHeader_(col))) {
      headers.push(col);
      normSet.add(normalizeHeader_(col));
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function toNumber_(v) {
  if (v == null || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function parseBool_(v) {
  if (v == null) return false;
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  const s = String(v).trim().toLowerCase();
  if (!s) return false;
  return (s === "true" || s === "1" || s === "si" || s === "sí" || s === "yes" || s === "y" || s === "on");
}

function parseISODate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function pad2_(n) { return String(n).padStart(2, "0"); }

function formatSeconds_(sec) {
  if (sec == null) return "";
  const total = Number(sec);
  if (!Number.isFinite(total)) return "";
  const minutes = Math.floor(total / 60);
  const seconds = total - minutes * 60;
  const sInt = Math.floor(seconds);
  const hundredths = Math.round((seconds - sInt) * 100);
  return `${minutes}:${pad2_(sInt)}.${String(hundredths).padStart(2, "0")}`;
}

function parseTimeToSeconds_(t) {
  if (t == null) return null;
  if (typeof t === "number") return Number.isFinite(t) ? t : null;
  const s = String(t).trim();
  if (!s) return null;

  // 1) mm:ss.xx
  const m1 = s.match(/^(\d+):(\d{1,2})(?:\.(\d{1,2}))?$/);
  if (m1) {
    const mm = Number(m1[1]);
    const ss = Number(m1[2]);
    const hh = m1[3] ? Number(String(m1[3]).padEnd(2, "0")) : 0;
    if ([mm, ss, hh].some(x => !Number.isFinite(x))) return null;
    return mm * 60 + ss + hh / 100;
  }

  // 2) ss.xx
  const m2 = s.match(/^(\d+)(?:\.(\d{1,2}))?$/);
  if (m2) {
    const ss = Number(m2[1]);
    const hh = m2[2] ? Number(String(m2[2]).padEnd(2, "0")) : 0;
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

function normalizeSwimmerRow_(row, map) {
  const get = (key) => (map[key] == null ? "" : row[map[key]]);
  const swimmer = {
    coach_id: String(get("coach_id") || "").trim(),
    swimmer_id: String(get("swimmer_id") || "").trim(),
    nombre: String(get("nombre") || "").trim(),
    fecha_nac: get("fecha_nac") || "",
    genero: String(get("genero") || "").trim(),
    altura_cm: toNumber_(get("altura_cm")),
    peso_kg: toNumber_(get("peso_kg")),
    fc_reposo: toNumber_(get("fc_reposo")),
    created_at: get("created_at") || "",
    updated_at: get("updated_at") || "",
    allow_marks_edit: parseBool_(get("allow_marks_edit")),
    allow_marks_delete: parseBool_(get("allow_marks_delete")),
  };
  const age = calcAgeAt_(swimmer.fecha_nac, new Date());
  if (age != null) swimmer.edad = age;
  swimmer.categoria = categoryFromAge_(age);
  return swimmer;
}

function fixCourseSwap_(obj) {
  const knownCourses = ["SCM", "LCM", "SCY"];
  const t = String(obj.tipo_toma || "").toUpperCase();
  const c = String(obj.curso || "").toUpperCase();
  if (knownCourses.includes(t) && !knownCourses.includes(c)) {
    return { tipo_toma: obj.curso, curso: obj.tipo_toma };
  }
  return obj;
}

function normalizeMarkRow_(row, map) {
  const get = (key) => (map[key] == null ? "" : row[map[key]]);

  let tipo_toma = String(get("tipo_toma") || "").trim();
  let curso = String(get("curso") || "").trim();
  const fixed = fixCourseSwap_({ tipo_toma, curso });
  tipo_toma = String(fixed.tipo_toma || "").trim();
  curso = String(fixed.curso || "").trim();

  let tiempo_s = toNumber_(get("tiempo_s"));
  let tiempo_str = String(get("tiempo_str") || "").trim();
  const tiempo_raw = (get("tiempo_raw") == null ? "" : String(get("tiempo_raw")));

  if (tiempo_s == null) {
    tiempo_s = parseTimeToSeconds_(tiempo_str) ?? parseTimeToSeconds_(tiempo_raw);
  }
  if (!tiempo_str) tiempo_str = formatSeconds_(tiempo_s);

  let fecha = get("fecha") || "";
  let season_year = toNumber_(get("season_year"));
  if (season_year == null) {
    const d = parseISODate_(fecha) || parseISODate_(get("created_at")) || parseISODate_(get("updated_at"));
    if (d) season_year = d.getFullYear();
  }
  let age_chip = toNumber_(get("age_chip"));
  const edad_ref = toNumber_(get("edad_ref"));
  if (age_chip == null && edad_ref != null) age_chip = edad_ref;

  const mark_id = String(get("mark_id") || "").trim() || String(get("client_mark_id") || "").trim();

  return {
    coach_id: String(get("coach_id") || "").trim(),
    swimmer_id: String(get("swimmer_id") || "").trim(),
    fecha,
    season_year: season_year == null ? null : Number(season_year),
    age_chip: age_chip == null ? null : Number(age_chip),
    tipo_toma,
    curso,
    estilo: String(get("estilo") || "").trim(),
    distancia_m: toNumber_(get("distancia_m")),
    carril: String(get("carril") || "").trim(),
    tiempo_str,
    tiempo_s: (tiempo_s == null ? null : Number(tiempo_s)),
    created_at: get("created_at") || "",
    lugar_evento: String(get("lugar_evento") || "").trim(),
    tiempo_raw,
    updated_at: get("updated_at") || "",
    mark_id,
    edited_by: String(get("edited_by") || "").trim(),
    deleted_at: get("deleted_at") || "",
    source: String(get("source") || "").trim(),
    // Back-compat
    client_mark_id: String(get("client_mark_id") || "").trim(),
    edad_ref: (edad_ref == null ? null : Number(edad_ref)),
    categoria_ref: String(get("categoria_ref") || "").trim(),
  };
}

/** ---------------- HANDLERS ---------------- */

function handleDiag_(params) {
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);
  return {
    status: "ok",
    spreadsheet: ss.getName(),
    swimmers_sheet: cfg.swimmers_sheet,
    marks_sheet: cfg.marks_sheet,
    standards_webapp_url: cfg.standards_webapp_url || "",
    now: new Date().toISOString(),
  };
}

function handleGetSwimmers_(params) {
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);

  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();

  const sheet = getSheetOrThrow_(ss, cfg.swimmers_sheet);
  ensureColumns_(sheet, CANONICAL_SWIMMER_COLS);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { status: "ok", swimmers: [] };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = buildHeaderMap_(headers, buildAliasLookup_(SWIMMER_HEADER_ALIASES));
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const swimmers = [];
  for (let r = 0; r < values.length; r++) {
    const sw = normalizeSwimmerRow_(values[r], map);
    if (coachId && sw.coach_id !== coachId) continue;
    swimmers.push(sw);
  }

  return { status: "ok", swimmers };
}

function handleGetSwimmerProfile_(params) {
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);

  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();
  const swimmerId = String(params.swimmer_id || params.swimmerId || params.id_nadador || "").trim();
  if (!swimmerId) return { status: "error", error: "swimmer_id requerido" };

  const sheet = getSheetOrThrow_(ss, cfg.swimmers_sheet);
  ensureColumns_(sheet, CANONICAL_SWIMMER_COLS);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { status: "error", error: "No hay nadadores" };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = buildHeaderMap_(headers, buildAliasLookup_(SWIMMER_HEADER_ALIASES));
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const sidIdx = map["swimmer_id"];
  const cidIdx = map["coach_id"];

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const sid = sidIdx == null ? "" : String(row[sidIdx] || "").trim();
    const cid = cidIdx == null ? "" : String(row[cidIdx] || "").trim();
    if (sid === swimmerId && (!coachId || cid === coachId)) {
      return { status: "ok", swimmer: normalizeSwimmerRow_(row, map) };
    }
  }
  return { status: "error", error: "No se encontró nadador" };
}

function handleGetSwimmerPermissions_(params) {
  const prof = handleGetSwimmerProfile_(params);
  if (prof.status !== "ok") return prof;
  const sw = prof.swimmer || {};
  return {
    status: "ok",
    permissions: {
      allow_marks_edit: !!sw.allow_marks_edit,
      allow_marks_delete: !!sw.allow_marks_delete,
    }
  };
}

function handleSetSwimmerPermissions_(params) {
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);

  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();
  const swimmerId = String(params.swimmer_id || params.swimmerId || params.id_nadador || "").trim();
  if (!coachId || !swimmerId) return { status: "error", error: "coach_id y swimmer_id requeridos" };

  const allowEdit = parseBool_(params.allow_marks_edit ?? params.can_edit ?? params.edit ?? params.perm_edit);
  const allowDelete = parseBool_(params.allow_marks_delete ?? params.can_delete ?? params.delete ?? params.perm_delete);

  const sheet = getSheetOrThrow_(ss, cfg.swimmers_sheet);
  ensureColumns_(sheet, CANONICAL_SWIMMER_COLS);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { status: "error", error: "No hay nadadores" };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = buildHeaderMap_(headers, buildAliasLookup_(SWIMMER_HEADER_ALIASES));
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const sidIdx = map["swimmer_id"];
  const cidIdx = map["coach_id"];
  const eIdx = map["allow_marks_edit"];
  const dIdx = map["allow_marks_delete"];
  const uIdx = map["updated_at"];

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const sid = sidIdx == null ? "" : String(row[sidIdx] || "").trim();
    const cid = cidIdx == null ? "" : String(row[cidIdx] || "").trim();
    if (sid === swimmerId && cid === coachId) {
      const rowNumber = r + 2;
      if (eIdx != null) sheet.getRange(rowNumber, eIdx + 1).setValue(allowEdit ? 1 : 0);
      if (dIdx != null) sheet.getRange(rowNumber, dIdx + 1).setValue(allowDelete ? 1 : 0);
      if (uIdx != null) sheet.getRange(rowNumber, uIdx + 1).setValue(new Date().toISOString());
      SpreadsheetApp.flush();
      return { status: "ok", permissions: { allow_marks_edit: allowEdit, allow_marks_delete: allowDelete } };
    }
  }

  return { status: "error", error: "No se encontró swimmer_id para ese coach_id" };
}

function handleGetSwimmerMarksWithContext_(params) {
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);

  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();
  const swimmerId = String(params.swimmer_id || params.swimmerId || params.id_nadador || "").trim();
  if (!swimmerId) return { status: "error", error: "swimmer_id requerido" };

  const includeDeleted = parseBool_(params.include_deleted);

  // swimmer profile para fecha_nac y genero
  const prof = handleGetSwimmerProfile_({ coach_id: coachId, swimmer_id: swimmerId });
  const swimmer = (prof.status === "ok") ? prof.swimmer : null;

  const sheet = getSheetOrThrow_(ss, cfg.marks_sheet);
  ensureColumns_(sheet, CANONICAL_MARK_COLS);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { status: "ok", marks: [], swimmer };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = buildHeaderMap_(headers, buildAliasLookup_(HEADER_ALIASES));
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const sidIdx = map["swimmer_id"];
  const cidIdx = map["coach_id"];
  const delIdx = map["deleted_at"];

  let marks = [];
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const sid = sidIdx == null ? "" : String(row[sidIdx] || "").trim();
    const cid = cidIdx == null ? "" : String(row[cidIdx] || "").trim();
    if (sid !== swimmerId) continue;
    if (coachId && cid !== coachId) continue;

    const deletedVal = delIdx == null ? "" : row[delIdx];
    if (!includeDeleted && deletedVal) continue;

    marks.push(normalizeMarkRow_(row, map));
  }

  // Enriquecer con standards (si cfg.standards_webapp_url existe)
  marks = enrichMarksWithStandards_(marks, swimmer, cfg);

  return { status: "ok", marks, swimmer };
}

function handleAddMark_(params) {
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);

  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();
  const swimmerId = String(params.swimmer_id || params.swimmerId || params.id_nadador || "").trim();
  if (!coachId || !swimmerId) return { status: "error", error: "coach_id y swimmer_id requeridos" };

  const nowIso = new Date().toISOString();

  const mark = {
    coach_id: coachId,
    swimmer_id: swimmerId,

    fecha: params.fecha || params.date || "",
    season_year: toNumber_(params.season_year),
    age_chip: toNumber_(params.age_chip),

    tipo_toma: params.tipo_toma || params.tipo || "",
    curso: params.curso || params.pool || "",
    estilo: params.estilo || "",
    distancia_m: toNumber_(params.distancia_m),
    carril: params.carril || "",

    tiempo_raw: (params.tiempo_raw ?? params.tiempo ?? params.time ?? ""),
    tiempo_str: (params.tiempo_str ?? params.time_str ?? ""),
    tiempo_s: toNumber_(params.tiempo_s),

    created_at: params.created_at || nowIso,
    lugar_evento: params.lugar_evento || params.lugar || "",
    updated_at: params.updated_at || "",
    mark_id: params.mark_id || "",
    edited_by: params.edited_by || "",
    deleted_at: params.deleted_at || "",
    source: params.source || "webapp",

    // Back-compat
    client_mark_id: params.client_mark_id || "",
    edad_ref: null,
    categoria_ref: "",
  };

  // mark_id fallback
  if (!mark.mark_id) {
    try { mark.mark_id = Utilities.getUuid(); } catch (_) { mark.mark_id = String(new Date().getTime()); }
  }

  // normalizar swap curso/tipo_toma
  const fixed = fixCourseSwap_({ tipo_toma: mark.tipo_toma, curso: mark.curso });
  mark.tipo_toma = String(fixed.tipo_toma || "").trim();
  mark.curso = String(fixed.curso || "").trim();

  // tiempos fallback
  if (mark.tiempo_s == null) mark.tiempo_s = parseTimeToSeconds_(mark.tiempo_str) ?? parseTimeToSeconds_(mark.tiempo_raw);
  if (!mark.tiempo_str) mark.tiempo_str = formatSeconds_(mark.tiempo_s);
  if (!mark.tiempo_raw) mark.tiempo_raw = mark.tiempo_str;

  // calcular chips históricos si faltan
  const prof = handleGetSwimmerProfile_({ coach_id: coachId, swimmer_id: swimmerId });
  const swimmer = (prof.status === "ok") ? prof.swimmer : null;

  const eventDate = parseISODate_(mark.fecha) || parseISODate_(mark.created_at) || new Date();
  if (mark.season_year == null && eventDate) mark.season_year = eventDate.getFullYear();

  if (swimmer) {
    const ageRef = calcAgeAt_(swimmer.fecha_nac, eventDate);
    if (ageRef != null) {
      mark.edad_ref = ageRef;
      mark.categoria_ref = categoryFromAge_(ageRef);
      if (mark.age_chip == null) mark.age_chip = ageRef;
    }
  }

  // Guardar
  const sheet = getSheetOrThrow_(ss, cfg.marks_sheet);
  ensureColumns_(sheet, CANONICAL_MARK_COLS);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = buildHeaderMap_(headers, buildAliasLookup_(HEADER_ALIASES));

  const rowOut = new Array(headers.length).fill("");
  Object.keys(mark).forEach(k => {
    const idx = map[k];
    if (idx != null) rowOut[idx] = mark[k];
  });

  sheet.appendRow(rowOut);
  SpreadsheetApp.flush();

  return { status: "ok", mark };
}

/**
 * Etapa 2 – editar marca (por mark_id).
 * No se usa hasta que el dashboard lo llame.
 */
function handleEditMark_(params) {
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);

  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();
  const swimmerId = String(params.swimmer_id || params.swimmerId || "").trim();
  const markId = String(params.mark_id || params.id || "").trim();
  if (!coachId || !swimmerId || !markId) return { status: "error", error: "coach_id, swimmer_id y mark_id requeridos" };

  const sheet = getSheetOrThrow_(ss, cfg.marks_sheet);
  ensureColumns_(sheet, CANONICAL_MARK_COLS);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { status: "error", error: "No hay marcas" };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = buildHeaderMap_(headers, buildAliasLookup_(HEADER_ALIASES));
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const midIdx = map["mark_id"];
  const sidIdx = map["swimmer_id"];
  const cidIdx = map["coach_id"];

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const cid = cidIdx == null ? "" : String(row[cidIdx] || "").trim();
    const sid = sidIdx == null ? "" : String(row[sidIdx] || "").trim();
    const mid = midIdx == null ? "" : String(row[midIdx] || "").trim();

    if (cid === coachId && sid === swimmerId && mid === markId) {
      const rowNum = r + 2;

      const editableKeys = [
        "fecha","season_year","age_chip","tipo_toma","curso","estilo","distancia_m","carril",
        "tiempo_raw","tiempo_str","tiempo_s","lugar_evento","source"
      ];

      editableKeys.forEach(k => {
        if (params[k] == null) return;
        const idx = map[k];
        if (idx == null) return;
        sheet.getRange(rowNum, idx + 1).setValue(params[k]);
      });

      // recalcular tiempo_s/tiempo_str si vinieron como raw/str
      const tStr = params.tiempo_str ?? (map.tiempo_str != null ? sheet.getRange(rowNum, map.tiempo_str + 1).getValue() : "");
      const tRaw = params.tiempo_raw ?? (map.tiempo_raw != null ? sheet.getRange(rowNum, map.tiempo_raw + 1).getValue() : "");
      const tS = toNumber_(params.tiempo_s) ?? parseTimeToSeconds_(tStr) ?? parseTimeToSeconds_(tRaw);
      if (map.tiempo_s != null && tS != null) sheet.getRange(rowNum, map.tiempo_s + 1).setValue(tS);
      if (map.tiempo_str != null && tStr) sheet.getRange(rowNum, map.tiempo_str + 1).setValue(tStr);
      if (map.updated_at != null) sheet.getRange(rowNum, map.updated_at + 1).setValue(new Date().toISOString());
      if (map.edited_by != null && params.edited_by) sheet.getRange(rowNum, map.edited_by + 1).setValue(params.edited_by);

      SpreadsheetApp.flush();
      return { status: "ok" };
    }
  }

  return { status: "error", error: "mark_id no encontrado" };
}

/**
 * Etapa 2 – borrar marca (soft delete por defecto, hard delete opcional).
 * params.mode: "soft" | "hard"
 */
function handleDeleteMark_(params) {
  const ss = getActiveSS_();
  const cfg = getConfig_(ss);

  const coachId = String(params.coach_id || params.coachId || params.coach || "").trim();
  const swimmerId = String(params.swimmer_id || params.swimmerId || "").trim();
  const markId = String(params.mark_id || params.id || "").trim();
  const mode = String(params.mode || "soft").toLowerCase();
  if (!coachId || !swimmerId || !markId) return { status: "error", error: "coach_id, swimmer_id y mark_id requeridos" };

  const sheet = getSheetOrThrow_(ss, cfg.marks_sheet);
  ensureColumns_(sheet, CANONICAL_MARK_COLS);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { status: "error", error: "No hay marcas" };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = buildHeaderMap_(headers, buildAliasLookup_(HEADER_ALIASES));
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const midIdx = map["mark_id"];
  const sidIdx = map["swimmer_id"];
  const cidIdx = map["coach_id"];

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const cid = cidIdx == null ? "" : String(row[cidIdx] || "").trim();
    const sid = sidIdx == null ? "" : String(row[sidIdx] || "").trim();
    const mid = midIdx == null ? "" : String(row[midIdx] || "").trim();

    if (cid === coachId && sid === swimmerId && mid === markId) {
      const rowNum = r + 2;
      if (mode === "hard") {
        sheet.deleteRow(rowNum);
      } else {
        if (map.deleted_at != null) sheet.getRange(rowNum, map.deleted_at + 1).setValue(new Date().toISOString());
        if (map.updated_at != null) sheet.getRange(rowNum, map.updated_at + 1).setValue(new Date().toISOString());
      }
      SpreadsheetApp.flush();
      return { status: "ok" };
    }
  }

  return { status: "error", error: "mark_id no encontrado" };
}

/** ---------------- STANDARDS ENRICH ---------------- */

function buildUsaKey_(cycle, sex, age, style, dist, course, level) {
  return [
    normalizeHeader_(cycle),
    normalizeHeader_(sex),
    String(age == null ? "" : age),
    normalizeHeader_(style),
    String(dist == null ? "" : dist),
    normalizeHeader_(course),
    normalizeHeader_(level),
  ].join("|");
}

function buildCaddaKey_(year, sex, age, style, dist, course, tipoMarca) {
  return [
    String(year || ""),
    normalizeHeader_(sex),
    String(age == null ? "" : age),
    normalizeHeader_(style),
    String(dist == null ? "" : dist),
    normalizeHeader_(course),
    normalizeHeader_(tipoMarca),
  ].join("|");
}

function fetchStandard_(cfg, type, keyParts) {
  if (!cfg.standards_webapp_url) return { status: "skip" };

  const base = String(cfg.standards_webapp_url).trim();
  const qs = Object.keys(keyParts).map(k => `${encodeURIComponent(k)}=${encodeURIComponent(String(keyParts[k]))}`).join("&");
  const url = `${base}?action=lookup_standard&${qs}`;

  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const txt = res.getContentText();
    const obj = JSON.parse(txt);
    return obj;
  } catch (err) {
    return { status: "error", error: String(err && err.message ? err.message : err) };
  }
}

function enrichMarksWithStandards_(marks, swimmer, cfg) {
  if (!cfg.standards_webapp_url) return marks;

  const conv = Number(cfg.conv_scm_to_lcm || DEFAULT_CONV_SCM_TO_LCM);
  const sex = (swimmer && swimmer.genero) ? String(swimmer.genero).trim().toUpperCase() : "";
  const birth = swimmer ? swimmer.fecha_nac : null;

  // 1) Preparar keys únicas
  const needed = {};
  marks.forEach(m => {
    const style = m.estilo;
    const dist = m.distancia_m;
    const course = "LCM";
    const cycle = cfg.usa_cycle;
    const year = cfg.cadda_year;

    const markDate = parseISODate_(m.fecha) || parseISODate_(m.created_at) || new Date();
    const ageRef = (m.edad_ref != null) ? Number(m.edad_ref) : (m.age_chip != null ? Number(m.age_chip) : calcAgeAt_(birth, markDate));
    const ageKey = ageRef == null ? "" : ageRef;

    needed["USA|A|" + buildUsaKey_(cycle, sex, ageKey, style, dist, course, "A")] = {
      type: "USA", level: "A", cycle, sex, age: ageKey, style, dist, course
    };
    needed["USA|AA|" + buildUsaKey_(cycle, sex, ageKey, style, dist, course, "AA")] = {
      type: "USA", level: "AA", cycle, sex, age: ageKey, style, dist, course
    };
    needed["CADDA|MIN|" + buildCaddaKey_(year, sex, ageKey, style, dist, course, "Minima Nacional")] = {
      type: "CADDA", year, sex, age: ageKey, style, dist, course, tipo_marca: "Minima Nacional"
    };
  });

  // 2) Fetch + cache (CacheService)
  const cache = CacheService.getScriptCache();
  const fetched = {};

  Object.keys(needed).forEach(k => {
    const cached = cache.get(k);
    if (cached) {
      try { fetched[k] = JSON.parse(cached); return; } catch (_) {}
    }
    const spec = needed[k];

    let result;
    if (spec.type === "USA") {
      result = fetchStandard_(cfg, "USA", {
        source: "usa",
        ciclo: spec.cycle,
        genero: spec.sex,
        edad: spec.age,
        estilo: spec.style,
        distancia_m: spec.dist,
        curso: spec.course,
        nivel: spec.level,
      });
    } else {
      result = fetchStandard_(cfg, "CADDA", {
        source: "cadda",
        año: spec.year,
        genero: spec.sex,
        categoria: categoryFromAge_(Number(spec.age)),
        estilo: spec.style,
        distancia_m: spec.dist,
        curso: spec.course,
        tipo_marca: spec.tipo_marca,
      });
    }

    fetched[k] = result;
    try { cache.put(k, JSON.stringify(result), 60 * 60); } catch (_) {}
  });

  // 3) Enriquecer marcas
  return marks.map(m => {
    const out = Object.assign({}, m);

    // edad_ref/categoria_ref
    const markDate = parseISODate_(out.fecha) || parseISODate_(out.created_at) || new Date();
    const ageRef = (out.edad_ref != null) ? Number(out.edad_ref) : (out.age_chip != null ? Number(out.age_chip) : calcAgeAt_(birth, markDate));
    if (ageRef != null) out.edad_ref = ageRef;
    if (!out.categoria_ref) out.categoria_ref = categoryFromAge_(ageRef);

    // equivalencia a LCM
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

    // keys
    const cycle = cfg.usa_cycle;
    const year = cfg.cadda_year;
    const course = "LCM";
    const style = out.estilo;
    const dist = out.distancia_m;

    const usaAKey = "USA|A|" + buildUsaKey_(cycle, sex, ageRef, style, dist, course, "A");
    const usaAAKey = "USA|AA|" + buildUsaKey_(cycle, sex, ageRef, style, dist, course, "AA");
    const cKey = "CADDA|MIN|" + buildCaddaKey_(year, sex, ageRef, style, dist, course, "Minima Nacional");

    const usaA = fetched[usaAKey];
    const usaAA = fetched[usaAAKey];
    const cadda = fetched[cKey];

    // USA
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

    // CADDA
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

    // brechas
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

    // nivel (AA > A > CADDA > —)
    out.nivel = "—";
    if (equivLCM != null) {
      if (usa.AA_s != null && equivLCM <= usa.AA_s) out.nivel = "AA";
      else if (usa.A_s != null && equivLCM <= usa.A_s) out.nivel = "A";
      else if (caddaObj && caddaObj.tiempo_s != null && equivLCM <= caddaObj.tiempo_s) out.nivel = "AR";
    }

    return out;
  });
}
