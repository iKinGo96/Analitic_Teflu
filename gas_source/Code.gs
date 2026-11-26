/**
 * Backend del tablero de tendencias de maquinaria.
 *
 * Fuente predeterminada:
 *   Spreadsheet: 1PYRKlyaVaPOpTnXQ2WPZydkEOX_E_4bAlUz2XXeT9cY
 *   Hoja: H. MAQ
 *
 * Convencion especial de fechas detectada en la fuente:
 *   - el dia 25 representa el ejercicio 2025;
 *   - el dia 26 representa el ejercicio 2026.
 *
 * La convencion es configurable mediante propiedades del script:
 *   HMAQ_DATE_MODE = DAY_MARKER | CALENDAR_YEAR
 *   HMAQ_DATE_DAY_YEAR_MAP = {"25":2025,"26":2026}
 *
 * Tambien se pueden configurar:
 *   HMAQ_SPREADSHEET_ID, HMAQ_SHEET_NAME, HMAQ_CACHE_SECONDS,
 *   HMAQ_OEE_TARGET.
 *
 * Todos los porcentajes de salida se expresan como decimales (0 a 1).
 */

var APP_CONFIG_ = Object.freeze({
  spreadsheetId: '1PYRKlyaVaPOpTnXQ2WPZydkEOX_E_4bAlUz2XXeT9cY',
  sheetName: 'H. MAQ',
  defaultTimeZone: 'America/Mexico_City',
  cacheSeconds: 21600,
  oeeTarget: 0.85,
  cacheVersion: 'hmaq-v12-cdem-ot',
  dateMode: 'DAY_MARKER',
  dateDayYearMap: Object.freeze({ '25': 2025, '26': 2026 })
});

var COLUMN_HEADERS_ = Object.freeze({
  machine: 'MAQUINA',
  month: 'MES',
  availableMinutes: 'MINUTOS DISPONIBLES',
  scheduledMinutes: 'MINUTOS PROGRAMADOS',
  occupancy: 'OCUPACION REAL',
  rate: 'RATE',
  target: 'TARGET M',
  productionTotal: 'PRODUCCION TOTAL',
  maintenance: 'MIN MTTO',
  adjustment: 'MIN AJU',
  engineering: 'MIN ING',
  humanResources: 'MIN RH',
  qualityDowntime: 'MIN CALD',
  logistics: 'MIN LOG',
  warehouse: 'MIN ALM',
  productionDowntime: 'MIN PROD',
  scrapAdjustment: 'SCRAP AJUSTE',
  scrapMaintenance: 'SCRAP MTTO',
  scrapSupplier: 'SCRAP PROV',
  scrapEngineering: 'SCRAP ING',
  scrapProcess: 'SCRAP PA',
  scrapLogistics: 'LOG',
  scrapLabor: 'SCRAP MO',
  process: 'PROCESO',
  totalDowntime: 'TOTAL MIN PARO',
  availability: 'DISPONIBILIDAD',
  efficiency: 'EFICIENCIA',
  quality: 'CALIDAD',
  oee: 'OEE'
});

var DOWNTIME_FIELDS_ = Object.freeze([
  Object.freeze({ field: 'maintenance', key: 'mtto', label: 'Mantenimiento' }),
  Object.freeze({ field: 'adjustment', key: 'ajuste', label: 'Ajustes' }),
  Object.freeze({ field: 'engineering', key: 'ingenieria', label: 'Ingeniería' }),
  Object.freeze({ field: 'humanResources', key: 'rh', label: 'Recursos humanos' }),
  Object.freeze({ field: 'qualityDowntime', key: 'calidad', label: 'Calidad' }),
  Object.freeze({ field: 'logistics', key: 'logistica', label: 'Logística' }),
  Object.freeze({ field: 'warehouse', key: 'almacen', label: 'Almacén' })
]);

var MONTH_LABELS_ = Object.freeze([
  'ene', 'feb', 'mar', 'abr', 'may', 'jun',
  'jul', 'ago', 'sep', 'oct', 'nov', 'dic'
]);

/**
 * Punto de entrada del Web App.
 *
 * Sin parametros sirve Index.html. Para consumir JSON directamente:
 *   ?api=dashboard
 *   ?api=trend
 * Los filtros opcionales son process, machine, fromPeriod y toPeriod.
 *
 * @param {Object=} e Evento HTTP de Apps Script.
 * @return {GoogleAppsScript.HTML.HtmlOutput|GoogleAppsScript.Content.TextOutput}
 */
function doGet(e) {
  var parameters = e && e.parameter ? e.parameter : {};
  var api = cleanString_(parameters.api).toLowerCase();

  if (api === 'dashboard' || api === 'trend') {
    try {
      var filters = filtersFromQuery_(parameters);
      var payload = api === 'trend' ? getTrendData(filters) : getDashboardData(filters);
      return jsonOutput_(payload);
    } catch (error) {
      console.error('Error en doGet JSON: ' + errorStack_(error));
      return jsonOutput_({
        ok: false,
        error: {
          message: publicErrorMessage_(error)
        }
      });
    }
  }

  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Tendencias de máquinas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    console.error('Error al servir Index.html: ' + errorStack_(error));
    throw new Error(
      'No fue posible abrir la aplicación. Verifica que exista el archivo Index.html. Detalle: ' +
      publicErrorMessage_(error)
    );
  }
}

/**
 * Permite incluir parciales HTML desde Index.html.
 * Solo acepta nombres simples para evitar rutas no deseadas.
 *
 * @param {string} filename Nombre de archivo sin extension.
 * @return {string} Contenido del archivo HTML.
 */
function include(filename) {
  var safeName = cleanString_(filename);
  if (!/^[A-Za-z0-9_-]+$/.test(safeName)) {
    throw new Error('Nombre de parcial HTML no válido.');
  }
  return HtmlService.createHtmlOutputFromFile(safeName).getContent();
}

/**
 * API principal para google.script.run.
 *
 * Contrato de salida:
 * {
 *   meta, filters, selection, kpis, trend, downtime, ranking, insights
 * }
 *
 * @param {Object=} filters Filtros process, machine, fromPeriod y toPeriod.
 * @return {Object} Datos listos para gráficas y tarjetas KPI.
 */
function getDashboardData(filters) {
  try {
    var config = runtimeConfig_();
    var safeFilters = filters || {};
    var cacheKey = 'dash-resp-' + datasetCacheKey_(config) + '-' + hashString_(JSON.stringify(safeFilters));
    var cache = CacheService.getScriptCache();
    
    if (!safeFilters.forceRefresh) {
      var cached = cache.get(cacheKey);
      if (cached) {
        try {
          if (cached.indexOf('gzip:') === 0) {
            cached = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(cached.slice(5)), 'application/gzip')).getDataAsString('UTF-8');
          }
          var parsed = JSON.parse(cached);
          if (parsed && parsed.meta) return parsed;
        } catch (e) {
          console.warn('Dashboard cache inválido, se reconstruirá.');
        }
      }
    }

    var dashboard = buildDashboard_(safeFilters);
    
    try {
      var payload = JSON.stringify(dashboard);
      var byteLength = Utilities.newBlob(payload).getBytes().length;
      if (byteLength > 95000) {
        payload = 'gzip:' + Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(payload, 'application/json')).getBytes());
        byteLength = payload.length;
      }
      if (byteLength <= 95000) {
        cache.put(cacheKey, payload, config.cacheSeconds);
      } else {
        console.warn('La respuesta del dashboard supera 95 KB; se omite el caché.');
      }
    } catch (cacheError) {
      console.warn('No se pudo guardar la respuesta en caché: ' + publicErrorMessage_(cacheError));
    }
    
    return dashboard;
  } catch (error) {
    console.error('Error en getDashboardData: ' + errorStack_(error));
    throw new Error('No fue posible generar el tablero. Detalle: ' + publicErrorMessage_(error));
  }
}

/**
 * API ligera de tendencias para google.script.run.
 *
 * @param {Object=} filters Filtros process, machine, fromPeriod y toPeriod.
 * @return {Object} Metadatos, selección, KPI y serie temporal.
 */
function getTrendData(filters) {
  var dashboard = getDashboardData(filters || {});
  return {
    meta: dashboard.meta,
    selection: dashboard.selection,
    kpis: dashboard.kpis,
    trend: dashboard.trend,
    cdemPareto: dashboard.cdemPareto
  };
}

/**
 * Construye el objeto completo del tablero.
 *
 * @param {Object} requestedFilters Filtros solicitados.
 * @return {Object}
 * @private
 */
function buildDashboard_(requestedFilters) {
  var dataset = loadDataset_(requestedFilters && requestedFilters.forceRefresh === true);
  var allRows = dataset.rows;
  var filterOptions = buildFilterOptions_(allRows, requestedFilters, dataset.cdemSummary);
  var selection = resolveSelection_(requestedFilters, filterOptions, allRows);

  // La lista de máquinas se adapta al proceso seleccionado, pero la selección
  // sigue siendo explícita para que el frontend pueda mostrar un estado vacío.
  filterOptions.machines = uniqueSorted_(allRows
    .filter(function (row) {
      return !selection.process || row.process === selection.process;
    })
    .map(function (row) { return row.machine; }));

  var selectedRows = allRows.filter(function (row) {
    return (!selection.process || row.process === selection.process) &&
      (!selection.machine || row.machine === selection.machine) &&
      row.periodSort >= selection.fromPeriodSort &&
      row.periodSort <= selection.toPeriodSort;
  });

  var trend = buildTrend_(selectedRows, selection.fromPeriod, selection.toPeriod);
  var latestSelectedRows = selectedRows.filter(function (row) {
    return row.periodSort === selection.toPeriodSort;
  });
  var kpiPeriodKey = selection.toPeriod;
  if (!latestSelectedRows.length && selectedRows.length) {
    var maxPresentSort = Math.max.apply(null, selectedRows.map(function (r) { return r.periodSort; }));
    latestSelectedRows = selectedRows.filter(function (r) { return r.periodSort === maxPresentSort; });
    if (latestSelectedRows.length) {
      kpiPeriodKey = latestSelectedRows[0].period;
    }
  }
  var kpis = aggregateRows_(latestSelectedRows);
  kpis.periodLabel = periodLabelFromKey_(kpiPeriodKey);
  kpis.deltaOee = trend.length > 1 && isFiniteNumber_(trend[trend.length - 1].oee) &&
    isFiniteNumber_(trend[trend.length - 2].oee)
    ? round_(trend[trend.length - 1].oee - trend[trend.length - 2].oee, 6)
    : null;

  var downtime = buildDowntimePareto_(selectedRows);

  var cdemSummary = dataset.cdemSummary || [];
  var pPie = requestedFilters && requestedFilters.downtimePie ? requestedFilters.downtimePie : {};
  var cdemPeriods = uniqueSorted_(cdemSummary.map(function(row) { return row.p; }));
  var latestDowntimePeriod = cdemPeriods.length ? cdemPeriods[cdemPeriods.length - 1] : selection.toPeriod;
  var fromP = pPie.fromPeriod || latestDowntimePeriod;
  var toP = pPie.toPeriod || latestDowntimePeriod;
  var fromSort = fromP ? parseInt(String(fromP).replace('-', ''), 10) : 0;
  var toSort = toP ? parseInt(String(toP).replace('-', ''), 10) : 999999;
  if (fromSort > toSort) {
    var tmpS = fromSort; fromSort = toSort; toSort = tmpS;
    var tmpP = fromP; fromP = toP; toP = tmpP;
  }
  var downtimePieCdem = cdemSummary.filter(function(row) {
    return (!selection.process || row.proc === selection.process) && 
           (!selection.machine || row.m === selection.machine) && 
           row.ps >= fromSort && row.ps <= toSort;
  });

  var pieGroups = {};
  downtimePieCdem.forEach(function(r) {
    var key = r.dept || 'Sin Clasificar';
    if (!pieGroups[key]) pieGroups[key] = 0;
    pieGroups[key] += r.min;
  });

  var downtimePie = Object.keys(pieGroups).map(function(key) {
    return { label: key, minutes: pieGroups[key] };
  });
  downtimePie.sort(function(a, b) { return b.minutes - a.minutes; });

  var availability = buildAvailabilityView_(dataset.cdemSummary || [], dataset.cdemStatus, requestedFilters.availability);
  var scrap = buildScrapView_(dataset.scrapSummary || [], dataset.scrapStatus, requestedFilters.scrap);
  var ranking = buildRanking_(allRows, selection);
  
  var historicalRows = allRows.filter(function (row) {
    return (!selection.process || row.process === selection.process) &&
      (!selection.machine || row.machine === selection.machine) &&
      row.periodSort <= selection.toPeriodSort;
  });
  var last6Periods = periodRange_(shiftPeriod_(selection.toPeriod, -5), selection.toPeriod);
  var last6Rows = historicalRows.filter(function(row) { return last6Periods.indexOf(row.period) !== -1; });
  var last6Aggregate = aggregateRows_(last6Rows);
  last6Aggregate.count = uniqueSorted_(last6Rows.filter(hasOperationalBase_).map(function (row) { return row.period; })).length;

  var insights = buildInsights_(trend, kpis, downtime, ranking, selectedRows.length, last6Aggregate);
  var critical = buildCriticalMachines_(allRows, selection, dataset.oeeTarget, requestedFilters.criticalWindowMonths);
  critical.decisions = readCriticalDecisions_();
  critical.access = criticalAccess_();

  return {
    meta: {
      apiVersion: '1.0.0',
      oeeTarget: dataset.oeeTarget,
      updatedAt: dataset.loadedAt,
      generatedAt: new Date().toISOString(),
      sourceRows: dataset.sourceRows,
      filteredRows: selectedRows.length,
      skippedRows: dataset.skippedRows,
      source: {
        spreadsheetId: dataset.spreadsheetId,
        sheetName: dataset.sheetName
      },
      dateConvention: dataset.dateConvention,
      warnings: dataset.warnings
    },
    filters: filterOptions,
    selection: {
      process: selection.process,
      machine: selection.machine,
      fromPeriod: selection.fromPeriod,
      toPeriod: selection.toPeriod
    },
    kpis: kpis,
    trend: trend,
    downtime: downtime,
    downtimePie: downtimePie,
    downtimePieSelection: { fromPeriod: fromP, toPeriod: toP },
    cdemPareto: availability.rows,
    availability: availability,
    scrapPareto: scrap.rows,
    scrap: scrap,
    ranking: ranking,
    insights: insights,
    critical: critical,
    otMap: dataset.otMap,
    deviations: dataset.deviations || [],
    deviationsStatus: dataset.deviationsStatus || {}
  };
}

/**
 * Prioridad histórica, independiente del filtro de máquina y del inicio del
 * resumen. Usa meses calendario hasta el corte elegido; nunca rellena huecos.
 * Es una regla operativa explicable, no una predicción de fallas.
 */
function buildCriticalMachines_(allRows, selection, target, requestedWindow) {
  var isAll = requestedWindow === 'all';
  var months = isAll ? 'all' : ([3, 6, 12].indexOf(Number(requestedWindow)) >= 0 ? Number(requestedWindow) : 6);
  
  // Análisis de toda la planta: no se filtra por proceso ni por máquina.
  // El proceso seleccionado solo aplica a tendencias y KPIs, no a la identificación de críticas.
  var peers = allRows.filter(function (row) {
    return row.period <= selection.toPeriod;
  });

  var from;
  if (isAll) {
    from = peers.reduce(function(min, row) { return row.period < min ? row.period : min; }, peers[0] ? peers[0].period : selection.toPeriod);
  } else {
    from = shiftPeriod_(selection.toPeriod, 1 - months);
  }
  
  var minimum = Math.max(3, Math.ceil((isAll ? periodRange_(from, selection.toPeriod).length : months) / 2));
  var candidates = uniqueSorted_(peers.map(function (row) { return row.machine; })).map(function (machine) {
    var rows = peers.filter(function (row) { return row.machine === machine; });
    var points = buildTrend_(rows.filter(function (row) { return row.period >= from; }), from, selection.toPeriod);
    return analyzeCriticalTrend_(machine, rows[rows.length - 1].process, points, target, minimum);
  });
  candidates.sort(function (a, b) {
    return (b.score === null ? -1 : b.score) - (a.score === null ? -1 : a.score) || a.machine.localeCompare(b.machine);
  });
  return {
    methodVersion: 'trend-v1', windowMonths: months, fromPeriod: from,
    toPeriod: selection.toPeriod, process: selection.process, target: target,
    minimumValidMonths: minimum, candidates: candidates
  };
}

function analyzeCriticalTrend_(machine, process, points, target, minimum) {
  var valid = [];
  points.forEach(function (point, index) {
    if (isFiniteNumber_(point.oee)) valid.push({ point: point, x: index, weight: index + 1 });
  });
  var enough = valid.length >= minimum;
  var weightSum = valid.reduce(function (sum, item) { return sum + item.weight; }, 0);
  var weighted = weightSum ? valid.reduce(function (sum, item) { return sum + item.point.oee * item.weight; }, 0) / weightSum : null;
  var below = valid.filter(function (item) { return item.point.oee < target; }).length;
  var streak = 0;
  for (var i = points.length - 1; i >= 0 && isFiniteNumber_(points[i].oee) && points[i].oee < target; i--) streak++;
  var meanX = valid.length ? valid.reduce(function (sum, item) { return sum + item.x; }, 0) / valid.length : 0;
  var meanY = valid.length ? valid.reduce(function (sum, item) { return sum + item.point.oee; }, 0) / valid.length : 0;
  var numerator = 0;
  var denominator = 0;
  valid.forEach(function (item) {
    numerator += (item.x - meanX) * (item.point.oee - meanY);
    denominator += Math.pow(item.x - meanX, 2);
  });
  var slope = enough && denominator > 0 ? numerator / denominator : null;
  var declines = 0;
  points.forEach(function (point, index) {
    if (index && isFiniteNumber_(point.oee) && isFiniteNumber_(points[index - 1].oee) && point.oee < points[index - 1].oee - 0.00005) declines++;
  });
  var persistence = valid.length ? below / valid.length : 0;
  // 50% brecha ponderada (satura a 20 pp), 30% caída (3 pp/mes),
  // 20% persistencia. La puntuación no equivale al OEE ni a probabilidad.
  var score = enough ? 50 * Math.min(1, Math.max(0, target - weighted) / .20) +
    30 * Math.min(1, Math.max(0, -slope) / .03) + 20 * persistence : null;
  var persistent = enough && below >= 2 && persistence >= .5 && weighted < target;
  var declining = enough && slope <= -.01 && declines >= 2;
  
  // Evaluar también el periodo reciente (hasta 6 meses) para no diluir crisis actuales en ventanas largas
  var recentValid = valid.filter(function (item) { return item.x >= points.length - 6; });
  var recentBelow = recentValid.filter(function (item) { return item.point.oee < target; }).length;
  var recentPersistence = recentValid.length ? recentBelow / recentValid.length : 0;
  var recentMeanX = recentValid.length ? recentValid.reduce(function (sum, item) { return sum + item.x; }, 0) / recentValid.length : 0;
  var recentMeanY = recentValid.length ? recentValid.reduce(function (sum, item) { return sum + item.point.oee; }, 0) / recentValid.length : 0;
  var recentNum = 0, recentDen = 0;
  recentValid.forEach(function (item) {
    recentNum += (item.x - recentMeanX) * (item.point.oee - recentMeanY);
    recentDen += Math.pow(item.x - recentMeanX, 2);
  });
  var recentSlope = recentValid.length >= 3 && recentDen > 0 ? recentNum / recentDen : null;
  var recentWeightSum = recentValid.reduce(function (sum, item) { return sum + item.weight; }, 0);
  var recentWeighted = recentWeightSum ? recentValid.reduce(function (sum, item) { return sum + item.point.oee * item.weight; }, 0) / recentWeightSum : null;
  var recentDeclines = 0;
  recentValid.forEach(function (item, idx) {
    if (idx > 0 && item.x === recentValid[idx - 1].x + 1 && item.point.oee < recentValid[idx - 1].point.oee - 0.00005) recentDeclines++;
  });
  
  var recentPersistent = enough && recentValid.length >= 3 && recentBelow >= 2 && recentPersistence >= .5 && recentWeighted < target;
  var recentDeclining = enough && recentValid.length >= 3 && recentSlope <= -.01 && recentDeclines >= 2;
  var autoCritical = persistent || declining || recentPersistent || recentDeclining;

  var reasons = [];
  if (!enough) reasons.push('Historial insuficiente: ' + valid.length + ' de ' + points.length + ' meses con OEE; se requieren ' + minimum + '.');
  if (persistent) reasons.push(below + ' de ' + valid.length + ' meses con OEE por debajo de la meta; promedio reciente ' + (weighted * 100).toFixed(1) + '%.');
  else if (recentPersistent) reasons.push('Rezago reciente: ' + recentBelow + ' de los últimos ' + recentValid.length + ' meses evaluados por debajo de la meta.');
  if (declining) reasons.push('Tendencia descendente de ' + Math.abs(slope * 100).toFixed(2) + ' pp/mes, con ' + declines + ' caídas entre meses consecutivos.');
  else if (recentDeclining) reasons.push('Caída reciente aguda: tendencia de ' + Math.abs(recentSlope * 100).toFixed(2) + ' pp/mes en el último semestre.');
  if (streak >= 2) reasons.push(streak + ' meses consecutivos por debajo de la meta hasta el corte.');
  if (enough && !autoCritical) reasons.push('No cumple los criterios automáticos de rezago sostenido o caída continuada, ni siquiera en los últimos 6 meses.');
  if (points.length > valid.length) reasons.push((points.length - valid.length) + ' meses sin OEE válido; no se trataron como cero.');
  var latest = points.length ? points[points.length - 1] : {};
  if (!isFiniteNumber_(latest.oee)) reasons.push('Sin OEE válido en el mes de corte; la evaluación usa solo el historial disponible.');
  var components = ['availability', 'efficiency', 'quality'].map(function (key) {
    var items = valid.filter(function (item) { return isFiniteNumber_(item.point[key]); });
    var weights = items.reduce(function (sum, item) { return sum + item.weight; }, 0);
    return { key: key, value: weights ? items.reduce(function (sum, item) { return sum + item.point[key] * item.weight; }, 0) / weights : null };
  }).filter(function (item) { return isFiniteNumber_(item.value); }).sort(function (a, b) { return a.value - b.value; });
  return {
    machine: machine, process: process, eligible: enough, autoCritical: autoCritical,
    score: nullableRound_(score, 1), weightedOee: nullableRound_(weighted, 6),
    slopePp: slope === null ? null : round_(slope * 100, 3),
    latestOee: isFiniteNumber_(latest.oee) ? latest.oee : null,
    lastDataPeriod: valid.length ? valid[valid.length - 1].point.period : null,
    validMonths: valid.length, missingMonths: points.length - valid.length,
    belowTargetMonths: below, consecutiveBelow: streak,
    weakestComponent: components.length ? { key: components[0].key, value: round_(components[0].value, 6) } : null,
    reasons: reasons,
    points: points.map(function (point) { return { period: point.period, label: point.label, occupancy: point.occupancy, oee: point.oee, availability: point.availability, efficiency: point.efficiency, quality: point.quality }; })
  };
}

function criticalDecisionKey_() {
  var config = runtimeConfig_();
  return 'HMAQ_CRITICAL_DECISIONS_V1:' + config.spreadsheetId + ':' + config.sheetName;
}

function readCriticalDecisions_() {
  var raw = PropertiesService.getScriptProperties().getProperty(criticalDecisionKey_());
  if (!raw) return {};
  var parsed = JSON.parse(raw);
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error('No se pudieron leer las decisiones compartidas. Contacta al administrador.');
  }
  return parsed;
}

/** La identidad debe ser del visitante, nunca la del propietario efectivo. */
function criticalAccess_() {
  return { canEdit: false, message: 'Modo lectura. Haz clic en "Revisar y gestionar" e ingresa credenciales para editar.' };
}

function verifyCredentials(user, pass) {
  var props = PropertiesService.getScriptProperties();
  var expectedUser = props.getProperty('HMAQ_AUTH_USER') || 'admin';
  var expectedPass = props.getProperty('HMAQ_AUTH_PASS') || 'admin123';
  
  if (user === expectedUser && pass === expectedPass) {
    return { ok: true };
  }
  return { ok: false };
}

/** Ejecutar una sola vez desde el editor de Apps Script con tu propia cuenta.
 * El sufijo _ impide llamarla con google.script.run desde la aplicación web.
 * No cambia una lista existente: esa modificación corresponde al editor.
 */
function configurarMiCuentaAdministradora_() {
  var properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('HMAQ_ADMIN_EMAILS')) throw new Error('La lista de administradores ya existe. Revísala en Propiedades del script.');
  var email = cleanString_(Session.getActiveUser().getEmail()).toLowerCase();
  if (!email || email.indexOf('@') < 1) throw new Error('Google no pudo identificar tu cuenta. Ejecuta esta función desde el editor con tu sesión de Google.');
  properties.setProperty('HMAQ_ADMIN_EMAILS', email);
  console.log('Cuenta administradora configurada: ' + email);
}

/** Guarda solo decisiones manuales. No escribe en la hoja ni en su caché. */
function setCriticalDecision(request) {
  request = request || {};
  if (!verifyCredentials(request.authUser, request.authPass).ok) {
    throw new Error('No autorizado. Credenciales incorrectas.');
  }
  request = request || {};
  var machine = cleanString_(request.machine);
  var status = request.status;
  if (['pinned', 'dismissed', 'auto'].indexOf(status) < 0 ||
      typeof request.expectedRevision !== 'number' || request.expectedRevision < 0 ||
      !Number.isInteger(request.expectedRevision)) throw new Error('Decisión no válida.');
  if (!loadDataset_().rows.some(function (row) { return row.machine === machine; })) {
    throw new Error('La máquina no existe en la fuente configurada.');
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('Hay otra decisión en curso. Intenta guardar de nuevo.');
  try {
    var decisions = readCriticalDecisions_();
    var previous = Object.prototype.hasOwnProperty.call(decisions, machine) ? decisions[machine] : null;
    var revision = previous ? previous.revision : 0;
    if (request.expectedRevision !== revision) {
      return { ok: false, conflict: true, decisions: decisions };
    }
    // Mantener revisión incluso al volver a automático evita escrituras antiguas.
    Object.defineProperty(decisions, machine, { value: {
      status: status, revision: revision + 1, updatedAt: new Date().toISOString()
    }, enumerable: true, configurable: true, writable: true });
    PropertiesService.getScriptProperties().setProperty(criticalDecisionKey_(), JSON.stringify(decisions));
    return { ok: true, decisions: decisions };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Lee y normaliza la fuente, con caché de script para reducir lecturas.
 *
 * @param {boolean=} forceRefresh Omite la caché cuando es verdadero.
 * @return {Object}
 * @private
 */
function loadDataset_(forceRefresh) {
  var config = runtimeConfig_();
  var cache = CacheService.getScriptCache();
  var cacheKey = datasetCacheKey_(config);
  if (forceRefresh) {
    cache.remove(cacheKey);
  }
  var restored = forceRefresh ? null : getDatasetFromCache_(cache, cacheKey);
  if (restored) {
    return restored;
  }

  var lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    hasLock = lock.tryLock(5000);
    if (hasLock) {
      var afterLock = forceRefresh ? null : getDatasetFromCache_(cache, cacheKey);
      if (afterLock) {
        return afterLock;
      }
    }

    var dataset = readDataset_(config);
    putDatasetInCache_(cache, cacheKey, dataset, config.cacheSeconds);
    return dataset;
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Lee la hoja y valida los 29 encabezados conocidos.
 *
 * @param {Object} config Configuración efectiva.
 * @return {Object}
 * @private
 */
function readDataset_(config) {
  var spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  } catch (error) {
    throw new Error(
      'No se pudo abrir el archivo fuente. Confirma el ID y los permisos de lectura. ' +
      publicErrorMessage_(error)
    );
  }

  var sheet = spreadsheet.getSheetByName(config.sheetName);
  if (!sheet) {
    throw new Error('No existe la hoja "' + config.sheetName + '" en el archivo fuente.');
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    throw new Error('La hoja "' + config.sheetName + '" no contiene filas de datos.');
  }

  var values;
  try {
    values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  } catch (error) {
    throw new Error('No se pudo leer el rango de datos. ' + publicErrorMessage_(error));
  }

  var indexes = resolveColumnIndexes_(values[0]);
  var rows = [];
  var warnings = [];
  var skippedRows = 0;
  var timeZone = spreadsheet.getSpreadsheetTimeZone() || config.timeZone;

  for (var i = 1; i < values.length; i += 1) {
    var sourceRow = values[i];
    if (isEmptyRow_(sourceRow)) {
      continue;
    }

    var machine = cleanString_(sourceRow[indexes.machine]);
    var process = cleanString_(sourceRow[indexes.process]);
    var period = parsePeriod_(sourceRow[indexes.month], config, timeZone);

    if (!machine || !process || !period) {
      skippedRows += 1;
      if (warnings.length < 10) {
        warnings.push(
          'Fila ' + (i + 1) + ' omitida: falta máquina, proceso o un MES válido.'
        );
      }
      continue;
    }

    var row = {
      machine: machine,
      process: process,
      period: period.value,
      periodLabel: period.label,
      periodSort: period.sort,
      availableMinutes: numberOrNull_(sourceRow[indexes.availableMinutes], false),
      scheduledMinutes: numberOrNull_(sourceRow[indexes.scheduledMinutes], false),
      occupancy: numberOrNull_(sourceRow[indexes.occupancy], true),
      rate: numberOrNull_(sourceRow[indexes.rate], false),
      target: numberOrNull_(sourceRow[indexes.target], false),
      productionTotal: numberOrNull_(sourceRow[indexes.productionTotal], false),
      maintenance: numberOrNull_(sourceRow[indexes.maintenance], false),
      adjustment: numberOrNull_(sourceRow[indexes.adjustment], false),
      engineering: numberOrNull_(sourceRow[indexes.engineering], false),
      humanResources: numberOrNull_(sourceRow[indexes.humanResources], false),
      qualityDowntime: numberOrNull_(sourceRow[indexes.qualityDowntime], false),
      logistics: numberOrNull_(sourceRow[indexes.logistics], false),
      warehouse: numberOrNull_(sourceRow[indexes.warehouse], false),
      productionDowntime: numberOrNull_(sourceRow[indexes.productionDowntime], false),
      totalDowntime: numberOrNull_(sourceRow[indexes.totalDowntime], false),
      availability: numberOrNull_(sourceRow[indexes.availability], true),
      efficiency: numberOrNull_(sourceRow[indexes.efficiency], true),
      quality: numberOrNull_(sourceRow[indexes.quality], true),
      oee: numberOrNull_(sourceRow[indexes.oee], true),
      totalScrap: sumNullable_([
        numberOrNull_(sourceRow[indexes.scrapAdjustment], false),
        numberOrNull_(sourceRow[indexes.scrapMaintenance], false),
        numberOrNull_(sourceRow[indexes.scrapSupplier], false),
        numberOrNull_(sourceRow[indexes.scrapEngineering], false),
        numberOrNull_(sourceRow[indexes.scrapProcess], false),
        numberOrNull_(sourceRow[indexes.scrapLogistics], false),
        numberOrNull_(sourceRow[indexes.scrapLabor], false)
      ])
    };

    // En la fuente, algunas filas operativas usan guiones en todas las
    // causas de scrap y registran calidad de 100%. Ese indicador confirma
    // cero piezas rechazadas; sin esa evidencia se conserva el dato ausente.
    if (row.totalScrap === null && row.productionTotal > 0 && row.quality === 1) {
      row.totalScrap = 0;
    }
    rows.push(row);
  }

  if (!rows.length) {
    throw new Error(
      'No se encontraron filas válidas en "' + config.sheetName + '" después de normalizar los datos.'
    );
  }

  rows.sort(compareRows_);

  var cdem = readCdem_(spreadsheet, rows, config, timeZone);
  var scrap = readScrap_(spreadsheet, rows, config, timeZone);
  var otMap = readOtMtto_(spreadsheet);
  var deviations = readDesviaciones_(spreadsheet, config, timeZone);
  var actions = readAcciones_(spreadsheet, config);

  return {
    spreadsheetId: config.spreadsheetId,
    sheetName: config.sheetName,
    oeeTarget: config.oeeTarget,
    loadedAt: new Date().toISOString(),
    sourceRows: values.slice(1).filter(function (row) { return !isEmptyRow_(row); }).length,
    skippedRows: skippedRows,
    warnings: warnings,
    dateConvention: {
      mode: config.dateMode,
      dayYearMap: config.dateDayYearMap,
      description: config.dateMode === 'DAY_MARKER'
        ? 'El día del MES actúa como marcador del año de reporte.'
        : 'Se usa el año calendario almacenado en la celda MES.'
    },
    rows: rows,
    cdemSummary: cdem.rows,
    cdemStatus: cdem.status,
    scrapSummary: scrap.rows,
    scrapStatus: scrap.status,
    otMap: otMap,
    deviations: deviations.rows,
    deviationsStatus: deviations.status,
    actions: actions.rows,
    actionsActivities: actions.activities || [],
    actionsStatus: actions.status
  };
}

/** Diagnóstico explícito: una lectura fallida nunca equivale a cero paros. */
function readCdem_(spreadsheet, machineRows, config, timeZone) {
  var sourceName = 'C.DEM';
  try {
    var sheet = spreadsheet.getSheetByName('C.DEM') || spreadsheet.getSheetByName('C. DEM');
    if (!sheet) throw new Error('No existe la hoja C.DEM (ni su nombre anterior C. DEM).');
    sourceName = sheet.getName();
    if (!sheet.getLastRow()) throw new Error('La hoja ' + sourceName + ' no contiene encabezados.');
    var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    var result = normalizeCdem_(values, machineRows, config, timeZone);
    result.status.source = sourceName;
    return result;
  } catch (error) {
    return { rows: [], status: { state: 'error', source: sourceName, message: 'No fue posible leer ' + sourceName + ': ' + publicErrorMessage_(error) } };
  }
}

function normalizeCdem_(values, machineRows, config, timeZone) {
  var headers = (values[0] || []).map(normalizeHeader_);
  var indexes = {};
  ['MAQUINA', 'MES DE CORTE', 'CAUSA GENERAL', 'CAUSA ESPECIFICA', 'TOTAL MIN'].forEach(function (header) {
    var index = headers.indexOf(header);
    if (index < 0) throw new Error('Falta el encabezado ' + header + '.');
    indexes[header] = index;
  });
  var processIndex = headers.indexOf('PROCESO');
  var deptoIndex = headers.indexOf('DEPTO GENERO PARO');
  if (deptoIndex < 0) {
    for (var d = 0; d < headers.length; d++) {
      var hd = headers[d] || '';
      if (hd === 'DEPTO' || hd === 'DEPARTAMENTO' || hd.indexOf('DEPTO') >= 0 || hd === 'PARO') {
        deptoIndex = d;
        break;
      }
    }
  }

  var parteIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i] || '';
    if (
      h === 'NUMERO DE PARTE' ||
      h === 'NO DE PARTE' ||
      h === 'NUM DE PARTE' ||
      h === 'NUMERO PARTE' ||
      h === 'NO PARTE' ||
      h === 'PARTE' ||
      h.indexOf('PARTE') >= 0 ||
      h.indexOf('PIEZA') >= 0
    ) {
      parteIndex = i;
      break;
    }
  }
  if (parteIndex < 0 && headers.length > 5 && (headers[0] === 'MAQUINA' || headers[1] === 'FECHA')) {
    parteIndex = 5;
  }

  var semanaIndex = -1;
  for (var s = 0; s < headers.length; s++) {
    var hs = headers[s] || '';
    if (hs === 'SEMANA' || hs === 'SEM' || hs === 'NO. SEMANA' || hs === 'SEMANA DE CORTE' || hs.indexOf('SEMANA') >= 0) {
      semanaIndex = s;
      break;
    }
  }
  if (semanaIndex < 0 && headers.length > 2 && (headers[0] === 'MAQUINA' || headers[1] === 'FECHA')) {
    semanaIndex = 2; // Columna C en C.DEM estándar
  }

  var otIndex = -1;
  for (var o = 0; o < headers.length; o++) {
    var ho = headers[o] || '';
    if (
      ho === 'ORDEN DE TRABAJO' ||
      ho === 'ORDEN TRABAJO' ||
      ho === 'NO DE ORDEN' ||
      ho === 'NO. DE ORDEN' ||
      ho === 'NO. ORDEN' ||
      ho === 'NO ORDEN' ||
      ho === 'NUMERO DE ORDEN' ||
      ho === 'ORDEN' ||
      ho === 'OT' ||
      ho.indexOf('ORDEN') >= 0 ||
      ho.indexOf('TRABAJO') >= 0
    ) {
      otIndex = o;
      break;
    }
  }
  if (otIndex < 0 && (headers.length > 10 || (values[1] && values[1].length > 10))) {
    otIndex = 10; // Columna K en C.DEM estándar
  }
  var references = Object.create(null);
  var knownProcesses = uniqueSorted_(machineRows.map(function (row) { return row.process; }));
  machineRows.forEach(function (row) {
    var key = normalizeComparable_(row.machine);
    if (!references[key]) references[key] = [];
    references[key].push(row);
  });
  var groups = Object.create(null);
  var skipped = 0, unmapped = 0, sourceRows = 0, blankCauses = 0;
  values.slice(1).forEach(function (row) {
    if (isEmptyRow_(row)) return;
    sourceRows++;
    var machine = cleanString_(row[indexes.MAQUINA]);
    var period = parsePeriod_(row[indexes['MES DE CORTE']], config, timeZone);
    var minutes = numberOrNull_(row[indexes['TOTAL MIN']], false);
    if (!machine || !period || !isFiniteNumber_(minutes) || minutes < 0) { skipped++; return; }
    var causeGen = cleanString_(row[indexes['CAUSA GENERAL']]);
    if (!causeGen) { causeGen = 'Sin causa general'; blankCauses++; }
    var causeEsp = cleanString_(row[indexes['CAUSA ESPECIFICA']]);
      if (!causeEsp) { causeEsp = 'Sin causa específica'; }
      var depto = deptoIndex >= 0 ? cleanString_(row[deptoIndex]) : 'Sin departamento';
      if (!depto) depto = 'Sin departamento';
      var rawParte = parteIndex >= 0 ? row[parteIndex] : '';
      var parte = cleanString_(rawParte);
      if (!parte || parte.toUpperCase() === 'NAN' || parte.toUpperCase() === 'NULL' || parte === '-' || parte === '0' || parte.toUpperCase() === 'UNDEFINED') {
        parte = 'Sin número de parte';
      }
      var rawSemana = semanaIndex >= 0 ? cleanString_(row[semanaIndex]) : '';
      var semana = 'Sin semana';
      if (rawSemana) {
        var semClean = rawSemana.trim().toUpperCase();
        if (/^W\d+$/.test(semClean)) {
          semana = semClean;
        } else if (/^\d+$/.test(semClean)) {
          semana = 'W' + semClean;
        } else if (semClean.indexOf('SEM') >= 0) {
          var numMatch = semClean.match(/\d+/);
          semana = numMatch ? 'W' + numMatch[0] : semClean;
        } else {
          semana = semClean;
        }
      }
      var rawOt = otIndex >= 0 ? cleanString_(row[otIndex]) : '';
      var ot = 'Sin orden de trabajo';
      if (rawOt) {
        var upOt = rawOt.toUpperCase();
        if (upOt !== 'NAN' && upOt !== 'NULL' && upOt !== 'UNDEFINED' && rawOt !== '-' && rawOt !== '0' && upOt !== 'N/A' && upOt !== 'NO') {
          ot = rawOt;
        }
      }
    var matches = references[normalizeComparable_(machine)] || [];
    if (matches.length) machine = matches[0].machine;
    // La máquina y el mes enlazan ambos catálogos: BEWO → CORTE DISCO.
    var sameMonth = matches.filter(function (item) { return item.period === period.value; });
    var processes = uniqueSorted_((sameMonth.length ? sameMonth : matches).map(function (item) { return item.process; }));
    var rawProcess = processIndex < 0 ? '' : cleanString_(row[processIndex]);
    var process = processes.length === 1 ? processes[0] : knownProcesses.filter(function (name) {
      return normalizeComparable_(name) === normalizeComparable_(rawProcess);
    })[0];
    if (!process) { process = 'Sin proceso homologado'; unmapped++; }
    var key = JSON.stringify([machine, process, period.value, causeGen, causeEsp, depto, parte, semana, ot]);
    if (!groups[key]) groups[key] = { m: machine, proc: process, p: period.value, ps: period.sort, cg: causeGen, ce: causeEsp, dept: depto, part: parte, sem: semana, ot: ot, min: 0, count: 0 };
    groups[key].min += minutes;
    groups[key].count++;
  });
  return { rows: Object.keys(groups).map(function (key) { return groups[key]; }), status: {
    state: 'ready', source: 'C.DEM', sourceRows: sourceRows, skippedRows: skipped,
    unmappedRows: unmapped, blankCauseRows: blankCauses
  } };
}

function readScrap_(spreadsheet, machineRows, config, timeZone) {
  var sourceName = 'S.SCRAP';
  try {
    var sheet = spreadsheet.getSheetByName('S.SCRAP') || spreadsheet.getSheetByName('SCRAP') || spreadsheet.getSheetByName('S. SCRAP');
    if (!sheet) throw new Error('No existe la hoja S.SCRAP.');
    sourceName = sheet.getName();
    if (!sheet.getLastRow()) throw new Error('La hoja ' + sourceName + ' no contiene encabezados.');
    var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    var result = normalizeScrap_(values, machineRows, config, timeZone);
    result.status.source = sourceName;
    return result;
  } catch (error) {
    return { rows: [], status: { state: 'error', source: sourceName, message: 'No fue posible leer ' + sourceName + ': ' + publicErrorMessage_(error) } };
  }
}

function normalizeScrap_(values, machineRows, config, timeZone) {
  var headers = (values[0] || []).map(normalizeHeader_);
  var indexes = {};
  ['MAQUINA', 'PIEZAS S', 'MODO DE FALLA', 'DEFECTO'].forEach(function (header) {
    var index = headers.indexOf(header);
    if (index < 0) {
      // Fallback searches
      if (header === 'PIEZAS S') {
        for (var h = 0; h < headers.length; h++) {
          if (headers[h] === 'PIEZAS' || headers[h] === 'PIEZAS SCRAP' || headers[h] === 'CANTIDAD') { index = h; break; }
        }
      } else if (header === 'MODO DE FALLA') {
        for (var h = 0; h < headers.length; h++) {
          if (headers[h].indexOf('MODO') >= 0 || headers[h] === 'FALLA') { index = h; break; }
        }
      } else if (header === 'DEFECTO') {
        for (var h = 0; h < headers.length; h++) {
          if (headers[h].indexOf('DEFECTO') >= 0) { index = h; break; }
        }
      }
    }
    if (index < 0) throw new Error('Falta el encabezado ' + header + ' en S.SCRAP.');
    indexes[header] = index;
  });

  var fechaIndex = headers.indexOf('FECHA');
  var fechaCorteIndex = headers.indexOf('FECHA CORTE');
  var semanaIndex = headers.indexOf('SEMANA');
  var operacionIndex = headers.indexOf('OPERACION');
  if (operacionIndex < 0) operacionIndex = headers.indexOf('PROCESO');
  var deptoIndex = headers.indexOf('DEPARTAMENTO GENERADO');
  if (deptoIndex < 0) deptoIndex = headers.indexOf('DEPARTAMENTO');
  var parteIndex = headers.indexOf('NUMERO DE PARTE');
  if (parteIndex < 0) parteIndex = headers.indexOf('PARTE');
  var pesoIndex = headers.indexOf('PESO TONELADAS');
  if (pesoIndex < 0) pesoIndex = headers.indexOf('TONELADAS');

  var references = Object.create(null);
  var knownProcesses = uniqueSorted_(machineRows.map(function (row) { return row.process; }));
  machineRows.forEach(function (row) {
    var key = normalizeComparable_(row.machine);
    if (!references[key]) references[key] = [];
    references[key].push(row);
  });

  var groups = Object.create(null);
  var skipped = 0, unmapped = 0, sourceRows = 0;

  values.slice(1).forEach(function (row) {
    if (isEmptyRow_(row)) return;
    sourceRows++;
    var machine = cleanString_(row[indexes.MAQUINA]);
    if (!machine) { skipped++; return; }

    var period = null;
    if (fechaCorteIndex >= 0) {
      period = parsePeriod_(row[fechaCorteIndex], config, timeZone);
      if (!period) {
        var rawFc = cleanString_(row[fechaCorteIndex]).toLowerCase();
        var fcMatch = rawFc.match(/^([a-z]+)[-\s\/](\d{2,4})$/);
        if (fcMatch) {
          var mNames = { ene: 1, feb: 2, mar: 3, abr: 4, may: 5, jun: 6, jul: 7, ago: 8, sep: 9, oct: 10, nov: 11, dic: 12 };
          var mNum = mNames[fcMatch[1].slice(0, 3)];
          if (mNum) {
            var rawY = Number(fcMatch[2]);
            var yNum = rawY < 100 ? 2000 + rawY : rawY;
            period = { value: yNum + '-' + (mNum < 10 ? '0' + mNum : mNum), sort: yNum * 100 + mNum, year: yNum, month: mNum };
          }
        }
      }
    }
    if (!period && fechaIndex >= 0) {
      period = parsePeriod_(row[fechaIndex], config, timeZone);
    }
    if (!period) {
      period = { value: '2026-10', sort: 202610, year: 2026, month: 10 };
    }

    var pcs = numberOrNull_(row[indexes['PIEZAS S']], false);
    if (!isFiniteNumber_(pcs) || pcs <= 0) {
      // Still count rows if there's tonnage or piece is integer
      pcs = parseInt(row[indexes['PIEZAS S']], 10) || 0;
    }
    var tons = pesoIndex >= 0 ? numberOrNull_(row[pesoIndex], false) : 0;
    if (!isFiniteNumber_(tons) || tons < 0) tons = 0;
    var kg = tons * 1000;

    var modoFalla = cleanString_(row[indexes['MODO DE FALLA']]) || 'Sin modo de falla';
    var defecto = cleanString_(row[indexes.DEFECTO]) || 'Sin defecto';
    var depto = deptoIndex >= 0 ? cleanString_(row[deptoIndex]) : 'Sin departamento';
    if (!depto) depto = 'Sin departamento';
    var parte = parteIndex >= 0 ? cleanString_(row[parteIndex]) : 'Sin parte';
    if (!parte) parte = 'Sin parte';

    var rawSemana = semanaIndex >= 0 ? cleanString_(row[semanaIndex]) : '';
    var semana = 'Sin semana';
    if (rawSemana) {
      var semClean = rawSemana.trim().toUpperCase();
      if (/^W\d+$/.test(semClean)) {
        semana = semClean;
      } else if (/^\d+$/.test(semClean)) {
        semana = 'W' + semClean;
      } else if (semClean.indexOf('SEM') >= 0) {
        var numMatch = semClean.match(/\d+/);
        semana = numMatch ? 'W' + numMatch[0] : semClean;
      } else {
        semana = semClean;
      }
    }

    var matches = references[normalizeComparable_(machine)] || [];
    if (matches.length) machine = matches[0].machine;
    var sameMonth = matches.filter(function (item) { return item.period === period.value; });
    var processes = uniqueSorted_((sameMonth.length ? sameMonth : matches).map(function (item) { return item.process; }));
    var rawProcess = operacionIndex < 0 ? '' : cleanString_(row[operacionIndex]);
    if (rawProcess === '#N/D' || rawProcess === '#¡REF!' || rawProcess === '-' || rawProcess === 'N/A') rawProcess = '';
    var process = processes.length === 1 ? processes[0] : knownProcesses.filter(function (name) {
      return normalizeComparable_(name) === normalizeComparable_(rawProcess);
    })[0];
    if (!process && processes.length > 0) process = processes[0];
    if (!process) { process = 'Sin proceso homologado'; unmapped++; }

    var key = JSON.stringify([machine, process, period.value, modoFalla, defecto, depto, parte, semana]);
    if (!groups[key]) {
      groups[key] = {
        m: machine, proc: process, p: period.value, ps: period.sort,
        mf: modoFalla, def: defecto, dept: depto, part: parte, sem: semana,
        pcs: 0, tons: 0, kg: 0, count: 0
      };
    }
    groups[key].pcs += pcs;
    groups[key].tons += tons;
    groups[key].kg += kg;
    groups[key].count++;
  });

  var rows = Object.keys(groups).map(function (key) {
    var g = groups[key];
    g.tons = round_(g.tons, 5);
    g.kg = round_(g.kg, 2);
    return g;
  });

  return {
    rows: rows,
    status: {
      state: 'ready', source: 'S.SCRAP', sourceRows: sourceRows, skippedRows: skipped,
      unmappedRows: unmapped
    }
  };
}

/**
 * Resuelve y valida los encabezados de la hoja, sin depender del orden.
 *
 * @param {Array<*>} headers Encabezados de la primera fila.
 * @return {Object<string, number>}
 * @private
 */
function resolveColumnIndexes_(headers) {
  var normalizedIndexes = {};
  headers.forEach(function (header, index) {
    var normalized = normalizeHeader_(header);
    if (!normalized) {
      return;
    }
    if (Object.prototype.hasOwnProperty.call(normalizedIndexes, normalized)) {
      throw new Error('El encabezado "' + header + '" está duplicado en la hoja fuente.');
    }
    normalizedIndexes[normalized] = index;
  });

  var indexes = {};
  var missing = [];
  Object.keys(COLUMN_HEADERS_).forEach(function (field) {
    var expected = COLUMN_HEADERS_[field];
    var normalizedExpected = normalizeHeader_(expected);
    if (!Object.prototype.hasOwnProperty.call(normalizedIndexes, normalizedExpected)) {
      missing.push(expected);
    } else {
      indexes[field] = normalizedIndexes[normalizedExpected];
    }
  });

  if (missing.length) {
    throw new Error(
      'La estructura de "H. MAQ" cambió. Faltan los encabezados: ' + missing.join(', ') + '.'
    );
  }
  return indexes;
}

/**
 * Genera las opciones globales de filtros.
 *
 * @param {Array<Object>} rows Filas normalizadas.
 * @param {Object} requestedFilters Filtros solicitados.
 * @return {Object}
 * @private
 */
function buildFilterOptions_(rows, requestedFilters, cdemSummary) {
  var periodsByValue = {};
  (rows || []).forEach(function (row) {
    if (row.period) {
      periodsByValue[row.period] = {
        value: row.period,
        label: row.periodLabel || periodLabelFromKey_(row.period),
        sort: row.periodSort || periodSort_(row.period)
      };
    }
  });

  (cdemSummary || []).forEach(function (row) {
    if (row.p && !periodsByValue[row.p]) {
      periodsByValue[row.p] = {
        value: row.p,
        label: periodLabelFromKey_(row.p),
        sort: row.ps || periodSort_(row.p)
      };
    }
  });

  var periods = Object.keys(periodsByValue)
    .map(function (key) { return periodsByValue[key]; })
    .sort(function (a, b) { return a.sort - b.sort; })
    .map(function (period) {
      return { value: period.value, label: period.label };
    });

  return {
    processes: uniqueSorted_(rows.map(function (row) { return row.process; })),
    machines: uniqueSorted_(rows.map(function (row) { return row.machine; })),
    periods: periods
  };
}

/**
 * Normaliza la selección pedida y aplica periodos predeterminados.
 *
 * @param {Object} requestedFilters Filtros solicitados.
 * @param {Object} options Opciones válidas.
 * @param {Array<Object>} rows Filas normalizadas.
 * @return {Object}
 * @private
 */
function resolveSelection_(requestedFilters, options, rows) {
  var filters = requestedFilters && typeof requestedFilters === 'object'
    ? requestedFilters
    : {};
  var process = canonicalChoice_(filters.process, options.processes, 'proceso');
  var machine = canonicalChoice_(filters.machine, uniqueSorted_(rows.map(function (row) {
    return row.machine;
  })), 'máquina');

  if (!options.periods.length) {
    throw new Error('No hay periodos disponibles para construir la tendencia.');
  }

  var toPeriod = canonicalPeriod_(
    filters.toPeriod || filters.to,
    options.periods,
    options.periods[options.periods.length - 1].value,
    'periodo final'
  );
  var toYear = toPeriod.split('-')[0];
  var yearPeriods = options.periods.filter(function (period) { return (period.value || period).startsWith(toYear) && (period.value || period) <= toPeriod; });
  var defaultFrom = yearPeriods.length ? (yearPeriods[0].value || yearPeriods[0]) : shiftPeriod_(toPeriod, -11);
  var fromPeriod = canonicalPeriod_(
    filters.fromPeriod || filters.from,
    options.periods,
    defaultFrom,
    'periodo inicial'
  );

  var fromSort = periodSort_(fromPeriod);
  var toSort = periodSort_(toPeriod);
  if (fromSort > toSort) {
    var swapPeriod = fromPeriod;
    var swapSort = fromSort;
    fromPeriod = toPeriod;
    fromSort = toSort;
    toPeriod = swapPeriod;
    toSort = swapSort;
  }

  return {
    process: process,
    machine: machine,
    fromPeriod: fromPeriod,
    toPeriod: toPeriod,
    fromPeriodSort: fromSort,
    toPeriodSort: toSort
  };
}

/**
 * Construye la serie temporal agregada.
 *
 * @param {Array<Object>} rows Filas filtradas.
 * @param {string} fromPeriod Periodo inicial AAAA-MM.
 * @param {string} toPeriod Periodo final AAAA-MM.
 * @return {Array<Object>}
 * @private
 */
function buildTrend_(rows, fromPeriod, toPeriod) {
  var groups = groupBy_(rows, function (row) { return row.period; });
  var trend = periodRange_(fromPeriod, toPeriod)
    .map(function (period) {
      var periodRows = groups[period] || [];
      var aggregate = aggregateRows_(periodRows);
      return {
        period: period,
        label: periodLabelFromKey_(period),
        oee: aggregate.oee,
        availability: aggregate.availability,
        efficiency: aggregate.efficiency,
        quality: aggregate.quality,
        downtimeMinutes: aggregate.downtimeMinutes,
        production: aggregate.production,
        target: aggregate.target
      };
    });

  trend.forEach(function (point, index) {
    var window = trend.slice(Math.max(0, index - 2), index + 1);
    var validOee = window
      .map(function (item) { return item.oee; })
      .filter(isFiniteNumber_);
    // Se requieren tres observaciones válidas para no disfrazar huecos como
    // una media móvil completa.
    point.rollingOee3 = window.length === 3 && validOee.length === 3
      ? round_(validOee.reduce(function (sum, value) { return sum + value; }, 0) / 3, 6)
      : null;
  });
  return trend;
}

/**
 * Agrega métricas con ponderadores coherentes con cada indicador.
 *
 * @param {Array<Object>} rows Filas a agregar.
 * @return {Object}
 * @private
 */
function aggregateRows_(rows) {
  var activeRows = rows.filter(hasOperationalBase_);
  if (!activeRows.length) {
    return {
      oee: null,
      availability: null,
      efficiency: null,
      quality: null,
      downtimeMinutes: null,
      production: null,
      target: null,
      scrap: null,
      deltaOee: null
    };
  }

  // Los denominadores y numeradores deben cubrir las mismas filas. Un dato
  // ausente no equivale a cero producción, cero paro o cero scrap.
  var sumScheduled = sumComplete_(activeRows.map(function (row) { return row.scheduledMinutes; }));
  var sumDowntime = sumComplete_(activeRows.map(rowDowntimeNullable_));
  var sumTarget = sumComplete_(activeRows.map(function (row) { return row.target; }));
  var sumProductionTotal = sumComplete_(activeRows.map(function (row) { return row.productionTotal; }));
  var sumScrap = sumComplete_(activeRows.map(function (row) { return row.totalScrap; }));

  var availability = sumScheduled > 0 && isFiniteNumber_(sumDowntime)
    ? (sumScheduled - sumDowntime) / sumScheduled : null;
  var efficiency = sumTarget > 0 && isFiniteNumber_(sumProductionTotal)
    ? sumProductionTotal / sumTarget : null;
  var quality = sumProductionTotal > 0 && isFiniteNumber_(sumScrap)
    ? (sumProductionTotal - sumScrap) / sumProductionTotal : null;

  var oee = isFiniteNumber_(availability) && isFiniteNumber_(efficiency) && isFiniteNumber_(quality)
    ? availability * efficiency * quality
    : null;

  return {
    oee: nullableRound_(oee, 6),
    availability: nullableRound_(availability, 6),
    efficiency: nullableRound_(efficiency, 6),
    quality: nullableRound_(quality, 6),
    downtimeMinutes: nullableRound_(sumDowntime, 2),
    production: nullableRound_(sumProductionTotal, 2),
    target: nullableRound_(sumTarget, 2),
    scrap: nullableRound_(sumScrap, 2),
    deltaOee: null
  };
}

/**
 * Construye el Pareto de causas de paro.
 *
 * @param {Array<Object>} rows Filas filtradas.
 * @return {Array<Object>}
 * @private
 */
function buildAvailabilityView_(summary, status, requested) {
  status = status || { state: 'error', source: 'C.DEM', message: 'Los datos de C.DEM no están disponibles. Actualiza la aplicación.' };
  var options = {
    processes: uniqueSorted_(summary.map(function (row) { return row.proc; })),
    machines: [],
    periods: uniqueSorted_(summary.map(function (row) { return row.p; })).map(function (period) { return { value: period, label: periodLabelFromKey_(period) }; })
  };
  requested = requested || {};
  if (!options.periods.length) return { status: status, filters: options, selection: {}, rows: [], totalMinutes: 0, matchedRows: 0 };
  var latestPeriod = options.periods[options.periods.length - 1].value;
  var reqWithDefaults = {
    process: requested.process,
    machine: requested.machine,
    generalCause: requested.generalCause,
    fromPeriod: requested.fromPeriod || latestPeriod,
    toPeriod: requested.toPeriod || latestPeriod
  };
  var normalizedRows = summary.map(function (row) { return { machine: row.m }; });
  var selection = resolveSelection_(reqWithDefaults, options, normalizedRows);
  options.machines = uniqueSorted_(summary.filter(function (row) { return !selection.process || row.proc === selection.process; }).map(function (row) { return row.m; }));
  if (selection.machine && options.machines.indexOf(selection.machine) < 0) selection.machine = null;
  
  var filteredForCauses = summary.filter(function (row) {
    return (!selection.process || row.proc === selection.process) && (!selection.machine || row.m === selection.machine) &&
      row.ps >= selection.fromPeriodSort && row.ps <= selection.toPeriodSort;
  });
  
  options.generalCauses = uniqueSorted_(filteredForCauses.map(function(row) { return row.cg; }));
  selection.generalCause = requested.generalCause && options.generalCauses.indexOf(requested.generalCause) >= 0 ? requested.generalCause : null;
  
  var filtered = summary.filter(function (row) {
    return (!selection.process || row.proc === selection.process) && (!selection.machine || row.m === selection.machine) && (!selection.generalCause || row.cg === selection.generalCause) &&
      row.ps >= selection.fromPeriodSort && row.ps <= selection.toPeriodSort;
  });
  var paretoResult = buildCdemPareto_(summary, selection);
  var rows = paretoResult.rows || paretoResult;
  return {
    status: status, filters: options, selection: selection, rows: rows,
    weeks: paretoResult.weeks || [],
    totalMinutes: round_(rows.reduce(function (sum, row) { return sum + row.minutes; }, 0), 2),
    matchedRows: filtered.reduce(function (sum, row) { return sum + (row.count || 0); }, 0)
  };
}

function buildCdemPareto_(cdemSummary, selection) {
  var filtered = cdemSummary.filter(function(row) {
    var matchP = !selection.process || row.proc === selection.process;
    var matchM = !selection.machine || row.m === selection.machine;
    var matchCG = !selection.generalCause || row.cg === selection.generalCause;
    var matchTime = row.ps >= selection.fromPeriodSort && row.ps <= selection.toPeriodSort;
    return matchP && matchM && matchCG && matchTime;
  });

  var grouped = Object.create(null);
  var weeksSet = Object.create(null);
  filtered.forEach(function(r) {
    var key = r.m; // Level 1
    if (selection.machine && !selection.generalCause) key = r.cg; // Level 2
    if (selection.machine && selection.generalCause) key = r.ce; // Level 3

    if (!grouped[key]) {
      grouped[key] = {
        total: 0,
        byWeek: Object.create(null),
        byPart: Object.create(null),
        byDept: Object.create(null),
        byOt: Object.create(null)
      };
    }
    grouped[key].total += r.min;
    var sem = r.sem || 'Sin semana';
    weeksSet[sem] = true;
    if (!grouped[key].byWeek[sem]) grouped[key].byWeek[sem] = 0;
    grouped[key].byWeek[sem] += r.min;

    var partKey = r.part || 'Sin número de parte';
    if (!grouped[key].byPart[partKey]) grouped[key].byPart[partKey] = 0;
    grouped[key].byPart[partKey] += r.min;

    var deptKey = r.dept || 'Sin departamento';
    if (!grouped[key].byDept[deptKey]) grouped[key].byDept[deptKey] = 0;
    grouped[key].byDept[deptKey] += r.min;

    var rawOt = r.ot || 'Sin orden de trabajo';
    if (!rawOt || rawOt === 'Sin orden de trabajo' || rawOt === '-' || rawOt === '0') {
      var sinOt = 'Sin orden de trabajo';
      if (!grouped[key].byOt[sinOt]) grouped[key].byOt[sinOt] = 0;
      grouped[key].byOt[sinOt] += r.min;
    } else {
      var otTokens = rawOt.split(/[,;\/]+/).map(function(s) { return s.trim(); }).filter(Boolean);
      if (otTokens.length > 1) {
        var splitMin = r.min / otTokens.length;
        otTokens.forEach(function(tok) {
          if (!grouped[key].byOt[tok]) grouped[key].byOt[tok] = 0;
          grouped[key].byOt[tok] += splitMin;
        });
      } else {
        var tok = otTokens[0] || rawOt;
        if (!grouped[key].byOt[tok]) grouped[key].byOt[tok] = 0;
        grouped[key].byOt[tok] += r.min;
      }
    }
  });

  var result = [];
  for (var k in grouped) {
    if (grouped[k].total > 0) {
      var cleanByWeek = {};
      for (var w in grouped[k].byWeek) {
        cleanByWeek[w] = round_(grouped[k].byWeek[w], 2);
      }

      var cleanParts = Object.keys(grouped[k].byPart).map(function(p) {
        return { part: p, minutes: round_(grouped[k].byPart[p], 2) };
      }).sort(function(a, b) {
        return b.minutes - a.minutes || a.part.localeCompare(b.part, 'es');
      });

      var cleanDepts = Object.keys(grouped[k].byDept).map(function(d) {
        return { dept: d, minutes: round_(grouped[k].byDept[d], 2) };
      }).sort(function(a, b) {
        return b.minutes - a.minutes || a.dept.localeCompare(b.dept, 'es');
      });

      var cleanOts = Object.keys(grouped[k].byOt).map(function(o) {
        return { ot: o, minutes: round_(grouped[k].byOt[o], 2) };
      }).sort(function(a, b) {
        return b.minutes - a.minutes || a.ot.localeCompare(b.ot, 'es');
      });

      result.push({
        key: k,
        label: k,
        minutes: round_(grouped[k].total, 2),
        byWeek: cleanByWeek,
        parts: cleanParts,
        depts: cleanDepts,
        ots: cleanOts
      });
    }
  }

  result.sort(function(a, b) {
    return b.minutes - a.minutes || a.label.localeCompare(b.label, 'es');
  });
  var total = result.reduce(function (sum, row) { return sum + row.minutes; }, 0);
  var running = 0;
  result.forEach(function (row, index) {
    running += row.minutes;
    row.share = total ? row.minutes / total : 0;
    row.cumulative = total ? (index === result.length - 1 ? 1 : running / total) : 0;
  });

  var sortedWeeks = Object.keys(weeksSet).sort(function(a, b) {
    var numA = parseInt((a.match(/\d+/) || [999])[0], 10);
    var numB = parseInt((b.match(/\d+/) || [999])[0], 10);
    if (numA !== numB) return numA - numB;
    return a.localeCompare(b);
  });
  result.weeks = sortedWeeks;
  return result;
}

function buildScrapView_(summary, status, requested) {
  status = status || { state: 'error', source: 'S.SCRAP', message: 'Los datos de S.SCRAP no están disponibles. Actualiza la aplicación.' };
  var options = {
    processes: uniqueSorted_(summary.map(function (row) { return row.proc; })),
    machines: [],
    periods: uniqueSorted_(summary.map(function (row) { return row.p; })).map(function (period) { return { value: period, label: periodLabelFromKey_(period) }; })
  };
  requested = requested || {};
  if (!options.periods.length) return { status: status, filters: options, selection: {}, rows: [], totalPieces: 0, totalKg: 0, matchedRows: 0 };
  var latestPeriod = options.periods[options.periods.length - 1].value;
  var reqWithDefaults = {
    process: requested.process,
    machine: requested.machine,
    modoFalla: requested.modoFalla,
    fromPeriod: requested.fromPeriod || latestPeriod,
    toPeriod: requested.toPeriod || latestPeriod
  };
  var normalizedRows = summary.map(function (row) { return { machine: row.m }; });
  var selection = resolveSelection_(reqWithDefaults, options, normalizedRows);
  options.machines = uniqueSorted_(summary.filter(function (row) { return !selection.process || row.proc === selection.process; }).map(function (row) { return row.m; }));
  if (selection.machine && options.machines.indexOf(selection.machine) < 0) selection.machine = null;

  var filteredForModos = summary.filter(function (row) {
    return (!selection.process || row.proc === selection.process) && (!selection.machine || row.m === selection.machine) &&
      row.ps >= selection.fromPeriodSort && row.ps <= selection.toPeriodSort;
  });

  options.modosFalla = uniqueSorted_(filteredForModos.map(function(row) { return row.mf; }));
  selection.modoFalla = requested.modoFalla && options.modosFalla.indexOf(requested.modoFalla) >= 0 ? requested.modoFalla : null;

  var filtered = summary.filter(function (row) {
    return (!selection.process || row.proc === selection.process) && (!selection.machine || row.m === selection.machine) && (!selection.modoFalla || row.mf === selection.modoFalla) &&
      row.ps >= selection.fromPeriodSort && row.ps <= selection.toPeriodSort;
  });
  var paretoResult = buildScrapPareto_(summary, selection);
  var rows = paretoResult.rows || paretoResult;
  return {
    status: status, filters: options, selection: selection, rows: rows,
    weeks: paretoResult.weeks || [],
    totalPieces: paretoResult.totalPieces || rows.reduce(function (sum, row) { return sum + row.pieces; }, 0),
    totalKg: paretoResult.totalKg || round_(rows.reduce(function (sum, row) { return sum + row.kg; }, 0), 2),
    matchedRows: filtered.reduce(function (sum, row) { return sum + (row.count || 0); }, 0)
  };
}

function buildScrapPareto_(scrapSummary, selection) {
  var filtered = scrapSummary.filter(function(row) {
    var matchP = !selection.process || row.proc === selection.process;
    var matchM = !selection.machine || row.m === selection.machine;
    var matchMF = !selection.modoFalla || row.mf === selection.modoFalla;
    var matchTime = row.ps >= selection.fromPeriodSort && row.ps <= selection.toPeriodSort;
    return matchP && matchM && matchMF && matchTime;
  });

  var grouped = Object.create(null);
  var weeksSet = Object.create(null);
  filtered.forEach(function(r) {
    var key = r.m; // Nivel 1: Máquinas con mayor scrap
    if (selection.machine && !selection.modoFalla) key = r.mf; // Nivel 2: Modo de falla
    if (selection.machine && selection.modoFalla) key = r.def; // Nivel 3: Defecto específico

    if (!grouped[key]) {
      grouped[key] = { pieces: 0, kg: 0, tons: 0, byWeek: Object.create(null), byWeekKg: Object.create(null) };
    }
    grouped[key].pieces += r.pcs;
    grouped[key].kg += r.kg;
    grouped[key].tons += r.tons;
    var sem = r.sem || 'Sin semana';
    weeksSet[sem] = true;
    if (!grouped[key].byWeek[sem]) grouped[key].byWeek[sem] = 0;
    grouped[key].byWeek[sem] += r.pcs;
    if (!grouped[key].byWeekKg[sem]) grouped[key].byWeekKg[sem] = 0;
    grouped[key].byWeekKg[sem] += r.kg;
  });

  var result = [];
  for (var k in grouped) {
    if (grouped[k].pieces > 0 || grouped[k].kg > 0) {
      var cleanByWeek = {};
      var cleanByWeekKg = {};
      for (var w in grouped[k].byWeek) {
        cleanByWeek[w] = grouped[k].byWeek[w];
        cleanByWeekKg[w] = round_(grouped[k].byWeekKg[w], 2);
      }
      result.push({
        key: k,
        label: k,
        pieces: grouped[k].pieces,
        kg: round_(grouped[k].kg, 2),
        tons: round_(grouped[k].tons, 4),
        byWeek: cleanByWeek,
        byWeekKg: cleanByWeekKg
      });
    }
  }

  result.sort(function(a, b) {
    return b.pieces - a.pieces || b.kg - a.kg || a.label.localeCompare(b.label, 'es');
  });
  var totalPieces = result.reduce(function (sum, row) { return sum + row.pieces; }, 0);
  var totalKg = round_(result.reduce(function (sum, row) { return sum + row.kg; }, 0), 2);
  var running = 0;
  result.forEach(function (row, index) {
    running += row.pieces;
    row.share = totalPieces ? row.pieces / totalPieces : 0;
    row.cumulative = totalPieces ? (index === result.length - 1 ? 1 : running / totalPieces) : 0;
  });

  var sortedWeeks = Object.keys(weeksSet).sort(function(a, b) {
    var numA = parseInt((a.match(/\d+/) || [999])[0], 10);
    var numB = parseInt((b.match(/\d+/) || [999])[0], 10);
    if (numA !== numB) return numA - numB;
    return a.localeCompare(b);
  });
  result.weeks = sortedWeeks;
  result.totalPieces = totalPieces;
  result.totalKg = totalKg;
  return result;
}

function buildDowntimePareto_(rows) {
  var activeRows = rows.filter(hasOperationalBase_);
  var totalScheduledMinutes = sumNullable_(activeRows.map(function (row) {
    return row.scheduledMinutes;
  }));
  var result = DOWNTIME_FIELDS_.map(function (definition) {
    var minutes = activeRows.reduce(function (sum, row) {
      return sum + safeNumber_(row[definition.field]);
    }, 0);
    return {
      key: definition.key,
      label: definition.label,
      minutes: round_(minutes, 2),
      share: 0,
      normalizedPer1000: isFiniteNumber_(totalScheduledMinutes) && totalScheduledMinutes > 0
        ? round_(minutes * 1000 / totalScheduledMinutes, 4)
        : null
    };
  });
  var total = result.reduce(function (sum, item) { return sum + item.minutes; }, 0);
  result.forEach(function (item) {
    item.share = total > 0 ? round_(item.minutes / total, 6) : 0;
  });
  return result.sort(function (a, b) {
    return b.minutes - a.minutes || a.label.localeCompare(b.label, 'es');
  });
}

/**
 * Compara máquinas en el periodo final solicitado. Los faltantes permanecen
 * visibles con KPI null en lugar de retroceder silenciosamente a otro mes.
 * El filtro de máquina se ignora intencionalmente para conservar los pares.
 *
 * @param {Array<Object>} allRows Todas las filas normalizadas.
 * @param {Object} selection Selección resuelta.
 * @return {Array<Object>}
 * @private
 */
function buildRanking_(allRows, selection) {
  var peerRows = allRows.filter(function (row) {
    return !selection.process || row.process === selection.process;
  });
  if (!peerRows.length) {
    return [];
  }
  var machines = uniqueSorted_(peerRows.map(function (row) { return row.machine; }));
  var previousPeriod = shiftPeriod_(selection.toPeriod, -1);
  var previousSort = periodSort_(previousPeriod);

  return machines.map(function (machine) {
    var machineRows = peerRows.filter(function (row) { return row.machine === machine; });
    var currentRows = machineRows.filter(function (row) {
      return row.periodSort === selection.toPeriodSort;
    });
    var current = aggregateRows_(currentRows);
    var previous = aggregateRows_(machineRows.filter(function (row) {
      return row.periodSort === previousSort;
    }));
    var delta = isFiniteNumber_(current.oee) && isFiniteNumber_(previous.oee)
      ? round_(current.oee - previous.oee, 6)
      : null;
    var referenceRow = currentRows[0] || machineRows[0];

    return {
      machine: machine,
      process: referenceRow.process,
      oee: current.oee,
      availability: current.availability,
      efficiency: current.efficiency,
      quality: current.quality,
      deltaOee: delta,
      period: selection.toPeriod,
      periodLabel: periodLabelFromKey_(selection.toPeriod)
    };
  }).sort(function (a, b) {
    if (a.oee === null && b.oee === null) {
      return a.machine.localeCompare(b.machine);
    }
    if (a.oee === null) {
      return 1;
    }
    if (b.oee === null) {
      return -1;
    }
    return b.oee - a.oee || a.machine.localeCompare(b.machine);
  });
}

/**
 * Genera observaciones breves, estrictamente derivadas de los datos.
 *
 * @param {Array<Object>} trend Serie temporal.
 * @param {Object} kpis KPI agregados.
 * @param {Array<Object>} downtime Pareto.
 * @param {Array<Object>} ranking Ranking de máquinas.
 * @param {number} rowCount Número de filas filtradas.
 * @return {Array<string>}
 * @private
 */
function buildInsights_(trend, kpis, downtime, ranking, rowCount, last6Aggregate) {
  if (!rowCount) {
    return ['No hay datos para la combinación de filtros seleccionada.'];
  }

  var insights = [];
  if (trend.length > 1 && isFiniteNumber_(kpis.deltaOee)) {
    var direction = kpis.deltaOee > 0 ? 'mejoró' : kpis.deltaOee < 0 ? 'disminuyó' : 'se mantuvo';
    insights.push(
      'El OEE ' + direction + ' ' + Math.abs(kpis.deltaOee * 100).toFixed(3) +
      ' puntos porcentuales frente al periodo anterior.'
    );
  }

  var oeeTarget = runtimeConfig_().oeeTarget;
  var validMonths = 0;
  var outOfTarget = 0;
  for (var i = 0; i < trend.length; i++) {
    if (isFiniteNumber_(trend[i].oee)) {
      validMonths++;
      if (trend[i].oee < oeeTarget) outOfTarget++;
    }
  }
  
  if (validMonths > 0) {
    if (outOfTarget > 0) {
      insights.push('El OEE estuvo fuera de meta en ' + outOfTarget + ' de los ' + validMonths + ' meses analizados.');
    } else {
      insights.push('El OEE se mantuvo en la meta durante los ' + validMonths + ' meses analizados.');
    }
  }

  if (downtime.length && downtime[0].minutes > 0) {
    insights.push(
      downtime[0].label + ' concentra la mayor causa de paro: ' +
      formatInteger_(downtime[0].minutes) + ' min (' +
      (downtime[0].share * 100).toFixed(3) + '% del Pareto).'
    );
  }

  var comparableRanking = ranking.filter(function (item) { return isFiniteNumber_(item.oee); });
  if (comparableRanking.length) {
    insights.push(
      comparableRanking[0].machine + ' tiene el mayor OEE del último periodo: ' +
      (comparableRanking[0].oee * 100).toFixed(3) + '%.'
    );
  }
  
  if (comparableRanking.length > 1) {
    var lowest = comparableRanking[comparableRanking.length - 1];
    insights.push(
      lowest.machine + ' tiene el menor OEE del último periodo: ' +
      (lowest.oee * 100).toFixed(3) + '%.'
    );
  }
  
  return insights.slice(0, 5);
}

/**
 * Configuración efectiva a partir de propiedades del script.
 *
 * @return {Object}
 * @private
 */
function runtimeConfig_() {
  var properties = PropertiesService.getScriptProperties();
  var spreadsheetId = cleanString_(properties.getProperty('HMAQ_SPREADSHEET_ID')) ||
    APP_CONFIG_.spreadsheetId;
  var sheetName = cleanString_(properties.getProperty('HMAQ_SHEET_NAME')) ||
    APP_CONFIG_.sheetName;
  var mode = cleanString_(properties.getProperty('HMAQ_DATE_MODE') || APP_CONFIG_.dateMode)
    .toUpperCase();
  if (mode !== 'DAY_MARKER' && mode !== 'CALENDAR_YEAR') {
    throw new Error(
      'HMAQ_DATE_MODE debe ser DAY_MARKER o CALENDAR_YEAR; se recibió "' + mode + '".'
    );
  }

  var dayYearMap = Object.assign({}, APP_CONFIG_.dateDayYearMap);
  var rawMap = properties.getProperty('HMAQ_DATE_DAY_YEAR_MAP');
  if (rawMap) {
    try {
      var parsed = JSON.parse(rawMap);
      validateDayYearMap_(parsed);
      dayYearMap = parsed;
    } catch (error) {
      throw new Error(
        'HMAQ_DATE_DAY_YEAR_MAP debe ser JSON válido, por ejemplo ' +
        '{"25":2025,"26":2026}. Detalle: ' + publicErrorMessage_(error)
      );
    }
  }
  validateDayYearMap_(dayYearMap);

  var cacheSeconds = Number(properties.getProperty('HMAQ_CACHE_SECONDS') || APP_CONFIG_.cacheSeconds);
  if (!isFinite(cacheSeconds)) {
    cacheSeconds = APP_CONFIG_.cacheSeconds;
  }
  cacheSeconds = Math.max(30, Math.min(21600, Math.floor(cacheSeconds)));

  var oeeTarget = numberOrNull_(
    properties.getProperty('HMAQ_OEE_TARGET') || APP_CONFIG_.oeeTarget,
    true
  );
  if (!isFiniteNumber_(oeeTarget) || oeeTarget <= 0 || oeeTarget > 2) {
    throw new Error('HMAQ_OEE_TARGET debe ser un porcentaje positivo de hasta 200%.');
  }

  return {
    spreadsheetId: spreadsheetId,
    sheetName: sheetName,
    timeZone: APP_CONFIG_.defaultTimeZone,
    cacheSeconds: cacheSeconds,
    oeeTarget: oeeTarget,
    dateMode: mode,
    dateDayYearMap: dayYearMap
  };
}

/**
 * Interpreta MES y aplica la convención de día marcador cuando corresponde.
 *
 * @param {*} value Valor de celda.
 * @param {Object} config Configuración efectiva.
 * @param {string} timeZone Zona horaria de la hoja.
 * @return {?Object}
 * @private
 */
function parsePeriod_(value, config, timeZone) {
  var year;
  var month;
  var day;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    var parts = Utilities.formatDate(value, timeZone, 'yyyy-M-d').split('-');
    year = Number(parts[0]);
    month = Number(parts[1]);
    day = Number(parts[2]);
  } else if (typeof value === 'number' && isFinite(value)) {
    // Compatibilidad con seriales de Google Sheets/Excel (época 1899-12-30).
    var serialDate = new Date(Date.UTC(1899, 11, 30) + Math.round(value * 86400000));
    year = serialDate.getUTCFullYear();
    month = serialDate.getUTCMonth() + 1;
    day = serialDate.getUTCDate();
  } else {
    var parsedText = parsePeriodText_(cleanString_(value));
    if (!parsedText) {
      return null;
    }
    year = parsedText.year;
    month = parsedText.month;
    day = parsedText.day;
  }

  if (config.dateMode === 'DAY_MARKER' && day &&
      Object.prototype.hasOwnProperty.call(config.dateDayYearMap, String(day))) {
    year = Number(config.dateDayYearMap[String(day)]);
  }
  if (!isFinite(year) || !isFinite(month) || month < 1 || month > 12) {
    return null;
  }

  var valueKey = String(Math.floor(year)) + '-' + pad2_(month);
  return {
    value: valueKey,
    label: MONTH_LABELS_[month - 1] + ' ' + Math.floor(year),
    sort: Math.floor(year) * 100 + month
  };
}

/**
 * Acepta YYYY-MM, DD/MM/YYYY, MM/YYYY y nombres de mes en español/inglés.
 * En formatos sin día no se puede aplicar DAY_MARKER y se conserva el año.
 *
 * @param {string} text Texto de fecha.
 * @return {?Object}
 * @private
 */
function parsePeriodText_(text) {
  if (!text) {
    return null;
  }
  var normalized = text.toLowerCase().replace(/\./g, '').trim();
  var match = normalized.match(/^(\d{4})[-\/]([01]?\d)(?:[-\/]([0-3]?\d))?$/);
  if (match) {
    return { year: Number(match[1]), month: Number(match[2]), day: Number(match[3] || 0) };
  }
  match = normalized.match(/^([0-3]?\d)[-\/]([01]?\d)[-\/](\d{4})$/);
  if (match) {
    return { year: Number(match[3]), month: Number(match[2]), day: Number(match[1]) };
  }
  match = normalized.match(/^([01]?\d)[-\/](\d{4})$/);
  if (match) {
    return { year: Number(match[2]), month: Number(match[1]), day: 0 };
  }

  var monthNames = {
    ene: 1, enero: 1, jan: 1, january: 1,
    feb: 2, febrero: 2, february: 2,
    mar: 3, marzo: 3, march: 3,
    abr: 4, abril: 4, apr: 4, april: 4,
    may: 5, mayo: 5,
    jun: 6, junio: 6, june: 6,
    jul: 7, julio: 7, july: 7,
    ago: 8, agosto: 8, aug: 8, august: 8,
    sep: 9, sept: 9, septiembre: 9, september: 9,
    oct: 10, octubre: 10, october: 10,
    nov: 11, noviembre: 11, november: 11,
    dic: 12, diciembre: 12, dec: 12, december: 12
  };
  match = normalized.match(/^([a-záéíóúñ]+)[-\s\/](\d{4})$/i);
  if (match) {
    var monthKey = removeDiacritics_(match[1].toLowerCase());
    if (monthNames[monthKey]) {
      return { year: Number(match[2]), month: monthNames[monthKey], day: 0 };
    }
  }
  return null;
}

/**
 * Convierte números, guiones y porcentajes de forma tolerante.
 * Los guiones, blancos, N/A y errores comunes se convierten a null.
 *
 * @param {*} value Valor de celda.
 * @param {boolean} percentage Indica si el campo es porcentaje.
 * @return {?number}
 * @private
 */
function numberOrNull_(value, percentage) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (typeof value === 'number') {
    if (!isFinite(value)) {
      return null;
    }
    // getValues() devuelve porcentajes de Sheets como razones decimales:
    // 160% llega como 1.6. No se aplica una heurística por magnitud porque
    // recortaría silenciosamente indicadores válidos por encima de 150%.
    return value;
  }

  var text = cleanString_(value);
  if (!text || /^(?:-|—|–|n\/?a|na|n\.d\.|nd|s\/?d|#n\/?a)$/i.test(text)) {
    return null;
  }
  var hasPercent = text.indexOf('%') >= 0;
  text = text.replace(/[%\s\u00A0]/g, '');
  if (text.indexOf(',') >= 0 && text.indexOf('.') >= 0) {
    if (text.lastIndexOf(',') > text.lastIndexOf('.')) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  } else if (text.indexOf(',') >= 0) {
    if (hasPercent || /,\d{1,2}$/.test(text)) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  }
  // Rechaza errores de fórmula y texto no numérico; no los convierte en 0.
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) {
    return null;
  }
  var numeric = Number(text);
  if (!isFinite(numeric)) {
    return null;
  }
  if (percentage && (hasPercent || Math.abs(numeric) > 1.5)) {
    numeric /= 100;
  }
  return numeric;
}

/**
 * Serializa el conjunto en formato compacto para respetar el límite del caché.
 *
 * @param {GoogleAppsScript.Cache.Cache} cache Caché de script.
 * @param {string} key Clave.
 * @param {Object} dataset Conjunto normalizado.
 * @param {number} seconds TTL.
 * @private
 */
function putDatasetInCache_(cache, key, dataset, seconds) {
  try {
    var compact = {
      m: {
        spreadsheetId: dataset.spreadsheetId,
        sheetName: dataset.sheetName,
        oeeTarget: dataset.oeeTarget,
        loadedAt: dataset.loadedAt,
        sourceRows: dataset.sourceRows,
        skippedRows: dataset.skippedRows,
        warnings: dataset.warnings,
        dateConvention: dataset.dateConvention
      },
      r: dataset.rows.map(compactRow_),
      c: (dataset.cdemSummary || []).map(function (row) { return [row.m, row.proc, row.p, row.ps, row.cg, row.ce, row.min, row.count, row.dept, row.part || 'Sin número de parte', row.sem || 'Sin semana', row.ot || 'Sin orden de trabajo']; }),
      cs: dataset.cdemStatus,
      ot: dataset.otMap,
      dev: dataset.deviations || [],
      devs: dataset.deviationsStatus || {}
    };
    var payload = JSON.stringify(compact);
    var byteLength = Utilities.newBlob(payload).getBytes().length;
    if (byteLength > 95000) {
      payload = 'gzip:' + Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(payload, 'application/json')).getBytes());
      byteLength = payload.length;
    }
    if (byteLength <= 95000) {
      cache.put(key, payload, seconds);
      return;
    }
    // Cada clave del caché admite ~100 KB: se reparte en fragmentos para que
    // el conjunto completo se guarde y las cargas siguientes no relean las hojas.
    var chunkSize = 90000;
    var total = Math.ceil(payload.length / chunkSize);
    if (total > 40) {
      console.warn('El conjunto normalizado es demasiado grande para el caché.');
      return;
    }
    var entries = {};
    for (var i = 0; i < total; i += 1) {
      entries[key + ':' + i] = payload.slice(i * chunkSize, (i + 1) * chunkSize);
    }
    cache.putAll(entries, seconds);
    cache.put(key, 'chunks:' + total, seconds);
  } catch (error) {
    // El caché es una optimización: una falla no debe impedir cargar el tablero.
    console.warn('No fue posible guardar el caché: ' + publicErrorMessage_(error));
  }
}

/**
 * Lee el conjunto del caché, ya sea en una sola clave o en fragmentos.
 * Devuelve null si falta algún fragmento o el contenido es inválido.
 *
 * @param {GoogleAppsScript.Cache.Cache} cache Caché de script.
 * @param {string} key Clave.
 * @return {?Object}
 * @private
 */
function getDatasetFromCache_(cache, key) {
  var stored = cache.get(key);
  if (!stored) {
    return null;
  }
  if (stored.indexOf('chunks:') !== 0) {
    return restoreCachedDataset_(stored);
  }
  var total = Number(stored.slice(7));
  if (!isFinite(total) || total < 1) {
    return null;
  }
  var keys = [];
  for (var i = 0; i < total; i += 1) {
    keys.push(key + ':' + i);
  }
  var parts = cache.getAll(keys);
  var payload = '';
  for (var j = 0; j < keys.length; j += 1) {
    if (!parts[keys[j]]) {
      return null;
    }
    payload += parts[keys[j]];
  }
  return restoreCachedDataset_(payload);
}

/**
 * Restaura un conjunto del caché. Devuelve null si está corrupto.
 *
 * @param {string} payload JSON compacto.
 * @return {?Object}
 * @private
 */
function restoreCachedDataset_(payload) {
  try {
    if (payload.indexOf('gzip:') === 0) payload = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(payload.slice(5)), 'application/gzip')).getDataAsString('UTF-8');
    var compact = JSON.parse(payload);
    if (!compact || !compact.m || !Array.isArray(compact.r) || !Array.isArray(compact.c) || !compact.cs) {
      return null;
    }
    return {
      spreadsheetId: compact.m.spreadsheetId,
      sheetName: compact.m.sheetName,
      oeeTarget: isFiniteNumber_(compact.m.oeeTarget) ? compact.m.oeeTarget : APP_CONFIG_.oeeTarget,
      loadedAt: compact.m.loadedAt,
      sourceRows: compact.m.sourceRows,
      skippedRows: compact.m.skippedRows,
      warnings: compact.m.warnings || [],
      dateConvention: compact.m.dateConvention,
      rows: compact.r.map(expandRow_),
      cdemSummary: compact.c.map(function (row) { return { m: row[0], proc: row[1], p: row[2], ps: row[3], cg: row[4], ce: row[5], min: row[6], count: row[7], dept: row[8] || 'Sin departamento', part: row[9] || 'Sin número de parte', sem: row[10] || 'Sin semana', ot: row[11] || 'Sin orden de trabajo' }; }),
      cdemStatus: compact.cs,
      otMap: compact.ot || {},
      deviations: compact.dev || [],
      deviationsStatus: compact.devs || {}
    };
  } catch (error) {
    console.warn('Caché inválido; se volverá a leer la hoja.');
    return null;
  }
}

/** @param {Object} row @return {Array<*>} @private */
function compactRow_(row) {
  return [
    row.machine, row.process, row.period, row.periodLabel, row.periodSort,
    row.availableMinutes, row.scheduledMinutes, row.occupancy, row.rate,
    row.target, row.productionTotal, row.maintenance, row.adjustment, row.engineering,
    row.humanResources, row.qualityDowntime, row.logistics, row.warehouse,
    row.productionDowntime, row.totalDowntime, row.availability, row.efficiency,
    row.quality, row.oee, row.totalScrap
  ];
}

/** @param {Array<*>} row @return {Object} @private */
function expandRow_(row) {
  return {
    machine: row[0], process: row[1], period: row[2], periodLabel: row[3],
    periodSort: row[4], availableMinutes: row[5], scheduledMinutes: row[6],
    occupancy: row[7], rate: row[8], target: row[9], productionTotal: row[10],
    maintenance: row[11], adjustment: row[12], engineering: row[13],
    humanResources: row[14], qualityDowntime: row[15], logistics: row[16],
    warehouse: row[17], productionDowntime: row[18], totalDowntime: row[19],
    availability: row[20], efficiency: row[21], quality: row[22], oee: row[23],
    totalScrap: row[24]
  };
}

/** @param {Object} config @return {string} @private */
function datasetCacheKey_(config) {
  var signature = [
    APP_CONFIG_.cacheVersion,
    config.spreadsheetId,
    config.sheetName,
    config.dateMode,
    JSON.stringify(config.dateDayYearMap),
    config.oeeTarget
  ].join('|');
  return 'dashboard-dataset-' + hashString_(signature);
}

/** @param {Object} map @private */
function validateDayYearMap_(map) {
  if (!map || typeof map !== 'object' || Array.isArray(map) || !Object.keys(map).length) {
    throw new Error('El mapa de días y años no puede estar vacío.');
  }
  Object.keys(map).forEach(function (dayKey) {
    var day = Number(dayKey);
    var year = Number(map[dayKey]);
    if (!Number.isInteger(day) || day < 1 || day > 31 ||
        !Number.isInteger(year) || year < 2000 || year > 2100) {
      throw new Error('Entrada inválida en el mapa: "' + dayKey + '": ' + map[dayKey] + '.');
    }
    map[String(day)] = year;
  });
}

/** @param {Object} parameters @return {Object} @private */
function filtersFromQuery_(parameters) {
  return {
    process: parameters.process || null,
    machine: parameters.machine || null,
    fromPeriod: parameters.fromPeriod || parameters.from || null,
    toPeriod: parameters.toPeriod || parameters.to || null
  };
}

/** @param {Object} value @return {GoogleAppsScript.Content.TextOutput} @private */
function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

/** @param {*} requested @param {Array<string>} options @param {string} label @return {?string} @private */
function canonicalChoice_(requested, options, label) {
  var value = cleanString_(requested);
  if (!value || /^(?:all|todos|todas|\*)$/i.test(value)) {
    return null;
  }
  var normalized = normalizeComparable_(value);
  var match = options.filter(function (option) {
    return normalizeComparable_(option) === normalized;
  })[0];
  if (!match) {
    throw new Error('El filtro de ' + label + ' "' + value + '" no existe en la fuente.');
  }
  return match;
}

/** @param {*} requested @param {Array<Object>} options @param {string} fallback @param {string} label @return {string} @private */
function canonicalPeriod_(requested, options, fallback, label) {
  var value = cleanString_(requested);
  if (!value) {
    return fallback;
  }
  if (!/^\d{4}-(?:0[1-9]|1[0-2])$/.test(value)) {
    throw new Error('El ' + label + ' debe usar el formato AAAA-MM.');
  }
  var exists = options.some(function (option) { return option.value === value; });
  if (!exists) {
    throw new Error('El ' + label + ' "' + value + '" no existe en la fuente.');
  }
  return value;
}

/**
 * Promedio ponderado estricto. Si no existe un denominador positivo devuelve
 * null; nunca convierte una fila sin operación en un KPI de 0%.
 *
 * @param {Array<Object>} rows Filas activas.
 * @param {string} field Indicador.
 * @param {Function} weightGetter Función de denominador/peso.
 * @return {?number}
 * @private
 */
function weightedAverageStrict_(rows, field, weightGetter) {
  var weighted = 0;
  var totalWeight = 0;
  rows.forEach(function (row) {
    var value = row[field];
    if (!isFiniteNumber_(value)) {
      return;
    }
    var weight = weightGetter(row);
    if (weight > 0) {
      weighted += value * weight;
      totalWeight += weight;
    }
  });
  return totalWeight > 0 ? weighted / totalWeight : null;
}

/**
 * Determina si una fila tiene una base real de operación.
 *
 * @param {Object} row Fila normalizada.
 * @return {boolean}
 * @private
 */
function hasOperationalBase_(row) {
  return positiveWeight_(row.availableMinutes) > 0 ||
    positiveWeight_(row.scheduledMinutes) > 0 ||
    positiveWeight_(row.rate) > 0 ||
    positiveWeight_(row.target) > 0 ||
    positiveWeight_(row.productionTotal) > 0;
}

/**
 * Minutos de paro. Cuando no hay total ni desglose, los infiere de la
 * disponibilidad solo si existe una base programada; de otro modo devuelve null.
 *
 * @param {Object} row Fila normalizada.
 * @return {?number}
 * @private
 */
function rowDowntimeNullable_(row) {
  if (isFiniteNumber_(row.totalDowntime)) {
    return row.totalDowntime;
  }
  var categoryValues = DOWNTIME_FIELDS_.map(function (definition) {
    return row[definition.field];
  });
  var categoryTotal = sumNullable_(categoryValues);
  if (isFiniteNumber_(categoryTotal)) {
    return categoryTotal;
  }
  if (positiveWeight_(row.scheduledMinutes) && isFiniteNumber_(row.availability)) {
    return Math.max(0, row.scheduledMinutes * (1 - row.availability));
  }
  return null;
}

/** @param {Array<Object>} rows @param {Function} keyGetter @return {Object<string, Array<Object>>} @private */
function groupBy_(rows, keyGetter) {
  return rows.reduce(function (groups, row) {
    var key = String(keyGetter(row));
    if (!groups[key]) {
      groups[key] = [];
    }
    groups[key].push(row);
    return groups;
  }, {});
}

/** @param {Array<*>} values @return {?number} @private */
function sumNullable_(values) {
  var hasValue = false;
  var total = values.reduce(function (sum, value) {
    if (isFiniteNumber_(value)) {
      hasValue = true;
      return sum + value;
    }
    return sum;
  }, 0);
  return hasValue ? total : null;
}

/** Suma completa: un valor desconocido hace desconocido el total. @private */
function sumComplete_(values) {
  if (!values.length || !values.every(isFiniteNumber_)) {
    return null;
  }
  return values.reduce(function (sum, value) { return sum + value; }, 0);
}

/** @param {Array<*>} row @return {boolean} @private */
function isEmptyRow_(row) {
  return !row.some(function (value) {
    return value !== null && value !== undefined && cleanString_(value) !== '';
  });
}

/** @param {Object} a @param {Object} b @return {number} @private */
function compareRows_(a, b) {
  return a.periodSort - b.periodSort ||
    a.process.localeCompare(b.process, 'es') ||
    a.machine.localeCompare(b.machine, 'es');
}

/** @param {Array<string>} values @return {Array<string>} @private */
function uniqueSorted_(values) {
  var seen = {};
  return values.filter(function (value) {
    if (!value || seen[value]) {
      return false;
    }
    seen[value] = true;
    return true;
  }).sort(function (a, b) { return a.localeCompare(b, 'es', { numeric: true }); });
}

/** @param {*} value @return {string} @private */
function cleanString_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).replace(/\s+/g, ' ').trim();
}

/** @param {*} value @return {string} @private */
function normalizeHeader_(value) {
  return removeDiacritics_(cleanString_(value).toUpperCase())
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim();
}

/** @param {*} value @return {string} @private */
function normalizeComparable_(value) {
  return removeDiacritics_(cleanString_(value).toUpperCase());
}

/** @param {string} value @return {string} @private */
function removeDiacritics_(value) {
  if (value && typeof value.normalize === 'function') {
    return value.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }
  return value;
}

/** @param {string} period @return {number} @private */
function periodSort_(period) {
  var parts = String(period).split('-');
  return Number(parts[0]) * 100 + Number(parts[1]);
}

/** @param {string} period @return {string} @private */
function periodLabelFromKey_(period) {
  var parts = String(period).split('-');
  var year = Number(parts[0]);
  var month = Number(parts[1]);
  if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) {
    return String(period);
  }
  return MONTH_LABELS_[month - 1] + ' ' + year;
}

/** @param {string} period @param {number} offset @return {string} @private */
function shiftPeriod_(period, offset) {
  var parts = String(period).split('-');
  var year = Number(parts[0]);
  var month = Number(parts[1]);
  var absoluteMonth = year * 12 + (month - 1) + Number(offset || 0);
  var shiftedYear = Math.floor(absoluteMonth / 12);
  var shiftedMonth = absoluteMonth - shiftedYear * 12 + 1;
  return shiftedYear + '-' + pad2_(shiftedMonth);
}

/** @param {string} fromPeriod @param {string} toPeriod @return {Array<string>} @private */
function periodRange_(fromPeriod, toPeriod) {
  var periods = [];
  var current = fromPeriod;
  var limit = 1200;
  while (periodSort_(current) <= periodSort_(toPeriod) && periods.length < limit) {
    periods.push(current);
    current = shiftPeriod_(current, 1);
  }
  return periods;
}

/** @param {number} value @return {string} @private */
function pad2_(value) {
  return value < 10 ? '0' + value : String(value);
}

/** @param {*} value @return {boolean} @private */
function isFiniteNumber_(value) {
  return typeof value === 'number' && isFinite(value);
}

/** @param {*} value @return {number} @private */
function safeNumber_(value) {
  return isFiniteNumber_(value) ? value : 0;
}

/** @param {*} value @return {number} @private */
function positiveWeight_(value) {
  return isFiniteNumber_(value) && value > 0 ? value : 0;
}

/** @param {number} value @param {number} digits @return {number} @private */
function round_(value, digits) {
  var factor = Math.pow(10, digits);
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

/** @param {*} value @param {number} digits @return {?number} @private */
function nullableRound_(value, digits) {
  return isFiniteNumber_(value) ? round_(value, digits) : null;
}

/** @param {number} value @return {string} @private */
function formatInteger_(value) {
  return String(Math.round(value)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/** @param {string} value @return {string} @private */
function hashString_(value) {
  var hash = 0;
  for (var i = 0; i < value.length; i += 1) {
    hash = ((hash << 5) - hash) + value.charCodeAt(i);
    hash |= 0;
  }
  return String(Math.abs(hash));
}

/** @param {*} error @return {string} @private */
function publicErrorMessage_(error) {
  if (!error) {
    return 'Error desconocido.';
  }
  return cleanString_(error.message || error);
}

/** @param {*} error @return {string} @private */
function errorStack_(error) {
  return error && error.stack ? error.stack : publicErrorMessage_(error);
}
function readOtMtto_(spreadsheet) {

  var otMap = {};

  try {

    var sheet = spreadsheet.getSheetByName('OT MTTO') || spreadsheet.getSheetByName('OT_MTTO');

    if (!sheet) return otMap;

    

    var lastRow = sheet.getLastRow();

    var lastCol = sheet.getLastColumn();

    if (lastRow < 2 || lastCol < 1) return otMap;

    

    var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    var headers = values[0].map(normalizeHeader_);

    

    var indexFolio = -1, indexDesc = -1, indexAct = -1;

    for (var i = 0; i < headers.length; i++) {

      if (indexFolio < 0 && (headers[i] === 'ORDEN DE TRABAJO' || headers[i] === 'OT' || headers[i] === 'FOLIO')) indexFolio = i;

      if (indexDesc < 0 && (headers[i] === 'DESCRIPCION DETALLADA' || headers[i] === 'DESCRIPCION')) indexDesc = i;

      if (indexAct < 0 && (headers[i] === 'ACTIVIDADES' || headers[i] === 'TRABAJO REALIZADO' || headers[i] === 'ACTIVIDAD')) indexAct = i;

    }

    

    if (indexFolio < 0) return otMap;

    

    for (var r = 1; r < values.length; r++) {

      var row = values[r];

      var folio = String(row[indexFolio] || '').trim();

      if (!folio) continue;

      

      var desc = indexDesc >= 0 ? String(row[indexDesc] || '').trim() : '';

      var act = indexAct >= 0 ? String(row[indexAct] || '').trim() : '';

      if (desc || act) {

        otMap[folio] = { desc: desc, act: act };

      }

    }

  } catch (e) {

    console.warn('Error leyendo OT MTTO: ' + e.message);

  }

  return otMap;

}

/**
 * Lectura de la hoja C.DESV (Control de Desviaciones).
 */
function readDesviaciones_(spreadsheet, config, timeZone) {
  var sourceName = 'C.DESV';
  try {
    var sheet = spreadsheet.getSheetByName('C.DESV') || spreadsheet.getSheetByName('C. DESV') || spreadsheet.getSheetByName('DESVIACIONES');
    if (!sheet) throw new Error('No existe la hoja C.DESV (ni C. DESV).');
    sourceName = sheet.getName();
    if (!sheet.getLastRow()) throw new Error('La hoja ' + sourceName + ' no contiene datos.');
    var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    var result = normalizeDesviaciones_(values, config, timeZone);
    result.status.source = sourceName;
    return result;
  } catch (error) {
    return { rows: [], status: { state: 'error', source: sourceName, message: 'No fue posible leer ' + sourceName + ': ' + publicErrorMessage_(error) } };
  }
}

function normalizeDesviaciones_(values, config, timeZone) {
  if (!values || values.length < 2) {
    return { rows: [], status: { state: 'empty', source: 'C.DESV', message: 'Sin registros de desviaciones.' } };
  }
  var rawHeaders = values[0] || [];
  var headers = rawHeaders.map(function(h) { return cleanString_(h).toUpperCase(); });
  
  function findCol(candidates) {
    for (var i = 0; i < candidates.length; i++) {
      for (var j = 0; j < headers.length; j++) {
        if (headers[j].indexOf(candidates[i].toUpperCase()) >= 0) return j;
      }
    }
    return -1;
  }

  var idxNo = findCol(['NO. DESVIAC', 'DESVIACION', 'FOLIO']);
  var idxFecha = findCol(['FECHA SOLICITUD', 'FECHA']);
  var idxParteTeflu = findCol(['PARTE TEFLU', 'NO. DE PARTE TEFLU']);
  var idxDesc = findCol(['DESCRIPCI']);
  var idxCliente = findCol(['CLIENTE']);
  var idxParteCli = findCol(['PARTE CLIENTE']);
  var idxProv = findCol(['PROVEEDOR']);
  var idxArea = findCol(['AREA AFECTADA', 'ÁREA AFECTADA']);
  var idxRazon = findCol(['RAZON', 'RAZÓN']);
  var idxDefecto = findCol(['DEFECTO', 'MEDICION', 'MEDICIÓN']);
  var idxAfecta = findCol(['AFECTA']);
  var idxPz = findCol(['PIEZAS']);
  var idxBarras = findCol(['BARRAS']);
  var idxAtado = findCol(['ATADO']);
  var idxTon = findCol(['TONELADAS']);
  var idxCantOrig = findCol(['CANTIDAD ORIGINAL']);
  var idxResol = findCol(['RESOLUC']);
  var idxCond = findCol(['CONDICI', 'ACUERDOS']);
  var idxPdf = findCol(['PDF', 'ARCHIVO']);

  var monthNames = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  var rows = [];

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (isEmptyRow_(row)) continue;

    var noDesv = idxNo >= 0 ? cleanString_(row[idxNo]) : '';
    if (!noDesv && isEmptyRow_(row)) continue;

    var fechaRaw = idxFecha >= 0 ? row[idxFecha] : '';
    var year = 2026, month = 1, day = 1, dateStr = '';

    if (fechaRaw instanceof Date && !isNaN(fechaRaw.getTime())) {
      year = fechaRaw.getFullYear();
      month = fechaRaw.getMonth() + 1;
      day = fechaRaw.getDate();
      dateStr = (day < 10 ? '0' + day : day) + '/' + (month < 10 ? '0' + month : month) + '/' + year;
    } else {
      var s = cleanString_(fechaRaw);
      dateStr = s;
      var slashParts = s.split('/');
      if (slashParts.length === 3) {
        day = parseInt(slashParts[0], 10) || 1;
        month = parseInt(slashParts[1], 10) || 1;
        year = parseInt(slashParts[2], 10) || 2026;
        if (year < 100) year += 2000;
      }
    }

    var monthKey = year + '-' + (month < 10 ? '0' + month : month);
    var monthLabel = monthNames[month - 1] + ' ' + year;

    var pz = idxPz >= 0 ? (numberOrNull_(row[idxPz], false) || 0) : 0;
    var barras = idxBarras >= 0 ? (numberOrNull_(row[idxBarras], false) || 0) : 0;
    var atados = idxAtado >= 0 ? (numberOrNull_(row[idxAtado], false) || 0) : 0;
    var ton = idxTon >= 0 ? (numberOrNull_(row[idxTon], false) || 0) : 0;

    rows.push({
      id: noDesv || ('DESV-' + r),
      date: dateStr,
      year: year,
      month: month,
      period: monthKey,
      periodLabel: monthLabel,
      partTeflu: idxParteTeflu >= 0 ? cleanString_(row[idxParteTeflu]) : '',
      description: idxDesc >= 0 ? cleanString_(row[idxDesc]) : '',
      client: idxCliente >= 0 ? cleanString_(row[idxCliente]) : '',
      clientPart: idxParteCli >= 0 ? cleanString_(row[idxParteCli]) : '',
      supplier: idxProv >= 0 ? cleanString_(row[idxProv]) : '',
      area: idxArea >= 0 ? cleanString_(row[idxArea]) : '',
      reason: idxRazon >= 0 ? cleanString_(row[idxRazon]) : '',
      defect: idxDefecto >= 0 ? cleanString_(row[idxDefecto]) : '',
      affects: idxAfecta >= 0 ? cleanString_(row[idxAfecta]) : '',
      piezas: pz,
      barras: barras,
      atados: atados,
      toneladas: ton,
      quantityOriginal: idxCantOrig >= 0 ? cleanString_(row[idxCantOrig]) : '',
      resolution: idxResol >= 0 ? cleanString_(row[idxResol]) : '',
      agreements: idxCond >= 0 ? cleanString_(row[idxCond]) : '',
      pdf: idxPdf >= 0 ? cleanString_(row[idxPdf]) : ''
    });
  }

  return {
    rows: rows,
    status: { state: 'ok', source: 'C.DESV', total: rows.length }
  };
}

function getInitialData(forceRefresh) {
  var config = runtimeConfig_();
  var dataset = loadDataset_(forceRefresh);
  return {
    meta: {
      oeeTarget: config.oeeTarget,
      updatedAt: dataset.updatedAt || new Date().toISOString()
    },
    rows: dataset.rows,
    cdemSummary: dataset.cdemSummary,
    cdemStatus: dataset.cdemStatus,
    scrapSummary: dataset.scrapSummary,
    scrapStatus: dataset.scrapStatus,
    otMap: dataset.otMap,
    deviations: dataset.deviations || [],
    deviationsStatus: dataset.deviationsStatus || {},
    actions: dataset.actions || ACTIONS_SEED_,
    actionsActivities: dataset.actionsActivities || ACTIONS_ACTIVITIES_SEED_,
    actionsStatus: dataset.actionsStatus || { state: 'seed', source: 'ACCIONES 2026.xlsx' },
    criticalDecisions: readCriticalDecisions_(),
    criticalAccess: criticalAccess_()
  };
}

/**
 * Mantiene el cach fresco en segundo plano para que la primera carga sea inmediata.
 * Ejecutar esta funcin manualmente UNA VEZ desde el editor de Apps Script.
 */
function setupWarmingTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'warmupCache_') {
      return; // Ya existe
    }
  }
  ScriptApp.newTrigger('warmupCache_')
    .timeBased()
    .everyHours(1) // Actualiza el cach cada hora automticamente
    .create();
}

function warmupCache_() {
  try {
    loadDataset_(true); // Fuerza la lectura de Google Sheets y reconstruye el cach
  } catch (e) {
    console.error('Error en warmup: ' + e.message);
  }
}/**
 * Devuelve el resumen de producción para las nuevas tablas.
 * Exclusivamente para la vista de "Producción".
 */
function getProductionSummary() {
  var config = runtimeConfig_();
  var dataset = loadDataset_(false);
  var rows = dataset.rows || [];

  function processGroup(allowedProcesses) {
    var filtered = rows.filter(function(r) { return allowedProcesses.indexOf(r.process) >= 0; });
    var byYearAndMonth = {};
    filtered.forEach(function(r) {
      if (!r.period) return;
      var parts = r.period.split('-');
      var year = parts[0];
      var month = parseInt(parts[1], 10);
      if (!byYearAndMonth[year]) byYearAndMonth[year] = {};
      if (!byYearAndMonth[year][month]) byYearAndMonth[year][month] = { production: 0, scrap: 0 };
      
      var prod = (typeof r.piecesOk === 'number') ? r.piecesOk : r.production;
      if (typeof prod === 'number' && isFinite(prod)) {
        byYearAndMonth[year][month].production += prod;
      }
      var scrap = r.totalScrap;
      if (typeof scrap === 'number' && isFinite(scrap)) {
        byYearAndMonth[year][month].scrap += scrap;
      }
    });
    return byYearAndMonth;
  }

  return {
    corte: processGroup(['CORTE LASER', 'CORTE DISCO']),
    dobladoRouter: processGroup(['DOBLADO', 'ROUTER'])
  };
}
/**
 * Devuelve el resumen de producción para las nuevas tablas.
 * Exclusivamente para la vista de "Producción".
 */
function getProductionSummary() {
  var config = runtimeConfig_();
  var dataset = loadDataset_(false);
  var rows = dataset.rows || [];

  function processGroup(allowedProcesses) {
    var filtered = rows.filter(function(r) { return allowedProcesses.indexOf(r.process) >= 0; });
    var byYearAndMonth = {};
    filtered.forEach(function(r) {
      if (!r.period) return;
      var parts = r.period.split('-');
      var year = parts[0];
      var month = parseInt(parts[1], 10);
      if (!byYearAndMonth[year]) byYearAndMonth[year] = {};
      if (!byYearAndMonth[year][month]) byYearAndMonth[year][month] = { production: 0, scrap: 0 };
      
      var prod = r.productionTotal;
      if (typeof prod === 'number' && isFinite(prod)) {
        byYearAndMonth[year][month].production += prod;
      }
      var scrap = r.totalScrap;
      if (typeof scrap === 'number' && isFinite(scrap)) {
        byYearAndMonth[year][month].scrap += scrap;
      }
    });
    return byYearAndMonth;
  }

  return {
    corte: processGroup(['CORTE LASER', 'CORTE DISCO']),
    dobladoRouter: processGroup(['DOBLADO', 'ROUTER'])
  };
}
// ==========================================
// SEGUIMIENTO DE ACCIONES (ACCIONES 2026)
// ==========================================

var ACTIONS_SEED_ = [
  {
    "item": 1,
    "tipo": "SCRAP FUERA DE META",
    "categoria": "Modo de Falla",
    "numeroParte": "147-5190",
    "descripcion": "PORCENTAJE DE SCRAP FUERA DE META 3.33 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON AGULO ABIERTO Y BARRENO DESPLAZADO EN EL PROCESO DE DOBLADO EN LA MÁQUINA D-229.",
    "causas": [
      {
        "causeId": "1",
        "causeText": "1. MATERIA PRIMA HA PRESENTADO TORSIÓN Y PANDEAMIENTO, SE INSPECCIONA Y SE CORTAN BARRAS CON TOLERANCIA",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "1.SE IMPLEMENTA DISPOSITIVO EN CORTE LASER PARA ASEGURAR PIEZAS FUNCIONALES EN DOBLADO",
            "date": "2026-10-30",
            "responsible": "MIGUEL ALCANTARILLA\nGRECIA ESCUDERO",
            "status": "En Proceso",
            "compliance": 0.5,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1. EJECUTAR MODIFICACIÓN EN EL DISPOSITIVO DE CORTE LÁSER (DISPOSITIVO CORREGIDO EN PRODUCCIÓN ENTREGADO-VALIDAR NOTCH JUEGO)-(VIERNES 30/10/26) PENDIENTE VALIDAR.\nSE TIENEN 2 COTIZACIONES, FALTA 1, SIGUIENTE PASO, SELECCIONAR PROVEEDOR.",
            "date": "2026-10-30",
            "responsible": "MIGUEL ALCANTARILLA\nGRECIA ESCUDERO",
            "status": "En Proceso",
            "compliance": 0.9,
            "notes": ""
          },
          {
            "actionId": "1.1",
            "type": "Correctiva",
            "actionText": "1.1 SUPERVISOR LT8 DEBERÁ IDENTIFICAR Y REGISTRAR CONDICIONES Y AFECTACIONES DE MATERIA PRIMA EN EL PROCESO DE CORTE LÁSER (INMEDIATAMENTE-SE ENVIA POR CORREO)",
            "date": "INMEDIATO",
            "responsible": "TOMAS DEL CARMEN",
            "status": "Cerrada",
            "compliance": 1.0,
            "notes": ""
          }
        ]
      },
      {
        "causeId": "2",
        "causeText": "2. DISPOSITIVO EN DOBLADO ESTA RESTRINGIDO Y RECHAZA PIEZAS FUNCIONALES",
        "containments": [],
        "correctives": [
          {
            "actionId": "2",
            "type": "Correctiva",
            "actionText": "2. REALIZAR MODIFICACIÓN EN EL DISPOSITIVO DE DOBLADO (GESTIONAR COTIZACIÓN) DAR A LOS BLOQUES .025 DE TOLERANCIA EN AMBOS BLOQUES. MIGUEL/GRECIA (VIERNES 30/10/26)",
            "date": "COTIZACIÓN 30/10/2026",
            "responsible": "MIGUEL ALCANTARILLA\nGRECIA ESCUDERO",
            "status": "En Proceso",
            "compliance": 0.2,
            "notes": ""
          }
        ]
      },
      {
        "causeId": "3",
        "causeText": "3.OPERADOR NO LLEVA  A TOPE LA PIEZA PROVOCANDO DOBLEZ INCORRECTO",
        "containments": [],
        "correctives": [
          {
            "actionId": "3",
            "type": "Correctiva",
            "actionText": "3.VALIDAR POKA-YOKE EN EL CAÑÓN PARA ASEGURAR QUE LA PIEZA LLEGUE A TOPE PREVIO AL DOBLADO (MIGUEL SANCHEZ VALIDAR PROXIMA CORRIDA DE NP).",
            "date": "PENDIENTE FECHA",
            "responsible": "MIGUEL ALCANTARILLA\nMIGUEL SANCHEZ",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          },
          {
            "actionId": "3.1",
            "type": "Correctiva",
            "actionText": "3.1 IDENTIFICAR PERSONAL CAPACITADO Y DEFINIR PLANTILLA FIJA PARA ESTE NÚMERO DE PARTE-SE PERDIO PLANTILLA POR ROTACIÓN, ACERCAMIENTO CON PERSONAL PARA FIJAR PLATILLA (26/10/26), YA ESTA LA PLANTILLA FIJA.",
            "date": "2026-10-26",
            "responsible": "MIGUEL ALCANTARILLA",
            "status": "Cerrada",
            "compliance": 1.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "2026-10-26",
    "responsables": "GRECIA ESCUDERO, MIGUEL ALCANTARILLA, MIGUEL SANCHEZ, TOMAS DEL CARMEN",
    "cumplimiento": 0.62,
    "totalAcciones": 5
  },
  {
    "item": 2,
    "tipo": "SCRAP FUERA DE META",
    "categoria": "Modo de Falla",
    "numeroParte": "5092464",
    "descripcion": "PORCENTAJE DE SCRAP FUERA DE META 1.59 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON BARRENO DESCENTRADO EN CL-007",
    "causas": [
      {
        "causeId": "1",
        "causeText": "1. MATERIA PRIMA HA PRESENTADO TORSIÓN Y PANDEAMIENTO, SE INSPECCIONA Y SE CORTAN BARRAS CON TOLERANCIA",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "1.SE INSPECCIONA MATERIA PRIMA, SE SELECCIONA BARRAS CON TOLERANCIA PARA CORTE, SE REALIZAN AJUSTE DE COMPENSACIÓN PARA ABSORBER PANDEAMIENTO",
            "date": "-",
            "responsible": "ISAI PILOTZI",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1. SE SOLICITO AUTORIZACIÓN DE CAMBIO DE PROVEEDOR DE MP PARA ASEGURAR \"TORSIÓN Y RECTITUD\" DENTRO DE ESPECIFICACIÓN, NO SE AUTORIZO.",
            "date": "-",
            "responsible": "ISAI PILOTZI",
            "status": "Cerrada",
            "compliance": 1.0,
            "notes": ""
          },
          {
            "actionId": "2",
            "type": "Correctiva",
            "actionText": "2. DEJAR PROGRAMA A AJUSTADORES PARA REALIZAR COMPENSACIONES DE ACUERDO A CONDICIONES DE MATERIA PRIMA\n2.1 CAPACITAR A LOS AJUSTADORES EN COMPENSACIONES AMBAS NAVES",
            "date": "2026-10-30",
            "responsible": "MIGUEL SANCHEZ",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          },
          {
            "actionId": "3",
            "type": "Correctiva",
            "actionText": "3. ACTUALIZAR PLAN DE DESARROLLO DE AJUSTADORES ANA FERMAN, AL CERRAR CAPACITACIONES 04/11/2026.",
            "date": "PENDIENTE",
            "responsible": "ANA FERMAN",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "PENDIENTE",
    "responsables": "ANA FERMAN, ISAI PILOTZI, MIGUEL SANCHEZ",
    "cumplimiento": 0.3333333333333333,
    "totalAcciones": 3
  },
  {
    "item": 3,
    "tipo": "SCRAP FUERA DE META",
    "categoria": "Modo de Falla",
    "numeroParte": "BEH91595",
    "descripcion": "PORCENTAJE DE SCRAP FUERA DE META",
    "causas": [
      {
        "causeId": "1",
        "causeText": "1. EL SISTEMA DE BUSQUEDA DE SOLDADURA SE ENCUENTRA DESHABILITADO POR FALLAS EN MÁQUINA CL-009, EL OPERADOR TIENE QUE COLOCAR LA SOLDADURA MANUALMENTE",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "1. AYUDA VISUAL PARA OPERADOR COMO COLOCAR POSICIÓN DE SOLDADURA",
            "date": "2026-06-01",
            "responsible": "ROGELIO RODRIGUEZ",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1. SOLICITAR  SOPORTE TÉCNICO CON PROVEEDOR BLM PARA DIAGNÓSTICO DE LA FALLA.\n2. SE REALIZA EL REEMPLAZO DEL CABLE HMI.\n3. SE VALIDA ESTABILIDAD OPERATIVA SIN RECURRENCIA DE LA FALLA.",
            "date": "2026-06-01",
            "responsible": "ROGELIO RODRIGUEZ",
            "status": "Cerrada",
            "compliance": 1.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "2026-06-01",
    "responsables": "ROGELIO RODRIGUEZ",
    "cumplimiento": 1.0,
    "totalAcciones": 1
  },
  {
    "item": 4,
    "tipo": "SCRAP FUERA DE META",
    "categoria": "Modo de Falla",
    "numeroParte": "UC41059",
    "descripcion": "PORCENTAJE DE SCRAP FUERA DE META GENERADO PRINCIPALMENTE EN DOBLADORA 233 Y MB-001",
    "causas": [
      {
        "causeId": "1",
        "causeText": "1. LAS PIEZAS AL ACOMODARSE SE GOLPEAN GENERANDO MARCA PROFUNDA",
        "containments": [],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1.SEPARADOR ENTRE CAMAS DE PIEZAS (REPLICAR EN TRASPALEO DE PIEZAS)",
            "date": "2026-11-04",
            "responsible": "MIGUEL ALCANTARILLA",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      },
      {
        "causeId": "2",
        "causeText": "2. OPERADOR INTRODUCE LA PIEZA EN POSICIÓN INCORRECTA QUEDANDO PIEZA MAL DOBLADA",
        "containments": [],
        "correctives": [
          {
            "actionId": "2",
            "type": "Correctiva",
            "actionText": "2.VALIDAR AGREGAR PRECORTE EN PIEZA PARA IDENTIFICAR POSICIÓN DE PIEZAS, PENDIENTE VALIDAR CON MIGUEL SANCHEZ, PENDIENTE POR VALIDAR PRODUCCION FLUJO DE PROCESO PARA CORREGIR ÁREA DE OPORTUNIDAD DE OPERADORES.",
            "date": "PENDIENTE REVISION",
            "responsible": "MIGUEL ALCANTARILLA\nMIGUEL SANCHEZ",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      },
      {
        "causeId": "3",
        "causeText": "3. BARRENO DESCENTRADO DESDE CORTE LASER, GAGE COMPRA PIEZAS CON BARRENO DESCENTRADO (MATERIA PRIMA EN LA MAXIMA REDUCE TOLERANCIA DE CENTRADO, EN LA MINIMA TE DA MAYOR TOLERANCIA CON BARRENOS DESCENTRADOS), BARRENOS NO ESTAN CONCENTRICOS.",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "3. IMPLEMENTAR  GAGE CON PERNO PASADO EN EL CORTE LASER (MIGUEL ALCANTARILLA/GRECIA ESCUDERO)-15/11/2026\nDISEÑO DE GAGE PARA 28/10/26, AGREGAR A LA DOCUMENTACION TECNICA",
            "date": "PENDIENTE",
            "responsible": "GRECIA ESCUDERO\nMIGUEL ALCANTARILLA",
            "status": "En Proceso",
            "compliance": 0.5,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "3",
            "type": "Correctiva",
            "actionText": "3. SE SELECCIONO ATADOS CON BARRAS FUERA DE ESPECIFICACION PARA VALIDAR CON PROVEEDOR LAS CONDICIONES DE LA MATERIA PRIMA.",
            "date": "PENDIENTE",
            "responsible": "GRECIA ESCUDERO\nMIGUEL ALCANTARILLA",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      },
      {
        "causeId": "4",
        "causeText": "4. DISTANCIA DE WIPER MENOR A 43 MM NO ASIENTA CORRECTAMENTE EN LA RODAJA GENERANDO SCRAP .",
        "containments": [],
        "correctives": [
          {
            "actionId": "4",
            "type": "Correctiva",
            "actionText": "4. CABIOS DE WIPER E INSERTO RESPECTO A VIDA UTIL (INMEDIATO) CHRISTIAN, ESTAR MONITOREANDO EN AUTONOMOS, GRADUADO DE WIPER, PORTAINSERTO (), SE FIJA VIDA UTIL WIPER PORTA INSERTO, FALTA DEFINIR CADA CUANTO, PRIMER PASO QUE SE CAPTURE MANUAL EN LA BITACORA (INMEDIATO) BITACORAS EN PISO, LOS VIERNES CHRISTIAN AUDITA QUE SE ESTEN REALIZANDO LOS REPORTES.",
            "date": "INMEDIATO",
            "responsible": "CHRISTIAN",
            "status": "Cerrada",
            "compliance": 1.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "INMEDIATO",
    "responsables": "CHRISTIAN, GRECIA ESCUDERO, MIGUEL ALCANTARILLA, MIGUEL SANCHEZ",
    "cumplimiento": 0.25,
    "totalAcciones": 4
  },
  {
    "item": 5,
    "tipo": "SCRAP FUERA DE META",
    "categoria": "Modo de Falla",
    "numeroParte": "749-06174C\n749-06175C",
    "descripcion": "PORCENTAJE DE SCRAP FUERA DE META GENERADO PRINCIPALMENTE EN DOBADLORA D-232",
    "causas": [
      {
        "causeId": "1",
        "causeText": "1. MATERIA PRIMA VIENE CON COSTURA CORRIDA Y CHISPORROTEO GENERA QUE EL MANDRIL GENERE ABOMBAMIENTO EN EL DOBLEZ",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "1.1PARA MATERIA PRIMA PROCESADA EN MOLINO 4, REALIZAR LIBERACIÓN CON EL EQUIPO DE TEFLU (PRODUCCIÓN, CALIDAD, INGENIERÍA, GERENCIA, ETC.), DOCUMENTANDO LA CONDICIÓN BAJO DESVIACIÓN.",
            "date": "2026-04-30",
            "responsible": "MIGUEL ALCANTARILLA\nGRECIA ESCUDERO\nISAI PILOTZI",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1. DEFINIR DIMENSIONES FUNCIONALES DE ESLABONES  Y COMPARTIR DIBUJO A CALIDAD PARA FABRICAR GAGE 1. 03/04/2026\n1.2MANDAR A FABRICAR GAGE PARA CALIDAD Y PRODUCCIÓN 1. 03/04/2026\n1.3 LIBERAR LA MATERIA PRIMA EN RECIBO MEDIANTE GAGE  (ESLABONES CON VARILLA) ASEGURANDO QUE LA MATERIA PRIMA DE MOLINO 6 SE ENCUENTRE LIBRE DE CHISPORROTEO; CUALQUIER MATERIAL CON ESTA CONDICIÓN DEBERÁ SER RECHAZADO 18/04/26\n1.4 VALIDAR EN PROCESO DE CORTE LÁSER LA CONDICIÓN DE COSTURA MEDIANTE GAGE (ESLABONES CON VARILLA)  18/04/26.",
            "date": "2026-04-30",
            "responsible": "MIGUEL ALCANTARILLA\nGRECIA ESCUDERO\nISAI PILOTZI",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "2026-04-30",
    "responsables": "GRECIA ESCUDERO, ISAI PILOTZI, MIGUEL ALCANTARILLA",
    "cumplimiento": 0.0,
    "totalAcciones": 1
  },
  {
    "item": 6,
    "tipo": "DISPONIBILIDAD DE MÁQUINA FUERA DE META",
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CORTADORA DISCO C-123",
    "descripcion": "ALARMAS CONSTANTES EN MAQUINA LA CAUSA RAIZ NO  ESTA DETECTADA",
    "causas": [
      {
        "causeId": "1",
        "causeText": "CAUSA RAIZ NO ENCONTRADA AUN",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "1. SE REVISA TARJETA PROFIBUS, SE MANDA A REPARACION Y NO SE ENCUETRA LA FALLA\n2. SE SOLICITA TARJETA NUEVA CON TIEMPO DE ENTREGA DE 15 DIAS\n3. TARJETA SE ENTREGO EL 21/10/26 SE INSTALO, SE SIGUE PRESENTANDO LA FALLA\n4. SE RELEVAN ENTRADAS Y SALIDAS DE TARJETA\n5. SE RELEVAN BOBINAS Y SE PONE UNA FUENTE ALTERNA PARA DISMINUIR CONSUMO DE ENERGIA DIRECTA DE LA TARJETA \n6. SE SOLICITA SOPORTE DE PROVEEDOR, AUN NO SE CONFIRMA FECHA DE VISITA.",
            "date": "INMEDIATO",
            "responsible": "VICTOR HERNANDEZ",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ],
        "correctives": []
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "INMEDIATO",
    "responsables": "VICTOR HERNANDEZ",
    "cumplimiento": 0.0,
    "totalAcciones": 1
  },
  {
    "item": 7,
    "tipo": "DISPONIBILIDAD DE MÁQUINA FUERA DE META Y SCRAP",
    "categoria": "Modo de Falla",
    "numeroParte": "D-232\n749-06114",
    "descripcion": "VARILLA MODIFICADA NO TIENE SUJECION CORRECTA EN EL BLOQUE DE FIJACION DE LA VARILLA POR LO QUE ESTA GENERANDO TIEMPOS PROLONGADOS DE AJUSTE Y SCRAP",
    "causas": [
      {
        "causeId": "1",
        "causeText": "1.ESTA TENIENDO JUEGO EN LA PARTE DEL BLOQUE DE FIJACION DE LA VARILLA PORQUE EL DISEÑO QUE SE TIENE AL MOMENTO DE SOLDAR LAS ARANDELAS NO QUEDAN A LA MISMA DISTANCIA DE SEPARACIÓN, AL REALIZAR DOBLEZ SE MUEVE Y GENERA QUE LA PIEZA TENGA ANGULO ABIERTO/BARRENO DESPLAZADO",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "1. ARMAR VARILLA CON CONDICIONES ADECUADAS PARA EVITAR DESPLAZAMIENTO 29/10/2026,",
            "date": "",
            "responsible": "CHRISTIAN",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1.1 REALIZAR DISEÑO DE DISPOSITIVO PARA SOLDADURA\n1.1 FABRICAR DISPOSITIVO DE SOLDADURA\n1.2 GENERAR INSTRUCCIÓN DE TRABAJO PARA USO DE DISPOSITIVO Y CAPACITACION DEL SOLDADOR",
            "date": "",
            "responsible": "CHRISTIAN",
            "status": "Cerrada",
            "compliance": 1.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "-",
    "responsables": "CHRISTIAN",
    "cumplimiento": 1.0,
    "totalAcciones": 1
  },
  {
    "item": 8,
    "tipo": "DISPONIBILIDAD DE MÁQUINA FUERA DE META",
    "categoria": "Seguimiento Independiente",
    "numeroParte": "D-232",
    "descripcion": "PLANTILLA INCOMPLETA",
    "causas": [
      {
        "causeId": "1",
        "causeText": "NO SE CUENTA CON AYUDANTE GENERAL CONTRATADO EN EL TURNO DEL SUPERVISOR  BLADIMIR Y JOEL, EL PERSONAL CONTRATADO RENUNCIO POR TIPO DE TRABAJO, NO SE HAN CUBIERTO LAS VACANTES POR FALTA DE RECLUTADOR",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "1. YA SE CONTRATARON A LOS DOS AYUDANTES Y YA SE TIENE CUBIERTO ACTUALMENTE LAS VACANTES",
            "date": "PENDIENTE",
            "responsible": "ANA FERMAN",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "ANALIZAR ADAPATACION DE PERSONAL DE NUEVO INGRESO EN DOBLADO PARA DETERMINAR TIEMPO DE LIBERACION AL 100% A OPERACIONES\nSOLICITAR METODOS DE TRABAJO A SEGURIDAD INDUSTRIAL PARA DISMINUIR CARGA DE TRABAJO",
            "date": "PENDIENTE",
            "responsible": "ANA FERMAN",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "PENDIENTE",
    "responsables": "ANA FERMAN",
    "cumplimiento": 0.0,
    "totalAcciones": 1
  },
  {
    "item": 9,
    "tipo": "DISPONIBILIDAD DE MÁQUINA FUERA DE META",
    "categoria": "Seguimiento Independiente",
    "numeroParte": "VARIAS MÁQUINAS",
    "descripcion": "TRANSPORTE SE DEMORA EN LLEGADA A PLANTA POR TRÁFICO GENERANDO RETRASO EN OPERACIONES",
    "causas": [
      {
        "causeId": "1",
        "causeText": "1. TRÁFICO EN RUTA HABITUAL GENERA DEMORA EN LLEGADA DE PERSONAL A PLANTA",
        "containments": [],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1. REVISAR ANÁLISIS DE IMPACTO DE RUTA Y EVALUAR AJUSTE DE HORARIOS / RECORRIDO",
            "date": "PENDIENTE",
            "responsible": "RECURSOS HUMANOS\nLOGÍSTICA",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      },
      {
        "causeId": "2",
        "causeText": "2. FALTA DE COORDINACIÓN ENTRE NECESIDADES PARA TEFLU Y MAQUILA",
        "containments": [],
        "correctives": [
          {
            "actionId": "2",
            "type": "Correctiva",
            "actionText": "2. COORDINACIÓN ENTRE NECESIDADES PARA TEFLU Y MAQUILA (ESTABLECER CANAL DIRECTO DE MONITOREO)",
            "date": "PENDIENTE",
            "responsible": "PLANEACIÓN\nLOGÍSTICA",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "PENDIENTE",
    "responsables": "LOGÍSTICA, PLANEACIÓN, RECURSOS HUMANOS",
    "cumplimiento": 0.0,
    "totalAcciones": 2
  },
  {
    "item": 10,
    "tipo": "SCRAP FUERA DE META",
    "categoria": "Modo de Falla",
    "numeroParte": "BEWO  80142 5EF0A",
    "descripcion": "DIMENSIÓN FUERA DE ESPECIFICACIÓN: PIEZAS LARGAS SE PUEDEN RETRABAJAR, PERO SE TIRARON",
    "causas": [
      {
        "causeId": "1",
        "causeText": "1. DESCONOCIMIENTO EN PISO DEL PROCEDIMIENTO DE RECUPERACIÓN DE PIEZAS LARGAS; DISPOSITIVO EN EMPAQUE E INSPECCIÓN CON PIEZA MASTER GENERANDO FUGA DE MATERIAL",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "1. RETIRAR DISPOSITIVO DE EMPAQUE Y PONER EN PISO PARA INSPECCIÓN; SEGREGAR PIEZAS LARGAS PARA EVITAR SCRAP",
            "date": "INMEDIATO",
            "responsible": "CALIDAD\nPRODUCCIÓN",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1. CAPACITACIÓN A PERSONAL SOBRE CRITERIO DE RECUPERACIÓN Y RETRABAJO DE PIEZAS LARGAS CON GAGE OFICIAL",
            "date": "INMEDIATO",
            "responsible": "CALIDAD\nPRODUCCIÓN",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      },
      {
        "causeId": "2",
        "causeText": "2. NÚMEROS DE PARTE DE LUNKOMEX QUE SE CORTAN EN LÁSER: LONGITUD NO COMPATIBLE DIRECTAMENTE CON CORTE LÁSER",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "2. SOLICITAR A PROGRAMACIÓN LÁSER PLAN DE CORTE EMERGENTE",
            "date": "INMEDIATO",
            "responsible": "PROGRAMACIÓN\nCORTE LÁSER",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "2",
            "type": "Correctiva",
            "actionText": "2. SOLICITAR A PROGRAMACIÓN DESPUNTAR TUBO PARA PODER METERLO EN LA LÁSER Y REALIZAR PRUEBA PILOTO",
            "date": "INMEDIATO",
            "responsible": "PROGRAMACIÓN\nCORTE LÁSER",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "INMEDIATO",
    "responsables": "CALIDAD, CORTE LÁSER, PRODUCCIÓN, PROGRAMACIÓN",
    "cumplimiento": 0.0,
    "totalAcciones": 2
  },
  {
    "item": 11,
    "tipo": "DISPONIBILIDAD DE MÁQUINA FUERA DE META",
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CL-009",
    "descripcion": "3 VISITAS DE BLM, CALIBRAN SISTEMA DE VISIÓN TRABAJA BIEN 2 SEMANAS Y SE VUELVE A PRESENTAR FALLA, SE CAMBIARON LENTES, CABLES, PIERDE CALIBRACIÓN, SE PRESENTA 4 VISITA",
    "causas": [
      {
        "causeId": "1",
        "causeText": "NO SE TIENE CAUSA RAÍZ DEFINITIVA ACTUALMENTE (SE CAMBIARON CABLES, LENTES Y SWITCH SIN SOLUCIÓN DEFINITIVA, PROCESO DE CALIBRACIÓN TARDADO)",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "1. CONTENCIÓN: SE REINICIA LA MÁQUINA CADA VEZ QUE SE PRESENTA LA FALLA\n2. CALIBRACIÓN TEMPORAL POR PROVEEDOR BLM",
            "date": "PENDIENTE VISITA",
            "responsible": "ROGELIO RODRIGUEZ\nBLM",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1. SOLICITAR 4TA VISITA DE ESPECIALISTA BLM PARA DIAGNÓSTICO PROFUNDO EN HARDWARE/SOFTWARE Y SEGUIMIENTO A GARANTÍA",
            "date": "PENDIENTE VISITA",
            "responsible": "ROGELIO RODRIGUEZ\nBLM",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "PENDIENTE VISITA",
    "responsables": "BLM, ROGELIO RODRIGUEZ",
    "cumplimiento": 0.0,
    "totalAcciones": 1
  },
  {
    "item": 12,
    "tipo": "DISPONIBILIDAD DE MÁQUINA FUERA DE META",
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CL-005",
    "descripcion": "FALLA RECURRENTE DE CENTRADORES ALTOS EN CL-005",
    "causas": [
      {
        "causeId": "1",
        "causeText": "PIEZAS DESGASTADAS NYLAMIND (REFACCIÓN) Y MAL AJUSTE DE SENSORES (AGREGAR A LA INSTRUCCIÓN DE PREVENTIVO, CAMBIO CADA AÑO DE LAS PIEZAS), REBABAS EN CREMALLERAS.",
        "containments": [
          {
            "actionId": "1",
            "type": "Contención",
            "actionText": "REPARACIÓN Y AJUSTE MECÁNICO TEMPORAL EN MÁQUINA",
            "date": "2026-10-30",
            "responsible": "ROGELIO RODRIGUEZ",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1. DAR DE ALTA REFACCIONES 30/10/2026 (GESTIONAR LIBERACIÓN DE DIRECCIÓN)",
            "date": "2026-10-30",
            "responsible": "ROGELIO RODRIGUEZ",
            "status": "En Proceso",
            "compliance": 0.5,
            "notes": ""
          },
          {
            "actionId": "2",
            "type": "Correctiva",
            "actionText": "2. INSTRUCCIÓN DE TRABAJO PARA INSTALACIÓN DE CENTRADORES (13/11/26) Y EVALUACIÓN DE TÉCNICOS DE MANTENIMIENTO (20/11/26)",
            "date": "2026-11-13",
            "responsible": "ROGELIO RODRIGUEZ",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "2026-11-13",
    "responsables": "ROGELIO RODRIGUEZ",
    "cumplimiento": 0.25,
    "totalAcciones": 2
  },
  {
    "item": 13,
    "tipo": "PORCENTAJE DE PAROS MAYOR A PORCENTAJE PERMITIDO",
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CORTE LÁSER",
    "descripcion": "NO SE CUENTA CON PLANTILLA COMPLETA Y NO SE TIENE EL PERSONAL CALIFICADO (5 INTENTOS DE REPARACIÓN EN FALLA DE CENTRADORES)",
    "causas": [
      {
        "causeId": "1",
        "causeText": "TÉCNICOS NO ESTÁN FAMILIARIZADOS CON MÁQUINAS QUE SE TIENEN EN PLANTA; PERSONAL DE INGRESO TIENE CONOCIMIENTO PERO NO ESTÁ FAMILIARIZADO CON LAS MÁQUINAS DE PLANTA.",
        "containments": [],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1. PLANTILLA COMPLETA Y ACELERAMIENTO DE PERSONAL COMPETENTE (CONTRATAR PERSONAL Y BASE DE PLANTILLA PARA SUPLIR INMEDIATAMENTE)",
            "date": "PENDIENTE",
            "responsible": "RECURSOS HUMANOS\nMANTENIMIENTO",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          },
          {
            "actionId": "2",
            "type": "Correctiva",
            "actionText": "2. PLAN DUAL CON CONALEP PARA TRAER PRACTICANTES PARA FUTUROS TÉCNICOS",
            "date": "PENDIENTE",
            "responsible": "RECURSOS HUMANOS",
            "status": "Abierta",
            "compliance": 0.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "PENDIENTE",
    "responsables": "MANTENIMIENTO, RECURSOS HUMANOS",
    "cumplimiento": 0.0,
    "totalAcciones": 2
  },
  {
    "item": 14,
    "tipo": "FALLAS DE GARRAS",
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CL-007",
    "descripcion": "DISEÑO DE GARRA PARA TRAPECIO",
    "causas": [
      {
        "causeId": "1",
        "causeText": "ÁNGULO DE CONO ORIGINAL (45°) Y LONGITUD DE GARRA PROVOCABAN FALLAS DE SUJECIÓN EN PERFIL TRAPECIO",
        "containments": [],
        "correctives": [
          {
            "actionId": "1",
            "type": "Correctiva",
            "actionText": "1. SE REALIZÓ MODIFICACIÓN DE ANGULACIÓN DE CONO DE GARRAS, DE 45 GRADOS A 42 GRADOS, SE RECORTO 1 MM DE LONGITUD.",
            "date": "INMEDIATO",
            "responsible": "INGENIERÍA\nMANTENIMIENTO",
            "status": "Cerrada",
            "compliance": 1.0,
            "notes": ""
          }
        ]
      }
    ],
    "fechaInicial": "-",
    "fechaFinal": "INMEDIATO",
    "responsables": "INGENIERÍA, MANTENIMIENTO",
    "cumplimiento": 1.0,
    "totalAcciones": 1
  }
];

var ACTIONS_ACTIVITIES_SEED_ = [
  {
    "item": 1,
    "categoria": "Modo de Falla",
    "numeroParte": "147-5190",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META 3.33 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON AGULO ABIERTO Y BARRENO DESPLAZADO EN EL PROCESO DE DOBLADO EN LA MÁQUINA D-229.",
    "causaId": "1",
    "causaTexto": "1. MATERIA PRIMA HA PRESENTADO TORSIÓN Y PANDEAMIENTO, SE INSPECCIONA Y SE CORTAN BARRAS CON TOLERANCIA",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "1.SE IMPLEMENTA DISPOSITIVO EN CORTE LASER PARA ASEGURAR PIEZAS FUNCIONALES EN DOBLADO",
    "responsable": "MIGUEL ALCANTARILLA\nGRECIA ESCUDERO",
    "fechaCompromiso": "2026-10-30",
    "estado": "En Proceso",
    "cumplimiento": 0.5,
    "notes": ""
  },
  {
    "item": 1,
    "categoria": "Modo de Falla",
    "numeroParte": "147-5190",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META 3.33 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON AGULO ABIERTO Y BARRENO DESPLAZADO EN EL PROCESO DE DOBLADO EN LA MÁQUINA D-229.",
    "causaId": "1",
    "causaTexto": "1. MATERIA PRIMA HA PRESENTADO TORSIÓN Y PANDEAMIENTO, SE INSPECCIONA Y SE CORTAN BARRAS CON TOLERANCIA",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1. EJECUTAR MODIFICACIÓN EN EL DISPOSITIVO DE CORTE LÁSER (DISPOSITIVO CORREGIDO EN PRODUCCIÓN ENTREGADO-VALIDAR NOTCH JUEGO)-(VIERNES 30/10/26) PENDIENTE VALIDAR.\nSE TIENEN 2 COTIZACIONES, FALTA 1, SIGUIENTE PASO, SELECCIONAR PROVEEDOR.",
    "responsable": "MIGUEL ALCANTARILLA\nGRECIA ESCUDERO",
    "fechaCompromiso": "2026-10-30",
    "estado": "En Proceso",
    "cumplimiento": 0.9,
    "notes": ""
  },
  {
    "item": 1,
    "categoria": "Modo de Falla",
    "numeroParte": "147-5190",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META 3.33 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON AGULO ABIERTO Y BARRENO DESPLAZADO EN EL PROCESO DE DOBLADO EN LA MÁQUINA D-229.",
    "causaId": "1",
    "causaTexto": "1. MATERIA PRIMA HA PRESENTADO TORSIÓN Y PANDEAMIENTO, SE INSPECCIONA Y SE CORTAN BARRAS CON TOLERANCIA",
    "tipoAccion": "Correctiva",
    "accionId": "1.1",
    "accionTexto": "1.1 SUPERVISOR LT8 DEBERÁ IDENTIFICAR Y REGISTRAR CONDICIONES Y AFECTACIONES DE MATERIA PRIMA EN EL PROCESO DE CORTE LÁSER (INMEDIATAMENTE-SE ENVIA POR CORREO)",
    "responsable": "TOMAS DEL CARMEN",
    "fechaCompromiso": "INMEDIATO",
    "estado": "Cerrada",
    "cumplimiento": 1.0,
    "notes": ""
  },
  {
    "item": 1,
    "categoria": "Modo de Falla",
    "numeroParte": "147-5190",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META 3.33 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON AGULO ABIERTO Y BARRENO DESPLAZADO EN EL PROCESO DE DOBLADO EN LA MÁQUINA D-229.",
    "causaId": "2",
    "causaTexto": "2. DISPOSITIVO EN DOBLADO ESTA RESTRINGIDO Y RECHAZA PIEZAS FUNCIONALES",
    "tipoAccion": "Correctiva",
    "accionId": "2",
    "accionTexto": "2. REALIZAR MODIFICACIÓN EN EL DISPOSITIVO DE DOBLADO (GESTIONAR COTIZACIÓN) DAR A LOS BLOQUES .025 DE TOLERANCIA EN AMBOS BLOQUES. MIGUEL/GRECIA (VIERNES 30/10/26)",
    "responsable": "MIGUEL ALCANTARILLA\nGRECIA ESCUDERO",
    "fechaCompromiso": "COTIZACIÓN 30/10/2026",
    "estado": "En Proceso",
    "cumplimiento": 0.2,
    "notes": ""
  },
  {
    "item": 1,
    "categoria": "Modo de Falla",
    "numeroParte": "147-5190",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META 3.33 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON AGULO ABIERTO Y BARRENO DESPLAZADO EN EL PROCESO DE DOBLADO EN LA MÁQUINA D-229.",
    "causaId": "3",
    "causaTexto": "3.OPERADOR NO LLEVA  A TOPE LA PIEZA PROVOCANDO DOBLEZ INCORRECTO",
    "tipoAccion": "Correctiva",
    "accionId": "3",
    "accionTexto": "3.VALIDAR POKA-YOKE EN EL CAÑÓN PARA ASEGURAR QUE LA PIEZA LLEGUE A TOPE PREVIO AL DOBLADO (MIGUEL SANCHEZ VALIDAR PROXIMA CORRIDA DE NP).",
    "responsable": "MIGUEL ALCANTARILLA\nMIGUEL SANCHEZ",
    "fechaCompromiso": "PENDIENTE FECHA",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 1,
    "categoria": "Modo de Falla",
    "numeroParte": "147-5190",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META 3.33 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON AGULO ABIERTO Y BARRENO DESPLAZADO EN EL PROCESO DE DOBLADO EN LA MÁQUINA D-229.",
    "causaId": "3",
    "causaTexto": "3.OPERADOR NO LLEVA  A TOPE LA PIEZA PROVOCANDO DOBLEZ INCORRECTO",
    "tipoAccion": "Correctiva",
    "accionId": "3.1",
    "accionTexto": "3.1 IDENTIFICAR PERSONAL CAPACITADO Y DEFINIR PLANTILLA FIJA PARA ESTE NÚMERO DE PARTE-SE PERDIO PLANTILLA POR ROTACIÓN, ACERCAMIENTO CON PERSONAL PARA FIJAR PLATILLA (26/10/26), YA ESTA LA PLANTILLA FIJA.",
    "responsable": "MIGUEL ALCANTARILLA",
    "fechaCompromiso": "2026-10-26",
    "estado": "Cerrada",
    "cumplimiento": 1.0,
    "notes": ""
  },
  {
    "item": 2,
    "categoria": "Modo de Falla",
    "numeroParte": "5092464",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META 1.59 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON BARRENO DESCENTRADO EN CL-007",
    "causaId": "1",
    "causaTexto": "1. MATERIA PRIMA HA PRESENTADO TORSIÓN Y PANDEAMIENTO, SE INSPECCIONA Y SE CORTAN BARRAS CON TOLERANCIA",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "1.SE INSPECCIONA MATERIA PRIMA, SE SELECCIONA BARRAS CON TOLERANCIA PARA CORTE, SE REALIZAN AJUSTE DE COMPENSACIÓN PARA ABSORBER PANDEAMIENTO",
    "responsable": "ISAI PILOTZI",
    "fechaCompromiso": "-",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 2,
    "categoria": "Modo de Falla",
    "numeroParte": "5092464",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META 1.59 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON BARRENO DESCENTRADO EN CL-007",
    "causaId": "1",
    "causaTexto": "1. MATERIA PRIMA HA PRESENTADO TORSIÓN Y PANDEAMIENTO, SE INSPECCIONA Y SE CORTAN BARRAS CON TOLERANCIA",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1. SE SOLICITO AUTORIZACIÓN DE CAMBIO DE PROVEEDOR DE MP PARA ASEGURAR \"TORSIÓN Y RECTITUD\" DENTRO DE ESPECIFICACIÓN, NO SE AUTORIZO.",
    "responsable": "ISAI PILOTZI",
    "fechaCompromiso": "-",
    "estado": "Cerrada",
    "cumplimiento": 1.0,
    "notes": ""
  },
  {
    "item": 2,
    "categoria": "Modo de Falla",
    "numeroParte": "5092464",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META 1.59 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON BARRENO DESCENTRADO EN CL-007",
    "causaId": "1",
    "causaTexto": "1. MATERIA PRIMA HA PRESENTADO TORSIÓN Y PANDEAMIENTO, SE INSPECCIONA Y SE CORTAN BARRAS CON TOLERANCIA",
    "tipoAccion": "Correctiva",
    "accionId": "2",
    "accionTexto": "2. DEJAR PROGRAMA A AJUSTADORES PARA REALIZAR COMPENSACIONES DE ACUERDO A CONDICIONES DE MATERIA PRIMA\n2.1 CAPACITAR A LOS AJUSTADORES EN COMPENSACIONES AMBAS NAVES",
    "responsable": "MIGUEL SANCHEZ",
    "fechaCompromiso": "2026-10-30",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 2,
    "categoria": "Modo de Falla",
    "numeroParte": "5092464",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META 1.59 ACUMULADO DE ENERO 2026 AL CORTE 17 OCTUBRE 2027, LOS DEFECTOS QUE SE HAN GENERADO SON BARRENO DESCENTRADO EN CL-007",
    "causaId": "1",
    "causaTexto": "1. MATERIA PRIMA HA PRESENTADO TORSIÓN Y PANDEAMIENTO, SE INSPECCIONA Y SE CORTAN BARRAS CON TOLERANCIA",
    "tipoAccion": "Correctiva",
    "accionId": "3",
    "accionTexto": "3. ACTUALIZAR PLAN DE DESARROLLO DE AJUSTADORES ANA FERMAN, AL CERRAR CAPACITACIONES 04/11/2026.",
    "responsable": "ANA FERMAN",
    "fechaCompromiso": "PENDIENTE",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 3,
    "categoria": "Modo de Falla",
    "numeroParte": "BEH91595",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META",
    "causaId": "1",
    "causaTexto": "1. EL SISTEMA DE BUSQUEDA DE SOLDADURA SE ENCUENTRA DESHABILITADO POR FALLAS EN MÁQUINA CL-009, EL OPERADOR TIENE QUE COLOCAR LA SOLDADURA MANUALMENTE",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "1. AYUDA VISUAL PARA OPERADOR COMO COLOCAR POSICIÓN DE SOLDADURA",
    "responsable": "ROGELIO RODRIGUEZ",
    "fechaCompromiso": "2026-06-01",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 3,
    "categoria": "Modo de Falla",
    "numeroParte": "BEH91595",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META",
    "causaId": "1",
    "causaTexto": "1. EL SISTEMA DE BUSQUEDA DE SOLDADURA SE ENCUENTRA DESHABILITADO POR FALLAS EN MÁQUINA CL-009, EL OPERADOR TIENE QUE COLOCAR LA SOLDADURA MANUALMENTE",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1. SOLICITAR  SOPORTE TÉCNICO CON PROVEEDOR BLM PARA DIAGNÓSTICO DE LA FALLA.\n2. SE REALIZA EL REEMPLAZO DEL CABLE HMI.\n3. SE VALIDA ESTABILIDAD OPERATIVA SIN RECURRENCIA DE LA FALLA.",
    "responsable": "ROGELIO RODRIGUEZ",
    "fechaCompromiso": "2026-06-01",
    "estado": "Cerrada",
    "cumplimiento": 1.0,
    "notes": ""
  },
  {
    "item": 4,
    "categoria": "Modo de Falla",
    "numeroParte": "UC41059",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META GENERADO PRINCIPALMENTE EN DOBLADORA 233 Y MB-001",
    "causaId": "1",
    "causaTexto": "1. LAS PIEZAS AL ACOMODARSE SE GOLPEAN GENERANDO MARCA PROFUNDA",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1.SEPARADOR ENTRE CAMAS DE PIEZAS (REPLICAR EN TRASPALEO DE PIEZAS)",
    "responsable": "MIGUEL ALCANTARILLA",
    "fechaCompromiso": "2026-11-04",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 4,
    "categoria": "Modo de Falla",
    "numeroParte": "UC41059",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META GENERADO PRINCIPALMENTE EN DOBLADORA 233 Y MB-001",
    "causaId": "2",
    "causaTexto": "2. OPERADOR INTRODUCE LA PIEZA EN POSICIÓN INCORRECTA QUEDANDO PIEZA MAL DOBLADA",
    "tipoAccion": "Correctiva",
    "accionId": "2",
    "accionTexto": "2.VALIDAR AGREGAR PRECORTE EN PIEZA PARA IDENTIFICAR POSICIÓN DE PIEZAS, PENDIENTE VALIDAR CON MIGUEL SANCHEZ, PENDIENTE POR VALIDAR PRODUCCION FLUJO DE PROCESO PARA CORREGIR ÁREA DE OPORTUNIDAD DE OPERADORES.",
    "responsable": "MIGUEL ALCANTARILLA\nMIGUEL SANCHEZ",
    "fechaCompromiso": "PENDIENTE REVISION",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 4,
    "categoria": "Modo de Falla",
    "numeroParte": "UC41059",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META GENERADO PRINCIPALMENTE EN DOBLADORA 233 Y MB-001",
    "causaId": "3",
    "causaTexto": "3. BARRENO DESCENTRADO DESDE CORTE LASER, GAGE COMPRA PIEZAS CON BARRENO DESCENTRADO (MATERIA PRIMA EN LA MAXIMA REDUCE TOLERANCIA DE CENTRADO, EN LA MINIMA TE DA MAYOR TOLERANCIA CON BARRENOS DESCENTRADOS), BARRENOS NO ESTAN CONCENTRICOS.",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "3. IMPLEMENTAR  GAGE CON PERNO PASADO EN EL CORTE LASER (MIGUEL ALCANTARILLA/GRECIA ESCUDERO)-15/11/2026\nDISEÑO DE GAGE PARA 28/10/26, AGREGAR A LA DOCUMENTACION TECNICA",
    "responsable": "GRECIA ESCUDERO\nMIGUEL ALCANTARILLA",
    "fechaCompromiso": "PENDIENTE",
    "estado": "En Proceso",
    "cumplimiento": 0.5,
    "notes": ""
  },
  {
    "item": 4,
    "categoria": "Modo de Falla",
    "numeroParte": "UC41059",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META GENERADO PRINCIPALMENTE EN DOBLADORA 233 Y MB-001",
    "causaId": "3",
    "causaTexto": "3. BARRENO DESCENTRADO DESDE CORTE LASER, GAGE COMPRA PIEZAS CON BARRENO DESCENTRADO (MATERIA PRIMA EN LA MAXIMA REDUCE TOLERANCIA DE CENTRADO, EN LA MINIMA TE DA MAYOR TOLERANCIA CON BARRENOS DESCENTRADOS), BARRENOS NO ESTAN CONCENTRICOS.",
    "tipoAccion": "Correctiva",
    "accionId": "3",
    "accionTexto": "3. SE SELECCIONO ATADOS CON BARRAS FUERA DE ESPECIFICACION PARA VALIDAR CON PROVEEDOR LAS CONDICIONES DE LA MATERIA PRIMA.",
    "responsable": "GRECIA ESCUDERO\nMIGUEL ALCANTARILLA",
    "fechaCompromiso": "PENDIENTE",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 4,
    "categoria": "Modo de Falla",
    "numeroParte": "UC41059",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META GENERADO PRINCIPALMENTE EN DOBLADORA 233 Y MB-001",
    "causaId": "4",
    "causaTexto": "4. DISTANCIA DE WIPER MENOR A 43 MM NO ASIENTA CORRECTAMENTE EN LA RODAJA GENERANDO SCRAP .",
    "tipoAccion": "Correctiva",
    "accionId": "4",
    "accionTexto": "4. CABIOS DE WIPER E INSERTO RESPECTO A VIDA UTIL (INMEDIATO) CHRISTIAN, ESTAR MONITOREANDO EN AUTONOMOS, GRADUADO DE WIPER, PORTAINSERTO (), SE FIJA VIDA UTIL WIPER PORTA INSERTO, FALTA DEFINIR CADA CUANTO, PRIMER PASO QUE SE CAPTURE MANUAL EN LA BITACORA (INMEDIATO) BITACORAS EN PISO, LOS VIERNES CHRISTIAN AUDITA QUE SE ESTEN REALIZANDO LOS REPORTES.",
    "responsable": "CHRISTIAN",
    "fechaCompromiso": "INMEDIATO",
    "estado": "Cerrada",
    "cumplimiento": 1.0,
    "notes": ""
  },
  {
    "item": 5,
    "categoria": "Modo de Falla",
    "numeroParte": "749-06174C\n749-06175C",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META GENERADO PRINCIPALMENTE EN DOBADLORA D-232",
    "causaId": "1",
    "causaTexto": "1. MATERIA PRIMA VIENE CON COSTURA CORRIDA Y CHISPORROTEO GENERA QUE EL MANDRIL GENERE ABOMBAMIENTO EN EL DOBLEZ",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "1.1PARA MATERIA PRIMA PROCESADA EN MOLINO 4, REALIZAR LIBERACIÓN CON EL EQUIPO DE TEFLU (PRODUCCIÓN, CALIDAD, INGENIERÍA, GERENCIA, ETC.), DOCUMENTANDO LA CONDICIÓN BAJO DESVIACIÓN.",
    "responsable": "MIGUEL ALCANTARILLA\nGRECIA ESCUDERO\nISAI PILOTZI",
    "fechaCompromiso": "2026-04-30",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 5,
    "categoria": "Modo de Falla",
    "numeroParte": "749-06174C\n749-06175C",
    "problema": "PORCENTAJE DE SCRAP FUERA DE META GENERADO PRINCIPALMENTE EN DOBADLORA D-232",
    "causaId": "1",
    "causaTexto": "1. MATERIA PRIMA VIENE CON COSTURA CORRIDA Y CHISPORROTEO GENERA QUE EL MANDRIL GENERE ABOMBAMIENTO EN EL DOBLEZ",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1. DEFINIR DIMENSIONES FUNCIONALES DE ESLABONES  Y COMPARTIR DIBUJO A CALIDAD PARA FABRICAR GAGE 1. 03/04/2026\n1.2MANDAR A FABRICAR GAGE PARA CALIDAD Y PRODUCCIÓN 1. 03/04/2026\n1.3 LIBERAR LA MATERIA PRIMA EN RECIBO MEDIANTE GAGE  (ESLABONES CON VARILLA) ASEGURANDO QUE LA MATERIA PRIMA DE MOLINO 6 SE ENCUENTRE LIBRE DE CHISPORROTEO; CUALQUIER MATERIAL CON ESTA CONDICIÓN DEBERÁ SER RECHAZADO 18/04/26\n1.4 VALIDAR EN PROCESO DE CORTE LÁSER LA CONDICIÓN DE COSTURA MEDIANTE GAGE (ESLABONES CON VARILLA)  18/04/26.",
    "responsable": "MIGUEL ALCANTARILLA\nGRECIA ESCUDERO\nISAI PILOTZI",
    "fechaCompromiso": "2026-04-30",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 6,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CORTADORA DISCO C-123",
    "problema": "ALARMAS CONSTANTES EN MAQUINA LA CAUSA RAIZ NO  ESTA DETECTADA",
    "causaId": "1",
    "causaTexto": "CAUSA RAIZ NO ENCONTRADA AUN",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "1. SE REVISA TARJETA PROFIBUS, SE MANDA A REPARACION Y NO SE ENCUETRA LA FALLA\n2. SE SOLICITA TARJETA NUEVA CON TIEMPO DE ENTREGA DE 15 DIAS\n3. TARJETA SE ENTREGO EL 21/10/26 SE INSTALO, SE SIGUE PRESENTANDO LA FALLA\n4. SE RELEVAN ENTRADAS Y SALIDAS DE TARJETA\n5. SE RELEVAN BOBINAS Y SE PONE UNA FUENTE ALTERNA PARA DISMINUIR CONSUMO DE ENERGIA DIRECTA DE LA TARJETA \n6. SE SOLICITA SOPORTE DE PROVEEDOR, AUN NO SE CONFIRMA FECHA DE VISITA.",
    "responsable": "VICTOR HERNANDEZ",
    "fechaCompromiso": "INMEDIATO",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 7,
    "categoria": "Modo de Falla",
    "numeroParte": "D-232\n749-06114",
    "problema": "VARILLA MODIFICADA NO TIENE SUJECION CORRECTA EN EL BLOQUE DE FIJACION DE LA VARILLA POR LO QUE ESTA GENERANDO TIEMPOS PROLONGADOS DE AJUSTE Y SCRAP",
    "causaId": "1",
    "causaTexto": "1.ESTA TENIENDO JUEGO EN LA PARTE DEL BLOQUE DE FIJACION DE LA VARILLA PORQUE EL DISEÑO QUE SE TIENE AL MOMENTO DE SOLDAR LAS ARANDELAS NO QUEDAN A LA MISMA DISTANCIA DE SEPARACIÓN, AL REALIZAR DOBLEZ SE MUEVE Y GENERA QUE LA PIEZA TENGA ANGULO ABIERTO/BARRENO DESPLAZADO",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "1. ARMAR VARILLA CON CONDICIONES ADECUADAS PARA EVITAR DESPLAZAMIENTO 29/10/2026,",
    "responsable": "CHRISTIAN",
    "fechaCompromiso": "",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 7,
    "categoria": "Modo de Falla",
    "numeroParte": "D-232\n749-06114",
    "problema": "VARILLA MODIFICADA NO TIENE SUJECION CORRECTA EN EL BLOQUE DE FIJACION DE LA VARILLA POR LO QUE ESTA GENERANDO TIEMPOS PROLONGADOS DE AJUSTE Y SCRAP",
    "causaId": "1",
    "causaTexto": "1.ESTA TENIENDO JUEGO EN LA PARTE DEL BLOQUE DE FIJACION DE LA VARILLA PORQUE EL DISEÑO QUE SE TIENE AL MOMENTO DE SOLDAR LAS ARANDELAS NO QUEDAN A LA MISMA DISTANCIA DE SEPARACIÓN, AL REALIZAR DOBLEZ SE MUEVE Y GENERA QUE LA PIEZA TENGA ANGULO ABIERTO/BARRENO DESPLAZADO",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1.1 REALIZAR DISEÑO DE DISPOSITIVO PARA SOLDADURA\n1.1 FABRICAR DISPOSITIVO DE SOLDADURA\n1.2 GENERAR INSTRUCCIÓN DE TRABAJO PARA USO DE DISPOSITIVO Y CAPACITACION DEL SOLDADOR",
    "responsable": "CHRISTIAN",
    "fechaCompromiso": "",
    "estado": "Cerrada",
    "cumplimiento": 1.0,
    "notes": ""
  },
  {
    "item": 8,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "D-232",
    "problema": "PLANTILLA INCOMPLETA",
    "causaId": "1",
    "causaTexto": "NO SE CUENTA CON AYUDANTE GENERAL CONTRATADO EN EL TURNO DEL SUPERVISOR  BLADIMIR Y JOEL, EL PERSONAL CONTRATADO RENUNCIO POR TIPO DE TRABAJO, NO SE HAN CUBIERTO LAS VACANTES POR FALTA DE RECLUTADOR",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "1. YA SE CONTRATARON A LOS DOS AYUDANTES Y YA SE TIENE CUBIERTO ACTUALMENTE LAS VACANTES",
    "responsable": "ANA FERMAN",
    "fechaCompromiso": "PENDIENTE",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 8,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "D-232",
    "problema": "PLANTILLA INCOMPLETA",
    "causaId": "1",
    "causaTexto": "NO SE CUENTA CON AYUDANTE GENERAL CONTRATADO EN EL TURNO DEL SUPERVISOR  BLADIMIR Y JOEL, EL PERSONAL CONTRATADO RENUNCIO POR TIPO DE TRABAJO, NO SE HAN CUBIERTO LAS VACANTES POR FALTA DE RECLUTADOR",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "ANALIZAR ADAPATACION DE PERSONAL DE NUEVO INGRESO EN DOBLADO PARA DETERMINAR TIEMPO DE LIBERACION AL 100% A OPERACIONES\nSOLICITAR METODOS DE TRABAJO A SEGURIDAD INDUSTRIAL PARA DISMINUIR CARGA DE TRABAJO",
    "responsable": "ANA FERMAN",
    "fechaCompromiso": "PENDIENTE",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 9,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "VARIAS MÁQUINAS",
    "problema": "TRANSPORTE SE DEMORA EN LLEGADA A PLANTA POR TRÁFICO GENERANDO RETRASO EN OPERACIONES",
    "causaId": "1",
    "causaTexto": "1. TRÁFICO EN RUTA HABITUAL GENERA DEMORA EN LLEGADA DE PERSONAL A PLANTA",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1. REVISAR ANÁLISIS DE IMPACTO DE RUTA Y EVALUAR AJUSTE DE HORARIOS / RECORRIDO",
    "responsable": "RECURSOS HUMANOS\nLOGÍSTICA",
    "fechaCompromiso": "PENDIENTE",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 9,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "VARIAS MÁQUINAS",
    "problema": "TRANSPORTE SE DEMORA EN LLEGADA A PLANTA POR TRÁFICO GENERANDO RETRASO EN OPERACIONES",
    "causaId": "2",
    "causaTexto": "2. FALTA DE COORDINACIÓN ENTRE NECESIDADES PARA TEFLU Y MAQUILA",
    "tipoAccion": "Correctiva",
    "accionId": "2",
    "accionTexto": "2. COORDINACIÓN ENTRE NECESIDADES PARA TEFLU Y MAQUILA (ESTABLECER CANAL DIRECTO DE MONITOREO)",
    "responsable": "PLANEACIÓN\nLOGÍSTICA",
    "fechaCompromiso": "PENDIENTE",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 10,
    "categoria": "Modo de Falla",
    "numeroParte": "BEWO  80142 5EF0A",
    "problema": "DIMENSIÓN FUERA DE ESPECIFICACIÓN: PIEZAS LARGAS SE PUEDEN RETRABAJAR, PERO SE TIRARON",
    "causaId": "1",
    "causaTexto": "1. DESCONOCIMIENTO EN PISO DEL PROCEDIMIENTO DE RECUPERACIÓN DE PIEZAS LARGAS; DISPOSITIVO EN EMPAQUE E INSPECCIÓN CON PIEZA MASTER GENERANDO FUGA DE MATERIAL",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "1. RETIRAR DISPOSITIVO DE EMPAQUE Y PONER EN PISO PARA INSPECCIÓN; SEGREGAR PIEZAS LARGAS PARA EVITAR SCRAP",
    "responsable": "CALIDAD\nPRODUCCIÓN",
    "fechaCompromiso": "INMEDIATO",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 10,
    "categoria": "Modo de Falla",
    "numeroParte": "BEWO  80142 5EF0A",
    "problema": "DIMENSIÓN FUERA DE ESPECIFICACIÓN: PIEZAS LARGAS SE PUEDEN RETRABAJAR, PERO SE TIRARON",
    "causaId": "1",
    "causaTexto": "1. DESCONOCIMIENTO EN PISO DEL PROCEDIMIENTO DE RECUPERACIÓN DE PIEZAS LARGAS; DISPOSITIVO EN EMPAQUE E INSPECCIÓN CON PIEZA MASTER GENERANDO FUGA DE MATERIAL",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1. CAPACITACIÓN A PERSONAL SOBRE CRITERIO DE RECUPERACIÓN Y RETRABAJO DE PIEZAS LARGAS CON GAGE OFICIAL",
    "responsable": "CALIDAD\nPRODUCCIÓN",
    "fechaCompromiso": "INMEDIATO",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 10,
    "categoria": "Modo de Falla",
    "numeroParte": "BEWO  80142 5EF0A",
    "problema": "DIMENSIÓN FUERA DE ESPECIFICACIÓN: PIEZAS LARGAS SE PUEDEN RETRABAJAR, PERO SE TIRARON",
    "causaId": "2",
    "causaTexto": "2. NÚMEROS DE PARTE DE LUNKOMEX QUE SE CORTAN EN LÁSER: LONGITUD NO COMPATIBLE DIRECTAMENTE CON CORTE LÁSER",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "2. SOLICITAR A PROGRAMACIÓN LÁSER PLAN DE CORTE EMERGENTE",
    "responsable": "PROGRAMACIÓN\nCORTE LÁSER",
    "fechaCompromiso": "INMEDIATO",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 10,
    "categoria": "Modo de Falla",
    "numeroParte": "BEWO  80142 5EF0A",
    "problema": "DIMENSIÓN FUERA DE ESPECIFICACIÓN: PIEZAS LARGAS SE PUEDEN RETRABAJAR, PERO SE TIRARON",
    "causaId": "2",
    "causaTexto": "2. NÚMEROS DE PARTE DE LUNKOMEX QUE SE CORTAN EN LÁSER: LONGITUD NO COMPATIBLE DIRECTAMENTE CON CORTE LÁSER",
    "tipoAccion": "Correctiva",
    "accionId": "2",
    "accionTexto": "2. SOLICITAR A PROGRAMACIÓN DESPUNTAR TUBO PARA PODER METERLO EN LA LÁSER Y REALIZAR PRUEBA PILOTO",
    "responsable": "PROGRAMACIÓN\nCORTE LÁSER",
    "fechaCompromiso": "INMEDIATO",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 11,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CL-009",
    "problema": "3 VISITAS DE BLM, CALIBRAN SISTEMA DE VISIÓN TRABAJA BIEN 2 SEMANAS Y SE VUELVE A PRESENTAR FALLA, SE CAMBIARON LENTES, CABLES, PIERDE CALIBRACIÓN, SE PRESENTA 4 VISITA",
    "causaId": "1",
    "causaTexto": "NO SE TIENE CAUSA RAÍZ DEFINITIVA ACTUALMENTE (SE CAMBIARON CABLES, LENTES Y SWITCH SIN SOLUCIÓN DEFINITIVA, PROCESO DE CALIBRACIÓN TARDADO)",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "1. CONTENCIÓN: SE REINICIA LA MÁQUINA CADA VEZ QUE SE PRESENTA LA FALLA\n2. CALIBRACIÓN TEMPORAL POR PROVEEDOR BLM",
    "responsable": "ROGELIO RODRIGUEZ\nBLM",
    "fechaCompromiso": "PENDIENTE VISITA",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 11,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CL-009",
    "problema": "3 VISITAS DE BLM, CALIBRAN SISTEMA DE VISIÓN TRABAJA BIEN 2 SEMANAS Y SE VUELVE A PRESENTAR FALLA, SE CAMBIARON LENTES, CABLES, PIERDE CALIBRACIÓN, SE PRESENTA 4 VISITA",
    "causaId": "1",
    "causaTexto": "NO SE TIENE CAUSA RAÍZ DEFINITIVA ACTUALMENTE (SE CAMBIARON CABLES, LENTES Y SWITCH SIN SOLUCIÓN DEFINITIVA, PROCESO DE CALIBRACIÓN TARDADO)",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1. SOLICITAR 4TA VISITA DE ESPECIALISTA BLM PARA DIAGNÓSTICO PROFUNDO EN HARDWARE/SOFTWARE Y SEGUIMIENTO A GARANTÍA",
    "responsable": "ROGELIO RODRIGUEZ\nBLM",
    "fechaCompromiso": "PENDIENTE VISITA",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 12,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CL-005",
    "problema": "FALLA RECURRENTE DE CENTRADORES ALTOS EN CL-005",
    "causaId": "1",
    "causaTexto": "PIEZAS DESGASTADAS NYLAMIND (REFACCIÓN) Y MAL AJUSTE DE SENSORES (AGREGAR A LA INSTRUCCIÓN DE PREVENTIVO, CAMBIO CADA AÑO DE LAS PIEZAS), REBABAS EN CREMALLERAS.",
    "tipoAccion": "Contención",
    "accionId": "1",
    "accionTexto": "REPARACIÓN Y AJUSTE MECÁNICO TEMPORAL EN MÁQUINA",
    "responsable": "ROGELIO RODRIGUEZ",
    "fechaCompromiso": "2026-10-30",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 12,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CL-005",
    "problema": "FALLA RECURRENTE DE CENTRADORES ALTOS EN CL-005",
    "causaId": "1",
    "causaTexto": "PIEZAS DESGASTADAS NYLAMIND (REFACCIÓN) Y MAL AJUSTE DE SENSORES (AGREGAR A LA INSTRUCCIÓN DE PREVENTIVO, CAMBIO CADA AÑO DE LAS PIEZAS), REBABAS EN CREMALLERAS.",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1. DAR DE ALTA REFACCIONES 30/10/2026 (GESTIONAR LIBERACIÓN DE DIRECCIÓN)",
    "responsable": "ROGELIO RODRIGUEZ",
    "fechaCompromiso": "2026-10-30",
    "estado": "En Proceso",
    "cumplimiento": 0.5,
    "notes": ""
  },
  {
    "item": 12,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CL-005",
    "problema": "FALLA RECURRENTE DE CENTRADORES ALTOS EN CL-005",
    "causaId": "1",
    "causaTexto": "PIEZAS DESGASTADAS NYLAMIND (REFACCIÓN) Y MAL AJUSTE DE SENSORES (AGREGAR A LA INSTRUCCIÓN DE PREVENTIVO, CAMBIO CADA AÑO DE LAS PIEZAS), REBABAS EN CREMALLERAS.",
    "tipoAccion": "Correctiva",
    "accionId": "2",
    "accionTexto": "2. INSTRUCCIÓN DE TRABAJO PARA INSTALACIÓN DE CENTRADORES (13/11/26) Y EVALUACIÓN DE TÉCNICOS DE MANTENIMIENTO (20/11/26)",
    "responsable": "ROGELIO RODRIGUEZ",
    "fechaCompromiso": "2026-11-13",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 13,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CORTE LÁSER",
    "problema": "NO SE CUENTA CON PLANTILLA COMPLETA Y NO SE TIENE EL PERSONAL CALIFICADO (5 INTENTOS DE REPARACIÓN EN FALLA DE CENTRADORES)",
    "causaId": "1",
    "causaTexto": "TÉCNICOS NO ESTÁN FAMILIARIZADOS CON MÁQUINAS QUE SE TIENEN EN PLANTA; PERSONAL DE INGRESO TIENE CONOCIMIENTO PERO NO ESTÁ FAMILIARIZADO CON LAS MÁQUINAS DE PLANTA.",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1. PLANTILLA COMPLETA Y ACELERAMIENTO DE PERSONAL COMPETENTE (CONTRATAR PERSONAL Y BASE DE PLANTILLA PARA SUPLIR INMEDIATAMENTE)",
    "responsable": "RECURSOS HUMANOS\nMANTENIMIENTO",
    "fechaCompromiso": "PENDIENTE",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 13,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CORTE LÁSER",
    "problema": "NO SE CUENTA CON PLANTILLA COMPLETA Y NO SE TIENE EL PERSONAL CALIFICADO (5 INTENTOS DE REPARACIÓN EN FALLA DE CENTRADORES)",
    "causaId": "1",
    "causaTexto": "TÉCNICOS NO ESTÁN FAMILIARIZADOS CON MÁQUINAS QUE SE TIENEN EN PLANTA; PERSONAL DE INGRESO TIENE CONOCIMIENTO PERO NO ESTÁ FAMILIARIZADO CON LAS MÁQUINAS DE PLANTA.",
    "tipoAccion": "Correctiva",
    "accionId": "2",
    "accionTexto": "2. PLAN DUAL CON CONALEP PARA TRAER PRACTICANTES PARA FUTUROS TÉCNICOS",
    "responsable": "RECURSOS HUMANOS",
    "fechaCompromiso": "PENDIENTE",
    "estado": "Abierta",
    "cumplimiento": 0.0,
    "notes": ""
  },
  {
    "item": 14,
    "categoria": "Seguimiento Independiente",
    "numeroParte": "CL-007",
    "problema": "DISEÑO DE GARRA PARA TRAPECIO",
    "causaId": "1",
    "causaTexto": "ÁNGULO DE CONO ORIGINAL (45°) Y LONGITUD DE GARRA PROVOCABAN FALLAS DE SUJECIÓN EN PERFIL TRAPECIO",
    "tipoAccion": "Correctiva",
    "accionId": "1",
    "accionTexto": "1. SE REALIZÓ MODIFICACIÓN DE ANGULACIÓN DE CONO DE GARRAS, DE 45 GRADOS A 42 GRADOS, SE RECORTO 1 MM DE LONGITUD.",
    "responsable": "INGENIERÍA\nMANTENIMIENTO",
    "fechaCompromiso": "INMEDIATO",
    "estado": "Cerrada",
    "cumplimiento": 1.0,
    "notes": ""
  }
];

function readAcciones_(spreadsheet, config) {
  var sourceName = 'C. ACC';
  try {
    var sheet = spreadsheet.getSheetByName('C. ACC') || 
                spreadsheet.getSheetByName('Concentrado') || 
                spreadsheet.getSheetByName('CONCENTRADO') || 
                spreadsheet.getSheetByName('C.ACC') || 
                spreadsheet.getSheetByName('ACCIONES') || 
                spreadsheet.getSheetByName('SEGUIMIENTO DE ACCIONES') || 
                spreadsheet.getSheetByName('SEGUIMIENTO ACCIONES') ||
                spreadsheet.getSheetByName('PLAN DE ACCIONES');
    if (!sheet || sheet.getLastRow() < 2) {
      return { 
        rows: ACTIONS_SEED_, 
        activities: ACTIONS_ACTIVITIES_SEED_,
        status: { state: 'seed', source: 'ACCIONES 2026 FINAL.xlsx (C. ACC)', total: ACTIONS_SEED_.length, message: 'Mostrando datos de la hoja C. ACC.' } 
      };
    }
    sourceName = sheet.getName();
    var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    var result = normalizeAcciones_(values);
    result.status.source = sourceName;
    return result;
  } catch (error) {
    return { 
      rows: ACTIONS_SEED_, 
      activities: ACTIONS_ACTIVITIES_SEED_,
      status: { state: 'seed', source: 'ACCIONES 2026 FINAL.xlsx (C. ACC)', total: ACTIONS_SEED_.length, message: 'Usando datos base de C. ACC: ' + publicErrorMessage_(error) } 
    };
  }
}

function normalizeAcciones_(values) {
  if (!values || values.length < 2) {
    return { rows: ACTIONS_SEED_, activities: ACTIONS_ACTIVITIES_SEED_, status: { state: 'seed', source: 'C. ACC', total: ACTIONS_SEED_.length } };
  }

  // Find header row (can be row 0 or row 1 due to =TODAY() in row 1)
  var headerRowIdx = 0;
  for (var hR = 0; hR < Math.min(values.length, 4); hR++) {
    var rowStr = (values[hR] || []).map(function(c) { return cleanString_(c).toUpperCase(); }).join(' ');
    if (rowStr.indexOf('ITEM') >= 0 || rowStr.indexOf('NUMERO DE PARTE') >= 0 || rowStr.indexOf('ACCIONES CORRECTIVAS') >= 0) {
      headerRowIdx = hR;
      break;
    }
  }

  var rawHeaders = values[headerRowIdx] || [];
  var headers = rawHeaders.map(function(h) { return cleanString_(h).toUpperCase(); });

  function findCol(candidates) {
    for (var i = 0; i < candidates.length; i++) {
      for (var j = 0; j < headers.length; j++) {
        if (headers[j].indexOf(candidates[i].toUpperCase()) >= 0) return j;
      }
    }
    return -1;
  }

  var idxItem = findCol(['ITEM', 'NO', 'NUM']);
  var idxTipo = findCol(['TIPO DE ACCION', 'TIPO']);
  var idxParte = findCol(['NUMERO DE PARTE', 'PARTE', 'MAQUINA', 'EQUIPO']);
  var idxDesc = findCol(['DESCRIPCION DEL PROBLEMA', 'DESCRIPCION', 'PROBLEMA']);
  var idxCausa = findCol(['CAUSA RAIZ', 'CAUSA']);
  var idxCont = findCol(['ACCIONES CONTENEDORAS', 'CONTENEDORAS', 'CONTENCION']);
  var idxCorr = findCol(['ACCIONES CORRECTIVAS', 'CORRECTIVAS']);
  var idxResp = findCol(['RESPONSABLES', 'RESPONSABLE']);
  var idxFechaComp = findCol(['FECHA COMPROMISO', 'COMPROMISO', 'FECHA FINAL', 'FINAL']);
  var idxCump = findCol(['CUMPLIMIENTO', 'AVANCE', '%']);

  var lastItem = '', lastTipo = '', lastParte = '', lastDesc = '', lastCausa = '', lastCont = '';
  var itemGroups = {};
  var activitiesList = [];

  for (var r = headerRowIdx + 1; r < values.length; r++) {
    var row = values[r];
    if (!row || !row.some(function(c) { return c !== '' && c !== null && c !== undefined; })) continue;

    var itemVal = idxItem >= 0 && row[idxItem] !== '' && row[idxItem] !== null ? row[idxItem] : lastItem;
    var tipoVal = idxTipo >= 0 && row[idxTipo] !== '' && row[idxTipo] !== null ? cleanString_(row[idxTipo]) : lastTipo;
    var parteVal = idxParte >= 0 && row[idxParte] !== '' && row[idxParte] !== null ? cleanString_(row[idxParte]) : lastParte;
    var descVal = idxDesc >= 0 && row[idxDesc] !== '' && row[idxDesc] !== null ? cleanString_(row[idxDesc]) : lastDesc;
    var causaVal = idxCausa >= 0 && row[idxCausa] !== '' && row[idxCausa] !== null ? cleanString_(row[idxCausa]) : lastCausa;
    var contVal = idxCont >= 0 && row[idxCont] !== '' && row[idxCont] !== null ? cleanString_(row[idxCont]) : lastCont;

    lastItem = itemVal;
    lastTipo = tipoVal;
    lastParte = parteVal;
    lastDesc = descVal;
    lastCausa = causaVal;
    lastCont = contVal;

    var corrVal = idxCorr >= 0 ? cleanString_(row[idxCorr]) : '';
    var respVal = idxResp >= 0 ? cleanString_(row[idxResp]) : 'Por asignar';
    var fCompVal = idxFechaComp >= 0 ? formatActionDate_(row[idxFechaComp]) : '-';

    var cump = 0;
    var rawCump = idxCump >= 0 ? row[idxCump] : 0;
    if (typeof rawCump === 'number') {
      cump = rawCump;
    } else if (rawCump) {
      var num = parseFloat(String(rawCump).replace('%', '').replace(',', '.').trim());
      if (!isNaN(num)) cump = num > 1 ? num / 100 : num;
    }

    var estado = 'Abierta';
    if (cump >= 1) estado = 'Cerrada';
    else if (cump > 0) estado = 'En Proceso';

    var isModoFalla = tipoVal.indexOf('SCRAP') >= 0 || tipoVal.indexOf('DEFECTO') >= 0 || (parteVal !== 'CORTADORA DISCO C-123' && tipoVal.indexOf('DISPONIBILIDAD') < 0 && Number(itemVal) <= 5);
    var categoria = isModoFalla ? 'Modo de Falla' : 'Seguimiento Independiente';

    var itemKey = String(itemVal || r);
    if (!itemGroups[itemKey]) {
      itemGroups[itemKey] = {
        item: Number(itemVal) || r,
        tipo: tipoVal || 'ACCION',
        categoria: categoria,
        numeroParte: parteVal || '-',
        descripcion: descVal || '',
        causasMap: {},
        responsablesSet: {},
        totalAcciones: 0,
        accionesCerradas: 0,
        sumCumplimiento: 0,
        fechaFinal: '-'
      };
    }

    var g = itemGroups[itemKey];
    g.totalAcciones++;
    if (estado === 'Cerrada') g.accionesCerradas++;
    g.sumCumplimiento += cump;
    if (fCompVal && fCompVal !== '-') g.fechaFinal = fCompVal;
    if (respVal && respVal !== 'Por asignar') {
      var rParts = respVal.split(/\r?\n|,/);
      for (var rp = 0; rp < rParts.length; rp++) {
        var rpTrim = rParts[rp].trim();
        if (rpTrim) g.responsablesSet[rpTrim] = true;
      }
    }

    var causaKey = causaVal || 'Causa general';
    if (!g.causasMap[causaKey]) {
      g.causasMap[causaKey] = {
        causeId: String(Object.keys(g.causasMap).length + 1),
        causeText: causaKey,
        containments: [],
        correctives: []
      };
    }

    var cTarget = g.causasMap[causaKey];
    if (contVal && contVal !== '-' && cTarget.containments.length === 0) {
      cTarget.containments.push({
        actionId: '1',
        type: 'Contención',
        actionText: contVal,
        date: fCompVal,
        responsible: respVal,
        status: contVal.toUpperCase().indexOf('IMPLEMENTA') >= 0 ? 'En Proceso' : 'Abierta',
        compliance: contVal.toUpperCase().indexOf('IMPLEMENTA') >= 0 ? 0.5 : 0.0,
        notes: ''
      });
    }

    if (corrVal && corrVal !== '-') {
      var mNum = corrVal.match(/^(\d+(?:\.\d+)*)\s*[\.\-\)]?\s*(.*)$/);
      var actId = mNum ? mNum[1] : String(cTarget.correctives.length + 1);
      cTarget.correctives.push({
        actionId: actId,
        type: 'Correctiva',
        actionText: corrVal,
        date: fCompVal,
        responsible: respVal,
        status: estado,
        compliance: cump,
        notes: ''
      });

      activitiesList.push({
        item: g.item,
        categoria: g.categoria,
        numeroParte: g.numeroParte,
        problema: g.descripcion,
        causaId: cTarget.causeId,
        causaTexto: cTarget.causeText,
        tipoAccion: 'Correctiva',
        accionId: actId,
        accionTexto: corrVal,
        responsable: respVal,
        fechaCompromiso: fCompVal,
        estado: estado,
        cumplimiento: cump,
        notes: ''
      });
    }
  }

  var rows = [];
  for (var k in itemGroups) {
    var grp = itemGroups[k];
    var cArr = [];
    for (var cK in grp.causasMap) {
      cArr.push(grp.causasMap[cK]);
    }
    var respsList = Object.keys(grp.responsablesSet).join(', ') || 'Por asignar';
    var avgCump = grp.totalAcciones > 0 ? (grp.sumCumplimiento / grp.totalAcciones) : 0;
    rows.push({
      item: grp.item,
      tipo: grp.tipo,
      categoria: grp.categoria,
      numeroParte: grp.numeroParte,
      descripcion: grp.descripcion,
      causas: cArr,
      fechaInicial: '-',
      fechaFinal: grp.fechaFinal,
      responsables: respsList,
      cumplimiento: avgCump,
      totalAcciones: grp.totalAcciones
    });
  }

  return {
    rows: rows.length > 0 ? rows : ACTIONS_SEED_,
    activities: activitiesList.length > 0 ? activitiesList : ACTIONS_ACTIVITIES_SEED_,
    status: { state: 'ok', source: 'C. ACC', total: rows.length }
  };
}

function formatActionDate_(val) {
  if (!val || val === '-') return '-';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'America/Mexico_City', 'yyyy-MM-dd');
  }
  return cleanString_(val);
}

function crearOActualizarHojaAcciones() {
  var config = runtimeConfig_();
  var ss = SpreadsheetApp.openById(config.spreadsheetId);
  var sheetName = 'C. ACC';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  var headers = [
    'ITEM', 'TIPO DE ACCION', 'NUMERO DE PARTE', 'DESCRIPCION DEL PROBLEMA',
    'CAUSA RAIZ', 'ACCIONES CONTENEDORAS', 'ACCIONES CORRECTIVAS',
    'RESPONSABLES', 'FECHA COMPROMISO', 'CUMPLIMIENTO'
  ];
  
  sheet.getRange(1, 9).setFormula('=TODAY()');
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  var headerRange = sheet.getRange(2, 1, 1, headers.length);
  headerRange.setBackground('#0f1c2e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontFamily('Arial');
  headerRange.setFontSize(10);
  headerRange.setHorizontalAlignment('center');
  
  sheet.setFrozenRows(2);
  for (var c = 1; c <= headers.length; c++) {
    sheet.autoResizeColumn(c);
  }

  return { success: true, message: 'Hoja C. ACC lista en Spreadsheet.' };
}