/**
 * ============================================================================
 * BACKEND API - SISTEMA ÁGIL DE ENTRADAS, SALIDAS, ALISTAMIENTO Y ETIQUETAS
 * Proyecto: movimientos-eys-z (Zúñiga DecoCancún)
 * ============================================================================
 */

const EYS_CONFIG = {
  TIMEZONE: "America/Cancun",
  IDS: {
    SS_INVENTARIO: "1Av2UJLq7zOi4zlzt7PLYXhOV34qV_SIYUcIPwaF-od4",
    SS_ERP: "1qFbcn2alEpRsTQWREsElbLSbuPC4LJNdu7KDtdkn2UU",
    SS_OPERACION: "1VfgdfLmmIti2XljXAepw7_xcqv1ZQ73g3f071FQaSVk"
  },
  HOJAS_OP: {
    ODS: {
      NOMBRE: "ODS",
      COLUMNAS: [
        "id_ods",
        "nombre_evento",
        "id_vendedor",
        "vendedor_nombre",
        "id_locacion",
        "locacion_nombre",
        "fecha_montaje",
        "estado",
        "fecha_creacion",
        "ultima_mod",
        "activo"
      ]
    },
    ZONAS: {
      NOMBRE: "ZONAS_ALISTAMIENTO",
      COLUMNAS: [
        "id_reg",
        "zona",
        "bodega_base",
        "codigo_art",
        "nombre_art",
        "tipo_art",
        "cantidad",
        "id_ods",
        "fecha_montaje",
        "locacion_nombre",
        "nombre_evento",
        "fecha_actualizacion",
        "activo"
      ]
    },
    VIAJES: {
      NOMBRE: "VIAJES",
      COLUMNAS: [
        "id_viaje",
        "tipo_viaje",
        "bodega_origen",
        "destino",
        "id_ods_general",
        "nombre_evento_general",
        "fecha_montaje_general",
        "vehiculo",
        "chofer",
        "estado",
        "fecha_salida",
        "fecha_cierre",
        "observaciones_cierre",
        "activo"
      ]
    },
    VIAJES_DETALLE: {
      NOMBRE: "VIAJES_DETALLE",
      COLUMNAS: [
        "id_det",
        "id_viaje",
        "codigo_art",
        "nombre_art",
        "tipo_art",
        "zona_origen",
        "zona_destino",
        "id_ods",
        "nombre_evento",
        "locacion_nombre",
        "fecha_montaje",
        "cant_enviada",
        "cant_recibida",
        "cant_diferencia",
        "novedad",
        "estado_item",
        "activo"
      ]
    },
    HISTORIAL: {
      NOMBRE: "HISTORIAL_MOVIMIENTOS",
      COLUMNAS: [
        "id_mov",
        "fecha_hora",
        "fecha_iso",
        "id_viaje",
        "tipo_movimiento",
        "bodega_origen",
        "zona_origen",
        "bodega_destino",
        "zona_destino",
        "id_ods",
        "nombre_evento",
        "locacion_nombre",
        "vendedor_nombre",
        "fecha_ods",
        "codigo_art",
        "nombre_art",
        "cantidad",
        "vehiculo",
        "chofer",
        "novedad",
        "resumen_texto",
        "activo"
      ]
    }
  }
};

/**
 * Respuesta JSON estándar compatible con fetch desde GitHub Pages
 */
function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Formatea fecha/hora en zona horaria de Cancún
 */
function _formatearFechaHora(val) {
  if (val === null || val === undefined || val === "") return "";
  const tz = EYS_CONFIG.TIMEZONE;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, tz, "yyyy-MM-dd HH:mm");
  }
  const str = String(val).trim();
  if (!str) return "";
  if (str.includes("T")) {
    const p = str.split("T");
    return p[0] + " " + (p[1] || "").substring(0, 5);
  }
  return str;
}

function _ahoraStr() {
  return Utilities.formatDate(new Date(), EYS_CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm");
}

function _hoyIso() {
  return Utilities.formatDate(new Date(), EYS_CONFIG.TIMEZONE, "yyyy-MM-dd");
}

/**
 * Garantiza que las pestañas y encabezados existan en el Archivo 3 (Operación)
 */
function _garantizarHojasOperacion(ssOp) {
  const definiciones = EYS_CONFIG.HOJAS_OP;
  Object.keys(definiciones).forEach(function(key) {
    const def = definiciones[key];
    let sheet = ssOp.getSheetByName(def.NOMBRE);
    if (!sheet) {
      sheet = ssOp.insertSheet(def.NOMBRE);
      sheet.getRange(1, 1, 1, def.COLUMNAS.length).setValues([def.COLUMNAS]);
      sheet.getRange(1, 1, 1, def.COLUMNAS.length)
        .setFontWeight("bold")
        .setBackground("#003366")
        .setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    } else if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, def.COLUMNAS.length).setValues([def.COLUMNAS]);
      sheet.setFrozenRows(1);
    }
  });
}

/**
 * Determina la sub-zona de BONFIL según el tipo de artículo o zona explícita
 */
function _resolverZonaBonfilPorTipo(tipoArt, zonaSugerida) {
  const zUp = String(zonaSugerida || "").trim().toUpperCase();
  if (
    zUp === "BONFIL MOBILIARIO" ||
    zUp === "BONFIL MANUALIDADES" ||
    zUp === "BONFIL TELAS" ||
    zUp === "BONFIL ILUMINACIÓN" ||
    zUp === "BONFIL ILUMINACION"
  ) {
    return zUp === "BONFIL ILUMINACION" ? "BONFIL ILUMINACIÓN" : zUp;
  }
  const t = String(tipoArt || "").toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (t.includes("MANUALIDAD") || t.includes("DECORACION") || t.includes("ACCESORIO") || t.includes("FLOR")) {
    return "BONFIL MANUALIDADES";
  }
  if (t.includes("TELA") || t.includes("MANTEL") || t.includes(" CORTIN") || t.includes("COJIN") || t.includes("SERVILLETA")) {
    return "BONFIL TELAS";
  }
  if (t.includes("ILUMINACION") || t.includes("LUZ") || t.includes("AUDIO") || t.includes("ELECTRIC") || t.includes("LAMPARA") || t.includes("NEON")) {
    return "BONFIL ILUMINACIÓN";
  }
  return "BONFIL MOBILIARIO";
}

function _obtenerBodegaBaseDeZona(zona) {
  const z = String(zona || "").trim().toUpperCase();
  if (z.indexOf("BONFIL") === 0) return "BONFIL";
  if (z === "CEDIS") return "CEDIS";
  if (z === "PUERTO") return "PUERTO";
  if (z === "LOCACIONES") return "LOCACIONES";
  return z;
}

/**
 * Punto de entrada GET (para carga inicial o consulta de etiqueta)
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = String(params.action || "init").trim();

    if (action === "init") {
      return _jsonResponse(apiInicializarSistema());
    }
    if (action === "getDatos") {
      return _jsonResponse(apiObtenerDatosEtiqueta(params.ods, params.cop));
    }
    return _jsonResponse({ exito: true, mensaje: "API Movimientos EyS Zúñiga activa" });
  } catch (err) {
    return _jsonResponse({ exito: false, error: err.toString() });
  }
}

/**
 * Punto de entrada POST (para todas las operaciones de escritura y consulta)
 */
function doPost(e) {
  try {
    let body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    const action = String(body.action || "").trim();

    switch (action) {
      case "init":
        return _jsonResponse(apiInicializarSistema());
      case "guardarODS":
        return _jsonResponse(apiGuardarODS(body.payload || {}));
      case "cambiarEstadoODS":
        return _jsonResponse(apiCambiarEstadoODS(body.id_ods, body.nuevo_estado));
      case "agregarAZonaManual":
        return _jsonResponse(apiAgregarAZonaManual(body.zona, body.items || []));
      case "ajustarItemZona":
        return _jsonResponse(apiAjustarItemZona(body.id_reg, body.nueva_cantidad));
      case "crearViaje":
        return _jsonResponse(apiCrearViaje(body.payload || {}));
      case "completarViaje":
        return _jsonResponse(apiCompletarViaje(body.payload || {}));
      case "cancelarViaje":
        return _jsonResponse(apiCancelarViaje(body.id_viaje, body.motivo));
      default:
        return _jsonResponse({ exito: false, error: "Acción no reconocida: " + action });
    }
  } catch (err) {
    return _jsonResponse({ exito: false, error: err.toString() });
  }
}

/**
 * 1. CARGA INICIAL COMPLETA:
 *    - Catálogo de Artículos desde Archivo 1 (INVENTARIO fila 4+)
 *    - Locaciones, Vendedores, Vehículos y Choferes desde Archivo 2 (ERP)
 *    - ODS, Zonas de Alistamiento, Viajes e Historial desde Archivo 3 (Operación)
 */
function apiInicializarSistema() {
  const ssInv = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_INVENTARIO);
  const ssErp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_ERP);
  const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
  _garantizarHojasOperacion(ssOp);

  // A. Leer Artículos de la hoja INVENTARIO (Encabezados en fila 3, datos desde fila 4)
  const articulos = [];
  const sheetInv = ssInv.getSheetByName("INVENTARIO");
  if (sheetInv) {
    const dataInv = sheetInv.getDataRange().getValues();
    if (dataInv.length >= 4) {
      const headersInv = dataInv[2].map(function(h) {
        return String(h || "").trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      });
      const idxCod = headersInv.indexOf("CODIGO") >= 0 ? headersInv.indexOf("CODIGO") : 1;
      const idxNom = headersInv.findIndex(function(h) { return h.indexOf("NOMBRE") === 0; }) >= 0
        ? headersInv.findIndex(function(h) { return h.indexOf("NOMBRE") === 0; })
        : 2;
      const idxTipo = headersInv.findIndex(function(h) { return h.indexOf("TIPO") >= 0; }) >= 0
        ? headersInv.findIndex(function(h) { return h.indexOf("TIPO") >= 0; })
        : 3;

      const vistos = {};
      for (let i = 3; i < dataInv.length; i++) {
        const row = dataInv[i];
        const codigo = String(row[idxCod] || "").trim();
        const nombre = String(row[idxNom] || "").trim();
        const tipo = String(row[idxTipo] || "MOBILIARIO").trim();
        if (!codigo && !nombre) continue;
        const codKey = codigo || ("SIN-COD-" + i);
        if (vistos[codKey]) continue;
        vistos[codKey] = true;
        articulos.push({
          codigo: codigo,
          nombre: nombre,
          tipo: tipo,
          zona_bonfil: _resolverZonaBonfilPorTipo(tipo, "")
        });
      }
    }
  }

  // B. Leer Catálogos del ERP (LOCACION, VENDEDOR, VEHICULO, PLANTA)
  const locaciones = [];
  const sheetLoc = ssErp.getSheetByName("LOCACION");
  if (sheetLoc) {
    const dLoc = sheetLoc.getDataRange().getValues();
    for (let l = 1; l < dLoc.length; l++) {
      const idL = String(dLoc[l][0] || "").trim();
      const nomL = String(dLoc[l][1] || "").trim();
      if (idL && nomL) {
        locaciones.push({ id: idL, nombre: nomL });
      }
    }
  }

  const vendedores = [];
  const sheetVen = ssErp.getSheetByName("VENDEDOR");
  if (sheetVen) {
    const dVen = sheetVen.getDataRange().getValues();
    for (let v = 1; v < dVen.length; v++) {
      const idV = String(dVen[v][0] || "").trim();
      const nomV = String(dVen[v][1] || "").trim();
      if (idV && nomV) {
        vendedores.push({ id: idV, nombre: nomV });
      }
    }
  }

  const vehiculos = [];
  const sheetVeh = ssErp.getSheetByName("VEHICULO");
  if (sheetVeh) {
    const dVeh = sheetVeh.getDataRange().getValues();
    for (let k = 1; k < dVeh.length; k++) {
      const idVeh = String(dVeh[k][0] || "").trim();
      const nomVeh = String(dVeh[k][1] || "").trim();
      if (idVeh && nomVeh) {
        vehiculos.push({ id: idVeh, nombre: nomVeh });
      }
    }
  }

  const mapPuestos = {};
  const sheetPuesto = ssErp.getSheetByName("PUESTO");
  if (sheetPuesto) {
    const dPuesto = sheetPuesto.getDataRange().getValues();
    for (let p = 1; p < dPuesto.length; p++) {
      const idP = String(dPuesto[p][0] || "").trim();
      const nomP = String(dPuesto[p][1] || "").trim();
      if (idP && nomP) mapPuestos[idP] = nomP;
    }
  }

  const choferes = [];
  const otrosPlanta = [];
  const sheetPlanta = ssErp.getSheetByName("PLANTA");
  if (sheetPlanta) {
    const dPlanta = sheetPlanta.getDataRange().getValues();
    for (let c = 1; c < dPlanta.length; c++) {
      const rowP = dPlanta[c];
      const idEmp = String(rowP[0] || "").trim();
      const nom = String(rowP[1] || "").trim();
      const ape = String(rowP[2] || "").trim();
      const nombreCompleto = (nom + " " + ape).trim();
      if (!idEmp || !nombreCompleto) continue;

      const rawAct = rowP.length > 18 ? rowP[18] : 1;
      const strAct = String(rawAct).trim().toUpperCase();
      if (rawAct === 0 || strAct === "0" || strAct === "FALSE" || strAct === "NO") continue;

      const idPuesto = rowP.length > 14 ? String(rowP[14] || "").trim() : "";
      const nomPuesto = mapPuestos[idPuesto] || "";
      const puestoNorm = nomPuesto.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

      const itemEmp = {
        id: idEmp,
        nombre: nombreCompleto,
        puesto: nomPuesto || "Operativo"
      };
      if (puestoNorm.includes("chofer") || puestoNorm.includes("conductor") || puestoNorm.includes("operador")) {
        choferes.push(itemEmp);
      } else {
        otrosPlanta.push(itemEmp);
      }
    }
  }
  const listaChoferesFinal = choferes.concat(otrosPlanta);

  // C. Leer Datos Operativos del Archivo 3
  const estadoOp = _leerEstadoOperativoCompleto(ssOp);

  return {
    exito: true,
    catalogos: {
      articulos: articulos,
      locaciones: locaciones,
      vendedores: vendedores,
      vehiculos: vehiculos,
      choferes: listaChoferesFinal
    },
    ods: estadoOp.ods,
    zonas: estadoOp.zonas,
    viajes: estadoOp.viajes,
    historial: estadoOp.historial
  };
}

/**
 * Lee ODS, Zonas de Alistamiento, Viajes (con sus detalles) e Historial del Archivo 3
 */
function _leerEstadoOperativoCompleto(ssOp) {
  // 1. ODS
  const listaODS = [];
  const sheetOds = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ODS.NOMBRE);
  if (sheetOds && sheetOds.getLastRow() > 1) {
    const dOds = sheetOds.getDataRange().getValues();
    for (let i = 1; i < dOds.length; i++) {
      const r = dOds[i];
      const idOds = String(r[0] || "").trim();
      const act = r[10];
      if (!idOds || act === 0 || String(act) === "0") continue;
      const fMont = _formatearFechaHora(r[6]);
      listaODS.push({
        id_ods: idOds,
        nombre_evento: String(r[1] || "").trim(),
        id_vendedor: String(r[2] || "").trim(),
        vendedor_nombre: String(r[3] || "").trim(),
        id_locacion: String(r[4] || "").trim(),
        locacion_nombre: String(r[5] || "").trim(),
        fecha_montaje: fMont,
        fecha_iso: fMont ? fMont.substring(0, 10) : "",
        estado: String(r[7] || "ACTIVA").trim().toUpperCase(),
        fecha_creacion: _formatearFechaHora(r[8]),
        ultima_mod: _formatearFechaHora(r[9])
      });
    }
  }
  // Ordenar ODS desde la fecha de montaje más próxima hasta la más lejana
  listaODS.sort(function(a, b) {
    const fa = a.fecha_montaje || "9999-12-31";
    const fb = b.fecha_montaje || "9999-12-31";
    return fa.localeCompare(fb);
  });

  // 2. ZONAS_ALISTAMIENTO
  const listaZonas = [];
  const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);
  if (sheetZonas && sheetZonas.getLastRow() > 1) {
    const dZ = sheetZonas.getDataRange().getValues();
    for (let z = 1; z < dZ.length; z++) {
      const rz = dZ[z];
      const idReg = String(rz[0] || "").trim();
      const cant = Number(rz[6]) || 0;
      const actZ = rz[12];
      if (!idReg || cant <= 0 || actZ === 0 || String(actZ) === "0") continue;
      listaZonas.push({
        id_reg: idReg,
        zona: String(rz[1] || "").trim(),
        bodega_base: String(rz[2] || "").trim(),
        codigo_art: String(rz[3] || "").trim(),
        nombre_art: String(rz[4] || "").trim(),
        tipo_art: String(rz[5] || "").trim(),
        cantidad: cant,
        id_ods: String(rz[7] || "").trim(),
        fecha_montaje: _formatearFechaHora(rz[8]),
        locacion_nombre: String(rz[9] || "").trim(),
        nombre_evento: String(rz[10] || "").trim(),
        fecha_actualizacion: _formatearFechaHora(rz[11])
      });
    }
  }

  // 3. VIAJES y VIAJES_DETALLE
  const detallesPorViaje = {};
  const sheetDet = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES_DETALLE.NOMBRE);
  if (sheetDet && sheetDet.getLastRow() > 1) {
    const dDet = sheetDet.getDataRange().getValues();
    for (let d = 1; d < dDet.length; d++) {
      const rd = dDet[d];
      const idDet = String(rd[0] || "").trim();
      const idViaje = String(rd[1] || "").trim();
      const actD = rd[16];
      if (!idDet || !idViaje || actD === 0 || String(actD) === "0") continue;
      if (!detallesPorViaje[idViaje]) detallesPorViaje[idViaje] = [];
      detallesPorViaje[idViaje].push({
        id_det: idDet,
        id_viaje: idViaje,
        codigo_art: String(rd[2] || "").trim(),
        nombre_art: String(rd[3] || "").trim(),
        tipo_art: String(rd[4] || "").trim(),
        zona_origen: String(rd[5] || "").trim(),
        zona_destino: String(rd[6] || "").trim(),
        id_ods: String(rd[7] || "").trim(),
        nombre_evento: String(rd[8] || "").trim(),
        locacion_nombre: String(rd[9] || "").trim(),
        fecha_montaje: _formatearFechaHora(rd[10]),
        cant_enviada: Number(rd[11]) || 0,
        cant_recibida: Number(rd[12]) || 0,
        cant_diferencia: Number(rd[13]) || 0,
        novedad: String(rd[14] || "").trim(),
        estado_item: String(rd[15] || "EN_CURSO").trim()
      });
    }
  }

  const listaViajes = [];
  const sheetViajes = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES.NOMBRE);
  if (sheetViajes && sheetViajes.getLastRow() > 1) {
    const dV = sheetViajes.getDataRange().getValues();
    for (let v = 1; v < dV.length; v++) {
      const rv = dV[v];
      const idV = String(rv[0] || "").trim();
      const actV = rv[13];
      if (!idV || actV === 0 || String(actV) === "0") continue;
      listaViajes.push({
        id_viaje: idV,
        tipo_viaje: String(rv[1] || "TRASLADO_BODEGA").trim(),
        bodega_origen: String(rv[2] || "").trim(),
        destino: String(rv[3] || "").trim(),
        id_ods_general: String(rv[4] || "").trim(),
        nombre_evento_general: String(rv[5] || "").trim(),
        fecha_montaje_general: _formatearFechaHora(rv[6]),
        vehiculo: String(rv[7] || "").trim(),
        chofer: String(rv[8] || "").trim(),
        estado: String(rv[9] || "EN_CURSO").trim().toUpperCase(),
        fecha_salida: _formatearFechaHora(rv[10]),
        fecha_cierre: _formatearFechaHora(rv[11]),
        observaciones_cierre: String(rv[12] || "").trim(),
        items: detallesPorViaje[idV] || []
      });
    }
  }
  listaViajes.sort(function(a, b) {
    return String(b.fecha_salida || "").localeCompare(String(a.fecha_salida || ""));
  });

  // 4. HISTORIAL_MOVIMIENTOS
  const listaHistorial = [];
  const sheetHist = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.HISTORIAL.NOMBRE);
  if (sheetHist && sheetHist.getLastRow() > 1) {
    const dH = sheetHist.getDataRange().getValues();
    for (let h = 1; h < dH.length; h++) {
      const rh = dH[h];
      const idMov = String(rh[0] || "").trim();
      const actH = rh[21];
      if (!idMov || actH === 0 || String(actH) === "0") continue;
      listaHistorial.push({
        id_mov: idMov,
        fecha_hora: _formatearFechaHora(rh[1]),
        fecha_iso: String(rh[2] || "").trim() || (_formatearFechaHora(rh[1]).substring(0, 10)),
        id_viaje: String(rh[3] || "").trim(),
        tipo_movimiento: String(rh[4] || "").trim(),
        bodega_origen: String(rh[5] || "").trim(),
        zona_origen: String(rh[6] || "").trim(),
        bodega_destino: String(rh[7] || "").trim(),
        zona_destino: String(rh[8] || "").trim(),
        id_ods: String(rh[9] || "").trim(),
        nombre_evento: String(rh[10] || "").trim(),
        locacion_nombre: String(rh[11] || "").trim(),
        vendedor_nombre: String(rh[12] || "").trim(),
        fecha_ods: _formatearFechaHora(rh[13]),
        codigo_art: String(rh[14] || "").trim(),
        nombre_art: String(rh[15] || "").trim(),
        cantidad: Number(rh[16]) || 0,
        vehiculo: String(rh[17] || "").trim(),
        chofer: String(rh[18] || "").trim(),
        novedad: String(rh[19] || "").trim(),
        resumen_texto: String(rh[20] || "").trim()
      });
    }
  }
  listaHistorial.sort(function(a, b) {
    return String(b.fecha_hora || "").localeCompare(String(a.fecha_hora || ""));
  });

  return {
    ods: listaODS,
    zonas: listaZonas,
    viajes: listaViajes,
    historial: listaHistorial
  };
}

/**
 * 2. CREAR O EDITAR ODS
 */
function apiGuardarODS(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetOds = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ODS.NOMBRE);

    const idOdsOriginal = String(payload.id_ods_original || "").trim();
    const idOds = String(payload.id_ods || "").trim();
    const nombreEv = String(payload.nombre_evento || "").trim();
    const idVend = String(payload.id_vendedor || "").trim();
    const nomVend = String(payload.vendedor_nombre || "").trim();
    const idLoc = String(payload.id_locacion || "").trim();
    const nomLoc = String(payload.locacion_nombre || "").trim();
    const fMont = _formatearFechaHora(payload.fecha_montaje);
    const ahora = _ahoraStr();

    if (!idOds || !nombreEv || !nomLoc || !fMont) {
      return { exito: false, error: "Faltan datos obligatorios de la ODS." };
    }

    const data = sheetOds.getDataRange().getValues();
    let filaEditar = -1;
    const idBuscar = idOdsOriginal || idOds;

    for (let i = 1; i < data.length; i++) {
      const idFila = String(data[i][0] || "").trim();
      const actFila = data[i][10];
      if (actFila === 0 || String(actFila) === "0") continue;
      if (idFila === idBuscar) {
        filaEditar = i + 1;
        break;
      }
    }

    if (payload.modo_edicion && filaEditar > 0) {
      sheetOds.getRange(filaEditar, 1, 1, 10).setValues([[
        idOds,
        nombreEv,
        idVend,
        nomVend,
        idLoc,
        nomLoc,
        fMont,
        String(data[filaEditar - 1][7] || "ACTIVA"),
        data[filaEditar - 1][8] || ahora,
        ahora
      ]]);
    } else {
      if (filaEditar > 0 && !payload.modo_edicion) {
        return { exito: false, error: "La ODS #" + idOds + " ya existe." };
      }
      sheetOds.appendRow([
        idOds,
        nombreEv,
        idVend,
        nomVend,
        idLoc,
        nomLoc,
        fMont,
        "ACTIVA",
        ahora,
        ahora,
        1
      ]);
    }

    const estadoOp = _leerEstadoOperativoCompleto(ssOp);
    return {
      exito: true,
      mensaje: payload.modo_edicion ? "ODS actualizada correctamente." : "ODS creada correctamente.",
      ods: estadoOp.ods,
      zonas: estadoOp.zonas
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 3. CAMBIAR ESTADO DE UNA ODS (COMPLETADA / CANCELADA / ACTIVA)
 */
function apiCambiarEstadoODS(idOds, nuevoEstado) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetOds = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ODS.NOMBRE);
    const data = sheetOds.getDataRange().getValues();
    const target = String(idOds || "").trim();
    const est = String(nuevoEstado || "ACTIVA").trim().toUpperCase();
    const ahora = _ahoraStr();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === target && data[i][10] !== 0 && String(data[i][10]) !== "0") {
        sheetOds.getRange(i + 1, 8).setValue(est);
        sheetOds.getRange(i + 1, 10).setValue(ahora);
        break;
      }
    }

    const estadoOp = _leerEstadoOperativoCompleto(ssOp);
    return { exito: true, ods: estadoOp.ods };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper interno: Suma o resta stock en ZONAS_ALISTAMIENTO
 */
function _modificarStockZona(sheetZonas, params) {
  const zona = String(params.zona || "").trim().toUpperCase();
  const bodegaBase = _obtenerBodegaBaseDeZona(zona);
  const codigoArt = String(params.codigo_art || "").trim();
  const nombreArt = String(params.nombre_art || "").trim();
  const tipoArt = String(params.tipo_art || "MOBILIARIO").trim();
  const idOds = String(params.id_ods || "").trim();
  const fMont = _formatearFechaHora(params.fecha_montaje);
  const locNom = String(params.locacion_nombre || "").trim();
  const nomEv = String(params.nombre_evento || "").trim();
  const delta = Number(params.delta) || 0;
  const ahora = _ahoraStr();

  if (!zona || !codigoArt || delta === 0) return;

  const data = sheetZonas.getDataRange().getValues();
  let filaEncontrada = -1;
  let cantActual = 0;

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (r[12] === 0 || String(r[12]) === "0") continue;
    const zRow = String(r[1] || "").trim().toUpperCase();
    const cRow = String(r[3] || "").trim();
    const oRow = String(r[7] || "").trim();
    if (zRow === zona && cRow === codigoArt && oRow === idOds) {
      filaEncontrada = i + 1;
      cantActual = Number(r[6]) || 0;
      break;
    }
  }

  if (filaEncontrada > 0) {
    const nuevaCant = Math.max(0, cantActual + delta);
    sheetZonas.getRange(filaEncontrada, 7).setValue(nuevaCant);
    sheetZonas.getRange(filaEncontrada, 12).setValue(ahora);
    if (nuevaCant === 0) {
      sheetZonas.getRange(filaEncontrada, 13).setValue(0);
    }
  } else if (delta > 0) {
    const idReg = "ZA-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
    sheetZonas.appendRow([
      idReg,
      zona,
      bodegaBase,
      codigoArt,
      nombreArt,
      tipoArt,
      delta,
      idOds,
      fMont,
      locNom,
      nomEv,
      ahora,
      1
    ]);
  }
}

/**
 * Helper interno: Registra una fila en HISTORIAL_MOVIMIENTOS
 */
function _registrarHistorial(sheetHist, h) {
  const ahora = h.fecha_hora || _ahoraStr();
  const fechaIso = ahora.substring(0, 10);
  const idMov = "MOV-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
  sheetHist.appendRow([
    idMov,
    ahora,
    fechaIso,
    h.id_viaje || "",
    h.tipo_movimiento || "",
    h.bodega_origen || "",
    h.zona_origen || "",
    h.bodega_destino || "",
    h.zona_destino || "",
    h.id_ods || "",
    h.nombre_evento || "",
    h.locacion_nombre || "",
    h.vendedor_nombre || "",
    h.fecha_ods || "",
    h.codigo_art || "",
    h.nombre_art || "",
    Number(h.cantidad) || 0,
    h.vehiculo || "",
    h.chofer || "",
    h.novedad || "",
    h.resumen_texto || "",
    1
  ]);
}

/**
 * 4. AGREGAR UNO O VARIOS ARTÍCULOS MANUALMENTE A UNA ZONA DE ALISTAMIENTO
 */
function apiAgregarAZonaManual(zona, items) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);
    const sheetHist = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.HISTORIAL.NOMBRE);
    const zonaClean = String(zona || "").trim().toUpperCase();
    const bodegaBase = _obtenerBodegaBaseDeZona(zonaClean);
    const ahora = _ahoraStr();

    (items || []).forEach(function(it) {
      const cant = Number(it.cantidad) || 0;
      if (!it.codigo_art || cant <= 0) return;
      _modificarStockZona(sheetZonas, {
        zona: zonaClean,
        codigo_art: it.codigo_art,
        nombre_art: it.nombre_art,
        tipo_art: it.tipo_art || "MOBILIARIO",
        id_ods: it.id_ods,
        fecha_montaje: it.fecha_montaje,
        locacion_nombre: it.locacion_nombre,
        nombre_evento: it.nombre_evento,
        delta: cant
      });

      _registrarHistorial(sheetHist, {
        fecha_hora: ahora,
        id_viaje: "INGRESO-ZONA",
        tipo_movimiento: "INGRESO_ALISTAMIENTO",
        bodega_origen: bodegaBase,
        zona_origen: "ALMACÉN / TALLER",
        bodega_destino: bodegaBase,
        zona_destino: zonaClean,
        id_ods: it.id_ods,
        nombre_evento: it.nombre_evento,
        locacion_nombre: it.locacion_nombre,
        vendedor_nombre: it.vendedor_nombre || "",
        fecha_ods: it.fecha_montaje,
        codigo_art: it.codigo_art,
        nombre_art: it.nombre_art,
        cantidad: cant,
        vehiculo: "",
        chofer: "",
        novedad: "",
        resumen_texto: "ODS " + it.id_ods + " x " + cant + " (" + it.nombre_art + ") ingresó a zona " + zonaClean
      });
    });

    const estadoOp = _leerEstadoOperativoCompleto(ssOp);
    return {
      exito: true,
      zonas: estadoOp.zonas,
      historial: estadoOp.historial
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 5. AJUSTAR O RETIRAR ARTÍCULO DE ZONA DE ALISTAMIENTO
 */
function apiAjustarItemZona(idReg, nuevaCantidad) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);
    const data = sheetZonas.getDataRange().getValues();
    const target = String(idReg || "").trim();
    const cantNum = Math.max(0, Number(nuevaCantidad) || 0);
    const ahora = _ahoraStr();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === target && data[i][12] !== 0 && String(data[i][12]) !== "0") {
        sheetZonas.getRange(i + 1, 7).setValue(cantNum);
        sheetZonas.getRange(i + 1, 12).setValue(ahora);
        if (cantNum === 0) {
          sheetZonas.getRange(i + 1, 13).setValue(0);
        }
        break;
      }
    }

    const estadoOp = _leerEstadoOperativoCompleto(ssOp);
    return { exito: true, zonas: estadoOp.zonas };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 6. CREAR UN NUEVO VIAJE / MOVIMIENTO (TRASLADO DE BODEGA O ENVÍO A EVENTO)
 */
function apiCrearViaje(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetViajes = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES.NOMBRE);
    const sheetDet = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES_DETALLE.NOMBRE);
    const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);
    const sheetHist = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.HISTORIAL.NOMBRE);

    const tipoViaje = String(payload.tipo_viaje || "TRASLADO_BODEGA").trim().toUpperCase();
    const bodegaOrigen = String(payload.bodega_origen || "").trim().toUpperCase();
    const destino = String(payload.destino || "").trim();
    const vehiculo = String(payload.vehiculo || "").trim();
    const chofer = String(payload.chofer || "").trim();
    const idOdsGeneral = String(payload.id_ods_general || "").trim();
    const nomEvGeneral = String(payload.nombre_evento_general || "").trim();
    const fMontGeneral = _formatearFechaHora(payload.fecha_montaje_general);
    const items = Array.isArray(payload.items) ? payload.items : [];

    if (!bodegaOrigen || !destino || items.length === 0) {
      return { exito: false, error: "Completa origen, destino y al menos 1 artículo en la carga." };
    }

    if (tipoViaje === "TRASLADO_BODEGA" && _obtenerBodegaBaseDeZona(bodegaOrigen) === _obtenerBodegaBaseDeZona(destino)) {
      return { exito: false, error: "No se puede realizar un traslado hacia la misma bodega (" + bodegaOrigen + ")." };
    }

    const numViaje = sheetViajes.getLastRow();
    const idViaje = "VJ-" + String(numViaje).padStart(4, "0");
    const ahora = _ahoraStr();

    sheetViajes.appendRow([
      idViaje,
      tipoViaje,
      bodegaOrigen,
      destino,
      idOdsGeneral,
      nomEvGeneral,
      fMontGeneral,
      vehiculo,
      chofer,
      "EN_CURSO",
      ahora,
      "",
      "",
      1
    ]);

    items.forEach(function(it, idx) {
      const cant = Number(it.cantidad) || 0;
      if (!it.codigo_art || cant <= 0) return;

      const idOdsItem = tipoViaje === "ENVIO_EVENTO" ? idOdsGeneral : String(it.id_ods || "").trim();
      const nomEvItem = tipoViaje === "ENVIO_EVENTO" ? nomEvGeneral : String(it.nombre_evento || "").trim();
      const locItem = tipoViaje === "ENVIO_EVENTO" ? destino : String(it.locacion_nombre || "").trim();
      const fMontItem = tipoViaje === "ENVIO_EVENTO" ? fMontGeneral : _formatearFechaHora(it.fecha_montaje);
      const tipoArt = String(it.tipo_art || "MOBILIARIO").trim();

      const zonaOrigenItem = bodegaOrigen === "BONFIL"
        ? _resolverZonaBonfilPorTipo(tipoArt, it.zona_origen)
        : bodegaOrigen;

      let zonaDestinoItem = destino;
      if (tipoViaje === "TRASLADO_BODEGA") {
        zonaDestinoItem = destino === "BONFIL"
          ? _resolverZonaBonfilPorTipo(tipoArt, it.zona_destino)
          : destino;
      } else {
        zonaDestinoItem = "LOCACIONES";
      }

      // 1. Descontar de la Zona de Alistamiento Origen
      _modificarStockZona(sheetZonas, {
        zona: zonaOrigenItem,
        codigo_art: it.codigo_art,
        nombre_art: it.nombre_art,
        tipo_art: tipoArt,
        id_ods: idOdsItem,
        fecha_montaje: fMontItem,
        locacion_nombre: locItem,
        nombre_evento: nomEvItem,
        delta: -cant
      });

      // 2. Si es ENVIO_EVENTO, entra inmediatamente a LOCACIONES (en el hotel)
      if (tipoViaje === "ENVIO_EVENTO") {
        _modificarStockZona(sheetZonas, {
          zona: "LOCACIONES",
          codigo_art: it.codigo_art,
          nombre_art: it.nombre_art,
          tipo_art: tipoArt,
          id_ods: idOdsItem,
          fecha_montaje: fMontItem,
          locacion_nombre: destino,
          nombre_evento: nomEvItem,
          delta: cant
        });
      }

      const idDet = idViaje + "-D" + (idx + 1);
      sheetDet.appendRow([
        idDet,
        idViaje,
        it.codigo_art,
        it.nombre_art,
        tipoArt,
        zonaOrigenItem,
        zonaDestinoItem,
        idOdsItem,
        nomEvItem,
        locItem,
        fMontItem,
        cant,
        cant,
        0,
        "",
        "EN_CURSO",
        1
      ]);

      const textoHist = tipoViaje === "ENVIO_EVENTO"
        ? ("ODS " + idOdsItem + " x " + cant + " (" + it.nombre_art + ") se envió de " + zonaOrigenItem + " a hotel " + destino)
        : ("ODS " + idOdsItem + " x " + cant + " (" + it.nombre_art + ") salió en traslado de " + zonaOrigenItem + " hacia " + zonaDestinoItem);

      _registrarHistorial(sheetHist, {
        fecha_hora: ahora,
        id_viaje: idViaje,
        tipo_movimiento: tipoViaje === "ENVIO_EVENTO" ? "ENVIO_A_EVENTO" : "SALIDA_TRASLADO",
        bodega_origen: bodegaOrigen,
        zona_origen: zonaOrigenItem,
        bodega_destino: tipoViaje === "ENVIO_EVENTO" ? "LOCACIONES" : destino,
        zona_destino: zonaDestinoItem,
        id_ods: idOdsItem,
        nombre_evento: nomEvItem,
        locacion_nombre: locItem,
        vendedor_nombre: it.vendedor_nombre || "",
        fecha_ods: fMontItem,
        codigo_art: it.codigo_art,
        nombre_art: it.nombre_art,
        cantidad: cant,
        vehiculo: vehiculo,
        chofer: chofer,
        novedad: "",
        resumen_texto: textoHist
      });
    });

    const estadoOp = _leerEstadoOperativoCompleto(ssOp);
    return {
      exito: true,
      id_viaje: idViaje,
      zonas: estadoOp.zonas,
      viajes: estadoOp.viajes,
      historial: estadoOp.historial
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 7. COMPLETAR / CERRAR UN VIAJE (RECEPCIÓN DE TRASLADO O REGRESO DE EVENTO)
 */
function apiCompletarViaje(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetViajes = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES.NOMBRE);
    const sheetDet = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES_DETALLE.NOMBRE);
    const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);
    const sheetHist = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.HISTORIAL.NOMBRE);

    const idViaje = String(payload.id_viaje || "").trim();
    const obsGeneral = String(payload.observaciones_cierre || "").trim();
    const itemsRecibidos = payload.items_recibidos || {};
    const ahora = _ahoraStr();

    const dViajes = sheetViajes.getDataRange().getValues();
    let filaViaje = -1;
    let infoViaje = null;

    for (let v = 1; v < dViajes.length; v++) {
      if (String(dViajes[v][0] || "").trim() === idViaje && dViajes[v][13] !== 0 && String(dViajes[v][13]) !== "0") {
        filaViaje = v + 1;
        infoViaje = {
          tipo_viaje: String(dViajes[v][1] || "TRASLADO_BODEGA").trim().toUpperCase(),
          bodega_origen: String(dViajes[v][2] || "").trim(),
          destino: String(dViajes[v][3] || "").trim(),
          vehiculo: String(dViajes[v][7] || "").trim(),
          chofer: String(dViajes[v][8] || "").trim(),
          estado: String(dViajes[v][9] || "").trim().toUpperCase()
        };
        break;
      }
    }

    if (filaViaje === -1 || !infoViaje) {
      return { exito: false, error: "No se encontró el viaje " + idViaje };
    }
    if (infoViaje.estado === "COMPLETADO") {
      return { exito: false, error: "Este viaje ya fue marcado como completado." };
    }

    const dDet = sheetDet.getDataRange().getValues();
    for (let d = 1; d < dDet.length; d++) {
      const rd = dDet[d];
      if (String(rd[1] || "").trim() !== idViaje || rd[16] === 0 || String(rd[16]) === "0") continue;

      const idDet = String(rd[0] || "").trim();
      const codArt = String(rd[2] || "").trim();
      const nomArt = String(rd[3] || "").trim();
      const tipoArt = String(rd[4] || "MOBILIARIO").trim();
      const zonaOrigen = String(rd[5] || "").trim();
      const zonaDestino = String(rd[6] || "").trim();
      const idOds = String(rd[7] || "").trim();
      const nomEv = String(rd[8] || "").trim();
      const locNom = String(rd[9] || "").trim();
      const fMont = _formatearFechaHora(rd[10]);
      const cantEnv = Number(rd[11]) || 0;

      const ajuste = itemsRecibidos[idDet] || {};
      const cantRec = ajuste.cant_recibida !== undefined ? Math.max(0, Number(ajuste.cant_recibida) || 0) : cantEnv;
      const cantDif = Math.max(0, cantEnv - cantRec);
      const novedadItem = String(ajuste.novedad || "").trim();
      const estadoItem = cantDif > 0 || novedadItem ? "CON_NOVEDAD" : "COMPLETO";

      sheetDet.getRange(d + 1, 13, 1, 4).setValues([[cantRec, cantDif, novedadItem, estadoItem]]);

      if (infoViaje.tipo_viaje === "TRASLADO_BODEGA") {
        // En Traslado de Bodega: al completarse, las piezas recibidas entran a la Zona de Alistamiento Destino
        if (cantRec > 0) {
          _modificarStockZona(sheetZonas, {
            zona: zonaDestino,
            codigo_art: codArt,
            nombre_art: nomArt,
            tipo_art: tipoArt,
            id_ods: idOds,
            fecha_montaje: fMont,
            locacion_nombre: locNom,
            nombre_evento: nomEv,
            delta: cantRec
          });
        }

        const txtRec = "ODS " + idOds + " x " + cantRec + " (" + nomArt + ") se recibió en " + zonaDestino + " proveniente de " + zonaOrigen + (novedadItem ? (" [Novedad: " + novedadItem + "]") : "");
        _registrarHistorial(sheetHist, {
          fecha_hora: ahora,
          id_viaje: idViaje,
          tipo_movimiento: "RECEPCION_TRASLADO",
          bodega_origen: infoViaje.bodega_origen,
          zona_origen: zonaOrigen,
          bodega_destino: infoViaje.destino,
          zona_destino: zonaDestino,
          id_ods: idOds,
          nombre_evento: nomEv,
          locacion_nombre: locNom,
          vendedor_nombre: "",
          fecha_ods: fMont,
          codigo_art: codArt,
          nombre_art: nomArt,
          cantidad: cantRec,
          vehiculo: infoViaje.vehiculo,
          chofer: infoViaje.chofer,
          novedad: novedadItem,
          resumen_texto: txtRec
        });
      } else {
        // En ENVIO_EVENTO: al completarse significa que REGRESÓ del hotel a la bodega origen
        _modificarStockZona(sheetZonas, {
          zona: "LOCACIONES",
          codigo_art: codArt,
          nombre_art: nomArt,
          tipo_art: tipoArt,
          id_ods: idOds,
          fecha_montaje: fMont,
          locacion_nombre: infoViaje.destino,
          nombre_evento: nomEv,
          delta: -cantEnv
        });

        if (cantRec > 0) {
          _modificarStockZona(sheetZonas, {
            zona: zonaOrigen,
            codigo_art: codArt,
            nombre_art: nomArt,
            tipo_art: tipoArt,
            id_ods: idOds,
            fecha_montaje: fMont,
            locacion_nombre: locNom,
            nombre_evento: nomEv,
            delta: cantRec
          });
        }

        const txtReg = "ODS " + idOds + " x " + cantRec + " (" + nomArt + ") regresó de hotel " + infoViaje.destino + " a " + zonaOrigen + (novedadItem ? (" [Novedad/Daño: " + novedadItem + "]") : "");
        _registrarHistorial(sheetHist, {
          fecha_hora: ahora,
          id_viaje: idViaje,
          tipo_movimiento: "REGRESO_DE_EVENTO",
          bodega_origen: "LOCACIONES",
          zona_origen: infoViaje.destino,
          bodega_destino: infoViaje.bodega_origen,
          zona_destino: zonaOrigen,
          id_ods: idOds,
          nombre_evento: nomEv,
          locacion_nombre: infoViaje.destino,
          vendedor_nombre: "",
          fecha_ods: fMont,
          codigo_art: codArt,
          nombre_art: nomArt,
          cantidad: cantRec,
          vehiculo: infoViaje.vehiculo,
          chofer: infoViaje.chofer,
          novedad: novedadItem,
          resumen_texto: txtReg
        });
      }

      if (cantDif > 0) {
        _registrarHistorial(sheetHist, {
          fecha_hora: ahora,
          id_viaje: idViaje,
          tipo_movimiento: "NOVEDAD_FALTANTE",
          bodega_origen: infoViaje.bodega_origen,
          zona_origen: zonaOrigen,
          bodega_destino: infoViaje.destino,
          zona_destino: zonaDestino,
          id_ods: idOds,
          nombre_evento: nomEv,
          locacion_nombre: locNom || infoViaje.destino,
          vendedor_nombre: "",
          fecha_ods: fMont,
          codigo_art: codArt,
          nombre_art: nomArt,
          cantidad: cantDif,
          vehiculo: infoViaje.vehiculo,
          chofer: infoViaje.chofer,
          novedad: novedadItem || "Faltante al completar viaje",
          resumen_texto: "⚠️ ODS " + idOds + ": Faltaron " + cantDif + " pza(s) de " + nomArt + " en viaje " + idViaje + (novedadItem ? (" (" + novedadItem + ")") : "")
        });
      }
    }

    sheetViajes.getRange(filaViaje, 10).setValue("COMPLETADO");
    sheetViajes.getRange(filaViaje, 12).setValue(ahora);
    sheetViajes.getRange(filaViaje, 13).setValue(obsGeneral);

    const estadoOp = _leerEstadoOperativoCompleto(ssOp);
    return {
      exito: true,
      zonas: estadoOp.zonas,
      viajes: estadoOp.viajes,
      historial: estadoOp.historial
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 8. CANCELAR UN VIAJE EN CURSO (REVIERTE STOCK A ORIGEN)
 */
function apiCancelarViaje(idViaje, motivo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetViajes = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES.NOMBRE);
    const sheetDet = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES_DETALLE.NOMBRE);
    const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);
    const sheetHist = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.HISTORIAL.NOMBRE);

    const target = String(idViaje || "").trim();
    const ahora = _ahoraStr();

    const dViajes = sheetViajes.getDataRange().getValues();
    let filaViaje = -1;
    let infoViaje = null;
    for (let v = 1; v < dViajes.length; v++) {
      if (String(dViajes[v][0] || "").trim() === target && dViajes[v][13] !== 0 && String(dViajes[v][13]) !== "0") {
        filaViaje = v + 1;
        infoViaje = {
          tipo_viaje: String(dViajes[v][1] || "").trim().toUpperCase(),
          bodega_origen: String(dViajes[v][2] || "").trim(),
          destino: String(dViajes[v][3] || "").trim()
        };
        break;
      }
    }

    if (filaViaje === -1 || !infoViaje) {
      return { exito: false, error: "Viaje no encontrado." };
    }

    const dDet = sheetDet.getDataRange().getValues();
    for (let d = 1; d < dDet.length; d++) {
      const rd = dDet[d];
      if (String(rd[1] || "").trim() !== target || rd[16] === 0 || String(rd[16]) === "0") continue;

      const codArt = String(rd[2] || "").trim();
      const nomArt = String(rd[3] || "").trim();
      const tipoArt = String(rd[4] || "").trim();
      const zonaOrigen = String(rd[5] || "").trim();
      const idOds = String(rd[7] || "").trim();
      const nomEv = String(rd[8] || "").trim();
      const locNom = String(rd[9] || "").trim();
      const fMont = _formatearFechaHora(rd[10]);
      const cantEnv = Number(rd[11]) || 0;

      // Devolver a la zona origen
      _modificarStockZona(sheetZonas, {
        zona: zonaOrigen,
        codigo_art: codArt,
        nombre_art: nomArt,
        tipo_art: tipoArt,
        id_ods: idOds,
        fecha_montaje: fMont,
        locacion_nombre: locNom,
        nombre_evento: nomEv,
        delta: cantEnv
      });

      // Si era ENVIO_EVENTO, quitar de LOCACIONES
      if (infoViaje.tipo_viaje === "ENVIO_EVENTO") {
        _modificarStockZona(sheetZonas, {
          zona: "LOCACIONES",
          codigo_art: codArt,
          nombre_art: nomArt,
          tipo_art: tipoArt,
          id_ods: idOds,
          fecha_montaje: fMont,
          locacion_nombre: infoViaje.destino,
          nombre_evento: nomEv,
          delta: -cantEnv
        });
      }

      sheetDet.getRange(d + 1, 16).setValue("CANCELADO");
    }

    sheetViajes.getRange(filaViaje, 10).setValue("CANCELADO");
    sheetViajes.getRange(filaViaje, 12).setValue(ahora);
    sheetViajes.getRange(filaViaje, 13).setValue(String(motivo || "Viaje cancelado"));

    _registrarHistorial(sheetHist, {
      fecha_hora: ahora,
      id_viaje: target,
      tipo_movimiento: "VIAJE_CANCELADO",
      bodega_origen: infoViaje.bodega_origen,
      zona_origen: infoViaje.bodega_origen,
      bodega_destino: infoViaje.destino,
      zona_destino: infoViaje.destino,
      id_ods: "",
      nombre_evento: "",
      locacion_nombre: "",
      vendedor_nombre: "",
      fecha_ods: "",
      codigo_art: "",
      nombre_art: "",
      cantidad: 0,
      vehiculo: "",
      chofer: "",
      novedad: motivo || "Cancelado",
      resumen_texto: "🚫 Viaje " + target + " cancelado (artículos devueltos a " + infoViaje.bodega_origen + ")"
    });

    const estadoOp = _leerEstadoOperativoCompleto(ssOp);
    return {
      exito: true,
      zonas: estadoOp.zonas,
      viajes: estadoOp.viajes,
      historial: estadoOp.historial
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 9. COMPATIBILIDAD DIRECTA PARA CONSULTA DE ETIQUETA INDIVIDUAL
 */
function apiObtenerDatosEtiqueta(ods, cop) {
  const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
  const ssInv = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_INVENTARIO);
  _garantizarHojasOperacion(ssOp);

  const idOds = String(ods || "").trim();
  const idCop = String(cop || "").trim().toUpperCase();

  let locacion = "SIN LOCACIÓN";
  let fechaEvento = _hoyIso();

  const sheetOds = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ODS.NOMBRE);
  if (sheetOds && sheetOds.getLastRow() > 1) {
    const dOds = sheetOds.getDataRange().getValues();
    for (let i = 1; i < dOds.length; i++) {
      if (String(dOds[i][0] || "").trim() === idOds) {
        locacion = String(dOds[i][5] || "SIN LOCACIÓN").trim();
        fechaEvento = _formatearFechaHora(dOds[i][6]).substring(0, 10);
        break;
      }
    }
  }

  let nombreArticulo = "ARTÍCULO " + idCop;
  const sheetInv = ssInv.getSheetByName("INVENTARIO");
  if (sheetInv && sheetInv.getLastRow() >= 4) {
    const dInv = sheetInv.getDataRange().getValues();
    for (let r = 3; r < dInv.length; r++) {
      if (String(dInv[r][1] || "").trim().toUpperCase() === idCop) {
        nombreArticulo = String(dInv[r][2] || "").trim();
        break;
      }
    }
  }

  return {
    success: true,
    data: {
      ods: idOds,
      cop: idCop,
      locacion: locacion,
      fechaEvento: fechaEvento,
      nombreArticulo: nombreArticulo
    }
  };
}

