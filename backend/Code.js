/**
 * ============================================================================
 * SISTEMA DE CONTROL DE ENTRADAS Y SALIDAS (MOVIMIENTOS-EYS-Z)
 * Backend API en Google Apps Script — Alineado con ERP ZUÑIGA
 * ============================================================================
 * Conecta 3 archivos de Google Sheets privados/restringidos:
 * 1. SS_INVENTARIO : 1Av2UJLq7zOi4zlzt7PLYXhOV34qV_SIYUcIPwaF-od4 (Hoja INVENTARIO, encabezados en fila 3, datos en fila 4)
 * 2. SS_CATALOGOS  : 1qFbcn2alEpRsTQWREsElbLSbuPC4LJNdu7KDtdkn2UU (Hojas EMPLEADO, LOCACION, VENDEDOR, VEHICULO, PLANTA, PUESTO)
 * 3. SS_OPERACION  : 1VfgdfLmmIti2XljXAepw7_xcqv1ZQ73g3f071FQaSVk (Base de datos operativa nueva)
 */

const EYS_CONFIG = {
  IDS: {
    SS_INVENTARIO: "1Av2UJLq7zOi4zlzt7PLYXhOV34qV_SIYUcIPwaF-od4",
    SS_CATALOGOS: "1qFbcn2alEpRsTQWREsElbLSbuPC4LJNdu7KDtdkn2UU",
    SS_OPERACION: "1VfgdfLmmIti2XljXAepw7_xcqv1ZQ73g3f071FQaSVk"
  },
  ZONAS_VALIDAS: [
    "BONFIL MOBILIARIO",
    "BONFIL MANUALIDADES",
    "BONFIL TELAS",
    "BONFIL ILUMINACIÓN",
    "CEDIS",
    "PUERTO",
    "LOCACIONES"
  ],
  BODEGAS_FISICAS: ["BONFIL", "CEDIS", "PUERTO"],
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
 * Función de configuración inicial (ejecutar 1 sola vez en el editor de Apps Script
 * para autorizar permisos de lectura/escritura en las 3 hojas).
 */
function inicializarBaseDeDatosOperativa() {
  const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
  _garantizarHojasOperacion(ssOp);
  const ssInv = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_INVENTARIO);
  const ssCat = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_CATALOGOS);
  return "Conexión exitosa con: " + ssOp.getName() + ", " + ssInv.getName() + " y " + ssCat.getName();
}

/**
 * Endpoint HTTP GET
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = String(params.action || "init").trim();

    if (params.ver_art) {
      return _renderizarFichaEscaneoQR(params.ver_art, params.ods || "");
    }

    if (action === "init") {
      return _jsonResponse(apiObtenerDatosIniciales());
    }

    if (action === "obtenerDatosEtiqueta") {
      return _jsonResponse(apiObtenerDatosEtiqueta(params.ods || "", params.cop || ""));
    }

    if (action === "ping") {
      return _jsonResponse({ exito: true, mensaje: "API Movimientos EyS Z en línea", fecha: _ahoraStr() });
    }

    return _jsonResponse({ exito: false, error: "Acción GET no reconocida: " + action });
  } catch (err) {
    return _jsonResponse({ exito: false, error: err.message, stack: err.stack });
  }
}

/**
 * Endpoint HTTP POST
 */
function doPost(e) {
  try {
    let body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    const action = String(body.action || "").trim();

    switch (action) {
      case "loginUsuario":
        return _jsonResponse(apiLoginUsuario(body.email, body.pin));

      case "solicitarCodigoPin":
        return _jsonResponse(apiSolicitarCodigoPin(body.email));

      case "verificarCodigoPin":
        return _jsonResponse(apiVerificarCodigoPin(body.email, body.codigo));

      case "actualizarPin":
        return _jsonResponse(apiActualizarPin(body.email, body.nuevoPin));

      case "init":
        return _jsonResponse(apiObtenerDatosIniciales());

      case "guardarODS":
        return _jsonResponse(apiGuardarODS(body.payload || body));

      case "cambiarEstadoODS":
        return _jsonResponse(apiCambiarEstadoODS(body.id_ods, body.nuevo_estado));

      case "agregarAZonaManual":
        return _jsonResponse(apiAgregarAZonaManual(body.zona, body.items));

      case "ajustarItemZona":
      case "regresarAlmacenZona":
        return _jsonResponse(apiAjustarItemZona(body.id_reg, body.nueva_cantidad, body.motivo || "REGRESO_A_ALMACEN"));

      case "editarItemZona":
        return _jsonResponse(apiEditarItemZona(body.payload || body));

      case "crearVendedor":
        return _jsonResponse(apiCrearVendedor(body.payload || body));

      case "crearLocacion":
        return _jsonResponse(apiCrearLocacion(body.payload || body));

      case "crearCabeceraViaje":
        return _jsonResponse(apiCrearCabeceraViaje(body.payload || body));

      case "subirItemsAViaje":
        return _jsonResponse(apiSubirItemsAViaje(body.payload || body));

      case "bajarItemDeViaje":
        return _jsonResponse(apiBajarItemDeViaje(body.id_viaje, body.id_det));

      case "darSalidaViaje":
        return _jsonResponse(apiDarSalidaViaje(body.id_viaje, body.usuario_nombre || ""));

      case "crearViaje":
        return _jsonResponse(apiCrearViaje(body.payload || body));

      case "completarViaje":
      case "darEntradaViaje":
      case "tomarItemsDeVehiculo":
        return _jsonResponse(apiCompletarViaje(body.payload || body));

      case "crearReporteViajeLocacion":
        return _jsonResponse(apiCrearReporteViajeLocacion(body.payload || body));

      case "cancelarViaje":
        return _jsonResponse(apiCancelarViaje(body.id_viaje, body.motivo || ""));

      default:
        return _jsonResponse({ exito: false, error: "Acción POST no reconocida: " + action });
    }
  } catch (err) {
    return _jsonResponse({ exito: false, error: err.message });
  }
}

function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Vista HTML rápida cuando se escanea el QR de una etiqueta con la cámara del celular
 */
function _renderizarFichaEscaneoQR(codigoArt, idOds) {
  const cod = String(codigoArt || "").trim().toUpperCase();
  const ods = String(idOds || "").trim();
  const html = `
    <!DOCTYPE html>
    <html lang="es">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>Etiqueta QR · ${cod}</title>
      <style>
        body { font-family: system-ui, sans-serif; background:#0f172a; color:#f8fafc; display:flex; align-items:center; justify-content:center; min-height:100vh; margin:0; padding:16px; }
        .card { background:#1e293b; border:1px solid #334155; border-radius:18px; padding:24px; max-width:380px; width:100%; text-align:center; box-shadow:0 12px 30px rgba(0,0,0,0.4); }
        .badge { display:inline-block; background:#00a9e0; color:#fff; padding:6px 14px; border-radius:999px; font-weight:700; font-size:14px; margin-bottom:12px; }
        h2 { margin:6px 0; font-size:24px; color:#38bdf8; }
        p { color:#94a3b8; font-size:15px; margin:6px 0; }
      </style>
    </head>
    <body>
      <div class="card">
        <div class="badge">ERP ZUÑIGA · ETIQUETA VERIFICADA</div>
        <h2>C.OP: ${cod}</h2>
        ${ods ? `<p><strong>ODS Vinculada:</strong> #${ods}</p>` : ""}
        <p>Escaneo de control de Entradas y Salidas</p>
      </div>
    </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html)
    .setTitle("Etiqueta " + cod)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _ahoraStr() {
  return Utilities.formatDate(new Date(), "America/Cancun", "yyyy-MM-dd HH:mm");
}

function _hoyIso() {
  return Utilities.formatDate(new Date(), "America/Cancun", "yyyy-MM-dd");
}

function _formatearFechaHora(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, "America/Cancun", "yyyy-MM-dd HH:mm");
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(s)) {
    return s.replace("T", " ").substring(0, 16);
  }
  return s;
}

function _garantizarHojasOperacion(ssOp) {
  const definiciones = EYS_CONFIG.HOJAS_OP;
  Object.keys(definiciones).forEach(function (key) {
    const def = definiciones[key];
    let sheet = ssOp.getSheetByName(def.NOMBRE);
    if (!sheet) {
      sheet = ssOp.insertSheet(def.NOMBRE);
      sheet.appendRow(def.COLUMNAS);
      sheet.getRange(1, 1, 1, def.COLUMNAS.length)
        .setFontWeight("bold")
        .setBackground("#00a9e0")
        .setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(def.COLUMNAS);
      sheet.getRange(1, 1, 1, def.COLUMNAS.length)
        .setFontWeight("bold")
        .setBackground("#00a9e0")
        .setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }
  });
}

function _obtenerBodegaBaseDeZona(zona) {
  const z = String(zona || "").trim().toUpperCase();
  if (z.indexOf("BONFIL") === 0) return "BONFIL";
  if (z === "CEDIS") return "CEDIS";
  if (z === "PUERTO") return "PUERTO";
  if (z === "LOCACIONES") return "LOCACIONES";
  return z;
}

function _resolverZonaBonfilPorTipo(tipoArt, zonaPreferida) {
  const zp = String(zonaPreferida || "").trim().toUpperCase();
  if (zp.indexOf("BONFIL ") === 0) return zp;
  const t = String(tipoArt || "").trim().toUpperCase();
  if (t.indexOf("MANU") !== -1 || t.indexOf("DECO") !== -1) return "BONFIL MANUALIDADES";
  if (t.indexOf("TELA") !== -1 || t.indexOf("MANTEL") !== -1 || t.indexOf("LINO") !== -1) return "BONFIL TELAS";
  if (t.indexOf("ILUM") !== -1 || t.indexOf("LUZ") !== -1 || t.indexOf("ELEC") !== -1 || t.indexOf("AUDIO") !== -1) return "BONFIL ILUMINACIÓN";
  return "BONFIL MOBILIARIO";
}

function _normalizarUrlFotoDrive(rawFoto) {
  const s = String(rawFoto || "").trim();
  if (!s) return "";
  if (s.indexOf("http") === 0) {
    const m = s.match(/[-\w]{25,}/);
    if (m && (s.indexOf("drive.google.com") !== -1 || s.indexOf("docs.google.com") !== -1)) {
      return "https:" + "/" + "/drive.google.com/thumbnail?id=" + m[0] + "&sz=w200";
    }
    return s;
  }
  if (/^[-\w]{20,}$/.test(s)) {
    return "https:" + "/" + "/drive.google.com/thumbnail?id=" + s + "&sz=w200";
  }
  return "";
}

/**
 * ============================================================================
 * 0. AUTENTICACIÓN DE USUARIOS CONTRA LA HOJA "EMPLEADO" DEL ERP ZUÑIGA
 * ============================================================================
 */
function _generarHashPIN(pin) {
  if (!pin) return "";
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin).trim());
  return digest.map(function (byte) {
    const v = (byte < 0) ? byte + 256 : byte;
    return v.toString(16).padStart(2, "0");
  }).join("");
}

function apiLoginUsuario(email, pin) {
  const emailBuscado = String(email || "").trim().toLowerCase();
  const pinIngresado = String(pin || "").trim();

  if (!emailBuscado || !pinIngresado) {
    return { exito: false, message: "Ingresa tu correo y tu PIN de acceso." };
  }

  const hashIngresado = _generarHashPIN(pinIngresado);
  const ssCat = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_CATALOGOS);
  const sheetEmp = ssCat.getSheetByName("EMPLEADO");

  if (sheetEmp && sheetEmp.getLastRow() >= 2) {
    const data = sheetEmp.getDataRange().getValues();
    const headers = data[0].map(function (h) { return String(h || "").trim().toLowerCase(); });

    const idxId = headers.indexOf("id_empleado") !== -1 ? headers.indexOf("id_empleado") : 0;
    const idxNom = headers.indexOf("nombre") !== -1 ? headers.indexOf("nombre") : 1;
    const idxEmail = headers.indexOf("email") !== -1 ? headers.indexOf("email") : 2;
    const idxPin = headers.indexOf("pin") !== -1 ? headers.indexOf("pin") : 3;
    const idxDepto = headers.indexOf("departamento") !== -1 ? headers.indexOf("departamento") : 4;
    const idxRol = headers.indexOf("rol") !== -1 ? headers.indexOf("rol") : 5;
    const idxNivel = headers.indexOf("id_nivel") !== -1 ? headers.indexOf("id_nivel") : 6;
    const idxFoto = headers.indexOf("foto_emple_url") !== -1 ? headers.indexOf("foto_emple_url") : 10;
    const idxBode = headers.indexOf("id_bode") !== -1 ? headers.indexOf("id_bode") : 14;

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      const emailHoja = String(fila[idxEmail] || "").trim().toLowerCase();
      if (!emailHoja) continue;

      if (emailHoja === emailBuscado) {
        const pinHoja = String(fila[idxPin] || "").trim();
        if (pinHoja === hashIngresado || pinHoja === pinIngresado) {
          const fotoUrl = _normalizarUrlFotoDrive(fila[idxFoto]) || "https:" + "/" + "/cdn-icons-png.flaticon.com/512/3135/3135715.png";
          return {
            exito: true,
            user: {
              id_empleado: String(fila[idxId] || ""),
              nombre: String(fila[idxNom] || "Usuario ERP").trim(),
              email: emailHoja,
              departamento: String(fila[idxDepto] || "OPERACIÓN").trim(),
              rol: String(fila[idxRol] || "OPERATIVO").trim(),
              id_nivel: Number(fila[idxNivel]) || 1,
              foto: fotoUrl,
              id_bode: String(fila[idxBode] || "").trim()
            }
          };
        }
        return { exito: false, message: "El PIN ingresado es incorrecto." };
      }
    }
  }

  return { exito: false, message: "El correo ingresado no está registrado en la hoja EMPLEADO del ERP." };
}

function apiSolicitarCodigoPin(email) {
  const emailBuscado = String(email || "").trim().toLowerCase();
  if (!emailBuscado) return { exito: false, message: "Ingresa tu correo electrónico." };

  const ssCat = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_CATALOGOS);
  const sheetEmp = ssCat.getSheetByName("EMPLEADO");
  if (!sheetEmp) return { exito: false, message: "No se encontró la hoja EMPLEADO." };

  const data = sheetEmp.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || "").trim().toLowerCase(); });
  const idxEmail = headers.indexOf("email") !== -1 ? headers.indexOf("email") : 2;
  const idxCod = headers.indexOf("codi_temp") !== -1 ? headers.indexOf("codi_temp") : 8;
  const idxExp = headers.indexOf("expira_codi_temp") !== -1 ? headers.indexOf("expira_codi_temp") : 9;

  for (let i = 1; i < data.length; i++) {
    const emailHoja = String(data[i][idxEmail] || "").trim().toLowerCase();
    if (emailHoja === emailBuscado) {
      const codigo = String(Math.floor(100000 + Math.random() * 900000));
      const fechaExpira = new Date();
      fechaExpira.setHours(fechaExpira.getHours() + 1);

      sheetEmp.getRange(i + 1, idxCod + 1).setValue(codigo);
      sheetEmp.getRange(i + 1, idxExp + 1).setValue(fechaExpira);

      MailApp.sendEmail({
        to: emailBuscado,
        subject: "🔐 Código de Verificación - ERP ZUÑIGA",
        htmlBody: `<div style="font-family:sans-serif; padding:20px; border:1px solid #eee; border-radius:10px;">
          <h2>Recuperación de Acceso - ERP ZUÑIGA</h2>
          <p>Tu código de verificación es:</p>
          <h1 style="letter-spacing:5px; color:#00a9e0; background:#f4f4f4; text-align:center; padding:10px;">${codigo}</h1>
          <p>Este código expira en 1 hora.</p>
        </div>`
      });
      return { exito: true };
    }
  }
  return { exito: false, message: "Correo no encontrado en EMPLEADO." };
}

function apiVerificarCodigoPin(email, codigoIngresado) {
  const emailBuscado = String(email || "").trim().toLowerCase();
  const cod = String(codigoIngresado || "").trim();
  const ssCat = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_CATALOGOS);
  const sheetEmp = ssCat.getSheetByName("EMPLEADO");
  if (!sheetEmp) return { exito: false, message: "Hoja EMPLEADO no disponible." };

  const data = sheetEmp.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || "").trim().toLowerCase(); });
  const idxEmail = headers.indexOf("email") !== -1 ? headers.indexOf("email") : 2;
  const idxCod = headers.indexOf("codi_temp") !== -1 ? headers.indexOf("codi_temp") : 8;
  const idxExp = headers.indexOf("expira_codi_temp") !== -1 ? headers.indexOf("expira_codi_temp") : 9;

  for (let i = 1; i < data.length; i++) {
    const emailHoja = String(data[i][idxEmail] || "").trim().toLowerCase();
    if (emailHoja === emailBuscado) {
      const codGuardado = String(data[i][idxCod] || "").trim();
      const expira = data[i][idxExp] ? new Date(data[i][idxExp]).getTime() : 0;
      if (codGuardado && codGuardado === cod) {
        if (expira && new Date().getTime() > expira) {
          return { exito: false, message: "El código ha expirado." };
        }
        return { exito: true };
      }
      return { exito: false, message: "Código incorrecto." };
    }
  }
  return { exito: false, message: "Correo no encontrado." };
}

function apiActualizarPin(email, nuevoPin) {
  const emailBuscado = String(email || "").trim().toLowerCase();
  const pinLimpio = String(nuevoPin || "").trim();
  if (pinLimpio.length !== 4) return { exito: false, message: "El PIN debe tener exactamente 4 dígitos." };

  const ssCat = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_CATALOGOS);
  const sheetEmp = ssCat.getSheetByName("EMPLEADO");
  const data = sheetEmp.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || "").trim().toLowerCase(); });
  const idxEmail = headers.indexOf("email") !== -1 ? headers.indexOf("email") : 2;
  const idxPin = headers.indexOf("pin") !== -1 ? headers.indexOf("pin") : 3;
  const idxCod = headers.indexOf("codi_temp") !== -1 ? headers.indexOf("codi_temp") : 8;
  const idxExp = headers.indexOf("expira_codi_temp") !== -1 ? headers.indexOf("expira_codi_temp") : 9;

  for (let i = 1; i < data.length; i++) {
    const emailHoja = String(data[i][idxEmail] || "").trim().toLowerCase();
    if (emailHoja === emailBuscado) {
      const hash = _generarHashPIN(pinLimpio);
      sheetEmp.getRange(i + 1, idxPin + 1).setValue(hash);
      sheetEmp.getRange(i + 1, idxCod + 1).clearContent();
      sheetEmp.getRange(i + 1, idxExp + 1).clearContent();
      SpreadsheetApp.flush();
      return apiLoginUsuario(emailBuscado, pinLimpio);
    }
  }
  return { exito: false, message: "Usuario no encontrado." };
}

/**
 * ============================================================================
 * 1. CARGA INICIAL UNIFICADA (CATÁLOGOS + ESTADO OPERATIVO)
 * ============================================================================
 */
function apiObtenerDatosIniciales() {
  const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
  _garantizarHojasOperacion(ssOp);

  const catalogos = _leerCatalogosExternos();
  const estadoOp = _leerEstadoOperativoCompleto(ssOp);

  return {
    exito: true,
    catalogos: catalogos,
    ods: estadoOp.ods,
    zonas: estadoOp.zonas,
    viajes: estadoOp.viajes,
    historial: estadoOp.historial,
    zonas_validas: EYS_CONFIG.ZONAS_VALIDAS,
    bodegas_fisicas: EYS_CONFIG.BODEGAS_FISICAS
  };
}

function _leerCatalogosExternos() {
  const resultado = {
    articulos: [],
    locaciones: [],
    regiones: [],
    vendedores: [],
    vehiculos: [],
    choferes: []
  };

  // A. Leer INVENTARIO (encabezados en fila 3, datos desde fila 4)
  try {
    const ssInv = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_INVENTARIO);
    const sheetInv = ssInv.getSheetByName("INVENTARIO");
    if (sheetInv && sheetInv.getLastRow() >= 4) {
      const dataInv = sheetInv.getDataRange().getValues();
      for (let i = 3; i < dataInv.length; i++) {
        const fila = dataInv[i];
        const codigo = String(fila[1] || "").trim();
        const nombre = String(fila[2] || "").trim();
        if (!codigo && !nombre) continue;
        const estadoFila = String(fila[15] || "").trim().toUpperCase();
        if (estadoFila === "INACTIVO" || estadoFila === "BAJA") continue;

        resultado.articulos.push({
          codigo: codigo,
          nombre: nombre,
          tipo: String(fila[3] || "MOBILIARIO").trim(),
          foto: _normalizarUrlFotoDrive(fila[4]),
          bodega_origen: String(fila[11] || "").trim()
        });
      }
    }
  } catch (eInv) {
    Logger.log("Aviso leyendo INVENTARIO: " + eInv.message);
  }

  // B. Leer REGION, LOCACION, VENDEDOR, VEHICULO, PLANTA y PUESTO
  try {
    const ssCat = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_CATALOGOS);

    const mapaRegionesNombre = {};
    const sheetReg = ssCat.getSheetByName("REGION");
    if (sheetReg && sheetReg.getLastRow() > 1) {
      const dReg = sheetReg.getDataRange().getValues();
      for (let r = 1; r < dReg.length; r++) {
        const idReg = String(dReg[r][0] || "").trim();
        const nomReg = String(dReg[r][1] || "").trim();
        const idSedeR = String(dReg[r][2] || "1").trim();
        if (nomReg) {
          mapaRegionesNombre[idReg] = nomReg;
          resultado.regiones.push({
            id: idReg || nomReg,
            nombre: nomReg,
            id_sede: idSedeR
          });
        }
      }
    }
    if (resultado.regiones.length === 0) {
      resultado.regiones = [
        { id: "1", nombre: "Zona Hotelera", id_sede: "1" },
        { id: "2", nombre: "Costa Mujeres", id_sede: "1" },
        { id: "3", nombre: "Puerto Aventuras", id_sede: "1" },
        { id: "4", nombre: "Zona Cabos", id_sede: "2" }
      ];
    }

    const sheetLoc = ssCat.getSheetByName("LOCACION");
    if (sheetLoc && sheetLoc.getLastRow() > 1) {
      const dLoc = sheetLoc.getDataRange().getValues();
      for (let i = 1; i < dLoc.length; i++) {
        const idLoc = String(dLoc[i][0] || "").trim();
        const nomLoc = String(dLoc[i][1] || "").trim();
        const idRegLoc = String(dLoc[i][2] || "").trim();
        if (nomLoc) {
          resultado.locaciones.push({
            id: idLoc || nomLoc,
            nombre: nomLoc,
            id_region: idRegLoc,
            region_nombre: mapaRegionesNombre[idRegLoc] || idRegLoc || "",
            foto: _normalizarUrlFotoDrive(dLoc[i][3])
          });
        }
      }
      resultado.locaciones.sort(function (a, b) { return a.nombre.localeCompare(b.nombre); });
    }

    const sheetVen = ssCat.getSheetByName("VENDEDOR");
    if (sheetVen && sheetVen.getLastRow() > 1) {
      const dVen = sheetVen.getDataRange().getValues();
      for (let i = 1; i < dVen.length; i++) {
        const idVen = String(dVen[i][0] || "").trim();
        const nomVen = String(dVen[i][1] || "").trim();
        const sedeVen = String(dVen[i][2] || "1").trim();
        if (nomVen) {
          resultado.vendedores.push({ id: idVen || nomVen, nombre: nomVen, id_sede: sedeVen });
        }
      }
      resultado.vendedores.sort(function (a, b) { return a.nombre.localeCompare(b.nombre); });
    }

    const sheetVeh = ssCat.getSheetByName("VEHICULO");
    if (sheetVeh && sheetVeh.getLastRow() > 1) {
      const dVeh = sheetVeh.getDataRange().getValues();
      for (let i = 1; i < dVeh.length; i++) {
        const idVeh = String(dVeh[i][0] || "").trim();
        const nomVeh = String(dVeh[i][1] || "").trim();
        if (nomVeh) {
          resultado.vehiculos.push({
            id: idVeh || nomVeh,
            nombre: nomVeh,
            foto: _normalizarUrlFotoDrive(dVeh[i][2])
          });
        }
      }
    }

    const mapaPuestos = {};
    const sheetPuesto = ssCat.getSheetByName("PUESTO");
    if (sheetPuesto && sheetPuesto.getLastRow() > 1) {
      const dPue = sheetPuesto.getDataRange().getValues();
      for (let p = 1; p < dPue.length; p++) {
        const idP = String(dPue[p][0] || "").trim();
        const nomP = String(dPue[p][1] || "").trim();
        if (idP) mapaPuestos[idP] = nomP;
      }
    }

    const sheetPlanta = ssCat.getSheetByName("PLANTA");
    if (sheetPlanta && sheetPlanta.getLastRow() > 1) {
      const dPla = sheetPlanta.getDataRange().getValues();
      const headersPla = dPla[0].map(function (h) { return String(h || "").trim().toLowerCase(); });
      const idxIdEmp = headersPla.indexOf("id_empleado") !== -1 ? headersPla.indexOf("id_empleado") : 0;
      const idxNom = headersPla.indexOf("nombre") !== -1 ? headersPla.indexOf("nombre") : 1;
      const idxApe = headersPla.indexOf("apellido") !== -1 ? headersPla.indexOf("apellido") : 2;
      const idxFotoP = headersPla.indexOf("foto_id") !== -1 ? headersPla.indexOf("foto_id") : 4;
      const idxPue = headersPla.indexOf("id_puesto") !== -1 ? headersPla.indexOf("id_puesto") : 14;
      const idxAct = headersPla.indexOf("activo") !== -1 ? headersPla.indexOf("activo") : 18;

      const listaChoferesPuesto = [];
      const listaPlantaGeneral = [];

      for (let c = 1; c < dPla.length; c++) {
        const row = dPla[c];
        const activoVal = row[idxAct];
        if (activoVal === 0 || String(activoVal).trim() === "0" || String(activoVal).toUpperCase() === "FALSE") continue;

        const idEmp = String(row[idxIdEmp] || "").trim();
        const nomCompleto = (String(row[idxNom] || "").trim() + " " + String(row[idxApe] || "").trim()).trim();
        if (!nomCompleto) continue;

        const idPuesto = String(row[idxPue] || "").trim();
        const nomPuesto = mapaPuestos[idPuesto] || idPuesto || "";
        const itemEmp = {
          id: idEmp || nomCompleto,
          nombre: nomCompleto,
          puesto: nomPuesto,
          foto: _normalizarUrlFotoDrive(row[idxFotoP]),
          es_chofer: nomPuesto.toUpperCase().indexOf("CHOFER") !== -1
        };

        listaPlantaGeneral.push(itemEmp);
        if (itemEmp.es_chofer) {
          listaChoferesPuesto.push(itemEmp);
        }
      }

      resultado.choferes = (listaChoferesPuesto.length > 0 ? listaChoferesPuesto.concat(listaPlantaGeneral.filter(function (x) { return !x.es_chofer; })) : listaPlantaGeneral);
    }
  } catch (eCat) {
    Logger.log("Aviso leyendo Catálogos ERP: " + eCat.message);
  }

  return resultado;
}

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
      listaODS.push({
        id_ods: idOds,
        nombre_evento: String(r[1] || "").trim(),
        id_vendedor: String(r[2] || "").trim(),
        vendedor_nombre: String(r[3] || "").trim(),
        id_locacion: String(r[4] || "").trim(),
        locacion_nombre: String(r[5] || "").trim(),
        fecha_montaje: _formatearFechaHora(r[6]),
        estado: String(r[7] || "ACTIVA").trim().toUpperCase(),
        fecha_creacion: _formatearFechaHora(r[8]),
        ultima_mod: _formatearFechaHora(r[9])
      });
    }
  }
  listaODS.sort(function (a, b) {
    return String(a.fecha_montaje || "9999").localeCompare(String(b.fecha_montaje || "9999"));
  });

  // 2. ZONAS_ALISTAMIENTO
  const listaZonas = [];
  const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);
  if (sheetZonas && sheetZonas.getLastRow() > 1) {
    const dZ = sheetZonas.getDataRange().getValues();
    for (let j = 1; j < dZ.length; j++) {
      const rz = dZ[j];
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
        estado_item: String(rd[15] || "EN_CARGA").trim()
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
      let estViaje = String(rv[9] || "EN_CARGA").trim().toUpperCase();
      if (estViaje === "EN_CURSO") estViaje = "EN_RUTA"; // Compatibilidad con viajes anteriores

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
        estado: estViaje,
        fecha_salida: _formatearFechaHora(rv[10]),
        fecha_cierre: _formatearFechaHora(rv[11]),
        observaciones_cierre: String(rv[12] || "").trim(),
        items: detallesPorViaje[idV] || []
      });
    }
  }
  listaViajes.sort(function (a, b) {
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
  listaHistorial.sort(function (a, b) {
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
 * 2. CREAR O EDITAR ODS (1. PLANEACIÓN)
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

    (items || []).forEach(function (it) {
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
 * 5A-BIS. EDITAR UN REGISTRO EN ZONA DE ALISTAMIENTO (CORREGIR ODS, CANTIDAD O ARTÍCULO)
 */
function apiEditarItemZona(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);
    const sheetHist = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.HISTORIAL.NOMBRE);
    const data = sheetZonas.getDataRange().getValues();

    const idReg = String(payload.id_reg || "").trim();
    const nuevoCod = String(payload.codigo_art || "").trim().toUpperCase();
    const nuevoNom = String(payload.nombre_art || "").trim();
    const nuevoTipo = String(payload.tipo_art || "MOBILIARIO").trim();
    const nuevaCant = Math.max(1, Number(payload.cantidad) || 1);
    const nuevaOds = String(payload.id_ods || "").trim();
    const nuevaFechaMont = _formatearFechaHora(payload.fecha_montaje);
    const nuevaLoc = String(payload.locacion_nombre || "").trim();
    const nuevoEv = String(payload.nombre_evento || "").trim();
    const ahora = _ahoraStr();

    if (!idReg || !nuevoCod || !nuevoNom || !nuevaOds) {
      return { exito: false, error: "Completa Código, Artículo, ODS y Cantidad." };
    }

    let filaEditar = -1;
    let zonaActual = "";
    let bodegaBase = "";
    let antResumen = "";

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === idReg && data[i][12] !== 0 && String(data[i][12]) !== "0") {
        filaEditar = i + 1;
        zonaActual = String(data[i][1] || "").trim().toUpperCase();
        bodegaBase = String(data[i][2] || "").trim() || _obtenerBodegaBaseDeZona(zonaActual);
        antResumen = "ODS #" + data[i][7] + " · " + data[i][3] + " x " + data[i][6];
        break;
      }
    }

    if (filaEditar === -1) {
      return { exito: false, error: "No se encontró el registro en la zona de alistamiento." };
    }

    // Verificar si ya existe otro renglón activo en la misma zona con el mismo nuevoCod + nuevaOds
    let filaDuplicada = -1;
    let cantDuplicada = 0;
    for (let j = 1; j < data.length; j++) {
      if (j + 1 === filaEditar) continue;
      if (data[j][12] === 0 || String(data[j][12]) === "0") continue;
      const zRow = String(data[j][1] || "").trim().toUpperCase();
      const cRow = String(data[j][3] || "").trim().toUpperCase();
      const oRow = String(data[j][7] || "").trim();
      if (zRow === zonaActual && cRow === nuevoCod && oRow === nuevaOds) {
        filaDuplicada = j + 1;
        cantDuplicada = Number(data[j][6]) || 0;
        break;
      }
    }

    if (filaDuplicada > 0) {
      sheetZonas.getRange(filaDuplicada, 7).setValue(cantDuplicada + nuevaCant);
      sheetZonas.getRange(filaDuplicada, 8, 1, 5).setValues([[nuevaOds, nuevaFechaMont, nuevaLoc, nuevoEv, ahora]]);
      sheetZonas.getRange(filaEditar, 7).setValue(0);
      sheetZonas.getRange(filaEditar, 12).setValue(ahora);
      sheetZonas.getRange(filaEditar, 13).setValue(0);
    } else {
      sheetZonas.getRange(filaEditar, 4, 1, 9).setValues([[
        nuevoCod,
        nuevoNom,
        nuevoTipo,
        nuevaCant,
        nuevaOds,
        nuevaFechaMont,
        nuevaLoc,
        nuevoEv,
        ahora
      ]]);
    }

    if (sheetHist) {
      _registrarHistorial(sheetHist, {
        fecha_hora: ahora,
        id_viaje: "EDICION-ZONA",
        tipo_movimiento: "EDICION_ZONA_ALISTAMIENTO",
        bodega_origen: bodegaBase,
        zona_origen: zonaActual,
        bodega_destino: bodegaBase,
        zona_destino: zonaActual,
        id_ods: nuevaOds,
        nombre_evento: nuevoEv,
        locacion_nombre: nuevaLoc,
        vendedor_nombre: "",
        fecha_ods: nuevaFechaMont,
        codigo_art: nuevoCod,
        nombre_art: nuevoNom,
        cantidad: nuevaCant,
        vehiculo: "",
        chofer: "",
        novedad: "",
        resumen_texto: "✏️ Editado en " + zonaActual + " (Antes: " + antResumen + " -> Ahora: ODS #" + nuevaOds + " · " + nuevoCod + " x " + nuevaCant + ")"
      });
    }

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
 * 5. AJUSTAR O RETIRAR ARTÍCULO DE ZONA DE ALISTAMIENTO (REGRESAR A ALMACÉN)
 */
function apiAjustarItemZona(idReg, nuevaCantidad, motivo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);
    const sheetHist = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.HISTORIAL.NOMBRE);
    const data = sheetZonas.getDataRange().getValues();
    const target = String(idReg || "").trim();
    const cantNum = Math.max(0, Number(nuevaCantidad) || 0);
    const ahora = _ahoraStr();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === target && data[i][12] !== 0 && String(data[i][12]) !== "0") {
        const cantAnterior = Number(data[i][6]) || 0;
        const cantRetirada = Math.max(0, cantAnterior - cantNum);
        const zonaActual = String(data[i][1] || "").trim();
        const bodegaBase = String(data[i][2] || "").trim() || _obtenerBodegaBaseDeZona(zonaActual);
        const codArt = String(data[i][3] || "").trim();
        const nomArt = String(data[i][4] || "").trim();
        const idOds = String(data[i][7] || "").trim();
        const fMont = _formatearFechaHora(data[i][8]);
        const locNom = String(data[i][9] || "").trim();
        const nomEv = String(data[i][10] || "").trim();

        sheetZonas.getRange(i + 1, 7).setValue(cantNum);
        sheetZonas.getRange(i + 1, 12).setValue(ahora);
        if (cantNum === 0) {
          sheetZonas.getRange(i + 1, 13).setValue(0);
        }

        if (cantRetirada > 0 && sheetHist) {
          _registrarHistorial(sheetHist, {
            fecha_hora: ahora,
            id_viaje: "REGRESO-ALM",
            tipo_movimiento: motivo || "REGRESO_A_ALMACEN",
            bodega_origen: bodegaBase,
            zona_origen: zonaActual,
            bodega_destino: bodegaBase,
            zona_destino: "ALMACÉN",
            id_ods: idOds,
            nombre_evento: nomEv,
            locacion_nombre: locNom,
            vendedor_nombre: "",
            fecha_ods: fMont,
            codigo_art: codArt,
            nombre_art: nomArt,
            cantidad: cantRetirada,
            vehiculo: "",
            chofer: "",
            novedad: "",
            resumen_texto: "↩️ ODS " + idOds + " x " + cantRetirada + " (" + nomArt + ") regresó de zona " + zonaActual + " a Almacén"
          });
        }
        break;
      }
    }

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
 * 5B. CREAR NUEVO VENDEDOR EN SS_CATALOGOS (HOJA VENDEDOR) ASIGNADO A LA SEDE ACTIVA
 */
function apiCrearVendedor(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const nombre = String(payload.nombre || "").trim();
    const idSede = Number(payload.id_sede) || 1;
    if (!nombre) {
      return { exito: false, error: "Ingresa el nombre del vendedor." };
    }

    const ssCat = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_CATALOGOS);
    let sheetVen = ssCat.getSheetByName("VENDEDOR");
    if (!sheetVen) {
      sheetVen = ssCat.insertSheet("VENDEDOR");
      sheetVen.appendRow(["id_vent", "nombre", "id_sede"]);
    }

    const data = sheetVen.getDataRange().getValues();
    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
      const idNum = Number(data[i][0]) || 0;
      if (idNum > maxId) maxId = idNum;
      const nomExist = String(data[i][1] || "").trim().toLowerCase();
      if (nomExist === nombre.toLowerCase()) {
        return { exito: false, error: "Ya existe un vendedor con el nombre '" + nombre + "'." };
      }
    }

    const nuevoId = maxId + 1;
    sheetVen.appendRow([nuevoId, nombre, idSede]);
    SpreadsheetApp.flush();

    const catalogos = _leerCatalogosExternos();
    return {
      exito: true,
      vendedor: { id: String(nuevoId), nombre: nombre, id_sede: String(idSede) },
      vendedores: catalogos.vendedores
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 5C. CREAR NUEVA LOCACIÓN EN SS_CATALOGOS (HOJA LOCACION) CON SU REGIÓN
 */
function apiCrearLocacion(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const nombre = String(payload.nombre || "").trim();
    const idRegion = String(payload.id_region || "1").trim();
    if (!nombre) {
      return { exito: false, error: "Ingresa el nombre de la locación." };
    }

    const ssCat = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_CATALOGOS);
    let sheetLoc = ssCat.getSheetByName("LOCACION");
    if (!sheetLoc) {
      sheetLoc = ssCat.insertSheet("LOCACION");
      sheetLoc.appendRow(["id_locacion", "nombre", "id_region", "foto"]);
    }

    const data = sheetLoc.getDataRange().getValues();
    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
      const idNum = Number(data[i][0]) || 0;
      if (idNum > maxId) maxId = idNum;
      const nomExist = String(data[i][1] || "").trim().toLowerCase();
      if (nomExist === nombre.toLowerCase()) {
        return { exito: false, error: "Ya existe una locación con el nombre '" + nombre + "'." };
      }
    }

    const nuevoId = maxId + 1;
    sheetLoc.appendRow([nuevoId, nombre, Number(idRegion) || idRegion, ""]);
    SpreadsheetApp.flush();

    const catalogos = _leerCatalogosExternos();
    return {
      exito: true,
      locacion: { id: String(nuevoId), nombre: nombre, id_region: idRegion },
      locaciones: catalogos.locaciones
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * 6. FLUJO COLABORATIVO DE VIAJES (IGUAL AL ERP ZUÑIGA):
 *    PASO A: Crear Cabecera del Viaje en Andén (EN_CARGA)
 *    PASO B: Subir / Bajar cosas al Camión entre varias personas (EN_CARGA)
 *    PASO C: Dar Salida al Camión (pasa a EN_RUTA; si es Evento entra a LOCACIONES)
 *    PASO D: Dar Entrada / Recibir en otra bodega o regreso de evento (COMPLETADO)
 * ============================================================================
 */
function apiCrearCabeceraViaje(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetViajes = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES.NOMBRE);

    const tipoViaje = String(payload.tipo_viaje || "TRASLADO_BODEGA").trim().toUpperCase();
    const bodegaOrigen = String(payload.bodega_origen || "").trim().toUpperCase();
    const destino = String(payload.destino || "").trim();
    const vehiculo = String(payload.vehiculo || "").trim();
    const chofer = String(payload.chofer || "").trim();
    const idOdsGeneral = String(payload.id_ods_general || "").trim();
    const nomEvGeneral = String(payload.nombre_evento_general || "").trim();
    const fMontGeneral = _formatearFechaHora(payload.fecha_montaje_general);

    if (!bodegaOrigen || !destino || !vehiculo || !chofer) {
      return { exito: false, error: "Completa Bodega Origen, Destino, Unidad y Chofer." };
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
      "EN_CARGA",
      ahora,
      "",
      "",
      1
    ]);

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
 * PASO B.1: Subir artículos al camión mientras está en Andén (EN_CARGA)
 */
function apiSubirItemsAViaje(payload) {
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
    const items = Array.isArray(payload.items) ? payload.items : [];
    const ahora = _ahoraStr();

    if (!idViaje || items.length === 0) {
      return { exito: false, error: "Selecciona al menos 1 artículo para subir al camión." };
    }

    const dViajes = sheetViajes.getDataRange().getValues();
    let infoViaje = null;
    for (let v = 1; v < dViajes.length; v++) {
      if (String(dViajes[v][0] || "").trim() === idViaje && dViajes[v][13] !== 0 && String(dViajes[v][13]) !== "0") {
        infoViaje = {
          tipo_viaje: String(dViajes[v][1] || "TRASLADO_BODEGA").trim().toUpperCase(),
          bodega_origen: String(dViajes[v][2] || "").trim().toUpperCase(),
          destino: String(dViajes[v][3] || "").trim(),
          id_ods_general: String(dViajes[v][4] || "").trim(),
          nombre_evento_general: String(dViajes[v][5] || "").trim(),
          fecha_montaje_general: _formatearFechaHora(dViajes[v][6]),
          vehiculo: String(dViajes[v][7] || "").trim(),
          chofer: String(dViajes[v][8] || "").trim(),
          estado: String(dViajes[v][9] || "EN_CARGA").trim().toUpperCase()
        };
        break;
      }
    }

    if (!infoViaje) {
      return { exito: false, error: "No se encontró el viaje " + idViaje };
    }
    if (infoViaje.estado !== "EN_CARGA" && infoViaje.estado !== "EN_RUTA") {
      return { exito: false, error: "Solo se pueden subir cosas cuando el viaje está en Andén o en Escala (En Ruta)." };
    }
    const estadoItemNuevo = infoViaje.estado === "EN_RUTA" ? "EN_RUTA" : "EN_CARGA";

    items.forEach(function (it) {
      const cant = Number(it.cantidad) || 0;
      if (!it.codigo_art || cant <= 0) return;

      const idOdsItem = String(it.id_ods || infoViaje.id_ods_general || "").trim();
      const nomEvItem = String(it.nombre_evento || infoViaje.nombre_evento_general || "").trim();
      const locItem = infoViaje.tipo_viaje === "ENVIO_EVENTO" ? infoViaje.destino : String(it.locacion_nombre || "").trim();
      const fMontItem = _formatearFechaHora(it.fecha_montaje || infoViaje.fecha_montaje_general);
      const tipoArt = String(it.tipo_art || "MOBILIARIO").trim();

      const zonaOrigenItem = it.zona_origen
        ? String(it.zona_origen).trim().toUpperCase()
        : (infoViaje.bodega_origen === "BONFIL" ? _resolverZonaBonfilPorTipo(tipoArt, "") : infoViaje.bodega_origen);

      let zonaDestinoItem = infoViaje.destino;
      if (infoViaje.tipo_viaje === "TRASLADO_BODEGA") {
        zonaDestinoItem = infoViaje.destino === "BONFIL"
          ? _resolverZonaBonfilPorTipo(tipoArt, it.zona_destino)
          : infoViaje.destino;
      } else {
        zonaDestinoItem = "LOCACIONES";
      }

      // Descontar de la Zona de Alistamiento Origen al subirlo al camión (salvo si es un EXTRA libre que no proviene de stock de zona)
      if (!it.es_extra) {
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
      }

      const obsItem = String(it.observaciones || it.novedad || "").trim();
      const idDet = idViaje + "-D" + new Date().getTime() + "-" + Math.floor(Math.random() * 900 + 100);
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
        obsItem,
        estadoItemNuevo,
        1
      ]);

      _registrarHistorial(sheetHist, {
        fecha_hora: ahora,
        id_viaje: idViaje,
        tipo_movimiento: "CARGA_ANDEN",
        bodega_origen: infoViaje.bodega_origen,
        zona_origen: zonaOrigenItem,
        bodega_destino: infoViaje.tipo_viaje === "ENVIO_EVENTO" ? "LOCACIONES" : infoViaje.destino,
        zona_destino: zonaDestinoItem,
        id_ods: idOdsItem,
        nombre_evento: nomEvItem,
        locacion_nombre: locItem,
        vendedor_nombre: it.vendedor_nombre || "",
        fecha_ods: fMontItem,
        codigo_art: it.codigo_art,
        nombre_art: it.nombre_art,
        cantidad: cant,
        vehiculo: infoViaje.vehiculo,
        chofer: infoViaje.chofer,
        novedad: "",
        resumen_texto: "ODS " + idOdsItem + " x " + cant + " (" + it.nombre_art + ") se subió al camión " + idViaje + " en Andén (" + zonaOrigenItem + ")"
      });
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
 * PASO B.2: Bajar un artículo del camión mientras sigue en Andén (EN_CARGA)
 */
function apiBajarItemDeViaje(idViaje, idDet) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetDet = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES_DETALLE.NOMBRE);
    const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);

    const dDet = sheetDet.getDataRange().getValues();
    for (let d = 1; d < dDet.length; d++) {
      const rd = dDet[d];
      if (String(rd[0] || "").trim() === String(idDet).trim() && String(rd[1] || "").trim() === String(idViaje).trim() && rd[16] !== 0) {
        const codArt = String(rd[2] || "").trim();
        const nomArt = String(rd[3] || "").trim();
        const tipoArt = String(rd[4] || "").trim();
        const zonaOrigen = String(rd[5] || "").trim();
        const idOds = String(rd[7] || "").trim();
        const nomEv = String(rd[8] || "").trim();
        const locNom = String(rd[9] || "").trim();
        const fMont = _formatearFechaHora(rd[10]);
        const cantEnv = Number(rd[11]) || 0;

        // Devolver a la Zona de Alistamiento Origen (salvo si fue un EXTRA sin código de catálogo)
        if (codArt !== "EXTRA" && codArt !== "S/C") {
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
        }

        sheetDet.getRange(d + 1, 16).setValue("BAJADO_ANDEN");
        sheetDet.getRange(d + 1, 17).setValue(0);
        break;
      }
    }

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
 * PASO C: DAR SALIDA AL CAMIÓN (Pasa de EN_CARGA -> EN_RUTA)
 */
function apiDarSalidaViaje(idViaje, usuarioNombre) {
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
          tipo_viaje: String(dViajes[v][1] || "TRASLADO_BODEGA").trim().toUpperCase(),
          bodega_origen: String(dViajes[v][2] || "").trim(),
          destino: String(dViajes[v][3] || "").trim(),
          vehiculo: String(dViajes[v][7] || "").trim(),
          chofer: String(dViajes[v][8] || "").trim(),
          estado: String(dViajes[v][9] || "EN_CARGA").trim().toUpperCase()
        };
        break;
      }
    }

    if (filaViaje === -1 || !infoViaje) {
      return { exito: false, error: "No se encontró el viaje " + target };
    }
    if (infoViaje.estado !== "EN_CARGA") {
      return { exito: false, error: "Este viaje ya fue despachado a ruta." };
    }

    const dDet = sheetDet.getDataRange().getValues();
    let totalItemsCargados = 0;

    for (let d = 1; d < dDet.length; d++) {
      const rd = dDet[d];
      if (String(rd[1] || "").trim() !== target || rd[16] === 0 || String(rd[16]) === "0") continue;
      totalItemsCargados++;

      const codArt = String(rd[2] || "").trim();
      const nomArt = String(rd[3] || "").trim();
      const tipoArt = String(rd[4] || "").trim();
      const zonaOrigen = String(rd[5] || "").trim();
      const zonaDestino = String(rd[6] || "").trim();
      const idOds = String(rd[7] || "").trim();
      const nomEv = String(rd[8] || "").trim();
      const locNom = String(rd[9] || "").trim();
      const fMont = _formatearFechaHora(rd[10]);
      const cantEnv = Number(rd[11]) || 0;

      const esEnvioEvento = infoViaje.tipo_viaje === "ENVIO_EVENTO";
      sheetDet.getRange(d + 1, 16).setValue(esEnvioEvento ? "ENTREGADO_LOCACION" : "EN_RUTA");

      // Si es ENVIO_EVENTO, al salir se descarga automáticamente en LOCACIONES (en el hotel con su ODS)
      if (esEnvioEvento) {
        _modificarStockZona(sheetZonas, {
          zona: "LOCACIONES",
          codigo_art: codArt,
          nombre_art: nomArt,
          tipo_art: tipoArt,
          id_ods: idOds,
          fecha_montaje: fMont,
          locacion_nombre: infoViaje.destino,
          nombre_evento: nomEv,
          delta: cantEnv
        });
      }

      const textoHist = esEnvioEvento
        ? ("🏨 ODS " + idOds + " x " + cantEnv + " (" + nomArt + ") salió de " + zonaOrigen + " y quedó entregado en locación " + infoViaje.destino)
        : ("🚚 ODS " + idOds + " x " + cantEnv + " (" + nomArt + ") salió en vehículo (bodega móvil) de " + zonaOrigen + " rumbo a " + zonaDestino);

      _registrarHistorial(sheetHist, {
        fecha_hora: ahora,
        id_viaje: target,
        tipo_movimiento: infoViaje.tipo_viaje === "ENVIO_EVENTO" ? "ENVIO_A_EVENTO" : "SALIDA_TRASLADO",
        bodega_origen: infoViaje.bodega_origen,
        zona_origen: zonaOrigen,
        bodega_destino: infoViaje.tipo_viaje === "ENVIO_EVENTO" ? "LOCACIONES" : infoViaje.destino,
        zona_destino: zonaDestino,
        id_ods: idOds,
        nombre_evento: nomEv,
        locacion_nombre: locNom || infoViaje.destino,
        vendedor_nombre: "",
        fecha_ods: fMont,
        codigo_art: codArt,
        nombre_art: nomArt,
        cantidad: cantEnv,
        vehiculo: infoViaje.vehiculo,
        chofer: infoViaje.chofer,
        novedad: "",
        resumen_texto: textoHist
      });
    }

    if (totalItemsCargados === 0) {
      return { exito: false, error: "El camión está vacío. Súbele al menos 1 artículo desde tu Zona de Alistamiento antes de darle salida." };
    }

    // Si el viaje es A EVENTO / HOTEL, se descarga automáticamente en la locación al darle salida y el camión queda libre (COMPLETADO)
    if (infoViaje.tipo_viaje === "ENVIO_EVENTO") {
      sheetViajes.getRange(filaViaje, 10).setValue("COMPLETADO");
      sheetViajes.getRange(filaViaje, 11).setValue(ahora);
      sheetViajes.getRange(filaViaje, 12).setValue(ahora);
      sheetViajes.getRange(filaViaje, 13).setValue("Entregado automáticamente en locación: " + infoViaje.destino);
    } else {
      sheetViajes.getRange(filaViaje, 10).setValue("EN_RUTA");
      sheetViajes.getRange(filaViaje, 11).setValue(ahora);
    }

    const estadoOp = _leerEstadoOperativoCompleto(ssOp);
    return {
      exito: true,
      entregado_en_locacion: infoViaje.tipo_viaje === "ENVIO_EVENTO",
      zonas: estadoOp.zonas,
      viajes: estadoOp.viajes,
      historial: estadoOp.historial
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Compatibilidad directa (si se envía viaje con items de un solo golpe)
 */
function apiCrearViaje(payload) {
  const resCab = apiCrearCabeceraViaje(payload);
  if (!resCab || !resCab.exito) return resCab;
  if (Array.isArray(payload.items) && payload.items.length > 0) {
    apiSubirItemsAViaje({ id_viaje: resCab.id_viaje, items: payload.items });
  }
  const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
  const estadoOp = _leerEstadoOperativoCompleto(ssOp);
  return {
    exito: true,
    id_viaje: resCab.id_viaje,
    zonas: estadoOp.zonas,
    viajes: estadoOp.viajes,
    historial: estadoOp.historial
  };
}

/**
 * 7. TOMAR ARTÍCULOS DE UN VEHÍCULO EN RUTA (BODEGA MÓVIL — DESCARGA PARCIAL O TOTAL HACIA ZONA DE ALISTAMIENTO PROPIA)
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
    const zonaDestinoSeleccionada = String(payload.zona_destino || "").trim().toUpperCase();
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
      return { exito: false, error: "Este vehículo ya fue descargado por completo." };
    }

    const dDet = sheetDet.getDataRange().getValues();
    let piezasRestantesEnVehiculo = 0;
    let piezasDescargadasEnEstaAccion = 0;

    for (let d = 1; d < dDet.length; d++) {
      const rd = dDet[d];
      if (String(rd[1] || "").trim() !== idViaje || rd[16] === 0 || String(rd[16]) === "0") continue;

      const idDet = String(rd[0] || "").trim();
      const codArt = String(rd[2] || "").trim();
      const nomArt = String(rd[3] || "").trim();
      const tipoArt = String(rd[4] || "MOBILIARIO").trim();
      const zonaOrigen = String(rd[5] || "").trim();
      const zonaDestinoDefault = String(rd[6] || "").trim();
      const idOds = String(rd[7] || "").trim();
      const nomEv = String(rd[8] || "").trim();
      const locNom = String(rd[9] || "").trim();
      const fMont = _formatearFechaHora(rd[10]);
      const cantEnCamion = Number(rd[11]) || 0;

      if (cantEnCamion <= 0) continue;

      const ajuste = itemsRecibidos[idDet];
      if (!ajuste) {
        piezasRestantesEnVehiculo += cantEnCamion;
        continue;
      }

      const cantTomada = ajuste.cant_recibida !== undefined
        ? Math.min(cantEnCamion, Math.max(0, Number(ajuste.cant_recibida) || 0))
        : 0;
      const novedadItem = String(ajuste.novedad || "").trim();
      const zonaFinalItem = String(ajuste.zona_destino || zonaDestinoSeleccionada || zonaDestinoDefault || infoViaje.destino).trim().toUpperCase();
      const bodegaFinalItem = _obtenerBodegaBaseDeZona(zonaFinalItem);

      if (cantTomada <= 0 && !novedadItem) {
        piezasRestantesEnVehiculo += cantEnCamion;
        continue;
      }

      piezasDescargadasEnEstaAccion += cantTomada;
      const cantRestanteItem = Math.max(0, cantEnCamion - cantTomada);

      if (cantTomada > 0) {
        _modificarStockZona(sheetZonas, {
          zona: zonaFinalItem,
          codigo_art: codArt,
          nombre_art: nomArt,
          tipo_art: tipoArt,
          id_ods: idOds,
          fecha_montaje: fMont,
          locacion_nombre: locNom,
          nombre_evento: nomEv,
          delta: cantTomada
        });

        const txtRec = "📥 ODS " + idOds + " x " + cantTomada + " (" + nomArt + ") se bajó del vehículo " + infoViaje.vehiculo + " (" + idViaje + ") hacia zona " + zonaFinalItem + (novedadItem ? (" [Novedad: " + novedadItem + "]") : "");
        _registrarHistorial(sheetHist, {
          fecha_hora: ahora,
          id_viaje: idViaje,
          tipo_movimiento: "DESCARGA_VEHICULO_A_ZONA",
          bodega_origen: infoViaje.bodega_origen,
          zona_origen: zonaOrigen,
          bodega_destino: bodegaFinalItem,
          zona_destino: zonaFinalItem,
          id_ods: idOds,
          nombre_evento: nomEv,
          locacion_nombre: locNom,
          vendedor_nombre: "",
          fecha_ods: fMont,
          codigo_art: codArt,
          nombre_art: nomArt,
          cantidad: cantTomada,
          vehiculo: infoViaje.vehiculo,
          chofer: infoViaje.chofer,
          novedad: novedadItem,
          resumen_texto: txtRec
        });
      }

      if (novedadItem) {
        _registrarHistorial(sheetHist, {
          fecha_hora: ahora,
          id_viaje: idViaje,
          tipo_movimiento: "NOVEDAD_FALTANTE",
          bodega_origen: infoViaje.bodega_origen,
          zona_origen: zonaOrigen,
          bodega_destino: bodegaFinalItem,
          zona_destino: zonaFinalItem,
          id_ods: idOds,
          nombre_evento: nomEv,
          locacion_nombre: locNom,
          vendedor_nombre: "",
          fecha_ods: fMont,
          codigo_art: codArt,
          nombre_art: nomArt,
          cantidad: cantTomada,
          vehiculo: infoViaje.vehiculo,
          chofer: infoViaje.chofer,
          novedad: novedadItem,
          resumen_texto: "⚠️ ODS " + idOds + " (" + nomArt + ") en " + idViaje + ": " + novedadItem
        });
      }

      if (cantRestanteItem > 0) {
        // Descarga parcial: el resto del artículo sigue arriba del vehículo (bodega móvil)
        sheetDet.getRange(d + 1, 12).setValue(cantRestanteItem);
        sheetDet.getRange(d + 1, 13).setValue(cantRestanteItem);
        if (novedadItem) sheetDet.getRange(d + 1, 15).setValue(novedadItem);
        piezasRestantesEnVehiculo += cantRestanteItem;
      } else {
        // Se bajó el total de este renglón
        sheetDet.getRange(d + 1, 7).setValue(zonaFinalItem);
        sheetDet.getRange(d + 1, 13).setValue(cantTomada);
        sheetDet.getRange(d + 1, 15).setValue(novedadItem);
        sheetDet.getRange(d + 1, 16).setValue("DESCARGADO");
        sheetDet.getRange(d + 1, 17).setValue(0);
      }
    }

    if (piezasDescargadasEnEstaAccion === 0) {
      return { exito: false, error: "Indica al menos 1 pieza a bajar del vehículo hacia tu Zona de Alistamiento." };
    }

    if (piezasRestantesEnVehiculo === 0) {
      sheetViajes.getRange(filaViaje, 10).setValue("COMPLETADO");
      sheetViajes.getRange(filaViaje, 12).setValue(ahora);
      if (obsGeneral) sheetViajes.getRange(filaViaje, 13).setValue(obsGeneral);
    } else if (obsGeneral) {
      sheetViajes.getRange(filaViaje, 13).setValue(obsGeneral);
    }

    const estadoOp = _leerEstadoOperativoCompleto(ssOp);
    return {
      exito: true,
      viaje_completado: piezasRestantesEnVehiculo === 0,
      piezas_restantes: piezasRestantesEnVehiculo,
      zonas: estadoOp.zonas,
      viajes: estadoOp.viajes,
      historial: estadoOp.historial
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 7B. CREAR REPORTE DE VIAJE (RECIBIR DE LOCACIÓN / HOTELES HACIA MI ZONA DE ALISTAMIENTO)
 *     Solo permite recibir artículos que estén actualmente en la zona LOCACIONES,
 *     sin importar si el chofer los trajo de 1 hotel o de varios hoteles.
 */
function apiCrearReporteViajeLocacion(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ssOp = SpreadsheetApp.openById(EYS_CONFIG.IDS.SS_OPERACION);
    _garantizarHojasOperacion(ssOp);
    const sheetViajes = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES.NOMBRE);
    const sheetDet = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.VIAJES_DETALLE.NOMBRE);
    const sheetZonas = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.ZONAS.NOMBRE);
    const sheetHist = ssOp.getSheetByName(EYS_CONFIG.HOJAS_OP.HISTORIAL.NOMBRE);

    const vehiculo = String(payload.vehiculo || "").trim();
    const chofer = String(payload.chofer || "").trim();
    const zonaDestino = String(payload.zona_destino || "").trim().toUpperCase();
    const bodegaDestino = _obtenerBodegaBaseDeZona(zonaDestino);
    const observaciones = String(payload.observaciones || "").trim();
    const items = Array.isArray(payload.items) ? payload.items : [];
    const ahora = _ahoraStr();

    if (!vehiculo || !chofer || !zonaDestino) {
      return { exito: false, error: "Selecciona Unidad/Vehículo, Chofer y tu Zona de Alistamiento destino." };
    }
    if (items.length === 0) {
      return { exito: false, error: "Selecciona al menos un artículo de LOCACIONES que haya traído el chofer." };
    }

    const numViaje = sheetViajes.getLastRow();
    const idViaje = "RV-" + String(numViaje).padStart(4, "0");

    // Recolectar hoteles de los que trajo carga en este reporte
    const hotelesSet = {};
    items.forEach(function(it) {
      if (it.locacion_nombre) hotelesSet[String(it.locacion_nombre).trim()] = true;
    });
    const origenHotelesTxt = Object.keys(hotelesSet).length > 0
      ? ("LOCACIONES (" + Object.keys(hotelesSet).join(", ") + ")")
      : "LOCACIONES";

    sheetViajes.appendRow([
      idViaje,
      "REPORTE_LOCACION",
      origenHotelesTxt,
      zonaDestino,
      "",
      "Reporte de Recepción desde Locación",
      ahora,
      vehiculo,
      chofer,
      "COMPLETADO",
      ahora,
      ahora,
      observaciones || ("Recibido en " + zonaDestino + " desde " + origenHotelesTxt),
      1
    ]);

    items.forEach(function(it) {
      const cantRec = Math.max(0, Number(it.cantidad) || 0);
      if (!it.codigo_art || cantRec <= 0) return;

      const codArt = String(it.codigo_art).trim();
      const nomArt = String(it.nombre_art || "").trim();
      const tipoArt = String(it.tipo_art || "MOBILIARIO").trim();
      const idOds = String(it.id_ods || "").trim();
      const nomEv = String(it.nombre_evento || "").trim();
      const locNom = String(it.locacion_nombre || "").trim();
      const fMont = _formatearFechaHora(it.fecha_montaje);
      const novedadItem = String(it.novedad || "").trim();

      // 1. Descontar de la zona LOCACIONES
      _modificarStockZona(sheetZonas, {
        zona: "LOCACIONES",
        codigo_art: codArt,
        nombre_art: nomArt,
        tipo_art: tipoArt,
        id_ods: idOds,
        fecha_montaje: fMont,
        locacion_nombre: locNom,
        nombre_evento: nomEv,
        delta: -cantRec
      });

      // 2. Ingresar a la Zona de Alistamiento propia del encargado que recibe
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

      const idDet = idViaje + "-D" + new Date().getTime() + "-" + Math.floor(Math.random() * 900 + 100);
      sheetDet.appendRow([
        idDet,
        idViaje,
        codArt,
        nomArt,
        tipoArt,
        "LOCACIONES (" + (locNom || "Hotel") + ")",
        zonaDestino,
        idOds,
        nomEv,
        locNom,
        fMont,
        cantRec,
        cantRec,
        0,
        novedadItem,
        novedadItem ? "CON_NOVEDAD" : "COMPLETO",
        1
      ]);

      const txtHist = "📋 Reporte " + idViaje + ": ODS " + idOds + " x " + cantRec + " (" + nomArt + ") recibido de locación " + (locNom || "Hotel") + " hacia zona " + zonaDestino + " (Chofer: " + chofer + ")" + (novedadItem ? (" [Novedad: " + novedadItem + "]") : "");
      _registrarHistorial(sheetHist, {
        fecha_hora: ahora,
        id_viaje: idViaje,
        tipo_movimiento: "REPORTE_RECEPCION_LOCACION",
        bodega_origen: "LOCACIONES",
        zona_origen: locNom || "LOCACIONES",
        bodega_destino: bodegaDestino,
        zona_destino: zonaDestino,
        id_ods: idOds,
        nombre_evento: nomEv,
        locacion_nombre: locNom,
        vendedor_nombre: "",
        fecha_ods: fMont,
        codigo_art: codArt,
        nombre_art: nomArt,
        cantidad: cantRec,
        vehiculo: vehiculo,
        chofer: chofer,
        novedad: novedadItem,
        resumen_texto: txtHist
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
 * 8. CANCELAR UN VIAJE (EN ANDÉN O EN RUTA) -> REVIERTE STOCK A ORIGEN
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
          destino: String(dViajes[v][3] || "").trim(),
          estado: String(dViajes[v][9] || "EN_CARGA").trim().toUpperCase()
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

      // Si ya estaba EN_RUTA y era ENVIO_EVENTO, quitar de LOCACIONES
      if ((infoViaje.estado === "EN_RUTA" || infoViaje.estado === "EN_CURSO") && infoViaje.tipo_viaje === "ENVIO_EVENTO") {
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
