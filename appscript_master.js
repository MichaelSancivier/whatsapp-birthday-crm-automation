/**
 * ==============================================================================
 * ATOM & REVOPS AUTOMATION: SYSTEM TEMPLATE FOR BIRTHDAYS & CRM AUDIT
 * ==============================================================================
 * Integración: Google Drive -> Google Sheets CRM -> Atom Webhook API (WhatsApp)
 * Descripción: Ingesta de bases, sanitización E.164, motor de cumpleaños,
 *              candado antiduplicidad, log de auditoría y reporte diario HTML.
 * Autor: Michael Sancivier (Senior Onboarding Specialist / RevOps)
 * ==============================================================================
 */

// ⚙️ 1. CONFIGURACIÓN GLOBAL DEL PROYECTO
const CONFIG = {
  CLIENTE_NOMBRE: "NOMBRE_CLIENTE",
  EMAIL_CONTACTO_CLIENTE: "Contacto",
  EMAILS_NOTIFICACION: "manager@empresa.com, ops@empresa.com",
  
  // Integración Webhook Atom
  ATOM_WEBHOOK_URL: "https://api.empresa.com/v1/webhooks/hsm/trigger",
  ATOM_TOKEN: "YOUR_BEARER_TOKEN_HERE",
  
  // Repositorios en Google Drive
  FOLDER_ENTRADA_ID: "YOUR_INPUT_FOLDER_ID",
  FOLDER_PROCESADOS_ID: "YOUR_PROCESSED_FOLDER_ID",
  
  // Reglas Telefonía Internacional (E.164)
  CODIGO_PAIS_DEFAULT: "591",  // Ej: 591 para Bolivia, 55 para Brasil, 52 para México
  LARGO_NUMERO_LOCAL: 8,       // Cantidad de dígitos locales
  
  // Zona Horaria
  TIMEZONE: "GMT-4"
};

// 🚀 2. MENÚ DE INTERFAZ Y AUTOMATIZACIÓN (UI)
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Control Atom')
    .addItem('🚀 Disparo de Emergencia (Procesar HOY)', 'ejecutarDisparoManual')
    .addItem('📥 Importar Archivos de Drive Ahora', 'importarArchivosManual')
    .addSeparator()
    .addItem('🔧 Inicializar Estructura de Hojas', 'inicializarEstructura')
    .addToUi();
}

function ejecutarDisparoManual() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('⚠️ Ejecución Manual', '¿Deseas procesar la base de Drive y disparar los cumpleaños de HOY?', ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    ejecucionDiariaCumpleanos();
    ui.alert('✅ Proceso completado. Revisa la pestaña "Historial_Enviados" y el correo.');
  }
}

function importarArchivosManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMaestra = ss.getSheetByName("Base_Maestra_Contactos") || crearHojaMaestra(ss);
  importarNuevosArchivosDrive(hojaMaestra);
  SpreadsheetApp.getUi().alert('✅ Archivos importados a la "Base_Maestra_Contactos".');
}

function inicializarEstructura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let hModelo = ss.getSheetByName("Plantilla_Modelo") || ss.insertSheet("Plantilla_Modelo");
  hModelo.clear();
  hModelo.appendRow(["nombre", "fechanacimiento", "cumpleaños", "nrodetelefono4"]);
  hModelo.getRange("A1:D1").setFontWeight("bold").setBackground("#FF5B00").setFontColor("#FFFFFF");
  hModelo.appendRow(["EJEMPLO CLIENTE PROSPECTO", "20/11/1985", "20/11", "70000000"]);

  crearHojaMaestra(ss);
  crearHojaLog(ss);
  
  let hDefault = ss.getSheetByName("Hoja 1") || ss.getSheetByName("Sheet1");
  if (hDefault && ss.getSheets().length > 1) ss.deleteSheet(hDefault);
  
  SpreadsheetApp.getUi().alert("✅ Estructura inicializada correctamente.");
}

// 🔄 3. MOTOR PRINCIPAL DE PROCESAMIENTO
function ejecucionDiariaCumpleanos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaMaestra = ss.getSheetByName("Base_Maestra_Contactos") || crearHojaMaestra(ss);
  const hojaLog = ss.getSheetByName("Historial_Enviados") || crearHojaLog(ss);

  importarNuevosArchivosDrive(hojaMaestra);
  procesarEnvioCumpleanos(ss, hojaMaestra, hojaLog);
}

function importarNuevosArchivosDrive(hojaMaestra) {
  if (CONFIG.FOLDER_ENTRADA_ID.includes("YOUR_")) return;
  
  const folderEntrada = DriveApp.getFolderById(CONFIG.FOLDER_ENTRADA_ID);
  const folderProcesados = DriveApp.getFolderById(CONFIG.FOLDER_PROCESADOS_ID);
  const archivos = folderEntrada.getFiles();

  while (archivos.hasNext()) {
    let archivo = archivos.next();
    if (archivo.getMimeType() === MimeType.FOLDER) continue;

    let nombreArchivo = archivo.getName();
    let fechaCargaArchivo = Utilities.formatDate(archivo.getLastUpdated(), CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm");

    try {
      let filas = [];
      if (nombreArchivo.toLowerCase().endsWith(".xlsx")) {
        filas = leerFilasDesdeXlsx(archivo);
      } else {
        let contenidoCsv = archivo.getBlob().getDataAsString();
        filas = Utilities.parseCsv(contenidoCsv);
      }

      if (filas.length > 1) {
        let encabezados = filas[0].map(e => e.toString().toLowerCase().trim());
        
        let colNombre = encabezados.findIndex(e => e.includes("nombre") || e.includes("client"));
        let colTel = encabezados.findIndex(e => e.includes("telefono") || e.includes("nrodetelefono") || e.includes("tel") || e.includes("cel") || e.includes("phone"));
        let colCumple = encabezados.findIndex(e => e.includes("cumpleaños") || e.includes("cumpleanos"));
        let colFechaNac = encabezados.findIndex(e => e.includes("fechanacimiento") || e.includes("nacimiento"));
        
        let colFechaFinal = colCumple !== -1 ? colCumple : colFechaNac;

        if (colNombre !== -1 && colTel !== -1 && colFechaFinal !== -1) {
          upsertContactosMaestra(hojaMaestra, filas, colNombre, colTel, colFechaFinal, nombreArchivo, fechaCargaArchivo);
        }
      }

      archivo.moveTo(folderProcesados);
    } catch (error) {
      Logger.log("Error procesando " + nombreArchivo + ": " + error.toString());
    }
  }
}

function upsertContactosMaestra(hoja, filas, colNombre, colTel, colFecha, nombreArchivo, fechaCarga) {
  hoja.getRange("C:E").setNumberFormat("@");
  let datosExistentes = hoja.getDataRange().getValues();
  let mapaMaster = new Map();

  for (let i = 1; i < datosExistentes.length; i++) {
    let tel = String(datosExistentes[i][1]).replace(/\D/g, '');
    if (tel) {
      mapaMaster.set(tel, [
        datosExistentes[i][0],
        datosExistentes[i][1],
        datosExistentes[i][2],
        datosExistentes[i][3],
        "'" + String(datosExistentes[i][4]).replace(/^'/, '')
      ]);
    }
  }

  let fechaCargaTexto = "'" + fechaCarga;

  for (let i = 1; i < filas.length; i++) {
    let nombre = String(filas[i][colNombre]).trim();
    let telRaw = String(filas[i][colTel]);
    let valorCumple = filas[i][colFecha];

    if (!nombre || !telRaw || !valorCumple) continue;

    let telLimpio = sanitizarTelefonoUniversal(telRaw);
    let fechaNormalizada = formatearCumpleanosDDMM(valorCumple);

    if (!telLimpio || !fechaNormalizada) continue;

    mapaMaster.set(telLimpio, [nombre, telLimpio, "'" + fechaNormalizada, nombreArchivo, fechaCargaTexto]);
  }

  let matrizFinal = Array.from(mapaMaster.values());
  if (matrizFinal.length > 0) {
    let filasAfectadas = hoja.getLastRow() > 1 ? hoja.getLastRow() - 1 : 1;
    hoja.getRange(2, 1, filasAfectadas, 5).clearContent();
    hoja.getRange(2, 1, matrizFinal.length, 5).setValues(matrizFinal);
  }
}

function procesarEnvioCumpleanos(ss, hojaMaestra, hojaLog) {
  let contactos = hojaMaestra.getDataRange().getValues();
  let totalGeneral = Math.max(0, contactos.length - 1);
  let mapaArchivos = {};

  for (let i = 1; i < contactos.length; i++) {
    let nombreArch = String(contactos[i][3] || "Carga Manual").trim();
    let fechaCargaArch = String(contactos[i][4] || "N/A").replace(/^'/, '');

    if (!mapaArchivos[nombreArch]) {
      mapaArchivos[nombreArch] = { count: 0, fechaCarga: fechaCargaArch };
    }
    mapaArchivos[nombreArch].count++;
  }

  let totalArchivos = Object.keys(mapaArchivos).length;
  let hoy = new Date();
  let diaHoyStr = Utilities.formatDate(hoy, CONFIG.TIMEZONE, "dd/MM");
  let fechaHoyISO = Utilities.formatDate(hoy, CONFIG.TIMEZONE, "yyyy-MM-dd");

  let enviadosHoy = obtenerTelefonosEnviadosHoy(hojaLog, fechaHoyISO);
  let procesadosEnEstaCorrida = {};
  let resumenEnvios = [];
  let totalCumpleHoy = 0;

  for (let i = 1; i < contactos.length; i++) {
    let nombre = String(contactos[i][0]).trim();
    let tel = String(contactos[i][1]).trim();
    let valorCeldaCumple = contactos[i][2];
    let archivoOrigen = contactos[i][3];
    let fechaCargaOrigen = String(contactos[i][4]).replace(/^'/, '');

    let cumpleDDMM = formatearCumpleanosDDMM(valorCeldaCumple);

    if (cumpleDDMM === diaHoyStr) {
      totalCumpleHoy++;

      if (enviadosHoy.includes(tel) || procesadosEnEstaCorrida[tel]) {
        resumenEnvios.push({
          nombre: nombre,
          telefono: tel,
          archivo: archivoOrigen,
          estado: "🟡 Omitido: Ya se procesó una orden de disparo hoy"
        });
        continue;
      }

      let exitoAtom = dispararWebhookAtom(tel, nombre);
      procesadosEnEstaCorrida[tel] = true;

      hojaLog.appendRow([
        new Date(), nombre, tel, cumpleDDMM, archivoOrigen, fechaCargaOrigen, exitoAtom ? "DISPARO_EXITOSO" : "ERROR_CONEXION"
      ]);

      resumenEnvios.push({
        nombre: nombre,
        telefono: tel,
        archivo: archivoOrigen,
        estado: exitoAtom ? "🟢 Orden enviada a Atom con éxito" : "🔴 Error de conexión con Webhook de Atom"
      });

      Utilities.sleep(300);
    }
  }

  let stats = {
    totalGeneral: totalGeneral,
    totalArchivos: totalArchivos,
    archivos: mapaArchivos,
    totalCumpleHoy: totalCumpleHoy
  };

  enviarReporteCorreo(CONFIG.EMAILS_NOTIFICACION, resumenEnvios, stats);
}

// 🛠️ 4. FUNCIONES AUXILIARES
function sanitizarTelefonoUniversal(telefonoRaw) {
  let num = String(telefonoRaw).replace(/\D/g, '');
  if (!num) return "";
  if (num.length === CONFIG.LARGO_NUMERO_LOCAL) {
    return CONFIG.CODIGO_PAIS_DEFAULT + num;
  }
  return num;
}

function formatearCumpleanosDDMM(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, CONFIG.TIMEZONE, "dd/MM");
  }
  let valStr = String(valor).trim().replace(/^'/, '');
  if (valStr.includes("/")) {
    let partes = valStr.split("/");
    if (partes.length >= 2) {
      return partes[0].padStart(2, '0') + "/" + partes[1].padStart(2, '0');
    }
  }
  return valStr;
}

function leerFilasDesdeXlsx(archivo) {
  let blob = archivo.getBlob();
  let resource = { name: "Temp_" + archivo.getName(), mimeType: MimeType.GOOGLE_SHEETS };
  let tempFile = Drive.Files.create(resource, blob);
  let tempSheet = SpreadsheetApp.openById(tempFile.id).getSheets()[0];
  let datos = tempSheet.getDataRange().getValues();
  DriveApp.getFileById(tempFile.id).setTrashed(true);
  return datos;
}

function dispararWebhookAtom(telefono, nombre) {
  try {
    let payload = { "first_name": String(nombre), "phone": String(telefono) };
    let options = {
      "method": "post",
      "contentType": "application/json",
      "headers": {
        "Authorization": "Bearer " + CONFIG.ATOM_TOKEN,
        "x-api-key": CONFIG.ATOM_TOKEN
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    let response = UrlFetchApp.fetch(CONFIG.ATOM_WEBHOOK_URL, options);
    let codigo = response.getResponseCode();
    return codigo === 200 || codigo === 201 || codigo === 202;
  } catch (e) {
    Logger.log("🔴 Excepción de red: " + e.toString());
    return false;
  }
}

function obtenerTelefonosEnviadosHoy(hojaLog, fechaHoyISO) {
  let datos = hojaLog.getDataRange().getValues();
  let enviados = [];
  for (let i = 1; i < datos.length; i++) {
    let fecha = datos[i][0];
    let tel = String(datos[i][2]);
    let st = datos[i][6];
    if (fecha instanceof Date) {
      let fStr = Utilities.formatDate(fecha, CONFIG.TIMEZONE, "yyyy-MM-dd");
      if (fStr === fechaHoyISO && st === "DISPARO_EXITOSO") enviados.push(tel);
    }
  }
  return enviados;
}

function crearHojaMaestra(ss) {
  let h = ss.insertSheet("Base_Maestra_Contactos");
  h.getRange("C:E").setNumberFormat("@");
  h.appendRow(["Nombre", "Teléfono", "Fecha Cumpleaños (DD/MM)", "Archivo Origen", "Fecha Carga Archivo"]);
  h.getRange("A1:E1").setFontWeight("bold").setBackground("#2C3E50").setFontColor("#FFFFFF");
  return h;
}

function crearHojaLog(ss) {
  let h = ss.insertSheet("Historial_Enviados");
  h.appendRow(["Fecha/Hora Envío", "Nombre", "Teléfono", "Fecha Cumpleaños", "Archivo Origen", "Fecha Carga Archivo", "Estado Disparo Atom"]);
  h.getRange("A1:G1").setFontWeight("bold").setBackground("#34495E").setFontColor("#FFFFFF");
  return h;
}

// 📧 5. GENERADOR DE CORREO DE AUDITORÍA HTML
function enviarReporteCorreo(destinatario, lista, stats) {
  let fechaHoyStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "dd/MM/yyyy");
  let asunto = `[Atom Auto] Resumen Diario de Cumpleaños & Auditoría CRM — ${fechaHoyStr}`;

  let html = `
  <div style="font-family: Arial, Helvetica, sans-serif; color: #333333; max-width: 650px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff;">
    <p style="font-size: 15px; margin-bottom: 20px;">Hola ${CONFIG.EMAIL_CONTACTO_CLIENTE},</p>
    <p style="font-size: 14px; color: #555555; line-height: 1.5;">A continuación se presenta el reporte diario de salud de la base CRM y el procesamiento de cumpleaños hacia la plataforma <strong>Atom</strong> para <strong>${CONFIG.CLIENTE_NOMBRE}</strong>:</p>
    
    <div style="background-color: #f8f9fa; border-left: 4px solid #FF5B00; padding: 15px; margin: 20px 0; border-radius: 4px;">
      <h3 style="margin: 0 0 10px 0; font-size: 15px; color: #111111;">📊 1. RESUMEN GENERAL DE LA BASE CRM</h3>
      <ul style="margin: 0; padding-left: 20px; font-size: 13px; color: #444444; line-height: 1.8;">
        <li><strong>Total de clientes en base general:</strong> ${stats.totalGeneral.toLocaleString()}</li>
        <li><strong>Total de cumpleañeros identificados para HOY (${fechaHoyStr}):</strong> <span style="color: #FF5B00; font-weight: bold;">${stats.totalCumpleHoy}</span></li>
        <li><strong>Total de archivos procesados e integrados:</strong> ${stats.totalArchivos}</li>
      </ul>
    </div>
    
    <div style="background-color: #f8f9fa; border-left: 4px solid #2C3E50; padding: 15px; margin: 20px 0; border-radius: 4px;">
      <h3 style="margin: 0 0 10px 0; font-size: 15px; color: #111111;">📁 2. HISTORIAL Y DESGLOSE POR ARCHIVO ORIGEN</h3>
      <div style="font-size: 13px; color: #444444;">`;
      
      if (stats.totalArchivos === 0) {
        html += `<p style="margin: 0; color: #777;">No hay archivos cargados en la base maestra.</p>`;
      } else {
        for (let nombreArchivo in stats.archivos) {
          let info = stats.archivos[nombreArchivo];
          html += `
          <div style="margin-bottom: 10px; padding-bottom: 8px; border-bottom: 1px dashed #e0e0e0;">
            <strong>• Archivo:</strong> '${nombreArchivo}'<br>
            <span style="color: #666; font-size: 12px;">Fecha de Carga: ${info.fechaCarga} | Clientes Registrados: ${info.count.toLocaleString()}</span>
          </div>`;
        }
      }
      
  html += `
      </div>
    </div>
    
    <div style="background-color: #f8f9fa; border-left: 4px solid #27ae60; padding: 15px; margin: 20px 0; border-radius: 4px;">
      <h3 style="margin: 0 0 10px 0; font-size: 15px; color: #111111;">🎉 3. DETALLE DE DISPAROS REALIZADOS HOY (${lista.length})</h3>
      <div style="font-size: 13px; color: #444444;">`;
      
      if (lista.length === 0) {
        html += `<p style="margin: 0; color: #777;">📅 Estado: Para hoy (${fechaHoyStr}) no se registran cumpleañeros en la base de datos.</p>`;
      } else {
        lista.forEach(item => {
          html += `
          <div style="margin-bottom: 8px; font-size: 12px;">
            <strong>• ${item.nombre}</strong> (${item.telefono})<br>
            <span style="color: #666;">Archivo Origen: '${item.archivo}' | Estatus: ${item.estado}</span>
          </div>`;
        });
      }
      
  html += `
      </div>
    </div>
    
    <p style="font-size: 11px; color: #888888; font-style: italic; margin-top: 25px; border-top: 1px solid #eeeeee; padding-top: 10px;">
      Nota: Este reporte confirma el procesamiento interno y el disparo de las órdenes a Atom. La entrega final del mensaje de WhatsApp se monitorea directamente en la consola de Atom.
    </p>

    <div style="margin-top: 35px; border-top: 2px solid #FF5B00; padding-top: 20px; font-family: Arial, sans-serif;">
      <table border="0" cellpadding="0" cellspacing="0" style="width: 100%;">
        <tr>
          <td style="vertical-align: top; padding-bottom: 10px;">
            <table border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td style="vertical-align: middle; padding-right: 8px;">
                  <div style="width: 28px; height: 28px; background-color: #FF5B00; border-radius: 50%; text-align: center; line-height: 28px; color: white; font-weight: bold; font-size: 16px;">⚛</div>
                </td>
                <td style="vertical-align: middle;">
                  <span style="font-size: 24px; font-weight: 900; color: #000000; letter-spacing: 0.5px;">ATOM</span>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="padding-bottom: 4px;">
            <span style="font-size: 16px; font-weight: bold; color: #666666;">Michael Sancivier</span>
          </td>
        </tr>
        <tr>
          <td style="padding-bottom: 12px;">
            <span style="font-size: 14px; font-weight: bold; color: #222222;">Especialista Onboarding (CSM)</span>
          </td>
        </tr>
        <tr>
          <td style="font-size: 13px; color: #444444; line-height: 1.8;">
            <span style="color: #FF5B00; font-weight: bold;">📞</span> +55 4198514.4526<br>
            <span style="color: #FF5B00; font-weight: bold;">📍</span> Brasil<br>
            <span style="color: #FF5B00; font-weight: bold;">🌐</span> <a href="https://atomchat.io" style="color: #FF5B00; text-decoration: underline; font-weight: bold;">atomchat.io</a>
          </td>
        </tr>
      </table>
    </div>

  </div>`;

  try {
    MailApp.sendEmail({
      to: destinatario,
      subject: asunto,
      htmlBody: html
    });
  } catch (e) {
    Logger.log("Error enviando email: " + e.toString());
  }
}
