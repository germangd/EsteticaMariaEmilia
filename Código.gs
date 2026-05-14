// ========== CONFIGURACIÓN ==========
/**
 * Propiedades del proyecto (Apps Script → Archivo → Propiedades del proyecto → Propiedades de secuencias de comandos):
 * - ADMIN_PASSWORD — Contraseña del panel de administración. Si no existe, se usa ADMIN_PASSWORD_DEFAULT en código (cambiala en producción).
 * - ADMIN_RECOVERY_EMAIL — Email (minúsculas) que recibe el enlace de “olvidé mi contraseña”. Si no existe, se usa EMAIL_DUENIO.
 * - WEB_APP_URL — URL completa del despliegue web (…/exec). Solo si falla ScriptApp.getService().getUrl() al generar el enlace del correo.
 */
const HOJA_TURNOS = "Turnos";
const HOJA_CONFIG = "Config";
const EMAIL_DUENIO = "germangd23@gmail.com";
/** Clave en Propiedades del proyecto: ADMIN_PASSWORD. Si no existe, se usa el valor por defecto (cambialo en producción). */
const ADMIN_PASSWORD_DEFAULT = "mestetica2026";
/** Prefijos CacheService (máx. ~6 h de sesión; reset 15 min). */
const CACHE_ADMIN_SESION = "adm_ses_";
const CACHE_ADMIN_RESET = "adm_rst_";
const ADMIN_SESION_SEGUNDOS = 21600; // 6 horas
const ADMIN_RESET_SEGUNDOS = 900; // 15 min
const ADMIN_PASSWORD_MIN_LEN = 6;

// ========== OBTENER ZONA HORARIA (cache por ejecución) ==========
var _zonaHorariaCache = null;
function getZonaHoraria() {
  if (_zonaHorariaCache === null) {
    _zonaHorariaCache = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  }
  return _zonaHorariaCache;
}

function obtenerPasswordAdminEsperada() {
  const configured = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");
  return configured && String(configured).length > 0 ? String(configured) : ADMIN_PASSWORD_DEFAULT;
}

/** Email donde se envía el enlace de recuperación: propiedad ADMIN_RECOVERY_EMAIL o EMAIL_DUENIO. */
function obtenerEmailRecuperacionAdmin() {
  const props = PropertiesService.getScriptProperties();
  const custom = props.getProperty("ADMIN_RECOVERY_EMAIL");
  if (custom && String(custom).trim().length > 0) return String(custom).trim().toLowerCase();
  return String(EMAIL_DUENIO || "").trim().toLowerCase();
}

function obtenerUrlWebApp() {
  let url = "";
  try {
    url = ScriptApp.getService().getUrl();
  } catch (e) {}
  if (!url) {
    url = PropertiesService.getScriptProperties().getProperty("WEB_APP_URL") || "";
  }
  return String(url).trim();
}

function crearSesionAdmin() {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(CACHE_ADMIN_SESION + token, "1", ADMIN_SESION_SEGUNDOS);
  return token;
}

function sesionAdminValida(token) {
  if (token == null || String(token).trim() === "") return false;
  return CacheService.getScriptCache().get(CACHE_ADMIN_SESION + String(token)) != null;
}

/** Invalida la sesión (cerrar sesión). */
function cerrarSesionAdmin(sessionToken) {
  if (sessionToken) CacheService.getScriptCache().remove(CACHE_ADMIN_SESION + String(sessionToken));
  return { exito: true };
}

/**
 * Login admin: devuelve token de sesión para usar en obtenerTurnosHoy / obtenerTodosTurnosFuturos.
 * Respuesta: { exito: true, token } o { exito: false }.
 */
function verificarPasswordAdmin(password) {
  if (password == null || String(password).trim() === "") return { exito: false };
  const expected = obtenerPasswordAdminEsperada();
  if (String(password) !== expected) return { exito: false };
  return { exito: true, token: crearSesionAdmin() };
}

/**
 * Solicita recuperación: solo envía email si el correo coincide con ADMIN_RECOVERY_EMAIL o EMAIL_DUENIO.
 * Respuesta siempre genérica por privacidad.
 */
function solicitarRecuperacionAdmin(email) {
  const mensajeGenerico = {
    exito: true,
    mensaje:
      "Si el correo está registrado para recuperación, recibirás un enlace en los próximos minutos. Revisá también spam."
  };
  const em = email == null ? "" : String(email).trim().toLowerCase();
  if (!em || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) return mensajeGenerico;

  const permitido = obtenerEmailRecuperacionAdmin();
  if (!permitido || em !== permitido) return mensajeGenerico;

  const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  CacheService.getScriptCache().put(CACHE_ADMIN_RESET + token, "1", ADMIN_RESET_SEGUNDOS);

  const base = obtenerUrlWebApp();
  if (!base) {
    return {
      exito: false,
      mensaje:
        "Falta la URL del despliegue. En Apps Script: Implementar → URL de la aplicación web, o configurá la propiedad WEB_APP_URL en Propiedades del proyecto."
    };
  }
  const sep = base.indexOf("?") >= 0 ? "&" : "?";
  const link = base + sep + "page=adminReset&token=" + encodeURIComponent(token);

  const html =
    "<div style=\"font-family:Arial,sans-serif;max-width:520px;margin:auto;padding:24px;\">" +
    "<p style=\"color:#2C2420;\">Recibimos una solicitud para <strong>cambiar la contraseña de administración</strong> de María Emilia Estética.</p>" +
    "<p style=\"color:#666;font-size:14px;\">Si no fuiste vos, ignorá este mensaje. El enlace vence en 15 minutos.</p>" +
    "<p style=\"margin:28px 0;\"><a href=\"" +
    link +
    "\" style=\"background:#d46b6b;color:#fff;padding:14px 24px;border-radius:10px;text-decoration:none;font-weight:bold;display:inline-block;\">Elegir nueva contraseña</a></p>" +
    "<p style=\"color:#aaa;font-size:12px;word-break:break-all;\">" +
    link +
    "</p></div>";

  try {
    MailApp.sendEmail({
      to: em,
      subject: "Recuperar contraseña de administración — María Emilia Estética",
      htmlBody: html
    });
  } catch (e) {
    return { exito: false, mensaje: "No se pudo enviar el email. Revisá cuotas y permisos de MailApp." };
  }
  return mensajeGenerico;
}

/**
 * Define nueva contraseña de admin (propiedad ADMIN_PASSWORD). Invalida el token de un solo uso.
 */
function restablecerPasswordAdmin(token, nuevaPassword) {
  const t = token == null ? "" : String(token).trim();
  const p = nuevaPassword == null ? "" : String(nuevaPassword);
  if (!t) return { exito: false, mensaje: "Enlace inválido o incompleto." };
  if (p.length < ADMIN_PASSWORD_MIN_LEN) {
    return { exito: false, mensaje: "La contraseña debe tener al menos " + ADMIN_PASSWORD_MIN_LEN + " caracteres." };
  }
  const key = CACHE_ADMIN_RESET + t;
  if (CacheService.getScriptCache().get(key) == null) {
    return { exito: false, mensaje: "El enlace expiró o ya fue usado. Solicitá uno nuevo." };
  }
  try {
    PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", p);
    CacheService.getScriptCache().remove(key);
    return { exito: true, mensaje: "Contraseña actualizada. Ya podés iniciar sesión con la nueva clave." };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

// ========== SERVIDOR WEB (con manejo de páginas) ==========
function doGet(e) {
  try {
    const page = e && e.parameter && e.parameter.page ? e.parameter.page : "index";
    if (page === "agenda") {
      return HtmlService.createHtmlOutputFromFile("agenda")
        .setTitle("Reservar turno - María Emilia Estética")
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (page === "adminReset") {
      const tpl = HtmlService.createTemplateFromFile("adminReset");
      tpl.token = e && e.parameter && e.parameter.token ? String(e.parameter.token) : "";
      return tpl
        .evaluate()
        .setTitle("Nueva contraseña — Administración")
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    return HtmlService.createHtmlOutputFromFile("index")
      .setTitle("María Emilia Estética")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput("<h1>Error</h1><p>" + error.toString() + "</p>");
  }
}

// ========== FUNCIONES DE BACKEND ==========
function obtenerServicios() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(HOJA_CONFIG);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange("A2:F" + lastRow).getValues();
    const servicios = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const nombre = row[0];
      if (!nombre || nombre.toString().trim() === "") continue;
      let inicio = row[4] ? row[4].toString().trim() : "09:00";
      let fin = row[5] ? row[5].toString().trim() : "18:00";
      if (inicio.match(/^\d{1,2}:\d{2}$/) && !inicio.match(/^\d{2}:/)) inicio = "0" + inicio;
      if (fin.match(/^\d{1,2}:\d{2}$/) && !fin.match(/^\d{2}:/)) fin = "0" + fin;
      servicios.push({
        nombre: nombre.toString().trim(),
        duracion: row[1] ? Number(row[1]) : 30,
        responsable: row[2] ? row[2].toString().trim() : "No asignado",
        capacidad: row[3] ? Number(row[3]) : 1,
        horarioInicio: inicio,
        horarioFin: fin
      });
    }
    return servicios;
  } catch (e) {
    return [];
  }
}

function normalizarFecha(fecha) {
  const tz = getZonaHoraria();
  if (!fecha) return "";
  if (fecha instanceof Date) {
    return Utilities.formatDate(fecha, tz, "yyyy-MM-dd");
  }
  if (typeof fecha === "string") {
    if (/^\d{4}-\d{2}-\d{2}$/.test(fecha)) return fecha;
    let d = new Date(fecha);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz, "yyyy-MM-dd");
    }
  }
  if (typeof fecha === "number") {
    let d = new Date(Math.round((fecha - 25569) * 86400000));
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz, "yyyy-MM-dd");
    }
  }
  return fecha.toString();
}

function normalizarHora(hora) {
  const tz = getZonaHoraria();
  if (!hora) return "";
  if (hora instanceof Date) return Utilities.formatDate(hora, tz, "HH:mm");
  if (typeof hora === "string") {
    let match = hora.match(/(\d{2}):(\d{2})/);
    if (match) return match[0];
    let match2 = hora.match(/(\d{1,2}):(\d{2})/);
    if (match2) return match2[1].padStart(2,'0') + ":" + match2[2];
  }
  return hora.toString();
}

function generarHorarios(inicioStr, finStr) {
  try {
    const horarios = [];
    const tz = getZonaHoraria();
    const inicioNormal = normalizarHora(inicioStr);
    const finNormal = normalizarHora(finStr);
    if (!inicioNormal || !finNormal) return [];
    const inicio = new Date(`2000-01-01T${inicioNormal}:00`);
    const fin = new Date(`2000-01-01T${finNormal}:00`);
    let actual = new Date(inicio);
    while (actual <= fin) {
      horarios.push(Utilities.formatDate(actual, tz, "HH:mm"));
      actual.setMinutes(actual.getMinutes() + 30);
    }
    return horarios;
  } catch (e) {
    return [];
  }
}

function obtenerHorariosDisponibles(servicioNombre, fecha) {
  try {
    const servicios = obtenerServicios();
    const servicio = servicios.find(s => s.nombre === servicioNombre);
    if (!servicio) return [];
    let horariosPosibles = generarHorarios(servicio.horarioInicio, servicio.horarioFin);
    if (horariosPosibles.length === 0) return [];
    const tz = getZonaHoraria();
    const hoyStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    if (fecha === hoyStr) {
      const ahora = new Date();
      const horaActual = Utilities.formatDate(ahora, tz, "HH:mm");
      horariosPosibles = horariosPosibles.filter(h => h >= horaActual);
    }
    const sheetTurnos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TURNOS);
    const todosTurnos = sheetTurnos.getDataRange().getValues();
    const ocupados = new Map();
    for (let i = 1; i < todosTurnos.length; i++) {
      const fechaTurno = normalizarFecha(todosTurnos[i][1]);
      const horaTurno = normalizarHora(todosTurnos[i][2]);
      const responsableTurno = todosTurnos[i][6];
      const servicioTurno = todosTurnos[i][5];
      const estado = todosTurnos[i][8] || "activo";
      if (fechaTurno === fecha && responsableTurno === servicio.responsable && servicioTurno === servicioNombre && estado !== "cancelado") {
        ocupados.set(horaTurno, (ocupados.get(horaTurno) || 0) + 1);
      }
    }
    const disponibles = horariosPosibles.filter(hora => (ocupados.get(hora) || 0) < servicio.capacidad);
    return disponibles;
  } catch (e) {
    return [];
  }
}

function esFechaHoraValida(fechaStr, horaStr) {
  if (!fechaStr || !horaStr) return false;
  const tz = getZonaHoraria();
  const hoy = new Date();
  const hoyStr = Utilities.formatDate(hoy, tz, "yyyy-MM-dd");
  if (fechaStr < hoyStr) return false;
  const fecha = new Date(fechaStr + "T12:00:00");
  if (fecha.getDay() === 0) return false;
  if (fechaStr === hoyStr) {
    const ahora = Utilities.formatDate(hoy, tz, "HH:mm");
    if (horaStr < ahora) return false;
  }
  return true;
}

function generarCodigo(longitud = 6) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let codigo = "";
  for (let i = 0; i < longitud; i++) codigo += chars.charAt(Math.floor(Math.random() * chars.length));
  return codigo;
}

function guardarTurno(datos) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return { exito: false, mensaje: "Sistema ocupado, intentá de nuevo." };
  try {
    const servicios = obtenerServicios();
    const servicio = servicios.find(s => s.nombre === datos.servicio);
    if (!servicio) throw new Error("Servicio no encontrado");
    if (!esFechaHoraValida(datos.fecha, datos.hora)) {
      lock.releaseLock();
      return { exito: false, mensaje: "Fecha u hora inválida (pasada, domingo o ya transcurrida)." };
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TURNOS);
    const todosTurnos = sheet.getDataRange().getValues();
    let ocupadosEnHora = 0;
    for (let i = 1; i < todosTurnos.length; i++) {
      const fechaTurno = normalizarFecha(todosTurnos[i][1]);
      const horaTurno = normalizarHora(todosTurnos[i][2]);
      const responsableTurno = todosTurnos[i][6];
      const servicioTurno = todosTurnos[i][5];
      const estado = todosTurnos[i][8] || "activo";
      if (fechaTurno === datos.fecha && horaTurno === datos.hora && responsableTurno === servicio.responsable && servicioTurno === datos.servicio && estado !== "cancelado") {
        ocupadosEnHora++;
      }
    }
    if (ocupadosEnHora >= servicio.capacidad) {
      lock.releaseLock();
      return { exito: false, mensaje: `No hay cupo para ${datos.servicio} a las ${datos.hora}. Capacidad: ${servicio.capacidad}.` };
    }
    const nuevoId = sheet.getLastRow() + 1;
    const codigoCancel = generarCodigo();
    sheet.appendRow([nuevoId, datos.fecha, datos.hora, datos.nombre, datos.telefono, datos.servicio, servicio.responsable, codigoCancel, "activo"]);
    // Formatear fecha legible
    const partes = datos.fecha.split('-');
    const fechaLegible = partes[2] + '/' + partes[1] + '/' + partes[0];

    // Email al CLIENTE
    if (datos.email) {
      const htmlCliente = `
        <div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;border:1px solid #eee;border-radius:12px;overflow:hidden;">
          <div style="background:linear-gradient(135deg,#F2D9DF,#E8D9F0);padding:32px;text-align:center;">
            <p style="font-size:12px;letter-spacing:3px;text-transform:uppercase;color:#A07830;margin:0 0 8px;">María Emilia Estética</p>
            <h1 style="font-family:Georgia,serif;font-size:28px;font-weight:300;color:#2C2420;margin:0;">✅ ¡Turno confirmado!</h1>
          </div>
          <div style="padding:32px;">
            <p style="color:#4A3F3A;font-size:15px;">Hola <strong>${datos.nombre}</strong>, tu turno fue reservado con éxito. Acá están los detalles:</p>
            <table style="width:100%;border-collapse:collapse;margin:20px 0;">
              <tr><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">💆 Servicio</td><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;color:#2C2420;">${datos.servicio}</td></tr>
              <tr><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">👩 Responsable</td><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;color:#2C2420;">${servicio.responsable}</td></tr>
              <tr><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📆 Fecha</td><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;color:#2C2420;">${fechaLegible}</td></tr>
              <tr><td style="padding:10px 0;color:#8A7A74;font-size:13px;">⏰ Hora</td><td style="padding:10px 0;font-weight:bold;color:#2C2420;">${datos.hora}</td></tr>
            </table>
            <div style="background:#fdf5f5;border-radius:10px;padding:20px;text-align:center;margin:20px 0;">
              <p style="color:#8A7A74;font-size:12px;margin:0 0 8px;">🔑 Tu código de cancelación</p>
              <p style="font-size:28px;font-weight:bold;letter-spacing:8px;color:#d46b6b;margin:0;">${codigoCancel}</p>
              <p style="color:#aaa;font-size:11px;margin:8px 0 0;">Guardalo para cancelar tu turno si lo necesitás.</p>
            </div>
            <p style="color:#8A7A74;font-size:13px;line-height:1.7;">Para cancelar, ingresá a la web y usá la sección <strong>❌ Cancelar turno</strong> con este código.</p>
          </div>
          <div style="background:#f9f0f0;padding:20px;text-align:center;border-top:1px solid #eee;">
            <p style="color:#aaa;font-size:11px;margin:0;">© María Emilia Estética · Ensenada · Bartolomé Bavio · Magdalena</p>
          </div>
        </div>`;
      try {
        MailApp.sendEmail({ to: datos.email, subject: `✅ Turno confirmado — ${datos.servicio} el ${fechaLegible}`, htmlBody: htmlCliente });
      } catch (e) {}
    }

    // Email al DUEÑO
    const htmlDuenio = `
      <div style="font-family:Arial,sans-serif;max-width:480px;margin:auto;border:1px solid #eee;border-radius:12px;overflow:hidden;">
        <div style="background:#2C2420;padding:24px;text-align:center;">
          <p style="color:#C9A84C;font-size:13px;letter-spacing:3px;text-transform:uppercase;margin:0;">Nuevo turno reservado</p>
        </div>
        <div style="padding:28px;">
          <table style="width:100%;border-collapse:collapse;">
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">👤 Cliente</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;">${datos.nombre}</td></tr>
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📞 Teléfono</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${datos.telefono}</td></tr>
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📧 Email</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${datos.email || '—'}</td></tr>
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">💆 Servicio</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${datos.servicio}</td></tr>
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">👩 Responsable</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${servicio.responsable}</td></tr>
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📆 Fecha</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${fechaLegible}</td></tr>
            <tr><td style="padding:8px 0;color:#8A7A74;font-size:13px;">⏰ Hora</td><td style="padding:8px 0;">${datos.hora}</td></tr>
          </table>
        </div>
      </div>`;
    try {
      MailApp.sendEmail({ to: EMAIL_DUENIO, subject: `📅 Nuevo turno — ${datos.nombre} · ${fechaLegible} ${datos.hora}`, htmlBody: htmlDuenio });
    } catch (e) {}
    lock.releaseLock();
    return { exito: true, mensaje: `Turno guardado con ${servicio.responsable}. Código: ${codigoCancel}`, codigo: codigoCancel };
  } catch (error) {
    lock.releaseLock();
    return { exito: false, mensaje: error.toString() };
  }
}

function cancelarTurno(codigo) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return { exito: false, mensaje: "Sistema ocupado, intentá de nuevo." };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TURNOS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][7] === codigo && data[i][8] !== "cancelado") {
        sheet.getRange(i + 1, 9).setValue("cancelado");
        return { exito: true, mensaje: "Turno cancelado correctamente." };
      }
    }
    return { exito: false, mensaje: "Código inválido." };
  } catch (error) {
    return { exito: false, mensaje: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function obtenerTurnosHoy(sessionToken) {
  try {
    if (!sesionAdminValida(sessionToken)) return { ok: false, error: "NO_AUTH" };
    const tz = getZonaHoraria();
    const hoyStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TURNOS);
    if (!sheet) return { ok: true, turnos: [] };
    const data = sheet.getDataRange().getValues();
    const turnos = [];
    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      let fechaTurnoStr = normalizarFecha(fila[1]);
      const estado = fila[8] || "activo";
      if (fechaTurnoStr === hoyStr && estado !== "cancelado") {
        turnos.push({
          id: String(fila[0]),
          fecha: fechaTurnoStr,
          hora: normalizarHora(fila[2]),
          nombre: String(fila[3] || ""),
          telefono: String(fila[4] || ""),
          servicio: String(fila[5] || ""),
          responsable: String(fila[6] || ""),
          codigo: String(fila[7] || "")
        });
      }
    }
    turnos.sort((a, b) => a.hora.localeCompare(b.hora));
    return { ok: true, turnos: turnos };
  } catch (e) {
    return { ok: false, error: "SERVER", mensaje: e.toString() };
  }
}

function obtenerTodosTurnosFuturos(sessionToken) {
  try {
    if (!sesionAdminValida(sessionToken)) return { ok: false, error: "NO_AUTH" };
    const tz = getZonaHoraria();
    const hoyStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TURNOS);
    if (!sheet) return { ok: true, turnos: [] };
    const data = sheet.getDataRange().getValues();
    const turnos = [];
    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      let fechaTurnoStr = normalizarFecha(fila[1]);
      const estado = fila[8] || "activo";
      if (fechaTurnoStr >= hoyStr && estado !== "cancelado") {
        turnos.push({
          id: String(fila[0]),
          fecha: fechaTurnoStr,
          hora: normalizarHora(fila[2]),
          nombre: String(fila[3] || ""),
          telefono: String(fila[4] || ""),
          servicio: String(fila[5] || ""),
          responsable: String(fila[6] || ""),
          codigo: String(fila[7] || "")
        });
      }
    }
    turnos.sort((a, b) => {
      if (a.fecha === b.fecha) return a.hora.localeCompare(b.hora);
      return a.fecha.localeCompare(b.fecha);
    });
    return { ok: true, turnos: turnos };
  } catch (e) {
    return { ok: false, error: "SERVER", mensaje: e.toString() };
  }
}
