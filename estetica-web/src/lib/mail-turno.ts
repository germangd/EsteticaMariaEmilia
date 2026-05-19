import nodemailer from "nodemailer";
import { Resend } from "resend";

export type TurnoMailPayload = {
  nombre: string;
  telefono: string;
  emailCliente: string | null;
  servicio: string;
  sede?: string | null;
  responsable: string;
  fecha: string;
  hora: string;
  codigoCancelacion: string;
  anticipoPorcentaje?: number;
  anticipoMontoPesos?: number;
};

type MailSendOpts = {
  from: string;
  to: string;
  subject: string;
  html: string;
};

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fechaLegible(fechaIso: string): string {
  const partes = fechaIso.split("-");
  if (partes.length !== 3) return fechaIso;
  return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

function htmlCliente(p: TurnoMailPayload): string {
  const nombre = escapeHtml(p.nombre);
  const servicio = escapeHtml(p.servicio);
  const sede = p.sede ? escapeHtml(p.sede) : "";
  const responsable = escapeHtml(p.responsable);
  const fechaL = escapeHtml(fechaLegible(p.fecha));
  const hora = escapeHtml(p.hora);
  const codigo = escapeHtml(p.codigoCancelacion);
  return `
        <div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;border:1px solid #eee;border-radius:12px;overflow:hidden;">
          <div style="background:linear-gradient(135deg,#F2D9DF,#E8D9F0);padding:32px;text-align:center;">
            <p style="font-size:12px;letter-spacing:3px;text-transform:uppercase;color:#A07830;margin:0 0 8px;">María Emilia Estética</p>
            <h1 style="font-family:Georgia,serif;font-size:28px;font-weight:300;color:#2C2420;margin:0;">✅ ¡Turno confirmado!</h1>
          </div>
          <div style="padding:32px;">
            <p style="color:#4A3F3A;font-size:15px;">Hola <strong>${nombre}</strong>, tu turno fue reservado con éxito. Acá están los detalles:</p>
            <table style="width:100%;border-collapse:collapse;margin:20px 0;">
              <tr><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">💆 Servicio</td><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;color:#2C2420;">${servicio}</td></tr>
              ${sede ? `<tr><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📍 Sede</td><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;color:#2C2420;">${sede}</td></tr>` : ""}
              <tr><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">👩 Responsable</td><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;color:#2C2420;">${responsable}</td></tr>
              <tr><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📆 Fecha</td><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;color:#2C2420;">${fechaL}</td></tr>
              <tr><td style="padding:10px 0;color:#8A7A74;font-size:13px;">⏰ Hora</td><td style="padding:10px 0;font-weight:bold;color:#2C2420;">${hora}</td></tr>
            </table>
            <div style="background:#fdf5f5;border-radius:10px;padding:20px;text-align:center;margin:20px 0;">
              <p style="color:#8A7A74;font-size:12px;margin:0 0 8px;">🔑 Tu código de cancelación</p>
              <p style="font-size:28px;font-weight:bold;letter-spacing:8px;color:#d46b6b;margin:0;">${codigo}</p>
              <p style="color:#aaa;font-size:11px;margin:8px 0 0;">Guardalo para cancelar tu turno si lo necesitás.</p>
            </div>
            <p style="color:#8A7A74;font-size:13px;line-height:1.7;">Para cancelar, ingresá a la web y usá la sección <strong>❌ Cancelar turno</strong> con este código.</p>
          </div>
          <div style="background:#f9f0f0;padding:20px;text-align:center;border-top:1px solid #eee;">
            <p style="color:#aaa;font-size:11px;margin:0;">© María Emilia Estética · Ensenada · Bartolomé Bavio · Magdalena</p>
          </div>
        </div>`;
}

function htmlPendienteCliente(p: TurnoMailPayload): string {
  const nombre = escapeHtml(p.nombre);
  const servicio = escapeHtml(p.servicio);
  const sede = p.sede ? escapeHtml(p.sede) : "";
  const fechaL = escapeHtml(fechaLegible(p.fecha));
  const hora = escapeHtml(p.hora);
  const codigo = escapeHtml(p.codigoCancelacion);
  const anticipo =
    p.anticipoMontoPesos && p.anticipoMontoPesos > 0
      ? `$${p.anticipoMontoPesos.toLocaleString("es-AR")}`
      : p.anticipoPorcentaje
        ? `${p.anticipoPorcentaje}% del tratamiento`
        : "según lo acordado";
  return `
        <div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;border:1px solid #eee;border-radius:12px;overflow:hidden;">
          <div style="background:linear-gradient(135deg,#F2D9DF,#E8D9F0);padding:32px;text-align:center;">
            <p style="font-size:12px;letter-spacing:3px;text-transform:uppercase;color:#A07830;margin:0 0 8px;">María Emilia Estética</p>
            <h1 style="font-family:Georgia,serif;font-size:26px;font-weight:300;color:#2C2420;margin:0;">Solicitud de turno recibida</h1>
          </div>
          <div style="padding:32px;">
            <p style="color:#4A3F3A;font-size:15px;">Hola <strong>${nombre}</strong>, registramos tu solicitud. El turno queda <strong>pendiente de confirmación</strong> hasta abonar el anticipo.</p>
            <table style="width:100%;border-collapse:collapse;margin:20px 0;">
              <tr><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">Servicio</td><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;color:#2C2420;">${servicio}</td></tr>
              ${sede ? `<tr><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">Sede</td><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;color:#2C2420;">${sede}</td></tr>` : ""}
              <tr><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">Fecha</td><td style="padding:10px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;color:#2C2420;">${fechaL}</td></tr>
              <tr><td style="padding:10px 0;color:#8A7A74;font-size:13px;">Hora</td><td style="padding:10px 0;font-weight:bold;color:#2C2420;">${hora}</td></tr>
            </table>
            <div style="background:#fff8e6;border-radius:10px;padding:16px;margin:16px 0;border:1px solid #e8d4a8;">
              <p style="margin:0;font-size:14px;color:#5c4a20;"><strong>Anticipo:</strong> ${escapeHtml(anticipo)}. Coordiná el pago con el salón (WhatsApp o en local). Al confirmarlo, te enviaremos la confirmación definitiva.</p>
            </div>
            <p style="color:#8A7A74;font-size:12px;">Código de solicitud: <strong>${codigo}</strong></p>
          </div>
        </div>`;
}

function htmlPendienteDuenio(p: TurnoMailPayload): string {
  const nombre = escapeHtml(p.nombre);
  const telefono = escapeHtml(p.telefono);
  const servicio = escapeHtml(p.servicio);
  const fechaL = escapeHtml(fechaLegible(p.fecha));
  const hora = escapeHtml(p.hora);
  const codigo = escapeHtml(p.codigoCancelacion);
  const anticipo =
    p.anticipoMontoPesos && p.anticipoMontoPesos > 0
      ? `$${p.anticipoMontoPesos.toLocaleString("es-AR")}`
      : p.anticipoPorcentaje
        ? `${p.anticipoPorcentaje}%`
        : "—";
  return `
        <div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;padding:20px;">
          <h2 style="color:#A07830;">Turno pendiente de anticipo</h2>
          <p><strong>${nombre}</strong> · ${telefono}</p>
          <p>${servicio} · ${fechaL} ${hora}</p>
          <p>Anticipo: <strong>${anticipo}</strong></p>
          <p>Código: <strong>${codigo}</strong></p>
          <p style="color:#666;font-size:13px;">Confirmá el turno en Admin → Turnos cuando recibas el pago.</p>
        </div>`;
}

function htmlDuenio(p: TurnoMailPayload): string {
  const nombre = escapeHtml(p.nombre);
  const telefono = escapeHtml(p.telefono);
  const email = p.emailCliente ? escapeHtml(p.emailCliente) : "—";
  const servicio = escapeHtml(p.servicio);
  const sede = p.sede ? escapeHtml(p.sede) : "";
  const responsable = escapeHtml(p.responsable);
  const fechaL = escapeHtml(fechaLegible(p.fecha));
  const hora = escapeHtml(p.hora);
  return `
      <div style="font-family:Arial,sans-serif;max-width:480px;margin:auto;border:1px solid #eee;border-radius:12px;overflow:hidden;">
        <div style="background:#2C2420;padding:24px;text-align:center;">
          <p style="color:#C9A84C;font-size:13px;letter-spacing:3px;text-transform:uppercase;margin:0;">Nuevo turno reservado</p>
        </div>
        <div style="padding:28px;">
          <table style="width:100%;border-collapse:collapse;">
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">👤 Cliente</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;font-weight:bold;">${nombre}</td></tr>
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📞 Teléfono</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${telefono}</td></tr>
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📧 Email</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${email}</td></tr>
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">💆 Servicio</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${servicio}</td></tr>
            ${sede ? `<tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📍 Sede</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${sede}</td></tr>` : ""}
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">👩 Responsable</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${responsable}</td></tr>
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📆 Fecha</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${fechaL}</td></tr>
            <tr><td style="padding:8px 0;color:#8A7A74;font-size:13px;">⏰ Hora</td><td style="padding:8px 0;">${hora}</td></tr>
          </table>
        </div>
      </div>`;
}

function envVar(name: string): string | undefined {
  const raw = process.env[name];
  if (raw == null) return undefined;
  const v = raw.trim().replace(/^["']|["']$/g, "");
  return v || undefined;
}

function isVercelRuntime(): boolean {
  return Boolean(process.env.VERCEL);
}

function smtpConfigured(): boolean {
  return Boolean(
    process.env.SMTP_USER?.trim() && process.env.SMTP_PASS?.trim()
  );
}

/** Gmail SMTP solo en local; Vercel bloquea puertos SMTP salientes. */
function canUseSmtp(): boolean {
  return smtpConfigured() && !isVercelRuntime();
}

function resendConfigured(): boolean {
  return Boolean(envVar("RESEND_API_KEY") && envVar("EMAIL_FROM"));
}

function mailProvider(): "smtp" | "resend" | null {
  if (canUseSmtp()) return "smtp";
  if (resendConfigured()) return "resend";
  return null;
}

function logMail(msg: string): void {
  if (process.env.NODE_ENV === "development" || isVercelRuntime()) {
    console.warn(`[mail] ${msg}`);
  }
}

/** Diagnóstico seguro (sin exponer la clave) para logs y /api/health. */
export function resendKeyDiagnostics(): {
  present: boolean;
  formatOk: boolean;
  length: number;
  hasInnerWhitespace: boolean;
} {
  const key = envVar("RESEND_API_KEY");
  return {
    present: Boolean(key),
    formatOk: Boolean(key?.startsWith("re_") && key.length >= 24),
    length: key?.length ?? 0,
    hasInnerWhitespace: Boolean(
      key && (key.includes(" ") || key.includes("\n") || key.includes("\r"))
    ),
  };
}

export function resendFromDiagnostics(): {
  set: boolean;
  fromOk: boolean;
  hint: string | null;
} {
  const from = envVar("EMAIL_FROM");
  if (!from) {
    return {
      set: false,
      fromOk: false,
      hint: "Definí EMAIL_FROM (ej. María Emilia Estética <onboarding@resend.dev>).",
    };
  }
  const invalid = resendFromAddressInvalid(from);
  return {
    set: true,
    fromOk: !invalid,
    hint: invalid,
  };
}

function logResendKeyHint(): void {
  const d = resendKeyDiagnostics();
  console.error(
    `[mail] Resend rechazó la API key. Diagnóstico: presente=${d.present}, formatoOk=${d.formatOk}, largo=${d.length}, espaciosInternos=${d.hasInnerWhitespace}. Creá una key nueva en resend.com/api-keys, pegala en Vercel como RESEND_API_KEY (Production), redeploy.`
  );
}

/** Remitente visible (Gmail u otro SMTP). */
function resolveFromAddress(): string | null {
  const from = envVar("EMAIL_FROM");
  if (from) return from;
  const user = envVar("SMTP_USER");
  if (user) return `María Emilia Estética <${user}>`;
  return null;
}

async function sendViaSmtp(opts: MailSendOpts): Promise<void> {
  const port = Number(process.env.SMTP_PORT) || 587;
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST?.trim() || "smtp.gmail.com",
    port,
    secure: port === 465,
    auth: {
      user: process.env.SMTP_USER!.trim(),
      pass: process.env.SMTP_PASS!.trim(),
    },
  });
  await transporter.sendMail({
    from: opts.from,
    to: opts.to,
    subject: opts.subject,
    html: opts.html,
  });
}

const RESEND_BLOCKED_FROM_DOMAINS = [
  "gmail.com",
  "googlemail.com",
  "hotmail.com",
  "outlook.com",
  "yahoo.com",
  "live.com",
];

function resendFromAddressInvalid(from: string): string | null {
  const match = from.match(/<([^>]+)>/);
  const email = (match ? match[1] : from).trim().toLowerCase();
  const domain = email.split("@")[1];
  if (!domain) return "EMAIL_FROM debe incluir un correo válido.";
  if (RESEND_BLOCKED_FROM_DOMAINS.includes(domain)) {
    return `EMAIL_FROM no puede ser @${domain} con Resend. Usá onboarding@resend.dev (prueba) o un dominio verificado en resend.com/domains. OWNER_EMAIL puede seguir siendo tu Gmail.`;
  }
  return null;
}

async function sendViaResend(opts: MailSendOpts): Promise<void> {
  const fromError = resendFromAddressInvalid(opts.from);
  if (fromError) {
    logMail(fromError);
    throw new Error(fromError);
  }

  const apiKey = envVar("RESEND_API_KEY");
  if (!apiKey?.startsWith("re_")) {
    throw new Error(
      "RESEND_API_KEY inválida: debe empezar con re_ (copiala desde resend.com → API Keys)"
    );
  }
  const resend = new Resend(apiKey);
  const { error } = await resend.emails.send({
    from: opts.from,
    to: opts.to,
    subject: opts.subject,
    html: opts.html,
  });
  if (error) {
    const msg =
      typeof error === "object" &&
      error !== null &&
      "message" in error &&
      String((error as { message: unknown }).message).includes("invalid");
    if (msg) logResendKeyHint();
    throw error;
  }
}

async function sendMail(opts: MailSendOpts): Promise<void> {
  const provider = mailProvider();
  if (provider === "smtp") {
    await sendViaSmtp(opts);
    return;
  }
  if (provider === "resend") {
    await sendViaResend(opts);
    return;
  }
  throw new Error("mail_not_configured");
}

/**
 * Mails al reservar. Prioridad: SMTP (Gmail) si hay SMTP_USER+SMTP_PASS; si no, Resend.
 * Errores solo en consola; no falla la reserva.
 */
export async function enviarMailsTurnoConfirmado(p: TurnoMailPayload): Promise<void> {
  const from = resolveFromAddress();
  const ownerEmail = envVar("OWNER_EMAIL");
  const provider = mailProvider();

  if (!from) {
    logMail("Falta EMAIL_FROM o SMTP_USER para el remitente.");
    return;
  }

  if (isVercelRuntime() && smtpConfigured() && !resendConfigured()) {
    logMail(
      "En Vercel no funciona SMTP/Gmail. Agregá RESEND_API_KEY y un EMAIL_FROM de Resend (ej. onboarding@resend.dev), redeploy."
    );
    return;
  }

  if (!provider) {
    logMail("Configurá SMTP (local) o RESEND_API_KEY + EMAIL_FROM (Vercel).");
    return;
  }

  const fechaL = fechaLegible(p.fecha);
  const via = provider;

  if (p.emailCliente) {
    try {
      await sendMail({
        from,
        to: p.emailCliente,
        subject: `✅ Turno confirmado — ${p.servicio} el ${fechaL}`,
        html: htmlCliente(p),
      });
    } catch (err) {
      console.error(`[mail:${via}] cliente (${p.emailCliente}):`, err);
      if (via === "resend" && envVar("EMAIL_FROM")?.includes("onboarding@resend.dev")) {
        logMail(
          "Con onboarding@resend.dev no se envía a correos arbitrarios del cliente. Verificá un dominio en resend.com/domains y usá EMAIL_FROM tipo turnos@tudominio.com."
        );
      }
    }
  }

  if (ownerEmail) {
    try {
      await sendMail({
        from,
        to: ownerEmail,
        subject: `📅 Nuevo turno — ${p.nombre} · ${fechaL} ${p.hora}`,
        html: htmlDuenio(p),
      });
    } catch (err) {
      console.error(`[mail:${via}] dueño:`, err);
    }
  } else if (process.env.NODE_ENV === "development") {
    console.warn("[mail] Sin OWNER_EMAIL: no se notifica al dueño.");
  }
}

/** Aviso de solicitud pendiente de anticipo (no es confirmación definitiva). */
export async function enviarMailsTurnoPendienteAnticipo(
  p: TurnoMailPayload
): Promise<void> {
  const from = resolveFromAddress();
  const ownerEmail = envVar("OWNER_EMAIL");
  const provider = mailProvider();

  if (!from || !provider) return;

  const fechaL = fechaLegible(p.fecha);
  const via = provider;

  if (p.emailCliente) {
    try {
      await sendMail({
        from,
        to: p.emailCliente,
        subject: `⏳ Turno pendiente — anticipo ${p.servicio} · ${fechaL}`,
        html: htmlPendienteCliente(p),
      });
    } catch (err) {
      console.error(`[mail:${via}] cliente pendiente:`, err);
    }
  }

  if (ownerEmail) {
    try {
      await sendMail({
        from,
        to: ownerEmail,
        subject: `⏳ Pendiente anticipo — ${p.nombre} · ${fechaL} ${p.hora}`,
        html: htmlPendienteDuenio(p),
      });
    } catch (err) {
      console.error(`[mail:${via}] dueño pendiente:`, err);
    }
  }
}
