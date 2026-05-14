import { Resend } from "resend";

export type TurnoMailPayload = {
  nombre: string;
  telefono: string;
  emailCliente: string | null;
  servicio: string;
  responsable: string;
  fecha: string;
  hora: string;
  codigoCancelacion: string;
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

function htmlDuenio(p: TurnoMailPayload): string {
  const nombre = escapeHtml(p.nombre);
  const telefono = escapeHtml(p.telefono);
  const email = p.emailCliente ? escapeHtml(p.emailCliente) : "—";
  const servicio = escapeHtml(p.servicio);
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
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">👩 Responsable</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${responsable}</td></tr>
            <tr><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;color:#8A7A74;font-size:13px;">📆 Fecha</td><td style="padding:8px 0;border-bottom:1px solid #f0e8e8;">${fechaL}</td></tr>
            <tr><td style="padding:8px 0;color:#8A7A74;font-size:13px;">⏰ Hora</td><td style="padding:8px 0;">${hora}</td></tr>
          </table>
        </div>
      </div>`;
}

/**
 * Envía los mismos correos que `guardarTurno` en `Código.gs` (errores solo en consola; no falla la reserva).
 */
export async function enviarMailsTurnoConfirmado(p: TurnoMailPayload): Promise<void> {
  const apiKey = process.env.RESEND_API_KEY?.trim();
  const from = process.env.EMAIL_FROM?.trim();
  const ownerEmail = process.env.OWNER_EMAIL?.trim();

  if (!apiKey || !from) {
    if (process.env.NODE_ENV === "development") {
      console.warn(
        "[mail] Sin RESEND_API_KEY o EMAIL_FROM: no se envían mails de turno."
      );
    }
    return;
  }

  const resend = new Resend(apiKey);
  const fechaL = fechaLegible(p.fecha);

  if (p.emailCliente) {
    const { error } = await resend.emails.send({
      from,
      to: p.emailCliente,
      subject: `✅ Turno confirmado — ${p.servicio} el ${fechaL}`,
      html: htmlCliente(p),
    });
    if (error) console.error("[mail] cliente:", error);
  }

  if (ownerEmail) {
    const { error } = await resend.emails.send({
      from,
      to: ownerEmail,
      subject: `📅 Nuevo turno — ${p.nombre} · ${fechaL} ${p.hora}`,
      html: htmlDuenio(p),
    });
    if (error) console.error("[mail] dueño:", error);
  } else if (process.env.NODE_ENV === "development") {
    console.warn("[mail] Sin OWNER_EMAIL: no se notifica al dueño.");
  }
}
