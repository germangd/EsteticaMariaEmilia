/** Número del salón (Argentina, sin +). Mismo que en la home si no hay env. */
export const WHATSAPP_NUMBER_DEFAULT = "5492215918286";

export function getWhatsAppNumber(): string {
  const raw =
    process.env.NEXT_PUBLIC_WHATSAPP_NUMBER?.trim() ||
    process.env.WHATSAPP_NUMBER?.trim();
  const digits = raw?.replace(/\D/g, "");
  return digits || WHATSAPP_NUMBER_DEFAULT;
}

function fechaLegible(fechaIso: string): string {
  const p = fechaIso.split("-");
  if (p.length !== 3) return fechaIso;
  return `${p[2]}/${p[1]}/${p[0]}`;
}

/** Link wa.me con mensaje de turno confirmado (el usuario debe tocar Enviar). */
export function buildWhatsAppTurnoUrl(params: {
  nombre: string;
  servicio: string;
  fecha: string;
  hora: string;
  codigo: string;
  pendienteAnticipo?: boolean;
  anticipoMontoPesos?: number;
  anticipoPorcentaje?: number;
}): string {
  const lineas = [
    params.pendienteAnticipo
      ? "Hola! Solicité un turno en María Emilia Estética (pendiente de confirmación por anticipo):"
      : "Hola! Acabo de reservar turno en María Emilia Estética:",
    `• ${params.servicio}`,
    `• ${fechaLegible(params.fecha)} ${params.hora}`,
    `• A nombre de: ${params.nombre}`,
    `• Código: ${params.codigo}`,
  ];
  if (params.pendienteAnticipo) {
    if (params.anticipoMontoPesos && params.anticipoMontoPesos > 0) {
      lineas.push(
        `• Anticipo a abonar: $${params.anticipoMontoPesos.toLocaleString("es-AR")}${params.anticipoPorcentaje ? ` (${params.anticipoPorcentaje}%)` : ""}`
      );
    } else if (params.anticipoPorcentaje) {
      lineas.push(`• Anticipo: ${params.anticipoPorcentaje}% del tratamiento`);
    }
    lineas.push("• Quiero coordinar el pago del anticipo.");
  } else {
    lineas.push(`• Código de cancelación: ${params.codigo}`);
  }
  const text = lineas.join("\n");

  return `https://wa.me/${getWhatsAppNumber()}?text=${encodeURIComponent(text)}`;
}
