import { getBusinessName, getWhatsAppNumberDigits } from "@/config/site";

export function getWhatsAppNumber(): string {
  return getWhatsAppNumberDigits();
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
  const negocio = getBusinessName();
  const lineas = [
    params.pendienteAnticipo
      ? `Hola! Solicité un turno en ${negocio} (pendiente de confirmación por anticipo):`
      : `Hola! Acabo de reservar turno en ${negocio}:`,
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
