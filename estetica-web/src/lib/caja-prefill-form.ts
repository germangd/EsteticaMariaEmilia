import type { PrefillCobroPaquete, PrefillCobroTurno } from "@/lib/caja-repo";
import { fmtPesos } from "@/lib/fmt-pesos";

export type LineaColaPrefill = {
  key: string;
  tipo: "servicio" | "paquete" | "otro";
  descripcion: string;
  cantidad: number;
  precioUnitarioPesos: number;
  serviceId?: number;
  servicePackageId?: number;
};

export function lineasColaDesdePrefillTurno(
  prefill: PrefillCobroTurno
): LineaColaPrefill[] {
  if (prefill.yaCobrado) return [];
  return [
    {
      key: `turno-${prefill.appointmentId}`,
      tipo: "servicio",
      descripcion: `${prefill.servicioNombre} (${prefill.fecha} ${prefill.hora})`,
      cantidad: 1,
      precioUnitarioPesos: prefill.importeCobroSugeridoPesos,
      serviceId: prefill.serviceId ?? undefined,
    },
  ];
}

export function lineasColaDesdePrefillPaquete(
  prefill: PrefillCobroPaquete
): LineaColaPrefill[] {
  if (prefill.yaCobrado) return [];
  return [
    {
      key: `paquete-${prefill.clientPackageId}`,
      tipo: "paquete",
      descripcion: `Paquete: ${prefill.paqueteNombre} (desde ${prefill.fechaCompra})`,
      cantidad: 1,
      precioUnitarioPesos: prefill.precioSugeridoPesos,
      servicePackageId: prefill.packageId,
    },
  ];
}

export function mensajePrefillTurno(prefill: PrefillCobroTurno): string {
  const ant = prefill.anticipoRequerido
    ? prefill.anticipoSugeridoPesos > 0
      ? ` Anticipo configurado: ${prefill.anticipoPorcentaje}% (${fmtPesos(prefill.anticipoSugeridoPesos)}).`
      : ` Anticipo configurado: ${prefill.anticipoPorcentaje}% (sin precio de referencia).`
    : "";

  if (prefill.totalAbonadoPesos > 0 && prefill.saldoPendientePesos > 0) {
    return `Saldo del turno #${prefill.appointmentId}: ${fmtPesos(prefill.saldoPendientePesos)} (ya abonado ${fmtPesos(prefill.totalAbonadoPesos)} de ${fmtPesos(prefill.precioSugeridoPesos)}). Revisá el importe y confirmá.${ant}`;
  }

  if (prefill.pendienteAnticipo) {
    return `Cobro de anticipo del turno #${prefill.appointmentId}: revisá el importe y confirmá. Luego confirmá el turno en Agenda.${ant}`;
  }

  if (prefill.importeCobroSugeridoPesos > 0) {
    return `Cobro del turno #${prefill.appointmentId}: revisá el importe sugerido y confirmá.${ant}`;
  }

  return `Cobro del turno #${prefill.appointmentId}: indicá el importe y confirmá.${ant}`;
}
