/** URL para cobrar un turno desde la agenda. */
export function urlCobrarTurno(appointmentId: number): string {
  return `/admin/caja?turno=${appointmentId}#nueva-venta`;
}

/** URL para cobrar una asignación de paquete desde Paquetes. */
export function urlCobrarPaquete(clientPackageId: number): string {
  return `/admin/caja?paquete=${clientPackageId}#nueva-venta`;
}

/** URL del ticket de una venta. */
export function urlTicketVenta(ventaId: number): string {
  return `/admin/caja/ticket/${ventaId}`;
}

/** Descarga CSV del historial de caja con filtros opcionales. */
export function urlExportHistorialCaja(params?: {
  desde?: string;
  hasta?: string;
  cliente?: string;
}): string {
  const q = new URLSearchParams();
  if (params?.desde) q.set("desde", params.desde);
  if (params?.hasta) q.set("hasta", params.hasta);
  if (params?.cliente?.trim()) q.set("cliente", params.cliente.trim());
  const qs = q.toString();
  return `/api/admin/caja/historial/export${qs ? `?${qs}` : ""}`;
}
