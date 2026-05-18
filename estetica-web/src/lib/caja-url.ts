/** URL para cobrar un turno desde la agenda. */
export function urlCobrarTurno(appointmentId: number): string {
  return `/admin/caja?turno=${appointmentId}#nueva-venta`;
}

/** URL del ticket de una venta. */
export function urlTicketVenta(ventaId: number): string {
  return `/admin/caja/ticket/${ventaId}`;
}
