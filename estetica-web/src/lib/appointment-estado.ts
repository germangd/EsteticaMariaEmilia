/** Estados de `appointments.estado`. */
export const ESTADO_TURNO = {
  ACTIVO: "activo",
  PENDIENTE_ANTICIPO: "pendiente_anticipo",
  CANCELADO: "cancelado",
} as const;

export type EstadoTurno = (typeof ESTADO_TURNO)[keyof typeof ESTADO_TURNO];

/** Estados que ocupan cupo en la agenda. */
export const ESTADOS_OCUPAN_CUPO: EstadoTurno[] = [
  ESTADO_TURNO.ACTIVO,
  ESTADO_TURNO.PENDIENTE_ANTICIPO,
];

export function etiquetaEstadoTurno(estado: string): string {
  if (estado === ESTADO_TURNO.PENDIENTE_ANTICIPO) {
    return "Pendiente anticipo";
  }
  if (estado === ESTADO_TURNO.ACTIVO) return "Confirmado";
  if (estado === ESTADO_TURNO.CANCELADO) return "Cancelado";
  return estado;
}

export function turnoEsConfirmado(estado: string): boolean {
  return estado === ESTADO_TURNO.ACTIVO;
}

export function turnoEsPendienteAnticipo(estado: string): boolean {
  return estado === ESTADO_TURNO.PENDIENTE_ANTICIPO;
}
