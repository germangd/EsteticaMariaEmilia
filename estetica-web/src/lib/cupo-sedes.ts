/**
 * Cupo de turnos entre sedes.
 *
 * Por defecto (sin variable): agenda única — si hay turno en una sede,
 * ese horario no se ofrece en las demás (mismo responsable).
 *
 * Futuro: en Vercel / `.env.local` definí
 * `CUPO_INDEPENDIENTE_POR_SEDE=true` para permitir turnos simultáneos
 * en sedes distintas.
 */
export function cupoCompartidoEntreSedes(): boolean {
  const v = process.env.CUPO_INDEPENDIENTE_POR_SEDE?.trim().toLowerCase();
  return v !== "true" && v !== "1";
}
