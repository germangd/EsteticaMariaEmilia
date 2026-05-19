/** Claves de ítem en reserva web (`s:` servicio, `p:` paquete). Sin dependencias de servidor. */

export function claveReservaServicio(nombre: string): string {
  return `s:${nombre}`;
}

export function claveReservaPaquete(id: number): string {
  return `p:${id}`;
}

export function parsearClaveReserva(
  clave: string
): { tipo: "servicio"; nombre: string } | { tipo: "paquete"; id: number } | null {
  const t = clave.trim();
  if (t.startsWith("s:")) {
    const nombre = t.slice(2).trim();
    return nombre ? { tipo: "servicio", nombre } : null;
  }
  if (t.startsWith("p:")) {
    const id = Number(t.slice(2));
    if (Number.isFinite(id) && id > 0) return { tipo: "paquete", id };
  }
  return null;
}
