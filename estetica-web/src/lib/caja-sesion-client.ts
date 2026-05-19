/** Mensajes y helpers para validar caja abierta en el panel (cliente). */

export const MENSAJE_CAJA_NO_ABIERTA =
  "No hay caja abierta. Abrí la caja (Estado de caja → Abrir caja) antes de registrar un cobro o una venta.";

export function alertarCajaNoAbierta(): void {
  if (typeof window !== "undefined") {
    window.alert(MENSAJE_CAJA_NO_ABIERTA);
  }
}

export async function haySesionCajaAbierta(): Promise<boolean> {
  try {
    const r = await fetch("/api/admin/caja/sesion", {
      credentials: "same-origin",
      cache: "no-store",
    });
    const data = (await r.json()) as { ok?: boolean; sesion?: { id: number } | null };
    return Boolean(r.ok && data.ok && data.sesion?.id);
  } catch {
    return false;
  }
}

export function esErrorCajaCerrada(mensaje: string | undefined): boolean {
  if (!mensaje) return false;
  const m = mensaje.toLowerCase();
  return (
    m.includes("abrí la caja") ||
    m.includes("abri la caja") ||
    m.includes("caja está cerrada") ||
    m.includes("caja esta cerrada") ||
    m.includes("no hay caja")
  );
}
