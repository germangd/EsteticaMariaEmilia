import type { ServicioInput } from "@/lib/servicios-repo";

/** Parsea body JSON de crear/actualizar servicio (admin). */
export function parseServicioBody(body: unknown): ServicioInput | null {
  if (!body || typeof body !== "object") return null;
  const b = body as Record<string, unknown>;

  const nombre = typeof b.nombre === "string" ? b.nombre : "";
  const tipo =
    b.tipo === "grupo" || b.tipo === "sub" || b.tipo === "suelto" ? b.tipo : null;
  const esGrupo = b.esGrupo === true || tipo === "grupo";
  const esSub = tipo === "sub" || (!esGrupo && b.parentId != null && b.parentId !== "");

  const duracionMin = Number(b.duracionMin ?? b.duracion);
  const responsable = typeof b.responsable === "string" ? b.responsable : "";
  const capacidad = Number(b.capacidad);
  const horarioInicio =
    typeof b.horarioInicio === "string" ? b.horarioInicio : "09:00";
  const horarioFin = typeof b.horarioFin === "string" ? b.horarioFin : "18:00";
  const precioPesos =
    b.precioPesos != null
      ? Number(b.precioPesos)
      : b.precio != null
        ? Number(b.precio)
        : 0;
  const anticipoRequerido = b.anticipoRequerido === true;
  const anticipoPorcentaje = Number(b.anticipoPorcentaje ?? 0);

  let parentId: number | null = null;
  if (esSub) {
    const pid = Number(b.parentId);
    if (!Number.isFinite(pid) || pid < 1) return null;
    parentId = pid;
  }

  if (!nombre.trim()) return null;

  if (esGrupo) {
    return {
      nombre,
      duracionMin: 0,
      responsable: responsable || "No asignado",
      capacidad: 0,
      horarioInicio,
      horarioFin,
      precioPesos: 0,
      parentId: null,
      esGrupo: true,
    };
  }

  if (!Number.isFinite(duracionMin) || !Number.isFinite(capacidad)) return null;

  if (anticipoRequerido) {
    if (
      !Number.isFinite(anticipoPorcentaje) ||
      anticipoPorcentaje < 1 ||
      anticipoPorcentaje > 100
    ) {
      return null;
    }
  }

  return {
    nombre,
    duracionMin,
    responsable,
    capacidad,
    horarioInicio,
    horarioFin,
    precioPesos: Number.isFinite(precioPesos) ? precioPesos : 0,
    parentId,
    esGrupo: false,
    anticipoRequerido,
    anticipoPorcentaje: anticipoRequerido ? anticipoPorcentaje : 0,
  };
}

export function mensajeErrorServicio(
  reason: string | undefined,
  conflicto?: { id: number; nombre: string } | null
): string {
  switch (reason) {
    case "duplicado":
      return conflicto
        ? `Ya existe «${conflicto.nombre}» en esta categoría. No podés renombrar a ese nombre: editá el registro existente o eliminá el duplicado (por ejemplo el que tiene el error de tipeo).`
        : "Ya existe un servicio con ese nombre en la misma categoría.";
    case "parent_invalido":
      return "Elegí una categoría válida (servicio tipo grupo).";
    case "tiene_hijos":
      return "Esta categoría tiene sub-servicios. Eliminalos o reasignalos antes.";
    case "invalido":
      return "Datos incompletos o inválidos.";
    case "not_found":
      return "Servicio no encontrado.";
    default:
      return "No se pudo guardar el servicio.";
  }
}
