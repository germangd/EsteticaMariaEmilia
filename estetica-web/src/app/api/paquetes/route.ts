import { NextResponse } from "next/server";
import { listarPaquetesPublicos } from "@/lib/paquetes-repo";
import { resolverItemReserva } from "@/lib/reserva-catalogo";

export const dynamic = "force-dynamic";

/** Combos / paquetes activos para reserva web. */
export async function GET() {
  const rows = await listarPaquetesPublicos();
  if (!Array.isArray(rows)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible.", paquetes: [] },
      { status: 503 }
    );
  }

  const paquetes = [];
  for (const p of rows) {
    const item = await resolverItemReserva({ tipo: "paquete", id: p.id });
    if (!item) continue;
    paquetes.push({
      id: p.id,
      nombre: p.nombre,
      descripcion: p.descripcion,
      precioPesos: p.precioPesos,
      sesionesTotal: p.sesionesTotal,
      serviciosIncluidos: p.serviciosIncluidos,
      duracion: item.duracionMin,
      responsable: item.responsable,
      capacidad: item.capacidad,
    });
  }

  return NextResponse.json({ ok: true, paquetes });
}
