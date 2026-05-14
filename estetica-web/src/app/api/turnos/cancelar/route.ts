import { NextResponse } from "next/server";
import { cancelarTurnoPorCodigo } from "@/lib/turnos-repo";

export const dynamic = "force-dynamic";

/**
 * Cancelar turno por código (paridad con `cancelarTurno` en `Código.gs`).
 * Body JSON: { "codigo": "ABC123" }
 */
export async function POST(request: Request) {
  let codigo = "";
  try {
    const body = (await request.json()) as { codigo?: string };
    codigo = body.codigo?.trim() ?? "";
  } catch {
    return NextResponse.json(
      { exito: false, mensaje: "JSON inválido." },
      { status: 400 }
    );
  }

  if (!codigo) {
    return NextResponse.json(
      { exito: false, mensaje: "Falta el código de cancelación." },
      { status: 400 }
    );
  }

  const c = codigo.toUpperCase();
  const res = await cancelarTurnoPorCodigo(c);

  if (res.ok === false) {
    if (res.reason === "no_db") {
      return NextResponse.json(
        { exito: false, mensaje: "Base de datos no configurada." },
        { status: 503 }
      );
    }
    return NextResponse.json({
      exito: false,
      mensaje: "Código no válido o turno ya cancelado.",
    });
  }

  return NextResponse.json({
    exito: true,
    mensaje: "Turno cancelado correctamente.",
  });
}
