import { NextRequest, NextResponse } from "next/server";
import { ADMIN_SESSION_COOKIE, verifySignedSessionValue } from "@/lib/admin-session";
import { cancelarTurnoPorCodigo } from "@/lib/turnos-repo";

export const dynamic = "force-dynamic";

export async function POST(request: NextRequest) {
  if (!verifySignedSessionValue(request.cookies.get(ADMIN_SESSION_COOKIE)?.value)) {
    return NextResponse.json({ ok: false, error: "no_autorizado" }, { status: 401 });
  }

  let body: { codigo?: string };
  try {
    body = (await request.json()) as { codigo?: string };
  } catch {
    return NextResponse.json({ ok: false, error: "json" }, { status: 400 });
  }

  const codigo = body.codigo?.trim().toUpperCase() ?? "";
  if (!codigo) {
    return NextResponse.json(
      { ok: false, mensaje: "Falta el código de cancelación." },
      { status: 400 }
    );
  }

  const res = await cancelarTurnoPorCodigo(codigo);
  if (!res.ok) {
    return NextResponse.json(
      {
        ok: false,
        mensaje: "No se encontró un turno activo con ese código.",
      },
      { status: 404 }
    );
  }

  return NextResponse.json({ ok: true, mensaje: "Turno cancelado." });
}
