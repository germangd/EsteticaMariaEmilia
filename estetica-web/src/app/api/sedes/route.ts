import { NextResponse } from "next/server";
import { listarSedesActivas } from "@/lib/sedes-repo";

export const dynamic = "force-dynamic";

export async function GET() {
  const rows = await listarSedesActivas();
  if (!Array.isArray(rows)) {
    return NextResponse.json(
      { ok: false, mensaje: "Base de datos no disponible." },
      { status: 503 }
    );
  }
  return NextResponse.json({ ok: true, sedes: rows });
}
