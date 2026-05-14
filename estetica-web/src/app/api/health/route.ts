import { NextResponse } from "next/server";
import { getNeonSql } from "@/lib/db";

/** Comprueba que las rutas API en Vercel respondan; si hay `DATABASE_URL`, prueba Neon. */
export async function GET() {
  const sql = getNeonSql();
  if (!sql) {
    return NextResponse.json({
      ok: true,
      service: "estetica-web",
      db: "not_configured",
      hint: "Definí DATABASE_URL (Neon, conexión pooled) en .env.local o en Vercel.",
    });
  }

  try {
    await sql`SELECT 1 AS ping`;
    return NextResponse.json({ ok: true, service: "estetica-web", db: "ok" });
  } catch {
    return NextResponse.json(
      { ok: false, service: "estetica-web", db: "error" },
      { status: 503 }
    );
  }
}
