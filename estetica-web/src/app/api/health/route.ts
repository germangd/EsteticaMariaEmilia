import { NextResponse } from "next/server";
import { getNeonSql } from "@/lib/db";
import { resendFromDiagnostics, resendKeyDiagnostics } from "@/lib/mail-turno";

/** Comprueba que las rutas API en Vercel respondan; si hay `DATABASE_URL`, prueba Neon. */
export async function GET() {
  const fromDiag = resendFromDiagnostics();
  const mail = {
    healthVersion: 2,
    onVercel: Boolean(process.env.VERCEL),
    resend: resendKeyDiagnostics(),
    emailFrom: fromDiag,
  };

  const sql = getNeonSql();
  if (!sql) {
    return NextResponse.json({
      ok: true,
      service: "estetica-web",
      db: "not_configured",
      mail,
      hint: "Definí DATABASE_URL (Neon, conexión pooled) en .env.local o en Vercel.",
    });
  }

  try {
    await sql`SELECT 1 AS ping`;
    return NextResponse.json({ ok: true, service: "estetica-web", db: "ok", mail });
  } catch {
    return NextResponse.json(
      { ok: false, service: "estetica-web", db: "error", mail },
      { status: 503 }
    );
  }
}
