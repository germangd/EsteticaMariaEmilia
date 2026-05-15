import { NextRequest, NextResponse } from "next/server";
import {
  ADMIN_SESSION_COOKIE,
  adminAuthConfigured,
  createSignedSessionValue,
  sessionCookieOptions,
  verifyAdminPasswordAttempt,
} from "@/lib/admin-session";

export const dynamic = "force-dynamic";

export async function POST(request: NextRequest) {
  if (!adminAuthConfigured()) {
    return NextResponse.json(
      {
        ok: false,
        error: "not_configured",
        mensaje:
          "Falta ADMIN_SESSION_SECRET (o ADMIN_SECRET) y ADMIN_PASSWORD en el servidor.",
      },
      { status: 503 }
    );
  }

  let body: { password?: string };
  try {
    body = (await request.json()) as { password?: string };
  } catch {
    return NextResponse.json({ ok: false, error: "json" }, { status: 400 });
  }

  if (!verifyAdminPasswordAttempt(body.password ?? "")) {
    return NextResponse.json(
      { ok: false, error: "credenciales", mensaje: "Contraseña incorrecta." },
      { status: 401 }
    );
  }

  const token = createSignedSessionValue();
  if (!token) {
    return NextResponse.json({ ok: false, error: "session" }, { status: 500 });
  }

  const res = NextResponse.json({ ok: true });
  res.cookies.set(ADMIN_SESSION_COOKIE, token, sessionCookieOptions());
  return res;
}
