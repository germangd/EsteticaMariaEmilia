import type { NextRequest } from "next/server";
import { NextResponse } from "next/server";
import {
  ADMIN_SESSION_COOKIE,
  verifySignedSessionValue,
} from "@/lib/admin-session";

export function isAdminRequest(request: NextRequest): boolean {
  return verifySignedSessionValue(
    request.cookies.get(ADMIN_SESSION_COOKIE)?.value
  );
}

export function adminUnauthorizedResponse(): NextResponse {
  return NextResponse.json({ ok: false, error: "no_autorizado" }, { status: 401 });
}
