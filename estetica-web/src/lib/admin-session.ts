import { createHash, createHmac, timingSafeEqual } from "node:crypto";

/** Cookie httpOnly con sesión firmada (HMAC). */
export const ADMIN_SESSION_COOKIE = "me_admin_sess";

const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000;

/** Secreto para firmar la cookie (no es la contraseña que escribe la dueña). */
export function getAdminSigningSecret(): string | null {
  const s =
    process.env.ADMIN_SESSION_SECRET?.trim() ||
    process.env.ADMIN_SECRET?.trim();
  return s || null;
}

/** Contraseña de acceso al panel (comparación por SHA-256 + timingSafeEqual). */
export function getAdminPasswordPlain(): string | null {
  return process.env.ADMIN_PASSWORD?.trim() || null;
}

export function adminAuthConfigured(): boolean {
  return Boolean(getAdminSigningSecret() && getAdminPasswordPlain());
}

function sha256buf(s: string): Buffer {
  return createHash("sha256").update(s, "utf8").digest();
}

export function verifyAdminPasswordAttempt(pw: string): boolean {
  const envp = getAdminPasswordPlain();
  if (!envp || !pw) return false;
  const a = sha256buf(pw);
  const b = sha256buf(envp);
  return a.length === b.length && timingSafeEqual(a, b);
}

export function createSignedSessionValue(): string | null {
  const secret = getAdminSigningSecret();
  if (!secret) return null;
  const exp = Date.now() + SESSION_TTL_MS;
  const payload = Buffer.from(JSON.stringify({ exp }), "utf8").toString(
    "base64url"
  );
  const sig = createHmac("sha256", secret).update(payload).digest("base64url");
  return `${payload}.${sig}`;
}

export function verifySignedSessionValue(token: string | undefined): boolean {
  if (!token?.includes(".")) return false;
  const secret = getAdminSigningSecret();
  if (!secret) return false;
  const i = token.lastIndexOf(".");
  const payload = token.slice(0, i);
  const sig = token.slice(i + 1);
  if (!payload || !sig) return false;
  const expected = createHmac("sha256", secret).update(payload).digest("base64url");
  const a = Buffer.from(sig, "utf8");
  const b = Buffer.from(expected, "utf8");
  if (a.length !== b.length) return false;
  if (!timingSafeEqual(a, b)) return false;
  try {
    const json = JSON.parse(
      Buffer.from(payload, "base64url").toString("utf8")
    ) as { exp?: number };
    if (typeof json.exp !== "number" || json.exp < Date.now()) return false;
    return true;
  } catch {
    return false;
  }
}

export function sessionCookieOptions() {
  return {
    httpOnly: true as const,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax" as const,
    path: "/",
    maxAge: Math.floor(SESSION_TTL_MS / 1000),
  };
}
