import { neon } from "@neondatabase/serverless";

let cached: ReturnType<typeof neon> | null = null;

/**
 * Cliente SQL para Neon (HTTP). En Vercel usá la URL **pooled** del panel.
 * Sin `DATABASE_URL`, devuelve `null` (build y dev sin DB siguen funcionando).
 */
export function getNeonSql(): ReturnType<typeof neon> | null {
  const url = process.env.DATABASE_URL?.trim();
  if (!url) return null;
  if (!cached) cached = neon(url);
  return cached;
}
