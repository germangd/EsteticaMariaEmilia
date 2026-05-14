import { config as loadEnv } from "dotenv";
import { defineConfig } from "drizzle-kit";
import { dirname, resolve } from "path";
import { fileURLToPath } from "url";
import WebSocketNode from "ws";

/**
 * drizzle-kit + Neon usa WebSocket en Node; sin esto en Windows/Node 20 a veces
 * `db:push` se queda en "Pulling schema..." y termina sin mensaje.
 */
(globalThis as unknown as { WebSocket: typeof globalThis.WebSocket }).WebSocket =
  WebSocketNode as unknown as typeof globalThis.WebSocket;

/** Carpeta `estetica-web/` (donde vive este archivo), aunque `npm run` se ejecute desde la raíz del repo. */
const esteticaWebRoot = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(esteticaWebRoot, "..");

// Raíz del repo primero, luego esta carpeta con override (gana `estetica-web/.env.local`).
loadEnv({ path: resolve(repoRoot, ".env") });
loadEnv({ path: resolve(repoRoot, ".env.local") });
loadEnv({ path: resolve(esteticaWebRoot, ".env"), override: true });
loadEnv({ path: resolve(esteticaWebRoot, ".env.local"), override: true });

/**
 * Drizzle Kit (push / migrate / studio): Neon recomienda URL **directa** (sin pooler);
 * la pooled puede fallar o usar WebSocket de forma rara en Windows.
 * La app Next sigue usando solo `DATABASE_URL` (pooled) en `src/lib/db.ts`.
 */
const pooledUrl = process.env.DATABASE_URL?.trim() ?? "";
const directUrl = process.env.DATABASE_URL_DIRECT?.trim() ?? "";
const drizzleDbUrl = directUrl || pooledUrl;

if (pooledUrl.includes("-pooler.") && !directUrl) {
  console.warn(
    "[drizzle-kit] Tu DATABASE_URL usa el host pooled (-pooler). " +
      "Si push/migrate cuelga o falla, añadí DATABASE_URL_DIRECT con la conexión **directa** del panel de Neon (sin pooler)."
  );
}

export default defineConfig({
  schema: "./src/db/schema.ts",
  out: "./drizzle",
  dialect: "postgresql",
  dbCredentials: {
    url: drizzleDbUrl,
  },
});
