import { drizzle } from "drizzle-orm/neon-http";
import { getNeonSql } from "@/lib/db";
import * as schema from "./schema";

export function getDb() {
  const sql = getNeonSql();
  if (!sql) return null;
  return drizzle({ client: sql, schema });
}
