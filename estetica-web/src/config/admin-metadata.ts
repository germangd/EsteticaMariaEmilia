import type { Metadata } from "next";
import { adminPageTitle } from "@/config/site";

/** Título de pestaña para páginas del panel admin. */
export function adminMetadata(section: string): Metadata {
  return {
    title: adminPageTitle(section),
    robots: { index: false, follow: false },
  };
}
