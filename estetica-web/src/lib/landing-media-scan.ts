import { existsSync } from "fs";
import { readdir, stat } from "fs/promises";
import path from "path";
import { HERO_SLIDES, type HeroSlide } from "@/lib/landing-media";
import landingManifest from "./landing-media-manifest.json";
import type { ServicioMediaFolder } from "./servicio-media-folders";

const MAX_HERO_SLIDES = 16;

const IMAGE_RE = /\.(jpe?g|png|webp|gif)$/i;
const VIDEO_RE = /\.(mp4|webm)$/i;

function shuffle<T>(items: T[]): T[] {
  const a = [...items];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j]!, a[i]!];
  }
  return a;
}

function publicUrl(sub: "hero" | "servicios", file: string): string {
  return `/landing/${sub}/${encodeURIComponent(file)}`;
}

function servicioSubdirPublicUrl(folder: string, file: string): string {
  return `/landing/servicios/${encodeURIComponent(folder)}/${encodeURIComponent(file)}`;
}

/** Raíz del proyecto Next o monorepo (cuando `cwd` es la carpeta padre). */
function landingRootCandidates(): string[] {
  const cwd = process.cwd();
  return [
    path.join(cwd, "public", "landing"),
    path.join(cwd, "estetica-web", "public", "landing"),
  ];
}

function resolveLandingRoot(): string {
  for (const root of landingRootCandidates()) {
    if (existsSync(path.join(root, "hero"))) return root;
  }
  return path.join(process.cwd(), "public", "landing");
}

function landingAbs(...parts: string[]): string {
  return path.join(resolveLandingRoot(), ...parts);
}

function manifestData(): {
  hero: string[];
  serviciosPorCarpeta: Partial<Record<ServicioMediaFolder, string[]>>;
} {
  const m = landingManifest as {
    hero?: unknown;
    servicios?: unknown;
  };
  const hero = Array.isArray(m.hero) ? (m.hero as string[]) : [];

  let serviciosPorCarpeta: Partial<Record<ServicioMediaFolder, string[]>> =
    {};
  const rawServ = m.servicios;
  if (rawServ && typeof rawServ === "object" && !Array.isArray(rawServ)) {
    serviciosPorCarpeta = rawServ as Partial<
      Record<ServicioMediaFolder, string[]>
    >;
  }

  return { hero, serviciosPorCarpeta };
}

async function listBasenames(dir: string): Promise<string[]> {
  if (!existsSync(dir)) return [];
  const names = await readdir(dir);
  const files: string[] = [];
  for (const name of names) {
    if (name.startsWith(".") || name === "README.md") continue;
    const abs = path.join(dir, name);
    try {
      const st = await stat(abs);
      if (st.isFile()) files.push(name);
    } catch {
      /* skip */
    }
  }
  return files;
}

function posterForVideo(heroDir: string, videoBasename: string): string {
  const stem = videoBasename.replace(/\.[^.]+$/, "");
  for (const ext of [".jpg", ".jpeg", ".png", ".webp"]) {
    const posterName = stem + ext;
    if (existsSync(path.join(heroDir, posterName))) {
      return publicUrl("hero", posterName);
    }
  }
  return "";
}

function buildHeroSlides(heroDir: string, names: string[]): HeroSlide[] {
  const slides: HeroSlide[] = [];
  for (const name of names) {
    if (IMAGE_RE.test(name)) {
      slides.push({
        kind: "image",
        src: publicUrl("hero", name),
        alt: `María Emilia Estética — ${name.replace(/\.[^.]+$/, "")}`,
      });
    } else if (VIDEO_RE.test(name)) {
      slides.push({
        kind: "video",
        src: publicUrl("hero", name),
        poster: posterForVideo(heroDir, name),
        alt: `Video — ${name.replace(/\.[^.]+$/, "")}`,
      });
    }
  }
  return slides;
}

/**
 * Lee `public/landing/hero`, arma slides (imagen / video + poster opcional),
 * mezcla al azar y limita cantidad. Si no hay archivos, usa el manifiesto
 * generado en build/dev; si tampoco hay, devuelve `HERO_SLIDES`.
 */
export async function loadHeroSlidesFromPublic(): Promise<HeroSlide[]> {
  const heroDir = landingAbs("hero");
  const fromDisk = buildHeroSlides(heroDir, await listBasenames(heroDir));
  if (fromDisk.length > 0) {
    return shuffle(fromDisk).slice(0, MAX_HERO_SLIDES);
  }

  const mf = manifestData();
  const fromManifest = buildHeroSlides(heroDir, mf.hero);
  if (fromManifest.length > 0) {
    return shuffle(fromManifest).slice(0, MAX_HERO_SLIDES);
  }

  return HERO_SLIDES;
}

const DEFAULT_CARD_IMAGES: string[] = [
  "https://images.unsplash.com/photo-1519823551278-64ac92734fb1?w=900&h=675&fit=crop&q=80",
  "https://images.unsplash.com/photo-1570172619644-dfd03ed5d881?w=900&h=675&fit=crop&q=80",
  "https://images.unsplash.com/photo-1522337360788-8b13dee7a37e?w=900&h=675&fit=crop&q=80",
  "https://images.unsplash.com/photo-1544161515-4ab6ce6db874?w=900&h=675&fit=crop&q=80",
  "https://images.unsplash.com/photo-1560869713-da86a43ec442?w=900&h=675&fit=crop&q=80",
  "https://images.unsplash.com/photo-1522338242992-e2a54887f5f0?w=900&h=675&fit=crop&q=80",
];

async function listImagesInServicioFolder(
  folder: ServicioMediaFolder
): Promise<string[]> {
  const dir = landingAbs("servicios", folder);
  return (await listBasenames(dir)).filter((n) => IMAGE_RE.test(n));
}

function pickRandom<T>(items: T[]): T | undefined {
  if (items.length === 0) return undefined;
  return items[Math.floor(Math.random() * items.length)]!;
}

/**
 * Una URL por servicio: imagen aleatoria dentro de
 * `public/landing/servicios/<carpeta>/` (o el manifiesto si el disco no lista
 * archivos). Si esa carpeta está vacía, **no** se usan fotos de otras carpetas
 * ni de la raíz de `servicios/`: solo la imagen por defecto (Unsplash) de ese
 * índice.
 */
export async function pickServicioCardImagesByFolder(
  folders: readonly ServicioMediaFolder[]
): Promise<string[]> {
  const mf = manifestData();
  const out: string[] = [];
  for (let i = 0; i < folders.length; i++) {
    const folder = folders[i]!;
    let pool = shuffle(await listImagesInServicioFolder(folder));
    if (pool.length === 0) {
      const fromMf = mf.serviciosPorCarpeta[folder];
      pool = shuffle(
        Array.isArray(fromMf)
          ? fromMf.filter((n) => typeof n === "string" && IMAGE_RE.test(n))
          : []
      );
    }
    const file = pickRandom(pool);
    if (file) {
      out.push(servicioSubdirPublicUrl(folder, file));
    } else {
      out.push(DEFAULT_CARD_IMAGES[i % DEFAULT_CARD_IMAGES.length]!);
    }
  }
  return out;
}
