/**
 * Genera `src/lib/landing-media-manifest.json` listando archivos en
 * `public/landing/hero` y `public/landing/servicios/<carpeta>/` (+ opcional
 * imágenes sueltas en la raíz de `servicios/` como respaldo legacy).
 *
 * Mantener alineado con `src/lib/servicio-media-folders.ts`.
 */
"use strict";

const fs = require("fs");
const path = require("path");

const IMAGE_RE = /\.(jpe?g|png|webp|gif)$/i;
const VIDEO_RE = /\.(mp4|webm)$/i;

/** @type {readonly string[]} */
const SERVICIO_FOLDERS = [
  "depilacion-laser",
  "faciales",
  "unas-esculpidas",
  "podologia",
  "coloracion",
  "alisado",
];

function landingRootCandidates(cwd) {
  return [
    path.join(cwd, "public", "landing"),
    path.join(cwd, "estetica-web", "public", "landing"),
  ];
}

function resolveLandingRoot(cwd) {
  for (const root of landingRootCandidates(cwd)) {
    if (fs.existsSync(path.join(root, "hero"))) return root;
  }
  return path.join(cwd, "public", "landing");
}

function listFiles(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter((name) => {
    if (name.startsWith(".") || name === "README.md") return false;
    const abs = path.join(dir, name);
    try {
      return fs.statSync(abs).isFile();
    } catch {
      return false;
    }
  });
}

function main() {
  const cwd = process.cwd();
  const root = resolveLandingRoot(cwd);
  const heroDir = path.join(root, "hero");
  const servDir = path.join(root, "servicios");

  const hero = listFiles(heroDir).filter(
    (n) => IMAGE_RE.test(n) || VIDEO_RE.test(n)
  );

  /** @type {Record<string, string[]>} */
  const servicios = {};
  for (const folder of SERVICIO_FOLDERS) {
    const sub = path.join(servDir, folder);
    servicios[folder] = listFiles(sub).filter((n) => IMAGE_RE.test(n));
  }

  const serviciosLegacyRoot = listFiles(servDir).filter((n) => {
    if (!IMAGE_RE.test(n)) return false;
    const abs = path.join(servDir, n);
    try {
      return fs.statSync(abs).isFile();
    } catch {
      return false;
    }
  });

  const outPath = path.join(
    __dirname,
    "..",
    "src",
    "lib",
    "landing-media-manifest.json"
  );
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  const payload = { hero, servicios, serviciosLegacyRoot };
  fs.writeFileSync(outPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");

  const subTotal = SERVICIO_FOLDERS.reduce(
    (acc, f) => acc + (servicios[f]?.length ?? 0),
    0
  );
  console.log(
    `[landing-sync-manifest] hero: ${hero.length}, servicios (subcarpetas): ${subTotal}, legacy raíz: ${serviciosLegacyRoot.length} -> ${path.relative(cwd, outPath)}`
  );
}

main();
