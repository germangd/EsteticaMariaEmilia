/**
 * Genera `src/lib/landing-media-manifest.json` listando archivos en
 * `public/landing/hero` y `public/landing/servicios` para que la home pueda
 * usar esos nombres cuando `fs.readdir` no ve la carpeta (p. ej. algunos runtimes).
 *
 * Se ejecuta en `prebuild` y al arrancar `npm run dev` (vía scripts/dev.cjs).
 */
"use strict";

const fs = require("fs");
const path = require("path");

const IMAGE_RE = /\.(jpe?g|png|webp|gif)$/i;
const VIDEO_RE = /\.(mp4|webm)$/i;

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
  const servicios = listFiles(servDir).filter((n) => IMAGE_RE.test(n));

  const outPath = path.join(
    __dirname,
    "..",
    "src",
    "lib",
    "landing-media-manifest.json"
  );
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(
    outPath,
    `${JSON.stringify({ hero, servicios }, null, 2)}\n`,
    "utf8"
  );
  console.log(
    `[landing-sync-manifest] hero: ${hero.length}, servicios: ${servicios.length} -> ${path.relative(cwd, outPath)}`
  );
}

main();
