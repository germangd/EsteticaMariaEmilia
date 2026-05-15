# Medios para la landing

Colocá archivos en estas carpetas (la home las lee en cada request; el orden del carrusel es **aleatorio**):

- **`hero/`** — fotos y videos del carrusel principal.  
  - Imágenes: `.jpg`, `.jpeg`, `.png`, `.webp`, `.gif`  
  - Video: `.mp4`, `.webm` (opcional: mismo nombre + `.jpg`/`.png` como *poster*, p. ej. `promo.mp4` + `promo.jpg`)

- **`servicios/`** — **una subcarpeta por servicio** (solo imágenes). Cada tarjeta toma **al azar** una foto de **su** carpeta; así no se mezclan con otros servicios.

  | Carpeta | Tarjeta en la web |
  |---------|-------------------|
  | `servicios/depilacion-laser/` | Depilación Láser |
  | `servicios/faciales/` | Faciales |
  | `servicios/unas-esculpidas/` | Uñas & Esculpidas |
  | `servicios/podologia/` | Podología |
  | `servicios/coloracion/` | Coloración |
  | `servicios/alisado/` | Alisado |

  Los nombres de carpeta tienen que coincidir **exactamente** con la tabla (kebab-case). Los nombres de archivo pueden ser los que quieras (`.jpg`, `.jpeg`, `.png`, `.webp`, `.gif`).

Si una subcarpeta está **vacía**, esa tarjeta muestra la **foto de ejemplo** (Unsplash) de ese índice; no se reutilizan fotos de otras carpetas ni imágenes sueltas en la raíz de `servicios/` (evita mezclar servicios).

Las imágenes en la **raíz** de `servicios/` no se usan para las tarjetas: colocá todo dentro de la subcarpeta que corresponda.

Al arrancar `npm run dev` o `npm run build` se actualiza `src/lib/landing-media-manifest.json` (útil en algunos deploys).

No subas archivos enormes sin optimizar: afectan el peso del deploy y la carga.
