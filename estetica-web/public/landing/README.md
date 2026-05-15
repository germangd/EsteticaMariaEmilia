# Medios para la landing

Colocá archivos en estas carpetas (se listan en **cada request** de la home y se elige un orden **aleatorio**):

- **`hero/`** — fotos y videos del carrusel principal.  
  - Imágenes: `.jpg`, `.jpeg`, `.png`, `.webp`, `.gif`  
  - Video: `.mp4`, `.webm` (opcional: mismo nombre + `.jpg`/`.png` como *poster*, p. ej. `promo.mp4` + `promo.jpg`)

- **`servicios/`** — solo imágenes para las tarjetas de servicios (se asignan al azar a cada card).

Si una carpeta está **vacía**, la app usa los **valores por defecto** del código (Unsplash / ejemplo).

## Si ves siempre las fotos de ejemplo

1. Los archivos tienen que estar en **`estetica-web/public/landing/hero/`** y **`estetica-web/public/landing/servicios/`** (no en la raíz del repo ni solo en `index.html` estático).
2. Abrí la app con **Next.js** (`npm run dev` desde la raíz del repo o desde `estetica-web`). Al arrancar se actualiza `src/lib/landing-media-manifest.json`; si agregás fotos con el dev server ya abierto, **reiniciá** `npm run dev` una vez (o ejecutá `node scripts/landing-sync-manifest.cjs` y reiniciá).
3. En **Vercel**, las imágenes tienen que estar **commiteadas** y redeploy tras cambiar la carpeta `public/landing`.

No subas archivos enormes sin optimizar: afectan el peso del deploy y la carga.
