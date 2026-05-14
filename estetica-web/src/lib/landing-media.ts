/**
 * URLs de ejemplo (Unsplash + un video CC0). Reemplazá por fotos/videos propios
 * en `public/landing/…` o por tu CDN cuando los tengas.
 */

export type HeroSlide =
  | { kind: "image"; src: string; alt: string }
  | { kind: "video"; src: string; poster: string; alt: string };

export const HERO_SLIDES: HeroSlide[] = [
  {
    kind: "image",
    src: "https://images.unsplash.com/photo-1516975080664-ed2fc6a32937?w=1400&h=1050&fit=crop&q=80",
    alt: "Espacio de estética luminoso y acogedor",
  },
  {
    kind: "image",
    src: "https://images.unsplash.com/photo-1570172619644-dfd03ed5d881?w=1400&h=1050&fit=crop&q=80",
    alt: "Tratamiento facial y cuidado de la piel",
  },
  {
    kind: "image",
    src: "https://images.unsplash.com/photo-1522337360788-8b13dee7a37e?w=1400&h=1050&fit=crop&q=80",
    alt: "Manicura y detalle de uñas",
  },
  {
    kind: "video",
    src: "https://interactive-examples.mdn.mozilla.net/media/cc0-videos/flower.mp4",
    poster:
      "https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=1400&h=1050&fit=crop&q=80",
    alt: "Video de ejemplo — sustituí por reel propio del salón",
  },
];
