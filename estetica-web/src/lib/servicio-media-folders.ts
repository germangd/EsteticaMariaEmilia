/**
 * Subcarpetas de `public/landing/servicios/<carpeta>/` — una por tarjeta de servicio
 * (mismo orden que `SERVICIOS_BASE` en la home).
 */
export const SERVICIO_MEDIA_FOLDERS = [
  "depilacion-laser",
  "faciales",
  "unas-esculpidas",
  "podologia",
  "coloracion",
  "alisado",
] as const;

export type ServicioMediaFolder = (typeof SERVICIO_MEDIA_FOLDERS)[number];
