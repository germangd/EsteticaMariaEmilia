/**
 * Ejemplo mínimo para otro cliente.
 * Copiá la estructura a `src/config/site.config.ts` y completá con tus datos.
 *
 * Ver docs/NUEVO-CLIENTE.md
 */
import {
  SERVICIO_MEDIA_FOLDERS,
  type ServicioMediaFolder,
} from "@/lib/servicio-media-folders";

export const siteConfig = {
  businessName: "Nombre del Salón",
  businessNameShort: "Salón",
  tagline: "Tu eslogan",
  description: "Descripción para Google y redes.",
  locationsLine: "Ciudad · Barrio",

  themeColor: "#FAF7F4",

  metadata: {
    ogImage: "https://ejemplo.com/og.jpg",
  },

  contact: {
    whatsappNumber: "5491100000000",
    whatsappDisplay: "+54 9 11 0000-0000",
    whatsappConsultaMessage: "Hola! Quiero consultar sobre sus servicios",
    mapsUrl: "https://www.google.com/maps/search/?api=1&query=...",
    instagramUrl: "https://instagram.com/tu_cuenta",
    instagramHandle: "@tu_cuenta",
  },

  mail: {
    accentColor: "#A07830",
    footerLine: "© Nombre del Salón · Ciudad",
    emailFromExample: "Nombre del Salón <onboarding@resend.dev>",
  },

  landing: {
    navBrand: "Salón",
    hero: {
      kicker: "Estética profesional",
      titleLines: ["Nombre", "del", "Salón"] as const,
      subtitle: "Frase corta del hero",
      locationsLine: "Ciudad · Barrio",
    },
    about: {
      paragraph: "Texto quienes somos...",
      cardTitle: "Nombre",
      cardKicker: "Profesional",
      stats: [
        { value: "5+", label: "Servicios" },
        { value: "1", label: "Sede" },
        { value: "100%", label: "Profesional" },
      ] as const,
    },
    zones: [{ name: "Ciudad", region: "Provincia" }],
    footer: {
      blurb: "Texto pie de página.",
      serviceList: ["Servicio 1", "Servicio 2"],
    },
    servicios: [
      {
        n: "01",
        icon: "✨",
        name: "Servicio 1",
        desc: "Descripción breve.",
        mediaFolder: SERVICIO_MEDIA_FOLDERS[0] as ServicioMediaFolder,
      },
    ],
  },
} as const;
