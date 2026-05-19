/**
 * Configuración de marca y landing por implementación.
 * Para un cliente nuevo: copiá el proyecto, editá este archivo y reemplazá `public/landing/`.
 *
 * Opcional en Vercel (.env): NEXT_PUBLIC_SITE_NAME, NEXT_PUBLIC_WHATSAPP_NUMBER
 */
import {
  SERVICIO_MEDIA_FOLDERS,
  type ServicioMediaFolder,
} from "@/lib/servicio-media-folders";

export type LandingServicioCard = {
  n: string;
  icon: string;
  name: string;
  desc: string;
  mediaFolder: ServicioMediaFolder;
};

export type ZonaAtencion = {
  name: string;
  region: string;
};

export const siteConfig = {
  /** Nombre legal / comercial completo */
  businessName: "María Emilia Estética",
  /** Marca corta (nav, footer) */
  businessNameShort: "ME Estética",
  /** Línea principal SEO */
  tagline: "Belleza y bienestar",
  /** Meta description (home, Open Graph) */
  description:
    "Depilación láser, faciales, uñas, podología y más. Reservá turno online. Ensenada, Bartolomé Bavio y Magdalena.",
  /** Texto bajo hero y pie de mails */
  locationsLine: "Ensenada · Bartolomé Bavio · Magdalena",

  themeColor: "#FAF7F4",

  metadata: {
    ogImage:
      "https://images.unsplash.com/photo-1560066984-138dadb4c035?w=1200&h=630&fit=crop&q=80",
  },

  contact: {
    /** Argentina, solo dígitos (sin +). Sobreescribible con NEXT_PUBLIC_WHATSAPP_NUMBER */
    whatsappNumber: "5492215918286",
    whatsappDisplay: "+54 9 221 591-8286",
    whatsappConsultaMessage:
      "Hola! Quiero consultar sobre los servicios de María Emilia Estética",
    mapsUrl:
      "https://www.google.com/maps/search/?api=1&query=Ensenada%2C+Provincia+de+Buenos+Aires%2C+Argentina",
    instagramUrl: "https://instagram.com/mariaemilia_estetica_",
    instagramHandle: "@mariaemilia_estetica_",
  },

  mail: {
    accentColor: "#A07830",
    /** Pie de correos al cliente */
    footerLine: "© María Emilia Estética · Ensenada · Bartolomé Bavio · Magdalena",
    /** Ejemplo en hint de EMAIL_FROM */
    emailFromExample: "María Emilia Estética <onboarding@resend.dev>",
  },

  landing: {
    navBrand: "ME Estética",
    hero: {
      kicker: "Estética profesional",
      titleLines: ["María", "Emilia", "Estética"] as const,
      subtitle: "Tu espacio de belleza y bienestar",
      locationsLine: "Ensenada · Bartolomé Bavio · Magdalena",
    },
    about: {
      paragraph:
        "En María Emilia Estética creemos que el cuidado personal es una forma de amor propio. Ofrecemos un ambiente cálido, profesional y personalizado donde cada cliente recibe la atención que merece. Trabajamos con los mejores productos y técnicas actualizadas para que salgas sintiéndote increíble.",
      cardTitle: "María Emilia",
      cardKicker: "Estética profesional",
      stats: [
        { value: "7+", label: "Servicios" },
        { value: "3", label: "Zonas" },
        { value: "100%", label: "Profesional" },
      ] as const,
    },
    zones: [
      { name: "Ensenada", region: "Provincia de Buenos Aires" },
      { name: "Bartolomé Bavio", region: "Provincia de Buenos Aires" },
      { name: "Magdalena", region: "Provincia de Buenos Aires" },
    ] satisfies ZonaAtencion[],
    footer: {
      blurb:
        "Tu espacio de belleza y bienestar profesional en la zona de Ensenada, Bartolomé Bavio y Magdalena.",
      serviceList: [
        "Depilación Láser",
        "Faciales",
        "Uñas & Esculpidas",
        "Podología",
        "Coloración",
        "Alisado",
      ],
    },
    /** Tarjetas de la home (orden = carpetas en public/landing/servicios/) */
    servicios: [
      {
        n: "01",
        icon: "✨",
        name: "Depilación Láser",
        desc: "Tecnología de última generación para una depilación definitiva, segura y sin dolor. Resultados duraderos desde la primera sesión.",
        mediaFolder: SERVICIO_MEDIA_FOLDERS[0],
      },
      {
        n: "02",
        icon: "🌸",
        name: "Faciales",
        desc: "Tratamientos personalizados para limpiar, hidratar y rejuvenecer tu piel. Protocolos adaptados a cada tipo de cutis.",
        mediaFolder: SERVICIO_MEDIA_FOLDERS[1],
      },
      {
        n: "03",
        icon: "💅",
        name: "Uñas & Esculpidas",
        desc: "Manicuría, esmaltado semipermanente y uñas esculpidas en acrílico o gel. Diseños únicos para cada ocasión.",
        mediaFolder: SERVICIO_MEDIA_FOLDERS[2],
      },
      {
        n: "04",
        icon: "🦶",
        name: "Podología",
        desc: "Cuidado profesional de pies para tu salud y bienestar. Tratamientos preventivos y estéticos a cargo de especialistas.",
        mediaFolder: SERVICIO_MEDIA_FOLDERS[3],
      },
      {
        n: "05",
        icon: "🎨",
        name: "Coloración",
        desc: "Tintura, mechas, balayage y técnicas de color actuales. Transformá tu look con los mejores productos del mercado.",
        mediaFolder: SERVICIO_MEDIA_FOLDERS[4],
      },
      {
        n: "06",
        icon: "💫",
        name: "Alisado",
        desc: "Alisado progresivo y keratinas para un cabello liso, brillante y sin frizz. Resultados que duran meses.",
        mediaFolder: SERVICIO_MEDIA_FOLDERS[5],
      },
    ] satisfies LandingServicioCard[],
  },
} as const;

export type SiteConfig = typeof siteConfig;
