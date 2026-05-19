import type { Metadata } from "next";
import { siteConfig } from "@/config/site.config";

/** Nombre del negocio (env opcional para despliegues sin tocar el archivo). */
export function getBusinessName(): string {
  return (
    process.env.NEXT_PUBLIC_SITE_NAME?.trim() ||
    process.env.SITE_NAME?.trim() ||
    siteConfig.businessName
  );
}

export function getBusinessNameShort(): string {
  return siteConfig.businessNameShort;
}

export function getLocationsLine(): string {
  return siteConfig.locationsLine;
}

export function adminPageTitle(section: string): string {
  return `Admin — ${section} | ${getBusinessName()}`;
}

export function reservarPageTitle(): string {
  return `Reservar turno | ${getBusinessName()}`;
}

export function buildSiteMetadata(): Metadata {
  const name = getBusinessName();
  const title = `${name} — ${siteConfig.tagline}`;
  const description = siteConfig.description;

  return {
    title,
    description,
    openGraph: {
      type: "website",
      locale: "es_AR",
      title,
      description,
      images: [
        {
          url: siteConfig.metadata.ogImage,
          width: 1200,
          height: 630,
          alt: name,
        },
      ],
    },
    twitter: {
      card: "summary_large_image",
      title: name,
      description: siteConfig.description.split(".")[0] + ".",
      images: [siteConfig.metadata.ogImage],
    },
  };
}

export function getWhatsAppNumberDigits(): string {
  const raw =
    process.env.NEXT_PUBLIC_WHATSAPP_NUMBER?.trim() ||
    process.env.WHATSAPP_NUMBER?.trim();
  const digits = raw?.replace(/\D/g, "");
  return digits || siteConfig.contact.whatsappNumber;
}

export function buildWhatsAppUrl(text: string): string {
  return `https://wa.me/${getWhatsAppNumberDigits()}?text=${encodeURIComponent(text)}`;
}

export function buildWhatsAppConsultaUrl(): string {
  return buildWhatsAppUrl(siteConfig.contact.whatsappConsultaMessage);
}

export function getMapsUrl(): string {
  return siteConfig.contact.mapsUrl;
}

export function getInstagramUrl(): string {
  return siteConfig.contact.instagramUrl;
}

export function getWhatsAppDisplay(): string {
  return siteConfig.contact.whatsappDisplay;
}

export function getMailFooterLine(): string {
  return siteConfig.mail.footerLine.replace(
    siteConfig.businessName,
    getBusinessName()
  );
}

export function getMailAccentColor(): string {
  return siteConfig.mail.accentColor;
}

export { siteConfig };
