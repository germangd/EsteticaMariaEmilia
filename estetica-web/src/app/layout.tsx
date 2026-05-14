import type { Metadata, Viewport } from "next";
import { Cormorant_Garamond, Raleway } from "next/font/google";
import "./globals.css";

const raleway = Raleway({
  subsets: ["latin"],
  variable: "--font-raleway",
  weight: ["300", "400", "500"],
});

const cormorant = Cormorant_Garamond({
  subsets: ["latin"],
  variable: "--font-cormorant",
  weight: ["300", "400", "600"],
  style: ["normal", "italic"],
});

export const viewport: Viewport = {
  themeColor: "#FAF7F4",
};

export const metadata: Metadata = {
  title: "María Emilia Estética — Belleza y bienestar",
  description:
    "Depilación láser, faciales, uñas, podología y más. Reservá turno online. Ensenada, Bartolomé Bavio y Magdalena.",
  openGraph: {
    type: "website",
    locale: "es_AR",
    title: "María Emilia Estética — Belleza y bienestar",
    description:
      "Depilación láser, faciales, uñas, podología y más. Reservá turno online. Ensenada, Bartolomé Bavio y Magdalena.",
    images: [
      {
        url: "https://images.unsplash.com/photo-1560066984-138dadb4c035?w=1200&h=630&fit=crop&q=80",
        width: 1200,
        height: 630,
        alt: "María Emilia Estética",
      },
    ],
  },
  twitter: {
    card: "summary_large_image",
    title: "María Emilia Estética",
    description:
      "Reservá turno online. Depilación láser, faciales, uñas y más en tu zona.",
    images: [
      "https://images.unsplash.com/photo-1560066984-138dadb4c035?w=1200&h=630&fit=crop&q=80",
    ],
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="es" className="scroll-smooth">
      <body
        className={`${raleway.variable} ${cormorant.variable} font-sans antialiased bg-cream text-ink`}
      >
        {children}
      </body>
    </html>
  );
}
