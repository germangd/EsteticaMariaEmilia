import type { Viewport } from "next";
import { Cormorant_Garamond, Raleway } from "next/font/google";
import { buildSiteMetadata } from "@/config/site";
import { siteConfig } from "@/config/site.config";
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
  themeColor: siteConfig.themeColor,
};

export const metadata = buildSiteMetadata();

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
