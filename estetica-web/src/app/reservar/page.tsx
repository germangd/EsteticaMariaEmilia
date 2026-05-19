import type { Metadata } from "next";
import { ReservarClient } from "@/components/reservar/reservar-client";
import { reservarPageTitle, siteConfig } from "@/config/site";

export const metadata: Metadata = {
  title: reservarPageTitle(),
  description: siteConfig.description,
};

export default function ReservarPage() {
  return <ReservarClient />;
}
