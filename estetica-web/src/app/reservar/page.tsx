import type { Metadata } from "next";
import { ReservarClient } from "@/components/reservar/reservar-client";

export const metadata: Metadata = {
  title: "Reservar turno | María Emilia Estética",
  description:
    "Reservá o cancelá tu turno online — Ensenada, Bartolomé Bavio y Magdalena.",
};

export default function ReservarPage() {
  return <ReservarClient />;
}
