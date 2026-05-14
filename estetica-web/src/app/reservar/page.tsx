import type { Metadata } from "next";
import Link from "next/link";

export const metadata: Metadata = {
  title: "Reservar turno | María Emilia Estética",
  description:
    "Reservá tu turno online en María Emilia Estética — Ensenada, Bartolomé Bavio y Magdalena.",
};

const WA_URL =
  "https://wa.me/5492215918286?text=Hola!%20Quiero%20reservar%20turno%20en%20Mar%C3%ADa%20Emilia%20Est%C3%A9tica";

export default function ReservarPage() {
  return (
    <main className="flex min-h-screen flex-col items-center justify-center bg-cream px-6 py-24 text-center">
      <p className="mb-3 text-xs font-medium uppercase tracking-[0.3em] text-gold">
        Turnos
      </p>
      <h1 className="mb-4 max-w-lg font-serif text-3xl font-light text-ink-dark md:text-4xl">
        La reserva online en este sitio se está armando
      </h1>
      <p className="mb-10 max-w-md text-sm font-light leading-relaxed text-ink-muted">
        Pronto vas a poder elegir servicio, fecha y horario acá mismo. Mientras
        tanto, escribinos por WhatsApp y coordinamos tu visita.
      </p>
      <div className="flex flex-col gap-3 sm:flex-row sm:gap-4">
        <a
          href={WA_URL}
          target="_blank"
          rel="noopener noreferrer"
          className="inline-flex items-center justify-center gap-2 rounded-[2px] bg-[#25D366] px-8 py-3.5 text-xs font-medium uppercase tracking-[0.12em] text-white transition-colors hover:bg-[#1da851]"
        >
          WhatsApp
        </a>
        <Link
          href="/"
          className="inline-flex items-center justify-center rounded-[2px] border border-gold/40 px-8 py-3.5 text-xs font-medium uppercase tracking-[0.12em] text-ink transition-colors hover:bg-white/80"
        >
          Volver al inicio
        </Link>
      </div>
    </main>
  );
}
