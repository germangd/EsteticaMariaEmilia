export default function Home() {
  return (
    <main className="min-h-screen bg-[#FAF7F4] text-[#2C2420] flex flex-col items-center justify-center px-6 py-16">
      <p className="text-xs tracking-[0.35em] uppercase text-[#C9A84C] mb-4">
        Estética profesional
      </p>
      <h1 className="font-serif text-4xl sm:text-5xl font-light text-center leading-tight mb-3">
        María <em className="text-[#C9A84C] not-italic">Emilia</em> Estética
      </h1>
      <p className="text-[#8A7A74] text-center max-w-md mb-10 font-light">
        Esqueleto Next.js para migrar la landing y la reserva a un solo sitio en
        Vercel. La lógica de turnos hay que portarla desde{" "}
        <code className="text-sm bg-white/80 px-1 rounded">Código.gs</code>.
      </p>
      <div className="flex flex-col sm:flex-row gap-3">
        <a
          href="/api/health"
          className="inline-flex items-center justify-center bg-[#C9A84C] text-white text-xs font-medium tracking-widest uppercase px-8 py-3 rounded-sm hover:bg-[#A07830] transition-colors"
        >
          Probar API /api/health
        </a>
        <a
          href="https://github.com/germangd/EsteticaMariaEmilia/blob/main/docs/migracion-stack-propio.md"
          className="inline-flex items-center justify-center border border-[#C9A84C]/40 text-[#4A3F3A] text-xs font-medium tracking-widest uppercase px-8 py-3 rounded-sm hover:bg-white/60 transition-colors"
          target="_blank"
          rel="noopener noreferrer"
        >
          Guía de migración
        </a>
      </div>
    </main>
  );
}
