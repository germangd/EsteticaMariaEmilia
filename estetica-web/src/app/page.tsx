import Image from "next/image";
import Link from "next/link";
import { HeroCarousel } from "@/components/landing/hero-carousel";
import { MeLogo } from "@/components/landing/me-logo";

const MAPS_URL =
  "https://www.google.com/maps/search/?api=1&query=Ensenada%2C+Provincia+de+Buenos+Aires%2C+Argentina";

const WA_URL =
  "https://wa.me/5492215918286?text=Hola!%20Quiero%20consultar%20sobre%20los%20servicios%20de%20Mar%C3%ADa%20Emilia%20Est%C3%A9tica";

const SERVICIOS = [
  {
    n: "01",
    icon: "✨",
    name: "Depilación Láser",
    desc: "Tecnología de última generación para una depilación definitiva, segura y sin dolor. Resultados duraderos desde la primera sesión.",
    image:
      "https://images.unsplash.com/photo-1519823551278-64ac92734fb1?w=900&h=675&fit=crop&q=80",
  },
  {
    n: "02",
    icon: "🌸",
    name: "Faciales",
    desc: "Tratamientos personalizados para limpiar, hidratar y rejuvenecer tu piel. Protocolos adaptados a cada tipo de cutis.",
    image:
      "https://images.unsplash.com/photo-1570172619644-dfd03ed5d881?w=900&h=675&fit=crop&q=80",
  },
  {
    n: "03",
    icon: "💅",
    name: "Uñas & Esculpidas",
    desc: "Manicuría, esmaltado semipermanente y uñas esculpidas en acrílico o gel. Diseños únicos para cada ocasión.",
    image:
      "https://images.unsplash.com/photo-1522337360788-8b13dee7a37e?w=900&h=675&fit=crop&q=80",
  },
  {
    n: "04",
    icon: "🦶",
    name: "Podología",
    desc: "Cuidado profesional de pies para tu salud y bienestar. Tratamientos preventivos y estéticos a cargo de especialistas.",
    image:
      "https://images.unsplash.com/photo-1544161515-4ab6ce6db874?w=900&h=675&fit=crop&q=80",
  },
  {
    n: "05",
    icon: "🎨",
    name: "Coloración",
    desc: "Tintura, mechas, balayage y técnicas de color actuales. Transformá tu look con los mejores productos del mercado.",
    image:
      "https://images.unsplash.com/photo-1560869713-da86a43ec442?w=900&h=675&fit=crop&q=80",
  },
  {
    n: "06",
    icon: "💫",
    name: "Alisado",
    desc: "Alisado progresivo y keratinas para un cabello liso, brillante y sin frizz. Resultados que duran meses.",
    image:
      "https://images.unsplash.com/photo-1522338242992-e2a54887f5f0?w=900&h=675&fit=crop&q=80",
  },
];

function WaIcon({ className }: { className?: string }) {
  return (
    <svg
      className={className}
      width="18"
      height="18"
      viewBox="0 0 24 24"
      fill="currentColor"
      aria-hidden
    >
      <path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z" />
    </svg>
  );
}

export default function Home() {
  return (
    <>
      <div
        className="pointer-events-none fixed inset-0 z-[1000] opacity-[0.12] motion-reduce:opacity-[0.06]"
        style={{
          backgroundImage: `url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.04'/%3E%3C/svg%3E")`,
        }}
        aria-hidden
      />

      <nav className="fixed left-0 right-0 top-0 z-[100] flex items-center justify-between border-b border-gold/20 bg-cream/85 px-5 py-4 backdrop-blur-md md:px-10">
        <Link
          href="/#inicio"
          className="cursor-pointer font-serif text-[1.05rem] font-semibold uppercase tracking-[0.15em] text-gold-dark"
        >
          ME Estética
        </Link>
        <ul className="hidden list-none items-center gap-9 md:flex">
          <li>
            <a
              href="#servicios"
              className="cursor-pointer text-[0.72rem] font-medium uppercase tracking-[0.18em] text-ink transition-colors hover:text-gold"
            >
              Servicios
            </a>
          </li>
          <li>
            <a
              href="#nosotros"
              className="cursor-pointer text-[0.72rem] font-medium uppercase tracking-[0.18em] text-ink transition-colors hover:text-gold"
            >
              Nosotros
            </a>
          </li>
          <li>
            <a
              href="#zonas"
              className="cursor-pointer text-[0.72rem] font-medium uppercase tracking-[0.18em] text-ink transition-colors hover:text-gold"
            >
              Zonas
            </a>
          </li>
          <li>
            <Link
              href="/reservar"
              className="cursor-pointer rounded-[2px] bg-gold px-[22px] py-2.5 text-[0.72rem] font-medium uppercase tracking-[0.15em] text-white transition-colors hover:bg-gold-dark"
            >
              Reservar turno
            </Link>
          </li>
        </ul>
        <Link
          href="/reservar"
          className="rounded-[2px] bg-gold px-4 py-2 text-[0.65rem] font-medium uppercase tracking-[0.12em] text-white md:hidden"
        >
          Reservar
        </Link>
      </nav>

      <main id="inicio">
        <section className="relative grid min-h-screen overflow-hidden md:grid-cols-2 md:items-stretch">
          <div className="relative z-[2] flex flex-col justify-center px-8 pb-16 pt-28 md:pl-20 md:pr-12 md:pt-32">
            <p className="mb-6 text-[0.68rem] font-medium uppercase tracking-[0.3em] text-gold motion-safe:animate-[fadeUp_0.8s_ease_both] motion-reduce:opacity-100">
              Estética profesional
            </p>
            <h1 className="mb-4 font-serif text-[clamp(2.5rem,5vw,5rem)] font-light leading-[1.05] text-ink-dark motion-safe:animate-[fadeUp_0.8s_ease_0.15s_both] motion-reduce:opacity-100">
              María
              <br />
              <em className="font-serif not-italic text-gold">Emilia</em>
              <br />
              Estética
            </h1>
            <p className="mb-10 max-w-md font-serif text-xl font-light italic text-ink-muted motion-safe:animate-[fadeUp_0.8s_ease_0.3s_both] motion-reduce:opacity-100">
              Tu espacio de belleza y bienestar
            </p>
            <div className="mb-10 flex flex-wrap gap-4 motion-safe:animate-[fadeUp_0.8s_ease_0.45s_both] motion-reduce:opacity-100">
              <Link
                href="/reservar"
                className="inline-block rounded-[2px] bg-gold px-9 py-4 text-[0.72rem] font-medium uppercase tracking-[0.2em] text-white transition-all hover:-translate-y-0.5 hover:bg-gold-dark"
              >
                Reservar turno
              </Link>
              <a
                href="#servicios"
                className="inline-flex items-center gap-2 border-b border-gold pb-4 text-[0.72rem] font-medium uppercase tracking-[0.2em] text-ink transition-colors hover:text-gold"
              >
                Ver servicios →
              </a>
            </div>
            <p className="text-[0.65rem] font-medium uppercase tracking-[0.25em] text-ink-muted motion-safe:animate-[fadeUp_0.8s_ease_0.6s_both] motion-reduce:opacity-100">
              Ensenada · Bartolomé Bavio · Magdalena
            </p>
          </div>

          <div className="relative min-h-[380px] md:h-full md:min-h-0">
            <div className="md:absolute md:inset-0 md:min-h-0">
              <HeroCarousel />
            </div>
            <div className="pointer-events-none absolute bottom-20 left-1/2 z-20 hidden -translate-x-1/2 rounded-full border border-white/30 bg-cream/80 p-3 shadow-lg backdrop-blur-sm md:block">
              <MeLogo gradientId="meHeroCarousel" className="h-10 w-[4.5rem]" />
            </div>
          </div>
        </section>

        <section id="servicios" className="bg-white px-6 py-20 md:px-20 md:py-24">
          <p className="mb-4 text-center text-[0.65rem] uppercase tracking-[0.35em] text-gold">
            Lo que hacemos
          </p>
          <h2 className="mb-14 text-center font-serif text-[clamp(1.75rem,3.5vw,3rem)] font-light leading-tight text-ink-dark">
            Nuestros <em className="font-serif not-italic text-gold">Servicios</em>
          </h2>
          <div className="mx-auto grid max-w-6xl grid-cols-1 gap-px bg-gold/20 sm:grid-cols-2 lg:grid-cols-3">
            {SERVICIOS.map((s) => (
              <article
                key={s.n}
                className="group relative cursor-default overflow-hidden bg-cream transition-transform duration-300 hover:-translate-y-1"
              >
                <div className="pointer-events-none absolute inset-0 z-[2] bg-gradient-to-br from-rose to-lilac opacity-0 transition-opacity duration-300 group-hover:opacity-100" />
                <div className="relative aspect-[4/3] w-full overflow-hidden">
                  <Image
                    src={s.image}
                    alt={s.name}
                    fill
                    sizes="(max-width: 640px) 100vw, (max-width: 1024px) 50vw, 33vw"
                    className="object-cover transition-transform duration-500 group-hover:scale-[1.04]"
                  />
                  <span
                    className="absolute left-3 top-3 text-2xl drop-shadow-md"
                    aria-hidden
                  >
                    {s.icon}
                  </span>
                  <span className="pointer-events-none absolute right-3 top-2 font-serif text-[3rem] font-light leading-none text-white/25 drop-shadow-sm">
                    {s.n}
                  </span>
                </div>
                <div className="relative z-[3] p-8 md:p-10">
                  <h3 className="mb-3 font-serif text-2xl font-normal text-ink-dark">
                    {s.name}
                  </h3>
                  <p className="text-[0.82rem] font-light leading-relaxed text-ink-muted">
                    {s.desc}
                  </p>
                </div>
              </article>
            ))}
          </div>
        </section>

        <section
          id="nosotros"
          className="grid items-center gap-12 bg-cream px-6 py-20 md:grid-cols-2 md:gap-20 md:px-20 md:py-24"
        >
          <div>
            <p className="mb-4 text-[0.65rem] uppercase tracking-[0.35em] text-gold">
              Quiénes somos
            </p>
            <div className="mb-5 h-px w-[60px] bg-gold" />
            <h2 className="mb-6 font-serif text-[clamp(1.75rem,3.5vw,3rem)] font-light leading-tight text-ink-dark">
              Un espacio creado
              <br />
              para <em className="font-serif not-italic text-gold">vos</em>
            </h2>
            <p className="mb-8 text-[0.95rem] font-light leading-[1.9] text-ink-muted">
              En María Emilia Estética creemos que el cuidado personal es una forma
              de amor propio. Ofrecemos un ambiente cálido, profesional y
              personalizado donde cada cliente recibe la atención que merece.
              Trabajamos con los mejores productos y técnicas actualizadas para
              que salgas sintiéndote increíble.
            </p>
            <div className="grid grid-cols-3 gap-6 border-t border-gold/20 py-8">
              <div>
                <div className="mb-1 font-serif text-4xl font-light text-gold">
                  7+
                </div>
                <div className="text-[0.68rem] uppercase tracking-[0.15em] text-ink-muted">
                  Servicios
                </div>
              </div>
              <div>
                <div className="mb-1 font-serif text-4xl font-light text-gold">
                  3
                </div>
                <div className="text-[0.68rem] uppercase tracking-[0.15em] text-ink-muted">
                  Zonas
                </div>
              </div>
              <div>
                <div className="mb-1 font-serif text-4xl font-light text-gold">
                  100%
                </div>
                <div className="text-[0.68rem] uppercase tracking-[0.15em] text-ink-muted">
                  Profesional
                </div>
              </div>
            </div>
          </div>
          <div className="relative">
            <div className="relative overflow-hidden rounded bg-gradient-to-br from-rose to-lilac p-12 text-center after:absolute after:-right-10 after:-top-10 after:h-[200px] after:w-[200px] after:rounded-full after:bg-gold/15">
              <div className="relative z-[1] mx-auto mb-6 flex h-40 w-40 items-center justify-center rounded-full bg-cream/90 shadow-lg">
                <MeLogo gradientId="meCardGrad" className="w-[110px]" />
              </div>
              <p className="relative z-[1] mb-2 font-serif text-2xl italic text-ink-dark">
                María Emilia
              </p>
              <p className="relative z-[1] text-[0.7rem] uppercase tracking-[0.2em] text-gold-dark">
                Estética profesional
              </p>
            </div>
          </div>
        </section>

        <section
          id="zonas"
          className="bg-ink-dark px-6 py-16 text-center md:px-20 md:py-20"
        >
          <p className="mb-4 text-[0.65rem] uppercase tracking-[0.35em] text-gold-light">
            Dónde encontrarnos
          </p>
          <h2 className="mb-10 font-serif text-[clamp(1.75rem,3.5vw,3rem)] font-light text-white">
            Atendemos en <em className="font-serif not-italic text-gold-light">tu zona</em>
          </h2>
          <div className="mx-auto flex max-w-3xl flex-wrap justify-center gap-10 md:gap-16">
            {[
              ["Ensenada", "Provincia de Buenos Aires"],
              ["Bartolomé Bavio", "Provincia de Buenos Aires"],
              ["Magdalena", "Provincia de Buenos Aires"],
            ].map(([name, sub]) => (
              <div key={name} className="text-center">
                <div className="mx-auto mb-4 h-2 w-2 rounded-full bg-gold" />
                <p className="mb-1 font-serif text-xl font-light text-white">
                  {name}
                </p>
                <p className="text-[0.65rem] uppercase tracking-[0.2em] text-white/40">
                  {sub}
                </p>
              </div>
            ))}
          </div>
          <p className="mt-10">
            <a
              href={MAPS_URL}
              target="_blank"
              rel="noopener noreferrer"
              className="inline-flex items-center gap-2 rounded-[2px] border border-gold-light/45 px-6 py-3 text-[0.68rem] font-medium uppercase tracking-[0.18em] text-gold-light transition-colors hover:bg-gold/15 hover:text-white"
            >
              Cómo llegar — mapa
            </a>
          </p>
        </section>

        <section
          id="turnos"
          className="relative overflow-hidden bg-gradient-to-br from-rose via-lilac to-[#edd9f5] px-6 py-20 text-center md:px-20 md:py-28"
        >
          <div className="pointer-events-none absolute -left-24 -top-24 h-[400px] w-[400px] rounded-full bg-gold/10" />
          <div className="pointer-events-none absolute -bottom-20 -right-20 h-[300px] w-[300px] rounded-full bg-gold/[0.08]" />
          <p className="relative z-[1] mb-4 text-[0.65rem] uppercase tracking-[0.35em] text-gold">
            Fácil y rápido
          </p>
          <h2 className="relative z-[1] mb-5 font-serif text-[clamp(1.75rem,3.5vw,3rem)] font-light text-ink-dark">
            Reservá tu <em className="font-serif not-italic text-gold">turno online</em>
          </h2>
          <p className="relative z-[1] mx-auto mb-5 max-w-xl text-[0.9rem] font-light text-ink-muted">
            Elegí el servicio, la fecha y el horario que más te convenga.
            <br />
            Sin esperas, sin llamadas.
          </p>
          <p className="relative z-[1] mx-auto mb-8 max-w-lg text-[0.78rem] font-light leading-relaxed text-ink-muted">
            Podés cancelar tu turno con el código que recibís por mail hasta{" "}
            <strong>24 horas antes</strong> del horario reservado. Después de
            ese plazo, escribinos por WhatsApp y lo vemos juntas.
          </p>
          <div className="relative z-[1] flex flex-wrap items-center justify-center gap-5">
            <Link
              href="/reservar"
              className="inline-block rounded-[2px] bg-gold px-11 py-[18px] text-[0.8rem] font-medium uppercase tracking-[0.15em] text-white transition-all hover:-translate-y-0.5 hover:bg-gold-dark"
            >
              ✦ &nbsp;Reservar turno online
            </Link>
            <a
              href={MAPS_URL}
              target="_blank"
              rel="noopener noreferrer"
              className="inline-flex items-center gap-2 rounded-[2px] border border-gold-dark/35 px-6 py-3 text-[0.68rem] font-medium uppercase tracking-[0.18em] text-gold-dark transition-colors hover:bg-white/50"
            >
              Cómo llegar
            </a>
            <a
              href={WA_URL}
              target="_blank"
              rel="noopener noreferrer"
              className="inline-flex items-center gap-2.5 rounded-[2px] bg-[#25D366] px-8 py-4 text-[0.72rem] font-medium uppercase tracking-[0.15em] text-white transition-all hover:-translate-y-0.5 hover:bg-[#1da851]"
            >
              <WaIcon />
              Escribinos por WhatsApp
            </a>
          </div>
        </section>

        <footer className="bg-footer px-6 pb-10 pt-14 text-white md:px-20 md:pt-16">
          <div className="mx-auto grid max-w-6xl grid-cols-1 gap-10 border-b border-gold/15 pb-10 md:grid-cols-3 md:gap-16">
            <div>
              <span className="mb-4 block font-serif text-[1.1rem] font-semibold uppercase tracking-[0.15em] text-gold-light">
                ME Estética
              </span>
              <p className="text-[0.8rem] font-light leading-relaxed text-white/40">
                Tu espacio de belleza y bienestar profesional en la zona de
                Ensenada, Bartolomé Bavio y Magdalena.
              </p>
            </div>
            <div>
              <h4 className="mb-5 font-serif text-base font-normal tracking-wide text-gold-light">
                Servicios
              </h4>
              <ul className="list-none space-y-2.5 text-[0.78rem] tracking-wide text-white/45">
                <li>Depilación Láser</li>
                <li>Faciales</li>
                <li>Uñas & Esculpidas</li>
                <li>Podología</li>
                <li>Coloración</li>
                <li>Alisado</li>
              </ul>
            </div>
            <div>
              <h4 className="mb-5 font-serif text-base font-normal tracking-wide text-gold-light">
                Contacto
              </h4>
              <ul className="list-none space-y-2.5 text-[0.78rem]">
                <li>
                  <a
                    href={MAPS_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-2 text-white/45 transition-colors hover:text-gold-light"
                  >
                    📍 Cómo llegar (mapa)
                  </a>
                </li>
                <li>
                  <a
                    href="https://wa.me/5492215918286"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-2 text-white/45 transition-colors hover:text-gold-light"
                  >
                    📱 +54 9 221 591-8286
                  </a>
                </li>
                <li>
                  <a
                    href="https://instagram.com/mariaemilia_estetica_"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-2 text-white/45 transition-colors hover:text-gold-light"
                  >
                    📸 @mariaemilia_estetica_
                  </a>
                </li>
                <li className="pt-4">
                  <Link
                    href="/reservar"
                    className="inline-block rounded-[2px] bg-gold px-6 py-3 text-[0.68rem] font-medium uppercase tracking-[0.15em] text-white transition-colors hover:bg-gold-dark"
                  >
                    Reservar turno
                  </Link>
                </li>
              </ul>
            </div>
          </div>
          <div className="mx-auto mt-10 flex max-w-6xl flex-col items-center justify-between gap-3 text-[0.68rem] tracking-wide text-white/25 md:flex-row">
            <span>© {new Date().getFullYear()} María Emilia Estética. Todos los derechos reservados.</span>
            <span>Ensenada · Bartolomé Bavio · Magdalena</span>
          </div>
          <div className="mx-auto mt-8 max-w-6xl border-t border-gold/10 pt-6 text-center">
            <a
              href="https://gestion-ya.vercel.app"
              target="_blank"
              rel="noopener noreferrer"
              className="inline-flex items-center gap-2 text-[0.72rem] tracking-wide text-white/35 transition-colors hover:text-gold-light"
            >
              <span>Desarrollado con</span>
              <span className="inline-flex items-center gap-1 font-semibold text-[#b8f542]">
                <span className="h-1.5 w-1.5 rounded-full bg-[#b8f542] shadow-[0_0_8px_rgba(184,245,66,0.6)]" />
                GestiónYa
              </span>
              <span className="opacity-60">→</span>
            </a>
          </div>
        </footer>
      </main>
    </>
  );
}
