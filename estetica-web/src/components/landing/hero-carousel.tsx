"use client";

import useEmblaCarousel from "embla-carousel-react";
import Image from "next/image";
import { useCallback, useEffect, useRef, useState } from "react";
import { HERO_SLIDES, type HeroSlide } from "@/lib/landing-media";

function SlideVideo({
  active,
  slide,
}: {
  active: boolean;
  slide: Extract<HeroSlide, { kind: "video" }>;
}) {
  const ref = useRef<HTMLVideoElement>(null);

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    if (active) {
      void el.play().catch(() => {});
    } else {
      el.pause();
      try {
        el.currentTime = 0;
      } catch {
        /* noop */
      }
    }
  }, [active]);

  return (
    <video
      ref={ref}
      className="h-full w-full object-cover"
      src={slide.src}
      poster={slide.poster}
      muted
      playsInline
      loop
      preload="metadata"
      aria-label={slide.alt}
    />
  );
}

export function HeroCarousel() {
  const [emblaRef, emblaApi] = useEmblaCarousel({ loop: true, align: "start" });
  const [selected, setSelected] = useState(0);

  const onSelect = useCallback(() => {
    if (!emblaApi) return;
    setSelected(emblaApi.selectedScrollSnap());
  }, [emblaApi]);

  useEffect(() => {
    if (!emblaApi) return;
    onSelect();
    emblaApi.on("select", onSelect);
    emblaApi.on("reInit", onSelect);
    return () => {
      emblaApi.off("select", onSelect);
      emblaApi.off("reInit", onSelect);
    };
  }, [emblaApi, onSelect]);

  useEffect(() => {
    if (!emblaApi) return;
    const reduce = window.matchMedia("(prefers-reduced-motion: reduce)");
    if (reduce.matches) return;
    const id = window.setInterval(() => {
      emblaApi.scrollNext();
    }, 6000);
    return () => window.clearInterval(id);
  }, [emblaApi]);

  const scrollPrev = useCallback(() => emblaApi?.scrollPrev(), [emblaApi]);
  const scrollNext = useCallback(() => emblaApi?.scrollNext(), [emblaApi]);
  const scrollTo = useCallback(
    (i: number) => emblaApi?.scrollTo(i),
    [emblaApi]
  );

  return (
    <div className="relative h-full min-h-[380px] w-full md:min-h-full">
      <div className="absolute inset-0 bg-gradient-to-br from-rose/40 via-lilac/35 to-[#edd9f5]/50" />

      <div className="overflow-hidden" ref={emblaRef}>
        <div className="flex h-full">
          {HERO_SLIDES.map((slide, i) => (
            <div
              className="relative min-h-[380px] w-0 flex-[0_0_100%] md:min-h-full"
              key={`${slide.kind}-${i}`}
            >
              {slide.kind === "image" ? (
                <Image
                  src={slide.src}
                  alt={slide.alt}
                  fill
                  priority={i === 0}
                  sizes="(max-width: 768px) 100vw, 50vw"
                  className="object-cover"
                />
              ) : (
                <SlideVideo active={selected === i} slide={slide} />
              )}
              <div className="pointer-events-none absolute inset-x-0 bottom-0 h-1/3 bg-gradient-to-t from-black/35 to-transparent" />
            </div>
          ))}
        </div>
      </div>

      <button
        type="button"
        onClick={scrollPrev}
        className="absolute left-1 top-1/2 z-10 -translate-y-1/2 rounded border border-white/40 bg-black/25 px-2 py-2.5 text-base text-white backdrop-blur-sm transition hover:bg-black/40 sm:py-3 sm:text-lg md:left-2"
        aria-label="Anterior"
      >
        ‹
      </button>
      <button
        type="button"
        onClick={scrollNext}
        className="absolute right-1 top-1/2 z-10 -translate-y-1/2 rounded border border-white/40 bg-black/25 px-2 py-2.5 text-base text-white backdrop-blur-sm transition hover:bg-black/40 sm:py-3 sm:text-lg md:right-2"
        aria-label="Siguiente"
      >
        ›
      </button>

      <div
        className="absolute bottom-4 left-0 right-0 z-10 flex justify-center gap-2"
        role="tablist"
        aria-label="Indicadores del carrusel"
      >
        {HERO_SLIDES.map((_, i) => (
          <button
            key={i}
            type="button"
            role="tab"
            aria-selected={selected === i}
            aria-label={`Ir a la diapositiva ${i + 1}`}
            onClick={() => scrollTo(i)}
            className={`h-2 rounded-full transition-all ${
              selected === i
                ? "w-8 bg-white"
                : "w-2 bg-white/50 hover:bg-white/75"
            }`}
          />
        ))}
      </div>
    </div>
  );
}
