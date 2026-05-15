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
      className="h-full w-full object-cover object-[center_22%] md:object-center"
      src={slide.src}
      poster={slide.poster || undefined}
      muted
      playsInline
      loop
      preload="metadata"
      aria-label={slide.alt}
    />
  );
}

export function HeroCarousel({ slides }: { slides: HeroSlide[] }) {
  const list = slides.length > 0 ? slides : HERO_SLIDES;
  const slideKey = list.map((s) => `${s.kind}:${s.src}`).join("|");

  const [emblaRef, emblaApi] = useEmblaCarousel({
    loop: true,
    align: "start",
    skipSnaps: false,
    dragFree: false,
  });
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
    emblaApi?.reInit();
  }, [emblaApi, slideKey]);

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
    <div className="relative h-full w-full min-h-0 md:min-h-full">
      <div className="absolute inset-0 bg-gradient-to-br from-rose/40 via-lilac/35 to-[#edd9f5]/50" />

      <div
        className="h-full min-h-0 overflow-hidden [-webkit-tap-highlight-color:transparent]"
        ref={emblaRef}
      >
        <div className="flex h-full min-h-0">
          {list.map((slide, i) => (
            <div
              className="relative h-full min-h-0 w-0 flex-[0_0_100%]"
              key={`${slide.kind}-${slide.src}-${i}`}
            >
              {slide.kind === "image" ? (
                <Image
                  src={slide.src}
                  alt={slide.alt}
                  fill
                  priority={i === 0}
                  sizes="(max-width: 768px) 100vw, (max-width: 1280px) 64vw, 58vw"
                  className="object-cover object-[center_22%] md:object-center"
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
        className="absolute left-1 top-1/2 z-10 flex h-11 min-w-11 -translate-y-1/2 items-center justify-center rounded-md border border-white/40 bg-black/30 px-0 text-lg leading-none text-white shadow-sm backdrop-blur-sm transition active:scale-95 hover:bg-black/45 sm:h-12 sm:min-w-12 sm:text-xl md:left-2"
        aria-label="Anterior"
      >
        ‹
      </button>
      <button
        type="button"
        onClick={scrollNext}
        className="absolute right-1 top-1/2 z-10 flex h-11 min-w-11 -translate-y-1/2 items-center justify-center rounded-md border border-white/40 bg-black/30 px-0 text-lg leading-none text-white shadow-sm backdrop-blur-sm transition active:scale-95 hover:bg-black/45 sm:h-12 sm:min-w-12 sm:text-xl md:right-2"
        aria-label="Siguiente"
      >
        ›
      </button>

      <div
        className="absolute bottom-0 left-0 right-0 z-10 flex justify-center px-12 pb-[max(0.75rem,env(safe-area-inset-bottom))] pt-2 md:bottom-1 md:px-14 md:pb-4"
        role="tablist"
        aria-label="Indicadores del carrusel"
      >
        <div className="flex max-w-full gap-1.5 overflow-x-auto overflow-y-visible py-2 [-ms-overflow-style:none] [scrollbar-width:none] md:gap-2 [&::-webkit-scrollbar]:hidden">
          {list.map((_, i) => (
            <button
              key={i}
              type="button"
              role="tab"
              aria-selected={selected === i}
              aria-label={`Ir a la diapositiva ${i + 1}`}
              onClick={() => scrollTo(i)}
              className={`shrink-0 rounded-full transition-all ${
                selected === i
                  ? "h-2 w-7 bg-white md:w-8"
                  : "h-2 w-2 bg-white/50 hover:bg-white/75 active:bg-white/90"
              }`}
            />
          ))}
        </div>
      </div>
    </div>
  );
}
