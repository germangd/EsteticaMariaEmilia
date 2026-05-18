/** Clases reutilizables: reservas online + panel admin/staff. */

export const uiLabel =
  "mb-1.5 block text-xs font-semibold uppercase tracking-wide text-ink-dark";

export const uiHint =
  "text-xs font-medium leading-relaxed text-ink";

export const uiInput =
  "w-full rounded-sm border border-gold/55 bg-white px-3 py-2.5 text-sm font-medium text-ink-dark shadow-sm outline-none transition placeholder:font-normal placeholder:text-ink-muted/75 focus:border-gold focus:bg-white focus:ring-2 focus:ring-gold/35";

export const uiSelect = `${uiInput} cursor-pointer`;

/** Panel cálido (usa token `panel` de tailwind — no color arbitrario). */
const uiWarmPanel =
  "rounded-sm border border-gold/60 bg-panel shadow-md ring-1 ring-gold/25";

export const uiCard = `${uiWarmPanel} p-5 md:p-6`;

export const uiCardSoft = `${uiWarmPanel} p-6`;

export const uiReservarForm = `${uiWarmPanel} p-6 md:p-7`;

export const uiBtnPrimary =
  "inline-flex items-center justify-center rounded-sm bg-gold px-5 py-2.5 text-[0.72rem] font-semibold uppercase tracking-wider text-white shadow-sm transition hover:bg-gold-dark hover:shadow focus:outline-none focus:ring-2 focus:ring-gold/40 disabled:cursor-not-allowed disabled:opacity-50";

export const uiBtnPrimaryFull =
  "w-full rounded-sm bg-gold py-3 text-xs font-bold uppercase tracking-widest text-white shadow-sm transition hover:bg-gold-dark focus:outline-none focus:ring-2 focus:ring-gold/40 disabled:cursor-not-allowed disabled:opacity-50";

export const uiBtnSecondary =
  "inline-flex items-center justify-center rounded-sm border border-gold/55 bg-white px-5 py-2.5 text-[0.72rem] font-semibold uppercase tracking-wider text-gold-dark shadow-sm transition hover:border-gold hover:bg-cream focus:outline-none focus:ring-2 focus:ring-gold/30";

export const uiBtnDark =
  "inline-flex w-full items-center justify-center rounded-sm bg-ink-dark py-3 text-xs font-bold uppercase tracking-widest text-white shadow-sm transition hover:bg-ink focus:outline-none focus:ring-2 focus:ring-ink/30 disabled:opacity-50";

export const uiBtnDanger =
  "inline-flex items-center justify-center rounded-sm border border-red-400/70 bg-white px-4 py-2 text-[0.72rem] font-semibold uppercase tracking-wider text-red-800 shadow-sm transition hover:bg-red-50 disabled:opacity-50";

export const uiTableWrap = `overflow-x-auto ${uiWarmPanel}`;

export const uiTableHead =
  "border-b border-gold/45 bg-panel-head text-[0.65rem] font-bold uppercase tracking-[0.14em] text-ink-dark";

export const uiPanelKicker =
  "mb-1 text-[0.65rem] font-bold uppercase tracking-[0.25em] text-gold-dark";

export const uiPanelTitle =
  "font-serif text-3xl font-medium text-ink-dark md:text-4xl";

export const uiPanelDesc = "mt-2 text-sm font-medium text-ink";

export const uiSectionTitle =
  "mb-4 font-serif text-xl font-semibold text-ink-dark";

export const uiSubsectionTitle =
  "mb-2 font-serif text-lg font-semibold text-ink-dark";

export const uiTabBar =
  "mb-8 flex rounded-sm border border-gold/60 bg-panel p-1 shadow-sm ring-1 ring-gold/20";

export const uiTabActive =
  "flex-1 rounded-sm bg-gold py-2.5 text-xs font-bold uppercase tracking-wide text-white shadow-sm";

export const uiTabInactive =
  "flex-1 rounded-sm py-2.5 text-xs font-semibold uppercase tracking-wide text-ink transition-colors hover:bg-cream hover:text-ink-dark";

export const uiTimeSlot =
  "rounded-full border border-gold/50 bg-white px-3 py-1.5 text-sm font-medium text-ink shadow-sm transition-colors hover:border-gold hover:bg-cream";

export const uiTimeSlotActive =
  "rounded-full border border-gold-dark bg-gold px-3 py-1.5 text-sm font-semibold text-white shadow-sm";
