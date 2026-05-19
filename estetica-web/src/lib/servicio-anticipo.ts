/** Monto de anticipo en ARS a partir del precio de referencia y el % configurado. */
export function calcularAnticipoPesos(
  precioPesos: number,
  anticipoRequerido: boolean,
  anticipoPorcentaje: number
): number {
  if (!anticipoRequerido || precioPesos <= 0) return 0;
  const pct = Math.min(100, Math.max(1, Math.round(anticipoPorcentaje)));
  return Math.round((precioPesos * pct) / 100);
}

export function normalizarAnticipoPorcentaje(
  requerido: boolean,
  porcentaje: number
): number {
  if (!requerido) return 0;
  return Math.min(100, Math.max(1, Math.round(porcentaje)));
}

export function etiquetaAnticipo(
  precioPesos: number,
  anticipoRequerido: boolean,
  anticipoPorcentaje: number
): string | null {
  if (!anticipoRequerido) return null;
  const pct = normalizarAnticipoPorcentaje(true, anticipoPorcentaje);
  const monto = calcularAnticipoPesos(precioPesos, true, pct);
  if (monto <= 0) {
    return `Anticipo ${pct}% (definí precio de referencia para calcular el monto)`;
  }
  return `Anticipo ${pct}%`;
}
