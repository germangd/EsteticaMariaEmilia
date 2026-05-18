"use client";

import { useEffect, useState } from "react";
import { fmtPesos } from "@/lib/fmt-pesos";

type Props = {
  value: number;
  disabled?: boolean;
  onSave: (precioPesos: number) => Promise<boolean>;
};

export function PrecioInlineEditor({ value, disabled, onSave }: Props) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(String(value));
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (!editing) setDraft(String(value));
  }, [value, editing]);

  async function commit() {
    const n = Math.max(0, Math.round(Number(draft) || 0));
    setSaving(true);
    try {
      const ok = await onSave(n);
      if (ok) setEditing(false);
    } finally {
      setSaving(false);
    }
  }

  if (!editing) {
    return (
      <button
        type="button"
        disabled={disabled || saving}
        title="Clic para cambiar el precio"
        onClick={() => {
          setDraft(String(value));
          setEditing(true);
        }}
        className="tabular-nums text-left underline decoration-gold/40 decoration-dotted underline-offset-2 hover:text-gold-dark disabled:opacity-50"
      >
        {value > 0 ? fmtPesos(value) : "Sin precio"}
      </button>
    );
  }

  return (
    <input
      type="number"
      min={0}
      step={1}
      autoFocus
      disabled={saving}
      className="w-28 rounded-sm border border-gold/35 bg-white px-2 py-1 text-sm tabular-nums"
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={() => void commit()}
      onKeyDown={(e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          void commit();
        }
        if (e.key === "Escape") {
          setDraft(String(value));
          setEditing(false);
        }
      }}
    />
  );
}
