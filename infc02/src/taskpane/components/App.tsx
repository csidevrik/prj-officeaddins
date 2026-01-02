// src/taskpane/components/App.tsx
import React, { useMemo, useState } from "react";
import { REQUIRED_TAGS, applyVars, validateTemplate } from "../office-document";

type Vars = Record<string, string>;

export default function App() {
  const initialVars = useMemo<Vars>(() => ({
    ACTA_NUM: "",
    PERIODO_DESDE: "",
    PERIODO_HASTA: "",
    CIUDAD: "Cuenca",
    FECHA_ACTA: "",
    USD_TOTAL: "7120.00",
    USD_CONECTIVIDAD: "2150.00",
    USD_DC: "4970.00",
  }), []);

  const [vars, setVars] = useState<Vars>(initialVars);
  const [status, setStatus] = useState<string>("");
  const [counts, setCounts] = useState<Record<string, number> | null>(null);
  const [busy, setBusy] = useState<boolean>(false);

  function setField(tag: string, value: string) {
    setVars((prev) => ({ ...prev, [tag]: value }));
  }

  async function onValidate() {
    try {
      setBusy(true);
      setStatus("Validando plantilla...");
      const c = await validateTemplate();
      setCounts(c);

      const missing = REQUIRED_TAGS.filter((t) => (c[t] ?? 0) === 0);
      if (missing.length) {
        setStatus(`Faltan tags: ${missing.join(", ")}`);
      } else {
        setStatus("Plantilla OK ✅ (todos los tags existen).");
      }
    } catch (e: any) {
      setStatus(`Error validando: ${e?.message || String(e)}`);
    } finally {
      setBusy(false);
    }
  }

  async function onApply() {
    try {
      // validación mínima para evitar aplicar vacío
      if (!vars.ACTA_NUM.trim() || !vars.DATE_INIT || !vars.DATE_END.trim()) {
        setStatus("Completa ACTA_NUM, PERIODO_DESDE y PERIODO_HASTA.");
        return;
      }

      setBusy(true);
      setStatus("Aplicando datos al documento...");
      await applyVars(vars);
      setStatus("Aplicado ✅ Revisa el documento.");
    } catch (e: any) {
      setStatus(`Error aplicando: ${e?.message || String(e)}`);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={{ fontFamily: "Segoe UI, Arial", padding: 12 }}>
      <h2 style={{ marginTop: 0 }}>Actas SD-WAN</h2>

      <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
        <button onClick={onValidate} disabled={busy}>Validar plantilla</button>
        <button onClick={onApply} disabled={busy}>Aplicar</button>
      </div>

      <div style={{ marginBottom: 12, fontSize: 13 }}>
        {status && <div><b>Estado:</b> {status}</div>}
      </div>

      <hr />

      {REQUIRED_TAGS.map((tag) => (
        <div key={tag} style={{ margin: "10px 0" }}>
          <label style={{ display: "block", fontSize: 12, opacity: 0.85 }}>{tag}</label>
          <input
            style={{ width: "100%", padding: 6, boxSizing: "border-box" }}
            value={vars[tag] ?? ""}
            onChange={(e) => setField(tag, e.target.value)}
            placeholder={tag}
          />
          {counts && (
            <div style={{ fontSize: 12, opacity: 0.8, marginTop: 2 }}>
              Ocurrencias en documento: {counts[tag] ?? 0}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}
