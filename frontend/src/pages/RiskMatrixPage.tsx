import { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type RiskMatrix, type RiskItem, type Template } from "../api/client";

const RISK_COLORS: Record<string, string> = {
  low: "#2F9E44",
  medium: "#F76707",
  high: "#E03131",
};

const CELL_BG: Record<string, string> = {
  low: "#D3F9D8",
  medium: "#FFE8CC",
  high: "#FFE3E3",
};

const CELL_BORDER: Record<string, string> = {
  low: "#2F9E4450",
  medium: "#F7670750",
  high: "#E0313150",
};

function riskLevel(score: number): "low" | "medium" | "high" {
  if (score <= 6) return "low";
  if (score <= 14) return "medium";
  return "high";
}

const statusLabel: Record<string, string> = {
  open: "Åpen",
  mitigated: "Mitigert",
  closed: "Lukket",
};

type RiskStatus = "open" | "mitigated" | "closed";

interface RiskForm {
  description: string;
  probability: number;
  consequence: number;
  mitigation: string;
  owner: string;
  status: RiskStatus;
  fagomrade: string;
  risk_owner: string;
  residual_probability: number;
  residual_consequence: number;
}

const emptyRisk: RiskForm = {
  description: "",
  probability: 3,
  consequence: 3,
  mitigation: "",
  owner: "",
  status: "open",
  fagomrade: "",
  risk_owner: "",
  residual_probability: 1,
  residual_consequence: 1,
};

export default function RiskMatrixPage() {
  const { projectId, matrixId } = useParams<{ projectId: string; matrixId: string }>();
  const navigate = useNavigate();
  const [matrix, setMatrix] = useState<RiskMatrix | null>(null);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState<RiskForm>(emptyRisk);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);
  const [residualEnabled, setResidualEnabled] = useState(false);
  const [activeTab, setActiveTab] = useState<"table" | "heatmap">("table");

  const [templateModal, setTemplateModal] = useState(false);
  const [templateMode, setTemplateMode] = useState<"new" | "existing">("new");
  const [templateName, setTemplateName] = useState("");
  const [existingTemplates, setExistingTemplates] = useState<Template[]>([]);
  const [selectedExistingId, setSelectedExistingId] = useState("");
  const [templateSaving, setTemplateSaving] = useState(false);

  useEffect(() => {
    if (projectId && matrixId) {
      api.riskMatrices.get(projectId, matrixId).then(setMatrix);
    }
  }, [projectId, matrixId]);

  function openEdit(risk: RiskItem) {
    const rp = risk.residual_probability ?? 0;
    const rc = risk.residual_consequence ?? 0;
    setForm({
      description: risk.description,
      probability: risk.probability,
      consequence: risk.consequence,
      mitigation: risk.mitigation ?? "",
      owner: risk.owner ?? "",
      status: risk.status,
      fagomrade: risk.fagomrade ?? "",
      risk_owner: risk.risk_owner ?? "",
      residual_probability: rp > 0 ? rp : 1,
      residual_consequence: rc > 0 ? rc : 1,
    });
    setResidualEnabled(rp > 0 || rc > 0);
    setEditingId(risk.id);
    setShowForm(true);
  }

  async function handleSave() {
    if (!projectId || !matrixId || !form.description) return;
    setSaving(true);
    try {
      const data = {
        ...form,
        mitigation: form.mitigation || null,
        owner: form.owner || null,
        fagomrade: form.fagomrade || null,
        risk_owner: form.risk_owner || null,
        residual_probability: residualEnabled ? form.residual_probability : null,
        residual_consequence: residualEnabled ? form.residual_consequence : null,
      };
      if (editingId) {
        const updated = await api.riskMatrices.updateRisk(projectId, matrixId, editingId, data);
        setMatrix((prev) => prev ? { ...prev, risks: prev.risks.map((r) => r.id === editingId ? updated : r) } : prev);
      } else {
        const created = await api.riskMatrices.addRisk(projectId, matrixId, data);
        setMatrix((prev) => prev ? { ...prev, risks: [...prev.risks, created] } : prev);
      }
      setShowForm(false);
      setForm(emptyRisk);
      setEditingId(null);
    } finally {
      setSaving(false);
    }
  }

  async function openTemplateModal() {
    const ts = await api.templates.list("risk_matrix");
    setExistingTemplates(ts);
    setTemplateMode("new");
    setTemplateName(matrix?.title ?? "");
    setSelectedExistingId(ts[0]?.id ?? "");
    setTemplateModal(true);
  }

  async function saveAsTemplate() {
    if (!matrix) return;
    setTemplateSaving(true);
    try {
      const data = JSON.stringify({ risks: matrix.risks.map(({ description, probability, consequence, mitigation, owner, status }) => ({ description, probability, consequence, mitigation, owner, status })) });
      if (templateMode === "new") {
        await api.templates.create({ name: templateName.trim() || matrix.title, type: "risk_matrix", data });
      } else {
        await api.templates.update(selectedExistingId, { name: existingTemplates.find(t => t.id === selectedExistingId)?.name ?? templateName, data });
      }
      setTemplateModal(false);
    } finally {
      setTemplateSaving(false);
    }
  }

  async function handleDelete(riskId: string) {
    if (!projectId || !matrixId) return;
    await api.riskMatrices.deleteRisk(projectId, matrixId, riskId);
    setMatrix((prev) => prev ? { ...prev, risks: prev.risks.filter((r) => r.id !== riskId) } : prev);
  }

  if (!matrix) return <div style={{ padding: "2rem" }}>Laster...</div>;

  return (
    <div style={{ maxWidth: 1400, margin: "0 auto", padding: "2rem" }}>
      <button
        onClick={() => navigate(`/projects/${projectId}`)}
        style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", marginBottom: "1rem" }}
      >
        ← Tilbake til prosjekt
      </button>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
        <h1 className="bf-h2">{matrix.title}</h1>
        <div style={{ display: "flex", gap: "0.5rem" }}>
          <Button variant="outline" onClick={openTemplateModal}>Lagre som mal</Button>
          <Button variant="filled" onClick={() => { setShowForm(true); setEditingId(null); setForm(emptyRisk); setResidualEnabled(false); }}>+ Legg til risiko</Button>
        </div>
      </div>

      <Modal isOpen={templateModal} onRequestClose={() => setTemplateModal(false)} header="Lagre som mal">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.5rem" }}>
            {(["new", "existing"] as const).map((mode) => (
              <button key={mode} onClick={() => setTemplateMode(mode)}
                style={{ padding: "0.75rem", borderRadius: 8, border: `2px solid ${templateMode === mode ? "#E03131" : "var(--bfc-base-dimmed)"}`, background: templateMode === mode ? "#E0313118" : "var(--bfc-base-3)", cursor: "pointer", fontWeight: 600, fontSize: "0.9rem", color: templateMode === mode ? "#E03131" : "var(--bfc-base-c-1)", transition: "all 0.15s" }}>
                {mode === "new" ? "Ny mal" : "Oppdater eksisterende"}
              </button>
            ))}
          </div>
          {templateMode === "new" ? (
            <Input label="Navn på malen" value={templateName} onChange={(e) => setTemplateName(e.target.value)} placeholder="f.eks. Standard risikomatrise" autoFocus />
          ) : (
            <div>
              <label className="bf-label">Velg mal å oppdatere</label>
              {existingTemplates.length === 0 ? (
                <p style={{ color: "var(--bfc-base-c-2)", fontSize: "0.9rem" }}>Ingen eksisterende maler</p>
              ) : (
                <select value={selectedExistingId} onChange={(e) => setSelectedExistingId(e.target.value)} style={{ width: "100%", padding: "0.5rem", borderRadius: 4, border: "1px solid var(--bfc-base-dimmed)", marginTop: 4 }}>
                  {existingTemplates.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                </select>
              )}
            </div>
          )}
          <p style={{ margin: 0, color: "var(--bfc-base-c-2)", fontSize: "0.85rem" }}>{matrix.risks.length} risikoer vil bli lagret i malen.</p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setTemplateModal(false)}>Avbryt</Button>
            <Button variant="filled" onClick={saveAsTemplate} state={templateSaving || (templateMode === "existing" && !selectedExistingId) ? "inactive" : "default"}>
              {templateSaving ? "Lagrer..." : "Lagre mal"}
            </Button>
          </div>
        </div>
      </Modal>

      {showForm && (
        <div style={{ background: "var(--bfc-base-3)", borderRadius: 8, padding: "1.5rem", marginBottom: "1.5rem", display: "grid", gap: "1rem" }}>
          <h3 className="bf-h4">{editingId ? "Rediger risiko" : "Ny risiko"}</h3>

          <Input label="Beskrivelse" value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} />

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1rem" }}>
            <Input label="Fagområde" value={form.fagomrade} onChange={(e) => setForm({ ...form, fagomrade: e.target.value })} placeholder="f.eks. Workplace, Network" />
            <Input label="Risiko eier" value={form.risk_owner} onChange={(e) => setForm({ ...form, risk_owner: e.target.value })} placeholder="Navn eller rolle" />
          </div>

          <div style={{ borderTop: "1px solid var(--bfc-base-dimmed)", paddingTop: "0.75rem" }}>
            <div style={{ fontSize: "0.8rem", fontWeight: 600, color: "var(--bfc-base-c-2)", marginBottom: "0.75rem", textTransform: "uppercase", letterSpacing: "0.05em" }}>
              Risikovurdering (før tiltak)
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1rem" }}>
              <div>
                <label className="bf-label">Sannsynlighet (1–5): {form.probability}</label>
                <input type="range" min={1} max={5} value={form.probability}
                  onChange={(e) => setForm({ ...form, probability: Number(e.target.value) })}
                  style={{ width: "100%", accentColor: "#1971C2", cursor: "pointer" }} />
              </div>
              <div>
                <label className="bf-label">Konsekvens (1–5): {form.consequence}</label>
                <input type="range" min={1} max={5} value={form.consequence}
                  onChange={(e) => setForm({ ...form, consequence: Number(e.target.value) })}
                  style={{ width: "100%", accentColor: "#1971C2", cursor: "pointer" }} />
              </div>
            </div>
            <div style={{ marginTop: "0.5rem", fontSize: "0.85rem", color: "var(--bfc-base-c-2)" }}>
              Score:{" "}
              <span style={{ background: RISK_COLORS[riskLevel(form.probability * form.consequence)], color: "#fff", borderRadius: 4, padding: "1px 8px", fontWeight: 600 }}>
                {form.probability * form.consequence}
              </span>
            </div>
          </div>

          <Input label="Tiltak" value={form.mitigation} onChange={(e) => setForm({ ...form, mitigation: e.target.value })} />
          <Input label="Ansvarlig (tiltak)" value={form.owner} onChange={(e) => setForm({ ...form, owner: e.target.value })} />

          <div style={{ borderTop: "1px solid var(--bfc-base-dimmed)", paddingTop: "0.75rem" }}>
            <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "0.75rem" }}>
              <div style={{ fontSize: "0.8rem", fontWeight: 600, color: "var(--bfc-base-c-2)", textTransform: "uppercase", letterSpacing: "0.05em" }}>
                Restrisiko (etter tiltak)
              </div>
              <label style={{ display: "flex", alignItems: "center", gap: "0.4rem", cursor: "pointer", fontSize: "0.85rem", color: "var(--bfc-base-c-2)", fontWeight: 400, userSelect: "none" }}>
                <input
                  type="checkbox"
                  checked={residualEnabled}
                  onChange={(e) => {
                    setResidualEnabled(e.target.checked);
                    if (!e.target.checked) setForm((f) => ({ ...f, residual_probability: 1, residual_consequence: 1 }));
                  }}
                  style={{ cursor: "pointer", accentColor: "#1971C2" }}
                />
                Vurdert
              </label>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1rem", opacity: residualEnabled ? 1 : 0.4, pointerEvents: residualEnabled ? "auto" : "none" }}>
              <div>
                <label className="bf-label">Oppdatert sannsynlighet (1–5): {residualEnabled ? form.residual_probability : "–"}</label>
                <input type="range" min={1} max={5} value={form.residual_probability}
                  onChange={(e) => setForm({ ...form, residual_probability: Number(e.target.value) })}
                  style={{ width: "100%", accentColor: "#1971C2", cursor: "pointer" }} />
              </div>
              <div>
                <label className="bf-label">Oppdatert konsekvens (1–5): {residualEnabled ? form.residual_consequence : "–"}</label>
                <input type="range" min={1} max={5} value={form.residual_consequence}
                  onChange={(e) => setForm({ ...form, residual_consequence: Number(e.target.value) })}
                  style={{ width: "100%", accentColor: "#1971C2", cursor: "pointer" }} />
              </div>
            </div>
            {residualEnabled && (
              <div style={{ marginTop: "0.5rem", fontSize: "0.85rem", color: "var(--bfc-base-c-2)" }}>
                Restrisiko-score:{" "}
                <span style={{ background: RISK_COLORS[riskLevel(form.residual_probability * form.residual_consequence)], color: "#fff", borderRadius: 4, padding: "1px 8px", fontWeight: 600 }}>
                  {form.residual_probability * form.residual_consequence}
                </span>
              </div>
            )}
          </div>

          <div>
            <label className="bf-label">Status</label>
            <select
              value={form.status}
              onChange={(e) => setForm({ ...form, status: e.target.value as RiskStatus })}
              style={{ width: "100%", padding: "0.5rem", borderRadius: 4, border: "1px solid var(--bfc-base-dimmed)" }}
            >
              <option value="open">Åpen</option>
              <option value="mitigated">Mitigert</option>
              <option value="closed">Lukket</option>
            </select>
          </div>
          <div style={{ display: "flex", gap: "0.5rem" }}>
            <Button onClick={() => setShowForm(false)}>Avbryt</Button>
            <Button variant="filled" onClick={handleSave} state={saving || !form.description ? "inactive" : "default"}>
              {saving ? "Lagrer..." : "Lagre"}
            </Button>
          </div>
        </div>
      )}

      {/* Tab switcher */}
      {matrix.risks.length > 0 && (
        <div style={{ display: "flex", gap: "0.25rem", marginBottom: "1.25rem", borderBottom: "1px solid var(--bfc-base-dimmed)", paddingBottom: 0 }}>
          {(["table", "heatmap"] as const).map((tab) => (
            <button
              key={tab}
              onClick={() => setActiveTab(tab)}
              style={{
                background: "none", border: "none", cursor: "pointer",
                padding: "0.6rem 1.1rem",
                fontSize: "0.9rem", fontWeight: activeTab === tab ? 600 : 400,
                color: activeTab === tab ? "#E03131" : "var(--bfc-base-c-2)",
                borderBottom: activeTab === tab ? "2px solid #E03131" : "2px solid transparent",
                marginBottom: -1, transition: "color 0.15s",
              }}
            >
              {tab === "table" ? "Tabell" : "Heatmap"}
            </button>
          ))}
        </div>
      )}

      {matrix.risks.length === 0 ? (
        <p style={{ color: "var(--bfc-base-c-2)" }}>Ingen risikoer lagt til ennå.</p>
      ) : activeTab === "table" ? (
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ textAlign: "left", borderBottom: "2px solid var(--bfc-base-dimmed)", fontSize: "0.85rem" }}>
              <th style={{ padding: "0.5rem" }}>Beskrivelse</th>
              <th style={{ padding: "0.5rem" }}>Fagområde</th>
              <th style={{ padding: "0.5rem" }}>Risiko eier</th>
              <th style={{ padding: "0.5rem" }}>S</th>
              <th style={{ padding: "0.5rem" }}>K</th>
              <th style={{ padding: "0.5rem" }}>Score</th>
              <th style={{ padding: "0.5rem" }}>Tiltak</th>
              <th style={{ padding: "0.5rem" }}>Ansvarlig</th>
              <th style={{ padding: "0.5rem" }}>Restrisiko</th>
              <th style={{ padding: "0.5rem" }}>Status</th>
              <th style={{ padding: "0.5rem" }}></th>
            </tr>
          </thead>
          <tbody>
            {matrix.risks.map((risk) => {
              const level = riskLevel(risk.risk_score);
              const residualLevel = risk.residual_score ? riskLevel(risk.residual_score) : null;
              return (
                <tr key={risk.id} style={{ borderBottom: "1px solid var(--bfc-base-dimmed)" }}>
                  <td style={{ padding: "0.5rem" }}>{risk.description}</td>
                  <td style={{ padding: "0.5rem", color: risk.fagomrade ? "inherit" : "var(--bfc-base-c-3)" }}>
                    {risk.fagomrade ?? "–"}
                  </td>
                  <td style={{ padding: "0.5rem", color: risk.risk_owner ? "inherit" : "var(--bfc-base-c-3)" }}>
                    {risk.risk_owner ?? "–"}
                  </td>
                  <td style={{ padding: "0.5rem" }}>{risk.probability}</td>
                  <td style={{ padding: "0.5rem" }}>{risk.consequence}</td>
                  <td style={{ padding: "0.5rem" }}>
                    <span style={{ background: RISK_COLORS[level], color: "#fff", borderRadius: 4, padding: "2px 8px", fontWeight: 600 }}>
                      {risk.risk_score}
                    </span>
                  </td>
                  <td style={{ padding: "0.5rem" }}>{risk.mitigation ?? "–"}</td>
                  <td style={{ padding: "0.5rem" }}>{risk.owner ?? "–"}</td>
                  <td style={{ padding: "0.5rem" }}>
                    {residualLevel && risk.residual_score ? (
                      <div>
                        <span style={{ background: RISK_COLORS[residualLevel], color: "#fff", borderRadius: 4, padding: "2px 8px", fontWeight: 600 }}>
                          {risk.residual_score}
                        </span>
                        <div style={{ fontSize: "0.72rem", color: "var(--bfc-base-c-3)", marginTop: 2 }}>
                          S{risk.residual_probability} × K{risk.residual_consequence}
                        </div>
                      </div>
                    ) : (
                      <span style={{ color: "var(--bfc-base-c-3)" }}>–</span>
                    )}
                  </td>
                  <td style={{ padding: "0.5rem" }}>{statusLabel[risk.status]}</td>
                  <td style={{ padding: "0.5rem", display: "flex", gap: "0.25rem" }}>
                    <Button onClick={() => openEdit(risk)}>Rediger</Button>
                    <Button state="alert" onClick={() => handleDelete(risk.id)}>Slett</Button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      ) : (
        <Heatmap risks={matrix.risks} onEdit={openEdit} />
      )}
    </div>
  );
}

// ─── Heatmap ──────────────────────────────────────────────────────────────────

const CELL_W = 118;
const CELL_H = 88;

function Heatmap({ risks, onEdit }: { risks: RiskItem[]; onEdit: (r: RiskItem) => void }) {
  // Index risks by cell key "P,C"
  const byCell: Record<string, { original: RiskItem[]; residual: RiskItem[] }> = {};
  for (let p = 1; p <= 5; p++) {
    for (let c = 1; c <= 5; c++) {
      byCell[`${p},${c}`] = { original: [], residual: [] };
    }
  }
  for (const r of risks) {
    byCell[`${r.probability},${r.consequence}`].original.push(r);
    if (r.residual_probability && r.residual_consequence) {
      byCell[`${r.residual_probability},${r.residual_consequence}`].residual.push(r);
    }
  }

  return (
    <div style={{ overflowX: "auto" }}>
      <div style={{ display: "inline-flex", gap: 0, alignItems: "stretch", minWidth: "fit-content" }}>
        {/* Y-axis label */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", width: 28, flexShrink: 0 }}>
          <span style={{
            writingMode: "vertical-rl", transform: "rotate(180deg)",
            fontSize: "0.78rem", fontWeight: 600, color: "var(--bfc-base-c-2)",
            letterSpacing: "0.04em", textTransform: "uppercase",
          }}>
            Sannsynlighet
          </span>
        </div>

        <div>
          {/* Grid rows: P=5 top, P=1 bottom */}
          {[5, 4, 3, 2, 1].map((p) => (
            <div key={p} style={{ display: "flex", alignItems: "stretch", gap: 3, marginBottom: 3 }}>
              {/* Row label */}
              <div style={{ width: 20, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "0.82rem", fontWeight: 700, color: "var(--bfc-base-c-1)", flexShrink: 0 }}>
                {p}
              </div>

              {[1, 2, 3, 4, 5].map((c) => {
                const score = p * c;
                const level = riskLevel(score);
                const cell = byCell[`${p},${c}`];
                const MAX_SHOW = 3;
                const shownOriginal = cell.original.slice(0, MAX_SHOW);
                const overflow = cell.original.length - MAX_SHOW;

                return (
                  <div key={c} style={{
                    width: CELL_W, minHeight: CELL_H,
                    background: CELL_BG[level],
                    border: `1px solid ${CELL_BORDER[level]}`,
                    borderRadius: 6, padding: "5px 5px 16px",
                    position: "relative",
                    display: "flex", flexWrap: "wrap",
                    gap: 3, alignContent: "flex-start",
                  }}>
                    {shownOriginal.map((r) => (
                      <RiskChip key={r.id} label={r.description} color={RISK_COLORS[riskLevel(r.risk_score)]} onClick={() => onEdit(r)} title={`${r.description} (S${r.probability}×K${r.consequence}=${r.risk_score})`} />
                    ))}
                    {overflow > 0 && (
                      <span style={{ fontSize: "0.68rem", color: "var(--bfc-base-c-2)", fontWeight: 600, alignSelf: "center" }}>
                        +{overflow}
                      </span>
                    )}
                    {/* Residual chips */}
                    {cell.residual.map((r) => (
                      <RiskChip
                        key={`res-${r.id}`}
                        label={r.description}
                        color={RISK_COLORS[riskLevel(r.residual_score ?? 0)]}
                        onClick={() => onEdit(r)}
                        title={`Restrisiko: ${r.description} (S${r.residual_probability}×K${r.residual_consequence}=${r.residual_score})`}
                        residual
                      />
                    ))}
                    {/* Score badge */}
                    <span style={{
                      position: "absolute", bottom: 3, right: 5,
                      fontSize: "0.62rem", fontWeight: 700,
                      color: RISK_COLORS[level], opacity: 0.7,
                    }}>
                      {score}
                    </span>
                  </div>
                );
              })}
            </div>
          ))}

          {/* X-axis labels */}
          <div style={{ display: "flex", alignItems: "center", gap: 3, marginTop: 4, paddingLeft: 23 }}>
            {[1, 2, 3, 4, 5].map((c) => (
              <div key={c} style={{ width: CELL_W, textAlign: "center", fontSize: "0.82rem", fontWeight: 700, color: "var(--bfc-base-c-1)" }}>
                {c}
              </div>
            ))}
          </div>
          <div style={{ textAlign: "center", fontSize: "0.78rem", fontWeight: 600, color: "var(--bfc-base-c-2)", marginTop: 6, paddingLeft: 23, textTransform: "uppercase", letterSpacing: "0.04em" }}>
            Konsekvens
          </div>
        </div>
      </div>

      {/* Legend */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: "1rem", marginTop: "1.5rem", alignItems: "center" }}>
        <span style={{ fontSize: "0.8rem", fontWeight: 600, color: "var(--bfc-base-c-2)" }}>Forklaring:</span>
        {([
          { level: "low", label: "Lav (1–6)" },
          { level: "medium", label: "Middels (7–14)" },
          { level: "high", label: "Høy (15–25)" },
        ] as const).map(({ level, label }) => (
          <div key={level} style={{ display: "flex", alignItems: "center", gap: "0.4rem" }}>
            <div style={{ width: 18, height: 14, borderRadius: 3, background: CELL_BG[level], border: `1px solid ${CELL_BORDER[level]}` }} />
            <span style={{ fontSize: "0.78rem", color: "var(--bfc-base-c-2)" }}>{label}</span>
          </div>
        ))}
        <div style={{ display: "flex", alignItems: "center", gap: "0.4rem" }}>
          <div style={{ width: 18, height: 14, borderRadius: 3, border: "2px dashed #888", background: "transparent" }} />
          <span style={{ fontSize: "0.78rem", color: "var(--bfc-base-c-2)" }}>Restrisiko</span>
        </div>
      </div>
    </div>
  );
}

function RiskChip({ label, color, onClick, title, residual = false }: {
  label: string;
  color: string;
  onClick: () => void;
  title: string;
  residual?: boolean;
}) {
  const [hovered, setHovered] = useState(false);
  return (
    <div
      onClick={onClick}
      title={title}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        background: residual ? "transparent" : color,
        border: `2px ${residual ? "dashed" : "solid"} ${color}`,
        color: residual ? color : "#fff",
        borderRadius: 4, padding: "1px 5px",
        fontSize: "0.68rem", fontWeight: 600,
        cursor: "pointer", maxWidth: CELL_W - 14,
        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
        boxShadow: hovered ? `0 0 0 2px ${color}40` : "none",
        transition: "box-shadow 0.12s",
        userSelect: "none",
      }}
    >
      {label.length > 14 ? label.slice(0, 14) + "…" : label}
    </div>
  );
}
