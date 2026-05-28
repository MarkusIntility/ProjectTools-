import { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input } from "@intility/bifrost-react";
import { api, type RiskMatrix, type RiskItem } from "../api/client";

const RISK_COLORS: Record<string, string> = {
  low: "#4caf50",
  medium: "#ff9800",
  high: "#f44336",
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
}

const emptyRisk: RiskForm = {
  description: "",
  probability: 3,
  consequence: 3,
  mitigation: "",
  owner: "",
  status: "open",
};

export default function RiskMatrixPage() {
  const { projectId, matrixId } = useParams<{ projectId: string; matrixId: string }>();
  const navigate = useNavigate();
  const [matrix, setMatrix] = useState<RiskMatrix | null>(null);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState<RiskForm>(emptyRisk);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (projectId && matrixId) {
      api.riskMatrices.get(projectId, matrixId).then(setMatrix);
    }
  }, [projectId, matrixId]);

  function openEdit(risk: RiskItem) {
    setForm({
      description: risk.description,
      probability: risk.probability,
      consequence: risk.consequence,
      mitigation: risk.mitigation ?? "",
      owner: risk.owner ?? "",
      status: risk.status,
    });
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

  async function handleDelete(riskId: string) {
    if (!projectId || !matrixId) return;
    await api.riskMatrices.deleteRisk(projectId, matrixId, riskId);
    setMatrix((prev) => prev ? { ...prev, risks: prev.risks.filter((r) => r.id !== riskId) } : prev);
  }

  if (!matrix) return <div style={{ padding: "2rem" }}>Laster...</div>;

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: "2rem" }}>
      <button
        onClick={() => navigate(`/projects/${projectId}`)}
        style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", marginBottom: "1rem" }}
      >
        ← Tilbake til prosjekt
      </button>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
        <h1 className="bf-h2">{matrix.title}</h1>
        <Button variant="filled" onClick={() => { setShowForm(true); setEditingId(null); setForm(emptyRisk); }}>
          + Legg til risiko
        </Button>
      </div>

      {showForm && (
        <div style={{ background: "var(--bfc-base-3)", borderRadius: 8, padding: "1.5rem", marginBottom: "1.5rem", display: "grid", gap: "1rem" }}>
          <h3 className="bf-h4">{editingId ? "Rediger risiko" : "Ny risiko"}</h3>
          <Input label="Beskrivelse" value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "1rem" }}>
            <div>
              <label className="bf-label">Sannsynlighet (1–5): {form.probability}</label>
              <input type="range" min={1} max={5} value={form.probability}
                onChange={(e) => setForm({ ...form, probability: Number(e.target.value) })}
                style={{ width: "100%" }} />
            </div>
            <div>
              <label className="bf-label">Konsekvens (1–5): {form.consequence}</label>
              <input type="range" min={1} max={5} value={form.consequence}
                onChange={(e) => setForm({ ...form, consequence: Number(e.target.value) })}
                style={{ width: "100%" }} />
            </div>
          </div>
          <Input label="Tiltak" value={form.mitigation} onChange={(e) => setForm({ ...form, mitigation: e.target.value })} />
          <Input label="Ansvarlig" value={form.owner} onChange={(e) => setForm({ ...form, owner: e.target.value })} />
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

      {matrix.risks.length === 0 ? (
        <p style={{ color: "var(--bfc-base-c-2)" }}>Ingen risikoer lagt til ennå.</p>
      ) : (
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ textAlign: "left", borderBottom: "2px solid var(--bfc-base-dimmed)" }}>
              <th style={{ padding: "0.5rem" }}>Beskrivelse</th>
              <th style={{ padding: "0.5rem" }}>S</th>
              <th style={{ padding: "0.5rem" }}>K</th>
              <th style={{ padding: "0.5rem" }}>Score</th>
              <th style={{ padding: "0.5rem" }}>Tiltak</th>
              <th style={{ padding: "0.5rem" }}>Ansvarlig</th>
              <th style={{ padding: "0.5rem" }}>Status</th>
              <th style={{ padding: "0.5rem" }}></th>
            </tr>
          </thead>
          <tbody>
            {matrix.risks.map((risk) => {
              const level = riskLevel(risk.risk_score);
              return (
                <tr key={risk.id} style={{ borderBottom: "1px solid var(--bfc-base-dimmed)" }}>
                  <td style={{ padding: "0.5rem" }}>{risk.description}</td>
                  <td style={{ padding: "0.5rem" }}>{risk.probability}</td>
                  <td style={{ padding: "0.5rem" }}>{risk.consequence}</td>
                  <td style={{ padding: "0.5rem" }}>
                    <span style={{ background: RISK_COLORS[level], color: "#fff", borderRadius: 4, padding: "2px 8px", fontWeight: 600 }}>
                      {risk.risk_score}
                    </span>
                  </td>
                  <td style={{ padding: "0.5rem" }}>{risk.mitigation ?? "–"}</td>
                  <td style={{ padding: "0.5rem" }}>{risk.owner ?? "–"}</td>
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
      )}
    </div>
  );
}
