import { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type Project, type CommunicationPlan, type CommunicationEntry, type Template } from "../api/client";
import { exportCommPlanPdf, exportCommPlanExcel } from "../utils/exportUtils";

const emptyEntry = {
  stakeholder: "",
  message: "",
  channel: "",
  frequency: "",
  responsible: "",
};

export default function CommunicationPlanPage() {
  const { projectId, planId } = useParams<{ projectId: string; planId: string }>();
  const navigate = useNavigate();
  const [plan, setPlan] = useState<CommunicationPlan | null>(null);
  const [project, setProject] = useState<Project | null>(null);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState(emptyEntry);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  const [templateModal, setTemplateModal] = useState(false);
  const [templateMode, setTemplateMode] = useState<"new" | "existing">("new");
  const [templateName, setTemplateName] = useState("");
  const [existingTemplates, setExistingTemplates] = useState<Template[]>([]);
  const [selectedExistingId, setSelectedExistingId] = useState("");
  const [templateSaving, setTemplateSaving] = useState(false);

  useEffect(() => {
    if (projectId && planId) {
      api.communicationPlans.get(projectId, planId).then(setPlan);
      api.projects.get(projectId).then(setProject);
    }
  }, [projectId, planId]);

  function openEdit(entry: CommunicationEntry) {
    setForm({ stakeholder: entry.stakeholder, message: entry.message, channel: entry.channel, frequency: entry.frequency, responsible: entry.responsible });
    setEditingId(entry.id);
    setShowForm(true);
  }

  async function handleSave() {
    if (!projectId || !planId || !form.stakeholder) return;
    setSaving(true);
    try {
      if (editingId) {
        const updated = await api.communicationPlans.updateEntry(projectId, planId, editingId, form);
        setPlan((prev) => prev ? { ...prev, entries: prev.entries.map((e) => e.id === editingId ? updated : e) } : prev);
      } else {
        const created = await api.communicationPlans.addEntry(projectId, planId, form);
        setPlan((prev) => prev ? { ...prev, entries: [...prev.entries, created] } : prev);
      }
      setShowForm(false);
      setForm(emptyEntry);
      setEditingId(null);
    } finally {
      setSaving(false);
    }
  }

  async function openTemplateModal() {
    const ts = await api.templates.list("communication_plan");
    setExistingTemplates(ts);
    setTemplateMode("new");
    setTemplateName(plan?.title ?? "");
    setSelectedExistingId(ts[0]?.id ?? "");
    setTemplateModal(true);
  }

  async function saveAsTemplate() {
    if (!plan) return;
    setTemplateSaving(true);
    try {
      const data = JSON.stringify({ entries: plan.entries.map(({ stakeholder, message, channel, frequency, responsible }) => ({ stakeholder, message, channel, frequency, responsible })) });
      if (templateMode === "new") {
        await api.templates.create({ name: templateName.trim() || plan.title, type: "communication_plan", data });
      } else {
        await api.templates.update(selectedExistingId, { name: existingTemplates.find(t => t.id === selectedExistingId)?.name ?? templateName, data });
      }
      setTemplateModal(false);
    } finally {
      setTemplateSaving(false);
    }
  }

  async function handleDelete(entryId: string) {
    if (!projectId || !planId) return;
    await api.communicationPlans.deleteEntry(projectId, planId, entryId);
    setPlan((prev) => prev ? { ...prev, entries: prev.entries.filter((e) => e.id !== entryId) } : prev);
  }

  if (!plan) return <div style={{ padding: "2rem" }}>Laster...</div>;

  return (
    <div style={{ maxWidth: 1000, margin: "0 auto", padding: "2rem" }}>
      <button
        onClick={() => navigate(`/projects/${projectId}`)}
        style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", marginBottom: "1rem" }}
      >
        ← Tilbake til prosjekt
      </button>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
        <h1 className="bf-h2">{plan.title}</h1>
        <div style={{ display: "flex", gap: "0.5rem" }}>
          {project && (<>
            <Button variant="outline" onClick={() => void exportCommPlanPdf(plan, project)}>↓ PDF</Button>
            <Button variant="outline" onClick={() => exportCommPlanExcel(plan, project)}>↓ Excel</Button>
          </>)}
          <Button variant="outline" onClick={openTemplateModal}>Lagre som mal</Button>
          <Button variant="filled" onClick={() => { setShowForm(true); setEditingId(null); setForm(emptyEntry); }}>+ Legg til oppføring</Button>
        </div>
      </div>

      <Modal isOpen={templateModal} onRequestClose={() => setTemplateModal(false)} header="Lagre som mal">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.5rem" }}>
            {(["new", "existing"] as const).map((mode) => (
              <button key={mode} onClick={() => setTemplateMode(mode)}
                style={{ padding: "0.75rem", borderRadius: 8, border: `2px solid ${templateMode === mode ? "#1971C2" : "var(--bfc-base-dimmed)"}`, background: templateMode === mode ? "#1971C218" : "var(--bfc-base-3)", cursor: "pointer", fontWeight: 600, fontSize: "0.9rem", color: templateMode === mode ? "#1971C2" : "var(--bfc-base-c-1)", transition: "all 0.15s" }}>
                {mode === "new" ? "Ny mal" : "Oppdater eksisterende"}
              </button>
            ))}
          </div>
          {templateMode === "new" ? (
            <Input label="Navn på malen" value={templateName} onChange={(e) => setTemplateName(e.target.value)} placeholder="f.eks. Standard kommunikasjonsplan" autoFocus />
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
          <p style={{ margin: 0, color: "var(--bfc-base-c-2)", fontSize: "0.85rem" }}>{plan.entries.length} oppføringer vil bli lagret i malen.</p>
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
          <h3 className="bf-h4">{editingId ? "Rediger oppføring" : "Ny oppføring"}</h3>
          <Input label="Interessent" value={form.stakeholder} onChange={(e) => setForm({ ...form, stakeholder: e.target.value })} />
          <Input label="Budskap / innhold" value={form.message} onChange={(e) => setForm({ ...form, message: e.target.value })} />
          <Input label="Kanal (f.eks. Teams, e-post, møte)" value={form.channel} onChange={(e) => setForm({ ...form, channel: e.target.value })} />
          <Input label="Frekvens (f.eks. Ukentlig, Månedlig)" value={form.frequency} onChange={(e) => setForm({ ...form, frequency: e.target.value })} />
          <Input label="Ansvarlig" value={form.responsible} onChange={(e) => setForm({ ...form, responsible: e.target.value })} />
          <div style={{ display: "flex", gap: "0.5rem" }}>
            <Button onClick={() => setShowForm(false)}>Avbryt</Button>
            <Button variant="filled" onClick={handleSave} state={saving || !form.stakeholder ? "inactive" : "default"}>
              {saving ? "Lagrer..." : "Lagre"}
            </Button>
          </div>
        </div>
      )}

      {plan.entries.length === 0 ? (
        <p style={{ color: "var(--bfc-base-c-2)" }}>Ingen oppføringer ennå.</p>
      ) : (
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ textAlign: "left", borderBottom: "2px solid var(--bfc-base-dimmed)" }}>
              <th style={{ padding: "0.5rem" }}>Interessent</th>
              <th style={{ padding: "0.5rem" }}>Budskap</th>
              <th style={{ padding: "0.5rem" }}>Kanal</th>
              <th style={{ padding: "0.5rem" }}>Frekvens</th>
              <th style={{ padding: "0.5rem" }}>Ansvarlig</th>
              <th style={{ padding: "0.5rem" }}></th>
            </tr>
          </thead>
          <tbody>
            {plan.entries.map((entry) => (
              <tr key={entry.id} style={{ borderBottom: "1px solid var(--bfc-base-dimmed)" }}>
                <td style={{ padding: "0.5rem" }}>{entry.stakeholder}</td>
                <td style={{ padding: "0.5rem" }}>{entry.message}</td>
                <td style={{ padding: "0.5rem" }}>{entry.channel}</td>
                <td style={{ padding: "0.5rem" }}>{entry.frequency}</td>
                <td style={{ padding: "0.5rem" }}>{entry.responsible}</td>
                <td style={{ padding: "0.5rem", display: "flex", gap: "0.25rem" }}>
                  <Button onClick={() => openEdit(entry)}>Rediger</Button>
                  <Button state="alert" onClick={() => handleDelete(entry.id)}>Slett</Button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}
