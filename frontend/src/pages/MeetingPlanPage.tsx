import { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal, TextArea } from "@intility/bifrost-react";
import { api, type MeetingPlan, type Meeting, type Template } from "../api/client";

const emptyMeeting = {
  title: "",
  date: "",
  location: "",
  agenda: "",
  participants: "",
  minutes: "",
};

export default function MeetingPlanPage() {
  const { projectId, planId } = useParams<{ projectId: string; planId: string }>();
  const navigate = useNavigate();
  const [plan, setPlan] = useState<MeetingPlan | null>(null);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState(emptyMeeting);
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
      api.meetingPlans.get(projectId, planId).then(setPlan);
    }
  }, [projectId, planId]);

  function openEdit(meeting: Meeting) {
    setForm({
      title: meeting.title,
      date: meeting.date.slice(0, 16),
      location: meeting.location ?? "",
      agenda: meeting.agenda ?? "",
      participants: meeting.participants ?? "",
      minutes: meeting.minutes ?? "",
    });
    setEditingId(meeting.id);
    setShowForm(true);
  }

  async function handleSave() {
    if (!projectId || !planId || !form.title || !form.date) return;
    setSaving(true);
    try {
      const data = {
        ...form,
        date: new Date(form.date).toISOString(),
        location: form.location || null,
        agenda: form.agenda || null,
        participants: form.participants || null,
        minutes: form.minutes || null,
      };
      if (editingId) {
        const updated = await api.meetingPlans.updateMeeting(projectId, planId, editingId, data);
        setPlan((prev) => prev ? { ...prev, meetings: prev.meetings.map((m) => m.id === editingId ? updated : m) } : prev);
      } else {
        const created = await api.meetingPlans.addMeeting(projectId, planId, data);
        setPlan((prev) => prev ? { ...prev, meetings: [...prev.meetings, created] } : prev);
      }
      setShowForm(false);
      setForm(emptyMeeting);
      setEditingId(null);
    } finally {
      setSaving(false);
    }
  }

  async function openTemplateModal() {
    const ts = await api.templates.list("meeting_plan");
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
      const data = JSON.stringify({ meetings: plan.meetings.map(({ title, date, location, agenda, participants, minutes }) => ({ title, date, location, agenda, participants, minutes })) });
      if (templateMode === "new") {
        await api.templates.create({ name: templateName.trim() || plan.title, type: "meeting_plan", data });
      } else {
        await api.templates.update(selectedExistingId, { name: existingTemplates.find(t => t.id === selectedExistingId)?.name ?? templateName, data });
      }
      setTemplateModal(false);
    } finally {
      setTemplateSaving(false);
    }
  }

  async function handleDelete(meetingId: string) {
    if (!projectId || !planId) return;
    await api.meetingPlans.deleteMeeting(projectId, planId, meetingId);
    setPlan((prev) => prev ? { ...prev, meetings: prev.meetings.filter((m) => m.id !== meetingId) } : prev);
  }

  if (!plan) return <div style={{ padding: "2rem" }}>Laster...</div>;

  const sorted = [...(plan.meetings ?? [])].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());

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
          <Button variant="outline" onClick={openTemplateModal}>Lagre som mal</Button>
          <Button variant="filled" onClick={() => { setShowForm(true); setEditingId(null); setForm(emptyMeeting); }}>+ Legg til møte</Button>
        </div>
      </div>

      <Modal isOpen={templateModal} onRequestClose={() => setTemplateModal(false)} header="Lagre som mal">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.5rem" }}>
            {(["new", "existing"] as const).map((mode) => (
              <button key={mode} onClick={() => setTemplateMode(mode)}
                style={{ padding: "0.75rem", borderRadius: 8, border: `2px solid ${templateMode === mode ? "#2F9E44" : "var(--bfc-base-dimmed)"}`, background: templateMode === mode ? "#2F9E4418" : "var(--bfc-base-3)", cursor: "pointer", fontWeight: 600, fontSize: "0.9rem", color: templateMode === mode ? "#2F9E44" : "var(--bfc-base-c-1)", transition: "all 0.15s" }}>
                {mode === "new" ? "Ny mal" : "Oppdater eksisterende"}
              </button>
            ))}
          </div>
          {templateMode === "new" ? (
            <Input label="Navn på malen" value={templateName} onChange={(e) => setTemplateName(e.target.value)} placeholder="f.eks. Standard møteplan" autoFocus />
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
          <p style={{ margin: 0, color: "var(--bfc-base-c-2)", fontSize: "0.85rem" }}>{plan.meetings.length} møter vil bli lagret i malen.</p>
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
          <h3 className="bf-h4">{editingId ? "Rediger møte" : "Nytt møte"}</h3>
          <Input label="Tittel" value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} />
          <Input label="Dato og tid" type="datetime-local" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
          <Input label="Sted / lenke" value={form.location} onChange={(e) => setForm({ ...form, location: e.target.value })} />
          <Input label="Deltakere (kommaseparert)" value={form.participants} onChange={(e) => setForm({ ...form, participants: e.target.value })} />
          <TextArea label="Agenda" value={form.agenda} onChange={(e) => setForm({ ...form, agenda: e.target.value })} />
          <TextArea label="Referat" value={form.minutes} onChange={(e) => setForm({ ...form, minutes: e.target.value })} />
          <div style={{ display: "flex", gap: "0.5rem" }}>
            <Button onClick={() => setShowForm(false)}>Avbryt</Button>
            <Button variant="filled" onClick={handleSave} state={saving || !form.title || !form.date ? "inactive" : "default"}>
              {saving ? "Lagrer..." : "Lagre"}
            </Button>
          </div>
        </div>
      )}

      {sorted.length === 0 ? (
        <p style={{ color: "var(--bfc-base-c-2)" }}>Ingen møter lagt til ennå.</p>
      ) : (
        <div style={{ display: "grid", gap: "1rem" }}>
          {sorted.map((meeting) => (
            <div key={meeting.id} style={{ border: "1px solid var(--bfc-base-dimmed)", borderRadius: 8, padding: "1rem" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <h3 className="bf-h4">{meeting.title}</h3>
                  <p style={{ color: "var(--bfc-base-c-2)", fontSize: "0.9rem" }}>
                    {new Date(meeting.date).toLocaleString("nb-NO")}
                    {meeting.location && ` · ${meeting.location}`}
                  </p>
                  {meeting.participants && <p style={{ fontSize: "0.85rem", marginTop: "0.25rem" }}>Deltakere: {meeting.participants}</p>}
                  {meeting.agenda && <p style={{ fontSize: "0.85rem", marginTop: "0.5rem", whiteSpace: "pre-wrap" }}><strong>Agenda:</strong> {meeting.agenda}</p>}
                  {meeting.minutes && <p style={{ fontSize: "0.85rem", marginTop: "0.5rem", whiteSpace: "pre-wrap" }}><strong>Referat:</strong> {meeting.minutes}</p>}
                </div>
                <div style={{ display: "flex", gap: "0.25rem", flexShrink: 0, marginLeft: "1rem" }}>
                  <Button onClick={() => openEdit(meeting)}>Rediger</Button>
                  <Button state="alert" onClick={() => handleDelete(meeting.id)}>Slett</Button>
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
