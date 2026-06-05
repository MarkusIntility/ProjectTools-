import { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal, TextArea } from "@intility/bifrost-react";
import { api, type Meeting, type MeetingPlan, type Template } from "../api/client";
import { isMsalConfigured, msalInstance } from "../auth/msalConfig";
import { fetchOutlookMeetingsByCategory, type OutlookEvent } from "../auth/calendarService";

const PLAN_COLOR = "#2F9E44";

function formatDate(iso: string): string {
  return new Date(iso).toLocaleString("nb-NO", {
    weekday: "short", day: "numeric", month: "short",
    year: "numeric", hour: "2-digit", minute: "2-digit",
  });
}

function relativeBadge(iso: string): { label: string; color: string } | null {
  const d = new Date(iso);
  const now = new Date();
  const diff = d.getTime() - now.getTime();
  if (diff < 0) return null;
  const days = Math.floor(diff / 86400000);
  if (days === 0) return { label: "I dag", color: "#E03131" };
  if (days === 1) return { label: "I morgen", color: "#F76707" };
  if (days <= 7) return { label: `Om ${days} dager`, color: "#F59F00" };
  return null;
}

export default function MeetingPlanPage() {
  const { projectId, planId } = useParams<{ projectId: string; planId: string }>();
  const navigate = useNavigate();

  const [plan, setPlan] = useState<MeetingPlan | null>(null);

  // ── Meeting form modal ───────────────────────────────────────────────────────
  const [showMeetingModal, setShowMeetingModal] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [form, setForm] = useState({ title: "", date: "", purpose: "" });
  const [saving, setSaving] = useState(false);

  // ── Outlook import modal ─────────────────────────────────────────────────────
  const [showImport, setShowImport] = useState(false);
  const [importCategory, setImportCategory] = useState("");
  const [importResults, setImportResults] = useState<OutlookEvent[] | null>(null);
  const [importFetching, setImportFetching] = useState(false);
  const [importSelected, setImportSelected] = useState<Set<string>>(new Set());
  const [importing, setImporting] = useState(false);
  const [importError, setImportError] = useState<string | null>(null);

  // ── Template modal ───────────────────────────────────────────────────────────
  const [templateModal, setTemplateModal] = useState(false);
  const [templateMode, setTemplateMode] = useState<"new" | "existing">("new");
  const [templateName, setTemplateName] = useState("");
  const [existingTemplates, setExistingTemplates] = useState<Template[]>([]);
  const [selectedExistingId, setSelectedExistingId] = useState("");
  const [templateSaving, setTemplateSaving] = useState(false);

  const msalAccount = isMsalConfigured ? (msalInstance.getAllAccounts()[0] ?? null) : null;

  useEffect(() => {
    if (projectId && planId) {
      api.meetingPlans.get(projectId, planId).then(setPlan);
    }
  }, [projectId, planId]);

  // ── Meeting CRUD ─────────────────────────────────────────────────────────────

  function openNew() {
    setForm({ title: "", date: "", purpose: "" });
    setEditingId(null);
    setShowMeetingModal(true);
  }

  function openEdit(m: Meeting) {
    setForm({ title: m.title, date: m.date.slice(0, 16), purpose: m.purpose ?? "" });
    setEditingId(m.id);
    setShowMeetingModal(true);
  }

  async function handleSave() {
    if (!projectId || !planId || !form.title || !form.date) return;
    setSaving(true);
    try {
      const payload = {
        title: form.title,
        date: new Date(form.date).toISOString(),
        purpose: form.purpose.trim() || null,
        outlook_id: null,
      };
      if (editingId) {
        const updated = await api.meetingPlans.updateMeeting(projectId, planId, editingId, payload);
        setPlan((p) => p ? { ...p, meetings: p.meetings.map((m) => m.id === editingId ? updated : m) } : p);
      } else {
        const created = await api.meetingPlans.addMeeting(projectId, planId, payload);
        setPlan((p) => p ? { ...p, meetings: [...p.meetings, created] } : p);
      }
      setShowMeetingModal(false);
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete(meetingId: string) {
    if (!projectId || !planId) return;
    await api.meetingPlans.deleteMeeting(projectId, planId, meetingId);
    setPlan((p) => p ? { ...p, meetings: p.meetings.filter((m) => m.id !== meetingId) } : p);
  }

  // ── Outlook import ───────────────────────────────────────────────────────────

  function openImport() {
    setImportCategory("");
    setImportResults(null);
    setImportSelected(new Set());
    setImportError(null);
    setShowImport(true);
  }

  async function handleLoginForImport() {
    sessionStorage.setItem("app.returnUrl", window.location.href);
    await msalInstance.loginRedirect({ scopes: ["Calendars.Read"] });
  }

  async function fetchImportResults() {
    if (!msalAccount || !importCategory.trim()) return;
    setImportFetching(true);
    setImportError(null);
    setImportResults(null);
    setImportSelected(new Set());
    try {
      const events = await fetchOutlookMeetingsByCategory(msalInstance, msalAccount, importCategory.trim());
      setImportResults(events);
    } catch (e) {
      setImportError(e instanceof Error ? e.message : "Ukjent feil");
    } finally {
      setImportFetching(false);
    }
  }

  function toggleImportSelect(id: string) {
    setImportSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  }

  function isAlreadyImported(eventId: string): boolean {
    return plan?.meetings.some((m) => m.outlook_id === eventId) ?? false;
  }

  async function handleImport() {
    if (!projectId || !planId || !importResults || importSelected.size === 0) return;
    setImporting(true);
    try {
      const toImport = importResults.filter((e) => importSelected.has(e.id));
      const created: Meeting[] = [];
      for (const event of toImport) {
        const m = await api.meetingPlans.addMeeting(projectId, planId, {
          title: event.subject,
          date: new Date(event.start.dateTime).toISOString(),
          purpose: event.bodyPreview?.trim() || null,
          outlook_id: event.id,
        });
        created.push(m);
      }
      setPlan((p) => p ? { ...p, meetings: [...p.meetings, ...created] } : p);
      setShowImport(false);
    } finally {
      setImporting(false);
    }
  }

  // ── Template ─────────────────────────────────────────────────────────────────

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
      const data = JSON.stringify({
        meetings: plan.meetings.map(({ title, date, purpose }) => ({ title, date, purpose })),
      });
      if (templateMode === "new") {
        await api.templates.create({ name: templateName.trim() || plan.title, type: "meeting_plan", data });
      } else {
        const existing = existingTemplates.find((t) => t.id === selectedExistingId);
        await api.templates.update(selectedExistingId, { name: existing?.name ?? templateName, data });
      }
      setTemplateModal(false);
    } finally {
      setTemplateSaving(false);
    }
  }

  // ── Render ───────────────────────────────────────────────────────────────────

  if (!plan) return <div style={{ padding: "2rem" }}>Laster...</div>;

  const sorted = [...plan.meetings].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
  const selectableCount = importResults?.filter((e) => !isAlreadyImported(e.id)).length ?? 0;
  const allSelected = selectableCount > 0 && importSelected.size === selectableCount;

  return (
    <div style={{ maxWidth: 900, margin: "0 auto", padding: "2rem" }}>
      <button
        onClick={() => navigate(`/projects/${projectId}`)}
        style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", marginBottom: "1rem", fontSize: "0.9rem", padding: 0 }}
      >
        ← Tilbake til prosjekt
      </button>

      {/* ── Header ── */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "2rem" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}>
          <div style={{ width: 4, height: 28, borderRadius: 2, background: PLAN_COLOR, flexShrink: 0 }} />
          <h1 className="bf-h2" style={{ margin: 0 }}>{plan.title}</h1>
          {sorted.length > 0 && (
            <span style={{ fontSize: "0.78rem", fontWeight: 600, padding: "2px 10px", borderRadius: 20, background: `${PLAN_COLOR}18`, color: PLAN_COLOR }}>
              {sorted.length} møter
            </span>
          )}
        </div>
        <div style={{ display: "flex", gap: "0.5rem" }}>
          <Button variant="outline" onClick={openTemplateModal}>Lagre som mal</Button>
          {isMsalConfigured && (
            <Button variant="outline" onClick={openImport}>Importer fra Outlook</Button>
          )}
          <Button variant="filled" onClick={openNew} style={{ background: PLAN_COLOR, borderColor: PLAN_COLOR }}>+ Legg til møte</Button>
        </div>
      </div>

      {/* ── Meeting list ── */}
      {sorted.length === 0 ? (
        <div style={{ padding: "2rem", textAlign: "center", borderRadius: 10, border: `1px dashed ${PLAN_COLOR}60`, background: `${PLAN_COLOR}06`, color: "var(--bfc-base-c-2)" }}>
          Ingen møter lagt til ennå. Klikk «+ Legg til møte» eller «Importer fra Outlook».
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: "0.6rem" }}>
          {sorted.map((meeting) => (
            <MeetingCard
              key={meeting.id}
              meeting={meeting}
              onEdit={() => openEdit(meeting)}
              onDelete={() => handleDelete(meeting.id)}
            />
          ))}
        </div>
      )}

      {/* ── Meeting form modal ── */}
      <Modal
        isOpen={showMeetingModal}
        onRequestClose={() => setShowMeetingModal(false)}
        header={editingId ? "Rediger møte" : "Nytt møte"}
      >
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <Input
            label="Møtenavn"
            value={form.title}
            onChange={(e) => setForm({ ...form, title: e.target.value })}
            placeholder="f.eks. Statusmøte uke 23"
            autoFocus
          />
          <Input
            label="Dato og tidspunkt"
            type="datetime-local"
            value={form.date}
            onChange={(e) => setForm({ ...form, date: e.target.value })}
          />
          <TextArea
            label="Hensikt"
            value={form.purpose}
            onChange={(e) => setForm({ ...form, purpose: e.target.value })}
            placeholder="Hva er formålet med møtet?"
          />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setShowMeetingModal(false)}>Avbryt</Button>
            <Button
              variant="filled"
              onClick={handleSave}
              state={saving || !form.title || !form.date ? "inactive" : "default"}
              style={!saving && form.title && form.date ? { background: PLAN_COLOR, borderColor: PLAN_COLOR } : undefined}
            >
              {saving ? "Lagrer..." : "Lagre"}
            </Button>
          </div>
        </div>
      </Modal>

      {/* ── Outlook import modal ── */}
      <Modal isOpen={showImport} onRequestClose={() => setShowImport(false)} header="Importer møter fra Outlook">
        {!isMsalConfigured ? (
          <p style={{ color: "var(--bfc-base-c-2)", fontSize: "0.9rem" }}>Azure AD er ikke konfigurert.</p>
        ) : !msalAccount ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem", alignItems: "flex-start" }}>
            <p style={{ margin: 0, color: "var(--bfc-base-c-2)", fontSize: "0.9rem" }}>
              Logg inn med Microsoft-kontoen din for å hente møter fra Outlook-kalenderen.
            </p>
            <Button variant="filled" onClick={handleLoginForImport} style={{ background: "#0078D4", borderColor: "#0078D4" }}>
              Logg inn med Microsoft
            </Button>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <p style={{ margin: 0, fontSize: "0.85rem", color: "var(--bfc-base-c-2)" }}>
              Innlogget som <strong>{msalAccount.username}</strong>
            </p>

            <div style={{ display: "flex", gap: "0.5rem", alignItems: "flex-end" }}>
              <div style={{ flex: 1 }}>
                <Input
                  label="Kategori i Outlook"
                  value={importCategory}
                  onChange={(e) => setImportCategory(e.target.value)}
                  placeholder="f.eks. Lunera"
                  onKeyDown={(e) => e.key === "Enter" && fetchImportResults()}
                />
              </div>
              <Button
                variant="filled"
                onClick={fetchImportResults}
                state={importFetching || !importCategory.trim() ? "inactive" : "default"}
                style={{ marginBottom: 2, background: "#0078D4", borderColor: "#0078D4" }}
              >
                {importFetching ? "Henter..." : "Søk"}
              </Button>
            </div>

            {importError && (
              <div style={{ padding: "0.75rem 1rem", borderRadius: 7, background: "#FFE3E3", border: "1px solid #E0313140", color: "#E03131", fontSize: "0.85rem" }}>
                {importError}
              </div>
            )}

            {importResults !== null && (
              <>
                {importResults.length === 0 ? (
                  <p style={{ margin: 0, color: "var(--bfc-base-c-2)", fontSize: "0.9rem" }}>
                    Ingen fremtidige møter funnet med kategorien «{importCategory}».
                  </p>
                ) : (
                  <>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: "0.85rem", color: "var(--bfc-base-c-2)" }}>
                        {importResults.length} møter funnet — {importSelected.size} valgt
                      </span>
                      {selectableCount > 0 && (
                        <button
                          onClick={() => {
                            if (allSelected) {
                              setImportSelected(new Set());
                            } else {
                              setImportSelected(new Set(importResults.filter((e) => !isAlreadyImported(e.id)).map((e) => e.id)));
                            }
                          }}
                          style={{ background: "none", border: "none", cursor: "pointer", color: "#0078D4", fontSize: "0.82rem", fontWeight: 600, padding: 0 }}
                        >
                          {allSelected ? "Fjern alle" : "Velg alle"}
                        </button>
                      )}
                    </div>

                    <div style={{ display: "flex", flexDirection: "column", gap: "0.4rem", maxHeight: 320, overflowY: "auto" }}>
                      {importResults.map((event) => {
                        const alreadyDone = isAlreadyImported(event.id);
                        const checked = importSelected.has(event.id);
                        return (
                          <label
                            key={event.id}
                            style={{
                              display: "flex", alignItems: "flex-start", gap: "0.75rem",
                              padding: "0.75rem 1rem", borderRadius: 7,
                              border: `1px solid ${checked ? `${PLAN_COLOR}60` : "var(--bfc-base-dimmed)"}`,
                              background: checked ? `${PLAN_COLOR}08` : "var(--bfc-base-3)",
                              cursor: alreadyDone ? "default" : "pointer",
                              opacity: alreadyDone ? 0.55 : 1,
                              transition: "border-color 0.15s, background 0.15s",
                            }}
                          >
                            <input
                              type="checkbox"
                              checked={checked}
                              disabled={alreadyDone}
                              onChange={() => !alreadyDone && toggleImportSelect(event.id)}
                              style={{ marginTop: 2, flexShrink: 0, cursor: alreadyDone ? "default" : "pointer" }}
                            />
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", flexWrap: "wrap" }}>
                                <span style={{ fontWeight: 600, fontSize: "0.9rem" }}>{event.subject}</span>
                                {alreadyDone && (
                                  <span style={{ fontSize: "0.7rem", fontWeight: 600, padding: "1px 7px", borderRadius: 20, background: "#868E9618", color: "#868E96" }}>
                                    Allerede importert
                                  </span>
                                )}
                              </div>
                              <div style={{ fontSize: "0.78rem", color: "var(--bfc-base-c-2)", marginTop: 2 }}>
                                {new Date(event.start.dateTime).toLocaleString("nb-NO", { weekday: "short", day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" })}
                              </div>
                              {event.bodyPreview?.trim() && (
                                <div style={{ fontSize: "0.78rem", color: "var(--bfc-base-c-3)", marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                                  {event.bodyPreview}
                                </div>
                              )}
                            </div>
                          </label>
                        );
                      })}
                    </div>

                    <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", paddingTop: "0.25rem" }}>
                      <Button onClick={() => setShowImport(false)}>Avbryt</Button>
                      <Button
                        variant="filled"
                        onClick={handleImport}
                        state={importing || importSelected.size === 0 ? "inactive" : "default"}
                        style={!importing && importSelected.size > 0 ? { background: PLAN_COLOR, borderColor: PLAN_COLOR } : undefined}
                      >
                        {importing ? "Importerer..." : `Importer ${importSelected.size > 0 ? importSelected.size + " møter" : ""}`}
                      </Button>
                    </div>
                  </>
                )}
              </>
            )}
          </div>
        )}
      </Modal>

      {/* ── Template modal ── */}
      <Modal isOpen={templateModal} onRequestClose={() => setTemplateModal(false)} header="Lagre som mal">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.5rem" }}>
            {(["new", "existing"] as const).map((mode) => (
              <button
                key={mode}
                onClick={() => setTemplateMode(mode)}
                style={{ padding: "0.75rem", borderRadius: 8, border: `2px solid ${templateMode === mode ? PLAN_COLOR : "var(--bfc-base-dimmed)"}`, background: templateMode === mode ? `${PLAN_COLOR}18` : "var(--bfc-base-3)", cursor: "pointer", fontWeight: 600, fontSize: "0.9rem", color: templateMode === mode ? PLAN_COLOR : "var(--bfc-base-c-1)", transition: "all 0.15s" }}
              >
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
    </div>
  );
}

// ── Meeting card ──────────────────────────────────────────────────────────────

function MeetingCard({ meeting, onEdit, onDelete }: {
  meeting: Meeting;
  onEdit: () => void;
  onDelete: () => void;
}) {
  const [hovered, setHovered] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const badge = relativeBadge(meeting.date);
  const isPast = new Date(meeting.date) < new Date();

  return (
    <div
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => { setHovered(false); setConfirmDelete(false); }}
      style={{
        display: "flex", alignItems: "flex-start", gap: "1rem",
        padding: "1rem 1.25rem", borderRadius: 8,
        background: "var(--bfc-base-3)",
        border: "1px solid var(--bfc-base-dimmed)",
        borderLeft: `3px solid ${hovered ? PLAN_COLOR : (isPast ? "var(--bfc-base-dimmed)" : PLAN_COLOR + "60")}`,
        opacity: isPast ? 0.65 : 1,
        transition: "border-color 0.15s, box-shadow 0.15s",
        boxShadow: hovered ? "0 2px 10px rgba(0,0,0,0.07)" : "none",
      }}
    >
      {/* Date block */}
      <div style={{ flexShrink: 0, minWidth: 52, textAlign: "center", paddingTop: 2 }}>
        <div style={{ fontSize: "1.4rem", fontWeight: 800, lineHeight: 1, color: isPast ? "var(--bfc-base-c-3)" : PLAN_COLOR }}>
          {new Date(meeting.date).getDate()}
        </div>
        <div style={{ fontSize: "0.72rem", fontWeight: 600, color: "var(--bfc-base-c-2)", textTransform: "uppercase", letterSpacing: "0.04em" }}>
          {new Date(meeting.date).toLocaleDateString("nb-NO", { month: "short" })}
        </div>
        <div style={{ fontSize: "0.72rem", color: "var(--bfc-base-c-3)" }}>
          {new Date(meeting.date).toLocaleTimeString("nb-NO", { hour: "2-digit", minute: "2-digit" })}
        </div>
      </div>

      {/* Divider */}
      <div style={{ width: 1, alignSelf: "stretch", background: "var(--bfc-base-dimmed)", flexShrink: 0 }} />

      {/* Content */}
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", flexWrap: "wrap" }}>
          <span style={{ fontWeight: 600, fontSize: "0.95rem" }}>{meeting.title}</span>
          {badge && (
            <span style={{ fontSize: "0.7rem", fontWeight: 600, padding: "1px 8px", borderRadius: 20, background: `${badge.color}18`, color: badge.color }}>
              {badge.label}
            </span>
          )}
          {meeting.outlook_id && (
            <span style={{ fontSize: "0.7rem", fontWeight: 600, padding: "1px 8px", borderRadius: 20, background: "#0078D418", color: "#0078D4" }}>
              Outlook
            </span>
          )}
        </div>
        {meeting.purpose && (
          <p style={{ margin: "0.3rem 0 0", fontSize: "0.85rem", color: "var(--bfc-base-c-2)", lineHeight: 1.45 }}>
            {meeting.purpose}
          </p>
        )}
        <div style={{ fontSize: "0.78rem", color: "var(--bfc-base-c-3)", marginTop: "0.25rem" }}>
          {formatDate(meeting.date)}
        </div>
      </div>

      {/* Actions */}
      <div style={{ display: "flex", alignItems: "center", gap: "0.4rem", flexShrink: 0, opacity: hovered ? 1 : 0, transition: "opacity 0.15s" }}>
        {confirmDelete ? (
          <>
            <span style={{ fontSize: "0.8rem", color: "#E03131", fontWeight: 500 }}>Slett?</span>
            <button
              onClick={onDelete}
              style={{ background: "#E03131", border: "none", borderRadius: 4, color: "#fff", padding: "3px 10px", fontSize: "0.8rem", fontWeight: 600, cursor: "pointer" }}
            >
              Ja
            </button>
            <button
              onClick={() => setConfirmDelete(false)}
              style={{ background: "none", border: "1px solid var(--bfc-base-dimmed)", borderRadius: 4, color: "var(--bfc-base-c-2)", padding: "3px 10px", fontSize: "0.8rem", cursor: "pointer" }}
            >
              Nei
            </button>
          </>
        ) : (
          <>
            <button
              onClick={onEdit}
              style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "4px 8px", borderRadius: 4, fontSize: "0.8rem", fontWeight: 500 }}
              onMouseEnter={(e) => (e.currentTarget.style.color = PLAN_COLOR)}
              onMouseLeave={(e) => (e.currentTarget.style.color = "var(--bfc-base-c-2)")}
            >
              Endre
            </button>
            <div style={{ width: 1, height: 14, background: "var(--bfc-base-dimmed)" }} />
            <button
              onClick={() => setConfirmDelete(true)}
              style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "4px 8px", borderRadius: 4, fontSize: "0.8rem", fontWeight: 500 }}
              onMouseEnter={(e) => (e.currentTarget.style.color = "#E03131")}
              onMouseLeave={(e) => (e.currentTarget.style.color = "var(--bfc-base-c-2)")}
            >
              Slett
            </button>
          </>
        )}
      </div>
    </div>
  );
}
