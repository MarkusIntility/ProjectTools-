import { useEffect, useState, useCallback } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type Runbook, type RunbookActivity } from "../api/client";
import { isMsalConfigured, msalInstance, PLANNER_SCOPES } from "../auth/msalConfig";
import { fetchPlannerData, parsePlanId, taskStatus, type PlannerData, type PlannerTask } from "../auth/plannerService";
import type { AccountInfo } from "@azure/msal-browser";

// ─── Status config ────────────────────────────────────────────────────────────

const STATUS_CONFIG: Record<RunbookActivity["status"], { label: string; color: string; bg: string }> = {
  not_started: { label: "Ikke startet", color: "#868E96", bg: "#868E9618" },
  in_progress:  { label: "Pågående",     color: "#1971C2", bg: "#1971C218" },
  done:         { label: "Ferdig",       color: "#2F9E44", bg: "#2F9E4418" },
  cancelled:    { label: "Kansellert",   color: "#E03131", bg: "#E0313118" },
};

const SOURCE_CONFIG = {
  planner:    { label: "Microsoft Planner", color: "#0078D4", initial: "P" },
  smartsheet: { label: "Smartsheet",        color: "#00A88E", initial: "S" },
  own:        { label: "Egen",              color: "#7950F2", initial: "E" },
};

// ─── Empty activity form ──────────────────────────────────────────────────────

const EMPTY_FORM = {
  name: "",
  phase: "",
  status: "not_started" as RunbookActivity["status"],
  start_date: "",
  end_date: "",
  responsible: "",
  description: "",
};

// ─── Main page ────────────────────────────────────────────────────────────────

export default function RunbookPage() {
  const { projectId, runbookId } = useParams<{ projectId: string; runbookId: string }>();
  const navigate = useNavigate();
  const [runbook, setRunbook] = useState<Runbook | null>(null);

  // Planner state
  const [msalReady, setMsalReady] = useState(false);
  const [plannerAccount, setPlannerAccount] = useState<AccountInfo | null>(null);
  const [plannerData, setPlannerData] = useState<PlannerData | null>(null);
  const [plannerLoading, setPlannerLoading] = useState(false);
  const [plannerError, setPlannerError] = useState<string | null>(null);

  const [activityModal, setActivityModal] = useState(false);
  const [editTarget, setEditTarget] = useState<RunbookActivity | null>(null);
  const [form, setForm] = useState(EMPTY_FORM);
  const [saving, setSaving] = useState(false);

  const [deleteTarget, setDeleteTarget] = useState<RunbookActivity | null>(null);
  const [deleting, setDeleting] = useState(false);

  const [editRunbookModal, setEditRunbookModal] = useState(false);
  const [runbookTitle, setRunbookTitle] = useState("");
  const [runbookUrl, setRunbookUrl] = useState("");

  useEffect(() => {
    if (!projectId || !runbookId) return;
    api.runbooks.get(projectId, runbookId).then(setRunbook);
  }, [projectId, runbookId]);

  // Initialize MSAL once and check for existing login
  useEffect(() => {
    if (!isMsalConfigured) return;
    msalInstance.initialize().then(() => {
      setMsalReady(true);
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) setPlannerAccount(accounts[0]);
    });
  }, []);

  const loadPlannerData = useCallback(async (account: AccountInfo, planId: string) => {
    setPlannerLoading(true);
    setPlannerError(null);
    try {
      const data = await fetchPlannerData(msalInstance, account, planId);
      setPlannerData(data);
    } catch (err) {
      setPlannerError(err instanceof Error ? err.message : "Ukjent feil");
    } finally {
      setPlannerLoading(false);
    }
  }, []);

  // Auto-fetch when account + runbook ready
  useEffect(() => {
    if (!plannerAccount || !runbook || runbook.source !== "planner" || !runbook.external_url) return;
    const planId = parsePlanId(runbook.external_url);
    if (planId) loadPlannerData(plannerAccount, planId);
  }, [plannerAccount, runbook, loadPlannerData]);

  async function loginPlanner() {
    if (!isMsalConfigured || !msalReady) return;
    try {
      const result = await msalInstance.loginPopup({ scopes: PLANNER_SCOPES });
      setPlannerAccount(result.account);
    } catch (err) {
      setPlannerError("Innlogging avbrutt eller mislyktes.");
    }
  }

  function refreshPlanner() {
    if (!plannerAccount || !runbook?.external_url) return;
    const planId = parsePlanId(runbook.external_url);
    if (planId) loadPlannerData(plannerAccount, planId);
  }

  function openAdd(defaultPhase?: string) {
    setEditTarget(null);
    setForm({ ...EMPTY_FORM, phase: defaultPhase ?? "" });
    setActivityModal(true);
  }

  function openEdit(activity: RunbookActivity) {
    setEditTarget(activity);
    setForm({
      name: activity.name,
      phase: activity.phase ?? "",
      status: activity.status,
      start_date: activity.start_date ? activity.start_date.slice(0, 16) : "",
      end_date: activity.end_date ? activity.end_date.slice(0, 16) : "",
      responsible: activity.responsible ?? "",
      description: activity.description ?? "",
    });
    setActivityModal(true);
  }

  async function handleSaveActivity() {
    if (!projectId || !runbookId || !form.name.trim()) return;
    setSaving(true);
    try {
      const payload = {
        name: form.name,
        phase: form.phase || null,
        status: form.status,
        start_date: form.start_date || null,
        end_date: form.end_date || null,
        responsible: form.responsible || null,
        description: form.description || null,
        sort_order: editTarget?.sort_order ?? 0,
      };
      if (editTarget) {
        const updated = await api.runbooks.updateActivity(projectId, runbookId, editTarget.id, payload);
        setRunbook((prev) => prev ? { ...prev, activities: prev.activities.map((a) => a.id === updated.id ? updated : a) } : prev);
      } else {
        const created = await api.runbooks.addActivity(projectId, runbookId, payload);
        setRunbook((prev) => prev ? { ...prev, activities: [...prev.activities, created] } : prev);
      }
      setActivityModal(false);
    } finally {
      setSaving(false);
    }
  }

  async function handleDeleteActivity() {
    if (!projectId || !runbookId || !deleteTarget) return;
    setDeleting(true);
    try {
      await api.runbooks.deleteActivity(projectId, runbookId, deleteTarget.id);
      setRunbook((prev) => prev ? { ...prev, activities: prev.activities.filter((a) => a.id !== deleteTarget.id) } : prev);
      setDeleteTarget(null);
    } finally {
      setDeleting(false);
    }
  }

  async function toggleStatus(activity: RunbookActivity) {
    if (!projectId || !runbookId) return;
    const next = activity.status === "done" ? "not_started" : "done";
    const updated = await api.runbooks.updateActivity(projectId, runbookId, activity.id, { ...activity, status: next });
    setRunbook((prev) => prev ? { ...prev, activities: prev.activities.map((a) => a.id === updated.id ? updated : a) } : prev);
  }

  async function handleSaveRunbook() {
    if (!projectId || !runbookId || !runbookTitle.trim()) return;
    const updated = await api.runbooks.update(projectId, runbookId, {
      title: runbookTitle,
      external_url: runbookUrl || null,
    });
    setRunbook(updated);
    setEditRunbookModal(false);
  }

  if (!runbook) return <div style={{ padding: "2rem" }}>Laster...</div>;

  const srcCfg = SOURCE_CONFIG[runbook.source];
  const phases = [...new Set(runbook.activities.map((a) => a.phase ?? ""))];
  const orderedPhases = phases.filter(Boolean);
  if (runbook.activities.some((a) => !a.phase)) orderedPhases.push("");

  const done = runbook.activities.filter((a) => a.status === "done").length;
  const inProgress = runbook.activities.filter((a) => a.status === "in_progress").length;
  const total = runbook.activities.length;

  return (
    <div style={{ maxWidth: 960, margin: "0 auto", padding: "2rem" }}>
      {/* Back */}
      <button
        onClick={() => navigate(`/projects/${projectId}`)}
        style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", marginBottom: "1.5rem", fontSize: "0.9rem", padding: 0 }}
      >
        ← Tilbake til prosjekt
      </button>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: "2rem", gap: "1rem" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "1rem" }}>
          <div style={{
            width: 48, height: 48, borderRadius: 10, background: srcCfg.color,
            color: "#fff", display: "flex", alignItems: "center", justifyContent: "center",
            fontWeight: 700, fontSize: "1.1rem", flexShrink: 0,
          }}>
            {srcCfg.initial}
          </div>
          <div>
            <h1 className="bf-h2" style={{ margin: 0 }}>{runbook.title}</h1>
            <span style={{ fontSize: "0.8rem", color: srcCfg.color, fontWeight: 600 }}>{srcCfg.label}</span>
          </div>
        </div>
        <div style={{ display: "flex", gap: "0.5rem", flexShrink: 0 }}>
          <Button variant="outline" onClick={() => { setRunbookTitle(runbook.title); setRunbookUrl(runbook.external_url ?? ""); setEditRunbookModal(true); }}>
            Rediger
          </Button>
          {runbook.source === "own" && (
            <Button variant="filled" onClick={() => openAdd()}>+ Legg til aktivitet</Button>
          )}
        </div>
      </div>

      {/* Planner integration view */}
      {runbook.source === "planner" && (
        <PlannerView
          runbook={runbook}
          srcCfg={srcCfg}
          isMsalConfigured={isMsalConfigured}
          msalReady={msalReady}
          account={plannerAccount}
          data={plannerData}
          loading={plannerLoading}
          error={plannerError}
          onLogin={loginPlanner}
          onRefresh={refreshPlanner}
        />
      )}

      {/* Smartsheet — link only for now */}
      {runbook.source === "smartsheet" && (
        <ExternalLinkCard runbook={runbook} srcCfg={srcCfg} />
      )}

      {/* Own runbook: stats + activities */}
      {runbook.source === "own" && (
        <>
          {/* Stats */}
          {total > 0 && (
            <div style={{ display: "flex", gap: "1rem", marginBottom: "1.75rem", flexWrap: "wrap" }}>
              {[
                { label: "Totalt", value: total, color: "#868E96" },
                { label: "Ferdig", value: done, color: "#2F9E44" },
                { label: "Pågående", value: inProgress, color: "#1971C2" },
                { label: "Gjenstår", value: total - done - inProgress, color: "#F76707" },
              ].map(({ label, value, color }) => (
                <div key={label} style={{
                  padding: "0.6rem 1.1rem", borderRadius: 8,
                  background: `${color}12`, border: `1px solid ${color}30`,
                  textAlign: "center", minWidth: 80,
                }}>
                  <div style={{ fontSize: "1.4rem", fontWeight: 700, color }}>{value}</div>
                  <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)" }}>{label}</div>
                </div>
              ))}
            </div>
          )}

          {/* Activities grouped by phase */}
          {total === 0 ? (
            <div style={{
              textAlign: "center", padding: "3rem",
              border: "2px dashed var(--bfc-base-dimmed)", borderRadius: 8,
              color: "var(--bfc-base-c-2)",
            }}>
              <p style={{ marginBottom: "1rem" }}>Ingen aktiviteter ennå.</p>
              <Button variant="filled" onClick={() => openAdd()}>+ Legg til første aktivitet</Button>
            </div>
          ) : (
            <div style={{ display: "grid", gap: "1.5rem" }}>
              {orderedPhases.map((phase) => {
                const acts = runbook.activities.filter((a) => (a.phase ?? "") === phase);
                return (
                  <PhaseGroup
                    key={phase || "__none__"}
                    phase={phase || "Uten fase"}
                    activities={acts}
                    onAdd={() => openAdd(phase)}
                    onEdit={openEdit}
                    onDelete={setDeleteTarget}
                    onToggleStatus={toggleStatus}
                  />
                );
              })}
            </div>
          )}
        </>
      )}

      {/* Activity form modal */}
      <Modal isOpen={activityModal} onRequestClose={() => setActivityModal(false)} header={editTarget ? "Rediger aktivitet" : "Ny aktivitet"}>
        <ActivityForm
          form={form}
          setForm={setForm}
          existingPhases={[...new Set(runbook.activities.map((a) => a.phase ?? "").filter(Boolean))]}
          saving={saving}
          onSave={handleSaveActivity}
          onCancel={() => setActivityModal(false)}
        />
      </Modal>

      {/* Delete confirmation */}
      <Modal isOpen={!!deleteTarget} onRequestClose={() => setDeleteTarget(null)} header="Slett aktivitet">
        <div style={{ display: "flex", flexDirection: "column", gap: "1.25rem" }}>
          <p style={{ margin: 0 }}>
            Er du sikker på at du vil slette <strong>«{deleteTarget?.name}»</strong>?
          </p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setDeleteTarget(null)}>Avbryt</Button>
            <Button variant="filled" state={deleting ? "inactive" : "default"} onClick={handleDeleteActivity}
              style={{ background: "#E03131", borderColor: "#E03131" }}>
              {deleting ? "Sletter..." : "Ja, slett"}
            </Button>
          </div>
        </div>
      </Modal>

      {/* Edit runbook modal */}
      <Modal isOpen={editRunbookModal} onRequestClose={() => setEditRunbookModal(false)} header="Rediger runbook">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <Input label="Tittel" value={runbookTitle} onChange={(e) => setRunbookTitle(e.target.value)} />
          {runbook.source !== "own" && (
            <Input
              label={`URL til ${srcCfg.label}`}
              value={runbookUrl}
              onChange={(e) => setRunbookUrl(e.target.value)}
              placeholder="https://..."
            />
          )}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setEditRunbookModal(false)}>Avbryt</Button>
            <Button variant="filled" onClick={handleSaveRunbook} state={!runbookTitle.trim() ? "inactive" : "default"}>
              Lagre
            </Button>
          </div>
        </div>
      </Modal>
    </div>
  );
}

// ─── Phase group ──────────────────────────────────────────────────────────────

function PhaseGroup({
  phase, activities, onAdd, onEdit, onDelete, onToggleStatus,
}: {
  phase: string;
  activities: RunbookActivity[];
  onAdd: () => void;
  onEdit: (a: RunbookActivity) => void;
  onDelete: (a: RunbookActivity) => void;
  onToggleStatus: (a: RunbookActivity) => void;
}) {
  const done = activities.filter((a) => a.status === "done").length;
  const pct = activities.length > 0 ? Math.round((done / activities.length) * 100) : 0;

  return (
    <div>
      {/* Phase header */}
      <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "0.5rem" }}>
        <h3 className="bf-h4" style={{ margin: 0, flex: 1 }}>{phase}</h3>
        <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>
          {done}/{activities.length} ferdig
        </span>
        {/* Progress bar */}
        <div style={{ width: 80, height: 6, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
          <div style={{ width: `${pct}%`, height: "100%", background: "#2F9E44", borderRadius: 3, transition: "width 0.3s" }} />
        </div>
        <button
          onClick={onAdd}
          style={{
            background: "none", border: "1px dashed var(--bfc-base-dimmed)", cursor: "pointer",
            color: "var(--bfc-base-c-2)", padding: "2px 10px", borderRadius: 4, fontSize: "0.8rem",
          }}
        >
          + Legg til
        </button>
      </div>

      {/* Activity rows */}
      <div style={{ display: "grid", gap: "0.35rem" }}>
        {activities.map((activity) => (
          <ActivityRow
            key={activity.id}
            activity={activity}
            onEdit={() => onEdit(activity)}
            onDelete={() => onDelete(activity)}
            onToggle={() => onToggleStatus(activity)}
          />
        ))}
      </div>
    </div>
  );
}

// ─── Activity row ─────────────────────────────────────────────────────────────

function ActivityRow({
  activity, onEdit, onDelete, onToggle,
}: {
  activity: RunbookActivity;
  onEdit: () => void;
  onDelete: () => void;
  onToggle: () => void;
}) {
  const [hovered, setHovered] = useState(false);
  const cfg = STATUS_CONFIG[activity.status];
  const isDone = activity.status === "done";
  const isCancelled = activity.status === "cancelled";

  return (
    <div
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        display: "flex", alignItems: "center", gap: "0.75rem",
        padding: "0.65rem 1rem",
        borderRadius: 7,
        background: "var(--bfc-base-3)",
        border: "1px solid var(--bfc-base-dimmed)",
        opacity: isCancelled ? 0.6 : 1,
        boxShadow: hovered ? "0 2px 8px rgba(0,0,0,0.07)" : "none",
        transition: "box-shadow 0.15s",
      }}
    >
      {/* Status toggle */}
      <button
        onClick={onToggle}
        title={isDone ? "Merk som ikke startet" : "Merk som ferdig"}
        style={{
          width: 22, height: 22, borderRadius: "50%", flexShrink: 0,
          border: `2px solid ${cfg.color}`,
          background: isDone ? cfg.color : "transparent",
          cursor: "pointer",
          display: "flex", alignItems: "center", justifyContent: "center",
          transition: "background 0.15s",
        }}
      >
        {isDone && <span style={{ color: "#fff", fontSize: "0.7rem", fontWeight: 700 }}>✓</span>}
      </button>

      {/* Name */}
      <span
        style={{
          flex: 1, minWidth: 0,
          fontSize: "0.9rem",
          textDecoration: isDone || isCancelled ? "line-through" : "none",
          color: isDone || isCancelled ? "var(--bfc-base-c-3)" : "inherit",
          overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
        }}
      >
        {activity.name}
      </span>

      {/* Responsible */}
      {activity.responsible && (
        <span style={{
          fontSize: "0.75rem", padding: "2px 8px", borderRadius: 20,
          background: "var(--bfc-base-dimmed)", color: "var(--bfc-base-c-2)",
          flexShrink: 0,
        }}>
          {activity.responsible}
        </span>
      )}

      {/* Date range */}
      {(activity.start_date || activity.end_date) && (
        <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)", flexShrink: 0, whiteSpace: "nowrap" }}>
          {activity.start_date && new Date(activity.start_date).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
          {activity.start_date && activity.end_date && " → "}
          {activity.end_date && new Date(activity.end_date).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
        </span>
      )}

      {/* Status badge */}
      <span style={{
        fontSize: "0.72rem", fontWeight: 600, padding: "2px 8px", borderRadius: 20,
        background: cfg.bg, color: cfg.color, flexShrink: 0,
      }}>
        {cfg.label}
      </span>

      {/* Actions */}
      <div
        style={{ display: "flex", gap: "0.25rem", flexShrink: 0, opacity: hovered ? 1 : 0, transition: "opacity 0.15s" }}
        onClick={(e) => e.stopPropagation()}
      >
        <button onClick={onEdit} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "2px 6px", borderRadius: 4, fontSize: "0.8rem" }}>
          Endre
        </button>
        <button onClick={onDelete} style={{ background: "none", border: "none", cursor: "pointer", color: "#E03131", padding: "2px 6px", borderRadius: 4, fontSize: "0.8rem" }}>
          Slett
        </button>
      </div>
    </div>
  );
}

// ─── Activity form ─────────────────────────────────────────────────────────────

type FormState = typeof EMPTY_FORM;

function ActivityForm({
  form, setForm, existingPhases, saving, onSave, onCancel,
}: {
  form: FormState;
  setForm: (f: FormState) => void;
  existingPhases: string[];
  saving: boolean;
  onSave: () => void;
  onCancel: () => void;
}) {
  const set = (key: keyof FormState) => (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) =>
    setForm({ ...form, [key]: e.target.value });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: "0.85rem" }}>
      <Input label="Aktivitetsnavn *" value={form.name} onChange={set("name")} placeholder="Beskriv aktiviteten" autoFocus />

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.75rem" }}>
        <div>
          <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 500, marginBottom: "0.35rem" }}>Fase</label>
          <input
            list="phases-list"
            value={form.phase}
            onChange={set("phase")}
            placeholder="f.eks. Forberedelser"
            style={{
              width: "100%", padding: "0.5rem 0.75rem", borderRadius: 6,
              border: "1px solid var(--bfc-base-dimmed)", background: "var(--bfc-base-3)",
              color: "inherit", fontSize: "0.9rem", boxSizing: "border-box",
            }}
          />
          <datalist id="phases-list">
            {existingPhases.map((p) => <option key={p} value={p} />)}
          </datalist>
        </div>

        <div>
          <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 500, marginBottom: "0.35rem" }}>Status</label>
          <select
            value={form.status}
            onChange={set("status")}
            style={{
              width: "100%", padding: "0.5rem 0.75rem", borderRadius: 6,
              border: "1px solid var(--bfc-base-dimmed)", background: "var(--bfc-base-3)",
              color: "inherit", fontSize: "0.9rem",
            }}
          >
            <option value="not_started">Ikke startet</option>
            <option value="in_progress">Pågående</option>
            <option value="done">Ferdig</option>
            <option value="cancelled">Kansellert</option>
          </select>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.75rem" }}>
        <Input label="Start dato/tid" type="datetime-local" value={form.start_date} onChange={set("start_date")} />
        <Input label="Slutt dato/tid" type="datetime-local" value={form.end_date} onChange={set("end_date")} />
      </div>

      <Input label="Ansvarlig" value={form.responsible} onChange={set("responsible")} placeholder="Navn eller team" />

      <div>
        <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 500, marginBottom: "0.35rem" }}>Beskrivelse</label>
        <textarea
          value={form.description}
          onChange={set("description")}
          rows={3}
          placeholder="Valgfrie notater eller instruksjoner"
          style={{
            width: "100%", padding: "0.5rem 0.75rem", borderRadius: 6,
            border: "1px solid var(--bfc-base-dimmed)", background: "var(--bfc-base-3)",
            color: "inherit", fontSize: "0.9rem", resize: "vertical", boxSizing: "border-box",
          }}
        />
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", marginTop: "0.25rem" }}>
        <Button onClick={onCancel}>Avbryt</Button>
        <Button variant="filled" onClick={onSave} state={!form.name.trim() || saving ? "inactive" : "default"}>
          {saving ? "Lagrer..." : "Lagre"}
        </Button>
      </div>
    </div>
  );
}

// ─── External link card (Smartsheet) ─────────────────────────────────────────

function ExternalLinkCard({ runbook, srcCfg }: { runbook: Runbook; srcCfg: { label: string; color: string } }) {
  return (
    <div style={{
      padding: "2rem", borderRadius: 10, background: "var(--bfc-base-3)",
      border: "1px solid var(--bfc-base-dimmed)", borderTop: `4px solid ${srcCfg.color}`,
      textAlign: "center", maxWidth: 540, margin: "0 auto",
    }}>
      <div style={{ fontSize: "2rem", marginBottom: "0.75rem" }}>📊</div>
      <h2 className="bf-h4" style={{ margin: "0 0 0.5rem" }}>Styres i {srcCfg.label}</h2>
      <p style={{ color: "var(--bfc-base-c-2)", marginBottom: "1.5rem", fontSize: "0.9rem" }}>
        Bruk lenken under for å se og redigere innholdet i {srcCfg.label}.
      </p>
      {runbook.external_url ? (
        <a href={runbook.external_url} target="_blank" rel="noopener noreferrer"
          style={{ display: "inline-block", padding: "0.6rem 1.5rem", background: srcCfg.color, color: "#fff", borderRadius: 6, fontWeight: 600, textDecoration: "none" }}>
          Åpne i {srcCfg.label} →
        </a>
      ) : (
        <p style={{ color: "var(--bfc-base-c-3)", fontSize: "0.85rem" }}>Ingen URL lagt til. Trykk «Rediger» for å legge til.</p>
      )}
    </div>
  );
}

// ─── Planner integration view ─────────────────────────────────────────────────

function PlannerView({
  runbook, srcCfg, isMsalConfigured: configured, msalReady, account, data, loading, error, onLogin, onRefresh,
}: {
  runbook: Runbook;
  srcCfg: { label: string; color: string };
  isMsalConfigured: boolean;
  msalReady: boolean;
  account: AccountInfo | null;
  data: PlannerData | null;
  loading: boolean;
  error: string | null;
  onLogin: () => void;
  onRefresh: () => void;
}) {
  const planId = runbook.external_url ? parsePlanId(runbook.external_url) : null;

  if (!configured) {
    return (
      <div style={{ padding: "1.5rem", borderRadius: 10, background: "#FFF3BF", border: "1px solid #FAB005", maxWidth: 600, margin: "0 auto" }}>
        <h3 style={{ margin: "0 0 0.5rem", color: "#5C3A00" }}>Azure AD ikke konfigurert</h3>
        <p style={{ margin: 0, fontSize: "0.9rem", color: "#7C4D00" }}>
          For å koble til Microsoft Planner kreves en Azure AD-app-registrering. Legg til
          {" "}<code>VITE_AZURE_CLIENT_ID</code> og <code>VITE_AZURE_TENANT_ID</code> som
          GitHub Secrets og bygg på nytt.
        </p>
      </div>
    );
  }

  if (!runbook.external_url || !planId) {
    return (
      <div style={{ textAlign: "center", padding: "2rem", color: "var(--bfc-base-c-2)" }}>
        <p>Ingen Planner-URL er lagt til ennå.</p>
        <p style={{ fontSize: "0.85rem" }}>Trykk «Rediger» og lim inn lenken til Planner-planen.</p>
      </div>
    );
  }

  return (
    <div>
      {/* Toolbar */}
      <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "1.25rem", flexWrap: "wrap" }}>
        {account ? (
          <>
            <span style={{ fontSize: "0.85rem", color: "var(--bfc-base-c-2)" }}>
              Innlogget som <strong>{account.name ?? account.username}</strong>
            </span>
            <Button variant="outline" onClick={onRefresh} state={loading ? "inactive" : "default"}>
              {loading ? "Henter..." : "Oppdater"}
            </Button>
            {runbook.external_url && (
              <a href={runbook.external_url} target="_blank" rel="noopener noreferrer"
                style={{ fontSize: "0.85rem", color: srcCfg.color, fontWeight: 600, textDecoration: "none" }}>
                Åpne i Planner →
              </a>
            )}
          </>
        ) : (
          <Button variant="filled" onClick={onLogin} state={!msalReady ? "inactive" : "default"}
            style={{ background: "#0078D4", borderColor: "#0078D4" }}>
            Logg inn med Microsoft for å hente aktiviteter
          </Button>
        )}
      </div>

      {error && (
        <div style={{ padding: "0.75rem 1rem", borderRadius: 6, background: "#FFE3E3", color: "#C92A2A", marginBottom: "1rem", fontSize: "0.9rem" }}>
          {error}
        </div>
      )}

      {loading && (
        <div style={{ textAlign: "center", padding: "3rem", color: "var(--bfc-base-c-2)" }}>
          Henter aktiviteter fra Microsoft Planner…
        </div>
      )}

      {data && !loading && <PlannerTaskList data={data} srcCfg={srcCfg} />}
    </div>
  );
}

// ─── Planner task list ────────────────────────────────────────────────────────

function PlannerTaskList({ data, srcCfg }: { data: PlannerData; srcCfg: { color: string } }) {
  const bucketMap = Object.fromEntries(data.buckets.map((b) => [b.id, b.name]));
  const grouped: Record<string, PlannerTask[]> = {};
  for (const task of data.tasks) {
    const bucket = bucketMap[task.bucketId] ?? "Ukjent fase";
    if (!grouped[bucket]) grouped[bucket] = [];
    grouped[bucket].push(task);
  }

  const done = data.tasks.filter((t) => t.percentComplete === 100).length;
  const inProgress = data.tasks.filter((t) => t.percentComplete > 0 && t.percentComplete < 100).length;
  const total = data.tasks.length;

  return (
    <div>
      {/* Stats */}
      <div style={{ display: "flex", gap: "1rem", marginBottom: "1.75rem", flexWrap: "wrap" }}>
        {[
          { label: "Totalt",    value: total,               color: "#868E96" },
          { label: "Ferdig",    value: done,                color: "#2F9E44" },
          { label: "Pågående",  value: inProgress,          color: "#1971C2" },
          { label: "Gjenstår",  value: total - done - inProgress, color: "#F76707" },
        ].map(({ label, value, color }) => (
          <div key={label} style={{ padding: "0.6rem 1.1rem", borderRadius: 8, background: `${color}12`, border: `1px solid ${color}30`, textAlign: "center", minWidth: 80 }}>
            <div style={{ fontSize: "1.4rem", fontWeight: 700, color }}>{value}</div>
            <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)" }}>{label}</div>
          </div>
        ))}
      </div>

      {/* Buckets */}
      <div style={{ display: "grid", gap: "1.5rem" }}>
        {Object.entries(grouped).map(([bucket, tasks]) => {
          const bucketDone = tasks.filter((t) => t.percentComplete === 100).length;
          const pct = tasks.length > 0 ? Math.round((bucketDone / tasks.length) * 100) : 0;
          return (
            <div key={bucket}>
              <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "0.5rem" }}>
                <h3 className="bf-h4" style={{ margin: 0, flex: 1 }}>{bucket}</h3>
                <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>{bucketDone}/{tasks.length} ferdig</span>
                <div style={{ width: 80, height: 6, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
                  <div style={{ width: `${pct}%`, height: "100%", background: "#2F9E44", borderRadius: 3, transition: "width 0.3s" }} />
                </div>
              </div>
              <div style={{ display: "grid", gap: "0.35rem" }}>
                {tasks.map((task) => <PlannerTaskRow key={task.id} task={task} />)}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function PlannerTaskRow({ task }: { task: PlannerTask }) {
  const [hovered, setHovered] = useState(false);
  const status = taskStatus(task.percentComplete);
  const cfg = STATUS_CONFIG[status];
  const isDone = status === "done";

  return (
    <div
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        display: "flex", alignItems: "center", gap: "0.75rem",
        padding: "0.65rem 1rem", borderRadius: 7,
        background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)",
        boxShadow: hovered ? "0 2px 8px rgba(0,0,0,0.07)" : "none",
        transition: "box-shadow 0.15s",
      }}
    >
      {/* Status circle (read-only) */}
      <div style={{
        width: 22, height: 22, borderRadius: "50%", flexShrink: 0,
        border: `2px solid ${cfg.color}`,
        background: isDone ? cfg.color : "transparent",
        display: "flex", alignItems: "center", justifyContent: "center",
      }}>
        {isDone && <span style={{ color: "#fff", fontSize: "0.7rem", fontWeight: 700 }}>✓</span>}
      </div>

      <span style={{
        flex: 1, fontSize: "0.9rem",
        textDecoration: isDone ? "line-through" : "none",
        color: isDone ? "var(--bfc-base-c-3)" : "inherit",
        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
      }}>
        {task.title}
      </span>

      {/* Progress bar for partially done */}
      {task.percentComplete > 0 && task.percentComplete < 100 && (
        <div style={{ width: 60, height: 5, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden", flexShrink: 0 }}>
          <div style={{ width: `${task.percentComplete}%`, height: "100%", background: cfg.color, borderRadius: 3 }} />
        </div>
      )}

      {/* Dates */}
      {(task.startDateTime || task.dueDateTime) && (
        <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)", flexShrink: 0, whiteSpace: "nowrap" }}>
          {task.startDateTime && new Date(task.startDateTime).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
          {task.startDateTime && task.dueDateTime && " → "}
          {task.dueDateTime && new Date(task.dueDateTime).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
        </span>
      )}

      <span style={{ fontSize: "0.72rem", fontWeight: 600, padding: "2px 8px", borderRadius: 20, background: cfg.bg, color: cfg.color, flexShrink: 0 }}>
        {cfg.label}
      </span>
    </div>
  );
}
