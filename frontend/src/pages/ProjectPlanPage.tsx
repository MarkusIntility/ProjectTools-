import { useEffect, useState, useCallback } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type ProjectPlan, type ProjectPlanTask } from "../api/client";
import { isMsalConfigured, msalInstance, PLANNER_SCOPES } from "../auth/msalConfig";
import { fetchPlannerData, parsePlanId, type PlannerData, type PlannerTask } from "../auth/plannerService";
import type { AccountInfo } from "@azure/msal-browser";

// ─── Config ───────────────────────────────────────────────────────────────────

const SOURCE_CONFIG = {
  planner:    { label: "Microsoft Planner", color: "#0078D4", initial: "P" },
  smartsheet: { label: "Smartsheet",        color: "#00A88E", initial: "S" },
  own:        { label: "Egen",              color: "#7950F2", initial: "E" },
};

const PCT_OPTIONS = [0, 10, 25, 50, 75, 90, 100];

const EMPTY_FORM = {
  name: "",
  bucket: "",
  percent_complete: 0,
  start_date: "",
  end_date: "",
  responsible: "",
  description: "",
};

type FormState = typeof EMPTY_FORM;

// ─── Main page ────────────────────────────────────────────────────────────────

export default function ProjectPlanPage() {
  const { projectId, planId } = useParams<{ projectId: string; planId: string }>();
  const navigate = useNavigate();
  const [plan, setPlan] = useState<ProjectPlan | null>(null);

  // Planner state
  const [msalReady] = useState(isMsalConfigured);
  const [plannerAccount, setPlannerAccount] = useState<AccountInfo | null>(null);
  const [plannerData, setPlannerData] = useState<PlannerData | null>(null);
  const [plannerLoading, setPlannerLoading] = useState(false);
  const [plannerError, setPlannerError] = useState<string | null>(null);

  const [taskModal, setTaskModal] = useState(false);
  const [editTarget, setEditTarget] = useState<ProjectPlanTask | null>(null);
  const [form, setForm] = useState<FormState>(EMPTY_FORM);
  const [saving, setSaving] = useState(false);

  const [deleteTarget, setDeleteTarget] = useState<ProjectPlanTask | null>(null);
  const [deleting, setDeleting] = useState(false);

  const [editPlanModal, setEditPlanModal] = useState(false);
  const [planTitle, setPlanTitle] = useState("");
  const [planUrl, setPlanUrl] = useState("");

  useEffect(() => {
    if (!projectId || !planId) return;
    api.projectPlans.get(projectId, planId).then(setPlan);
  }, [projectId, planId]);

  useEffect(() => {
    if (!isMsalConfigured) return;
    const bootError = sessionStorage.getItem("msal.bootError");
    if (bootError) {
      sessionStorage.removeItem("msal.bootError");
      setPlannerError(`Innlogging feilet: ${bootError}`);
      return;
    }
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) setPlannerAccount(accounts[0]);
  }, []);

  const loadPlannerData = useCallback(async (account: AccountInfo, planUrl: string) => {
    setPlannerLoading(true);
    setPlannerError(null);
    try {
      const data = await fetchPlannerData(msalInstance, account, planUrl);
      setPlannerData(data);
    } catch (err) {
      setPlannerError(err instanceof Error ? err.message : "Ukjent feil");
    } finally {
      setPlannerLoading(false);
    }
  }, []);

  useEffect(() => {
    if (!plannerAccount || !plan || plan.source !== "planner" || !plan.external_url) return;
    loadPlannerData(plannerAccount, plan.external_url);
  }, [plannerAccount, plan, loadPlannerData]);

  async function loginPlanner() {
    if (!isMsalConfigured || !msalReady) return;
    Object.keys(sessionStorage)
      .filter((k) => k.includes("interaction"))
      .forEach((k) => sessionStorage.removeItem(k));
    sessionStorage.setItem("app.returnUrl", window.location.href);
    await msalInstance.loginRedirect({ scopes: PLANNER_SCOPES });
  }

  function refreshPlanner() {
    if (!plannerAccount || !plan?.external_url) return;
    loadPlannerData(plannerAccount, plan.external_url);
  }

  function openAdd(defaultBucket?: string) {
    setEditTarget(null);
    setForm({ ...EMPTY_FORM, bucket: defaultBucket ?? "" });
    setTaskModal(true);
  }

  function openEdit(task: ProjectPlanTask) {
    setEditTarget(task);
    setForm({
      name: task.name,
      bucket: task.bucket ?? "",
      percent_complete: task.percent_complete,
      start_date: task.start_date ? task.start_date.slice(0, 16) : "",
      end_date: task.end_date ? task.end_date.slice(0, 16) : "",
      responsible: task.responsible ?? "",
      description: task.description ?? "",
    });
    setTaskModal(true);
  }

  async function handleSaveTask() {
    if (!projectId || !planId || !form.name.trim()) return;
    setSaving(true);
    try {
      const payload = {
        name: form.name,
        bucket: form.bucket || null,
        percent_complete: form.percent_complete,
        start_date: form.start_date || null,
        end_date: form.end_date || null,
        responsible: form.responsible || null,
        description: form.description || null,
        sort_order: editTarget?.sort_order ?? 0,
      };
      if (editTarget) {
        const updated = await api.projectPlans.updateTask(projectId, planId, editTarget.id, payload);
        setPlan((prev) => prev ? { ...prev, tasks: prev.tasks.map((t) => t.id === updated.id ? updated : t) } : prev);
      } else {
        const created = await api.projectPlans.addTask(projectId, planId, payload);
        setPlan((prev) => prev ? { ...prev, tasks: [...prev.tasks, created] } : prev);
      }
      setTaskModal(false);
    } finally {
      setSaving(false);
    }
  }

  async function handleDeleteTask() {
    if (!projectId || !planId || !deleteTarget) return;
    setDeleting(true);
    try {
      await api.projectPlans.deleteTask(projectId, planId, deleteTarget.id);
      setPlan((prev) => prev ? { ...prev, tasks: prev.tasks.filter((t) => t.id !== deleteTarget.id) } : prev);
      setDeleteTarget(null);
    } finally {
      setDeleting(false);
    }
  }

  async function updateTaskProgress(task: ProjectPlanTask, pct: number) {
    if (!projectId || !planId) return;
    const updated = await api.projectPlans.updateTask(projectId, planId, task.id, { ...task, percent_complete: pct });
    setPlan((prev) => prev ? { ...prev, tasks: prev.tasks.map((t) => t.id === updated.id ? updated : t) } : prev);
  }

  async function handleSavePlan() {
    if (!projectId || !planId || !planTitle.trim()) return;
    const updated = await api.projectPlans.update(projectId, planId, {
      title: planTitle,
      external_url: planUrl || null,
    });
    setPlan(updated);
    setEditPlanModal(false);
  }

  if (!plan) return <div style={{ padding: "2rem" }}>Laster...</div>;

  const srcCfg = SOURCE_CONFIG[plan.source];

  return (
    <div style={{ maxWidth: 1000, margin: "0 auto", padding: "2rem" }}>
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
            <h1 className="bf-h2" style={{ margin: 0 }}>{plan.title}</h1>
            <span style={{ fontSize: "0.8rem", color: srcCfg.color, fontWeight: 600 }}>{srcCfg.label}</span>
          </div>
        </div>
        <div style={{ display: "flex", gap: "0.5rem", flexShrink: 0 }}>
          <Button variant="outline" onClick={() => { setPlanTitle(plan.title); setPlanUrl(plan.external_url ?? ""); setEditPlanModal(true); }}>
            Rediger
          </Button>
          {plan.source === "own" && (
            <Button variant="filled" onClick={() => openAdd()}>+ Legg til oppgave</Button>
          )}
        </div>
      </div>

      {/* Planner view */}
      {plan.source === "planner" && (
        <PlannerView
          plan={plan}
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

      {/* Smartsheet */}
      {plan.source === "smartsheet" && (
        <ExternalLinkCard plan={plan} srcCfg={srcCfg} />
      )}

      {/* Own plan */}
      {plan.source === "own" && (
        <OwnPlanView
          plan={plan}
          onAdd={openAdd}
          onEdit={openEdit}
          onDelete={setDeleteTarget}
          onUpdateProgress={updateTaskProgress}
        />
      )}

      {/* Task form modal */}
      <Modal isOpen={taskModal} onRequestClose={() => setTaskModal(false)} header={editTarget ? "Rediger oppgave" : "Ny oppgave"}>
        <TaskForm
          form={form}
          setForm={setForm}
          existingBuckets={[...new Set(plan.tasks.map((t) => t.bucket ?? "").filter(Boolean))]}
          saving={saving}
          onSave={handleSaveTask}
          onCancel={() => setTaskModal(false)}
        />
      </Modal>

      {/* Delete confirmation */}
      <Modal isOpen={!!deleteTarget} onRequestClose={() => setDeleteTarget(null)} header="Slett oppgave">
        <div style={{ display: "flex", flexDirection: "column", gap: "1.25rem" }}>
          <p style={{ margin: 0 }}>Er du sikker på at du vil slette <strong>«{deleteTarget?.name}»</strong>?</p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setDeleteTarget(null)}>Avbryt</Button>
            <Button variant="filled" state={deleting ? "inactive" : "default"} onClick={handleDeleteTask}
              style={{ background: "#E03131", borderColor: "#E03131" }}>
              {deleting ? "Sletter..." : "Ja, slett"}
            </Button>
          </div>
        </div>
      </Modal>

      {/* Edit plan modal */}
      <Modal isOpen={editPlanModal} onRequestClose={() => setEditPlanModal(false)} header="Rediger prosjektplan">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <Input label="Tittel" value={planTitle} onChange={(e) => setPlanTitle(e.target.value)} />
          {plan.source !== "own" && (
            <Input
              label={`URL til ${srcCfg.label}`}
              value={planUrl}
              onChange={(e) => setPlanUrl(e.target.value)}
              placeholder="https://..."
            />
          )}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setEditPlanModal(false)}>Avbryt</Button>
            <Button variant="filled" onClick={handleSavePlan} state={!planTitle.trim() ? "inactive" : "default"}>Lagre</Button>
          </div>
        </div>
      </Modal>
    </div>
  );
}

// ─── Filter helpers ───────────────────────────────────────────────────────────

type Filter = "all" | "done" | "in_progress" | "remaining";

const FILTER_LABELS: Record<Filter, string> = {
  all:         "Alle",
  done:        "Ferdig",
  in_progress: "Pågående",
  remaining:   "Gjenstår",
};

function matchOwnFilter(pct: number, f: Filter): boolean {
  if (f === "done")        return pct === 100;
  if (f === "in_progress") return pct > 0 && pct < 100;
  if (f === "remaining")   return pct === 0;
  return true;
}

function matchPlannerFilter(pct: number, f: Filter): boolean {
  if (f === "done")        return pct === 100;
  if (f === "in_progress") return pct > 0 && pct < 100;
  if (f === "remaining")   return pct === 0;
  return true;
}

// ─── Own plan view ────────────────────────────────────────────────────────────

function OwnPlanView({
  plan, onAdd, onEdit, onDelete, onUpdateProgress,
}: {
  plan: ProjectPlan;
  onAdd: (bucket?: string) => void;
  onEdit: (t: ProjectPlanTask) => void;
  onDelete: (t: ProjectPlanTask) => void;
  onUpdateProgress: (t: ProjectPlanTask, pct: number) => void;
}) {
  const [filter, setFilter] = useState<Filter>("all");

  const total = plan.tasks.length;
  const done = plan.tasks.filter((t) => t.percent_complete === 100).length;
  const inProg = plan.tasks.filter((t) => t.percent_complete > 0 && t.percent_complete < 100).length;
  const remaining = total - done - inProg;
  const overallPct = total > 0 ? Math.round(plan.tasks.reduce((sum, t) => sum + t.percent_complete, 0) / total) : 0;

  const visibleTasks = plan.tasks.filter((t) => matchOwnFilter(t.percent_complete, filter));

  const buckets = [...new Set(plan.tasks.map((t) => t.bucket ?? ""))];
  const ordered = buckets.filter(Boolean);
  if (plan.tasks.some((t) => !t.bucket)) ordered.push("");

  if (total === 0) {
    return (
      <div style={{ textAlign: "center", padding: "3rem", border: "2px dashed var(--bfc-base-dimmed)", borderRadius: 8, color: "var(--bfc-base-c-2)" }}>
        <p style={{ marginBottom: "1rem" }}>Ingen oppgaver ennå.</p>
        <Button variant="filled" onClick={() => onAdd()}>+ Legg til første oppgave</Button>
      </div>
    );
  }

  const STAT_CARDS = [
    { key: "all"         as Filter, label: "Totalt",   value: total,     color: "#868E96" },
    { key: "done"        as Filter, label: "Ferdig",   value: done,      color: "#2F9E44" },
    { key: "in_progress" as Filter, label: "Pågående", value: inProg,    color: "#1971C2" },
    { key: "remaining"   as Filter, label: "Gjenstår", value: remaining, color: "#F76707" },
  ];

  return (
    <div>
      {/* Stats + overall progress */}
      <div style={{ display: "flex", gap: "1rem", marginBottom: "1rem", flexWrap: "wrap", alignItems: "center" }}>
        {STAT_CARDS.map(({ key, label, value, color }) => {
          const isActive = filter === key;
          return (
            <div
              key={key}
              onClick={() => setFilter(prev => prev === key ? "all" : key)}
              style={{
                padding: "0.6rem 1.1rem", borderRadius: 8, textAlign: "center", minWidth: 80,
                background: isActive ? `${color}22` : `${color}12`,
                border: `1px solid ${isActive ? color : `${color}30`}`,
                cursor: key === "all" ? "default" : "pointer",
                outline: isActive && key !== "all" ? `2px solid ${color}` : "none",
                outlineOffset: 1,
                transition: "all 0.15s",
              }}
            >
              <div style={{ fontSize: "1.4rem", fontWeight: 700, color }}>{value}</div>
              <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)" }}>{label}</div>
            </div>
          );
        })}
        <div style={{ flex: 1, minWidth: 200 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.8rem", color: "var(--bfc-base-c-2)", marginBottom: "0.3rem" }}>
            <span>Samlet fremdrift</span><span>{overallPct}%</span>
          </div>
          <div style={{ height: 8, borderRadius: 4, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
            <div style={{ width: `${overallPct}%`, height: "100%", background: overallPct === 100 ? "#2F9E44" : "#1971C2", borderRadius: 4, transition: "width 0.3s" }} />
          </div>
        </div>
      </div>

      {/* Active filter indicator */}
      {filter !== "all" && (
        <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "1rem" }}>
          <span style={{ fontSize: "0.8rem", color: "var(--bfc-base-c-2)" }}>
            Viser: <strong>{FILTER_LABELS[filter]}</strong> ({visibleTasks.length} oppgaver)
          </span>
          <button
            onClick={() => setFilter("all")}
            style={{ background: "none", border: "1px solid var(--bfc-base-dimmed)", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "1px 8px", borderRadius: 4, fontSize: "0.78rem" }}
          >
            × Fjern filter
          </button>
        </div>
      )}

      {/* Buckets */}
      <div style={{ display: "grid", gap: "1.5rem" }}>
        {ordered.map((bucket) => {
          const tasks = visibleTasks.filter((t) => (t.bucket ?? "") === bucket);
          if (tasks.length === 0 && filter !== "all") return null;
          const allTasks = plan.tasks.filter((t) => (t.bucket ?? "") === bucket);
          return (
            <BucketGroup
              key={bucket || "__none__"}
              bucket={bucket || "Uten fase"}
              tasks={tasks}
              allTaskCount={allTasks.length}
              filtered={filter !== "all"}
              onAdd={() => onAdd(bucket)}
              onEdit={onEdit}
              onDelete={onDelete}
              onUpdateProgress={onUpdateProgress}
            />
          );
        })}
      </div>
    </div>
  );
}

// ─── Bucket group ─────────────────────────────────────────────────────────────

function BucketGroup({
  bucket, tasks, allTaskCount, filtered, onAdd, onEdit, onDelete, onUpdateProgress,
}: {
  bucket: string;
  tasks: ProjectPlanTask[];
  allTaskCount: number;
  filtered: boolean;
  onAdd: () => void;
  onEdit: (t: ProjectPlanTask) => void;
  onDelete: (t: ProjectPlanTask) => void;
  onUpdateProgress: (t: ProjectPlanTask, pct: number) => void;
}) {
  const avg = allTaskCount > 0 ? Math.round(tasks.reduce((s, t) => s + t.percent_complete, 0) / allTaskCount) : 0;
  const done = tasks.filter((t) => t.percent_complete === 100).length;

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "0.5rem" }}>
        <h3 className="bf-h4" style={{ margin: 0, flex: 1 }}>{bucket}</h3>
        <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>
          {done}/{allTaskCount} ferdig · {avg}%
          {filtered && tasks.length < allTaskCount && ` (viser ${tasks.length})`}
        </span>
        <div style={{ width: 80, height: 6, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
          <div style={{ width: `${avg}%`, height: "100%", background: avg === 100 ? "#2F9E44" : "#1971C2", borderRadius: 3, transition: "width 0.3s" }} />
        </div>
        <button
          onClick={onAdd}
          style={{ background: "none", border: "1px dashed var(--bfc-base-dimmed)", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "2px 10px", borderRadius: 4, fontSize: "0.8rem" }}
        >
          + Legg til
        </button>
      </div>
      <div style={{ display: "grid", gap: "0.35rem" }}>
        {tasks.map((task) => (
          <TaskRow
            key={task.id}
            task={task}
            onEdit={() => onEdit(task)}
            onDelete={() => onDelete(task)}
            onUpdateProgress={(pct) => onUpdateProgress(task, pct)}
          />
        ))}
      </div>
    </div>
  );
}

// ─── Task row ─────────────────────────────────────────────────────────────────

function TaskRow({ task, onEdit, onDelete, onUpdateProgress }: {
  task: ProjectPlanTask;
  onEdit: () => void;
  onDelete: () => void;
  onUpdateProgress: (pct: number) => void;
}) {
  const [hovered, setHovered] = useState(false);
  const pct = task.percent_complete;
  const isDone = pct === 100;
  const color = isDone ? "#2F9E44" : pct > 0 ? "#1971C2" : "#868E96";

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
      {/* Progress circle toggle */}
      <button
        onClick={() => onUpdateProgress(isDone ? 0 : 100)}
        title={isDone ? "Merk som ikke startet" : "Merk som ferdig"}
        style={{
          width: 24, height: 24, borderRadius: "50%", flexShrink: 0,
          border: `2px solid ${color}`,
          background: isDone ? color : "transparent",
          cursor: "pointer",
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: "0.65rem", fontWeight: 700, color: isDone ? "#fff" : color,
          transition: "background 0.15s",
        }}
      >
        {isDone ? "✓" : pct > 0 ? `${pct}` : ""}
      </button>

      {/* Name */}
      <span style={{
        flex: 1, fontSize: "0.9rem",
        textDecoration: isDone ? "line-through" : "none",
        color: isDone ? "var(--bfc-base-c-3)" : "inherit",
        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
      }}>
        {task.name}
      </span>

      {/* Responsible */}
      {task.responsible && (
        <span style={{ fontSize: "0.75rem", padding: "2px 8px", borderRadius: 20, background: "var(--bfc-base-dimmed)", color: "var(--bfc-base-c-2)", flexShrink: 0 }}>
          {task.responsible}
        </span>
      )}

      {/* Dates */}
      {(task.start_date || task.end_date) && (
        <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)", flexShrink: 0, whiteSpace: "nowrap" }}>
          {task.start_date && new Date(task.start_date).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
          {task.start_date && task.end_date && " → "}
          {task.end_date && new Date(task.end_date).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
        </span>
      )}

      {/* % progress bar */}
      <div style={{ width: 64, flexShrink: 0 }}>
        <div style={{ display: "flex", justifyContent: "flex-end", fontSize: "0.7rem", color, marginBottom: 2 }}>{pct}%</div>
        <div style={{ height: 4, borderRadius: 2, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
          <div style={{ width: `${pct}%`, height: "100%", background: color, borderRadius: 2, transition: "width 0.3s" }} />
        </div>
      </div>

      {/* Actions */}
      <div style={{ display: "flex", gap: "0.25rem", flexShrink: 0, opacity: hovered ? 1 : 0, transition: "opacity 0.15s" }} onClick={(e) => e.stopPropagation()}>
        <button onClick={onEdit} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "2px 6px", borderRadius: 4, fontSize: "0.8rem" }}>Endre</button>
        <button onClick={onDelete} style={{ background: "none", border: "none", cursor: "pointer", color: "#E03131", padding: "2px 6px", borderRadius: 4, fontSize: "0.8rem" }}>Slett</button>
      </div>
    </div>
  );
}

// ─── Task form ────────────────────────────────────────────────────────────────

function TaskForm({ form, setForm, existingBuckets, saving, onSave, onCancel }: {
  form: FormState;
  setForm: (f: FormState) => void;
  existingBuckets: string[];
  saving: boolean;
  onSave: () => void;
  onCancel: () => void;
}) {
  const set = (key: keyof FormState) => (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) =>
    setForm({ ...form, [key]: key === "percent_complete" ? Number(e.target.value) : e.target.value });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: "0.85rem" }}>
      <Input label="Oppgavenavn *" value={form.name} onChange={set("name")} placeholder="Beskriv oppgaven" autoFocus />

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.75rem" }}>
        <div>
          <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 500, marginBottom: "0.35rem" }}>Fase / bucket</label>
          <input
            list="buckets-list"
            value={form.bucket}
            onChange={set("bucket")}
            placeholder="f.eks. Kartlegging"
            style={{ width: "100%", padding: "0.5rem 0.75rem", borderRadius: 6, border: "1px solid var(--bfc-base-dimmed)", background: "var(--bfc-base-3)", color: "inherit", fontSize: "0.9rem", boxSizing: "border-box" }}
          />
          <datalist id="buckets-list">
            {existingBuckets.map((b) => <option key={b} value={b} />)}
          </datalist>
        </div>
        <div>
          <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 500, marginBottom: "0.35rem" }}>% fullført</label>
          <select
            value={form.percent_complete}
            onChange={set("percent_complete")}
            style={{ width: "100%", padding: "0.5rem 0.75rem", borderRadius: 6, border: "1px solid var(--bfc-base-dimmed)", background: "var(--bfc-base-3)", color: "inherit", fontSize: "0.9rem" }}
          >
            {PCT_OPTIONS.map((p) => <option key={p} value={p}>{p}%</option>)}
          </select>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.75rem" }}>
        <Input label="Start dato" type="datetime-local" value={form.start_date} onChange={set("start_date")} />
        <Input label="Slutt dato" type="datetime-local" value={form.end_date} onChange={set("end_date")} />
      </div>

      <Input label="Ansvarlig" value={form.responsible} onChange={set("responsible")} placeholder="Navn eller team" />

      <div>
        <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 500, marginBottom: "0.35rem" }}>Beskrivelse</label>
        <textarea
          value={form.description}
          onChange={set("description")}
          rows={3}
          placeholder="Valgfrie notater"
          style={{ width: "100%", padding: "0.5rem 0.75rem", borderRadius: 6, border: "1px solid var(--bfc-base-dimmed)", background: "var(--bfc-base-3)", color: "inherit", fontSize: "0.9rem", resize: "vertical", boxSizing: "border-box" }}
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

// ─── External link card ───────────────────────────────────────────────────────

function ExternalLinkCard({ plan, srcCfg }: { plan: ProjectPlan; srcCfg: { label: string; color: string } }) {
  return (
    <div style={{ padding: "2rem", borderRadius: 10, background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)", borderTop: `4px solid ${srcCfg.color}`, textAlign: "center", maxWidth: 540, margin: "0 auto" }}>
      <div style={{ fontSize: "2rem", marginBottom: "0.75rem" }}>📊</div>
      <h2 className="bf-h4" style={{ margin: "0 0 0.5rem" }}>Styres i {srcCfg.label}</h2>
      <p style={{ color: "var(--bfc-base-c-2)", marginBottom: "1.5rem", fontSize: "0.9rem" }}>
        Bruk lenken under for å se og redigere innholdet i {srcCfg.label}.
      </p>
      {plan.external_url ? (
        <a href={plan.external_url} target="_blank" rel="noopener noreferrer"
          style={{ display: "inline-block", padding: "0.6rem 1.5rem", background: srcCfg.color, color: "#fff", borderRadius: 6, fontWeight: 600, textDecoration: "none" }}>
          Åpne i {srcCfg.label} →
        </a>
      ) : (
        <p style={{ color: "var(--bfc-base-c-3)", fontSize: "0.85rem" }}>Ingen URL lagt til. Trykk «Rediger» for å legge til.</p>
      )}
    </div>
  );
}

// ─── Planner view ─────────────────────────────────────────────────────────────

function PlannerView({
  plan, srcCfg, isMsalConfigured: configured, msalReady, account, data, loading, error, onLogin, onRefresh,
}: {
  plan: ProjectPlan;
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
  const planId = plan.external_url ? parsePlanId(plan.external_url) : null;

  if (!configured) {
    return (
      <div style={{ padding: "1.5rem", borderRadius: 10, background: "#FFF3BF", border: "1px solid #FAB005", maxWidth: 600 }}>
        <h3 style={{ margin: "0 0 0.5rem", color: "#5C3A00" }}>Azure AD ikke konfigurert</h3>
        <p style={{ margin: 0, fontSize: "0.9rem", color: "#7C4D00" }}>
          Legg til <code>VITE_AZURE_CLIENT_ID</code> og <code>VITE_AZURE_TENANT_ID</code> som GitHub Secrets og bygg på nytt.
        </p>
      </div>
    );
  }

  if (!plan.external_url || !planId) {
    return (
      <div style={{ textAlign: "center", padding: "2rem", color: "var(--bfc-base-c-2)" }}>
        <p>Ingen Planner-URL er lagt til ennå.</p>
        <p style={{ fontSize: "0.85rem" }}>Trykk «Rediger» og lim inn lenken til Planner-planen.</p>
      </div>
    );
  }

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "1.25rem", flexWrap: "wrap" }}>
        {account ? (
          <>
            <span style={{ fontSize: "0.85rem", color: "var(--bfc-base-c-2)" }}>
              Innlogget som <strong>{account.name ?? account.username}</strong>
            </span>
            <Button variant="outline" onClick={onRefresh} state={loading ? "inactive" : "default"}>
              {loading ? "Henter..." : "Oppdater"}
            </Button>
            {plan.external_url && (
              <a href={plan.external_url} target="_blank" rel="noopener noreferrer"
                style={{ fontSize: "0.85rem", color: srcCfg.color, fontWeight: 600, textDecoration: "none" }}>
                Åpne i Planner →
              </a>
            )}
          </>
        ) : (
          <Button variant="filled" onClick={onLogin} state={!msalReady ? "inactive" : "default"}
            style={{ background: "#0078D4", borderColor: "#0078D4" }}>
            Logg inn med Microsoft for å hente oppgaver
          </Button>
        )}
      </div>

      {error && (
        <div style={{ padding: "1rem 1.25rem", borderRadius: 8, background: error.startsWith("PREMIUM_PLAN") ? "#FFF3BF" : "#FFE3E3", border: `1px solid ${error.startsWith("PREMIUM_PLAN") ? "#FAB005" : "#FFA8A8"}`, marginBottom: "1rem" }}>
          {error.startsWith("PREMIUM_PLAN") ? (
            <>
              <div style={{ fontWeight: 600, color: "#5C3A00", marginBottom: "0.4rem" }}>Planner Premium — Dataverse-feil</div>
              <p style={{ margin: "0 0 0.75rem", fontSize: "0.875rem", color: "#7C4D00", fontFamily: "monospace", wordBreak: "break-word" }}>
                {error.replace(/^PREMIUM_PLAN:\s*/, "")}
              </p>
              {plan.external_url && (
                <a href={plan.external_url} target="_blank" rel="noopener noreferrer"
                  style={{ display: "inline-block", padding: "0.45rem 1.1rem", background: "#0078D4", color: "#fff", borderRadius: 5, fontWeight: 600, textDecoration: "none", fontSize: "0.875rem" }}>
                  Åpne plan i Planner →
                </a>
              )}
            </>
          ) : (
            <span style={{ color: "#C92A2A", fontSize: "0.9rem" }}>{error}</span>
          )}
        </div>
      )}

      {loading && <div style={{ textAlign: "center", padding: "3rem", color: "var(--bfc-base-c-2)" }}>Henter oppgaver fra Microsoft Planner…</div>}

      {data && !loading && <PlannerTaskGrid data={data} />}
    </div>
  );
}

// ─── Planner task grid ────────────────────────────────────────────────────────

function PlannerTaskGrid({ data }: { data: PlannerData }) {
  const [filter, setFilter] = useState<Filter>("all");

  const bucketMap = Object.fromEntries(data.buckets.map((b) => [b.id, b.name]));

  // All tasks grouped by bucket (unfiltered, for counts/progress)
  const allGrouped: Record<string, PlannerTask[]> = {};
  for (const task of data.tasks) {
    const bucket = bucketMap[task.bucketId] ?? "Ukjent";
    if (!allGrouped[bucket]) allGrouped[bucket] = [];
    allGrouped[bucket].push(task);
  }

  // Ordered bucket names: follow API bucket order, then any orphaned groups last
  const orderedBucketNames = [
    ...data.buckets.map((b) => b.name).filter((n) => allGrouped[n]),
    ...Object.keys(allGrouped).filter((n) => !data.buckets.find((b) => b.name === n)),
  ];

  const total = data.tasks.length;
  const done = data.tasks.filter((t) => t.percentComplete === 100).length;
  const inProg = data.tasks.filter((t) => t.percentComplete > 0 && t.percentComplete < 100).length;
  const remaining = total - done - inProg;
  const overallPct = total > 0 ? Math.round(data.tasks.reduce((s, t) => s + t.percentComplete, 0) / total) : 0;

  const visibleTasks = data.tasks.filter((t) => matchPlannerFilter(t.percentComplete, filter));

  const STAT_CARDS = [
    { key: "all"         as Filter, label: "Totalt",   value: total,     color: "#868E96" },
    { key: "done"        as Filter, label: "Ferdig",   value: done,      color: "#2F9E44" },
    { key: "in_progress" as Filter, label: "Pågående", value: inProg,    color: "#1971C2" },
    { key: "remaining"   as Filter, label: "Gjenstår", value: remaining, color: "#F76707" },
  ];

  return (
    <div>
      {/* Stats */}
      <div style={{ display: "flex", gap: "1rem", marginBottom: "1rem", flexWrap: "wrap", alignItems: "center" }}>
        {STAT_CARDS.map(({ key, label, value, color }) => {
          const isActive = filter === key;
          return (
            <div
              key={key}
              onClick={() => setFilter((prev) => (prev === key ? "all" : key))}
              style={{
                padding: "0.6rem 1.1rem", borderRadius: 8, textAlign: "center", minWidth: 80,
                background: isActive ? `${color}22` : `${color}12`,
                border: `1px solid ${isActive ? color : `${color}30`}`,
                cursor: key === "all" ? "default" : "pointer",
                outline: isActive && key !== "all" ? `2px solid ${color}` : "none",
                outlineOffset: 1,
                transition: "all 0.15s",
              }}
            >
              <div style={{ fontSize: "1.4rem", fontWeight: 700, color }}>{value}</div>
              <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)" }}>{label}</div>
            </div>
          );
        })}
        <div style={{ flex: 1, minWidth: 200 }}>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.8rem", color: "var(--bfc-base-c-2)", marginBottom: "0.3rem" }}>
            <span>Samlet fremdrift</span><span>{overallPct}%</span>
          </div>
          <div style={{ height: 8, borderRadius: 4, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
            <div style={{ width: `${overallPct}%`, height: "100%", background: overallPct === 100 ? "#2F9E44" : "#1971C2", borderRadius: 4, transition: "width 0.3s" }} />
          </div>
        </div>
      </div>

      {/* Active filter indicator */}
      {filter !== "all" && (
        <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "1rem" }}>
          <span style={{ fontSize: "0.8rem", color: "var(--bfc-base-c-2)" }}>
            Viser: <strong>{FILTER_LABELS[filter]}</strong> ({visibleTasks.length} oppgaver)
          </span>
          <button
            onClick={() => setFilter("all")}
            style={{ background: "none", border: "1px solid var(--bfc-base-dimmed)", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "1px 8px", borderRadius: 4, fontSize: "0.78rem" }}
          >
            × Fjern filter
          </button>
        </div>
      )}

      {/* Buckets */}
      <div style={{ display: "grid", gap: "1.5rem" }}>
        {orderedBucketNames.map((bucket) => {
          const allTasks = allGrouped[bucket] ?? [];
          const tasks = visibleTasks.filter((t) => (bucketMap[t.bucketId] ?? "Ukjent") === bucket);
          if (tasks.length === 0 && filter !== "all") return null;
          const avg = allTasks.length > 0 ? Math.round(allTasks.reduce((s, t) => s + t.percentComplete, 0) / allTasks.length) : 0;
          const bucketDone = allTasks.filter((t) => t.percentComplete === 100).length;
          return (
            <div key={bucket}>
              <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "0.5rem" }}>
                <h3 className="bf-h4" style={{ margin: 0, flex: 1 }}>{bucket}</h3>
                <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>
                  {bucketDone}/{allTasks.length} ferdig · {avg}%
                  {filter !== "all" && tasks.length < allTasks.length && ` (viser ${tasks.length})`}
                </span>
                <div style={{ width: 80, height: 6, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
                  <div style={{ width: `${avg}%`, height: "100%", background: avg === 100 ? "#2F9E44" : "#1971C2", borderRadius: 3, transition: "width 0.3s" }} />
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
  const pct = task.percentComplete;
  const isDone = pct === 100;
  const color = isDone ? "#2F9E44" : pct > 0 ? "#1971C2" : "#868E96";

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
      <div style={{
        width: 24, height: 24, borderRadius: "50%", flexShrink: 0,
        border: `2px solid ${color}`, background: isDone ? color : "transparent",
        display: "flex", alignItems: "center", justifyContent: "center",
        fontSize: "0.65rem", fontWeight: 700, color: isDone ? "#fff" : pct > 0 ? color : "transparent",
      }}>
        {isDone ? "✓" : pct > 0 ? `${pct}` : ""}
      </div>

      <span style={{
        flex: 1, fontSize: "0.9rem",
        textDecoration: isDone ? "line-through" : "none",
        color: isDone ? "var(--bfc-base-c-3)" : "inherit",
        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
      }}>
        {task.title}
      </span>

      {/* Labels */}
      {task.labels && task.labels.length > 0 && (
        <div style={{ display: "flex", gap: "0.3rem", flexShrink: 0, flexWrap: "wrap", maxWidth: 200 }}>
          {task.labels.map((label) => (
            <span key={label} style={{
              fontSize: "0.7rem", fontWeight: 600,
              padding: "2px 7px", borderRadius: 20,
              background: "#7950F218", color: "#7950F2",
              whiteSpace: "nowrap",
            }}>
              {label}
            </span>
          ))}
        </div>
      )}

      {(task.startDateTime || task.dueDateTime) && (
        <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)", flexShrink: 0, whiteSpace: "nowrap" }}>
          {task.startDateTime && new Date(task.startDateTime).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
          {task.startDateTime && task.dueDateTime && " → "}
          {task.dueDateTime && new Date(task.dueDateTime).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
        </span>
      )}

      <div style={{ width: 64, flexShrink: 0 }}>
        <div style={{ display: "flex", justifyContent: "flex-end", fontSize: "0.7rem", color, marginBottom: 2 }}>{pct}%</div>
        <div style={{ height: 4, borderRadius: 2, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
          <div style={{ width: `${pct}%`, height: "100%", background: color, borderRadius: 2 }} />
        </div>
      </div>
    </div>
  );
}
