import { useEffect, useState, useCallback } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type ProjectPlan, type ProjectPlanTask, type Template } from "../api/client";
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

// ─── Gantt types ──────────────────────────────────────────────────────────────

interface GanttPhase {
  name: string;
  start: Date;
  end: Date;
  color: string;
  taskCount: number;
  done: number;
}

const GANTT_COLORS = [
  "#7950F2",
  "#0CA678",
  "#F59F00",
  "#2F9E44",
  "#1971C2",
  "#C2255C",
  "#E8590C",
  "#6741D9",
];

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

  const [templateModal, setTemplateModal] = useState(false);
  const [templateMode, setTemplateMode] = useState<"new" | "existing">("new");
  const [templateName, setTemplateName] = useState("");
  const [existingTemplates, setExistingTemplates] = useState<Template[]>([]);
  const [selectedExistingId, setSelectedExistingId] = useState("");
  const [templateSaving, setTemplateSaving] = useState(false);
  const [view, setView] = useState<"list" | "gantt">("list");

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

  async function openTemplateModal() {
    const ts = await api.templates.list("project_plan");
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
      const data = JSON.stringify({ tasks: plan.tasks.map(({ name, bucket, percent_complete, start_date, end_date, responsible, description }) => ({ name, bucket, percent_complete, start_date, end_date, responsible, description })) });
      if (templateMode === "new") {
        await api.templates.create({ name: templateName.trim() || plan.title, type: "project_plan", data });
      } else {
        await api.templates.update(selectedExistingId, { name: existingTemplates.find(t => t.id === selectedExistingId)?.name ?? templateName, data });
      }
      setTemplateModal(false);
    } finally {
      setTemplateSaving(false);
    }
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
            <>
              <Button variant="outline" onClick={openTemplateModal}>Lagre som mal</Button>
              <Button variant="filled" onClick={() => openAdd()}>+ Legg til oppgave</Button>
            </>
          )}
        </div>
      </div>

      {/* Tab switcher */}
      {plan.source !== "smartsheet" && (
        <div style={{ display: "flex", borderBottom: "2px solid var(--bfc-base-dimmed)", marginBottom: "1.5rem" }}>
          {(["list", "gantt"] as const).map((v) => (
            <button
              key={v}
              onClick={() => setView(v)}
              style={{
                background: "none", border: "none", cursor: "pointer",
                padding: "0.5rem 1.25rem", fontSize: "0.9rem",
                fontWeight: view === v ? 600 : 400,
                color: view === v ? "#7950F2" : "var(--bfc-base-c-2)",
                borderBottom: `2px solid ${view === v ? "#7950F2" : "transparent"}`,
                marginBottom: -2,
                transition: "all 0.15s",
              }}
            >
              {v === "list" ? "Oppgaver" : "Gantt"}
            </button>
          ))}
        </div>
      )}

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
          view={view}
        />
      )}

      {/* Smartsheet */}
      {plan.source === "smartsheet" && (
        <ExternalLinkCard plan={plan} srcCfg={srcCfg} />
      )}

      {/* Own plan */}
      {plan.source === "own" && view === "list" && (
        <OwnPlanView
          plan={plan}
          onAdd={openAdd}
          onEdit={openEdit}
          onDelete={setDeleteTarget}
          onUpdateProgress={updateTaskProgress}
        />
      )}
      {plan.source === "own" && view === "gantt" && (
        <GanttView phases={extractOwnGanttPhases(plan.tasks)} />
      )}

      {/* Template modal */}
      <Modal isOpen={templateModal} onRequestClose={() => setTemplateModal(false)} header="Lagre som mal">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.5rem" }}>
            {(["new", "existing"] as const).map((mode) => (
              <button key={mode} onClick={() => setTemplateMode(mode)}
                style={{ padding: "0.75rem", borderRadius: 8, border: `2px solid ${templateMode === mode ? "#0CA678" : "var(--bfc-base-dimmed)"}`, background: templateMode === mode ? "#0CA67818" : "var(--bfc-base-3)", cursor: "pointer", fontWeight: 600, fontSize: "0.9rem", color: templateMode === mode ? "#0CA678" : "var(--bfc-base-c-1)", transition: "all 0.15s" }}>
                {mode === "new" ? "Ny mal" : "Oppdater eksisterende"}
              </button>
            ))}
          </div>
          {templateMode === "new" ? (
            <Input label="Navn på malen" value={templateName} onChange={(e) => setTemplateName(e.target.value)} placeholder="f.eks. Standard prosjektplan" autoFocus />
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
          <p style={{ margin: 0, color: "var(--bfc-base-c-2)", fontSize: "0.85rem" }}>{plan.tasks.length} oppgaver vil bli lagret i malen.</p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setTemplateModal(false)}>Avbryt</Button>
            <Button variant="filled" onClick={saveAsTemplate} state={templateSaving || (templateMode === "existing" && !selectedExistingId) ? "inactive" : "default"}>
              {templateSaving ? "Lagrer..." : "Lagre mal"}
            </Button>
          </div>
        </div>
      </Modal>

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
  plan, srcCfg, isMsalConfigured: configured, msalReady, account, data, loading, error, onLogin, onRefresh, view,
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
  view: "list" | "gantt";
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

      {data && !loading && (
        view === "gantt"
          ? <GanttView phases={extractPlannerGanttPhases(data)} />
          : <PlannerTaskGrid data={data} />
      )}
    </div>
  );
}

// ─── Planner task grid ────────────────────────────────────────────────────────

function PlannerTaskGrid({ data }: { data: PlannerData }) {
  const [filter, setFilter] = useState<Filter>("all");
  const [fagFilter, setFagFilter] = useState<string | null>(null);

  const bucketMap = Object.fromEntries(data.buckets.map((b) => [b.id, b.name]));
  const taskMap = Object.fromEntries(data.tasks.map((t) => [t.id, t]));

  // Hierarchical = Planner Premium plan with L1 (fase) / L2 (leveranse) structure
  const isHierarchical = data.tasks.some((t) => (t.outlineLevel ?? 1) > 1);

  function getFagomrade(task: PlannerTask): string {
    if (bucketMap[task.bucketId]) return bucketMap[task.bucketId];
    if (task.parentTaskId) {
      const parent = taskMap[task.parentTaskId];
      if (parent && bucketMap[parent.bucketId]) return bucketMap[parent.bucketId];
    }
    return "";
  }

  // Stats count: L2 leveranser in hierarchical mode, all tasks in flat mode
  const statSource = isHierarchical
    ? data.tasks.filter((t) => (t.outlineLevel ?? 1) === 2)
    : data.tasks;

  const total = statSource.length;
  const done = statSource.filter((t) => t.percentComplete === 100).length;
  const inProg = statSource.filter((t) => t.percentComplete > 0 && t.percentComplete < 100).length;
  const remaining = total - done - inProg;
  const overallPct = total > 0 ? Math.round(statSource.reduce((s, t) => s + t.percentComplete, 0) / total) : 0;
  const visibleStats = statSource.filter((t) => matchPlannerFilter(t.percentComplete, filter));

  const STAT_CARDS = [
    { key: "all"         as Filter, label: "Totalt",   value: total,     color: "#868E96" },
    { key: "done"        as Filter, label: "Ferdig",   value: done,      color: "#2F9E44" },
    { key: "in_progress" as Filter, label: "Pågående", value: inProg,    color: "#1971C2" },
    { key: "remaining"   as Filter, label: "Gjenstår", value: remaining, color: "#F76707" },
  ];

  // ── Hierarchical data setup ──────────────────────────────────────────────────
  const l1Tasks = isHierarchical
    ? data.tasks.filter((t) => (t.outlineLevel ?? 1) === 1)
    : [];

  const l2ByL1: Record<string, PlannerTask[]> = {};
  if (isHierarchical) {
    const l2Tasks = data.tasks.filter((t) => (t.outlineLevel ?? 1) === 2);
    for (const t of l2Tasks) {
      const key = t.parentTaskId ?? "__orphan__";
      if (!l2ByL1[key]) l2ByL1[key] = [];
      l2ByL1[key].push(t);
    }
  }

  // ── Flat data setup ──────────────────────────────────────────────────────────
  const allGrouped: Record<string, PlannerTask[]> = {};
  if (!isHierarchical) {
    for (const task of data.tasks) {
      const bucket = bucketMap[task.bucketId] ?? "Ukjent";
      if (!allGrouped[bucket]) allGrouped[bucket] = [];
      allGrouped[bucket].push(task);
    }
  }
  const orderedBucketNames = isHierarchical ? [] : [
    ...data.buckets.map((b) => b.name).filter((n) => allGrouped[n]),
    ...Object.keys(allGrouped).filter((n) => !data.buckets.find((b) => b.name === n)),
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

      {/* Active status filter indicator */}
      {filter !== "all" && (
        <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "0.75rem" }}>
          <span style={{ fontSize: "0.8rem", color: "var(--bfc-base-c-2)" }}>
            Viser: <strong>{FILTER_LABELS[filter]}</strong> ({visibleStats.length} {isHierarchical ? "leveranser" : "oppgaver"})
          </span>
          <button
            onClick={() => setFilter("all")}
            style={{ background: "none", border: "1px solid var(--bfc-base-dimmed)", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "1px 8px", borderRadius: 4, fontSize: "0.78rem" }}
          >
            × Fjern filter
          </button>
        </div>
      )}

      {/* Fagområde filter chips (hierarchical Premium only) */}
      {isHierarchical && data.buckets.length > 0 && (
        <div style={{ display: "flex", gap: "0.4rem", marginBottom: "1.25rem", flexWrap: "wrap", alignItems: "center" }}>
          <span style={{ fontSize: "0.78rem", color: "var(--bfc-base-c-2)", marginRight: "0.2rem" }}>Fagområde:</span>
          {data.buckets.map((b) => {
            const isActive = fagFilter === b.name;
            return (
              <button
                key={b.id}
                onClick={() => setFagFilter((prev) => (prev === b.name ? null : b.name))}
                style={{
                  padding: "3px 12px", borderRadius: 20, cursor: "pointer", fontSize: "0.78rem", fontWeight: 600,
                  border: `1px solid ${isActive ? "#0078D4" : "#0078D430"}`,
                  background: isActive ? "#0078D4" : "#0078D408",
                  color: isActive ? "#fff" : "#0078D4",
                  transition: "all 0.15s",
                }}
              >
                {b.name}
              </button>
            );
          })}
          {fagFilter && (
            <button
              onClick={() => setFagFilter(null)}
              style={{ background: "none", border: "1px solid var(--bfc-base-dimmed)", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "2px 8px", borderRadius: 4, fontSize: "0.78rem", marginLeft: "0.25rem" }}
            >
              × Fjern
            </button>
          )}
        </div>
      )}

      <div style={{ display: "grid", gap: "1.5rem" }}>
        {isHierarchical ? (
          // ── Hierarchical: L1 (fase) as section headers, L2 (leveranse) as rows with fagomrade chip
          <>
            {l1Tasks.map((l1) => {
              const allL2 = l2ByL1[l1.id] ?? [];
              const visibleL2 = allL2.filter((t) =>
                matchPlannerFilter(t.percentComplete, filter) &&
                (!fagFilter || getFagomrade(t) === fagFilter)
              );
              if (visibleL2.length === 0 && (filter !== "all" || fagFilter !== null)) return null;
              const avg = allL2.length > 0 ? Math.round(allL2.reduce((s, t) => s + t.percentComplete, 0) / allL2.length) : 0;
              const sectionDone = allL2.filter((t) => t.percentComplete === 100).length;
              return (
                <div key={l1.id}>
                  <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "0.5rem" }}>
                    <h3 className="bf-h4" style={{ margin: 0, flex: 1 }}>{l1.title}</h3>
                    <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>
                      {sectionDone}/{allL2.length} ferdig · {avg}%
                      {filter !== "all" && visibleL2.length < allL2.length && ` (viser ${visibleL2.length})`}
                    </span>
                    <div style={{ width: 80, height: 6, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
                      <div style={{ width: `${avg}%`, height: "100%", background: avg === 100 ? "#2F9E44" : "#1971C2", borderRadius: 3, transition: "width 0.3s" }} />
                    </div>
                  </div>
                  <div style={{ display: "grid", gap: "0.35rem" }}>
                    {visibleL2.map((t) => <PlannerTaskRow key={t.id} task={t} fagomrade={getFagomrade(t)} />)}
                  </div>
                </div>
              );
            })}
            {(l2ByL1["__orphan__"] ?? []).length > 0 && (() => {
              const visible = (l2ByL1["__orphan__"] ?? []).filter((t) => matchPlannerFilter(t.percentComplete, filter));
              if (visible.length === 0 && filter !== "all") return null;
              return (
                <div>
                  <div style={{ display: "flex", alignItems: "center", marginBottom: "0.5rem" }}>
                    <h3 className="bf-h4" style={{ margin: 0, flex: 1, color: "var(--bfc-base-c-2)" }}>Uten fase</h3>
                  </div>
                  <div style={{ display: "grid", gap: "0.35rem" }}>
                    {visible.map((t) => <PlannerTaskRow key={t.id} task={t} fagomrade={getFagomrade(t)} />)}
                  </div>
                </div>
              );
            })()}
          </>
        ) : (
          // ── Flat: group by bucket (Basic Planner / flat Premium)
          <>
            {orderedBucketNames.map((bucket) => {
              const allTasks = allGrouped[bucket] ?? [];
              const tasks = allTasks.filter((t) => matchPlannerFilter(t.percentComplete, filter));
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
          </>
        )}
      </div>
    </div>
  );
}

function PlannerTaskRow({ task, fagomrade }: { task: PlannerTask; fagomrade?: string }) {
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

      {/* Fagområde chip (hierarchical Premium) or labels (Basic Planner) */}
      {fagomrade ? (
        <span style={{
          fontSize: "0.7rem", fontWeight: 600, flexShrink: 0,
          padding: "2px 8px", borderRadius: 20,
          background: "#0078D418", color: "#0078D4",
          whiteSpace: "nowrap",
        }}>
          {fagomrade}
        </span>
      ) : task.labels && task.labels.length > 0 ? (
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
      ) : null}

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

// ─── Gantt helpers ────────────────────────────────────────────────────────────

function extractOwnGanttPhases(tasks: ProjectPlanTask[]): GanttPhase[] {
  const buckets = [...new Set(tasks.map((t) => t.bucket ?? ""))];
  const phases: GanttPhase[] = [];
  buckets.forEach((bucket, i) => {
    const bTasks = tasks.filter((t) => (t.bucket ?? "") === bucket);
    const starts = bTasks.filter((t) => t.start_date).map((t) => new Date(t.start_date!).getTime());
    const ends = bTasks.filter((t) => t.end_date).map((t) => new Date(t.end_date!).getTime());
    if (starts.length === 0 && ends.length === 0) return;
    const minStart = starts.length > 0 ? Math.min(...starts) : Math.min(...ends);
    const maxEnd = ends.length > 0 ? Math.max(...ends) : Math.max(...starts);
    phases.push({
      name: bucket || "Uten fase",
      start: new Date(minStart),
      end: new Date(maxEnd),
      color: GANTT_COLORS[i % GANTT_COLORS.length],
      taskCount: bTasks.length,
      done: bTasks.filter((t) => t.percent_complete === 100).length,
    });
  });
  return phases;
}

function extractPlannerGanttPhases(data: PlannerData): GanttPhase[] {
  const isHierarchical = data.tasks.some((t) => (t.outlineLevel ?? 1) > 1);
  const phases: GanttPhase[] = [];

  if (isHierarchical) {
    const l1Tasks = data.tasks.filter((t) => (t.outlineLevel ?? 1) === 1);
    const l2ByL1: Record<string, PlannerTask[]> = {};
    for (const t of data.tasks.filter((t) => (t.outlineLevel ?? 1) === 2)) {
      const key = t.parentTaskId ?? "__orphan__";
      if (!l2ByL1[key]) l2ByL1[key] = [];
      l2ByL1[key].push(t);
    }
    l1Tasks.forEach((l1, i) => {
      const children = l2ByL1[l1.id] ?? [];

      // Use the phase's own dates if set — these match what Planner shows on the phase header row
      if (l1.startDateTime && l1.dueDateTime) {
        phases.push({
          name: l1.title,
          start: new Date(l1.startDateTime),
          end: new Date(l1.dueDateTime),
          color: GANTT_COLORS[i % GANTT_COLORS.length],
          taskCount: children.length,
          done: children.filter((t) => t.percentComplete === 100).length,
        });
        return;
      }

      // Fallback: derive range from child task dates
      const starts = children.filter((t) => t.startDateTime).map((t) => new Date(t.startDateTime!).getTime());
      const ends = children.filter((t) => t.dueDateTime).map((t) => new Date(t.dueDateTime!).getTime());
      if (starts.length === 0 && ends.length === 0) return;
      const minStart = starts.length > 0 ? Math.min(...starts) : Math.min(...ends);
      const maxEnd = ends.length > 0 ? Math.max(...ends) : Math.max(...starts);
      phases.push({
        name: l1.title,
        start: new Date(minStart),
        end: new Date(maxEnd),
        color: GANTT_COLORS[i % GANTT_COLORS.length],
        taskCount: children.length,
        done: children.filter((t) => t.percentComplete === 100).length,
      });
    });
  } else {
    data.buckets.forEach((bucket, i) => {
      const bTasks = data.tasks.filter((t) => t.bucketId === bucket.id);
      const starts = bTasks.filter((t) => t.startDateTime).map((t) => new Date(t.startDateTime!).getTime());
      const ends = bTasks.filter((t) => t.dueDateTime).map((t) => new Date(t.dueDateTime!).getTime());
      if (starts.length === 0 && ends.length === 0) return;
      const minStart = starts.length > 0 ? Math.min(...starts) : Math.min(...ends);
      const maxEnd = ends.length > 0 ? Math.max(...ends) : Math.max(...starts);
      phases.push({
        name: bucket.name,
        start: new Date(minStart),
        end: new Date(maxEnd),
        color: GANTT_COLORS[i % GANTT_COLORS.length],
        taskCount: bTasks.length,
        done: bTasks.filter((t) => t.percentComplete === 100).length,
      });
    });
  }
  return phases;
}

// ─── Gantt phase sorting ──────────────────────────────────────────────────────

const PHASE_ORDER_MAP: [RegExp, number][] = [
  [/kartlegg|mapping|analyse|survey/i,           10],
  [/gjennomfør|implementer|etabler|rollout/i,    20],
  [/test|verifi|qa|validering/i,                 30],
  [/etter|post[\s\-]*go|after[\s\-]*go/i,        50], // must come before go-live check
  [/go[\s\-. ]*live|golive|launch/i,             40],
];

function phasePriority(name: string): number {
  for (const [rx, order] of PHASE_ORDER_MAP) {
    if (rx.test(name)) return order;
  }
  return 99;
}

// ─── Gantt view ───────────────────────────────────────────────────────────────

function GanttView({ phases }: { phases: GanttPhase[] }) {
  if (phases.length === 0) {
    return (
      <div style={{ textAlign: "center", padding: "3rem", border: "2px dashed var(--bfc-base-dimmed)", borderRadius: 8, color: "var(--bfc-base-c-2)" }}>
        <div style={{ fontSize: "2rem", marginBottom: "0.75rem" }}>📅</div>
        <p style={{ margin: "0 0 0.5rem" }}>Ingen oppgaver har dato satt.</p>
        <p style={{ margin: 0, fontSize: "0.85rem" }}>Legg til start- og sluttdato på oppgavene for å se Gantt-visning.</p>
      </div>
    );
  }

  // Sort phases by defined project phase order, reassign colors to match position
  const sorted = [...phases]
    .sort((a, b) => phasePriority(a.name) - phasePriority(b.name))
    .map((p, i) => ({ ...p, color: GANTT_COLORS[i % GANTT_COLORS.length] }));

  const MS_PER_DAY = 86_400_000;
  const pad = 2 * MS_PER_DAY;
  const rangeStartMs = Math.min(...sorted.map((p) => p.start.getTime())) - pad;
  const rangeEndMs = Math.max(...sorted.map((p) => p.end.getTime())) + pad;
  const totalMs = rangeEndMs - rangeStartMs;
  const totalDays = Math.ceil(totalMs / MS_PER_DAY);
  const totalWeeks = Math.ceil(totalDays / 7);

  function toPct(d: Date): number {
    return ((d.getTime() - rangeStartMs) / totalMs) * 100;
  }

  // Adaptive tick interval: fewer ticks for longer ranges to avoid label crowding
  const tickIntervalDays = totalDays <= 42 ? 7 : totalDays <= 98 ? 14 : 28;

  const ticks: Date[] = [];
  const seed = new Date(rangeStartMs);
  if (tickIntervalDays === 7) {
    // Align to nearest Monday
    const dow = seed.getDay();
    seed.setDate(seed.getDate() + (dow === 1 ? 7 : (8 - dow) % 7 || 7));
  } else {
    seed.setDate(seed.getDate() + tickIntervalDays);
  }
  let tick = new Date(seed);
  while (tick.getTime() <= rangeEndMs) {
    ticks.push(new Date(tick));
    tick.setDate(tick.getDate() + tickIntervalDays);
  }

  const minW = Math.max(700, ticks.length * 95 + 140);

  const fmt = (d: Date) => d.toLocaleDateString("nb-NO", { day: "numeric", month: "short" });
  const daysBetween = (a: Date, b: Date) => Math.max(1, Math.ceil((b.getTime() - a.getTime()) / MS_PER_DAY));

  return (
    <div>
      {/* Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.25rem" }}>
        <h3 className="bf-h4" style={{ margin: 0 }}>Tidslinje</h3>
        <span style={{ fontSize: "0.8rem", color: "var(--bfc-base-c-2)" }}>
          {fmt(sorted.reduce((a, p) => p.start < a ? p.start : a, sorted[0].start))} – {fmt(sorted.reduce((a, p) => p.end > a ? p.end : a, sorted[0].end))} · {totalWeeks} {totalWeeks === 1 ? "uke" : "uker"}
        </span>
      </div>

      {/* Timeline */}
      <div style={{ overflowX: "auto", borderRadius: 10, border: "1px solid var(--bfc-base-dimmed)", background: "var(--bfc-base-3)", padding: "1rem 1rem 0.75rem" }}>
        <div style={{ minWidth: minW }}>
          {/* Date axis row */}
          <div style={{ display: "flex" }}>
            <div style={{ width: 140, flexShrink: 0 }} />
            <div style={{ flex: 1, position: "relative", height: 22, marginBottom: "0.75rem" }}>
              {ticks.map((t, i) => (
                <span key={i} style={{
                  position: "absolute",
                  left: `${toPct(t)}%`,
                  transform: "translateX(-50%)",
                  fontSize: "0.75rem",
                  color: "var(--bfc-base-c-2)",
                  whiteSpace: "nowrap",
                  userSelect: "none",
                }}>
                  {fmt(t)}
                </span>
              ))}
            </div>
          </div>

          {/* Phase rows */}
          {sorted.map((phase) => {
            const leftPct = toPct(phase.start);
            const widthPct = Math.max(1.5, toPct(phase.end) - leftPct);
            return (
              <div key={phase.name} style={{ display: "flex", alignItems: "center", marginBottom: "0.5rem" }}>
                {/* Phase label */}
                <div style={{
                  width: 140, flexShrink: 0,
                  fontSize: "0.82rem", fontWeight: 600,
                  color: "var(--bfc-base-c-1)",
                  paddingRight: "0.75rem",
                  overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                }}>
                  {phase.name}
                </div>
                {/* Bar track */}
                <div style={{ flex: 1, position: "relative", height: 36 }}>
                  {/* Grid lines */}
                  {ticks.map((t, i) => (
                    <div key={i} style={{
                      position: "absolute",
                      left: `${toPct(t)}%`,
                      top: 0, bottom: 0, width: 1,
                      background: "var(--bfc-base-dimmed)",
                      pointerEvents: "none",
                    }} />
                  ))}
                  {/* Bar */}
                  <div style={{
                    position: "absolute",
                    left: `${leftPct}%`,
                    width: `${widthPct}%`,
                    top: 3, bottom: 3,
                    borderRadius: 6,
                    background: `${phase.color}20`,
                    border: `2px solid ${phase.color}`,
                    display: "flex", alignItems: "center",
                    paddingLeft: "0.5rem",
                    overflow: "hidden",
                    zIndex: 1,
                  }}>
                    <span style={{ fontSize: "0.71rem", fontWeight: 600, color: phase.color, whiteSpace: "nowrap" }}>
                      {fmt(phase.start)} → {fmt(phase.end)}
                    </span>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Summary cards */}
      <div style={{
        display: "grid",
        gridTemplateColumns: `repeat(${Math.min(sorted.length, 4)}, 1fr)`,
        gap: "0.75rem",
        marginTop: "1.5rem",
      }}>
        {sorted.map((phase) => {
          const days = daysBetween(phase.start, phase.end);
          const weeks = Math.floor(days / 7);
          const rem = days % 7;
          const durationStr = weeks > 0
            ? `${weeks} ${weeks === 1 ? "uke" : "uker"}${rem > 0 ? `, ${rem} dager` : ""}`
            : `${days} dager`;
          return (
            <div key={phase.name} style={{
              borderRadius: 10,
              border: `1px solid ${phase.color}40`,
              borderTop: `4px solid ${phase.color}`,
              padding: "1rem",
              background: "var(--bfc-base-3)",
            }}>
              <div style={{ fontSize: "0.68rem", fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: phase.color, marginBottom: "0.5rem" }}>
                {phase.name}
              </div>
              <div style={{ fontSize: "1.5rem", fontWeight: 700, color: "var(--bfc-base-c-1)", lineHeight: 1.2 }}>
                {days} dager
              </div>
              <div style={{ fontSize: "0.8rem", color: "var(--bfc-base-c-2)", marginTop: "0.2rem", marginBottom: "0.5rem" }}>
                {durationStr}
              </div>
              <div style={{ fontSize: "0.8rem", color: "var(--bfc-base-c-2)" }}>
                {fmt(phase.start)}<br />→ {fmt(phase.end)}
              </div>
              {phase.taskCount > 0 && (
                <div style={{ fontSize: "0.73rem", color: "var(--bfc-base-c-3)", marginTop: "0.5rem" }}>
                  {phase.done}/{phase.taskCount} ferdig
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
