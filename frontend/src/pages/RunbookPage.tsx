import React, { useEffect, useState, useCallback } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type Project, type Runbook, type RunbookActivity, type Template } from "../api/client";
import { exportRunbookPdf, exportRunbookExcel } from "../utils/exportUtils";
import { isMsalConfigured, msalInstance, PLANNER_SCOPES } from "../auth/msalConfig";
import { fetchPlannerData, parsePlanId, taskStatus, togglePlannerTask, type PlannerData, type PlannerTask } from "../auth/plannerService";
import { probeOnboardApi } from "../auth/onboardService";
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

type Filter = "all" | "done" | "in_progress" | "remaining";

const FILTER_LABELS: Record<Filter, string> = {
  all:         "Alle",
  done:        "Ferdig",
  in_progress: "Pågående",
  remaining:   "Gjenstår",
};

function matchActivityFilter(status: RunbookActivity["status"], f: Filter): boolean {
  if (f === "done")        return status === "done";
  if (f === "in_progress") return status === "in_progress";
  if (f === "remaining")   return status !== "done" && status !== "in_progress";
  return true;
}

function matchPlannerFilter(pct: number, f: Filter): boolean {
  if (f === "done")        return pct === 100;
  if (f === "in_progress") return pct > 0 && pct < 100;
  if (f === "remaining")   return pct === 0;
  return true;
}

type Tab = "dashboard" | "aktiviteter" | "pcer" | "lokasjoner" | "applikasjoner";
const TABS: { id: Tab; label: string }[] = [
  { id: "dashboard",     label: "Dashboard" },
  { id: "aktiviteter",   label: "Aktiviteter" },
  { id: "pcer",          label: "PCer" },
  { id: "lokasjoner",    label: "Lokasjoner" },
  { id: "applikasjoner", label: "Applikasjoner" },
];

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
  const [project, setProject] = useState<Project | null>(null);
  const [activeTab, setActiveTab] = useState<Tab>("dashboard");

  // Planner state — MSAL is initialized in main.tsx before React renders
  const [msalReady] = useState(isMsalConfigured);
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

  const [ownFilter, setOwnFilter] = useState<Filter>("all");

  const [editRunbookModal, setEditRunbookModal] = useState(false);
  const [runbookTitle, setRunbookTitle] = useState("");
  const [runbookUrl, setRunbookUrl] = useState("");

  const [templateModal, setTemplateModal] = useState(false);
  const [templateMode, setTemplateMode] = useState<"new" | "existing">("new");
  const [templateName, setTemplateName] = useState("");
  const [existingTemplates, setExistingTemplates] = useState<Template[]>([]);
  const [selectedExistingId, setSelectedExistingId] = useState("");
  const [templateSaving, setTemplateSaving] = useState(false);

  useEffect(() => {
    if (!projectId || !runbookId) return;
    api.runbooks.get(projectId, runbookId).then(setRunbook);
    api.projects.get(projectId).then(setProject);
  }, [projectId, runbookId]);

  // Check for existing MSAL login on mount (including after loginRedirect return)
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

  // Auto-fetch when account + runbook ready
  useEffect(() => {
    if (!plannerAccount || !runbook || runbook.source !== "planner" || !runbook.external_url) return;
    loadPlannerData(plannerAccount, runbook.external_url);
  }, [plannerAccount, runbook, loadPlannerData]);

  async function loginPlanner() {
    if (!isMsalConfigured || !msalReady) return;
    Object.keys(sessionStorage)
      .filter((k) => k.includes("interaction"))
      .forEach((k) => sessionStorage.removeItem(k));
    sessionStorage.setItem("app.returnUrl", window.location.href);
    await msalInstance.loginRedirect({ scopes: PLANNER_SCOPES });
  }

  function refreshPlanner() {
    if (!plannerAccount || !runbook?.external_url) return;
    loadPlannerData(plannerAccount, runbook.external_url);
  }

  async function handleToggleTask(taskId: string, done: boolean): Promise<void> {
    if (!plannerData || !plannerAccount || !runbook || !runbook.external_url) return;
    setPlannerData((prev) => prev ? {
      ...prev,
      tasks: prev.tasks.map((t) => t.id === taskId ? { ...t, percentComplete: done ? 100 : 0 } : t),
    } : null);
    try {
      await togglePlannerTask(msalInstance, plannerAccount, plannerData, taskId, done);
    } catch (err) {
      setPlannerError(err instanceof Error ? err.message : "Kunne ikke oppdatere oppgave i Planner");
      loadPlannerData(plannerAccount, runbook.external_url);
    }
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

  async function openTemplateModal() {
    const ts = await api.templates.list("runbook");
    setExistingTemplates(ts);
    setTemplateMode("new");
    setTemplateName(runbook?.title ?? "");
    setSelectedExistingId(ts[0]?.id ?? "");
    setTemplateModal(true);
  }

  async function saveAsTemplate() {
    if (!runbook) return;
    setTemplateSaving(true);
    try {
      const data = JSON.stringify({ activities: runbook.activities.map(({ name, phase, status, start_date, end_date, responsible, description }) => ({ name, phase, status, start_date, end_date, responsible, description })) });
      if (templateMode === "new") {
        await api.templates.create({ name: templateName.trim() || runbook.title, type: "runbook", data });
      } else {
        await api.templates.update(selectedExistingId, { name: existingTemplates.find(t => t.id === selectedExistingId)?.name ?? templateName, data });
      }
      setTemplateModal(false);
    } finally {
      setTemplateSaving(false);
    }
  }

  if (!runbook) return <div style={{ padding: "2rem" }}>Laster...</div>;

  const srcCfg = SOURCE_CONFIG[runbook.source];
  const phases = [...new Set(runbook.activities.map((a) => a.phase ?? ""))];
  const orderedPhases = phases.filter(Boolean);
  if (runbook.activities.some((a) => !a.phase)) orderedPhases.push("");

  const done = runbook.activities.filter((a) => a.status === "done").length;
  const inProgress = runbook.activities.filter((a) => a.status === "in_progress").length;
  const total = runbook.activities.length;

  const activityCount = runbook.source === "own" ? total : (plannerData?.tasks.length ?? 0);

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
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: "1.5rem", gap: "1rem" }}>
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
        <div style={{ display: "flex", gap: "0.5rem" }}>
          {project && (<>
            <Button variant="outline" onClick={() => void exportRunbookPdf(runbook, project, plannerData)}>↓ PDF</Button>
            <Button variant="outline" onClick={() => exportRunbookExcel(runbook, project, plannerData)}>↓ Excel</Button>
          </>)}
          <Button variant="outline" onClick={() => { setRunbookTitle(runbook.title); setRunbookUrl(runbook.external_url ?? ""); setEditRunbookModal(true); }}>
            Rediger
          </Button>
        </div>
      </div>

      {/* Tab bar */}
      <div style={{ display: "flex", borderBottom: "2px solid var(--bfc-base-dimmed)", marginBottom: "1.75rem" }}>
        {TABS.map(({ id, label }) => {
          const isActive = activeTab === id;
          const count = id === "aktiviteter" ? activityCount : 0;
          return (
            <button
              key={id}
              onClick={() => setActiveTab(id)}
              style={{
                background: "none", border: "none", cursor: "pointer",
                padding: "0.6rem 1.1rem",
                fontWeight: isActive ? 600 : 400,
                color: isActive ? srcCfg.color : "var(--bfc-base-c-2)",
                borderBottom: isActive ? `2px solid ${srcCfg.color}` : "2px solid transparent",
                marginBottom: "-2px",
                fontSize: "0.9rem",
                transition: "color 0.15s",
                display: "flex", alignItems: "center", gap: "0.4rem",
              }}
            >
              {label}
              {count > 0 && (
                <span style={{
                  fontSize: "0.7rem", fontWeight: 600,
                  padding: "1px 6px", borderRadius: 20,
                  background: isActive ? `${srcCfg.color}18` : "var(--bfc-base-dimmed)",
                  color: isActive ? srcCfg.color : "var(--bfc-base-c-2)",
                }}>
                  {count}
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* Dashboard */}
      {activeTab === "dashboard" && (
        <DashboardTab
          runbook={runbook}
          plannerData={plannerData}
          plannerLoading={plannerLoading}
          srcColor={srcCfg.color}
          onGoToActiviteter={() => setActiveTab("aktiviteter")}
        />
      )}

      {/* Aktiviteter */}
      {activeTab === "aktiviteter" && (
        <>
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
              onToggleTask={handleToggleTask}
            />
          )}

          {runbook.source === "smartsheet" && (
            <ExternalLinkCard runbook={runbook} srcCfg={srcCfg} />
          )}

          {runbook.source === "own" && (
            <>
              <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", marginBottom: "1rem" }}>
                <Button variant="outline" onClick={openTemplateModal}>Lagre som mal</Button>
                <Button variant="filled" onClick={() => openAdd()}>+ Legg til aktivitet</Button>
              </div>

              {total > 0 && (
                <>
                  <div style={{ display: "flex", gap: "1rem", marginBottom: "1rem", flexWrap: "wrap" }}>
                    {([
                      { key: "all"         as Filter, label: "Totalt",   value: total,                     color: "#868E96" },
                      { key: "done"        as Filter, label: "Ferdig",   value: done,                      color: "#2F9E44" },
                      { key: "in_progress" as Filter, label: "Pågående", value: inProgress,                color: "#1971C2" },
                      { key: "remaining"   as Filter, label: "Gjenstår", value: total - done - inProgress, color: "#F76707" },
                    ]).map(({ key, label, value, color }) => {
                      const isActive = ownFilter === key;
                      return (
                        <div
                          key={key}
                          onClick={() => key !== "all" && setOwnFilter((prev) => (prev === key ? "all" : key))}
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
                  </div>
                  {ownFilter !== "all" && (
                    <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "1rem" }}>
                      <span style={{ fontSize: "0.8rem", color: "var(--bfc-base-c-2)" }}>
                        Viser: <strong>{FILTER_LABELS[ownFilter]}</strong>
                      </span>
                      <button
                        onClick={() => setOwnFilter("all")}
                        style={{ background: "none", border: "1px solid var(--bfc-base-dimmed)", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "1px 8px", borderRadius: 4, fontSize: "0.78rem" }}
                      >
                        × Fjern filter
                      </button>
                    </div>
                  )}
                </>
              )}

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
                    const allActs = runbook.activities.filter((a) => (a.phase ?? "") === phase);
                    const acts = allActs.filter((a) => matchActivityFilter(a.status, ownFilter));
                    if (acts.length === 0 && ownFilter !== "all") return null;
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
        </>
      )}

      {/* Placeholder tabs */}
      {activeTab === "pcer" && (
        <PlaceholderTab label="PCer" icon="💻" color="#1971C2" description="Her vil en liste over alle PCer i prosjektet vises — inkl. serienummer, bruker og status.">
          {plannerAccount && (
            <button
              onClick={() => probeOnboardApi(msalInstance, plannerAccount)}
              style={{
                marginTop: "1.25rem", padding: "0.5rem 1.25rem",
                borderRadius: 6, border: "1px solid #1971C2",
                background: "#1971C218", color: "#1971C2",
                fontSize: "0.82rem", fontWeight: 600, cursor: "pointer",
              }}
            >
              Probe Onboard API (sjekk konsoll)
            </button>
          )}
        </PlaceholderTab>
      )}
      {activeTab === "lokasjoner" && (
        <PlaceholderTab label="Lokasjoner" icon="📍" color="#F76707" description="Her vil en liste over alle lokasjoner i prosjektet vises — inkl. adresse og antall enheter." />
      )}
      {activeTab === "applikasjoner" && (
        <PlaceholderTab label="Applikasjoner" icon="📦" color="#7950F2" description="Her vil en liste over alle applikasjoner i prosjektet vises — inkl. versjon og status." />
      )}

      {/* Activity form modal */}
      <Modal isOpen={templateModal} onRequestClose={() => setTemplateModal(false)} header="Lagre som mal">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.5rem" }}>
            {(["new", "existing"] as const).map((mode) => (
              <button key={mode} onClick={() => setTemplateMode(mode)}
                style={{ padding: "0.75rem", borderRadius: 8, border: `2px solid ${templateMode === mode ? "#7950F2" : "var(--bfc-base-dimmed)"}`, background: templateMode === mode ? "#7950F218" : "var(--bfc-base-3)", cursor: "pointer", fontWeight: 600, fontSize: "0.9rem", color: templateMode === mode ? "#7950F2" : "var(--bfc-base-c-1)", transition: "all 0.15s" }}>
                {mode === "new" ? "Ny mal" : "Oppdater eksisterende"}
              </button>
            ))}
          </div>
          {templateMode === "new" ? (
            <Input label="Navn på malen" value={templateName} onChange={(e) => setTemplateName(e.target.value)} placeholder="f.eks. Standard cutover-runbook" autoFocus />
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
          <p style={{ margin: 0, color: "var(--bfc-base-c-2)", fontSize: "0.85rem" }}>{runbook.activities.length} aktiviteter vil bli lagret i malen.</p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setTemplateModal(false)}>Avbryt</Button>
            <Button variant="filled" onClick={saveAsTemplate} state={templateSaving || (templateMode === "existing" && !selectedExistingId) ? "inactive" : "default"}>
              {templateSaving ? "Lagrer..." : "Lagre mal"}
            </Button>
          </div>
        </div>
      </Modal>

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

// ─── Dashboard tab ────────────────────────────────────────────────────────────

function DashboardTab({ runbook, plannerData, plannerLoading, srcColor, onGoToActiviteter }: {
  runbook: Runbook;
  plannerData: PlannerData | null;
  plannerLoading: boolean;
  srcColor: string;
  onGoToActiviteter: () => void;
}) {
  let actTotal = 0, actDone = 0, actInProgress = 0;
  let phases: { name: string; done: number; total: number }[] = [];

  if (runbook.source === "own") {
    const acts = runbook.activities;
    actTotal = acts.length;
    actDone = acts.filter((a) => a.status === "done").length;
    actInProgress = acts.filter((a) => a.status === "in_progress").length;

    const phaseMap: Record<string, { done: number; total: number }> = {};
    for (const a of acts) {
      const p = a.phase ?? "Uten fase";
      if (!phaseMap[p]) phaseMap[p] = { done: 0, total: 0 };
      phaseMap[p].total++;
      if (a.status === "done") phaseMap[p].done++;
    }
    phases = Object.entries(phaseMap).map(([name, s]) => ({ name, ...s }));
  } else if (plannerData) {
    actTotal = plannerData.tasks.length;
    actDone = plannerData.tasks.filter((t) => t.percentComplete === 100).length;
    actInProgress = plannerData.tasks.filter((t) => t.percentComplete > 0 && t.percentComplete < 100).length;

    const bucketMap = Object.fromEntries(plannerData.buckets.map((b) => [b.id, b.name]));
    const phaseMap: Record<string, { done: number; total: number }> = {};
    for (const t of plannerData.tasks) {
      const p = bucketMap[t.bucketId] ?? "Ukjent";
      if (!phaseMap[p]) phaseMap[p] = { done: 0, total: 0 };
      phaseMap[p].total++;
      if (t.percentComplete === 100) phaseMap[p].done++;
    }
    phases = Object.entries(phaseMap).map(([name, s]) => ({ name, ...s }));
  }

  const overallPct = actTotal > 0 ? Math.round((actDone / actTotal) * 100) : 0;
  const actRemaining = actTotal - actDone - actInProgress;
  const noData = actTotal === 0 && !plannerLoading;

  return (
    <div style={{ display: "grid", gap: "2rem" }}>

      {/* Activity stats */}
      <div>
        <h3 className="bf-h4" style={{ margin: "0 0 0.85rem", color: "var(--bfc-base-c-2)", fontWeight: 500, fontSize: "0.8rem", textTransform: "uppercase", letterSpacing: "0.06em" }}>Aktiviteter</h3>

        {noData ? (
          <div style={{
            padding: "2rem", borderRadius: 10,
            border: "1px dashed var(--bfc-base-dimmed)",
            color: "var(--bfc-base-c-2)", fontSize: "0.875rem", textAlign: "center",
          }}>
            {runbook.source === "own" ? (
              <>
                Ingen aktiviteter ennå.{" "}
                <button
                  onClick={onGoToActiviteter}
                  style={{ background: "none", border: "none", color: srcColor, cursor: "pointer", fontWeight: 600, padding: 0, fontSize: "inherit" }}
                >
                  Gå til Aktiviteter →
                </button>
              </>
            ) : (
              <>Logg inn i Aktiviteter-fanen for å se data fra Planner.</>
            )}
          </div>
        ) : (
          <>
            {/* Stat cards */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "0.75rem", marginBottom: "1rem" }}>
              {[
                { label: "Totalt",    value: actTotal,      color: "#868E96", sub: "aktiviteter" },
                { label: "Ferdig",    value: actDone,       color: "#2F9E44", sub: `${overallPct}% fullført` },
                { label: "Pågående",  value: actInProgress, color: "#1971C2", sub: "aktive nå" },
                { label: "Gjenstår",  value: actRemaining,  color: "#F76707", sub: "ikke startet" },
              ].map(({ label, value, color, sub }) => (
                <div
                  key={label}
                  onClick={onGoToActiviteter}
                  style={{
                    padding: "1rem 1.25rem", borderRadius: 10,
                    background: `${color}10`, border: `1px solid ${color}28`,
                    cursor: "pointer", transition: "box-shadow 0.15s",
                  }}
                  onMouseEnter={(e) => (e.currentTarget.style.boxShadow = `0 2px 12px ${color}28`)}
                  onMouseLeave={(e) => (e.currentTarget.style.boxShadow = "none")}
                >
                  <div style={{ fontSize: "2rem", fontWeight: 700, color, lineHeight: 1 }}>{value}</div>
                  <div style={{ fontSize: "0.85rem", fontWeight: 600, color, marginTop: "0.25rem" }}>{label}</div>
                  <div style={{ fontSize: "0.72rem", color: "var(--bfc-base-c-3)", marginTop: "0.15rem" }}>{sub}</div>
                </div>
              ))}
            </div>

            {/* Overall progress bar */}
            <div style={{
              padding: "0.85rem 1.1rem", borderRadius: 8,
              background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)",
            }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "0.5rem" }}>
                <span style={{ fontSize: "0.85rem", fontWeight: 500 }}>Samlet fremdrift</span>
                <span style={{ fontSize: "0.9rem", fontWeight: 700, color: overallPct === 100 ? "#2F9E44" : "#1971C2" }}>
                  {overallPct}%
                </span>
              </div>
              <div style={{ height: 10, borderRadius: 5, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
                <div style={{
                  width: `${overallPct}%`, height: "100%",
                  background: overallPct === 100 ? "#2F9E44" : "#1971C2",
                  borderRadius: 5, transition: "width 0.5s",
                }} />
              </div>
            </div>
          </>
        )}

        {plannerLoading && (
          <div style={{ color: "var(--bfc-base-c-2)", fontSize: "0.85rem", marginTop: "0.75rem" }}>
            Henter data fra Planner...
          </div>
        )}
      </div>

      {/* Phase breakdown */}
      {phases.length > 0 && (
        <div>
          <h3 className="bf-h4" style={{ margin: "0 0 0.85rem", color: "var(--bfc-base-c-2)", fontWeight: 500, fontSize: "0.8rem", textTransform: "uppercase", letterSpacing: "0.06em" }}>Fremdrift per fagområde</h3>
          <div style={{ display: "grid", gap: "0.5rem" }}>
            {phases.map(({ name, done: pDone, total: phTotal }) => {
              const pct = phTotal > 0 ? Math.round((pDone / phTotal) * 100) : 0;
              const barColor = pct === 100 ? "#2F9E44" : "#1971C2";
              return (
                <div key={name} style={{
                  padding: "0.75rem 1.1rem", borderRadius: 8,
                  background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)",
                }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "0.4rem" }}>
                    <span style={{ fontSize: "0.875rem", fontWeight: 500 }}>{name}</span>
                    <div style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}>
                      <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>{pDone}/{phTotal} ferdig</span>
                      <span style={{ fontSize: "0.75rem", fontWeight: 700, color: barColor, minWidth: 32, textAlign: "right" }}>{pct}%</span>
                    </div>
                  </div>
                  <div style={{ height: 6, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
                    <div style={{ width: `${pct}%`, height: "100%", background: barColor, borderRadius: 3, transition: "width 0.4s" }} />
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Summary tiles for future tabs */}
      <div>
        <h3 className="bf-h4" style={{ margin: "0 0 0.85rem", color: "var(--bfc-base-c-2)", fontWeight: 500, fontSize: "0.8rem", textTransform: "uppercase", letterSpacing: "0.06em" }}>Oversikt</h3>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: "0.75rem" }}>
          {[
            { label: "PCer",          icon: "💻", color: "#1971C2", count: 0 },
            { label: "Lokasjoner",    icon: "📍", color: "#F76707", count: 0 },
            { label: "Applikasjoner", icon: "📦", color: "#7950F2", count: 0 },
          ].map(({ label, icon, color, count }) => (
            <div key={label} style={{
              padding: "1.5rem 1rem", borderRadius: 10,
              background: `${color}08`, border: `1px solid ${color}20`,
              textAlign: "center",
            }}>
              <div style={{ fontSize: "2rem", marginBottom: "0.5rem" }}>{icon}</div>
              <div style={{ fontSize: "1.75rem", fontWeight: 700, color }}>{count}</div>
              <div style={{ fontSize: "0.85rem", fontWeight: 500, color, marginTop: "0.2rem" }}>{label}</div>
              <div style={{ fontSize: "0.72rem", color: "var(--bfc-base-c-3)", marginTop: "0.3rem" }}>Kommer snart</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ─── Placeholder tab ──────────────────────────────────────────────────────────

function PlaceholderTab({ label, icon, color, description, children }: {
  label: string;
  icon: string;
  color: string;
  description: string;
  children?: React.ReactNode;
}) {
  return (
    <div style={{
      textAlign: "center", padding: "4rem 2rem",
      border: `2px dashed ${color}40`, borderRadius: 12,
      background: `${color}05`,
    }}>
      <div style={{ fontSize: "3.5rem", marginBottom: "1rem" }}>{icon}</div>
      <h3 className="bf-h3" style={{ margin: "0 0 0.5rem", color }}>{label}</h3>
      <p style={{ fontSize: "0.9rem", color: "var(--bfc-base-c-2)", maxWidth: 400, margin: "0 auto" }}>
        {description}
      </p>
      {children}
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
      <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "0.5rem" }}>
        <h3 className="bf-h4" style={{ margin: 0, flex: 1 }}>{phase}</h3>
        <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>
          {done}/{activities.length} ferdig
        </span>
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

      <span style={{
        flex: 1, minWidth: 0, fontSize: "0.9rem",
        textDecoration: isDone || isCancelled ? "line-through" : "none",
        color: isDone || isCancelled ? "var(--bfc-base-c-3)" : "inherit",
        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
      }}>
        {activity.name}
      </span>

      {activity.responsible && (
        <span style={{
          fontSize: "0.75rem", padding: "2px 8px", borderRadius: 20,
          background: "var(--bfc-base-dimmed)", color: "var(--bfc-base-c-2)", flexShrink: 0,
        }}>
          {activity.responsible}
        </span>
      )}

      {(activity.start_date || activity.end_date) && (
        <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)", flexShrink: 0, whiteSpace: "nowrap" }}>
          {activity.start_date && new Date(activity.start_date).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
          {activity.start_date && activity.end_date && " → "}
          {activity.end_date && new Date(activity.end_date).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
        </span>
      )}

      <span style={{
        fontSize: "0.72rem", fontWeight: 600, padding: "2px 8px", borderRadius: 20,
        background: cfg.bg, color: cfg.color, flexShrink: 0,
      }}>
        {cfg.label}
      </span>

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
  runbook, srcCfg, isMsalConfigured: configured, msalReady, account, data, loading, error, onLogin, onRefresh, onToggleTask,
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
  onToggleTask: (taskId: string, done: boolean) => Promise<void>;
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
        <div style={{ padding: "1rem 1.25rem", borderRadius: 8, background: error.startsWith("PREMIUM_PLAN") ? "#FFF3BF" : "#FFE3E3", border: `1px solid ${error.startsWith("PREMIUM_PLAN") ? "#FAB005" : "#FFA8A8"}`, marginBottom: "1rem" }}>
          {error.startsWith("PREMIUM_PLAN") ? (
            <>
              <div style={{ fontWeight: 600, color: "#5C3A00", marginBottom: "0.4rem" }}>Planner Premium — Dataverse-feil</div>
              <p style={{ margin: "0 0 0.75rem", fontSize: "0.875rem", color: "#7C4D00", fontFamily: "monospace", wordBreak: "break-word" }}>
                {error.replace(/^PREMIUM_PLAN:\s*/, "")}
              </p>
              {runbook.external_url && (
                <a href={runbook.external_url} target="_blank" rel="noopener noreferrer"
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

      {loading && (
        <div style={{ textAlign: "center", padding: "3rem", color: "var(--bfc-base-c-2)" }}>
          Henter aktiviteter fra Microsoft Planner…
        </div>
      )}

      {data && !loading && <PlannerTaskList data={data} onToggleTask={data.source === "premium" ? undefined : onToggleTask} />}
    </div>
  );
}

// ─── Planner task list ────────────────────────────────────────────────────────

function PlannerTaskList({ data, onToggleTask }: { data: PlannerData; onToggleTask?: (taskId: string, done: boolean) => Promise<void> }) {
  const [filter, setFilter] = useState<Filter>("all");
  const [fagFilter, setFagFilter] = useState<string | null>(null);

  const bucketMap = Object.fromEntries(data.buckets.map((b) => [b.id, b.name]));
  const taskMap = Object.fromEntries(data.tasks.map((t) => [t.id, t]));

  const isHierarchical = data.tasks.some((t) => (t.outlineLevel ?? 1) > 1);

  function getFagomrade(task: PlannerTask): string {
    if (bucketMap[task.bucketId]) return bucketMap[task.bucketId];
    if (task.parentTaskId) {
      const parent = taskMap[task.parentTaskId];
      if (parent && bucketMap[parent.bucketId]) return bucketMap[parent.bucketId];
    }
    return "";
  }

  const statSource = isHierarchical
    ? data.tasks.filter((t) => (t.outlineLevel ?? 1) === 2)
    : data.tasks;

  const total = statSource.length;
  const done = statSource.filter((t) => t.percentComplete === 100).length;
  const inProg = statSource.filter((t) => t.percentComplete > 0 && t.percentComplete < 100).length;
  const remaining = total - done - inProg;
  const visibleStats = statSource.filter((t) => matchPlannerFilter(t.percentComplete, filter));

  // ── Hierarchical setup ──────────────────────────────────────────────────────
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

  // L3+ tasks grouped by parent — used for expand/collapse in task rows
  const childrenByParent: Record<string, PlannerTask[]> = {};
  if (isHierarchical) {
    for (const t of data.tasks) {
      if ((t.outlineLevel ?? 1) >= 3 && t.parentTaskId) {
        if (!childrenByParent[t.parentTaskId]) childrenByParent[t.parentTaskId] = [];
        childrenByParent[t.parentTaskId].push(t);
      }
    }
  }

  // ── Flat setup ──────────────────────────────────────────────────────────────
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
      {/* Stat cards */}
      <div style={{ display: "flex", gap: "1rem", marginBottom: "1rem", flexWrap: "wrap" }}>
        {([
          { key: "all"         as Filter, label: "Totalt",   value: total,     color: "#868E96" },
          { key: "done"        as Filter, label: "Ferdig",   value: done,      color: "#2F9E44" },
          { key: "in_progress" as Filter, label: "Pågående", value: inProg,    color: "#1971C2" },
          { key: "remaining"   as Filter, label: "Gjenstår", value: remaining, color: "#F76707" },
        ]).map(({ key, label, value, color }) => {
          const isActive = filter === key;
          return (
            <div
              key={key}
              onClick={() => key !== "all" && setFilter((prev) => (prev === key ? "all" : key))}
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
      </div>

      {/* Status filter indicator */}
      {filter !== "all" && (
        <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "0.75rem" }}>
          <span style={{ fontSize: "0.8rem", color: "var(--bfc-base-c-2)" }}>
            Viser: <strong>{FILTER_LABELS[filter]}</strong> ({visibleStats.length} {isHierarchical ? "leveranser" : "aktiviteter"})
          </span>
          <button
            onClick={() => setFilter("all")}
            style={{ background: "none", border: "1px solid var(--bfc-base-dimmed)", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "1px 8px", borderRadius: 4, fontSize: "0.78rem" }}
          >
            × Fjern filter
          </button>
        </div>
      )}

      {/* Fagområde filter chips (hierarchical only) */}
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
          // ── Hierarchical: L1 (fase) as headers, L2 (leveranse) as rows with fagomrade chip
          <>
            {l1Tasks.map((l1) => {
              const allL2 = l2ByL1[l1.id] ?? [];
              const visibleL2 = allL2.filter((t) =>
                matchPlannerFilter(t.percentComplete, filter) &&
                (!fagFilter || getFagomrade(t) === fagFilter)
              );
              if (visibleL2.length === 0 && (filter !== "all" || fagFilter !== null)) return null;
              const sectionDone = allL2.filter((t) => t.percentComplete === 100).length;
              const pct = allL2.length > 0 ? Math.round((sectionDone / allL2.length) * 100) : 0;
              return (
                <div key={l1.id}>
                  <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "0.5rem" }}>
                    <h3 className="bf-h4" style={{ margin: 0, flex: 1 }}>{l1.title}</h3>
                    <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>
                      {sectionDone}/{allL2.length} ferdig
                      {(filter !== "all" || fagFilter) && visibleL2.length < allL2.length && ` (viser ${visibleL2.length})`}
                    </span>
                    <div style={{ width: 80, height: 6, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
                      <div style={{ width: `${pct}%`, height: "100%", background: pct === 100 ? "#2F9E44" : "#1971C2", borderRadius: 3, transition: "width 0.3s" }} />
                    </div>
                  </div>
                  <div style={{ display: "grid", gap: "0.35rem" }}>
                    {visibleL2.map((t) => <PlannerTaskRow key={t.id} task={t} fagomrade={getFagomrade(t)} onToggle={onToggleTask} childrenByParent={childrenByParent} />)}
                  </div>
                </div>
              );
            })}
            {(l2ByL1["__orphan__"] ?? []).length > 0 && (() => {
              const visible = (l2ByL1["__orphan__"] ?? []).filter((t) =>
                matchPlannerFilter(t.percentComplete, filter) &&
                (!fagFilter || getFagomrade(t) === fagFilter)
              );
              if (visible.length === 0 && (filter !== "all" || fagFilter !== null)) return null;
              return (
                <div>
                  <h3 className="bf-h4" style={{ margin: "0 0 0.5rem", color: "var(--bfc-base-c-2)" }}>Uten fase</h3>
                  <div style={{ display: "grid", gap: "0.35rem" }}>
                    {visible.map((t) => <PlannerTaskRow key={t.id} task={t} fagomrade={getFagomrade(t)} onToggle={onToggleTask} childrenByParent={childrenByParent} />)}
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
              const bucketDone = allTasks.filter((t) => t.percentComplete === 100).length;
              const pct = allTasks.length > 0 ? Math.round((bucketDone / allTasks.length) * 100) : 0;
              return (
                <div key={bucket}>
                  <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "0.5rem" }}>
                    <h3 className="bf-h4" style={{ margin: 0, flex: 1 }}>{bucket}</h3>
                    <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>
                      {bucketDone}/{allTasks.length} ferdig
                      {filter !== "all" && tasks.length < allTasks.length && ` (viser ${tasks.length})`}
                    </span>
                    <div style={{ width: 80, height: 6, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
                      <div style={{ width: `${pct}%`, height: "100%", background: "#2F9E44", borderRadius: 3, transition: "width 0.3s" }} />
                    </div>
                  </div>
                  <div style={{ display: "grid", gap: "0.35rem" }}>
                    {tasks.map((task) => <PlannerTaskRow key={task.id} task={task} onToggle={onToggleTask} childrenByParent={childrenByParent} />)}
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

function PlannerTaskRow({
  task, fagomrade, onToggle, childrenByParent, depth = 0,
}: {
  task: PlannerTask;
  fagomrade?: string;
  onToggle?: (taskId: string, done: boolean) => Promise<void>;
  childrenByParent?: Record<string, PlannerTask[]>;
  depth?: number;
}) {
  const [hovered, setHovered] = useState(false);
  const [toggling, setToggling] = useState(false);
  const [expanded, setExpanded] = useState(false);
  const status = taskStatus(task.percentComplete);
  const cfg = STATUS_CONFIG[status];
  const isDone = status === "done";
  const circleColor = toggling ? "#ADB5BD" : cfg.color;
  const children = childrenByParent?.[task.id] ?? [];
  const hasChildren = children.length > 0;

  async function handleToggleClick() {
    if (!onToggle || toggling) return;
    setToggling(true);
    try {
      await onToggle(task.id, !isDone);
    } finally {
      setToggling(false);
    }
  }

  return (
    <div>
      <div
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        style={{
          display: "flex", alignItems: "center", gap: "0.75rem",
          padding: "0.65rem 1rem", borderRadius: 7,
          background: "var(--bfc-base-3)",
          border: "1px solid var(--bfc-base-dimmed)",
          borderLeft: depth > 0 ? "3px solid #0078D440" : "1px solid var(--bfc-base-dimmed)",
          marginLeft: depth > 0 ? `${depth * 1.5}rem` : undefined,
          boxShadow: hovered ? "0 2px 8px rgba(0,0,0,0.07)" : "none",
          transition: "box-shadow 0.15s",
        }}
      >
        {childrenByParent !== undefined && (
          hasChildren ? (
            <button
              onClick={() => setExpanded((v) => !v)}
              style={{
                background: "none", border: "none", cursor: "pointer", padding: "0 2px",
                fontSize: "0.75rem", color: "var(--bfc-base-c-2)", flexShrink: 0, lineHeight: 1,
              }}
              title={expanded ? "Skjul deloppgaver" : "Vis deloppgaver"}
            >
              {expanded ? "▾" : "▸"}
            </button>
          ) : (
            <div style={{ width: 18, flexShrink: 0 }} />
          )
        )}

        {onToggle && (
          <div
            onClick={() => { void handleToggleClick(); }}
            title={isDone ? "Merk som ikke ferdig" : "Merk som ferdig"}
            style={{
              width: 22, height: 22, borderRadius: "50%", flexShrink: 0,
              border: `2px solid ${circleColor}`,
              background: isDone && !toggling ? circleColor : "transparent",
              display: "flex", alignItems: "center", justifyContent: "center",
              cursor: toggling ? "wait" : "pointer",
              opacity: toggling ? 0.5 : 1,
              transition: "opacity 0.15s",
            }}
          >
            {isDone && !toggling && <span style={{ color: "#fff", fontSize: "0.7rem", fontWeight: 700 }}>✓</span>}
          </div>
        )}

        <span style={{
          flex: 1, fontSize: depth > 0 ? "0.85rem" : "0.9rem",
          textDecoration: isDone ? "line-through" : "none",
          color: isDone ? "var(--bfc-base-c-3)" : "inherit",
          overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
        }}>
          {task.title}
        </span>

        {depth === 0 && fagomrade ? (
          <span style={{ fontSize: "0.7rem", fontWeight: 600, flexShrink: 0, padding: "2px 8px", borderRadius: 20, background: "#0078D418", color: "#0078D4", whiteSpace: "nowrap" }}>
            {fagomrade}
          </span>
        ) : depth === 0 && task.labels && task.labels.length > 0 ? (
          <div style={{ display: "flex", gap: "0.3rem", flexShrink: 0, flexWrap: "wrap", maxWidth: 200 }}>
            {task.labels.map((label) => (
              <span key={label} style={{ fontSize: "0.7rem", fontWeight: 600, padding: "2px 7px", borderRadius: 20, background: "#7950F218", color: "#7950F2", whiteSpace: "nowrap" }}>
                {label}
              </span>
            ))}
          </div>
        ) : null}

        {task.percentComplete > 0 && task.percentComplete < 100 && (
          <div style={{ width: 60, height: 5, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden", flexShrink: 0 }}>
            <div style={{ width: `${task.percentComplete}%`, height: "100%", background: cfg.color, borderRadius: 3 }} />
          </div>
        )}

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

      {hasChildren && expanded && (
        <div style={{ display: "grid", gap: "0.25rem", marginTop: "0.25rem" }}>
          {children.map((child) => (
            <PlannerTaskRow
              key={child.id}
              task={child}
              childrenByParent={childrenByParent}
              depth={depth + 1}
              onToggle={onToggle}
            />
          ))}
        </div>
      )}
    </div>
  );
}
