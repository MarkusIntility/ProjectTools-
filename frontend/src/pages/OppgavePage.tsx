import { useEffect, useState, useCallback } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type OppgaveListe, type Oppgave, type Template } from "../api/client";
import { isMsalConfigured, msalInstance, PLANNER_SCOPES } from "../auth/msalConfig";
import { fetchPlannerData, parsePlanId, taskStatus, togglePlannerTask, type PlannerData, type PlannerTask } from "../auth/plannerService";
import type { AccountInfo } from "@azure/msal-browser";

// ─── Config ───────────────────────────────────────────────────────────────────

const SOURCE_CONFIG = {
  planner:    { label: "Microsoft Planner", color: "#0078D4", initial: "P" },
  smartsheet: { label: "Smartsheet",        color: "#00A88E", initial: "S" },
  own:        { label: "Egen",              color: "#7950F2", initial: "E" },
};

const STATUS_CONFIG = {
  not_started: { label: "Ikke startet", color: "#868E96", bg: "#868E9618" },
  in_progress:  { label: "Pågående",     color: "#1971C2", bg: "#1971C218" },
  done:         { label: "Ferdig",       color: "#2F9E44", bg: "#2F9E4418" },
};

type OppgaveStatus = "not_started" | "in_progress" | "done";
type Filter = "all" | "done" | "in_progress" | "remaining";

const FILTER_LABELS: Record<Filter, string> = {
  all: "Alle", done: "Ferdig", in_progress: "Pågående", remaining: "Gjenstår",
};

const EMPTY_FORM = {
  name: "",
  responsible: "",
  due_date: "",
  status: "not_started" as OppgaveStatus,
  description: "",
};

// ─── Main page ────────────────────────────────────────────────────────────────

export default function OppgavePage() {
  const { projectId, listeId } = useParams<{ projectId: string; listeId: string }>();
  const navigate = useNavigate();
  const [liste, setListe] = useState<OppgaveListe | null>(null);

  const [msalReady] = useState(isMsalConfigured);
  const [plannerAccount, setPlannerAccount] = useState<AccountInfo | null>(null);
  const [plannerData, setPlannerData] = useState<PlannerData | null>(null);
  const [plannerLoading, setPlannerLoading] = useState(false);
  const [plannerError, setPlannerError] = useState<string | null>(null);

  const [taskModal, setTaskModal] = useState(false);
  const [editTarget, setEditTarget] = useState<Oppgave | null>(null);
  const [form, setForm] = useState(EMPTY_FORM);
  const [saving, setSaving] = useState(false);

  const [deleteTarget, setDeleteTarget] = useState<Oppgave | null>(null);
  const [deleting, setDeleting] = useState(false);

  const [editListeModal, setEditListeModal] = useState(false);

  const [templateModal, setTemplateModal] = useState(false);
  const [templateMode, setTemplateMode] = useState<"new" | "existing">("new");
  const [templateName, setTemplateName] = useState("");
  const [existingTemplates, setExistingTemplates] = useState<Template[]>([]);
  const [selectedExistingId, setSelectedExistingId] = useState("");
  const [templateSaving, setTemplateSaving] = useState(false);
  const [listeTitle, setListeTitle] = useState("");
  const [listeUrl, setListeUrl] = useState("");

  const [filter, setFilter] = useState<Filter>("all");

  useEffect(() => {
    if (!projectId || !listeId) return;
    api.oppgaveLister.get(projectId, listeId).then(setListe);
  }, [projectId, listeId]);

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
    if (!plannerAccount || !liste || liste.source !== "planner" || !liste.external_url) return;
    loadPlannerData(plannerAccount, liste.external_url);
  }, [plannerAccount, liste, loadPlannerData]);

  async function loginPlanner() {
    if (!isMsalConfigured || !msalReady) return;
    Object.keys(sessionStorage).filter((k) => k.includes("interaction")).forEach((k) => sessionStorage.removeItem(k));
    sessionStorage.setItem("app.returnUrl", window.location.href);
    await msalInstance.loginRedirect({ scopes: PLANNER_SCOPES });
  }

  async function handleToggleTask(taskId: string, done: boolean): Promise<void> {
    if (!plannerData || !plannerAccount || !liste || !liste.external_url) return;
    setPlannerData((prev) => prev ? {
      ...prev,
      tasks: prev.tasks.map((t) => t.id === taskId ? { ...t, percentComplete: done ? 100 : 0 } : t),
    } : null);
    try {
      await togglePlannerTask(msalInstance, plannerAccount, plannerData, taskId, done);
    } catch (err) {
      setPlannerError(err instanceof Error ? err.message : "Kunne ikke oppdatere oppgave i Planner");
      loadPlannerData(plannerAccount, liste.external_url);
    }
  }

  function openAdd() {
    setEditTarget(null);
    setForm(EMPTY_FORM);
    setTaskModal(true);
  }

  function openEdit(o: Oppgave) {
    setEditTarget(o);
    setForm({
      name: o.name,
      responsible: o.responsible ?? "",
      due_date: o.due_date ? o.due_date.slice(0, 16) : "",
      status: o.status as OppgaveStatus,
      description: o.description ?? "",
    });
    setTaskModal(true);
  }

  async function handleSave() {
    if (!projectId || !listeId || !form.name.trim()) return;
    setSaving(true);
    try {
      const payload = {
        name: form.name,
        responsible: form.responsible || null,
        due_date: form.due_date || null,
        status: form.status,
        description: form.description || null,
        sort_order: editTarget?.sort_order ?? 0,
      };
      if (editTarget) {
        const updated = await api.oppgaveLister.updateOppgave(projectId, listeId, editTarget.id, payload);
        setListe((prev) => prev ? { ...prev, oppgaver: prev.oppgaver.map((o) => o.id === updated.id ? updated : o) } : prev);
      } else {
        const created = await api.oppgaveLister.addOppgave(projectId, listeId, payload);
        setListe((prev) => prev ? { ...prev, oppgaver: [...prev.oppgaver, created] } : prev);
      }
      setTaskModal(false);
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete() {
    if (!projectId || !listeId || !deleteTarget) return;
    setDeleting(true);
    try {
      await api.oppgaveLister.deleteOppgave(projectId, listeId, deleteTarget.id);
      setListe((prev) => prev ? { ...prev, oppgaver: prev.oppgaver.filter((o) => o.id !== deleteTarget.id) } : prev);
      setDeleteTarget(null);
    } finally {
      setDeleting(false);
    }
  }

  async function toggleStatus(o: Oppgave) {
    if (!projectId || !listeId) return;
    const next: OppgaveStatus = o.status === "done" ? "not_started" : "done";
    const updated = await api.oppgaveLister.updateOppgave(projectId, listeId, o.id, { ...o, status: next });
    setListe((prev) => prev ? { ...prev, oppgaver: prev.oppgaver.map((x) => x.id === updated.id ? updated : x) } : prev);
  }

  async function handleSaveListe() {
    if (!projectId || !listeId || !listeTitle.trim()) return;
    const updated = await api.oppgaveLister.update(projectId, listeId, {
      title: listeTitle,
      external_url: listeUrl || null,
    });
    setListe(updated);
    setEditListeModal(false);
  }

  async function openTemplateModal() {
    const ts = await api.templates.list("oppgave_liste");
    setExistingTemplates(ts);
    setTemplateMode("new");
    setTemplateName(liste?.title ?? "");
    setSelectedExistingId(ts[0]?.id ?? "");
    setTemplateModal(true);
  }

  async function saveAsTemplate() {
    if (!liste) return;
    setTemplateSaving(true);
    try {
      const data = JSON.stringify({ oppgaver: liste.oppgaver.map(({ name, responsible, due_date, status, description }) => ({ name, responsible, due_date, status, description })) });
      if (templateMode === "new") {
        await api.templates.create({ name: templateName.trim() || liste.title, type: "oppgave_liste", data });
      } else {
        await api.templates.update(selectedExistingId, { name: existingTemplates.find(t => t.id === selectedExistingId)?.name ?? templateName, data });
      }
      setTemplateModal(false);
    } finally {
      setTemplateSaving(false);
    }
  }

  if (!liste) return <div style={{ padding: "2rem" }}>Laster...</div>;

  const srcCfg = SOURCE_CONFIG[liste.source];

  return (
    <div style={{ maxWidth: 900, margin: "0 auto", padding: "2rem" }}>
      <button
        onClick={() => navigate(`/projects/${projectId}`)}
        style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", marginBottom: "1.5rem", fontSize: "0.9rem", padding: 0 }}
      >
        ← Tilbake til prosjekt
      </button>

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
            <h1 className="bf-h2" style={{ margin: 0 }}>{liste.title}</h1>
            <span style={{ fontSize: "0.8rem", color: srcCfg.color, fontWeight: 600 }}>{srcCfg.label}</span>
          </div>
        </div>
        <div style={{ display: "flex", gap: "0.5rem", flexShrink: 0 }}>
          <Button variant="outline" onClick={() => { setListeTitle(liste.title); setListeUrl(liste.external_url ?? ""); setEditListeModal(true); }}>
            Rediger
          </Button>
          {liste.source === "own" && (
            <>
              <Button variant="outline" onClick={openTemplateModal}>Lagre som mal</Button>
              <Button variant="filled" onClick={openAdd}>+ Legg til oppgave</Button>
            </>
          )}
        </div>
      </div>

      {/* Own source */}
      {liste.source === "own" && <OwnView liste={liste} filter={filter} setFilter={setFilter} onEdit={openEdit} onDelete={setDeleteTarget} onToggle={toggleStatus} />}

      {/* Planner source */}
      {liste.source === "planner" && (
        <PlannerView
          liste={liste}
          srcCfg={srcCfg}
          account={plannerAccount}
          data={plannerData}
          loading={plannerLoading}
          error={plannerError}
          msalReady={msalReady}
          onLogin={loginPlanner}
          onRefresh={() => { if (plannerAccount && liste.external_url) loadPlannerData(plannerAccount, liste.external_url); }}
          onToggleTask={handleToggleTask}
        />
      )}

      {/* Smartsheet source */}
      {liste.source === "smartsheet" && (
        <div style={{ padding: "2rem", borderRadius: 10, background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)", borderTop: `4px solid ${srcCfg.color}`, textAlign: "center", maxWidth: 540, margin: "0 auto" }}>
          <div style={{ fontSize: "2rem", marginBottom: "0.75rem" }}>📊</div>
          <h2 className="bf-h4" style={{ margin: "0 0 0.5rem" }}>Styres i Smartsheet</h2>
          {liste.external_url ? (
            <a href={liste.external_url} target="_blank" rel="noopener noreferrer"
              style={{ display: "inline-block", padding: "0.6rem 1.5rem", background: srcCfg.color, color: "#fff", borderRadius: 6, fontWeight: 600, textDecoration: "none" }}>
              Åpne i Smartsheet →
            </a>
          ) : (
            <p style={{ color: "var(--bfc-base-c-3)", fontSize: "0.85rem" }}>Ingen URL lagt til.</p>
          )}
        </div>
      )}

      {/* Task form modal */}
      <Modal isOpen={templateModal} onRequestClose={() => setTemplateModal(false)} header="Lagre som mal">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.5rem" }}>
            {(["new", "existing"] as const).map((mode) => (
              <button key={mode} onClick={() => setTemplateMode(mode)}
                style={{ padding: "0.75rem", borderRadius: 8, border: `2px solid ${templateMode === mode ? "#F59F00" : "var(--bfc-base-dimmed)"}`, background: templateMode === mode ? "#F59F0018" : "var(--bfc-base-3)", cursor: "pointer", fontWeight: 600, fontSize: "0.9rem", color: templateMode === mode ? "#F59F00" : "var(--bfc-base-c-1)", transition: "all 0.15s" }}>
                {mode === "new" ? "Ny mal" : "Oppdater eksisterende"}
              </button>
            ))}
          </div>
          {templateMode === "new" ? (
            <Input label="Navn på malen" value={templateName} onChange={(e) => setTemplateName(e.target.value)} placeholder="f.eks. Standard oppgaveliste" autoFocus />
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
          <p style={{ margin: 0, color: "var(--bfc-base-c-2)", fontSize: "0.85rem" }}>{liste.oppgaver.length} oppgaver vil bli lagret i malen.</p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setTemplateModal(false)}>Avbryt</Button>
            <Button variant="filled" onClick={saveAsTemplate} state={templateSaving || (templateMode === "existing" && !selectedExistingId) ? "inactive" : "default"}>
              {templateSaving ? "Lagrer..." : "Lagre mal"}
            </Button>
          </div>
        </div>
      </Modal>

      <Modal isOpen={taskModal} onRequestClose={() => setTaskModal(false)} header={editTarget ? "Rediger oppgave" : "Ny oppgave"}>
        <div style={{ display: "flex", flexDirection: "column", gap: "0.85rem" }}>
          <Input label="Oppgavenavn *" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} autoFocus />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.75rem" }}>
            <Input label="Ansvarlig" value={form.responsible} onChange={(e) => setForm({ ...form, responsible: e.target.value })} placeholder="Navn eller team" />
            <Input label="Forfallsdato" type="datetime-local" value={form.due_date} onChange={(e) => setForm({ ...form, due_date: e.target.value })} />
          </div>
          <div>
            <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 500, marginBottom: "0.35rem" }}>Status</label>
            <select
              value={form.status}
              onChange={(e) => setForm({ ...form, status: e.target.value as OppgaveStatus })}
              style={{ width: "100%", padding: "0.5rem 0.75rem", borderRadius: 6, border: "1px solid var(--bfc-base-dimmed)", background: "var(--bfc-base-3)", color: "inherit", fontSize: "0.9rem" }}
            >
              <option value="not_started">Ikke startet</option>
              <option value="in_progress">Pågående</option>
              <option value="done">Ferdig</option>
            </select>
          </div>
          <div>
            <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 500, marginBottom: "0.35rem" }}>Beskrivelse</label>
            <textarea
              value={form.description}
              onChange={(e) => setForm({ ...form, description: e.target.value })}
              rows={3}
              style={{ width: "100%", padding: "0.5rem 0.75rem", borderRadius: 6, border: "1px solid var(--bfc-base-dimmed)", background: "var(--bfc-base-3)", color: "inherit", fontSize: "0.9rem", resize: "vertical", boxSizing: "border-box" }}
            />
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setTaskModal(false)}>Avbryt</Button>
            <Button variant="filled" onClick={handleSave} state={!form.name.trim() || saving ? "inactive" : "default"}>
              {saving ? "Lagrer..." : "Lagre"}
            </Button>
          </div>
        </div>
      </Modal>

      <Modal isOpen={!!deleteTarget} onRequestClose={() => setDeleteTarget(null)} header="Slett oppgave">
        <div style={{ display: "flex", flexDirection: "column", gap: "1.25rem" }}>
          <p style={{ margin: 0 }}>Er du sikker på at du vil slette <strong>«{deleteTarget?.name}»</strong>?</p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setDeleteTarget(null)}>Avbryt</Button>
            <Button variant="filled" state={deleting ? "inactive" : "default"} onClick={handleDelete}
              style={{ background: "#E03131", borderColor: "#E03131" }}>
              {deleting ? "Sletter..." : "Ja, slett"}
            </Button>
          </div>
        </div>
      </Modal>

      <Modal isOpen={editListeModal} onRequestClose={() => setEditListeModal(false)} header="Rediger oppgaveliste">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <Input label="Tittel" value={listeTitle} onChange={(e) => setListeTitle(e.target.value)} />
          {liste.source !== "own" && (
            <Input label={`URL til ${srcCfg.label}`} value={listeUrl} onChange={(e) => setListeUrl(e.target.value)} placeholder="https://..." />
          )}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setEditListeModal(false)}>Avbryt</Button>
            <Button variant="filled" onClick={handleSaveListe} state={!listeTitle.trim() ? "inactive" : "default"}>Lagre</Button>
          </div>
        </div>
      </Modal>
    </div>
  );
}

// ─── Own view ─────────────────────────────────────────────────────────────────

function OwnView({ liste, filter, setFilter, onEdit, onDelete, onToggle }: {
  liste: OppgaveListe;
  filter: Filter;
  setFilter: (f: Filter) => void;
  onEdit: (o: Oppgave) => void;
  onDelete: (o: Oppgave) => void;
  onToggle: (o: Oppgave) => void;
}) {
  const total = liste.oppgaver.length;
  const done = liste.oppgaver.filter((o) => o.status === "done").length;
  const inProg = liste.oppgaver.filter((o) => o.status === "in_progress").length;
  const remaining = total - done - inProg;

  function matchFilter(o: Oppgave): boolean {
    if (filter === "done") return o.status === "done";
    if (filter === "in_progress") return o.status === "in_progress";
    if (filter === "remaining") return o.status === "not_started";
    return true;
  }

  const visible = liste.oppgaver.filter(matchFilter);

  if (total === 0) {
    return (
      <div style={{ textAlign: "center", padding: "3rem", border: "2px dashed var(--bfc-base-dimmed)", borderRadius: 8, color: "var(--bfc-base-c-2)" }}>
        <p style={{ marginBottom: "1rem" }}>Ingen oppgaver ennå.</p>
      </div>
    );
  }

  const STAT_CARDS = [
    { key: "all" as Filter, label: "Totalt", value: total, color: "#868E96" },
    { key: "done" as Filter, label: "Ferdig", value: done, color: "#2F9E44" },
    { key: "in_progress" as Filter, label: "Pågående", value: inProg, color: "#1971C2" },
    { key: "remaining" as Filter, label: "Gjenstår", value: remaining, color: "#F76707" },
  ];

  return (
    <div>
      <div style={{ display: "flex", gap: "1rem", marginBottom: "1rem", flexWrap: "wrap" }}>
        {STAT_CARDS.map(({ key, label, value, color }) => {
          const isActive = filter === key;
          return (
            <div
              key={key}
              onClick={() => key !== "all" && setFilter(filter === key ? "all" : key)}
              style={{
                padding: "0.6rem 1.1rem", borderRadius: 8, textAlign: "center", minWidth: 80,
                background: isActive ? `${color}22` : `${color}12`,
                border: `1px solid ${isActive ? color : `${color}30`}`,
                cursor: key === "all" ? "default" : "pointer",
                outline: isActive && key !== "all" ? `2px solid ${color}` : "none",
                outlineOffset: 1, transition: "all 0.15s",
              }}
            >
              <div style={{ fontSize: "1.4rem", fontWeight: 700, color }}>{value}</div>
              <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)" }}>{label}</div>
            </div>
          );
        })}
      </div>

      {filter !== "all" && (
        <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "1rem" }}>
          <span style={{ fontSize: "0.8rem", color: "var(--bfc-base-c-2)" }}>
            Viser: <strong>{FILTER_LABELS[filter]}</strong> ({visible.length} oppgaver)
          </span>
          <button onClick={() => setFilter("all")}
            style={{ background: "none", border: "1px solid var(--bfc-base-dimmed)", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "1px 8px", borderRadius: 4, fontSize: "0.78rem" }}>
            × Fjern filter
          </button>
        </div>
      )}

      <div style={{ display: "grid", gap: "0.35rem" }}>
        {visible.map((o) => (
          <OppgaveRow key={o.id} oppgave={o} onEdit={() => onEdit(o)} onDelete={() => onDelete(o)} onToggle={() => onToggle(o)} />
        ))}
      </div>
    </div>
  );
}

// ─── Oppgave row ──────────────────────────────────────────────────────────────

function OppgaveRow({ oppgave, onEdit, onDelete, onToggle }: {
  oppgave: Oppgave;
  onEdit: () => void;
  onDelete: () => void;
  onToggle: () => void;
}) {
  const [hovered, setHovered] = useState(false);
  const cfg = STATUS_CONFIG[oppgave.status as keyof typeof STATUS_CONFIG] ?? STATUS_CONFIG.not_started;
  const isDone = oppgave.status === "done";

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
      <button
        onClick={onToggle}
        style={{
          width: 22, height: 22, borderRadius: "50%", flexShrink: 0,
          border: `2px solid ${cfg.color}`,
          background: isDone ? cfg.color : "transparent",
          cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center",
          transition: "background 0.15s",
        }}
      >
        {isDone && <span style={{ color: "#fff", fontSize: "0.7rem", fontWeight: 700 }}>✓</span>}
      </button>

      <span style={{
        flex: 1, fontSize: "0.9rem",
        textDecoration: isDone ? "line-through" : "none",
        color: isDone ? "var(--bfc-base-c-3)" : "inherit",
        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
      }}>
        {oppgave.name}
      </span>

      {oppgave.responsible && (
        <span style={{ fontSize: "0.75rem", padding: "2px 8px", borderRadius: 20, background: "var(--bfc-base-dimmed)", color: "var(--bfc-base-c-2)", flexShrink: 0 }}>
          {oppgave.responsible}
        </span>
      )}

      {oppgave.due_date && (
        <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)", flexShrink: 0, whiteSpace: "nowrap" }}>
          {new Date(oppgave.due_date).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
        </span>
      )}

      <span style={{ fontSize: "0.72rem", fontWeight: 600, padding: "2px 8px", borderRadius: 20, background: cfg.bg, color: cfg.color, flexShrink: 0 }}>
        {cfg.label}
      </span>

      <div style={{ display: "flex", gap: "0.25rem", flexShrink: 0, opacity: hovered ? 1 : 0, transition: "opacity 0.15s" }}
        onClick={(e) => e.stopPropagation()}>
        <button onClick={onEdit} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "2px 6px", borderRadius: 4, fontSize: "0.8rem" }}>Endre</button>
        <button onClick={onDelete} style={{ background: "none", border: "none", cursor: "pointer", color: "#E03131", padding: "2px 6px", borderRadius: 4, fontSize: "0.8rem" }}>Slett</button>
      </div>
    </div>
  );
}

// ─── Planner view (flat list) ─────────────────────────────────────────────────

function PlannerView({ liste, srcCfg, account, data, loading, error, msalReady, onLogin, onRefresh, onToggleTask }: {
  liste: OppgaveListe;
  srcCfg: { label: string; color: string };
  account: AccountInfo | null;
  data: PlannerData | null;
  loading: boolean;
  error: string | null;
  msalReady: boolean;
  onLogin: () => void;
  onRefresh: () => void;
  onToggleTask: (taskId: string, done: boolean) => Promise<void>;
}) {
  const planId = liste.external_url ? parsePlanId(liste.external_url) : null;

  if (!isMsalConfigured) {
    return (
      <div style={{ padding: "1.5rem", borderRadius: 10, background: "#FFF3BF", border: "1px solid #FAB005", maxWidth: 600 }}>
        <h3 style={{ margin: "0 0 0.5rem", color: "#5C3A00" }}>Azure AD ikke konfigurert</h3>
        <p style={{ margin: 0, fontSize: "0.9rem", color: "#7C4D00" }}>Legg til <code>VITE_AZURE_CLIENT_ID</code> og <code>VITE_AZURE_TENANT_ID</code> som GitHub Secrets.</p>
      </div>
    );
  }

  if (!liste.external_url || !planId) {
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
            {liste.external_url && (
              <a href={liste.external_url} target="_blank" rel="noopener noreferrer"
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
        <div style={{ padding: "1rem 1.25rem", borderRadius: 8, background: "#FFE3E3", border: "1px solid #FFA8A8", marginBottom: "1rem" }}>
          <span style={{ color: "#C92A2A", fontSize: "0.9rem" }}>{error.replace(/^PREMIUM_PLAN:\s*/, "")}</span>
        </div>
      )}

      {loading && <div style={{ textAlign: "center", padding: "3rem", color: "var(--bfc-base-c-2)" }}>Henter oppgaver fra Planner…</div>}

      {data && !loading && <FlatPlannerList data={data} onToggle={data.source === "premium" ? undefined : onToggleTask} />}
    </div>
  );
}

// ─── Planner task list ────────────────────────────────────────────────────────

function FlatPlannerList({ data, onToggle }: { data: PlannerData; onToggle?: (taskId: string, done: boolean) => Promise<void> }) {
  const assigneeMap = data.assigneeMap ?? {};
  const [filter, setFilter] = useState<Filter>("all");

  const isHierarchical = data.tasks.some((t) => (t.outlineLevel ?? 1) > 1);

  // Stats source: L2 tasks in hierarchical mode, all tasks in flat mode
  const statSource = isHierarchical
    ? data.tasks.filter((t) => (t.outlineLevel ?? 1) === 2)
    : data.tasks;
  const total = statSource.length;
  const done = statSource.filter((t) => t.percentComplete === 100).length;
  const inProg = statSource.filter((t) => t.percentComplete > 0 && t.percentComplete < 100).length;
  const remaining = total - done - inProg;

  function matchFilter(t: PlannerTask): boolean {
    if (filter === "done") return t.percentComplete === 100;
    if (filter === "in_progress") return t.percentComplete > 0 && t.percentComplete < 100;
    if (filter === "remaining") return t.percentComplete === 0;
    return true;
  }

  // ── Hierarchical setup ───────────────────────────────────────────────────────
  const l1Tasks = isHierarchical ? data.tasks.filter((t) => (t.outlineLevel ?? 1) === 1) : [];
  const l2ByL1: Record<string, PlannerTask[]> = {};
  if (isHierarchical) {
    for (const t of data.tasks.filter((t2) => (t2.outlineLevel ?? 1) === 2)) {
      const key = t.parentTaskId ?? "__orphan__";
      if (!l2ByL1[key]) l2ByL1[key] = [];
      l2ByL1[key].push(t);
    }
  }
  const childrenByParent: Record<string, PlannerTask[]> = {};
  if (isHierarchical) {
    for (const t of data.tasks) {
      if ((t.outlineLevel ?? 1) >= 3 && t.parentTaskId) {
        if (!childrenByParent[t.parentTaskId]) childrenByParent[t.parentTaskId] = [];
        childrenByParent[t.parentTaskId].push(t);
      }
    }
  }

  // ── Flat setup ───────────────────────────────────────────────────────────────
  const flatVisible = isHierarchical ? [] : statSource.filter(matchFilter);
  const visibleCount = isHierarchical ? statSource.filter(matchFilter).length : flatVisible.length;

  const STAT_CARDS = [
    { key: "all" as Filter, label: "Totalt", value: total, color: "#868E96" },
    { key: "done" as Filter, label: "Ferdig", value: done, color: "#2F9E44" },
    { key: "in_progress" as Filter, label: "Pågående", value: inProg, color: "#1971C2" },
    { key: "remaining" as Filter, label: "Gjenstår", value: remaining, color: "#F76707" },
  ];

  return (
    <div>
      <div style={{ display: "flex", gap: "1rem", marginBottom: "1rem", flexWrap: "wrap" }}>
        {STAT_CARDS.map(({ key, label, value, color }) => {
          const isActive = filter === key;
          return (
            <div
              key={key}
              onClick={() => key !== "all" && setFilter(filter === key ? "all" : key)}
              style={{
                padding: "0.6rem 1.1rem", borderRadius: 8, textAlign: "center", minWidth: 80,
                background: isActive ? `${color}22` : `${color}12`,
                border: `1px solid ${isActive ? color : `${color}30`}`,
                cursor: key === "all" ? "default" : "pointer",
                outline: isActive && key !== "all" ? `2px solid ${color}` : "none",
                outlineOffset: 1, transition: "all 0.15s",
              }}
            >
              <div style={{ fontSize: "1.4rem", fontWeight: 700, color }}>{value}</div>
              <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)" }}>{label}</div>
            </div>
          );
        })}
      </div>

      {filter !== "all" && (
        <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", marginBottom: "1rem" }}>
          <span style={{ fontSize: "0.8rem", color: "var(--bfc-base-c-2)" }}>
            Viser: <strong>{FILTER_LABELS[filter]}</strong> ({visibleCount} oppgaver)
          </span>
          <button onClick={() => setFilter("all")}
            style={{ background: "none", border: "1px solid var(--bfc-base-dimmed)", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "1px 8px", borderRadius: 4, fontSize: "0.78rem" }}>
            × Fjern filter
          </button>
        </div>
      )}

      <div style={{ display: "grid", gap: isHierarchical ? "1.5rem" : "0.35rem" }}>
        {isHierarchical ? (
          <>
            {l1Tasks.map((l1) => {
              const allL2 = l2ByL1[l1.id] ?? [];
              const visibleL2 = allL2.filter(matchFilter);
              if (visibleL2.length === 0 && filter !== "all") return null;
              const sectionDone = allL2.filter((t) => t.percentComplete === 100).length;
              const pct = allL2.length > 0 ? Math.round((sectionDone / allL2.length) * 100) : 0;
              return (
                <div key={l1.id}>
                  <div style={{ display: "flex", alignItems: "center", gap: "0.75rem", marginBottom: "0.5rem" }}>
                    <h3 className="bf-h4" style={{ margin: 0, flex: 1 }}>{l1.title}</h3>
                    <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>
                      {sectionDone}/{allL2.length} ferdig
                      {filter !== "all" && visibleL2.length < allL2.length && ` (viser ${visibleL2.length})`}
                    </span>
                    <div style={{ width: 80, height: 6, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
                      <div style={{ width: `${pct}%`, height: "100%", background: pct === 100 ? "#2F9E44" : "#1971C2", borderRadius: 3, transition: "width 0.3s" }} />
                    </div>
                  </div>
                  <div style={{ display: "grid", gap: "0.35rem" }}>
                    {visibleL2.map((t) => (
                      <PlannerTaskRowOppgave key={t.id} task={t} assigneeMap={assigneeMap} childrenByParent={childrenByParent} onToggle={onToggle} />
                    ))}
                  </div>
                </div>
              );
            })}
            {(l2ByL1["__orphan__"] ?? []).length > 0 && (() => {
              const orphanVisible = (l2ByL1["__orphan__"] ?? []).filter(matchFilter);
              if (orphanVisible.length === 0 && filter !== "all") return null;
              return (
                <div>
                  <h3 className="bf-h4" style={{ margin: "0 0 0.5rem", color: "var(--bfc-base-c-2)" }}>Uten fase</h3>
                  <div style={{ display: "grid", gap: "0.35rem" }}>
                    {orphanVisible.map((t) => (
                      <PlannerTaskRowOppgave key={t.id} task={t} assigneeMap={assigneeMap} childrenByParent={childrenByParent} onToggle={onToggle} />
                    ))}
                  </div>
                </div>
              );
            })()}
          </>
        ) : (
          flatVisible.map((task) => (
            <PlannerTaskRowOppgave key={task.id} task={task} assigneeMap={assigneeMap} onToggle={onToggle} />
          ))
        )}
      </div>
    </div>
  );
}

function PlannerTaskRowOppgave({
  task, assigneeMap, childrenByParent, depth = 0, onToggle,
}: {
  task: PlannerTask;
  assigneeMap: Record<string, string>;
  childrenByParent?: Record<string, PlannerTask[]>;
  depth?: number;
  onToggle?: (taskId: string, done: boolean) => Promise<void>;
}) {
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
      <div style={{
        display: "flex", alignItems: "center", gap: "0.75rem",
        padding: "0.65rem 1rem", borderRadius: 7,
        background: "var(--bfc-base-3)",
        border: "1px solid var(--bfc-base-dimmed)",
        borderLeft: depth > 0 ? "3px solid #0078D440" : "1px solid var(--bfc-base-dimmed)",
        marginLeft: depth > 0 ? `${depth * 1.5}rem` : undefined,
      }}>
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
        <div
          onClick={onToggle ? () => { void handleToggleClick(); } : undefined}
          title={onToggle ? (isDone ? "Merk som ikke ferdig" : "Merk som ferdig") : undefined}
          style={{
            width: 22, height: 22, borderRadius: "50%", flexShrink: 0,
            border: `2px solid ${circleColor}`, background: isDone && !toggling ? circleColor : "transparent",
            display: "flex", alignItems: "center", justifyContent: "center",
            cursor: onToggle ? (toggling ? "wait" : "pointer") : "default",
            opacity: toggling ? 0.5 : 1, transition: "opacity 0.15s",
          }}
        >
          {isDone && !toggling && <span style={{ color: "#fff", fontSize: "0.7rem", fontWeight: 700 }}>✓</span>}
        </div>
        <span style={{
          flex: 1, fontSize: depth > 0 ? "0.85rem" : "0.9rem",
          textDecoration: isDone ? "line-through" : "none",
          color: isDone ? "var(--bfc-base-c-3)" : "inherit",
          overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
        }}>
          {task.title}
        </span>
        {depth === 0 && (() => {
          const assignees = Object.keys(task.assignments ?? {}).map((id) => assigneeMap[id]).filter(Boolean);
          return assignees.length > 0 ? (
            <span style={{ fontSize: "0.75rem", padding: "2px 8px", borderRadius: 20, background: "var(--bfc-base-dimmed)", color: "var(--bfc-base-c-2)", flexShrink: 0, whiteSpace: "nowrap" }}>
              {assignees.join(", ")}
            </span>
          ) : null;
        })()}
        {task.dueDateTime && (
          <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)", flexShrink: 0, whiteSpace: "nowrap" }}>
            {new Date(task.dueDateTime).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })}
          </span>
        )}
        {depth === 0 && (
          <span style={{ fontSize: "0.72rem", fontWeight: 600, padding: "2px 8px", borderRadius: 20, background: cfg.bg, color: cfg.color, flexShrink: 0 }}>
            {cfg.label}
          </span>
        )}
      </div>
      {hasChildren && expanded && (
        <div style={{ display: "grid", gap: "0.25rem", marginTop: "0.25rem" }}>
          {children.map((child) => (
            <PlannerTaskRowOppgave
              key={child.id}
              task={child}
              assigneeMap={assigneeMap}
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
