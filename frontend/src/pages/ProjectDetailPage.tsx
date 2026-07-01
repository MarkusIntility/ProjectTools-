import { useCallback, useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type Project, type RiskMatrix, type RiskItem, type CommunicationPlan, type MeetingPlan, type Runbook, type ProjectPlan, type OppgaveListe, type Template } from "../api/client";
import { isMsalConfigured, msalInstance } from "../auth/msalConfig";
import { fetchPlannerData } from "../auth/plannerService";
import { exportDashboardPdf } from "../utils/exportUtils";

const ACCENT_COLORS = [
  "#4C6EF5", "#7950F2", "#E64980", "#F76707",
  "#2F9E44", "#1098AD", "#E67700", "#862E9C",
];

function accentColor(name: string): string {
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = name.charCodeAt(i) + ((hash << 5) - hash);
  }
  return ACCENT_COLORS[Math.abs(hash) % ACCENT_COLORS.length];
}

function initials(name: string): string {
  return name.split(" ").map((w) => w[0]).slice(0, 2).join("").toUpperCase();
}

const SECTION_CONFIG = {
  risk:        { color: "#E03131", label: "Risikomatriser" },
  comm:        { color: "#1971C2", label: "Kommunikasjonsplaner" },
  meeting:     { color: "#2F9E44", label: "Møteplaner" },
  projectplan: { color: "#0CA678", label: "Prosjektplaner" },
  oppgave:     { color: "#F59F00", label: "Oppgaver" },
  runbook:     { color: "#7950F2", label: "Runbooks" },
};

type PlanType = "risk" | "comm" | "meeting" | "projectplan" | "oppgave" | "runbook";
type PlanItem = { id: string; title: string; is_primary?: boolean };
type OwnStep = "choice" | "template" | "config";

function TemplatePicker({ templates, selected, onSelect }: { templates: Template[]; selected: Template | null; onSelect: (t: Template) => void }) {
  if (templates.length === 0) {
    return (
      <div style={{ padding: "1rem", border: "1px dashed var(--bfc-base-dimmed)", borderRadius: 8, color: "var(--bfc-base-c-2)", fontSize: "0.9rem", textAlign: "center" }}>
        Ingen maler funnet for denne typen
      </div>
    );
  }
  return (
    <div style={{ display: "grid", gap: "0.5rem", maxHeight: 240, overflowY: "auto" }}>
      {templates.map((t) => (
        <button
          key={t.id}
          onClick={() => onSelect(t)}
          style={{
            border: `2px solid ${selected?.id === t.id ? "#0078D4" : "var(--bfc-base-dimmed)"}`,
            borderRadius: 8, padding: "0.75rem 1rem",
            background: selected?.id === t.id ? "#0078D418" : "var(--bfc-base-3)",
            cursor: "pointer", textAlign: "left", transition: "border-color 0.15s, background 0.15s",
          }}
        >
          <div style={{ fontWeight: 600, fontSize: "0.9rem" }}>{t.name}</div>
          <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)", marginTop: 2 }}>
            Oppdatert: {new Date(t.updated_at).toLocaleDateString("nb-NO")}
          </div>
        </button>
      ))}
    </div>
  );
}

function OwnChoiceCards({ color, onTom, onFraMal }: { color: string; onTom: () => void; onFraMal: () => void }) {
  return (
    <>
      <p style={{ color: "var(--bfc-base-c-2)", margin: 0, fontSize: "0.9rem" }}>Velg utgangspunkt:</p>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0.75rem" }}>
        {(["tom", "mal"] as const).map((opt) => (
          <button
            key={opt}
            onClick={opt === "tom" ? onTom : onFraMal}
            style={{ border: `2px solid ${color}40`, borderRadius: 10, padding: "1.25rem 1rem", background: `${color}08`, cursor: "pointer", textAlign: "center", transition: "border-color 0.15s, background 0.15s" }}
            onMouseEnter={(e) => { (e.currentTarget as HTMLButtonElement).style.borderColor = color; (e.currentTarget as HTMLButtonElement).style.background = `${color}14`; }}
            onMouseLeave={(e) => { (e.currentTarget as HTMLButtonElement).style.borderColor = `${color}40`; (e.currentTarget as HTMLButtonElement).style.background = `${color}08`; }}
          >
            <div style={{ fontWeight: 700, fontSize: "1.4rem", color, marginBottom: "0.4rem" }}>{opt === "tom" ? "☐" : "≡"}</div>
            <div style={{ fontWeight: 600, fontSize: "0.9rem", color }}>{opt === "tom" ? "Ny tom" : "Fra mal"}</div>
            <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)", marginTop: 2 }}>
              {opt === "tom" ? "Start fra bunnen av" : "Bruk en eksisterende mal"}
            </div>
          </button>
        ))}
      </div>
    </>
  );
}

export default function ProjectDetailPage() {
  const { projectId } = useParams<{ projectId: string }>();
  const navigate = useNavigate();
  const [project, setProject] = useState<Project | null>(null);
  const [riskMatrices, setRiskMatrices] = useState<RiskMatrix[]>([]);
  const [commPlans, setCommPlans] = useState<CommunicationPlan[]>([]);
  const [meetingPlans, setMeetingPlans] = useState<MeetingPlan[]>([]);
  const [runbooks, setRunbooks] = useState<Runbook[]>([]);
  const [projectPlans, setProjectPlans] = useState<ProjectPlan[]>([]);
  const [oppgaveLister, setOppgaveLister] = useState<OppgaveListe[]>([]);

  // Simple modal (risk / comm / meeting)
  const [newSimpleType, setNewSimpleType] = useState<"risk" | "comm" | "meeting" | null>(null);
  const [simpleStep, setSimpleStep] = useState<"choice" | "template" | "config">("choice");
  const [simpleTitle, setSimpleTitle] = useState("");
  const [simpleSelectedTemplate, setSimpleSelectedTemplate] = useState<Template | null>(null);
  const [simpleTemplates, setSimpleTemplates] = useState<Template[]>([]);
  const [simpleSaving, setSimpleSaving] = useState(false);

  // Oppgaveliste modal
  const [newOppgaveModal, setNewOppgaveModal] = useState(false);
  const [olSource, setOlSource] = useState<"own" | "planner" | "smartsheet" | null>(null);
  const [olTitle, setOlTitle] = useState("");
  const [olUrl, setOlUrl] = useState("");
  const [olSaving, setOlSaving] = useState(false);
  const [olOwnStep, setOlOwnStep] = useState<OwnStep>("choice");

  // Runbook modal
  const [newRunbookModal, setNewRunbookModal] = useState(false);
  const [rbSource, setRbSource] = useState<"own" | "planner" | "smartsheet" | null>(null);
  const [rbTitle, setRbTitle] = useState("");
  const [rbUrl, setRbUrl] = useState("");
  const [rbSaving, setRbSaving] = useState(false);
  const [rbOwnStep, setRbOwnStep] = useState<OwnStep>("choice");

  // Project plan modal
  const [newPlanModal, setNewPlanModal] = useState(false);
  const [ppSource, setPpSource] = useState<"own" | "planner" | "smartsheet" | null>(null);
  const [ppTitle, setPpTitle] = useState("");
  const [ppUrl, setPpUrl] = useState("");
  const [ppSaving, setPpSaving] = useState(false);
  const [ppOwnStep, setPpOwnStep] = useState<OwnStep>("choice");

  // Shared own-template state (only one "own" modal open at a time)
  const [ownSelectedTemplate, setOwnSelectedTemplate] = useState<Template | null>(null);
  const [ownTemplates, setOwnTemplates] = useState<Template[]>([]);

  const [activeView, setActiveView] = useState<"dashboard" | "innhold">("dashboard");

  const [renameTarget, setRenameTarget] = useState<{ id: string; title: string; type: PlanType } | null>(null);
  const [renameValue, setRenameValue] = useState("");
  const [renameSaving, setRenameSaving] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<{ id: string; title: string; type: PlanType } | null>(null);
  const [deleting, setDeleting] = useState(false);

  const [editProjectModal, setEditProjectModal] = useState(false);
  const [editProjectManager, setEditProjectManager] = useState("");
  const [editProjectStatus, setEditProjectStatus] = useState<"active" | "not_started" | "completed">("active");
  const [editProjectSaving, setEditProjectSaving] = useState(false);

  useEffect(() => {
    if (!projectId) return;
    Promise.all([
      api.projects.get(projectId),
      api.riskMatrices.list(projectId),
      api.communicationPlans.list(projectId),
      api.meetingPlans.list(projectId),
      api.runbooks.list(projectId),
      api.projectPlans.list(projectId),
      api.oppgaveLister.list(projectId),
    ]).then(([p, rm, cp, mp, rb, pp, ol]) => {
      setProject(p); setRiskMatrices(rm); setCommPlans(cp); setMeetingPlans(mp);
      setRunbooks(rb); setProjectPlans(pp); setOppgaveLister(ol);
    });
  }, [projectId]);

  // ─── Simple modal helpers ─────────────────────────────────────────────────────

  function openSimpleModal(type: "risk" | "comm" | "meeting") {
    setNewSimpleType(type);
    setSimpleStep("choice");
    setSimpleTitle("");
    setSimpleSelectedTemplate(null);
    setSimpleTemplates([]);
  }

  async function goToSimpleTemplate() {
    if (!newSimpleType) return;
    const typeKey = newSimpleType === "risk" ? "risk_matrix" : newSimpleType === "comm" ? "communication_plan" : "meeting_plan";
    const ts = await api.templates.list(typeKey);
    setSimpleTemplates(ts);
    setSimpleStep("template");
  }

  async function createSimple() {
    if (!projectId || !newSimpleType) return;
    setSimpleSaving(true);
    try {
      const defaultTitles = { risk: "Ny risikomatrise", comm: "Ny kommunikasjonsplan", meeting: "Ny møteplan" };
      const title = simpleTitle.trim() || defaultTitles[newSimpleType];
      if (newSimpleType === "risk") {
        const m = await api.riskMatrices.create(projectId, { title });
        if (simpleSelectedTemplate) {
          const d = JSON.parse(simpleSelectedTemplate.data) as { risks?: Array<{ description: string; probability: number; consequence: number; mitigation?: string | null; owner?: string | null; status?: string }> };
          for (const r of d.risks ?? []) {
            await api.riskMatrices.addRisk(projectId, m.id, { description: r.description, probability: r.probability, consequence: r.consequence, mitigation: r.mitigation ?? null, owner: r.owner ?? null, status: (r.status ?? "open") as "open" | "mitigated" | "closed", fagomrade: null, risk_owner: null, residual_probability: null, residual_consequence: null });
          }
        }
        navigate(`/projects/${projectId}/risk-matrix/${m.id}`);
      } else if (newSimpleType === "comm") {
        const p = await api.communicationPlans.create(projectId, { title });
        if (simpleSelectedTemplate) {
          const d = JSON.parse(simpleSelectedTemplate.data) as { entries?: Array<{ stakeholder: string; message: string; channel: string; frequency: string; responsible: string }> };
          for (const e of d.entries ?? []) await api.communicationPlans.addEntry(projectId, p.id, e);
        }
        navigate(`/projects/${projectId}/communication-plan/${p.id}`);
      } else {
        const p = await api.meetingPlans.create(projectId, { title });
        if (simpleSelectedTemplate) {
          const d = JSON.parse(simpleSelectedTemplate.data) as { meetings?: Array<{ title: string; date: string; purpose?: string | null }> };
          for (const mt of d.meetings ?? []) await api.meetingPlans.addMeeting(projectId, p.id, { title: mt.title, date: mt.date, purpose: mt.purpose ?? null, outlook_id: null });
        }
        navigate(`/projects/${projectId}/meeting-plan/${p.id}`);
      }
    } finally {
      setSimpleSaving(false);
    }
  }

  async function loadOwnTemplates(type: string) {
    const ts = await api.templates.list(type);
    setOwnTemplates(ts);
  }

  // ─── Create functions ─────────────────────────────────────────────────────────

  async function handleRename() {
    if (!projectId || !renameTarget || !renameValue.trim()) return;
    setRenameSaving(true);
    try {
      if (renameTarget.type === "risk") {
        const u = await api.riskMatrices.update(projectId, renameTarget.id, { title: renameValue });
        setRiskMatrices((prev) => prev.map((m) => (m.id === u.id ? u : m)));
      } else if (renameTarget.type === "comm") {
        const u = await api.communicationPlans.update(projectId, renameTarget.id, { title: renameValue });
        setCommPlans((prev) => prev.map((p) => (p.id === u.id ? u : p)));
      } else if (renameTarget.type === "meeting") {
        const u = await api.meetingPlans.update(projectId, renameTarget.id, { title: renameValue });
        setMeetingPlans((prev) => prev.map((p) => (p.id === u.id ? u : p)));
      } else if (renameTarget.type === "projectplan") {
        const u = await api.projectPlans.update(projectId, renameTarget.id, { title: renameValue });
        setProjectPlans((prev) => prev.map((p) => (p.id === u.id ? u : p)));
      } else if (renameTarget.type === "oppgave") {
        const u = await api.oppgaveLister.update(projectId, renameTarget.id, { title: renameValue });
        setOppgaveLister((prev) => prev.map((p) => (p.id === u.id ? u : p)));
      } else {
        const u = await api.runbooks.update(projectId, renameTarget.id, { title: renameValue });
        setRunbooks((prev) => prev.map((p) => (p.id === u.id ? u : p)));
      }
      setRenameTarget(null);
    } finally {
      setRenameSaving(false);
    }
  }

  async function createProjectPlan() {
    if (!projectId || !ppSource) return;
    setPpSaving(true);
    try {
      const title = ppTitle.trim() || (ppSource === "own" ? "Ny prosjektplan" : ppSource === "planner" ? "Planner-prosjektplan" : "Smartsheet-prosjektplan");
      const pp = await api.projectPlans.create(projectId, { title, source: ppSource, external_url: ppUrl.trim() || undefined });
      if (ppSource === "own" && ownSelectedTemplate) {
        const d = JSON.parse(ownSelectedTemplate.data) as { tasks?: Array<{ name: string; bucket?: string | null; percent_complete?: number; start_date?: string | null; end_date?: string | null; responsible?: string | null; description?: string | null }> };
        for (const t of d.tasks ?? []) {
          await api.projectPlans.addTask(projectId, pp.id, { name: t.name, bucket: t.bucket ?? null, percent_complete: t.percent_complete ?? 0, start_date: t.start_date ?? null, end_date: t.end_date ?? null, responsible: t.responsible ?? null, description: t.description ?? null });
        }
      }
      navigate(`/projects/${projectId}/project-plan/${pp.id}`);
    } finally {
      setPpSaving(false);
    }
  }

  async function createOppgaveListe() {
    if (!projectId || !olSource) return;
    setOlSaving(true);
    try {
      const title = olTitle.trim() || (olSource === "own" ? "Ny oppgaveliste" : olSource === "planner" ? "Planner-oppgaver" : "Smartsheet-oppgaver");
      const ol = await api.oppgaveLister.create(projectId, { title, source: olSource, external_url: olUrl.trim() || undefined });
      if (olSource === "own" && ownSelectedTemplate) {
        const d = JSON.parse(ownSelectedTemplate.data) as { oppgaver?: Array<{ name: string; responsible?: string | null; due_date?: string | null; status?: string; description?: string | null }> };
        for (const o of d.oppgaver ?? []) {
          await api.oppgaveLister.addOppgave(projectId, ol.id, { name: o.name, responsible: o.responsible ?? null, due_date: o.due_date ?? null, status: (o.status ?? "not_started") as "not_started" | "in_progress" | "done", description: o.description ?? null });
        }
      }
      navigate(`/projects/${projectId}/oppgave/${ol.id}`);
    } finally {
      setOlSaving(false);
    }
  }

  async function createRunbook() {
    if (!projectId || !rbSource) return;
    setRbSaving(true);
    try {
      const title = rbTitle.trim() || (rbSource === "own" ? "Ny runbook" : rbSource === "planner" ? "Planner-runbook" : "Smartsheet-runbook");
      const rb = await api.runbooks.create(projectId, { title, source: rbSource, external_url: rbUrl.trim() || undefined });
      if (rbSource === "own" && ownSelectedTemplate) {
        const d = JSON.parse(ownSelectedTemplate.data) as { activities?: Array<{ name: string; phase?: string | null; status?: string; start_date?: string | null; end_date?: string | null; responsible?: string | null; description?: string | null }> };
        for (const a of d.activities ?? []) {
          await api.runbooks.addActivity(projectId, rb.id, { name: a.name, phase: a.phase ?? null, status: (a.status ?? "not_started") as "not_started" | "in_progress" | "done" | "cancelled", start_date: a.start_date ?? null, end_date: a.end_date ?? null, responsible: a.responsible ?? null, description: a.description ?? null });
        }
      }
      navigate(`/projects/${projectId}/runbook/${rb.id}`);
    } finally {
      setRbSaving(false);
    }
  }

  async function handleDelete() {
    if (!projectId || !deleteTarget) return;
    setDeleting(true);
    try {
      if (deleteTarget.type === "risk") { await api.riskMatrices.delete(projectId, deleteTarget.id); setRiskMatrices((prev) => prev.filter((m) => m.id !== deleteTarget.id)); }
      else if (deleteTarget.type === "comm") { await api.communicationPlans.delete(projectId, deleteTarget.id); setCommPlans((prev) => prev.filter((p) => p.id !== deleteTarget.id)); }
      else if (deleteTarget.type === "meeting") { await api.meetingPlans.delete(projectId, deleteTarget.id); setMeetingPlans((prev) => prev.filter((p) => p.id !== deleteTarget.id)); }
      else if (deleteTarget.type === "projectplan") { await api.projectPlans.delete(projectId, deleteTarget.id); setProjectPlans((prev) => prev.filter((p) => p.id !== deleteTarget.id)); }
      else if (deleteTarget.type === "oppgave") { await api.oppgaveLister.delete(projectId, deleteTarget.id); setOppgaveLister((prev) => prev.filter((p) => p.id !== deleteTarget.id)); }
      else { await api.runbooks.delete(projectId, deleteTarget.id); setRunbooks((prev) => prev.filter((p) => p.id !== deleteTarget.id)); }
      setDeleteTarget(null);
    } finally {
      setDeleting(false);
    }
  }

  async function handleSetPrimary(type: PlanType, id: string) {
    if (!projectId) return;
    if (type === "projectplan") {
      const u = await api.projectPlans.setPrimary(projectId, id);
      setProjectPlans((prev) => prev.map((p) => ({ ...p, is_primary: p.id === u.id })));
    } else if (type === "oppgave") {
      const u = await api.oppgaveLister.setPrimary(projectId, id);
      setOppgaveLister((prev) => prev.map((p) => ({ ...p, is_primary: p.id === u.id })));
    } else if (type === "risk") {
      const u = await api.riskMatrices.setPrimary(projectId, id);
      setRiskMatrices((prev) => prev.map((m) => ({ ...m, is_primary: m.id === u.id })));
    } else if (type === "meeting") {
      const u = await api.meetingPlans.setPrimary(projectId, id);
      setMeetingPlans((prev) => prev.map((p) => ({ ...p, is_primary: p.id === u.id })));
    }
  }

  async function handleEditProject() {
    if (!projectId || !project) return;
    setEditProjectSaving(true);
    try {
      const updated = await api.projects.update(projectId, {
        name: project.name,
        description: project.description ?? undefined,
        project_manager: editProjectManager.trim() || null,
        status: editProjectStatus,
      });
      setProject(updated);
      setEditProjectModal(false);
    } finally {
      setEditProjectSaving(false);
    }
  }

  if (!project) return <div style={{ padding: "2rem" }}>Laster...</div>;

  const color = accentColor(project.name);

  const sourcePickerCfg = {
    planner:    { color: "#0078D4", label: "Microsoft Planner", sub: "Lenk til eksisterende plan", initial: "P" },
    smartsheet: { color: "#00A88E", label: "Smartsheet",        sub: "Lenk til eksisterende plan", initial: "S" },
  } as const;

  return (
    <div style={{ maxWidth: activeView === "dashboard" ? 1200 : 900, margin: "0 auto", padding: "2rem" }}>
      <button onClick={() => navigate("/projects")} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", marginBottom: "1.5rem", fontSize: "0.9rem", padding: 0 }}>
        ← Tilbake til prosjekter
      </button>

      <div style={{ display: "flex", alignItems: "center", gap: "1.25rem", padding: "1.5rem", borderRadius: 10, background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)", borderTop: `4px solid ${color}`, marginBottom: "2.5rem" }}>
        <div style={{ width: 56, height: 56, borderRadius: "50%", background: color, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: "1.1rem", flexShrink: 0 }}>
          {initials(project.name)}
        </div>
        <div style={{ flex: 1 }}>
          <h1 className="bf-h2" style={{ margin: 0 }}>{project.name}</h1>
          {project.description && <p style={{ color: "var(--bfc-base-c-2)", margin: "0.25rem 0 0", fontSize: "0.95rem" }}>{project.description}</p>}
          {project.project_manager && (
            <p style={{ color: "var(--bfc-base-c-2)", margin: "0.25rem 0 0", fontSize: "0.85rem" }}>
              <span style={{ fontWeight: 600 }}>Prosjektleder:</span> {project.project_manager}
            </p>
          )}
        </div>
        <button
          onClick={() => { setEditProjectManager(project.project_manager ?? ""); setEditProjectStatus(project.status ?? "active"); setEditProjectModal(true); }}
          style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", fontSize: "0.82rem", padding: "4px 8px", borderRadius: 4, flexShrink: 0 }}
          title="Rediger prosjektleder"
        >
          ✎ Rediger
        </button>
      </div>

      {/* Tab switcher */}
      <div style={{ display: "flex", gap: 0, marginBottom: "1.75rem", borderBottom: "1px solid var(--bfc-base-dimmed)" }}>
        {(["dashboard", "innhold"] as const).map((view) => (
          <button key={view} onClick={() => setActiveView(view)} style={{
            background: "none", border: "none", cursor: "pointer",
            padding: "0.6rem 1.4rem",
            fontSize: "0.92rem", fontWeight: activeView === view ? 600 : 400,
            color: activeView === view ? color : "var(--bfc-base-c-2)",
            borderBottom: activeView === view ? `2px solid ${color}` : "2px solid transparent",
            marginBottom: -1, transition: "color 0.15s",
          }}>
            {view === "dashboard" ? "Dashboard" : "Innhold"}
          </button>
        ))}
      </div>

      {activeView === "dashboard" ? (
        <DashboardView
          project={project}
          riskMatrices={riskMatrices}
          projectPlans={projectPlans}
          oppgaveLister={oppgaveLister}
          runbooks={runbooks}
          meetingPlans={meetingPlans}
          projectId={projectId!}
          navigate={navigate}
        />
      ) : (<>

      <Section type="projectplan" items={projectPlans}
        onNew={() => { setPpSource(null); setPpTitle(""); setPpUrl(""); setPpOwnStep("choice"); setOwnSelectedTemplate(null); setNewPlanModal(true); }}
        onOpen={(id) => navigate(`/projects/${projectId}/project-plan/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "projectplan" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "projectplan" })}
        onSetPrimary={(item) => handleSetPrimary("projectplan", item.id)}
        countLabel={(pp) => { const src = (pp as ProjectPlan).source; return src === "planner" ? "Planner" : src === "smartsheet" ? "Smartsheet" : `${(pp as ProjectPlan).tasks.length} oppgaver`; }}
      />
      <Section type="oppgave" items={oppgaveLister}
        onNew={() => { setOlSource(null); setOlTitle(""); setOlUrl(""); setOlOwnStep("choice"); setOwnSelectedTemplate(null); setNewOppgaveModal(true); }}
        onOpen={(id) => navigate(`/projects/${projectId}/oppgave/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "oppgave" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "oppgave" })}
        onSetPrimary={(item) => handleSetPrimary("oppgave", item.id)}
        countLabel={(ol) => { const src = (ol as OppgaveListe).source; return src === "planner" ? "Planner" : src === "smartsheet" ? "Smartsheet" : `${(ol as OppgaveListe).oppgaver.length} oppgaver`; }}
      />
      <Section type="risk" items={riskMatrices}
        onNew={() => openSimpleModal("risk")}
        onOpen={(id) => navigate(`/projects/${projectId}/risk-matrix/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "risk" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "risk" })}
        onSetPrimary={(item) => handleSetPrimary("risk", item.id)}
        countLabel={(m) => `${(m as RiskMatrix).risks.length} risikoer`}
      />
      <Section type="comm" items={commPlans}
        onNew={() => openSimpleModal("comm")}
        onOpen={(id) => navigate(`/projects/${projectId}/communication-plan/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "comm" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "comm" })}
        countLabel={(p) => `${(p as CommunicationPlan).entries.length} oppføringer`}
      />
      <Section type="meeting" items={meetingPlans}
        onNew={() => openSimpleModal("meeting")}
        onOpen={(id) => navigate(`/projects/${projectId}/meeting-plan/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "meeting" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "meeting" })}
        onSetPrimary={(item) => handleSetPrimary("meeting", item.id)}
        countLabel={(p) => `${(p as MeetingPlan).meetings.length} møter`}
      />
      <Section type="runbook" items={runbooks}
        onNew={() => { setRbSource(null); setRbTitle(""); setRbUrl(""); setRbOwnStep("choice"); setOwnSelectedTemplate(null); setNewRunbookModal(true); }}
        onOpen={(id) => navigate(`/projects/${projectId}/runbook/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "runbook" as PlanType }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "runbook" as PlanType })}
        countLabel={(rb) => { const src = (rb as Runbook).source; return src === "planner" ? "Planner" : src === "smartsheet" ? "Smartsheet" : `${(rb as Runbook).activities.length} aktiviteter`; }}
      />

      </>)} {/* end innhold */}

      {/* ─── Simple modal (Risk / Comm / Meeting) ─────────────────────────────── */}
      <Modal
        isOpen={newSimpleType !== null}
        onRequestClose={() => setNewSimpleType(null)}
        header={newSimpleType === "risk" ? "Ny risikomatrise" : newSimpleType === "comm" ? "Ny kommunikasjonsplan" : "Ny møteplan"}
      >
        {simpleStep === "choice" && (() => {
          const sColor = newSimpleType === "risk" ? "#E03131" : newSimpleType === "comm" ? "#1971C2" : "#2F9E44";
          return (
            <OwnChoiceCards color={sColor}
              onTom={() => setSimpleStep("config")}
              onFraMal={goToSimpleTemplate}
            />
          );
        })()}
        {simpleStep === "template" && (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => { setSimpleStep("choice"); setSimpleSelectedTemplate(null); }} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <p style={{ margin: 0, fontWeight: 600, fontSize: "0.9rem" }}>Velg mal:</p>
            <TemplatePicker templates={simpleTemplates} selected={simpleSelectedTemplate} onSelect={setSimpleSelectedTemplate} />
            <Input label="Navn" value={simpleTitle} onChange={(e) => setSimpleTitle(e.target.value)} placeholder={newSimpleType === "risk" ? "f.eks. Risikomatrise Migration" : newSimpleType === "comm" ? "f.eks. Kommunikasjonsplan Farvatn" : "f.eks. Møteplan Q1"} />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewSimpleType(null)}>Avbryt</Button>
              <Button variant="filled" onClick={createSimple} state={simpleSaving || !simpleSelectedTemplate ? "inactive" : "default"}>{simpleSaving ? "Oppretter..." : "Opprett fra mal"}</Button>
            </div>
          </div>
        )}
        {simpleStep === "config" && (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => setSimpleStep("choice")} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <Input label="Navn" value={simpleTitle} onChange={(e) => setSimpleTitle(e.target.value)} placeholder={newSimpleType === "risk" ? "f.eks. Risikomatrise Migration" : newSimpleType === "comm" ? "f.eks. Kommunikasjonsplan Farvatn" : "f.eks. Møteplan Q1"} autoFocus />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewSimpleType(null)}>Avbryt</Button>
              <Button variant="filled" onClick={createSimple} state={simpleSaving ? "inactive" : "default"}>{simpleSaving ? "Oppretter..." : "Opprett og åpne"}</Button>
            </div>
          </div>
        )}
      </Modal>

      {/* ─── New oppgaveliste modal ────────────────────────────────────────────── */}
      <Modal isOpen={newOppgaveModal} onRequestClose={() => setNewOppgaveModal(false)} header="Ny oppgaveliste">
        {olSource === null ? (
          <div>
            <p style={{ color: "var(--bfc-base-c-2)", marginBottom: "1.25rem", fontSize: "0.9rem" }}>Velg hvor oppgavelisten skal hentes fra:</p>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0.75rem" }}>
              {(["planner", "smartsheet", "own"] as const).map((src) => {
                const cfg = src === "own" ? { color: "#F59F00", label: "Opprett egen", sub: "Legg til oppgaver manuelt", initial: "+" } : { ...sourcePickerCfg[src], initial: sourcePickerCfg[src].initial };
                return (
                  <button key={src} onClick={() => setOlSource(src)}
                    style={{ border: `2px solid ${cfg.color}40`, borderRadius: 10, padding: "1.25rem 1rem", background: `${cfg.color}08`, cursor: "pointer", textAlign: "center", transition: "border-color 0.15s, background 0.15s" }}
                    onMouseEnter={(e) => { (e.currentTarget as HTMLButtonElement).style.borderColor = cfg.color; (e.currentTarget as HTMLButtonElement).style.background = `${cfg.color}14`; }}
                    onMouseLeave={(e) => { (e.currentTarget as HTMLButtonElement).style.borderColor = `${cfg.color}40`; (e.currentTarget as HTMLButtonElement).style.background = `${cfg.color}08`; }}
                  >
                    <div style={{ width: 40, height: 40, borderRadius: "50%", background: cfg.color, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: "1.1rem", margin: "0 auto 0.75rem" }}>{cfg.initial}</div>
                    <div style={{ fontWeight: 600, fontSize: "0.85rem", color: cfg.color, marginBottom: "0.3rem" }}>{cfg.label}</div>
                    <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)" }}>{cfg.sub}</div>
                  </button>
                );
              })}
            </div>
          </div>
        ) : olSource !== "own" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => setOlSource(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <Input label="Navn på oppgaveliste" value={olTitle} onChange={(e) => setOlTitle(e.target.value)} placeholder="f.eks. Oppgaver" autoFocus />
            <Input label={`URL til ${olSource === "planner" ? "Microsoft Planner" : "Smartsheet"}-planen`} value={olUrl} onChange={(e) => setOlUrl(e.target.value)} placeholder="https://..." />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewOppgaveModal(false)}>Avbryt</Button>
              <Button variant="filled" onClick={createOppgaveListe} state={olSaving ? "inactive" : "default"}>{olSaving ? "Oppretter..." : "Lagre"}</Button>
            </div>
          </div>
        ) : olOwnStep === "choice" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => setOlSource(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <OwnChoiceCards color="#F59F00" onTom={() => setOlOwnStep("config")} onFraMal={async () => { await loadOwnTemplates("oppgave_liste"); setOlOwnStep("template"); }} />
          </div>
        ) : olOwnStep === "template" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => { setOlOwnStep("choice"); setOwnSelectedTemplate(null); }} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <p style={{ margin: 0, fontWeight: 600, fontSize: "0.9rem" }}>Velg mal:</p>
            <TemplatePicker templates={ownTemplates} selected={ownSelectedTemplate} onSelect={setOwnSelectedTemplate} />
            <Input label="Navn på oppgaveliste" value={olTitle} onChange={(e) => setOlTitle(e.target.value)} placeholder="f.eks. Handlingspunkter" />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewOppgaveModal(false)}>Avbryt</Button>
              <Button variant="filled" onClick={createOppgaveListe} state={olSaving || !ownSelectedTemplate ? "inactive" : "default"}>{olSaving ? "Oppretter..." : "Opprett fra mal"}</Button>
            </div>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => setOlOwnStep("choice")} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <Input label="Navn på oppgaveliste" value={olTitle} onChange={(e) => setOlTitle(e.target.value)} placeholder="f.eks. Handlingspunkter" autoFocus />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewOppgaveModal(false)}>Avbryt</Button>
              <Button variant="filled" onClick={createOppgaveListe} state={olSaving ? "inactive" : "default"}>{olSaving ? "Oppretter..." : "Opprett og åpne"}</Button>
            </div>
          </div>
        )}
      </Modal>

      {/* ─── New runbook modal ─────────────────────────────────────────────────── */}
      <Modal isOpen={newRunbookModal} onRequestClose={() => setNewRunbookModal(false)} header="Ny runbook">
        {rbSource === null ? (
          <div>
            <p style={{ color: "var(--bfc-base-c-2)", marginBottom: "1.25rem", fontSize: "0.9rem" }}>Velg hvor runbooken skal hentes fra:</p>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0.75rem" }}>
              {(["planner", "smartsheet", "own"] as const).map((src) => {
                const cfg = src === "own" ? { color: "#7950F2", label: "Opprett egen", sub: "Legg til aktiviteter manuelt", initial: "+" } : { ...sourcePickerCfg[src], initial: sourcePickerCfg[src].initial };
                return (
                  <button key={src} onClick={() => setRbSource(src)}
                    style={{ border: `2px solid ${cfg.color}40`, borderRadius: 10, padding: "1.25rem 1rem", background: `${cfg.color}08`, cursor: "pointer", textAlign: "center", transition: "border-color 0.15s, background 0.15s" }}
                    onMouseEnter={(e) => { (e.currentTarget as HTMLButtonElement).style.borderColor = cfg.color; (e.currentTarget as HTMLButtonElement).style.background = `${cfg.color}14`; }}
                    onMouseLeave={(e) => { (e.currentTarget as HTMLButtonElement).style.borderColor = `${cfg.color}40`; (e.currentTarget as HTMLButtonElement).style.background = `${cfg.color}08`; }}
                  >
                    <div style={{ width: 40, height: 40, borderRadius: "50%", background: cfg.color, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: "1.1rem", margin: "0 auto 0.75rem" }}>{cfg.initial}</div>
                    <div style={{ fontWeight: 600, fontSize: "0.85rem", color: cfg.color, marginBottom: "0.3rem" }}>{cfg.label}</div>
                    <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)" }}>{cfg.sub}</div>
                  </button>
                );
              })}
            </div>
          </div>
        ) : rbSource !== "own" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => setRbSource(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <Input label="Navn på runbook" value={rbTitle} onChange={(e) => setRbTitle(e.target.value)} placeholder="f.eks. Farvatn Runbook" autoFocus />
            <Input label={`URL til ${rbSource === "planner" ? "Microsoft Planner" : "Smartsheet"}-planen`} value={rbUrl} onChange={(e) => setRbUrl(e.target.value)} placeholder="https://..." />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewRunbookModal(false)}>Avbryt</Button>
              <Button variant="filled" onClick={createRunbook} state={rbSaving ? "inactive" : "default"}>{rbSaving ? "Oppretter..." : "Lagre"}</Button>
            </div>
          </div>
        ) : rbOwnStep === "choice" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => setRbSource(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <OwnChoiceCards color="#7950F2" onTom={() => setRbOwnStep("config")} onFraMal={async () => { await loadOwnTemplates("runbook"); setRbOwnStep("template"); }} />
          </div>
        ) : rbOwnStep === "template" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => { setRbOwnStep("choice"); setOwnSelectedTemplate(null); }} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <p style={{ margin: 0, fontWeight: 600, fontSize: "0.9rem" }}>Velg mal:</p>
            <TemplatePicker templates={ownTemplates} selected={ownSelectedTemplate} onSelect={setOwnSelectedTemplate} />
            <Input label="Navn på runbook" value={rbTitle} onChange={(e) => setRbTitle(e.target.value)} placeholder="f.eks. Cutover-runbook" />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewRunbookModal(false)}>Avbryt</Button>
              <Button variant="filled" onClick={createRunbook} state={rbSaving || !ownSelectedTemplate ? "inactive" : "default"}>{rbSaving ? "Oppretter..." : "Opprett fra mal"}</Button>
            </div>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => setRbOwnStep("choice")} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <Input label="Navn på runbook" value={rbTitle} onChange={(e) => setRbTitle(e.target.value)} placeholder="f.eks. Cutover-runbook" autoFocus />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewRunbookModal(false)}>Avbryt</Button>
              <Button variant="filled" onClick={createRunbook} state={rbSaving ? "inactive" : "default"}>{rbSaving ? "Oppretter..." : "Opprett og åpne"}</Button>
            </div>
          </div>
        )}
      </Modal>

      {/* ─── New project plan modal ───────────────────────────────────────────── */}
      <Modal isOpen={newPlanModal} onRequestClose={() => setNewPlanModal(false)} header="Ny prosjektplan">
        {ppSource === null ? (
          <div>
            <p style={{ color: "var(--bfc-base-c-2)", marginBottom: "1.25rem", fontSize: "0.9rem" }}>Velg hvor prosjektplanen skal hentes fra:</p>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0.75rem" }}>
              {(["planner", "smartsheet", "own"] as const).map((src) => {
                const cfg = src === "own" ? { color: "#0CA678", label: "Opprett egen", sub: "Legg til oppgaver manuelt", initial: "+" } : { ...sourcePickerCfg[src], initial: sourcePickerCfg[src].initial };
                return (
                  <button key={src} onClick={() => setPpSource(src)}
                    style={{ border: `2px solid ${cfg.color}40`, borderRadius: 10, padding: "1.25rem 1rem", background: `${cfg.color}08`, cursor: "pointer", textAlign: "center", transition: "border-color 0.15s, background 0.15s" }}
                    onMouseEnter={(e) => { (e.currentTarget as HTMLButtonElement).style.borderColor = cfg.color; (e.currentTarget as HTMLButtonElement).style.background = `${cfg.color}14`; }}
                    onMouseLeave={(e) => { (e.currentTarget as HTMLButtonElement).style.borderColor = `${cfg.color}40`; (e.currentTarget as HTMLButtonElement).style.background = `${cfg.color}08`; }}
                  >
                    <div style={{ width: 40, height: 40, borderRadius: "50%", background: cfg.color, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: "1.1rem", margin: "0 auto 0.75rem" }}>{cfg.initial}</div>
                    <div style={{ fontWeight: 600, fontSize: "0.85rem", color: cfg.color, marginBottom: "0.3rem" }}>{cfg.label}</div>
                    <div style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)" }}>{cfg.sub}</div>
                  </button>
                );
              })}
            </div>
          </div>
        ) : ppSource !== "own" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => setPpSource(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <Input label="Navn på prosjektplan" value={ppTitle} onChange={(e) => setPpTitle(e.target.value)} placeholder="f.eks. Migration Services Plan" autoFocus />
            <Input label={`URL til ${ppSource === "planner" ? "Microsoft Planner" : "Smartsheet"}-planen`} value={ppUrl} onChange={(e) => setPpUrl(e.target.value)} placeholder="https://..." />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewPlanModal(false)}>Avbryt</Button>
              <Button variant="filled" onClick={createProjectPlan} state={ppSaving ? "inactive" : "default"}>{ppSaving ? "Oppretter..." : "Lagre"}</Button>
            </div>
          </div>
        ) : ppOwnStep === "choice" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => setPpSource(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <OwnChoiceCards color="#0CA678" onTom={() => setPpOwnStep("config")} onFraMal={async () => { await loadOwnTemplates("project_plan"); setPpOwnStep("template"); }} />
          </div>
        ) : ppOwnStep === "template" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => { setPpOwnStep("choice"); setOwnSelectedTemplate(null); }} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <p style={{ margin: 0, fontWeight: 600, fontSize: "0.9rem" }}>Velg mal:</p>
            <TemplatePicker templates={ownTemplates} selected={ownSelectedTemplate} onSelect={setOwnSelectedTemplate} />
            <Input label="Navn på prosjektplan" value={ppTitle} onChange={(e) => setPpTitle(e.target.value)} placeholder="f.eks. Prosjektplan Migration Services" />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewPlanModal(false)}>Avbryt</Button>
              <Button variant="filled" onClick={createProjectPlan} state={ppSaving || !ownSelectedTemplate ? "inactive" : "default"}>{ppSaving ? "Oppretter..." : "Opprett fra mal"}</Button>
            </div>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
            <button onClick={() => setPpOwnStep("choice")} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", alignSelf: "flex-start", padding: 0, fontSize: "0.9rem" }}>← Tilbake</button>
            <Input label="Navn på prosjektplan" value={ppTitle} onChange={(e) => setPpTitle(e.target.value)} placeholder="f.eks. Prosjektplan Migration Services" autoFocus />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button onClick={() => setNewPlanModal(false)}>Avbryt</Button>
              <Button variant="filled" onClick={createProjectPlan} state={ppSaving ? "inactive" : "default"}>{ppSaving ? "Oppretter..." : "Opprett og åpne"}</Button>
            </div>
          </div>
        )}
      </Modal>

      {/* ─── Rename modal ─────────────────────────────────────────────────────── */}
      <Modal isOpen={!!renameTarget} onRequestClose={() => setRenameTarget(null)} header="Gi nytt navn">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <Input label="Navn" value={renameValue} onChange={(e) => setRenameValue(e.target.value)} onKeyDown={(e) => e.key === "Enter" && handleRename()} autoFocus />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setRenameTarget(null)}>Avbryt</Button>
            <Button variant="filled" onClick={handleRename} state={!renameValue.trim() || renameSaving ? "inactive" : "default"}>{renameSaving ? "Lagrer..." : "Lagre"}</Button>
          </div>
        </div>
      </Modal>

      {/* ─── Delete confirmation ───────────────────────────────────────────────── */}
      <Modal isOpen={!!deleteTarget} onRequestClose={() => setDeleteTarget(null)} header="Bekreft sletting">
        <div style={{ display: "flex", flexDirection: "column", gap: "1.25rem" }}>
          <p style={{ margin: 0 }}>Er du sikker på at du vil slette <strong>«{deleteTarget?.title}»</strong>? Dette kan ikke angres.</p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setDeleteTarget(null)}>Avbryt</Button>
            <Button variant="filled" state={deleting ? "inactive" : "default"} onClick={handleDelete} style={{ background: "#E03131", borderColor: "#E03131" }}>
              {deleting ? "Sletter..." : "Ja, slett"}
            </Button>
          </div>
        </div>
      </Modal>

      {/* ─── Edit project modal ───────────────────────────────────────────────── */}
      <Modal isOpen={editProjectModal} onRequestClose={() => setEditProjectModal(false)} header="Rediger prosjekt">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <Input
            label="Prosjektleder"
            value={editProjectManager}
            onChange={(e) => setEditProjectManager(e.target.value)}
            placeholder="Navn på prosjektleder"
            autoFocus
          />
          <div>
            <label style={{ display: "block", fontSize: "0.85rem", fontWeight: 600, marginBottom: "0.4rem", color: "var(--bfc-base-c-1)" }}>
              Status
            </label>
            <div style={{ display: "flex", gap: "0.5rem" }}>
              {(["not_started", "active", "completed"] as const).map((s) => {
                const cfg = { not_started: { label: "Ikke startet", color: "#868E96" }, active: { label: "Aktivt", color: "#1971C2" }, completed: { label: "Ferdig", color: "#2F9E44" } }[s];
                const isSelected = editProjectStatus === s;
                return (
                  <button
                    key={s}
                    onClick={() => setEditProjectStatus(s)}
                    style={{
                      flex: 1, padding: "8px 4px", borderRadius: 6, border: `2px solid`,
                      borderColor: isSelected ? cfg.color : "var(--bfc-base-dimmed)",
                      background: isSelected ? `${cfg.color}15` : "var(--bfc-base-3)",
                      color: isSelected ? cfg.color : "var(--bfc-base-c-2)",
                      fontWeight: isSelected ? 600 : 400, fontSize: "0.82rem",
                      cursor: "pointer", transition: "all 0.15s ease",
                    }}
                  >
                    {cfg.label}
                  </button>
                );
              })}
            </div>
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setEditProjectModal(false)}>Avbryt</Button>
            <Button variant="filled" onClick={handleEditProject} state={editProjectSaving ? "inactive" : "default"}>
              {editProjectSaving ? "Lagrer..." : "Lagre"}
            </Button>
          </div>
        </div>
      </Modal>
    </div>
  );
}

function Section({ type, items, onNew, onOpen, onRename, onDelete, onSetPrimary, countLabel }: {
  type: PlanType; items: PlanItem[]; onNew: () => void;
  onOpen: (id: string) => void; onRename: (item: PlanItem) => void;
  onDelete: (item: PlanItem) => void; onSetPrimary?: (item: PlanItem) => void;
  countLabel: (item: object) => string;
}) {
  const { color, label } = SECTION_CONFIG[type];
  return (
    <div style={{ marginBottom: "2rem" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "0.75rem" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "0.65rem" }}>
          <div style={{ width: 4, height: 22, borderRadius: 2, background: color, flexShrink: 0 }} />
          <h2 className="bf-h3" style={{ margin: 0 }}>{label}</h2>
          {items.length > 0 && <span style={{ fontSize: "0.75rem", fontWeight: 600, padding: "1px 8px", borderRadius: 20, background: `${color}18`, color }}>{items.length}</span>}
        </div>
        <Button variant="outline" onClick={onNew}>+ Ny</Button>
      </div>
      {items.length === 0 ? (
        <div style={{ padding: "1rem 1.25rem", borderRadius: 6, border: `1px dashed ${color}60`, color: "var(--bfc-base-c-2)", fontSize: "0.9rem", background: `${color}06` }}>
          Ingen {label.toLowerCase()} ennå. Klikk «+ Ny» for å opprette.
        </div>
      ) : (
        <div style={{ display: "grid", gap: "0.5rem" }}>
          {items.map((item) => (
            <PlanRow key={item.id} item={item} color={color} countLabel={countLabel}
              showPrimary={!!onSetPrimary && items.length > 1}
              onOpen={() => onOpen(item.id)} onRename={() => onRename(item)}
              onDelete={() => onDelete(item)}
              onSetPrimary={onSetPrimary ? () => onSetPrimary(item) : undefined}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function PlanRow({ item, color, countLabel, showPrimary, onOpen, onRename, onDelete, onSetPrimary }: {
  item: PlanItem; color: string; countLabel: (item: object) => string;
  showPrimary?: boolean;
  onOpen: () => void; onRename: () => void; onDelete: () => void;
  onSetPrimary?: () => void;
}) {
  const [hovered, setHovered] = useState(false);
  return (
    <div
      onClick={onOpen}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{ display: "flex", alignItems: "center", padding: "0.85rem 1.1rem", borderRadius: 7, background: "var(--bfc-base-3)", border: `1px solid ${item.is_primary ? color : "var(--bfc-base-dimmed)"}`, borderLeft: `3px solid ${item.is_primary || hovered ? color : "var(--bfc-base-dimmed)"}`, cursor: "pointer", gap: "0.75rem", boxShadow: hovered ? "0 2px 10px rgba(0,0,0,0.08)" : "none", transform: hovered ? "translateX(2px)" : "translateX(0)", transition: "border-color 0.15s, box-shadow 0.15s, transform 0.15s" }}
    >
      <span style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: hovered ? 600 : 400, transition: "font-weight 0.1s" }}>
        {item.title}
      </span>
      <div style={{ display: "flex", alignItems: "center", gap: "0.6rem", flexShrink: 0 }} onClick={(e) => e.stopPropagation()}>
        {showPrimary && item.is_primary && (
          <span style={{ fontSize: "0.72rem", fontWeight: 700, padding: "2px 8px", borderRadius: 20, background: `${color}20`, color, border: `1px solid ${color}40` }}>
            Dashboard
          </span>
        )}
        {showPrimary && !item.is_primary && onSetPrimary && (
          <button onClick={onSetPrimary} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "4px 8px", borderRadius: 4, fontSize: "0.8rem", fontWeight: 500, opacity: hovered ? 0.7 : 0, transition: "opacity 0.15s" }}>
            Sett som primær
          </button>
        )}
        <span style={{ fontSize: "0.75rem", fontWeight: 600, padding: "2px 10px", borderRadius: 20, background: `${color}18`, color }}>{countLabel(item)}</span>
        <button onClick={onRename} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "4px 8px", borderRadius: 4, fontSize: "0.8rem", fontWeight: 500, opacity: hovered ? 1 : 0.4, transition: "opacity 0.15s, color 0.15s" }} onMouseEnter={(e) => (e.currentTarget.style.color = color)} onMouseLeave={(e) => (e.currentTarget.style.color = "var(--bfc-base-c-2)")}>Endre</button>
        <div style={{ width: 1, height: 14, background: "var(--bfc-base-dimmed)" }} />
        <button onClick={onDelete} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "4px 8px", borderRadius: 4, fontSize: "0.8rem", fontWeight: 500, opacity: hovered ? 1 : 0.4, transition: "opacity 0.15s, color 0.15s" }} onMouseEnter={(e) => (e.currentTarget.style.color = "#E03131")} onMouseLeave={(e) => (e.currentTarget.style.color = "var(--bfc-base-c-2)")}>Slett</button>
      </div>
    </div>
  );
}

// ─── Dashboard ────────────────────────────────────────────────────────────────

function dashRiskLevel(score: number): "low" | "medium" | "high" {
  if (score <= 6) return "low";
  if (score <= 14) return "medium";
  return "high";
}

const DASH_RISK = {
  low:    { color: "#2F9E44", bg: "#D3F9D8", border: "#2F9E4440", label: "Lav" },
  medium: { color: "#F76707", bg: "#FFE8CC", border: "#F7670740", label: "Middels" },
  high:   { color: "#E03131", bg: "#FFE3E3", border: "#E0313140", label: "Høy" },
};

type DeadlineItem = {
  key: string;
  type: "meeting" | "task" | "activity" | "oppgave";
  title: string;
  date: Date;
  done: boolean;
  context: string;
  source?: "planner";
};

const TYPE_CONFIG: Record<DeadlineItem["type"], { color: string; label: string }> = {
  meeting:  { color: "#7950F2", label: "Møte" },
  task:     { color: "#0CA678", label: "Oppgave" },
  activity: { color: "#7950F2", label: "Aktivitet" },
  oppgave:  { color: "#F59F00", label: "Oppgave" },
};

function relativeDateLabel(date: Date): string {
  const now = new Date();
  const diffMs = date.getTime() - new Date(now.toDateString()).getTime();
  const diffDays = Math.floor(diffMs / 86400000);
  if (diffDays === 0) return "I dag";
  if (diffDays === 1) return "I morgen";
  return date.toLocaleDateString("nb-NO", { day: "numeric", month: "short" });
}

// ── KPI card ─────────────────────────────────────────────────────────────────

function KpiCard({ value, label, sub, color, onClick }: {
  value: string | number;
  label: string;
  sub?: string;
  color: string;
  onClick?: () => void;
}) {
  const [hov, setHov] = useState(false);
  return (
    <div
      onClick={onClick}
      onMouseEnter={() => setHov(true)}
      onMouseLeave={() => setHov(false)}
      style={{
        padding: "1.75rem 2rem", borderRadius: 12,
        background: hov && onClick
          ? `linear-gradient(135deg, ${color}10 0%, var(--bfc-base-3) 60%)`
          : `linear-gradient(135deg, ${color}08 0%, var(--bfc-base-3) 50%)`,
        border: `1px solid ${hov && onClick ? color + "40" : "var(--bfc-base-dimmed)"}`,
        borderLeft: `5px solid ${color}`,
        cursor: onClick ? "pointer" : "default",
        boxShadow: hov && onClick ? `0 6px 24px ${color}20` : "0 1px 4px rgba(0,0,0,0.05)",
        transform: hov && onClick ? "translateY(-3px)" : "none",
        transition: "box-shadow 0.18s, transform 0.18s, border-color 0.18s, background 0.18s",
        minWidth: 0,
        position: "relative",
        overflow: "hidden",
      }}
    >
      <div style={{
        position: "absolute", top: 0, right: 0,
        width: 80, height: 80, borderRadius: "50%",
        background: `${color}08`, transform: "translate(20px, -20px)",
        pointerEvents: "none",
      }} />
      <div style={{ fontSize: "2.6rem", fontWeight: 800, color, lineHeight: 1, marginBottom: "0.4rem", letterSpacing: "-0.01em" }}>{value}</div>
      <div style={{ fontSize: "0.92rem", fontWeight: 700, color: "var(--bfc-base-c-1)", marginBottom: "0.25rem" }}>{label}</div>
      {sub && <div style={{ fontSize: "0.78rem", color: "var(--bfc-base-c-2)", lineHeight: 1.4 }}>{sub}</div>}
    </div>
  );
}

// ── Mini heatmap ──────────────────────────────────────────────────────────────

function MiniHeatmap({ risks }: { risks: RiskItem[] }) {
  const CELL = 40;
  const GAP = 2;

  const counts: Record<string, number> = {};
  for (const r of risks) {
    const k = `${r.probability},${r.consequence}`;
    counts[k] = (counts[k] ?? 0) + 1;
  }

  return (
    <div>
      <div style={{ display: "flex", alignItems: "flex-end", gap: GAP }}>
        {/* Y-axis label */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", width: 18 }}>
          <span style={{ writingMode: "vertical-rl", transform: "rotate(180deg)", fontSize: "0.65rem", fontWeight: 600, color: "var(--bfc-base-c-3)", textTransform: "uppercase", letterSpacing: "0.05em" }}>
            Sannsynlighet
          </span>
        </div>
        <div>
          {[5, 4, 3, 2, 1].map((p) => (
            <div key={p} style={{ display: "flex", alignItems: "center", gap: GAP, marginBottom: GAP }}>
              <div style={{ width: 14, textAlign: "right", fontSize: "0.65rem", fontWeight: 700, color: "var(--bfc-base-c-2)", flexShrink: 0 }}>{p}</div>
              {[1, 2, 3, 4, 5].map((c) => {
                const score = p * c;
                const lvl = dashRiskLevel(score);
                const cfg = DASH_RISK[lvl];
                const count = counts[`${p},${c}`] ?? 0;
                return (
                  <div key={c} style={{
                    width: CELL, height: CELL - 4,
                    background: cfg.bg, border: `1px solid ${cfg.border}`,
                    borderRadius: 4, display: "flex", alignItems: "center",
                    justifyContent: "center", position: "relative",
                  }}>
                    {count > 0 && (
                      <span style={{
                        background: cfg.color, color: "#fff",
                        borderRadius: "50%", width: 18, height: 18,
                        display: "flex", alignItems: "center", justifyContent: "center",
                        fontSize: "0.65rem", fontWeight: 800,
                      }}>{count}</span>
                    )}
                  </div>
                );
              })}
            </div>
          ))}
          {/* X labels */}
          <div style={{ display: "flex", gap: GAP, paddingLeft: 16 }}>
            {[1, 2, 3, 4, 5].map((c) => (
              <div key={c} style={{ width: CELL, textAlign: "center", fontSize: "0.65rem", fontWeight: 700, color: "var(--bfc-base-c-2)" }}>{c}</div>
            ))}
          </div>
          <div style={{ textAlign: "center", fontSize: "0.65rem", color: "var(--bfc-base-c-3)", paddingLeft: 16, marginTop: 2, textTransform: "uppercase", letterSpacing: "0.05em" }}>
            Konsekvens
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Resource progress bar ─────────────────────────────────────────────────────

function ResourceProgressRow({ name, done, total, color, typeLabel, external, externalUrl }: {
  name: string; done: number | null; total: number | null; color: string; typeLabel: string;
  external?: boolean; externalUrl?: string | null;
}) {
  const pct = (total != null && total > 0 && done != null) ? Math.round((done / total) * 100) : 0;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}>
      <div style={{ display: "flex", alignItems: "center", gap: "0.4rem", width: 220, flexShrink: 0, minWidth: 0 }}>
        <span style={{ fontSize: "0.7rem", padding: "1px 7px", borderRadius: 20, background: `${color}18`, color, fontWeight: 600, flexShrink: 0 }}>
          {typeLabel}
        </span>
        <span style={{ fontSize: "0.85rem", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          {name}
        </span>
      </div>
      {external ? (
        <div style={{ flex: 1, display: "flex", alignItems: "center", gap: "0.5rem" }}>
          <span style={{ fontSize: "0.75rem", padding: "1px 8px", borderRadius: 20, background: "#868E9618", color: "#868E96", fontWeight: 600 }}>
            Ekstern
          </span>
          {externalUrl && (
            <a href={externalUrl} target="_blank" rel="noreferrer"
              style={{ fontSize: "0.75rem", color: color, fontWeight: 500, textDecoration: "none" }}
              onMouseEnter={(e) => (e.currentTarget.style.textDecoration = "underline")}
              onMouseLeave={(e) => (e.currentTarget.style.textDecoration = "none")}
            >
              Åpne →
            </a>
          )}
        </div>
      ) : (
        <>
          <div style={{ flex: 1, height: 7, borderRadius: 4, background: "var(--bfc-base-dimmed)", overflow: "hidden", minWidth: 60 }}>
            <div style={{ height: "100%", width: `${pct}%`, background: pct === 100 ? "#2F9E44" : color, borderRadius: 4, transition: "width 0.4s ease" }} />
          </div>
          <span style={{ fontSize: "0.78rem", color: "var(--bfc-base-c-2)", whiteSpace: "nowrap", flexShrink: 0, minWidth: 70, textAlign: "right" }}>
            {done}/{total} · {pct}%
          </span>
        </>
      )}
    </div>
  );
}

// ── Dashboard view ────────────────────────────────────────────────────────────

function DashboardView({ project, riskMatrices, projectPlans, oppgaveLister, runbooks, meetingPlans, projectId, navigate }: {
  project: Project;
  riskMatrices: RiskMatrix[];
  projectPlans: ProjectPlan[];
  oppgaveLister: OppgaveListe[];
  runbooks: Runbook[];
  meetingPlans: MeetingPlan[];
  projectId: string;
  navigate: (path: string) => void;
}) {
  // ── Planner deadlines + stats (silent fetch) ────────────────────────────────
  const [plannerDeadlines, setPlannerDeadlines] = useState<DeadlineItem[]>([]);
  const [plannerLoading, setPlannerLoading] = useState(false);
  const [komendeExpanded, setKomendeExpanded] = useState(false);
  // Maps our DB id → { done, total } for Planner-backed plans
  const [plannerStats, setPlannerStats] = useState<Map<string, { done: number; total: number }>>(new Map());

  const fetchPlannerDeadlines = useCallback(async () => {
    if (!isMsalConfigured) return;
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return;
    const account = accounts[0];

    const resources: Array<{ id: string; url: string; context: string; type: DeadlineItem["type"] }> = [
      ...projectPlans
        .filter((p) => p.source === "planner" && p.external_url)
        .map((p) => ({ id: p.id, url: p.external_url!, context: p.title, type: "task" as const })),
      ...oppgaveLister
        .filter((ol) => ol.source === "planner" && ol.external_url)
        .map((ol) => ({ id: ol.id, url: ol.external_url!, context: ol.title, type: "oppgave" as const })),
      ...runbooks
        .filter((rb) => rb.source === "planner" && rb.external_url)
        .map((rb) => ({ id: rb.id, url: rb.external_url!, context: rb.title, type: "activity" as const })),
    ];

    if (resources.length === 0) return;

    setPlannerLoading(true);
    const fetchNow = new Date();
    const fetchTwoWeeks = new Date(fetchNow.getTime() + 14 * 86400000);

    const results = await Promise.allSettled(
      resources.map(async ({ id, url, context, type }) => {
        const data = await fetchPlannerData(msalInstance, account, url);
        const stats = {
          done: data.tasks.filter((t) => t.percentComplete === 100).length,
          total: data.tasks.length,
        };
        const deadlines = data.tasks
          .filter((t) => t.dueDateTime)
          .map((t) => ({
            key: `planner-${t.id}`,
            type,
            title: t.title,
            date: new Date(t.dueDateTime!),
            done: t.percentComplete === 100,
            context,
            source: "planner" as const,
          }))
          .filter((d) => d.date >= fetchNow && d.date <= fetchTwoWeeks);
        return { id, deadlines, stats };
      })
    );

    const items: DeadlineItem[] = [];
    const newStats = new Map<string, { done: number; total: number }>();
    for (const r of results) {
      if (r.status === "fulfilled") {
        items.push(...r.value.deadlines);
        newStats.set(r.value.id, r.value.stats);
      }
    }

    setPlannerDeadlines(items.sort((a, b) => a.date.getTime() - b.date.getTime()));
    setPlannerStats(newStats);
    setPlannerLoading(false);
  }, [projectPlans, oppgaveLister, runbooks]);

  useEffect(() => {
    fetchPlannerDeadlines();
  }, [fetchPlannerDeadlines]);

  // ── Primary resource selection (fallback to first if none marked) ───────────
  const primaryMatrix      = riskMatrices.find((m) => m.is_primary) ?? riskMatrices[0];
  const primaryOppgaveListe = oppgaveLister.find((ol) => ol.is_primary) ?? oppgaveLister[0];
  const primaryMeetingPlan  = meetingPlans.find((mp) => mp.is_primary) ?? meetingPlans[0];
  const primaryProjectPlan  = projectPlans.find((p) => p.is_primary) ?? projectPlans[0];

  // ── Derive data ─────────────────────────────────────────────────────────────
  const allRisks = (primaryMatrix?.risks ?? []);
  const openRisks = allRisks.filter((r) => r.status === "open");
  const highestScore = openRisks.reduce((max, r) => Math.max(max, r.risk_score), 0);
  const topRisks = [...openRisks].sort((a, b) => b.risk_score - a.risk_score).slice(0, 4);

  // "Ferdig"-kortet: ferdige oppgaver fra primær oppgaveliste
  const primaryOlStats = primaryOppgaveListe ? plannerStats.get(primaryOppgaveListe.id) : undefined;
  const doneOppgaver = primaryOppgaveListe?.source === "own"
    ? primaryOppgaveListe.oppgaver.filter((o) => o.status === "done").length
    : (primaryOlStats?.done ?? 0);
  const totalOppgaver = primaryOppgaveListe?.source === "own"
    ? primaryOppgaveListe.oppgaver.length
    : (primaryOlStats?.total ?? 0);

  // "Leveranser"-kortet: % ferdig fra primær prosjektplan
  const primaryPlanStats = primaryProjectPlan ? plannerStats.get(primaryProjectPlan.id) : undefined;
  const planTasksDone = primaryProjectPlan?.source === "own"
    ? primaryProjectPlan.tasks.filter((t) => t.percent_complete === 100).length
    : (primaryPlanStats?.done ?? 0);
  const planTasksTotal = primaryProjectPlan?.source === "own"
    ? primaryProjectPlan.tasks.length
    : (primaryPlanStats?.total ?? 0);
  const leveranserPct = planTasksTotal > 0 ? Math.round((planTasksDone / planTasksTotal) * 100) : null;

  const now = new Date();
  const nextMeeting = (primaryMeetingPlan?.meetings ?? [])
    .filter((m) => new Date(m.date) > now)
    .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())[0] ?? null;

  // ── Upcoming deadlines (14 days) ────────────────────────────────────────────
  const twoWeeks = new Date(now.getTime() + 14 * 86400000);
  const deadlines: DeadlineItem[] = [
    ...meetingPlans.flatMap((mp) =>
      mp.meetings.map((m) => ({
        key: `meeting-${m.id}`, type: "meeting" as const,
        title: m.title, date: new Date(m.date), done: false, context: mp.title,
      }))
    ),
    ...projectPlans.flatMap((p) =>
      p.tasks.filter((t) => t.end_date).map((t) => ({
        key: `task-${t.id}`, type: "task" as const,
        title: t.name, date: new Date(t.end_date!),
        done: t.percent_complete === 100, context: p.title,
      }))
    ),
    ...oppgaveLister.flatMap((ol) =>
      ol.oppgaver.filter((o) => o.due_date).map((o) => ({
        key: `oppgave-${o.id}`, type: "oppgave" as const,
        title: o.name, date: new Date(o.due_date!),
        done: o.status === "done", context: ol.title,
      }))
    ),
    ...runbooks.flatMap((rb) =>
      rb.activities.filter((a) => a.end_date).map((a) => ({
        key: `activity-${a.id}`, type: "activity" as const,
        title: a.name, date: new Date(a.end_date!),
        done: a.status === "done", context: rb.title,
      }))
    ),
  ]
    .filter((d) => d.date >= now && d.date <= twoWeeks)
    .concat(plannerDeadlines)
    .sort((a, b) => a.date.getTime() - b.date.getTime());

  // ── Resource progress rows — all sources ────────────────────────────────────
  type ProgressRow = {
    id: string; name: string;
    done: number | null; total: number | null;
    color: string; typeLabel: string;
    external: boolean; externalUrl?: string | null;
  };
  const progressRows: ProgressRow[] = [
    ...projectPlans.map((p) => ({
      id: p.id, name: p.title,
      done: p.source === "own" ? p.tasks.filter((t) => t.percent_complete === 100).length : null,
      total: p.source === "own" ? p.tasks.length : null,
      color: SECTION_CONFIG.projectplan.color, typeLabel: "Prosjektplan",
      external: p.source !== "own", externalUrl: p.external_url,
    })),
    ...oppgaveLister.map((ol) => ({
      id: ol.id, name: ol.title,
      done: ol.source === "own" ? ol.oppgaver.filter((o) => o.status === "done").length : null,
      total: ol.source === "own" ? ol.oppgaver.length : null,
      color: SECTION_CONFIG.oppgave.color, typeLabel: "Oppgaver",
      external: ol.source !== "own", externalUrl: ol.external_url,
    })),
    ...runbooks.map((rb) => ({
      id: rb.id, name: rb.title,
      done: rb.source === "own" ? rb.activities.filter((a) => a.status === "done").length : null,
      total: rb.source === "own" ? rb.activities.length : null,
      color: SECTION_CONFIG.runbook.color, typeLabel: "Runbook",
      external: rb.source !== "own", externalUrl: rb.external_url,
    })),
  ];

  // ── KPI config ──────────────────────────────────────────────────────────────
  const riskKpiColor = openRisks.length === 0
    ? "#868E96"
    : DASH_RISK[dashRiskLevel(highestScore)].color;

  const riskKpiSub = openRisks.length === 0
    ? "Ingen åpne risikoer"
    : `${DASH_RISK[dashRiskLevel(highestScore)].label} risiko · klikk for å se`;

  const hasPlannerOppgaver = isMsalConfigured && primaryOppgaveListe?.source === "planner";
  const hasPlannerPlans   = isMsalConfigured && primaryProjectPlan?.source === "planner";
  const msalAuthenticated = isMsalConfigured && msalInstance.getAllAccounts().length > 0;

  const oppgaverSub =
    plannerLoading && hasPlannerOppgaver ? "Laster Planner..." :
    hasPlannerOppgaver && !msalAuthenticated ? "Logg inn for Planner-data" :
    totalOppgaver > 0 ? `${doneOppgaver} av ${totalOppgaver} oppgaver` :
    "Ingen oppgaver registrert";

  const leveranserSub =
    plannerLoading && hasPlannerPlans ? "Laster Planner..." :
    hasPlannerPlans && !msalAuthenticated ? "Logg inn for Planner-data" :
    planTasksTotal > 0 ? `${planTasksDone} av ${planTasksTotal} oppgaver ferdig` :
    "Ingen prosjektplaner";

  function handleDashboardExport() {
    void exportDashboardPdf({
      project,
      openRisksCount: openRisks.length,
      highestRiskScore: highestScore,
      topRisks: topRisks.map((r) => ({ description: r.description, risk_score: r.risk_score, fagomrade: r.fagomrade ?? null, risk_owner: r.risk_owner ?? null })),
      doneOppgaver,
      totalOppgaver,
      nextMeetingTitle: nextMeeting?.title ?? null,
      nextMeetingDate: nextMeeting?.date ?? null,
      leveranserPct,
      planTasksDone,
      planTasksTotal,
      primaryPlanTitle: primaryProjectPlan?.title ?? null,
      primaryOppgaveTitle: primaryOppgaveListe?.title ?? null,
      primaryMeetingTitle: primaryMeetingPlan?.title ?? null,
      primaryMatrixTitle: primaryMatrix?.title ?? null,
    });
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: "2rem" }}>

      {/* ── Dashboard toolbar ───────────────────────────────────────────────── */}
      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <Button variant="outline" onClick={handleDashboardExport}>↓ Eksporter PDF</Button>
      </div>

      {/* ── KPI row ─────────────────────────────────────────────────────────── */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "1.25rem" }}>
        <KpiCard
          value={openRisks.length}
          label="Åpne risikoer"
          sub={riskKpiSub}
          color={riskKpiColor}
          onClick={primaryMatrix ? () => navigate(`/projects/${projectId}/risk-matrix/${primaryMatrix.id}`) : undefined}
        />
        <KpiCard
          value={plannerLoading && hasPlannerOppgaver ? "…" : doneOppgaver}
          label="Ferdige oppgaver"
          sub={oppgaverSub}
          color="#1971C2"
          onClick={primaryOppgaveListe ? () => navigate(`/projects/${projectId}/oppgave/${primaryOppgaveListe.id}`) : undefined}
        />
        <KpiCard
          value={nextMeeting
            ? new Date(nextMeeting.date).toLocaleDateString("nb-NO", { day: "numeric", month: "short" })
            : "–"}
          label="Neste møte"
          sub={nextMeeting?.title ?? "Ingen planlagte møter"}
          color="#7950F2"
        />
        <KpiCard
          value={plannerLoading && hasPlannerPlans ? "…" : (leveranserPct !== null ? `${leveranserPct}%` : "–")}
          label="Leveranser"
          sub={leveranserSub}
          color="#1098AD"
          onClick={primaryProjectPlan ? () => navigate(`/projects/${projectId}/project-plan/${primaryProjectPlan.id}`) : undefined}
        />
      </div>

      {/* ── Middle: risk alerts + deadlines ────────────────────────────────── */}
      <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)", gap: "1.25rem", alignItems: "start" }}>

        {/* Left: Top risks + mini heatmap */}
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <div style={{ borderRadius: 10, background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)", padding: "1.25rem" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
              <h3 className="bf-h4" style={{ margin: 0 }}>Topp risikoer</h3>
              {primaryMatrix && (
                <button onClick={() => navigate(`/projects/${projectId}/risk-matrix/${primaryMatrix.id}`)}
                  style={{ background: "none", border: "none", cursor: "pointer", color: "#E03131", fontSize: "0.8rem", fontWeight: 600, padding: 0 }}>
                  Se alle →
                </button>
              )}
            </div>
            {topRisks.length === 0 ? (
              <p style={{ color: "var(--bfc-base-c-2)", fontSize: "0.9rem", margin: 0 }}>Ingen åpne risikoer registrert.</p>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: "0.5rem" }}>
                {topRisks.map((r) => {
                  const lvl = dashRiskLevel(r.risk_score);
                  const cfg = DASH_RISK[lvl];
                  return (
                    <div key={r.id} style={{
                      display: "flex", alignItems: "center", gap: "0.65rem",
                      padding: "0.5rem 0.75rem", borderRadius: 7,
                      background: cfg.bg, border: `1px solid ${cfg.border}`,
                    }}>
                      <span style={{
                        background: cfg.color, color: "#fff",
                        borderRadius: 4, padding: "2px 8px",
                        fontSize: "0.78rem", fontWeight: 800, flexShrink: 0,
                      }}>
                        {r.risk_score}
                      </span>
                      <span style={{ fontSize: "0.85rem", flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                        {r.description}
                      </span>
                      {r.fagomrade && (
                        <span style={{ fontSize: "0.72rem", color: cfg.color, fontWeight: 600, flexShrink: 0 }}>
                          {r.fagomrade}
                        </span>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* Mini heatmap */}
          {allRisks.length > 0 && (
            <div style={{ borderRadius: 10, background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)", padding: "1.25rem" }}>
              <h3 className="bf-h4" style={{ margin: "0 0 1rem" }}>Risikokart</h3>
              <MiniHeatmap risks={openRisks} />
            </div>
          )}
        </div>

        {/* Right: Kommende frister */}
        <div style={{ borderRadius: 10, background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)", padding: "1.25rem" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
            <h3 className="bf-h4" style={{ margin: 0 }}>Kommende (14 dager)</h3>
            <div style={{ display: "flex", alignItems: "center", gap: "0.5rem" }}>
              {plannerLoading && (
                <span style={{ fontSize: "0.72rem", color: "#0078D4", fontWeight: 500 }}>
                  ↻ Laster Planner…
                </span>
              )}
              <button
                onClick={() => setKomendeExpanded(true)}
                title="Vis alle"
                style={{
                  background: "none", border: "1px solid var(--bfc-base-dimmed)",
                  borderRadius: 6, cursor: "pointer", padding: "3px 8px",
                  color: "var(--bfc-base-c-2)", fontSize: "0.78rem", fontWeight: 500,
                  display: "flex", alignItems: "center", gap: "0.3rem",
                  transition: "border-color 0.15s, color 0.15s",
                }}
                onMouseEnter={(e) => { e.currentTarget.style.borderColor = "#4C6EF5"; e.currentTarget.style.color = "#4C6EF5"; }}
                onMouseLeave={(e) => { e.currentTarget.style.borderColor = "var(--bfc-base-dimmed)"; e.currentTarget.style.color = "var(--bfc-base-c-2)"; }}
              >
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                  <polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/>
                  <line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/>
                </svg>
                Utvid
              </button>
            </div>
          </div>
          {deadlines.length === 0 && !plannerLoading ? (
            <p style={{ color: "var(--bfc-base-c-2)", fontSize: "0.9rem", margin: 0 }}>
              Ingen frister de neste 14 dagene.
            </p>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: "0.4rem" }}>
              {deadlines.map((d) => {
                const cfg = TYPE_CONFIG[d.type];
                return (
                  <div key={d.key} style={{
                    display: "flex", alignItems: "center", gap: "0.5rem",
                    padding: "0.5rem 0.75rem", borderRadius: 7,
                    background: d.done ? "var(--bfc-base-2)" : "var(--bfc-base-3)",
                    border: "1px solid var(--bfc-base-dimmed)",
                    opacity: d.done ? 0.55 : 1,
                    minWidth: 0, overflow: "hidden",
                  }}>
                    <span style={{
                      fontSize: "0.68rem", fontWeight: 600, padding: "1px 7px",
                      borderRadius: 20, background: `${cfg.color}18`, color: cfg.color,
                      flexShrink: 0, whiteSpace: "nowrap",
                    }}>
                      {cfg.label}
                    </span>
                    {d.source === "planner" && (
                      <span style={{
                        fontSize: "0.65rem", fontWeight: 600, padding: "1px 6px",
                        borderRadius: 20, background: "#0078D418", color: "#0078D4",
                        flexShrink: 0, whiteSpace: "nowrap",
                      }}>
                        Planner
                      </span>
                    )}
                    <span style={{
                      flex: 1, minWidth: 0, fontSize: "0.85rem",
                      overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                      textDecoration: d.done ? "line-through" : "none",
                      color: d.done ? "var(--bfc-base-c-3)" : "inherit",
                    }}>
                      {d.title}
                    </span>
                    <div style={{ flexShrink: 0, textAlign: "right", minWidth: 56 }}>
                      <div style={{ fontSize: "0.78rem", fontWeight: 600, whiteSpace: "nowrap", color: d.date.getTime() - now.getTime() < 86400000 ? "#E03131" : "var(--bfc-base-c-1)" }}>
                        {relativeDateLabel(d.date)}
                      </div>
                      <div style={{ fontSize: "0.68rem", color: "var(--bfc-base-c-3)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 100 }}>{d.context}</div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      {/* ── Resource progress ────────────────────────────────────────────────── */}
      <div style={{ borderRadius: 10, background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)", padding: "1.25rem" }}>
        <h3 className="bf-h4" style={{ margin: "0 0 1rem" }}>Ressursfremdrift</h3>
        {progressRows.length === 0 ? (
          <p style={{ color: "var(--bfc-base-c-2)", fontSize: "0.9rem", margin: 0 }}>
            Ingen prosjektplaner, oppgavelister eller runbooks registrert.
          </p>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: "0.75rem" }}>
            {progressRows.map((row) => (
              <ResourceProgressRow key={row.id} name={row.name} done={row.done} total={row.total} color={row.color} typeLabel={row.typeLabel} external={row.external} externalUrl={row.externalUrl} />
            ))}
          </div>
        )}
      </div>

      {/* ── Kommende expanded overlay ─────────────────────────────────────────── */}
      {komendeExpanded && (
        <div
          onClick={() => setKomendeExpanded(false)}
          style={{
            position: "fixed", inset: 0, zIndex: 1000,
            background: "rgba(0,0,0,0.45)",
            display: "flex", alignItems: "center", justifyContent: "center",
            backdropFilter: "blur(2px)",
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: "var(--bfc-base-3)", borderRadius: 14,
              border: "1px solid var(--bfc-base-dimmed)",
              width: "min(700px, 95vw)", maxHeight: "80vh",
              display: "flex", flexDirection: "column",
              boxShadow: "0 24px 64px rgba(0,0,0,0.22)",
            }}
          >
            {/* Header */}
            <div style={{
              display: "flex", justifyContent: "space-between", alignItems: "center",
              padding: "1.25rem 1.5rem", borderBottom: "1px solid var(--bfc-base-dimmed)", flexShrink: 0,
            }}>
              <div style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}>
                <h3 className="bf-h3" style={{ margin: 0 }}>Kommende (14 dager)</h3>
                <span style={{
                  fontSize: "0.78rem", fontWeight: 600, padding: "2px 10px",
                  borderRadius: 20, background: "#4C6EF518", color: "#4C6EF5",
                }}>
                  {deadlines.length} {deadlines.length === 1 ? "aktivitet" : "aktiviteter"}
                </span>
                {plannerLoading && (
                  <span style={{ fontSize: "0.72rem", color: "#0078D4", fontWeight: 500 }}>↻ Laster Planner…</span>
                )}
              </div>
              <button
                onClick={() => setKomendeExpanded(false)}
                style={{
                  background: "none", border: "1px solid var(--bfc-base-dimmed)",
                  borderRadius: 6, cursor: "pointer", padding: "4px 12px",
                  color: "var(--bfc-base-c-2)", fontSize: "0.85rem", fontWeight: 500,
                }}
                onMouseEnter={(e) => { e.currentTarget.style.borderColor = "#E03131"; e.currentTarget.style.color = "#E03131"; }}
                onMouseLeave={(e) => { e.currentTarget.style.borderColor = "var(--bfc-base-dimmed)"; e.currentTarget.style.color = "var(--bfc-base-c-2)"; }}
              >
                ✕ Lukk
              </button>
            </div>

            {/* Body */}
            <div style={{ overflowY: "auto", padding: "1.25rem 1.5rem", display: "flex", flexDirection: "column", gap: "0.4rem" }}>
              {deadlines.length === 0 && !plannerLoading ? (
                <p style={{ color: "var(--bfc-base-c-2)", fontSize: "0.9rem", margin: 0 }}>
                  Ingen frister de neste 14 dagene.
                </p>
              ) : (
                deadlines.map((d) => {
                  const cfg = TYPE_CONFIG[d.type];
                  return (
                    <div key={d.key} style={{
                      display: "flex", alignItems: "flex-start", gap: "0.5rem",
                      padding: "0.65rem 1rem", borderRadius: 8,
                      background: d.done ? "var(--bfc-base-2)" : "var(--bfc-base-3)",
                      border: "1px solid var(--bfc-base-dimmed)",
                      opacity: d.done ? 0.55 : 1,
                      minWidth: 0,
                    }}>
                      <span style={{
                        fontSize: "0.68rem", fontWeight: 600, padding: "1px 7px",
                        borderRadius: 20, background: `${cfg.color}18`, color: cfg.color,
                        flexShrink: 0, whiteSpace: "nowrap", marginTop: 2,
                      }}>
                        {cfg.label}
                      </span>
                      {d.source === "planner" && (
                        <span style={{
                          fontSize: "0.65rem", fontWeight: 600, padding: "1px 6px",
                          borderRadius: 20, background: "#0078D418", color: "#0078D4",
                          flexShrink: 0, whiteSpace: "nowrap", marginTop: 2,
                        }}>
                          Planner
                        </span>
                      )}
                      <span style={{
                        flex: 1, minWidth: 0, fontSize: "0.9rem",
                        lineHeight: 1.45,
                        textDecoration: d.done ? "line-through" : "none",
                        color: d.done ? "var(--bfc-base-c-3)" : "inherit",
                        wordBreak: "break-word",
                      }}>
                        {d.title}
                      </span>
                      <div style={{ flexShrink: 0, textAlign: "right", minWidth: 80, marginTop: 1 }}>
                        <div style={{
                          fontSize: "0.82rem", fontWeight: 600, whiteSpace: "nowrap",
                          color: d.date.getTime() - now.getTime() < 86400000 ? "#E03131" : "var(--bfc-base-c-1)",
                        }}>
                          {relativeDateLabel(d.date)}
                        </div>
                        <div style={{
                          fontSize: "0.72rem", color: "var(--bfc-base-c-3)",
                          whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 120,
                        }}>
                          {d.context}
                        </div>
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
