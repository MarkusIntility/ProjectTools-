import { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type Project, type RiskMatrix, type CommunicationPlan, type MeetingPlan, type Runbook, type ProjectPlan, type OppgaveListe, type Template } from "../api/client";

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
type PlanItem = { id: string; title: string };
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

  const [renameTarget, setRenameTarget] = useState<{ id: string; title: string; type: PlanType } | null>(null);
  const [renameValue, setRenameValue] = useState("");
  const [renameSaving, setRenameSaving] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<{ id: string; title: string; type: PlanType } | null>(null);
  const [deleting, setDeleting] = useState(false);

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
            await api.riskMatrices.addRisk(projectId, m.id, { description: r.description, probability: r.probability, consequence: r.consequence, mitigation: r.mitigation ?? null, owner: r.owner ?? null, status: (r.status ?? "open") as "open" | "mitigated" | "closed" });
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
          const d = JSON.parse(simpleSelectedTemplate.data) as { meetings?: Array<{ title: string; date: string; location?: string | null; agenda?: string | null; participants?: string | null; minutes?: string | null }> };
          for (const mt of d.meetings ?? []) await api.meetingPlans.addMeeting(projectId, p.id, { title: mt.title, date: mt.date, location: mt.location ?? null, agenda: mt.agenda ?? null, participants: mt.participants ?? null, minutes: mt.minutes ?? null });
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

  if (!project) return <div style={{ padding: "2rem" }}>Laster...</div>;

  const color = accentColor(project.name);

  const sourcePickerCfg = {
    planner:    { color: "#0078D4", label: "Microsoft Planner", sub: "Lenk til eksisterende plan", initial: "P" },
    smartsheet: { color: "#00A88E", label: "Smartsheet",        sub: "Lenk til eksisterende plan", initial: "S" },
  } as const;

  return (
    <div style={{ maxWidth: 900, margin: "0 auto", padding: "2rem" }}>
      <button onClick={() => navigate("/projects")} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", marginBottom: "1.5rem", fontSize: "0.9rem", padding: 0 }}>
        ← Tilbake til prosjekter
      </button>

      <div style={{ display: "flex", alignItems: "center", gap: "1.25rem", padding: "1.5rem", borderRadius: 10, background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)", borderTop: `4px solid ${color}`, marginBottom: "2.5rem" }}>
        <div style={{ width: 56, height: 56, borderRadius: "50%", background: color, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: "1.1rem", flexShrink: 0 }}>
          {initials(project.name)}
        </div>
        <div>
          <h1 className="bf-h2" style={{ margin: 0 }}>{project.name}</h1>
          {project.description && <p style={{ color: "var(--bfc-base-c-2)", margin: "0.25rem 0 0", fontSize: "0.95rem" }}>{project.description}</p>}
        </div>
      </div>

      <Section type="projectplan" items={projectPlans}
        onNew={() => { setPpSource(null); setPpTitle(""); setPpUrl(""); setPpOwnStep("choice"); setOwnSelectedTemplate(null); setNewPlanModal(true); }}
        onOpen={(id) => navigate(`/projects/${projectId}/project-plan/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "projectplan" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "projectplan" })}
        countLabel={(pp) => { const src = (pp as ProjectPlan).source; return src === "planner" ? "Planner" : src === "smartsheet" ? "Smartsheet" : `${(pp as ProjectPlan).tasks.length} oppgaver`; }}
      />
      <Section type="oppgave" items={oppgaveLister}
        onNew={() => { setOlSource(null); setOlTitle(""); setOlUrl(""); setOlOwnStep("choice"); setOwnSelectedTemplate(null); setNewOppgaveModal(true); }}
        onOpen={(id) => navigate(`/projects/${projectId}/oppgave/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "oppgave" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "oppgave" })}
        countLabel={(ol) => { const src = (ol as OppgaveListe).source; return src === "planner" ? "Planner" : src === "smartsheet" ? "Smartsheet" : `${(ol as OppgaveListe).oppgaver.length} oppgaver`; }}
      />
      <Section type="risk" items={riskMatrices}
        onNew={() => openSimpleModal("risk")}
        onOpen={(id) => navigate(`/projects/${projectId}/risk-matrix/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "risk" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "risk" })}
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
        countLabel={(p) => `${(p as MeetingPlan).meetings.length} møter`}
      />
      <Section type="runbook" items={runbooks}
        onNew={() => { setRbSource(null); setRbTitle(""); setRbUrl(""); setRbOwnStep("choice"); setOwnSelectedTemplate(null); setNewRunbookModal(true); }}
        onOpen={(id) => navigate(`/projects/${projectId}/runbook/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "runbook" as PlanType }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "runbook" as PlanType })}
        countLabel={(rb) => { const src = (rb as Runbook).source; return src === "planner" ? "Planner" : src === "smartsheet" ? "Smartsheet" : `${(rb as Runbook).activities.length} aktiviteter`; }}
      />

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
    </div>
  );
}

function Section({ type, items, onNew, onOpen, onRename, onDelete, countLabel }: {
  type: PlanType; items: PlanItem[]; onNew: () => void;
  onOpen: (id: string) => void; onRename: (item: PlanItem) => void;
  onDelete: (item: PlanItem) => void; countLabel: (item: object) => string;
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
          {items.map((item) => <PlanRow key={item.id} item={item} color={color} countLabel={countLabel} onOpen={() => onOpen(item.id)} onRename={() => onRename(item)} onDelete={() => onDelete(item)} />)}
        </div>
      )}
    </div>
  );
}

function PlanRow({ item, color, countLabel, onOpen, onRename, onDelete }: {
  item: PlanItem; color: string; countLabel: (item: object) => string;
  onOpen: () => void; onRename: () => void; onDelete: () => void;
}) {
  const [hovered, setHovered] = useState(false);
  return (
    <div
      onClick={onOpen}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{ display: "flex", alignItems: "center", padding: "0.85rem 1.1rem", borderRadius: 7, background: "var(--bfc-base-3)", border: "1px solid var(--bfc-base-dimmed)", borderLeft: `3px solid ${hovered ? color : "var(--bfc-base-dimmed)"}`, cursor: "pointer", gap: "0.75rem", boxShadow: hovered ? "0 2px 10px rgba(0,0,0,0.08)" : "none", transform: hovered ? "translateX(2px)" : "translateX(0)", transition: "border-color 0.15s, box-shadow 0.15s, transform 0.15s" }}
    >
      <span style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: hovered ? 600 : 400, transition: "font-weight 0.1s" }}>
        {item.title}
      </span>
      <div style={{ display: "flex", alignItems: "center", gap: "0.6rem", flexShrink: 0 }} onClick={(e) => e.stopPropagation()}>
        <span style={{ fontSize: "0.75rem", fontWeight: 600, padding: "2px 10px", borderRadius: 20, background: `${color}18`, color }}>{countLabel(item)}</span>
        <button onClick={onRename} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "4px 8px", borderRadius: 4, fontSize: "0.8rem", fontWeight: 500, opacity: hovered ? 1 : 0.4, transition: "opacity 0.15s, color 0.15s" }} onMouseEnter={(e) => (e.currentTarget.style.color = color)} onMouseLeave={(e) => (e.currentTarget.style.color = "var(--bfc-base-c-2)")}>Endre</button>
        <div style={{ width: 1, height: 14, background: "var(--bfc-base-dimmed)" }} />
        <button onClick={onDelete} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", padding: "4px 8px", borderRadius: 4, fontSize: "0.8rem", fontWeight: 500, opacity: hovered ? 1 : 0.4, transition: "opacity 0.15s, color 0.15s" }} onMouseEnter={(e) => (e.currentTarget.style.color = "#E03131")} onMouseLeave={(e) => (e.currentTarget.style.color = "var(--bfc-base-c-2)")}>Slett</button>
      </div>
    </div>
  );
}
