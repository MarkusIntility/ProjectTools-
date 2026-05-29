import { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type Project, type RiskMatrix, type CommunicationPlan, type MeetingPlan } from "../api/client";

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
  risk:    { color: "#E03131", label: "Risikomatriser" },
  comm:    { color: "#1971C2", label: "Kommunikasjonsplaner" },
  meeting: { color: "#2F9E44", label: "Møteplaner" },
};

type PlanType = "risk" | "comm" | "meeting";
type PlanItem = { id: string; title: string };

export default function ProjectDetailPage() {
  const { projectId } = useParams<{ projectId: string }>();
  const navigate = useNavigate();
  const [project, setProject] = useState<Project | null>(null);
  const [riskMatrices, setRiskMatrices] = useState<RiskMatrix[]>([]);
  const [commPlans, setCommPlans] = useState<CommunicationPlan[]>([]);
  const [meetingPlans, setMeetingPlans] = useState<MeetingPlan[]>([]);

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
    ]).then(([p, rm, cp, mp]) => {
      setProject(p);
      setRiskMatrices(rm);
      setCommPlans(cp);
      setMeetingPlans(mp);
    });
  }, [projectId]);

  async function createAndNavigate(type: PlanType) {
    if (!projectId) return;
    if (type === "risk") {
      const m = await api.riskMatrices.create(projectId, { title: "Ny risikomatrise" });
      navigate(`/projects/${projectId}/risk-matrix/${m.id}`);
    } else if (type === "comm") {
      const p = await api.communicationPlans.create(projectId, { title: "Ny kommunikasjonsplan" });
      navigate(`/projects/${projectId}/communication-plan/${p.id}`);
    } else {
      const p = await api.meetingPlans.create(projectId, { title: "Ny møteplan" });
      navigate(`/projects/${projectId}/meeting-plan/${p.id}`);
    }
  }

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
      } else {
        const u = await api.meetingPlans.update(projectId, renameTarget.id, { title: renameValue });
        setMeetingPlans((prev) => prev.map((p) => (p.id === u.id ? u : p)));
      }
      setRenameTarget(null);
    } finally {
      setRenameSaving(false);
    }
  }

  async function handleDelete() {
    if (!projectId || !deleteTarget) return;
    setDeleting(true);
    try {
      if (deleteTarget.type === "risk") {
        await api.riskMatrices.delete(projectId, deleteTarget.id);
        setRiskMatrices((prev) => prev.filter((m) => m.id !== deleteTarget.id));
      } else if (deleteTarget.type === "comm") {
        await api.communicationPlans.delete(projectId, deleteTarget.id);
        setCommPlans((prev) => prev.filter((p) => p.id !== deleteTarget.id));
      } else {
        await api.meetingPlans.delete(projectId, deleteTarget.id);
        setMeetingPlans((prev) => prev.filter((p) => p.id !== deleteTarget.id));
      }
      setDeleteTarget(null);
    } finally {
      setDeleting(false);
    }
  }

  if (!project) return <div style={{ padding: "2rem" }}>Laster...</div>;

  const color = accentColor(project.name);

  return (
    <div style={{ maxWidth: 900, margin: "0 auto", padding: "2rem" }}>
      {/* Back */}
      <button
        onClick={() => navigate("/projects")}
        style={{
          background: "none", border: "none", cursor: "pointer",
          color: "var(--bfc-base-c-2)", marginBottom: "1.5rem",
          fontSize: "0.9rem", padding: 0,
        }}
      >
        ← Tilbake til prosjekter
      </button>

      {/* Project header */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: "1.25rem",
          padding: "1.5rem",
          borderRadius: 10,
          background: "var(--bfc-base-3)",
          border: "1px solid var(--bfc-base-dimmed)",
          borderTop: `4px solid ${color}`,
          marginBottom: "2.5rem",
        }}
      >
        <div
          style={{
            width: 56, height: 56, borderRadius: "50%",
            background: color, color: "#fff",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontWeight: 700, fontSize: "1.1rem", flexShrink: 0,
          }}
        >
          {initials(project.name)}
        </div>
        <div>
          <h1 className="bf-h2" style={{ margin: 0 }}>{project.name}</h1>
          {project.description && (
            <p style={{ color: "var(--bfc-base-c-2)", margin: "0.25rem 0 0", fontSize: "0.95rem" }}>
              {project.description}
            </p>
          )}
        </div>
      </div>

      {/* Sections */}
      <Section
        type="risk"
        items={riskMatrices}
        onNew={() => createAndNavigate("risk")}
        onOpen={(id) => navigate(`/projects/${projectId}/risk-matrix/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "risk" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "risk" })}
        countLabel={(m) => `${(m as RiskMatrix).risks.length} risikoer`}
      />
      <Section
        type="comm"
        items={commPlans}
        onNew={() => createAndNavigate("comm")}
        onOpen={(id) => navigate(`/projects/${projectId}/communication-plan/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "comm" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "comm" })}
        countLabel={(p) => `${(p as CommunicationPlan).entries.length} oppføringer`}
      />
      <Section
        type="meeting"
        items={meetingPlans}
        onNew={() => createAndNavigate("meeting")}
        onOpen={(id) => navigate(`/projects/${projectId}/meeting-plan/${id}`)}
        onRename={(item) => { setRenameTarget({ ...item, type: "meeting" }); setRenameValue(item.title); }}
        onDelete={(item) => setDeleteTarget({ ...item, type: "meeting" })}
        countLabel={(p) => `${(p as MeetingPlan).meetings.length} møter`}
      />

      {/* Rename modal */}
      <Modal isOpen={!!renameTarget} onRequestClose={() => setRenameTarget(null)} header="Gi nytt navn">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <Input
            label="Navn"
            value={renameValue}
            onChange={(e) => setRenameValue(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleRename()}
            autoFocus
          />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setRenameTarget(null)}>Avbryt</Button>
            <Button variant="filled" onClick={handleRename} state={!renameValue.trim() || renameSaving ? "inactive" : "default"}>
              {renameSaving ? "Lagrer..." : "Lagre"}
            </Button>
          </div>
        </div>
      </Modal>

      {/* Delete confirmation */}
      <Modal isOpen={!!deleteTarget} onRequestClose={() => setDeleteTarget(null)} header="Bekreft sletting">
        <div style={{ display: "flex", flexDirection: "column", gap: "1.25rem" }}>
          <p style={{ margin: 0 }}>
            Er du sikker på at du vil slette <strong>«{deleteTarget?.title}»</strong>?
            {" "}Dette kan ikke angres.
          </p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setDeleteTarget(null)}>Avbryt</Button>
            <Button
              variant="filled"
              state={deleting ? "inactive" : "default"}
              onClick={handleDelete}
              style={{ background: "#E03131", borderColor: "#E03131" }}
            >
              {deleting ? "Sletter..." : "Ja, slett"}
            </Button>
          </div>
        </div>
      </Modal>
    </div>
  );
}

function Section({
  type, items, onNew, onOpen, onRename, onDelete, countLabel,
}: {
  type: PlanType;
  items: PlanItem[];
  onNew: () => void;
  onOpen: (id: string) => void;
  onRename: (item: PlanItem) => void;
  onDelete: (item: PlanItem) => void;
  countLabel: (item: object) => string;
}) {
  const { color, label } = SECTION_CONFIG[type];

  return (
    <div style={{ marginBottom: "2rem" }}>
      {/* Section header */}
      <div
        style={{
          display: "flex", justifyContent: "space-between", alignItems: "center",
          marginBottom: "0.75rem",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: "0.65rem" }}>
          <div style={{ width: 4, height: 22, borderRadius: 2, background: color, flexShrink: 0 }} />
          <h2 className="bf-h3" style={{ margin: 0 }}>{label}</h2>
          {items.length > 0 && (
            <span
              style={{
                fontSize: "0.75rem", fontWeight: 600,
                padding: "1px 8px", borderRadius: 20,
                background: `${color}18`, color: color,
              }}
            >
              {items.length}
            </span>
          )}
        </div>
        <Button variant="outline" onClick={onNew}>+ Ny</Button>
      </div>

      {/* Items */}
      {items.length === 0 ? (
        <div
          style={{
            padding: "1rem 1.25rem",
            borderRadius: 6,
            border: `1px dashed ${color}60`,
            color: "var(--bfc-base-c-2)",
            fontSize: "0.9rem",
            background: `${color}06`,
          }}
        >
          Ingen {label.toLowerCase()} ennå. Klikk «+ Ny» for å opprette.
        </div>
      ) : (
        <div style={{ display: "grid", gap: "0.5rem" }}>
          {items.map((item) => (
            <PlanRow
              key={item.id}
              item={item}
              color={color}
              countLabel={countLabel}
              onOpen={() => onOpen(item.id)}
              onRename={() => onRename(item)}
              onDelete={() => onDelete(item)}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function PlanRow({
  item, color, countLabel, onOpen, onRename, onDelete,
}: {
  item: PlanItem;
  color: string;
  countLabel: (item: object) => string;
  onOpen: () => void;
  onRename: () => void;
  onDelete: () => void;
}) {
  const [hovered, setHovered] = useState(false);

  return (
    <div
      onClick={onOpen}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        display: "flex",
        alignItems: "center",
        padding: "0.85rem 1.1rem",
        borderRadius: 7,
        background: "var(--bfc-base-3)",
        border: "1px solid var(--bfc-base-dimmed)",
        borderLeft: `3px solid ${hovered ? color : "var(--bfc-base-dimmed)"}`,
        cursor: "pointer",
        gap: "0.75rem",
        boxShadow: hovered ? "0 2px 10px rgba(0,0,0,0.08)" : "none",
        transform: hovered ? "translateX(2px)" : "translateX(0)",
        transition: "border-color 0.15s, box-shadow 0.15s, transform 0.15s",
      }}
    >
      <span
        style={{
          flex: 1, minWidth: 0,
          overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
          fontWeight: hovered ? 600 : 400,
          transition: "font-weight 0.1s",
        }}
      >
        {item.title}
      </span>

      <div
        style={{ display: "flex", alignItems: "center", gap: "0.6rem", flexShrink: 0 }}
        onClick={(e) => e.stopPropagation()}
      >
        <span
          style={{
            fontSize: "0.75rem", fontWeight: 600,
            padding: "2px 10px", borderRadius: 20,
            background: `${color}18`, color: color,
          }}
        >
          {countLabel(item)}
        </span>

        <button
          onClick={onRename}
          title="Gi nytt navn"
          style={{
            background: "none", border: "none", cursor: "pointer",
            color: "var(--bfc-base-c-2)",
            padding: "4px 8px", borderRadius: 4,
            fontSize: "0.8rem", fontWeight: 500,
            opacity: hovered ? 1 : 0.4,
            transition: "opacity 0.15s, color 0.15s",
          }}
          onMouseEnter={(e) => (e.currentTarget.style.color = color)}
          onMouseLeave={(e) => (e.currentTarget.style.color = "var(--bfc-base-c-2)")}
        >
          Endre
        </button>

        <div style={{ width: 1, height: 14, background: "var(--bfc-base-dimmed)" }} />

        <button
          onClick={onDelete}
          title="Slett"
          style={{
            background: "none", border: "none", cursor: "pointer",
            color: "var(--bfc-base-c-2)",
            padding: "4px 8px", borderRadius: 4,
            fontSize: "0.8rem", fontWeight: 500,
            opacity: hovered ? 1 : 0.4,
            transition: "opacity 0.15s, color 0.15s",
          }}
          onMouseEnter={(e) => (e.currentTarget.style.color = "#E03131")}
          onMouseLeave={(e) => (e.currentTarget.style.color = "var(--bfc-base-c-2)")}
        >
          Slett
        </button>
      </div>
    </div>
  );
}
