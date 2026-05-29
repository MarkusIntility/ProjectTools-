import { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Badge, Input, Modal } from "@intility/bifrost-react";
import { api, type Project, type RiskMatrix, type CommunicationPlan, type MeetingPlan } from "../api/client";

type PlanItem = { id: string; title: string };

export default function ProjectDetailPage() {
  const { projectId } = useParams<{ projectId: string }>();
  const navigate = useNavigate();
  const [project, setProject] = useState<Project | null>(null);
  const [riskMatrices, setRiskMatrices] = useState<RiskMatrix[]>([]);
  const [commPlans, setCommPlans] = useState<CommunicationPlan[]>([]);
  const [meetingPlans, setMeetingPlans] = useState<MeetingPlan[]>([]);

  // Rename modal
  const [renameTarget, setRenameTarget] = useState<{ id: string; title: string; type: "risk" | "comm" | "meeting" } | null>(null);
  const [renameValue, setRenameValue] = useState("");
  const [renameSaving, setRenameSaving] = useState(false);

  // Delete modal
  const [deleteTarget, setDeleteTarget] = useState<{ id: string; title: string; type: "risk" | "comm" | "meeting" } | null>(null);
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

  async function createAndNavigate(type: "risk" | "comm" | "meeting") {
    if (!projectId) return;
    if (type === "risk") {
      const m = await api.riskMatrices.create(projectId, { title: "Risikomatrise" });
      navigate(`/projects/${projectId}/risk-matrix/${m.id}`);
    } else if (type === "comm") {
      const p = await api.communicationPlans.create(projectId, { title: "Kommunikasjonsplan" });
      navigate(`/projects/${projectId}/communication-plan/${p.id}`);
    } else {
      const p = await api.meetingPlans.create(projectId, { title: "Møteplan" });
      navigate(`/projects/${projectId}/meeting-plan/${p.id}`);
    }
  }

  function openRename(item: PlanItem, type: "risk" | "comm" | "meeting") {
    setRenameTarget({ ...item, type });
    setRenameValue(item.title);
  }

  async function handleRename() {
    if (!projectId || !renameTarget || !renameValue.trim()) return;
    setRenameSaving(true);
    try {
      if (renameTarget.type === "risk") {
        const updated = await api.riskMatrices.update(projectId, renameTarget.id, { title: renameValue });
        setRiskMatrices((prev) => prev.map((m) => (m.id === updated.id ? updated : m)));
      } else if (renameTarget.type === "comm") {
        const updated = await api.communicationPlans.update(projectId, renameTarget.id, { title: renameValue });
        setCommPlans((prev) => prev.map((p) => (p.id === updated.id ? updated : p)));
      } else {
        const updated = await api.meetingPlans.update(projectId, renameTarget.id, { title: renameValue });
        setMeetingPlans((prev) => prev.map((p) => (p.id === updated.id ? updated : p)));
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

  return (
    <div style={{ maxWidth: 900, margin: "0 auto", padding: "2rem" }}>
      <button
        onClick={() => navigate("/projects")}
        style={{ background: "none", border: "none", cursor: "pointer", color: "var(--bfc-base-c-2)", marginBottom: "1rem" }}
      >
        ← Tilbake til prosjekter
      </button>

      <h1 className="bf-h1" style={{ marginBottom: "0.5rem" }}>{project.name}</h1>
      {project.description && (
        <p style={{ color: "var(--bfc-base-c-2)", marginBottom: "2rem" }}>{project.description}</p>
      )}

      <Section
        title="Risikomatriser"
        items={riskMatrices}
        type="risk"
        onNew={() => createAndNavigate("risk")}
        onOpen={(id) => navigate(`/projects/${projectId}/risk-matrix/${id}`)}
        onRename={(item) => openRename(item, "risk")}
        onDelete={(item) => setDeleteTarget({ ...item, type: "risk" })}
        countLabel={(m) => `${(m as RiskMatrix).risks.length} risikoer`}
      />

      <Section
        title="Kommunikasjonsplaner"
        items={commPlans}
        type="comm"
        onNew={() => createAndNavigate("comm")}
        onOpen={(id) => navigate(`/projects/${projectId}/communication-plan/${id}`)}
        onRename={(item) => openRename(item, "comm")}
        onDelete={(item) => setDeleteTarget({ ...item, type: "comm" })}
        countLabel={(p) => `${(p as CommunicationPlan).entries.length} oppføringer`}
      />

      <Section
        title="Møteplaner"
        items={meetingPlans}
        type="meeting"
        onNew={() => createAndNavigate("meeting")}
        onOpen={(id) => navigate(`/projects/${projectId}/meeting-plan/${id}`)}
        onRename={(item) => openRename(item, "meeting")}
        onDelete={(item) => setDeleteTarget({ ...item, type: "meeting" })}
        countLabel={(p) => `${(p as MeetingPlan).meetings.length} møter`}
      />

      {/* Rename modal */}
      <Modal
        isOpen={!!renameTarget}
        onRequestClose={() => setRenameTarget(null)}
        header="Gi nytt navn"
      >
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
            <Button
              variant="filled"
              onClick={handleRename}
              state={!renameValue.trim() || renameSaving ? "inactive" : "default"}
            >
              {renameSaving ? "Lagrer..." : "Lagre"}
            </Button>
          </div>
        </div>
      </Modal>

      {/* Delete confirmation modal */}
      <Modal
        isOpen={!!deleteTarget}
        onRequestClose={() => setDeleteTarget(null)}
        header="Slett"
      >
        <div style={{ display: "flex", flexDirection: "column", gap: "1.25rem" }}>
          <p>
            Er du sikker på at du vil slette <strong>{deleteTarget?.title}</strong>?
            Dette kan ikke angres.
          </p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
            <Button onClick={() => setDeleteTarget(null)}>Avbryt</Button>
            <Button
              variant="filled"
              state={deleting ? "inactive" : "default"}
              onClick={handleDelete}
              style={{ background: "var(--bfc-alert)", borderColor: "var(--bfc-alert)" }}
            >
              {deleting ? "Sletter..." : "Slett"}
            </Button>
          </div>
        </div>
      </Modal>
    </div>
  );
}

function Section({
  title,
  items,
  onNew,
  onOpen,
  onRename,
  onDelete,
  countLabel,
}: {
  title: string;
  items: PlanItem[];
  type: "risk" | "comm" | "meeting";
  onNew: () => void;
  onOpen: (id: string) => void;
  onRename: (item: PlanItem) => void;
  onDelete: (item: PlanItem) => void;
  countLabel: (item: object) => string;
}) {
  return (
    <div style={{ marginBottom: "2rem" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "0.75rem" }}>
        <h2 className="bf-h3">{title}</h2>
        <Button variant="outline" onClick={onNew}>+ Ny</Button>
      </div>

      {items.length === 0 ? (
        <div
          style={{
            padding: "1rem 1.25rem",
            borderRadius: 6,
            border: "1px dashed var(--bfc-base-dimmed)",
            color: "var(--bfc-base-c-2)",
            fontSize: "0.9rem",
          }}
        >
          Ingen {title.toLowerCase()} ennå.
        </div>
      ) : (
        <div style={{ display: "grid", gap: "0.4rem" }}>
          {items.map((item) => (
            <div
              key={item.id}
              onClick={() => onOpen(item.id)}
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                padding: "0.75rem 1rem",
                borderRadius: 6,
                background: "var(--bfc-base-3)",
                border: "1px solid var(--bfc-base-dimmed)",
                cursor: "pointer",
                gap: "0.75rem",
              }}
            >
              <span style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                {item.title}
              </span>

              <div
                style={{ display: "flex", alignItems: "center", gap: "0.5rem", flexShrink: 0 }}
                onClick={(e) => e.stopPropagation()}
              >
                <Badge state="neutral">{countLabel(item)}</Badge>
                <button
                  onClick={() => onRename(item)}
                  title="Gi nytt navn"
                  style={{
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    color: "var(--bfc-base-c-2)",
                    padding: "4px 6px",
                    borderRadius: 4,
                    fontSize: "0.85rem",
                    lineHeight: 1,
                  }}
                >
                  ✏️
                </button>
                <button
                  onClick={() => onDelete(item)}
                  title="Slett"
                  style={{
                    background: "none",
                    border: "none",
                    cursor: "pointer",
                    color: "var(--bfc-alert, #d63031)",
                    padding: "4px 6px",
                    borderRadius: 4,
                    fontSize: "0.85rem",
                    lineHeight: 1,
                  }}
                >
                  🗑️
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
