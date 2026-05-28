import { useEffect, useState } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { Button, Card, Badge } from "@intility/bifrost-react";
import { api, type Project, type RiskMatrix, type CommunicationPlan, type MeetingPlan } from "../api/client";

export default function ProjectDetailPage() {
  const { projectId } = useParams<{ projectId: string }>();
  const navigate = useNavigate();
  const [project, setProject] = useState<Project | null>(null);
  const [riskMatrices, setRiskMatrices] = useState<RiskMatrix[]>([]);
  const [commPlans, setCommPlans] = useState<CommunicationPlan[]>([]);
  const [meetingPlans, setMeetingPlans] = useState<MeetingPlan[]>([]);

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
      {project.description && <p style={{ color: "var(--bfc-base-c-2)", marginBottom: "2rem" }}>{project.description}</p>}

      <Section
        title="Risikomatriser"
        items={riskMatrices}
        onNew={() => createAndNavigate("risk")}
        onOpen={(id) => navigate(`/projects/${projectId}/risk-matrix/${id}`)}
        countLabel={(m) => `${(m as RiskMatrix).risks.length} risikoer`}
      />

      <Section
        title="Kommunikasjonsplaner"
        items={commPlans}
        onNew={() => createAndNavigate("comm")}
        onOpen={(id) => navigate(`/projects/${projectId}/communication-plan/${id}`)}
        countLabel={(p) => `${(p as CommunicationPlan).entries.length} oppføringer`}
      />

      <Section
        title="Møteplaner"
        items={meetingPlans}
        onNew={() => createAndNavigate("meeting")}
        onOpen={(id) => navigate(`/projects/${projectId}/meeting-plan/${id}`)}
        countLabel={(p) => `${(p as MeetingPlan).meetings.length} møter`}
      />
    </div>
  );
}

function Section({
  title,
  items,
  onNew,
  onOpen,
  countLabel,
}: {
  title: string;
  items: { id: string; title: string }[];
  onNew: () => void;
  onOpen: (id: string) => void;
  countLabel: (item: object) => string;
}) {
  return (
    <div style={{ marginBottom: "2rem" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "0.75rem" }}>
        <h2 className="bf-h3">{title}</h2>
        <Button variant="outline" onClick={onNew}>+ Ny</Button>
      </div>
      {items.length === 0 ? (
        <Card><p style={{ color: "var(--bfc-base-c-2)" }}>Ingen {title.toLowerCase()} ennå.</p></Card>
      ) : (
        <div style={{ display: "grid", gap: "0.5rem" }}>
          {items.map((item) => (
            <Card key={item.id} style={{ cursor: "pointer" }} onClick={() => onOpen(item.id)}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span>{item.title}</span>
                <Badge state="neutral">{countLabel(item)}</Badge>
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}
