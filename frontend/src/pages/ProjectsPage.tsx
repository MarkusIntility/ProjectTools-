import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type Project, type RiskItem } from "../api/client";

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

function riskLevel(score: number): "low" | "medium" | "high" {
  if (score <= 6) return "low";
  if (score <= 14) return "medium";
  return "high";
}

interface ProjectStats {
  openRisks: number;
  highestRiskScore: number;
  taskDone: number;
  taskTotal: number;
}

const RISK_LEVEL_CONFIG = {
  none:   { color: "#868E96", label: "Ingen risikoer" },
  low:    { color: "#2F9E44", label: "Lav risiko" },
  medium: { color: "#F76707", label: "Middels risiko" },
  high:   { color: "#E03131", label: "Høy risiko" },
};

function projectRiskConfig(stats: ProjectStats) {
  if (stats.openRisks === 0) return RISK_LEVEL_CONFIG.none;
  return RISK_LEVEL_CONFIG[riskLevel(stats.highestRiskScore)];
}

// ─── Stats bar ────────────────────────────────────────────────────────────────

function StatsBar({ stats }: { stats: ProjectStats | undefined }) {
  if (!stats) {
    return (
      <div style={{ display: "flex", gap: "0.6rem", marginTop: "0.85rem", opacity: 0.4 }}>
        <div style={{ height: 20, width: 90, borderRadius: 4, background: "var(--bfc-base-dimmed)" }} />
        <div style={{ height: 20, width: 110, borderRadius: 4, background: "var(--bfc-base-dimmed)" }} />
      </div>
    );
  }

  const riskCfg = projectRiskConfig(stats);
  const pct = stats.taskTotal > 0 ? Math.round((stats.taskDone / stats.taskTotal) * 100) : null;

  return (
    <div style={{ display: "flex", alignItems: "center", gap: "0.65rem", marginTop: "0.85rem", flexWrap: "wrap" }}>
      {/* Risk chip */}
      <div style={{
        display: "flex", alignItems: "center", gap: "0.3rem",
        fontSize: "0.75rem", fontWeight: 600,
        padding: "2px 8px", borderRadius: 20,
        background: `${riskCfg.color}18`, color: riskCfg.color,
        border: `1px solid ${riskCfg.color}30`,
        flexShrink: 0,
      }}>
        <span style={{ width: 7, height: 7, borderRadius: "50%", background: riskCfg.color, display: "inline-block" }} />
        {stats.openRisks > 0 ? `${stats.openRisks} åpne · ${riskCfg.label.toLowerCase()}` : riskCfg.label}
      </div>

      {/* Task progress */}
      {pct !== null && (
        <div style={{ display: "flex", alignItems: "center", gap: "0.4rem", flexShrink: 0 }}>
          <div style={{ width: 72, height: 5, borderRadius: 3, background: "var(--bfc-base-dimmed)", overflow: "hidden" }}>
            <div style={{
              height: "100%", borderRadius: 3,
              width: `${pct}%`,
              background: pct === 100 ? "#2F9E44" : "#1971C2",
              transition: "width 0.3s ease",
            }} />
          </div>
          <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-2)", fontWeight: 500, whiteSpace: "nowrap" }}>
            {pct}% ferdig
          </span>
        </div>
      )}

      {pct === null && stats.taskTotal === 0 && (
        <span style={{ fontSize: "0.75rem", color: "var(--bfc-base-c-3)" }}>Ingen oppgaver</span>
      )}
    </div>
  );
}

// ─── Project card ─────────────────────────────────────────────────────────────

function ProjectCard({ project, stats, onClick }: {
  project: Project;
  stats: ProjectStats | undefined;
  onClick: () => void;
}) {
  const color = accentColor(project.name);
  const [hovered, setHovered] = useState(false);

  return (
    <div
      onClick={onClick}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        cursor: "pointer", borderRadius: 8,
        background: "var(--bfc-base-3)",
        border: "1px solid var(--bfc-base-dimmed)",
        borderLeft: `4px solid ${color}`,
        padding: "1.25rem 1.5rem",
        display: "flex", gap: "1rem", alignItems: "flex-start",
        boxShadow: hovered ? "0 6px 20px rgba(0,0,0,0.12)" : "0 1px 4px rgba(0,0,0,0.06)",
        transform: hovered ? "translateY(-2px)" : "translateY(0)",
        transition: "box-shadow 0.15s ease, transform 0.15s ease",
      }}
    >
      <div style={{
        width: 46, height: 46, borderRadius: "50%",
        background: color, color: "#fff",
        display: "flex", alignItems: "center", justifyContent: "center",
        fontWeight: 700, fontSize: "1rem", flexShrink: 0, letterSpacing: "0.04em",
      }}>
        {initials(project.name)}
      </div>

      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "0.75rem" }}>
          <h2 className="bf-h4" style={{ margin: 0 }}>{project.name}</h2>
          <span style={{
            fontSize: "0.75rem", fontWeight: 600, padding: "2px 10px",
            borderRadius: 20, background: `${color}1A`, color, flexShrink: 0,
          }}>
            Aktivt
          </span>
        </div>

        {project.description && (
          <p style={{
            color: "var(--bfc-base-c-2)", marginTop: "0.3rem", fontSize: "0.9rem",
            overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
          }}>
            {project.description}
          </p>
        )}

        <StatsBar stats={stats} />

        <div style={{ marginTop: "0.6rem", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <span style={{ fontSize: "0.78rem", color: "var(--bfc-base-c-3)" }}>
            Opprettet{" "}
            {new Date(project.created_at).toLocaleDateString("nb-NO", { day: "numeric", month: "long", year: "numeric" })}
          </span>
          <span style={{ fontSize: "0.82rem", fontWeight: 600, color: hovered ? color : "var(--bfc-base-c-2)", transition: "color 0.15s ease" }}>
            Åpne &rarr;
          </span>
        </div>
      </div>
    </div>
  );
}

// ─── Page ─────────────────────────────────────────────────────────────────────

export default function ProjectsPage() {
  const navigate = useNavigate();
  const [projects, setProjects] = useState<Project[]>([]);
  const [loading, setLoading] = useState(true);
  const [statsMap, setStatsMap] = useState<Record<string, ProjectStats>>({});
  const [modalOpen, setModalOpen] = useState(false);
  const [name, setName] = useState("");
  const [description, setDescription] = useState("");
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    api.projects.list().then(setProjects).finally(() => setLoading(false));
  }, []);

  useEffect(() => {
    if (projects.length === 0) return;
    const fetchStats = async () => {
      const entries = await Promise.all(
        projects.map(async (p) => {
          const [matrices, plans, oppgaveLister, runbooks] = await Promise.all([
            api.riskMatrices.list(p.id),
            api.projectPlans.list(p.id),
            api.oppgaveLister.list(p.id),
            api.runbooks.list(p.id),
          ]);

          const allRisks: RiskItem[] = matrices.flatMap((m) => m.risks).filter((r) => r.status === "open");
          const highestRiskScore = allRisks.reduce((max, r) => Math.max(max, r.risk_score), 0);

          let taskDone = 0;
          let taskTotal = 0;

          for (const plan of plans.filter((pl) => pl.source === "own")) {
            taskDone += plan.tasks.filter((t) => t.percent_complete === 100).length;
            taskTotal += plan.tasks.length;
          }
          for (const liste of oppgaveLister.filter((ol) => ol.source === "own")) {
            taskDone += liste.oppgaver.filter((o) => o.status === "done").length;
            taskTotal += liste.oppgaver.length;
          }
          for (const rb of runbooks.filter((r) => r.source === "own")) {
            taskDone += rb.activities.filter((a) => a.status === "done").length;
            taskTotal += rb.activities.length;
          }

          return [p.id, { openRisks: allRisks.length, highestRiskScore, taskDone, taskTotal }] as [string, ProjectStats];
        })
      );
      setStatsMap(Object.fromEntries(entries));
    };
    fetchStats();
  }, [projects]);

  async function handleCreate() {
    if (!name.trim()) return;
    setSaving(true);
    try {
      const project = await api.projects.create({ name, description: description || undefined });
      setProjects((prev) => [project, ...prev]);
      setModalOpen(false);
      setName("");
      setDescription("");
    } finally {
      setSaving(false);
    }
  }

  return (
    <div style={{ maxWidth: 960, margin: "0 auto", padding: "2.5rem 2rem" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "2rem" }}>
        <div>
          <h1 className="bf-h1" style={{ margin: 0 }}>Prosjekter</h1>
          {!loading && projects.length > 0 && (
            <p style={{ margin: "0.25rem 0 0", color: "var(--bfc-base-c-2)", fontSize: "0.9rem" }}>
              {projects.length} {projects.length === 1 ? "prosjekt" : "prosjekter"}
            </p>
          )}
        </div>
        <Button variant="filled" onClick={() => setModalOpen(true)}>+ Nytt prosjekt</Button>
      </div>

      {loading ? (
        <p style={{ color: "var(--bfc-base-c-2)" }}>Laster...</p>
      ) : projects.length === 0 ? (
        <div style={{ textAlign: "center", padding: "3rem", border: "2px dashed var(--bfc-base-dimmed)", borderRadius: 8, color: "var(--bfc-base-c-2)" }}>
          <p style={{ fontSize: "1.1rem", marginBottom: "1rem" }}>Ingen prosjekter ennå</p>
          <Button variant="filled" onClick={() => setModalOpen(true)}>Opprett ditt første prosjekt</Button>
        </div>
      ) : (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(400px, 1fr))", gap: "1rem" }}>
          {projects.map((project) => (
            <ProjectCard
              key={project.id}
              project={project}
              stats={statsMap[project.id]}
              onClick={() => navigate(`/projects/${project.id}`)}
            />
          ))}
        </div>
      )}

      <Modal isOpen={modalOpen} onRequestClose={() => setModalOpen(false)} header="Nytt prosjekt">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <Input label="Prosjektnavn" value={name} onChange={(e) => setName(e.target.value)} placeholder="f.eks. IT-infrastruktur 2025" />
          <Input label="Beskrivelse (valgfri)" value={description} onChange={(e) => setDescription(e.target.value)} placeholder="Kort beskrivelse av prosjektet" />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", marginTop: "0.5rem" }}>
            <Button onClick={() => setModalOpen(false)}>Avbryt</Button>
            <Button variant="filled" onClick={handleCreate} state={!name.trim() || saving ? "inactive" : "default"}>
              {saving ? "Oppretter..." : "Opprett"}
            </Button>
          </div>
        </div>
      </Modal>
    </div>
  );
}
