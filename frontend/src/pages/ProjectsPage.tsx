import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Button, Input, Modal } from "@intility/bifrost-react";
import { api, type Project } from "../api/client";

const ACCENT_COLORS = [
  "#4C6EF5",
  "#7950F2",
  "#E64980",
  "#F76707",
  "#2F9E44",
  "#1098AD",
  "#E67700",
  "#862E9C",
];

function accentColor(name: string): string {
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = name.charCodeAt(i) + ((hash << 5) - hash);
  }
  return ACCENT_COLORS[Math.abs(hash) % ACCENT_COLORS.length];
}

function initials(name: string): string {
  return name
    .split(" ")
    .map((w) => w[0])
    .slice(0, 2)
    .join("")
    .toUpperCase();
}

function ProjectCard({ project, onClick }: { project: Project; onClick: () => void }) {
  const color = accentColor(project.name);
  const [hovered, setHovered] = useState(false);

  return (
    <div
      onClick={onClick}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        cursor: "pointer",
        borderRadius: 8,
        background: "var(--bfc-base-3)",
        border: "1px solid var(--bfc-base-dimmed)",
        borderLeft: `4px solid ${color}`,
        padding: "1.25rem 1.5rem",
        display: "flex",
        gap: "1rem",
        alignItems: "flex-start",
        boxShadow: hovered
          ? "0 6px 20px rgba(0,0,0,0.12)"
          : "0 1px 4px rgba(0,0,0,0.06)",
        transform: hovered ? "translateY(-2px)" : "translateY(0)",
        transition: "box-shadow 0.15s ease, transform 0.15s ease",
      }}
    >
      <div
        style={{
          width: 46,
          height: 46,
          borderRadius: "50%",
          background: color,
          color: "#fff",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontWeight: 700,
          fontSize: "1rem",
          flexShrink: 0,
          letterSpacing: "0.04em",
        }}
      >
        {initials(project.name)}
      </div>

      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "0.75rem" }}>
          <h2 className="bf-h4" style={{ margin: 0 }}>
            {project.name}
          </h2>
          <span
            style={{
              fontSize: "0.75rem",
              fontWeight: 600,
              padding: "2px 10px",
              borderRadius: 20,
              background: `${color}1A`,
              color: color,
              flexShrink: 0,
            }}
          >
            Aktivt
          </span>
        </div>

        {project.description && (
          <p
            style={{
              color: "var(--bfc-base-c-2)",
              marginTop: "0.3rem",
              fontSize: "0.9rem",
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
            }}
          >
            {project.description}
          </p>
        )}

        <div
          style={{
            marginTop: "0.85rem",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
          }}
        >
          <span style={{ fontSize: "0.78rem", color: "var(--bfc-base-c-3)" }}>
            Opprettet{" "}
            {new Date(project.created_at).toLocaleDateString("nb-NO", {
              day: "numeric",
              month: "long",
              year: "numeric",
            })}
          </span>
          <span
            style={{
              fontSize: "0.82rem",
              fontWeight: 600,
              color: hovered ? color : "var(--bfc-base-c-2)",
              transition: "color 0.15s ease",
            }}
          >
            Åpne &rarr;
          </span>
        </div>
      </div>
    </div>
  );
}

export default function ProjectsPage() {
  const navigate = useNavigate();
  const [projects, setProjects] = useState<Project[]>([]);
  const [loading, setLoading] = useState(true);
  const [modalOpen, setModalOpen] = useState(false);
  const [name, setName] = useState("");
  const [description, setDescription] = useState("");
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    api.projects.list().then(setProjects).finally(() => setLoading(false));
  }, []);

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
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: "2rem",
        }}
      >
        <div>
          <h1 className="bf-h1" style={{ margin: 0 }}>
            Prosjekter
          </h1>
          {!loading && projects.length > 0 && (
            <p style={{ margin: "0.25rem 0 0", color: "var(--bfc-base-c-2)", fontSize: "0.9rem" }}>
              {projects.length} {projects.length === 1 ? "prosjekt" : "prosjekter"}
            </p>
          )}
        </div>
        <Button variant="filled" onClick={() => setModalOpen(true)}>
          + Nytt prosjekt
        </Button>
      </div>

      {loading ? (
        <p style={{ color: "var(--bfc-base-c-2)" }}>Laster...</p>
      ) : projects.length === 0 ? (
        <div
          style={{
            textAlign: "center",
            padding: "3rem",
            border: "2px dashed var(--bfc-base-dimmed)",
            borderRadius: 8,
            color: "var(--bfc-base-c-2)",
          }}
        >
          <p style={{ fontSize: "1.1rem", marginBottom: "1rem" }}>Ingen prosjekter ennå</p>
          <Button variant="filled" onClick={() => setModalOpen(true)}>
            Opprett ditt første prosjekt
          </Button>
        </div>
      ) : (
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(400px, 1fr))",
            gap: "1rem",
          }}
        >
          {projects.map((project) => (
            <ProjectCard
              key={project.id}
              project={project}
              onClick={() => navigate(`/projects/${project.id}`)}
            />
          ))}
        </div>
      )}

      <Modal isOpen={modalOpen} onRequestClose={() => setModalOpen(false)} header="Nytt prosjekt">
        <div style={{ display: "flex", flexDirection: "column", gap: "1rem" }}>
          <Input
            label="Prosjektnavn"
            value={name}
            onChange={(e) => setName(e.target.value)}
            placeholder="f.eks. IT-infrastruktur 2025"
          />
          <Input
            label="Beskrivelse (valgfri)"
            value={description}
            onChange={(e) => setDescription(e.target.value)}
            placeholder="Kort beskrivelse av prosjektet"
          />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", marginTop: "0.5rem" }}>
            <Button onClick={() => setModalOpen(false)}>Avbryt</Button>
            <Button
              variant="filled"
              onClick={handleCreate}
              state={!name.trim() || saving ? "inactive" : "default"}
            >
              {saving ? "Oppretter..." : "Opprett"}
            </Button>
          </div>
        </div>
      </Modal>
    </div>
  );
}
