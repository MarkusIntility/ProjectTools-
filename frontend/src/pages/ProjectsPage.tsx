import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Button, Card, Input, Modal, Badge } from "@intility/bifrost-react";
import { api, type Project } from "../api/client";

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
    <div style={{ maxWidth: 900, margin: "0 auto", padding: "2rem" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
        <h1 className="bf-h1">Prosjekter</h1>
        <Button variant="filled" onClick={() => setModalOpen(true)}>
          Nytt prosjekt
        </Button>
      </div>

      {loading ? (
        <p>Laster...</p>
      ) : projects.length === 0 ? (
        <Card>
          <p style={{ textAlign: "center", color: "var(--bfc-base-c-2)" }}>
            Ingen prosjekter ennå. Opprett ditt første prosjekt.
          </p>
        </Card>
      ) : (
        <div style={{ display: "grid", gap: "1rem" }}>
          {projects.map((project) => (
            <Card
              key={project.id}
              style={{ cursor: "pointer" }}
              onClick={() => navigate(`/projects/${project.id}`)}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <h2 className="bf-h4">{project.name}</h2>
                  {project.description && (
                    <p style={{ color: "var(--bfc-base-c-2)", marginTop: "0.25rem" }}>{project.description}</p>
                  )}
                  <p style={{ fontSize: "0.85rem", color: "var(--bfc-base-c-3)", marginTop: "0.5rem" }}>
                    Opprettet {new Date(project.created_at).toLocaleDateString("nb-NO")}
                  </p>
                </div>
                <Badge state="default">Aktivt</Badge>
              </div>
            </Card>
          ))}
        </div>
      )}

      <Modal
        isOpen={modalOpen}
        onRequestClose={() => setModalOpen(false)}
        header="Nytt prosjekt"
      >
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
            <Button variant="filled" onClick={handleCreate} state={!name.trim() || saving ? "inactive" : "default"}>
              {saving ? "Oppretter..." : "Opprett"}
            </Button>
          </div>
        </div>
      </Modal>
    </div>
  );
}
