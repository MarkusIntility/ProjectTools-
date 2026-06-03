import { Routes, Route, Navigate } from "react-router-dom";
import ProjectsPage from "./pages/ProjectsPage";
import ProjectDetailPage from "./pages/ProjectDetailPage";
import RiskMatrixPage from "./pages/RiskMatrixPage";
import CommunicationPlanPage from "./pages/CommunicationPlanPage";
import MeetingPlanPage from "./pages/MeetingPlanPage";
import RunbookPage from "./pages/RunbookPage";
import ProjectPlanPage from "./pages/ProjectPlanPage";
import OppgavePage from "./pages/OppgavePage";
import { msalInstance, isMsalConfigured } from "./auth/msalConfig";

function TopBar() {
  if (!isMsalConfigured) return null;
  const account = msalInstance.getAllAccounts()[0];
  const name = account?.name ?? account?.username ?? "";

  return (
    <div style={{
      position: "sticky", top: 0, zIndex: 100,
      background: "var(--bfc-base-2)",
      borderBottom: "1px solid var(--bfc-base-dimmed)",
      display: "flex", alignItems: "center", justifyContent: "space-between",
      padding: "0 1.5rem", height: 44, flexShrink: 0,
    }}>
      <span style={{ fontWeight: 700, fontSize: "0.88rem", color: "var(--bfc-base-c-1)", letterSpacing: "0.01em" }}>
        ProjectTools
      </span>
      <div style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}>
        <span style={{ fontSize: "0.82rem", color: "var(--bfc-base-c-2)" }}>{name}</span>
        <button
          onClick={() => msalInstance.logoutRedirect({ account })}
          style={{
            fontSize: "0.78rem", color: "var(--bfc-base-c-2)",
            background: "none", border: "1px solid var(--bfc-base-dimmed)",
            borderRadius: 4, cursor: "pointer", padding: "3px 10px",
          }}
        >
          Logg ut
        </button>
      </div>
    </div>
  );
}

export default function App() {
  return (
    <>
      <TopBar />
      <Routes>
        <Route path="/" element={<Navigate to="/projects" replace />} />
        <Route path="/projects" element={<ProjectsPage />} />
        <Route path="/projects/:projectId" element={<ProjectDetailPage />} />
        <Route path="/projects/:projectId/risk-matrix/:matrixId" element={<RiskMatrixPage />} />
        <Route path="/projects/:projectId/communication-plan/:planId" element={<CommunicationPlanPage />} />
        <Route path="/projects/:projectId/meeting-plan/:planId" element={<MeetingPlanPage />} />
        <Route path="/projects/:projectId/runbook/:runbookId" element={<RunbookPage />} />
        <Route path="/projects/:projectId/project-plan/:planId" element={<ProjectPlanPage />} />
        <Route path="/projects/:projectId/oppgave/:listeId" element={<OppgavePage />} />
      </Routes>
    </>
  );
}
