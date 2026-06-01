import { Routes, Route, Navigate } from "react-router-dom";
import ProjectsPage from "./pages/ProjectsPage";
import ProjectDetailPage from "./pages/ProjectDetailPage";
import RiskMatrixPage from "./pages/RiskMatrixPage";
import CommunicationPlanPage from "./pages/CommunicationPlanPage";
import MeetingPlanPage from "./pages/MeetingPlanPage";
import RunbookPage from "./pages/RunbookPage";
import ProjectPlanPage from "./pages/ProjectPlanPage";
import OppgavePage from "./pages/OppgavePage";

export default function App() {
  return (
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
  );
}
