const BASE_URL = "/api";

async function request<T>(path: string, options?: RequestInit): Promise<T> {
  const res = await fetch(`${BASE_URL}${path}`, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  if (!res.ok) {
    const error = await res.text();
    throw new Error(error || res.statusText);
  }
  if (res.status === 204) return undefined as T;
  return res.json();
}

export interface Project {
  id: string;
  name: string;
  description: string | null;
  project_manager: string | null;
  status: "active" | "not_started" | "completed";
  created_at: string;
  updated_at: string;
}

export interface RiskItem {
  id: string;
  matrix_id: string;
  description: string;
  probability: number;
  consequence: number;
  risk_score: number;
  mitigation: string | null;
  owner: string | null;
  status: "open" | "mitigated" | "closed";
  fagomrade: string | null;
  risk_owner: string | null;
  residual_probability: number | null;
  residual_consequence: number | null;
  residual_score: number | null;
  fase?: string | null;
}

export interface RiskMatrix {
  id: string;
  project_id: string;
  title: string;
  is_primary: boolean;
  created_at: string;
  risks: RiskItem[];
}

export interface CommunicationEntry {
  id: string;
  plan_id: string;
  stakeholder: string;
  message: string;
  channel: string;
  frequency: string;
  responsible: string;
}

export interface CommunicationPlan {
  id: string;
  project_id: string;
  title: string;
  created_at: string;
  entries: CommunicationEntry[];
}

export interface Meeting {
  id: string;
  plan_id: string;
  title: string;
  date: string;
  purpose: string | null;
  outlook_id: string | null;
}

export interface MeetingPlan {
  id: string;
  project_id: string;
  title: string;
  is_primary: boolean;
  created_at: string;
  meetings: Meeting[];
}

export interface RunbookActivity {
  id: string;
  runbook_id: string;
  name: string;
  phase: string | null;
  status: "not_started" | "in_progress" | "done" | "cancelled";
  start_date: string | null;
  end_date: string | null;
  responsible: string | null;
  description: string | null;
  sort_order: number;
}

export interface ProjectPlanTask {
  id: string;
  plan_id: string;
  name: string;
  bucket: string | null;
  percent_complete: number;
  start_date: string | null;
  end_date: string | null;
  responsible: string | null;
  description: string | null;
  sort_order: number;
}

export interface ProjectPlan {
  id: string;
  project_id: string;
  title: string;
  source: "own" | "planner" | "smartsheet";
  external_url: string | null;
  is_primary: boolean;
  created_at: string;
  tasks: ProjectPlanTask[];
}

export interface Oppgave {
  id: string;
  liste_id: string;
  name: string;
  responsible: string | null;
  due_date: string | null;
  status: "not_started" | "in_progress" | "done";
  description: string | null;
  sort_order: number;
}

export interface OppgaveListe {
  id: string;
  project_id: string;
  title: string;
  source: "own" | "planner" | "smartsheet";
  external_url: string | null;
  is_primary: boolean;
  created_at: string;
  oppgaver: Oppgave[];
}

export interface Runbook {
  id: string;
  project_id: string;
  title: string;
  source: "own" | "planner" | "smartsheet";
  external_url: string | null;
  created_at: string;
  activities: RunbookActivity[];
}

export interface Template {
  id: string;
  name: string;
  type: string;
  data: string; // JSON string
  created_at: string;
  updated_at: string;
}

export const api = {
  projects: {
    list: () => request<Project[]>("/projects/"),
    get: (id: string) => request<Project>(`/projects/${id}`),
    create: (data: { name: string; description?: string; project_manager?: string | null; status?: string }) =>
      request<Project>("/projects/", { method: "POST", body: JSON.stringify(data) }),
    update: (id: string, data: { name: string; description?: string; project_manager?: string | null; status?: string }) =>
      request<Project>(`/projects/${id}`, { method: "PUT", body: JSON.stringify(data) }),
    delete: (id: string) => request<void>(`/projects/${id}`, { method: "DELETE" }),
  },
  riskMatrices: {
    list: (projectId: string) => request<RiskMatrix[]>(`/projects/${projectId}/risk-matrices/`),
    get: (projectId: string, matrixId: string) => request<RiskMatrix>(`/projects/${projectId}/risk-matrices/${matrixId}`),
    create: (projectId: string, data: { title: string }) =>
      request<RiskMatrix>(`/projects/${projectId}/risk-matrices/`, { method: "POST", body: JSON.stringify(data) }),
    update: (projectId: string, matrixId: string, data: { title: string }) =>
      request<RiskMatrix>(`/projects/${projectId}/risk-matrices/${matrixId}`, { method: "PUT", body: JSON.stringify(data) }),
    delete: (projectId: string, matrixId: string) =>
      request<void>(`/projects/${projectId}/risk-matrices/${matrixId}`, { method: "DELETE" }),
    addRisk: (projectId: string, matrixId: string, data: Omit<RiskItem, "id" | "matrix_id" | "risk_score" | "residual_score">) =>
      request<RiskItem>(`/projects/${projectId}/risk-matrices/${matrixId}/risks`, { method: "POST", body: JSON.stringify(data) }),
    updateRisk: (projectId: string, matrixId: string, riskId: string, data: Omit<RiskItem, "id" | "matrix_id" | "risk_score" | "residual_score">) =>
      request<RiskItem>(`/projects/${projectId}/risk-matrices/${matrixId}/risks/${riskId}`, { method: "PUT", body: JSON.stringify(data) }),
    deleteRisk: (projectId: string, matrixId: string, riskId: string) =>
      request<void>(`/projects/${projectId}/risk-matrices/${matrixId}/risks/${riskId}`, { method: "DELETE" }),
    setPrimary: (projectId: string, matrixId: string) =>
      request<RiskMatrix>(`/projects/${projectId}/risk-matrices/${matrixId}/set-primary`, { method: "PUT" }),
  },
  communicationPlans: {
    list: (projectId: string) => request<CommunicationPlan[]>(`/projects/${projectId}/communication-plans/`),
    get: (projectId: string, planId: string) => request<CommunicationPlan>(`/projects/${projectId}/communication-plans/${planId}`),
    create: (projectId: string, data: { title: string }) =>
      request<CommunicationPlan>(`/projects/${projectId}/communication-plans/`, { method: "POST", body: JSON.stringify(data) }),
    update: (projectId: string, planId: string, data: { title: string }) =>
      request<CommunicationPlan>(`/projects/${projectId}/communication-plans/${planId}`, { method: "PUT", body: JSON.stringify(data) }),
    delete: (projectId: string, planId: string) =>
      request<void>(`/projects/${projectId}/communication-plans/${planId}`, { method: "DELETE" }),
    addEntry: (projectId: string, planId: string, data: Omit<CommunicationEntry, "id" | "plan_id">) =>
      request<CommunicationEntry>(`/projects/${projectId}/communication-plans/${planId}/entries`, { method: "POST", body: JSON.stringify(data) }),
    updateEntry: (projectId: string, planId: string, entryId: string, data: Omit<CommunicationEntry, "id" | "plan_id">) =>
      request<CommunicationEntry>(`/projects/${projectId}/communication-plans/${planId}/entries/${entryId}`, { method: "PUT", body: JSON.stringify(data) }),
    deleteEntry: (projectId: string, planId: string, entryId: string) =>
      request<void>(`/projects/${projectId}/communication-plans/${planId}/entries/${entryId}`, { method: "DELETE" }),
  },
  meetingPlans: {
    list: (projectId: string) => request<MeetingPlan[]>(`/projects/${projectId}/meeting-plans/`),
    get: (projectId: string, planId: string) => request<MeetingPlan>(`/projects/${projectId}/meeting-plans/${planId}`),
    create: (projectId: string, data: { title: string }) =>
      request<MeetingPlan>(`/projects/${projectId}/meeting-plans/`, { method: "POST", body: JSON.stringify(data) }),
    update: (projectId: string, planId: string, data: { title: string }) =>
      request<MeetingPlan>(`/projects/${projectId}/meeting-plans/${planId}`, { method: "PUT", body: JSON.stringify(data) }),
    delete: (projectId: string, planId: string) =>
      request<void>(`/projects/${projectId}/meeting-plans/${planId}`, { method: "DELETE" }),
    addMeeting: (projectId: string, planId: string, data: Omit<Meeting, "id" | "plan_id">) =>
      request<Meeting>(`/projects/${projectId}/meeting-plans/${planId}/meetings`, { method: "POST", body: JSON.stringify(data) }),
    updateMeeting: (projectId: string, planId: string, meetingId: string, data: Omit<Meeting, "id" | "plan_id">) =>
      request<Meeting>(`/projects/${projectId}/meeting-plans/${planId}/meetings/${meetingId}`, { method: "PUT", body: JSON.stringify(data) }),
    deleteMeeting: (projectId: string, planId: string, meetingId: string) =>
      request<void>(`/projects/${projectId}/meeting-plans/${planId}/meetings/${meetingId}`, { method: "DELETE" }),
    setPrimary: (projectId: string, planId: string) =>
      request<MeetingPlan>(`/projects/${projectId}/meeting-plans/${planId}/set-primary`, { method: "PUT" }),
  },
  projectPlans: {
    list: (projectId: string) => request<ProjectPlan[]>(`/projects/${projectId}/project-plans/`),
    get: (projectId: string, planId: string) => request<ProjectPlan>(`/projects/${projectId}/project-plans/${planId}`),
    create: (projectId: string, data: { title: string; source: string; external_url?: string }) =>
      request<ProjectPlan>(`/projects/${projectId}/project-plans/`, { method: "POST", body: JSON.stringify(data) }),
    update: (projectId: string, planId: string, data: { title: string; external_url?: string | null }) =>
      request<ProjectPlan>(`/projects/${projectId}/project-plans/${planId}`, { method: "PUT", body: JSON.stringify(data) }),
    delete: (projectId: string, planId: string) =>
      request<void>(`/projects/${projectId}/project-plans/${planId}`, { method: "DELETE" }),
    addTask: (projectId: string, planId: string, data: Omit<ProjectPlanTask, "id" | "plan_id" | "sort_order">) =>
      request<ProjectPlanTask>(`/projects/${projectId}/project-plans/${planId}/tasks`, { method: "POST", body: JSON.stringify(data) }),
    updateTask: (projectId: string, planId: string, taskId: string, data: Omit<ProjectPlanTask, "id" | "plan_id">) =>
      request<ProjectPlanTask>(`/projects/${projectId}/project-plans/${planId}/tasks/${taskId}`, { method: "PUT", body: JSON.stringify(data) }),
    deleteTask: (projectId: string, planId: string, taskId: string) =>
      request<void>(`/projects/${projectId}/project-plans/${planId}/tasks/${taskId}`, { method: "DELETE" }),
    setPrimary: (projectId: string, planId: string) =>
      request<ProjectPlan>(`/projects/${projectId}/project-plans/${planId}/set-primary`, { method: "PUT" }),
  },
  oppgaveLister: {
    list: (projectId: string) => request<OppgaveListe[]>(`/projects/${projectId}/oppgave-lister/`),
    get: (projectId: string, listeId: string) => request<OppgaveListe>(`/projects/${projectId}/oppgave-lister/${listeId}`),
    create: (projectId: string, data: { title: string; source: string; external_url?: string }) =>
      request<OppgaveListe>(`/projects/${projectId}/oppgave-lister/`, { method: "POST", body: JSON.stringify(data) }),
    update: (projectId: string, listeId: string, data: { title: string; external_url?: string | null }) =>
      request<OppgaveListe>(`/projects/${projectId}/oppgave-lister/${listeId}`, { method: "PUT", body: JSON.stringify(data) }),
    delete: (projectId: string, listeId: string) =>
      request<void>(`/projects/${projectId}/oppgave-lister/${listeId}`, { method: "DELETE" }),
    addOppgave: (projectId: string, listeId: string, data: Omit<Oppgave, "id" | "liste_id" | "sort_order">) =>
      request<Oppgave>(`/projects/${projectId}/oppgave-lister/${listeId}/oppgaver`, { method: "POST", body: JSON.stringify(data) }),
    updateOppgave: (projectId: string, listeId: string, oppgaveId: string, data: Omit<Oppgave, "id" | "liste_id">) =>
      request<Oppgave>(`/projects/${projectId}/oppgave-lister/${listeId}/oppgaver/${oppgaveId}`, { method: "PUT", body: JSON.stringify(data) }),
    deleteOppgave: (projectId: string, listeId: string, oppgaveId: string) =>
      request<void>(`/projects/${projectId}/oppgave-lister/${listeId}/oppgaver/${oppgaveId}`, { method: "DELETE" }),
    setPrimary: (projectId: string, listeId: string) =>
      request<OppgaveListe>(`/projects/${projectId}/oppgave-lister/${listeId}/set-primary`, { method: "PUT" }),
  },
  templates: {
    list: (type?: string) => request<Template[]>(`/templates/${type ? `?type=${type}` : ""}`),
    get: (id: string) => request<Template>(`/templates/${id}`),
    create: (data: { name: string; type: string; data: string }) =>
      request<Template>("/templates/", { method: "POST", body: JSON.stringify(data) }),
    update: (id: string, data: { name: string; data: string }) =>
      request<Template>(`/templates/${id}`, { method: "PUT", body: JSON.stringify(data) }),
    delete: (id: string) => request<void>(`/templates/${id}`, { method: "DELETE" }),
  },
  runbooks: {
    list: (projectId: string) => request<Runbook[]>(`/projects/${projectId}/runbooks/`),
    get: (projectId: string, runbookId: string) => request<Runbook>(`/projects/${projectId}/runbooks/${runbookId}`),
    create: (projectId: string, data: { title: string; source: string; external_url?: string }) =>
      request<Runbook>(`/projects/${projectId}/runbooks/`, { method: "POST", body: JSON.stringify(data) }),
    update: (projectId: string, runbookId: string, data: { title: string; external_url?: string | null }) =>
      request<Runbook>(`/projects/${projectId}/runbooks/${runbookId}`, { method: "PUT", body: JSON.stringify(data) }),
    delete: (projectId: string, runbookId: string) =>
      request<void>(`/projects/${projectId}/runbooks/${runbookId}`, { method: "DELETE" }),
    addActivity: (projectId: string, runbookId: string, data: Omit<RunbookActivity, "id" | "runbook_id" | "sort_order">) =>
      request<RunbookActivity>(`/projects/${projectId}/runbooks/${runbookId}/activities`, { method: "POST", body: JSON.stringify(data) }),
    updateActivity: (projectId: string, runbookId: string, activityId: string, data: Omit<RunbookActivity, "id" | "runbook_id">) =>
      request<RunbookActivity>(`/projects/${projectId}/runbooks/${runbookId}/activities/${activityId}`, { method: "PUT", body: JSON.stringify(data) }),
    deleteActivity: (projectId: string, runbookId: string, activityId: string) =>
      request<void>(`/projects/${projectId}/runbooks/${runbookId}/activities/${activityId}`, { method: "DELETE" }),
  },
};
