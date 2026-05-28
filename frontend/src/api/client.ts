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
}

export interface RiskMatrix {
  id: string;
  project_id: string;
  title: string;
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
  location: string | null;
  agenda: string | null;
  participants: string | null;
  minutes: string | null;
}

export interface MeetingPlan {
  id: string;
  project_id: string;
  title: string;
  created_at: string;
  meetings: Meeting[];
}

export const api = {
  projects: {
    list: () => request<Project[]>("/projects/"),
    get: (id: string) => request<Project>(`/projects/${id}`),
    create: (data: { name: string; description?: string }) =>
      request<Project>("/projects/", { method: "POST", body: JSON.stringify(data) }),
    update: (id: string, data: { name: string; description?: string }) =>
      request<Project>(`/projects/${id}`, { method: "PUT", body: JSON.stringify(data) }),
    delete: (id: string) => request<void>(`/projects/${id}`, { method: "DELETE" }),
  },
  riskMatrices: {
    list: (projectId: string) => request<RiskMatrix[]>(`/projects/${projectId}/risk-matrices/`),
    get: (projectId: string, matrixId: string) => request<RiskMatrix>(`/projects/${projectId}/risk-matrices/${matrixId}`),
    create: (projectId: string, data: { title: string }) =>
      request<RiskMatrix>(`/projects/${projectId}/risk-matrices/`, { method: "POST", body: JSON.stringify(data) }),
    delete: (projectId: string, matrixId: string) =>
      request<void>(`/projects/${projectId}/risk-matrices/${matrixId}`, { method: "DELETE" }),
    addRisk: (projectId: string, matrixId: string, data: Omit<RiskItem, "id" | "matrix_id" | "risk_score">) =>
      request<RiskItem>(`/projects/${projectId}/risk-matrices/${matrixId}/risks`, { method: "POST", body: JSON.stringify(data) }),
    updateRisk: (projectId: string, matrixId: string, riskId: string, data: Omit<RiskItem, "id" | "matrix_id" | "risk_score">) =>
      request<RiskItem>(`/projects/${projectId}/risk-matrices/${matrixId}/risks/${riskId}`, { method: "PUT", body: JSON.stringify(data) }),
    deleteRisk: (projectId: string, matrixId: string, riskId: string) =>
      request<void>(`/projects/${projectId}/risk-matrices/${matrixId}/risks/${riskId}`, { method: "DELETE" }),
  },
  communicationPlans: {
    list: (projectId: string) => request<CommunicationPlan[]>(`/projects/${projectId}/communication-plans/`),
    get: (projectId: string, planId: string) => request<CommunicationPlan>(`/projects/${projectId}/communication-plans/${planId}`),
    create: (projectId: string, data: { title: string }) =>
      request<CommunicationPlan>(`/projects/${projectId}/communication-plans/`, { method: "POST", body: JSON.stringify(data) }),
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
    delete: (projectId: string, planId: string) =>
      request<void>(`/projects/${projectId}/meeting-plans/${planId}`, { method: "DELETE" }),
    addMeeting: (projectId: string, planId: string, data: Omit<Meeting, "id" | "plan_id">) =>
      request<Meeting>(`/projects/${projectId}/meeting-plans/${planId}/meetings`, { method: "POST", body: JSON.stringify(data) }),
    updateMeeting: (projectId: string, planId: string, meetingId: string, data: Omit<Meeting, "id" | "plan_id">) =>
      request<Meeting>(`/projects/${projectId}/meeting-plans/${planId}/meetings/${meetingId}`, { method: "PUT", body: JSON.stringify(data) }),
    deleteMeeting: (projectId: string, planId: string, meetingId: string) =>
      request<void>(`/projects/${projectId}/meeting-plans/${planId}/meetings/${meetingId}`, { method: "DELETE" }),
  },
};
