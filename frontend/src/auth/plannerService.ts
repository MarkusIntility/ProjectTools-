import type { AccountInfo, IPublicClientApplication } from "@azure/msal-browser";
import { PLANNER_SCOPES } from "./msalConfig";

export interface PlannerBucket {
  id: string;
  name: string;
  orderHint: string;
}

export interface PlannerTask {
  id: string;
  title: string;
  bucketId: string;
  percentComplete: number;
  startDateTime: string | null;
  dueDateTime: string | null;
  assignments: Record<string, unknown>;
  priority: number;
}

export interface PlannerData {
  buckets: PlannerBucket[];
  tasks: PlannerTask[];
}

async function graphRequest<T>(token: string, path: string): Promise<T> {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Graph API feil (${res.status}): ${err}`);
  }
  return res.json();
}

export async function fetchPlannerData(
  msal: IPublicClientApplication,
  account: AccountInfo,
  planId: string
): Promise<PlannerData> {
  let tokenResponse;
  try {
    tokenResponse = await msal.acquireTokenSilent({ scopes: PLANNER_SCOPES, account });
  } catch {
    tokenResponse = await msal.acquireTokenPopup({ scopes: PLANNER_SCOPES });
  }

  const token = tokenResponse.accessToken;
  const [bucketsData, tasksData] = await Promise.all([
    graphRequest<{ value: PlannerBucket[] }>(token, `/planner/plans/${planId}/buckets`),
    graphRequest<{ value: PlannerTask[] }>(token, `/planner/plans/${planId}/tasks`),
  ]);

  return { buckets: bucketsData.value, tasks: tasksData.value };
}

export function parsePlanId(url: string): string | null {
  try {
    // Hash-based: .../#/plantaskboard?groupId=xxx&planId=yyy
    const hashIndex = url.indexOf("#");
    if (hashIndex !== -1) {
      const qIndex = url.indexOf("?", hashIndex);
      if (qIndex !== -1) {
        const params = new URLSearchParams(url.slice(qIndex + 1));
        const id = params.get("planId");
        if (id) return id;
      }
    }
    // Regular query param: ?planId=yyy
    const u = new URL(url);
    const id = u.searchParams.get("planId");
    if (id) return id;
    // Premium Planner: /webui/premiumplan/{uuid}/
    const premiumMatch = url.match(/\/premiumplan\/([0-9a-f-]{36})/i);
    if (premiumMatch) return premiumMatch[1];
    // Classic new Planner app: /plan/[planId]/
    const match = url.match(/\/plan\/([A-Za-z0-9_-]+)/);
    return match?.[1] ?? null;
  } catch {
    return null;
  }
}

export function taskStatus(pct: number): "not_started" | "in_progress" | "done" {
  if (pct === 100) return "done";
  if (pct > 0) return "in_progress";
  return "not_started";
}
