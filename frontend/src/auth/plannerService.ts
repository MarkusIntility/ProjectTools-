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

async function graphRequest<T>(token: string, path: string, apiVersion = "v1.0"): Promise<T> {
  const res = await fetch(`https://graph.microsoft.com/${apiVersion}${path}`, {
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
    // Silent token acquisition failed (consent needed or token expired).
    // Use redirect — popup is unreliable in our SPA setup.
    sessionStorage.setItem("app.returnUrl", window.location.href);
    await msal.acquireTokenRedirect({ scopes: PLANNER_SCOPES, account });
    // Never reached — page navigates away
    throw new Error("Omdirigerer til innlogging…");
  }

  const token = tokenResponse.accessToken;

  // Try v1.0 first, fall back to beta for Planner Premium plans
  const apiVersion = await (async () => {
    const check = await fetch(`https://graph.microsoft.com/v1.0/planner/plans/${planId}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (check.ok) return "v1.0";
    if (check.status === 404) {
      // Try beta — Microsoft is rolling out Premium plan support there
      const betaCheck = await fetch(`https://graph.microsoft.com/beta/planner/plans/${planId}`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (betaCheck.ok) return "beta";
      // Neither version found the plan
      throw new Error(
        "PREMIUM_PLAN: Planen ble ikke funnet verken i Graph v1.0 eller beta. " +
        "Planner Premium-planer lagret i Dataverse er ikke tilgjengelige via " +
        "standard Graph Planner API. Bruk «Åpne i Planner»-lenken for å se " +
        "planen direkte, eller opprett en Basic Planner-plan for API-integrasjon."
      );
    }
    throw new Error(`Graph API feil (${check.status}): ${await check.text()}`);
  })();

  const [bucketsData, tasksData] = await Promise.all([
    graphRequest<{ value: PlannerBucket[] }>(token, `/planner/plans/${planId}/buckets`, apiVersion),
    graphRequest<{ value: PlannerTask[] }>(token, `/planner/plans/${planId}/tasks`, apiVersion),
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
