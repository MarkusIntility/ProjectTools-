import type { AccountInfo, IPublicClientApplication } from "@azure/msal-browser";
import { PLANNER_SCOPES } from "./msalConfig";

// ─── Types ────────────────────────────────────────────────────────────────────

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

// ─── Dataverse types ──────────────────────────────────────────────────────────

interface DataverseInstance {
  ApiUrl: string; // "https://org.api.crm4.dynamics.com/"
  Url: string;    // "https://org.crm4.dynamics.com/"
  FriendlyName: string;
}

interface DataverseTask {
  msdyn_projecttaskid: string;
  msdyn_subject: string;
  msdyn_progress?: number;        // 0.0–1.0 in some versions
  msdyn_percentcomplete?: number; // 0–100 in other versions
  msdyn_scheduledstart: string | null;
  msdyn_scheduledend: string | null;
  "_msdyn_projectbucket_value"?: string | null;
}

interface DataverseBucket {
  msdyn_projectbucketid: string;
  msdyn_name: string;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

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

async function acquireToken(
  msal: IPublicClientApplication,
  account: AccountInfo,
  scopes: string[]
): Promise<string> {
  try {
    const r = await msal.acquireTokenSilent({ scopes, account });
    return r.accessToken;
  } catch {
    // Silent failed — redirect for consent/re-auth
    sessionStorage.setItem("app.returnUrl", window.location.href);
    await msal.acquireTokenRedirect({ scopes, account });
    throw new Error("Omdirigerer for tilgang…"); // never reached
  }
}

// ─── Dataverse / Planner Premium ─────────────────────────────────────────────

const DATAVERSE_GD_URL = "https://globaldisco.crm.dynamics.com/api/discovery/v1.0/Instances";
const DATAVERSE_GD_SCOPE = "https://globaldisco.crm.dynamics.com/.default";

async function fetchPlannerPremiumData(
  msal: IPublicClientApplication,
  account: AccountInfo,
  planId: string
): Promise<PlannerData> {
  // 1. Discover the user's Dataverse environment via Global Discovery Service
  const gdToken = await acquireToken(msal, account, [DATAVERSE_GD_SCOPE]);

  const gdRes = await fetch(DATAVERSE_GD_URL, {
    headers: { Authorization: `Bearer ${gdToken}` },
  });
  if (!gdRes.ok) {
    throw new Error(`Fant ikke Dataverse-miljø (${gdRes.status}). Sjekk at du har Dynamics 365-tilgang.`);
  }

  const { value: instances }: { value: DataverseInstance[] } = await gdRes.json();
  if (!instances?.length) {
    throw new Error("Ingen Dataverse-miljøer funnet for denne kontoen. Planner Premium krever Power Platform-lisens.");
  }

  // Use first environment (most orgs have one; multi-env support can be added later)
  const env = instances[0];
  console.log("[Planner] Dataverse miljø:", env.FriendlyName, "ApiUrl:", env.ApiUrl, "Url:", env.Url);
  const apiUrl = env.ApiUrl.replace(/\/$/, "");

  // 2. Get token scoped to this specific Dataverse environment
  const dvScope = `${env.Url.replace(/\/$/, "")}/.default`;
  const dvToken = await acquireToken(msal, account, [dvScope]);

  const dvHeaders: Record<string, string> = {
    Authorization: `Bearer ${dvToken}`,
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    Accept: "application/json",
  };

  // 3. Fetch project tasks filtered by plan ID
  const taskFields = [
    "msdyn_projecttaskid",
    "msdyn_subject",
    "msdyn_progress",
    "msdyn_percentcomplete",
    "msdyn_scheduledstart",
    "msdyn_scheduledend",
    "_msdyn_projectbucket_value",
  ].join(",");

  const tasksUrl = `${apiUrl}/api/data/v9.2/msdyn_projecttasks?$filter=_msdyn_project_value eq ${planId}&$select=${taskFields}&$orderby=msdyn_scheduledstart asc`;
  console.log("[Planner] Dataverse tasks URL:", tasksUrl);
  const tasksRes = await fetch(tasksUrl, { headers: dvHeaders });

  if (!tasksRes.ok) {
    const errText = await tasksRes.text();
    console.error("[Planner] Dataverse tasks feil:", tasksRes.status, errText);
    if (tasksRes.status === 403) {
      throw new Error(`PREMIUM_PLAN: Ingen tilgang (403) til ${env.FriendlyName}. Du mangler lesetilgang til prosjektet i Dataverse.`);
    }
    if (tasksRes.status === 404) {
      throw new Error(`PREMIUM_PLAN: msdyn_projecttasks finnes ikke (404) på ${env.FriendlyName} (${apiUrl}). Svar: ${errText.slice(0, 300)}`);
    }
    throw new Error(`PREMIUM_PLAN: Dataverse feil (${tasksRes.status}) på ${env.FriendlyName}: ${errText.slice(0, 300)}`);
  }

  const { value: dvTasks }: { value: DataverseTask[] } = await tasksRes.json();

  // 4. Fetch buckets for this project
  let dvBuckets: DataverseBucket[] = [];
  const bucketsRes = await fetch(
    `${apiUrl}/api/data/v9.2/msdyn_projectbuckets?$filter=_msdyn_project_value eq ${planId}&$select=msdyn_projectbucketid,msdyn_name&$orderby=msdyn_name asc`,
    { headers: dvHeaders }
  );
  if (bucketsRes.ok) {
    const { value } = await bucketsRes.json();
    dvBuckets = value || [];
  }

  // 5. Map to PlannerData format
  const buckets: PlannerBucket[] =
    dvBuckets.length > 0
      ? dvBuckets.map((b) => ({
          id: b.msdyn_projectbucketid,
          name: b.msdyn_name,
          orderHint: b.msdyn_projectbucketid,
        }))
      : [{ id: "default", name: "Oppgaver", orderHint: "default" }];

  const defaultBucketId = buckets[0].id;

  const tasks: PlannerTask[] = dvTasks.map((t) => {
    const pct =
      t.msdyn_percentcomplete !== undefined
        ? t.msdyn_percentcomplete
        : Math.round((t.msdyn_progress ?? 0) * 100);
    return {
      id: t.msdyn_projecttaskid,
      title: t.msdyn_subject,
      bucketId: t["_msdyn_projectbucket_value"] ?? defaultBucketId,
      percentComplete: pct,
      startDateTime: t.msdyn_scheduledstart,
      dueDateTime: t.msdyn_scheduledend,
      assignments: {},
      priority: 5,
    };
  });

  return { buckets, tasks };
}

// ─── Main export ──────────────────────────────────────────────────────────────

export async function fetchPlannerData(
  msal: IPublicClientApplication,
  account: AccountInfo,
  planId: string
): Promise<PlannerData> {
  const token = await acquireToken(msal, account, PLANNER_SCOPES);

  // Check if plan exists in standard Graph API (v1.0 then beta)
  const v1Check = await fetch(`https://graph.microsoft.com/v1.0/planner/plans/${planId}`, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (v1Check.ok) {
    // Standard Basic Planner plan — use Graph API
    const [bucketsData, tasksData] = await Promise.all([
      graphRequest<{ value: PlannerBucket[] }>(token, `/planner/plans/${planId}/buckets`),
      graphRequest<{ value: PlannerTask[] }>(token, `/planner/plans/${planId}/tasks`),
    ]);
    return { buckets: bucketsData.value, tasks: tasksData.value };
  }

  if (v1Check.status === 404) {
    // Try beta (Microsoft is rolling out Premium support there)
    const betaCheck = await fetch(`https://graph.microsoft.com/beta/planner/plans/${planId}`, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (betaCheck.ok) {
      const [bucketsData, tasksData] = await Promise.all([
        graphRequest<{ value: PlannerBucket[] }>(token, `/planner/plans/${planId}/buckets`, "beta"),
        graphRequest<{ value: PlannerTask[] }>(token, `/planner/plans/${planId}/tasks`, "beta"),
      ]);
      return { buckets: bucketsData.value, tasks: tasksData.value };
    }

    // Both Graph versions failed — fall back to Dataverse for Planner Premium
    return await fetchPlannerPremiumData(msal, account, planId);
  }

  throw new Error(`Graph API feil (${v1Check.status}): ${await v1Check.text()}`);
}

// ─── Utilities ────────────────────────────────────────────────────────────────

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
