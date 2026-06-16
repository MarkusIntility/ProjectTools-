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
  labels?: string[];
  outlineLevel?: number;   // 1 = fase, 2 = leveranse, 3 = lokasjon/pulje
  parentTaskId?: string | null;
  durationMinutes?: number | null; // 0 = explicit milestone in Planner Premium; null = not fetched (Basic Planner)
  effort?: number | null;          // msdyn_effort (hours); null = not fetched (Basic Planner)
}

export interface PlannerData {
  buckets: PlannerBucket[];
  tasks: PlannerTask[];
  categoryDescriptions?: Record<string, string | null>;
  assigneeMap?: Record<string, string>; // userId → displayName
  source: "basic" | "premium";
  premiumApiUrl?: string;  // set when source === "premium"
  premiumDvScope?: string; // Dataverse OAuth scope, set when source === "premium"
}

// ─── Dataverse types ──────────────────────────────────────────────────────────

interface DataverseInstance {
  ApiUrl: string;         // "https://org.api.crm4.dynamics.com/"
  Url: string;            // "https://org.crm4.dynamics.com/"
  FriendlyName: string;
  Id?: string;            // GUID — Global Discovery returns "Id", not "OrganizationId"
}

interface DataverseTask {
  msdyn_projecttaskid: string;
  msdyn_subject: string;
  msdyn_progress?: number; // 0.0–1.0
  msdyn_scheduledstart: string | null;
  msdyn_scheduledend: string | null;
  "_msdyn_projectbucket_value"?: string | null;
  "_msdyn_resourcecategory_value"?: string | null;
  "_msdyn_projectsprint_value"?: string | null;
  "_msdyn_parenttask_value"?: string | null;
  msdyn_outlinelevel?: number;
  msdyn_duration?: number | null; // in minutes; 0 = milestone
  msdyn_effort?: number | null;   // planned effort in hours
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

async function resolveAssignees(
  token: string,
  tasks: (PlannerTask & { appliedCategories?: Record<string, boolean> })[]
): Promise<Record<string, string>> {
  const ids = [...new Set(tasks.flatMap((t) => Object.keys(t.assignments ?? {})))];
  if (ids.length === 0) return {};
  const map: Record<string, string> = {};
  await Promise.all(
    ids.map(async (id) => {
      try {
        const user = await graphRequest<{ displayName: string }>(token, `/users/${id}?$select=displayName`);
        map[id] = user.displayName;
      } catch {
        // leave entry absent — caller falls back to showing nothing
      }
    })
  );
  return map;
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
  planId: string,
  orgId: string | null
): Promise<PlannerData> {
  // 1. Discover Dataverse environments via Global Discovery Service
  const gdToken = await acquireToken(msal, account, [DATAVERSE_GD_SCOPE]);

  const gdRes = await fetch(DATAVERSE_GD_URL, {
    headers: { Authorization: `Bearer ${gdToken}` },
  });
  if (!gdRes.ok) {
    throw new Error(`PREMIUM_PLAN: Fant ikke Dataverse-miljø (${gdRes.status}). Sjekk at du har Dynamics 365-tilgang.`);
  }

  const { value: instances }: { value: DataverseInstance[] } = await gdRes.json();
  if (!instances?.length) {
    throw new Error("PREMIUM_PLAN: Ingen Dataverse-miljøer funnet for denne kontoen. Planner Premium krever Power Platform-lisens.");
  }

  console.log("[Planner] Dataverse-miljøer:", instances.map(i => `${i.FriendlyName} (${i.Id})`).join(", "));

  // 2. Find the environment matching the org ID from the Planner Premium URL
  let env: DataverseInstance | undefined;
  if (orgId) {
    env = instances.find(i => i.Id?.toLowerCase() === orgId.toLowerCase());
    if (env) {
      console.log("[Planner] Fant riktig miljø via orgId:", env.FriendlyName);
    } else {
      console.warn(`[Planner] orgId ${orgId} ikke funnet. Prøver alle miljøer.`);
    }
  }

  // 3. Try environments: matched first, then all others as fallback
  const ordered = env
    ? [env, ...instances.filter(i => i !== env)]
    : instances;

  for (const candidate of ordered) {
    const apiUrl = candidate.ApiUrl.replace(/\/$/, "");
    const dvScope = `${candidate.Url.replace(/\/$/, "")}/.default`;

    let dvToken: string;
    try {
      dvToken = await acquireToken(msal, account, [dvScope]);
    } catch {
      continue; // can't get token for this env, try next
    }

    const dvHeaders: Record<string, string> = {
      Authorization: `Bearer ${dvToken}`,
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      Accept: "application/json",
    };

    const taskFields = [
      "msdyn_projecttaskid",
      "msdyn_subject",
      "msdyn_progress",
      "msdyn_scheduledstart",
      "msdyn_scheduledend",
      "msdyn_duration",
      "msdyn_effort",
      "_msdyn_projectbucket_value",
      "_msdyn_resourcecategory_value",
      "_msdyn_projectsprint_value",
      "_msdyn_parenttask_value",
      "msdyn_outlinelevel",
    ].join(",");

    const tasksUrl = `${apiUrl}/api/data/v9.2/msdyn_projecttasks?$filter=_msdyn_project_value eq ${planId}&$select=${taskFields}&$orderby=msdyn_scheduledstart asc`;
    console.log("[Planner] Prøver:", candidate.FriendlyName, tasksUrl);

    const tasksRes = await fetch(tasksUrl, { headers: dvHeaders });

    if (tasksRes.status === 404) {
      // Entity doesn't exist in this env — try next
      console.warn(`[Planner] msdyn_projecttasks finnes ikke i ${candidate.FriendlyName}, prøver neste.`);
      continue;
    }

    if (!tasksRes.ok) {
      const errText = await tasksRes.text();
      console.error("[Planner] Dataverse feil:", tasksRes.status, errText);
      if (tasksRes.status === 403) {
        throw new Error(`PREMIUM_PLAN: Ingen tilgang (403) til prosjektet i ${candidate.FriendlyName}.`);
      }
      throw new Error(`PREMIUM_PLAN: Dataverse feil (${tasksRes.status}) i ${candidate.FriendlyName}: ${errText.slice(0, 200)}`);
    }

    const { value: dvTasks }: { value: DataverseTask[] } = await tasksRes.json();
    console.log(`[Planner] Fant ${dvTasks.length} oppgaver i ${candidate.FriendlyName}`);

    // Fetch buckets and try multiple label sources in parallel
    const [bucketsRes, sprintRes, rcatRes] = await Promise.all([
      fetch(
        `${apiUrl}/api/data/v9.2/msdyn_projectbuckets?$filter=_msdyn_project_value eq ${planId}&$select=msdyn_projectbucketid,msdyn_name&$orderby=msdyn_name asc`,
        { headers: dvHeaders }
      ),
      fetch(
        `${apiUrl}/api/data/v9.2/msdyn_projectsprints?$filter=_msdyn_project_value eq ${planId}&$select=msdyn_projectsprintid,msdyn_name`,
        { headers: dvHeaders }
      ),
      fetch(
        `${apiUrl}/api/data/v9.2/msdyn_resourcecategories?$select=msdyn_resourcecategoryid,msdyn_name`,
        { headers: dvHeaders }
      ),
    ]);

    let dvBuckets: DataverseBucket[] = [];
    if (bucketsRes.ok) {
      const { value } = await bucketsRes.json();
      dvBuckets = value || [];
    }

    // Build label map: try sprints first, then resource categories
    const labelMap: Record<string, string> = {};
    let labelField: "_msdyn_projectsprint_value" | "_msdyn_resourcecategory_value" | null = null;

    if (sprintRes.ok) {
      const { value: sprints } = await sprintRes.json();
      if (sprints?.length > 0) {
        console.log("[Planner Premium] Bruker sprints som labels:", sprints.map((s: Record<string, string>) => s.msdyn_name).join(", "));
        for (const s of sprints) {
          labelMap[s.msdyn_projectsprintid] = s.msdyn_name;
        }
        labelField = "_msdyn_projectsprint_value";
      }
    }

    if (!labelField && rcatRes.ok) {
      const { value: rcats } = await rcatRes.json();
      if (rcats?.length > 0) {
        console.log("[Planner Premium] Bruker resource categories som labels:", rcats.map((c: Record<string, string>) => c.msdyn_name).join(", "));
        for (const c of rcats) {
          labelMap[c.msdyn_resourcecategoryid] = c.msdyn_name;
        }
        labelField = "_msdyn_resourcecategory_value";
      }
    }

    if (!labelField) {
      console.log("[Planner Premium] Ingen label-kilde funnet (sprints:", sprintRes.status, "/ resourcecategories:", rcatRes.status, ")");
    }

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
      const pct = Math.round((t.msdyn_progress ?? 0) * 100);
      const labelId = labelField ? t[labelField] : null;
      const labels = labelId && labelMap[labelId] ? [labelMap[labelId]] : [];
      return {
        id: t.msdyn_projecttaskid,
        title: t.msdyn_subject,
        bucketId: t["_msdyn_projectbucket_value"] ?? defaultBucketId,
        percentComplete: pct,
        startDateTime: t.msdyn_scheduledstart,
        dueDateTime: t.msdyn_scheduledend,
        assignments: {},
        priority: 5,
        labels,
        outlineLevel: t.msdyn_outlinelevel,
        parentTaskId: t["_msdyn_parenttask_value"] ?? null,
        durationMinutes: t.msdyn_duration ?? null,
        effort: t.msdyn_effort ?? null,
      };
    });

    return { buckets, tasks, source: "premium" as const, premiumApiUrl: apiUrl, premiumDvScope: dvScope };
  }

  throw new Error("PREMIUM_PLAN: Planner Premium-planen ble ikke funnet i noen Dataverse-miljøer. Sjekk at du har tilgang til riktig Power Platform-miljø.");
}

// ─── Main export ──────────────────────────────────────────────────────────────

// Accepts the full Planner URL so both planId and orgId can be extracted
export async function fetchPlannerData(
  msal: IPublicClientApplication,
  account: AccountInfo,
  planUrl: string
): Promise<PlannerData> {
  const planId = parsePlanId(planUrl);
  if (!planId) throw new Error("Ugyldig Planner-URL — plan-ID ikke funnet.");

  const token = await acquireToken(msal, account, PLANNER_SCOPES);

  const v1Check = await fetch(`https://graph.microsoft.com/v1.0/planner/plans/${planId}`, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (v1Check.ok) {
    const [bucketsData, tasksData, planDetails] = await Promise.all([
      graphRequest<{ value: PlannerBucket[] }>(token, `/planner/plans/${planId}/buckets`),
      graphRequest<{ value: (PlannerTask & { appliedCategories?: Record<string, boolean> })[] }>(token, `/planner/plans/${planId}/tasks`),
      graphRequest<{ categoryDescriptions: Record<string, string | null> }>(token, `/planner/plans/${planId}/details`)
        .catch(() => ({ categoryDescriptions: {} as Record<string, string | null> })),
    ]);
    const catDesc = planDetails.categoryDescriptions ?? {};
    const tasks = tasksData.value.map((t) => ({
      ...t,
      labels: Object.entries(t.appliedCategories ?? {})
        .filter(([, v]) => v === true)
        .map(([k]) => catDesc[k])
        .filter((name): name is string => Boolean(name)),
    }));
    const assigneeMap = await resolveAssignees(token, tasksData.value);
    return { buckets: bucketsData.value, tasks, categoryDescriptions: catDesc, assigneeMap, source: "basic" as const };
  }

  if (v1Check.status === 404) {
    const betaCheck = await fetch(`https://graph.microsoft.com/beta/planner/plans/${planId}`, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (betaCheck.ok) {
      const [bucketsData, tasksData, planDetails] = await Promise.all([
        graphRequest<{ value: PlannerBucket[] }>(token, `/planner/plans/${planId}/buckets`, "beta"),
        graphRequest<{ value: (PlannerTask & { appliedCategories?: Record<string, boolean> })[] }>(token, `/planner/plans/${planId}/tasks`, "beta"),
        graphRequest<{ categoryDescriptions: Record<string, string | null> }>(token, `/planner/plans/${planId}/details`, "beta")
          .catch(() => ({ categoryDescriptions: {} as Record<string, string | null> })),
      ]);
      const catDesc = planDetails.categoryDescriptions ?? {};
      const tasks = tasksData.value.map((t) => ({
        ...t,
        labels: Object.entries(t.appliedCategories ?? {})
          .filter(([, v]) => v === true)
          .map(([k]) => catDesc[k])
          .filter((name): name is string => Boolean(name)),
      }));
      const assigneeMap = await resolveAssignees(token, tasksData.value);
      return { buckets: bucketsData.value, tasks, categoryDescriptions: catDesc, assigneeMap, source: "basic" as const };
    }

    // Both Graph versions 404 — Premium plan stored in Dataverse
    const orgId = parsePremiumOrgId(planUrl);
    return await fetchPlannerPremiumData(msal, account, planId, orgId);
  }

  throw new Error(`Graph API feil (${v1Check.status}): ${await v1Check.text()}`);
}

// ─── Toggle task complete ─────────────────────────────────────────────────────

export async function togglePlannerTask(
  msal: IPublicClientApplication,
  account: AccountInfo,
  data: PlannerData,
  taskId: string,
  done: boolean
): Promise<void> {
  if (data.source === "premium") {
    if (!data.premiumApiUrl || !data.premiumDvScope) throw new Error("Mangler Dataverse-konfigurasjon.");
    const dvToken = await acquireToken(msal, account, [data.premiumDvScope]);
    const apiBase = data.premiumApiUrl.replace(/\/$/, "");

    // msdyn_progress is a calculated field — update via msdyn_effortcompleted instead.
    // Set effortcompleted = effort (done) or 0 (not done); Dataverse recalculates progress.
    const task = data.tasks.find((t) => t.id === taskId);
    const effort = task?.effort ?? 8; // fallback to 8h if not fetched

    const res = await fetch(
      `${apiBase}/api/data/v9.2/msdyn_projecttasks(${taskId})`,
      {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${dvToken}`,
          "Content-Type": "application/json",
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
        },
        body: JSON.stringify({ msdyn_effortcompleted: done ? Math.max(effort, 0.001) : 0 }),
      }
    );
    if (!res.ok && res.status !== 204) {
      const err = await res.text();
      console.error("[Planner] Toggle PATCH error:", err);
      throw new Error(`Dataverse-feil (${res.status}): ${err.slice(0, 200)}`);
    }
  } else {
    const token = await acquireToken(msal, account, PLANNER_SCOPES);
    // GET the current task to retrieve the required ETag
    const getRes = await fetch(`https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!getRes.ok) throw new Error(`Kunne ikke hente oppgave (${getRes.status})`);
    const etag = getRes.headers.get("ETag") ?? "";
    const patchRes = await fetch(`https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`, {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        "If-Match": etag,
      },
      body: JSON.stringify({ percentComplete: done ? 100 : 0 }),
    });
    if (!patchRes.ok) {
      const err = await patchRes.text();
      throw new Error(`Planner-feil (${patchRes.status}): ${err.slice(0, 200)}`);
    }
  }
}

// ─── Utilities ────────────────────────────────────────────────────────────────

export function parsePlanId(url: string): string | null {
  try {
    const hashIndex = url.indexOf("#");
    if (hashIndex !== -1) {
      const qIndex = url.indexOf("?", hashIndex);
      if (qIndex !== -1) {
        const params = new URLSearchParams(url.slice(qIndex + 1));
        const id = params.get("planId");
        if (id) return id;
      }
    }
    const u = new URL(url);
    const id = u.searchParams.get("planId");
    if (id) return id;
    const premiumMatch = url.match(/\/premiumplan\/([0-9a-f-]{36})/i);
    if (premiumMatch) return premiumMatch[1];
    const match = url.match(/\/plan\/([A-Za-z0-9_-]+)/);
    return match?.[1] ?? null;
  } catch {
    return null;
  }
}

function parsePremiumOrgId(url: string): string | null {
  const match = url.match(/\/org\/([0-9a-f-]{36})/i);
  return match?.[1] ?? null;
}

export function taskStatus(pct: number): "not_started" | "in_progress" | "done" {
  if (pct === 100) return "done";
  if (pct > 0) return "in_progress";
  return "not_started";
}
