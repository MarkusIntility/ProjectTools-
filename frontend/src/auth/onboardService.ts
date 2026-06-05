import type { AccountInfo, IPublicClientApplication } from "@azure/msal-browser";

const ONBOARD_BASE = "https://onboard.intility.com";
const ONBOARD_SCOPES = [
  `api://2708edf9-362e-4d99-a54e-469d5a111a80/Onboard_Access`,
];

async function getOnboardToken(msal: IPublicClientApplication, account: AccountInfo): Promise<string> {
  const result = await msal.acquireTokenSilent({ scopes: ONBOARD_SCOPES, account });
  return result.accessToken;
}

async function onboardGet(token: string, path: string): Promise<unknown> {
  const res = await fetch(`${ONBOARD_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) return { __status: res.status, __statusText: res.statusText };
  const ct = res.headers.get("content-type") ?? "";
  if (ct.includes("json")) return res.json();
  return { __text: await res.text() };
}

/** Call this once from console / a button to discover what the Onboard API exposes.
 *  Results are logged to the browser console.
 */
export async function probeOnboardApi(msal: IPublicClientApplication, account: AccountInfo): Promise<void> {
  console.group("[Onboard probe] Fetching token...");
  const token = await getOnboardToken(msal, account);
  console.log("[Onboard probe] Token acquired ✓");
  console.groupEnd();

  const probeTargets = [
    // Discovery / meta
    "/api",
    "/api/v1",
    "/api/v2",
    "/swagger/v1/swagger.json",
    // Common resource collections (singular and plural)
    "/api/v1/computers",
    "/api/v1/computer",
    "/api/v1/devices",
    "/api/v1/locations",
    "/api/v1/location",
    "/api/v1/applications",
    "/api/v1/application",
    "/api/v1/apps",
    "/api/v1/users",
    "/api/v1/projects",
    "/api/v1/onboardings",
    "/api/v1/orders",
    "/api/v1/customers",
    "/api/v1/clients",
    "/api/v1/services",
    "/api/v1/licenses",
    // Top-level (no version prefix)
    "/computers",
    "/locations",
    "/applications",
  ];

  const results: Record<string, unknown> = {};

  for (const path of probeTargets) {
    try {
      const data = await onboardGet(token, path);
      results[path] = data;
      const status = (data as { __status?: number }).__status;
      if (!status) {
        console.log(`[Onboard probe] ✅ ${path}`, data);
      } else {
        console.log(`[Onboard probe] ❌ ${path} → HTTP ${status}`);
      }
    } catch (e) {
      results[path] = { __error: String(e) };
      console.log(`[Onboard probe] 💥 ${path} →`, e);
    }
  }

  console.group("[Onboard probe] Full results (copy this):");
  console.log(JSON.stringify(results, null, 2));
  console.groupEnd();
}

export { ONBOARD_SCOPES };
