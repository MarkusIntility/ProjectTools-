import type { AccountInfo, IPublicClientApplication } from "@azure/msal-browser";
import { CALENDAR_SCOPES } from "./msalConfig";

export interface OutlookEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  bodyPreview: string;
  categories: string[];
}

async function acquireCalendarToken(
  msal: IPublicClientApplication,
  account: AccountInfo
): Promise<string> {
  try {
    const r = await msal.acquireTokenSilent({ scopes: CALENDAR_SCOPES, account });
    return r.accessToken;
  } catch {
    sessionStorage.setItem("app.returnUrl", window.location.href);
    await msal.acquireTokenRedirect({ scopes: CALENDAR_SCOPES, account });
    throw new Error("Omdirigerer for tilgang…");
  }
}

export async function fetchOutlookMeetingsByCategory(
  msal: IPublicClientApplication,
  account: AccountInfo,
  categoryName: string
): Promise<OutlookEvent[]> {
  const token = await acquireCalendarToken(msal, account);

  const nowIso = new Date().toISOString();
  const safeCat = categoryName.replace(/'/g, "''");
  const filter = `categories/any(c:c eq '${safeCat}') and start/dateTime ge '${nowIso}'`;
  const params = new URLSearchParams({
    $filter: filter,
    $select: "id,subject,start,end,bodyPreview,categories",
    $orderby: "start/dateTime asc",
    $top: "200",
  });

  const res = await fetch(`https://graph.microsoft.com/v1.0/me/events?${params}`, {
    headers: {
      Authorization: `Bearer ${token}`,
      Prefer: 'outlook.timezone="UTC"',
    },
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Kalender-feil (${res.status}): ${err}`);
  }

  const data: { value: OutlookEvent[] } = await res.json();
  return data.value ?? [];
}
