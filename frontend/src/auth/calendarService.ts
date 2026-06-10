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

/**
 * Fetches all Outlook calendar instances (including recurring occurrences) with the
 * given category from now until `monthsAhead` months in the future.
 *
 * Uses /me/calendarView instead of /me/events so that each occurrence of a
 * recurring series is returned as a separate item.
 */
export async function fetchOutlookMeetingsByCategory(
  msal: IPublicClientApplication,
  account: AccountInfo,
  categoryName: string,
  monthsAhead = 12
): Promise<OutlookEvent[]> {
  const token = await acquireCalendarToken(msal, account);

  const now = new Date();
  const future = new Date(now);
  future.setMonth(future.getMonth() + monthsAhead);

  const safeCat = categoryName.replace(/'/g, "''");

  const params = new URLSearchParams({
    startDateTime: now.toISOString(),
    endDateTime: future.toISOString(),
    $filter: `categories/any(c:c eq '${safeCat}')`,
    $select: "id,subject,start,end,bodyPreview,categories",
    $orderby: "start/dateTime asc",
    $top: "500",
  });

  let url: string | null =
    `https://graph.microsoft.com/v1.0/me/calendarView?${params}`;
  const allEvents: OutlookEvent[] = [];

  while (url) {
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: 'outlook.timezone="UTC"',
      },
    });

    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Kalender-feil (${res.status}): ${err}`);
    }

    const data: { value: OutlookEvent[]; "@odata.nextLink"?: string } =
      await res.json();
    allEvents.push(...(data.value ?? []));
    url = data["@odata.nextLink"] ?? null;
  }

  return allEvents;
}
