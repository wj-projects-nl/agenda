// Pure Graph calendar reader. It does not create, update, delete or persist events.
const GRAPH_ORIGIN = 'https://graph.microsoft.com';
const CALENDAR_PATH = '/v1.0/me/calendarView';
const MAX_PAGES = 30;
const MAX_EVENTS = 1000;

export class CalendarReadError extends Error {
  constructor(code) {
    super(code);
    this.name = 'CalendarReadError';
    this.code = code;
  }
}

function validPageUrl(url) {
  const parsed = new URL(url);
  if (parsed.origin !== GRAPH_ORIGIN || parsed.pathname !== CALENDAR_PATH || parsed.username || parsed.password || parsed.hash) {
    throw new CalendarReadError('invalid_page');
  }
  return parsed.href;
}

function initialPage(now, days) {
  if (!(now instanceof Date) || Number.isNaN(now.getTime()) || !Number.isInteger(days) || days < 1 || days > 31) {
    throw new CalendarReadError('invalid_period');
  }
  const query = new URLSearchParams({
    startDateTime: now.toISOString(),
    endDateTime: new Date(now.getTime() + days * 86400000).toISOString(),
    '$select': 'id,subject,start,end,isAllDay,isCancelled',
    '$top': '50',
    '$orderby': 'start/dateTime',
  });
  return `${GRAPH_ORIGIN}${CALENDAR_PATH}?${query}`;
}

export async function readUpcomingEvents({accessToken, fetcher = fetch, now = new Date(), days = 14, signal} = {}) {
  if (!accessToken || typeof accessToken !== 'string') throw new CalendarReadError('missing_token');
  let next = initialPage(now, days);
  const visited = new Set();
  const byId = new Map();
  for (let page = 0; next; page++) {
    if (page >= MAX_PAGES) throw new CalendarReadError('too_many_pages');
    next = validPageUrl(next);
    if (visited.has(next)) throw new CalendarReadError('page_loop');
    visited.add(next);

    let response;
    try {
      response = await fetcher(next, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: 'application/json',
          Prefer: 'outlook.timezone="UTC"',
        },
        cache: 'no-store',
        signal,
      });
    } catch (error) {
      if (error?.name === 'AbortError') throw new CalendarReadError('timeout');
      throw new CalendarReadError('network');
    }
    if (!response?.ok) {
      const status = response?.status;
      if (status === 401 || status === 403) throw new CalendarReadError('permission');
      if (status === 429 || status === 503) throw new CalendarReadError('throttled');
      throw new CalendarReadError('http');
    }

    let payload;
    try { payload = await response.json(); }
    catch { throw new CalendarReadError('invalid_response'); }
    if (!payload || !Array.isArray(payload.value)) throw new CalendarReadError('invalid_response');
    for (const event of payload.value) {
      if (!event || typeof event.id !== 'string' || !event.id) throw new CalendarReadError('invalid_response');
      // Only retain the fields needed by the list view, not sensitive Graph metadata.
      byId.set(event.id, {
        id: event.id,
        subject: typeof event.subject === 'string' ? event.subject : '',
        start: event.start,
        end: event.end,
        isAllDay: event.isAllDay === true,
        isCancelled: event.isCancelled === true,
      });
      if (byId.size > MAX_EVENTS) throw new CalendarReadError('too_many_events');
    }
    next = payload['@odata.nextLink'] ?? null;
    if (next !== null && typeof next !== 'string') throw new CalendarReadError('invalid_page');
  }
  return [...byId.values()].sort((a, b) => {
    const start = (a.start?.dateTime || '').localeCompare(b.start?.dateTime || '');
    return start || a.id.localeCompare(b.id);
  });
}

export function formatEventDate(dateTime) {
  if (typeof dateTime !== 'string' || !dateTime) return 'Tijd onbekend';
  // Graph is explicitly requested to return UTC. A missing offset therefore means UTC.
  const utc = /(?:Z|[+-]\d\d:\d\d)$/i.test(dateTime) ? dateTime : `${dateTime}Z`;
  const date = new Date(utc);
  if (Number.isNaN(date.getTime())) return 'Tijd onbekend';
  return new Intl.DateTimeFormat('nl-NL', {
    timeZone: 'Europe/Amsterdam', weekday: 'short', day: 'numeric', month: 'short',
    hour: '2-digit', minute: '2-digit',
  }).format(date);
}

export function isWjTenantAccount(account, tenantId) {
  const tenant = account?.tenantId ?? account?.idTokenClaims?.tid;
  return Boolean(tenant && tenant.toLowerCase() === tenantId.toLowerCase());
}
