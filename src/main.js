import {PublicClientApplication, InteractionRequiredAuthError} from '@azure/msal-browser';
import {CalendarReadError, readUpcomingEvents, formatEventDate, isWjTenantAccount} from './graph-calendar.js';

const WJ_TENANT_ID = 'dcc6cd1a-0b51-4889-9c2f-e4d97c65f29d';

// Public OAuth client identifiers only. No application secret is used in this SPA.
const msal = new PublicClientApplication({
  auth: {
    clientId: '3f887f86-9c36-4dd4-83e7-857c2b231789',
    authority: `https://login.microsoftonline.com/${WJ_TENANT_ID}`,
    redirectUri: new URL('./redirect.html', window.location.href).href,
    ...(window.self !== window.top ? {
      popupRelayUri: new URL('./popup-relay.html', window.location.href).href,
    } : {}),
    postLogoutRedirectUri: new URL('./redirect.html', window.location.href).href,
  },
  cache: {cacheLocation: 'sessionStorage'},
});

const SCOPES = ['User.Read', 'Calendars.Read'];
const elements = Object.fromEntries(
  ['connect', 'refresh', 'signout', 'account', 'status', 'events']
    .map(id => [id, document.getElementById(id)]),
);

let busy = false;
let signedInAccount = null;

function report(message, error = false) {
  elements.status.textContent = message;
  elements.status.setAttribute('role', error ? 'alert' : 'status');
}

function renderControls() {
  elements.connect.disabled = busy || Boolean(signedInAccount);
  elements.refresh.disabled = busy || !signedInAccount;
  elements.signout.disabled = busy || !signedInAccount;
  elements.account.textContent = signedInAccount
    ? `Verbonden: ${signedInAccount.username ?? signedInAccount.name ?? 'Microsoft 365'}`
    : 'Niet aangemeld';
}

async function withBusy(task) {
  if (busy) return;
  busy = true;
  renderControls();
  try {
    await task();
  } catch (error) {
    // Never display a raw OAuth/Graph response, authorization code, token or URL.
    const code = String(error?.errorCode || error?.subError || '');
    if (error instanceof CalendarReadError) {
      const messages = {
        permission: 'Geen toegang tot deze agenda. Controleer de Microsoft-toestemming of meld je opnieuw aan.',
        throttled: 'Microsoft beperkt tijdelijk het aantal aanvragen. Wacht even en probeer opnieuw.',
        timeout: 'Het ophalen duurt te lang. Controleer de verbinding en probeer opnieuw.',
        too_many_pages: 'De agenda bevat te veel resultaten om veilig volledig te tonen. Verkort de periode.',
        too_many_events: 'De agenda bevat te veel afspraken om veilig volledig te tonen. Verkort de periode.',
      };
      report(messages[error.code] || 'De Outlook-agenda kon niet volledig worden opgehaald. Er is niets gewijzigd.', true);
    } else if (error instanceof InteractionRequiredAuthError) {
      report('Je Microsoft-sessie is verlopen. Klik opnieuw op Afspraken ophalen om je aan te melden.', true);
    } else if (/consent_required|65001|90094|65004/.test(code)) {
      report('Microsoft vereist goedkeuring van de WJ Projects-beheerder. Vraag eenmalige toestemming voor de organisatie aan.', true);
    } else if (code === 'wrong_tenant') {
      report('Dit account hoort niet bij de WJ Projects Microsoft 365-omgeving. Meld je aan met je werkaccount.', true);
    } else {
      report('De actie is niet gelukt. Controleer de browserinstellingen en Microsoft-toestemming.', true);
    }
  } finally {
    busy = false;
    renderControls();
  }
}

function renderEvents(events) {
  elements.events.replaceChildren();
  for (const event of events) {
    const item = document.createElement('li');
    const when = document.createElement('span');
    when.className = 'time';
    if (event.isAllDay) {
      const startDate = event.start?.dateTime?.slice(0, 10);
      when.textContent = `${startDate ?? 'Datum onbekend'} · hele dag`;
    } else {
      when.textContent = `${formatEventDate(event.start?.dateTime)} tot ${formatEventDate(event.end?.dateTime)}`;
    }
    const name = document.createElement('span');
    name.className = 'subject';
    name.textContent = event.isCancelled ? `[Geannuleerd] ${event.subject ?? 'Zonder titel'}` : event.subject ?? 'Zonder titel';
    item.append(when, name);
    elements.events.append(item);
  }
}

async function graphToken() {
  const request = {account: signedInAccount, scopes: SCOPES};
  try {
    return (await msal.acquireTokenSilent(request)).accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) throw error;
    return (await msal.acquireTokenPopup(request)).accessToken;
  }
}

async function refresh() {
  if (!signedInAccount) return;
  report('Outlook-afspraken worden opgehaald…');
  const token = await graphToken();
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 30000);
  try {
    const events = await readUpcomingEvents({accessToken: token, signal: controller.signal});
    // Only render after ALL Graph pages succeed; never display a partial result as complete.
    renderEvents(events);
    report(events.length ? `${events.length} afspraak/afspraken opgehaald. Alleen-lezen Microsoft Graph-koppeling.` : 'Geen afspraken in de komende 14 dagen.');
  } finally {
    clearTimeout(timeout);
  }
}

elements.connect.addEventListener('click', () => withBusy(async () => {
  report('Microsoft-aanmelding wordt geopend…');
  const result = await msal.loginPopup({scopes: SCOPES, prompt: 'select_account'});
  if (!isWjTenantAccount(result.account, WJ_TENANT_ID)) {
    const failure = new Error('Andere Microsoft tenant');
    failure.errorCode = 'wrong_tenant';
    throw failure;
  }
  // Never let another Microsoft account see the previous user's calendar in this tab.
  elements.events.replaceChildren();
  signedInAccount = result.account;
  msal.setActiveAccount(signedInAccount);
  renderControls();
  await refresh();
}));

elements.refresh.addEventListener('click', () => withBusy(refresh));

elements.signout.addEventListener('click', () => withBusy(async () => {
  // Local sign-out only: do not sign the user out of other Microsoft applications.
  if (signedInAccount) await msal.clearCache({account: signedInAccount});
  signedInAccount = null;
  msal.setActiveAccount(null);
  elements.events.replaceChildren();
  report('De lokale aanmeldsessie is gewist. Outlook blijft aangemeld.');
}));

withBusy(async () => {
  await msal.initialize();
  const candidates = [msal.getActiveAccount(), ...msal.getAllAccounts()];
  signedInAccount = candidates.find(account => isWjTenantAccount(account, WJ_TENANT_ID)) ?? null;
  report(signedInAccount ? 'Je sessie is gevonden. Klik op Afspraken ophalen.' : 'Klik op Inloggen met Microsoft. Er worden geen agenda-afspraken gewijzigd.');
});
