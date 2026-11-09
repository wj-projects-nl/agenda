import test from 'node:test';
import assert from 'node:assert/strict';
import {CalendarReadError, readUpcomingEvents, formatEventDate, isWjTenantAccount} from '../src/graph-calendar.js';

const now = new Date('2026-11-08T12:00:00.000Z');
const item = (id, start = '2026-11-09T08:00:00.0000000') => ({
  id, subject: `Afspraak ${id}`, start: {dateTime: start, timeZone: 'UTC'},
  end: {dateTime: '2026-11-09T09:00:00.0000000', timeZone: 'UTC'}, isAllDay: false,
});
const ok = body => ({ok: true, status: 200, json: async () => body});
const read = fetcher => readUpcomingEvents({accessToken: 'synthetisch-testtoken', fetcher, now});

test('reads every Graph page, sends only a bearer token and preserves unique events', async () => {
  const calls = [];
  const page2 = 'https://graph.microsoft.com/v1.0/me/calendarView?%24skiptoken=mock';
  const events = await read((url, options) => {
    calls.push({url, options});
    return calls.length === 1
      ? ok({value: [item('B'), item('A')], '@odata.nextLink': page2})
      : ok({value: [item('A'), item('C')]});
  });
  assert.deepEqual(events.map(e => e.id), ['A', 'B', 'C']);
  assert.equal(calls.length, 2);
  assert.match(calls[0].url, /startDateTime=2026-11-08T12%3A00%3A00.000Z/);
  assert.equal(calls[0].options.headers.Prefer, 'outlook.timezone="UTC"');
  assert.equal(calls[0].options.headers.Authorization, 'Bearer synthetisch-testtoken');
  assert.deepEqual(Object.keys(events[0]).sort(), ['end', 'id', 'isAllDay', 'isCancelled', 'start', 'subject']);
});

test('fails closed when Graph returns a different origin or path', async () => {
  for (const link of ['https://example.org/exfiltrate',
    'https://graph.microsoft.com/v1.0/me/messages',
    'https://graph.microsoft.com/v1.0/me/calendarView/other',
    'https://graph.microsoft.com/v1.0/me/calendarView#token']) {
    await assert.rejects(read(() => ok({value: [item('A')], '@odata.nextLink': link})),
      e => e instanceof CalendarReadError && e.code === 'invalid_page');
  }
});

test('detects repeated nextLink without looping or showing partial data', async () => {
  const repeat = 'https://graph.microsoft.com/v1.0/me/calendarView?skip=2';
  let count = 0;
  await assert.rejects(read(() => {
    count++;
    return ok({value: [item(String(count))], '@odata.nextLink': repeat});
  }), e => e.code === 'page_loop');
  assert.equal(count, 2);
});

test('rejects incomplete Graph responses and too many events', async () => {
  await assert.rejects(read(() => ok({detail: 'bad response'})), e => e.code === 'invalid_response');
  await assert.rejects(read(() => ok({value: [{subject: 'No id'}]})), e => e.code === 'invalid_response');
  await assert.rejects(read(() => ok({value: Array.from({length: 1001}, (_, i) => item(String(i)))})),
    e => e.code === 'too_many_events');
});

test('does not leak the Graph response or token in errors', async () => {
  for (const [status, expected] of [[401, 'permission'], [403, 'permission'], [429, 'throttled'], [503, 'throttled'], [500, 'http']]) {
    await assert.rejects(read(() => ({ok: false, status})), e => e.code === expected && !e.message.includes('synthetisch-testtoken'));
  }
  await assert.rejects(read(() => { throw new TypeError('private data'); }), e => e.code === 'network' && !e.message.includes('private'));
});

test('renders winter and summer local times correctly', () => {
  assert.match(formatEventDate('2026-11-09T08:00:00.0000000'), /09:00/);
  assert.match(formatEventDate('2026-07-09T08:00:00.0000000'), /10:00/);
  assert.equal(formatEventDate('not-a-date'), 'Tijd onbekend');
});

test('only accepts an account from the configured WJ tenant', () => {
  const tenant = 'dcc6cd1a-0b51-4889-9c2f-e4d97c65f29d';
  assert.equal(isWjTenantAccount({tenantId: tenant}, tenant), true);
  assert.equal(isWjTenantAccount({tenantId: 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'}, tenant), false);
  assert.equal(isWjTenantAccount({username: 'ahash@wj-projects.nl'}, tenant), false);
  assert.equal(isWjTenantAccount(null, tenant), false);
});
