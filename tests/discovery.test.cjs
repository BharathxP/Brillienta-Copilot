const test = require('node:test');
const assert = require('node:assert/strict');
const D = require('../lib-commonjs/webparts/brillientaCopilot/core/discovery.js');
const { publisherSignIn, adminConsentUrl, newGuid } = require('../lib-commonjs/webparts/brillientaCopilot/core/publisher.js');

const lists = (titles) => titles.map((t, i) => ({ Id: `id-${i}-${t}`, Title: t }));
const site = (titles) => ({
  url: 'https://contoso.sharepoint.com/sites/ops',
  host: 'contoso.sharepoint.com',
  siteId: '{11111111-1111-1111-1111-111111111111}',
  webId: '22222222-2222-2222-2222-222222222222',
  lists: lists(titles)
});
const HD = ['Tickets', 'Comments', 'KBArticles', 'Categories', 'Statuses', 'Priorities', 'Agents', 'CopilotActions', 'Settings'];
const AM = ['Assets', 'Requests', 'AssetTypes', 'Statuses', 'Comments'];
const ED = ['Employees', 'Departments', 'Profiles'];

test('finds all three apps on one site with their default prefixes', () => {
  const apps = D.findApps(site([...HD.map((s) => `HD_${s}`), ...AM.map((s) => `AM_${s}`), ...ED.map((s) => `BrillED_${s}`), 'Documents']));
  assert.deepEqual(apps.map((a) => [a.key, a.prefix]), [['helpdesk', 'HD_'], ['assets', 'AM_'], ['directory', 'BrillED_']]);
  assert.equal(apps[0].graphSiteId, 'contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222');
  assert.equal(apps[0].listIds.copilotActions, 'id-7-HD_CopilotActions');
});

test('finds an app installed with a custom prefix', () => {
  const apps = D.findApps(site(HD.map((s) => `ITDesk_${s}`)));
  assert.deepEqual(apps.map((a) => [a.key, a.prefix]), [['helpdesk', 'ITDesk_']]);
});

test('does not mistake other apps for Helpdesk or Assets', () => {
  // HRMS has "HRMS Tickets" and TaskManager has "TM_Comments", but neither has the rest of Helpdesk's lists.
  const apps = D.findApps(site(['HRMS Tickets', 'HRMS Employees', 'TM_Tasks', 'TM_Comments', 'TM_Statuses', 'HD_Tickets', 'HD_Comments']));
  assert.deepEqual(apps, []);
});

test('matching is case-insensitive', () => {
  assert.equal(D.findApps(site(ED.map((s) => `brilled_${s.toLowerCase()}`))).length, 1);
});

test('reads Helpdesk options, defaulting safely', () => {
  assert.deepEqual(D.helpdeskOptions(JSON.stringify({ ticketPrefix: 'IT-', copilot: { enabled: true } })), { ticketPrefix: 'IT-', copilotEnabled: 'true' });
  assert.deepEqual(D.helpdeskOptions(JSON.stringify({ copilot: { enabled: false } })), { ticketPrefix: 'HD-', copilotEnabled: 'false' });
  assert.deepEqual(D.helpdeskOptions('not json'), { ticketPrefix: 'HD-', copilotEnabled: 'false' });
  assert.deepEqual(D.helpdeskOptions(undefined), { ticketPrefix: 'HD-', copilotEnabled: 'false' });
});

test('normalises site addresses and rejects anything that is not SharePoint over https', () => {
  assert.equal(D.normaliseSiteUrl(' https://contoso.sharepoint.com/sites/hr/SitePages/Home.aspx '), 'https://contoso.sharepoint.com/sites/hr');
  assert.equal(D.normaliseSiteUrl('https://contoso.sharepoint.com/sites/hr/'), 'https://contoso.sharepoint.com/sites/hr');
  assert.equal(D.normaliseSiteUrl('https://contoso.sharepoint.com'), 'https://contoso.sharepoint.com');
  assert.equal(D.normaliseSiteUrl('http://contoso.sharepoint.com/sites/hr'), undefined);
  assert.equal(D.normaliseSiteUrl('https://evil.example.com/sites/hr'), undefined);
  assert.equal(D.normaliseSiteUrl('javascript:alert(1)'), undefined);
});

test('publisher sign-in is off until both values are filled in; consent URL and GUIDs are well formed', () => {
  assert.equal(publisherSignIn(), undefined);
  assert.ok(publisherSignIn({ publisher: 'B', appId: '0bb3470d-1450-4033-9c7e-38e3749dec6a', oauthRegistrationId: 'abcdefgh123' }));
  const url = adminConsentUrl('0bb3470d-1450-4033-9c7e-38e3749dec6a', 'tid');
  assert.match(url, /^https:\/\/login\.microsoftonline\.com\/tid\/v2\.0\/adminconsent\?client_id=0bb3470d/);
  assert.match(decodeURIComponent(url), /Sites\.ReadWrite\.All .*User\.Read offline_access/);
  for (let i = 0; i < 50; i++) assert.match(newGuid(), /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
});
