const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const Ajv = require('ajv');
const addFormats = require('ajv-formats');

const core = '../lib-commonjs/webparts/brillientaCopilot/core';
const { buildAgent, buildInstructions, validateAgentInput, MODULES, INSTRUCTIONS_LIMIT } = require(`${core}/compose.js`);
const { HELPDESK_ACTIONS, TICKET_FILTER_FIELDS } = require(`${core}/apps/helpdesk.js`);
const { ASSET_FILTER_FIELDS, REQUEST_FILTER_FIELDS, ASSET_ACTIONS } = require(`${core}/apps/assets.js`);
const { PEOPLE_FILTER_FIELDS, PEOPLE_FIELDS } = require(`${core}/apps/directory.js`);
const { buildPackageZip } = require(`${core}/package.js`);

const SITE = 'contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222';
const target = (key, over = {}) => {
  const m = MODULES[key];
  return {
    key,
    siteUrl: `https://contoso.sharepoint.com/sites/${key}`,
    graphSiteId: SITE,
    prefix: m.defaultPrefix,
    listIds: Object.fromEntries(m.lists.map((l) => [l.key, `${key}-${l.key}-id`])),
    options: key === 'helpdesk' ? { ticketPrefix: 'HD-', copilotEnabled: 'true' } : key === 'assets' ? { copilotEnabled: 'true', acknowledgementsEnabled: 'true' } : {},
    ...over
  };
};
const input = (apps = ['helpdesk', 'assets', 'directory'], over = {}) => ({
  agentName: 'Brillienta',
  orgName: 'Contoso',
  appId: '0bb3470d-1450-4033-9c7e-38e3749dec6a',
  appVersion: '1.0.0.0',
  oauthRegistrationId: 'reg-123',
  developerName: 'Contoso',
  websiteUrl: 'https://contoso.com',
  privacyUrl: '',
  termsUrl: '',
  brandColor: '#008d62',
  apps: apps.map((k) => target(k)),
  ...over
});

const SCHEMAS = path.join(__dirname, '..', 'node_modules', '@microsoft', 'app-manifest', 'build', 'json-schemas');
function validator(rel) {
  const ajv = new Ajv({ strict: false, allErrors: true, validateFormats: true });
  addFormats(ajv);
  const schema = JSON.parse(fs.readFileSync(path.join(SCHEMAS, rel), 'utf8'));
  delete schema.$schema;
  return ajv.compile(schema);
}

test('all three apps: one core plugin plus one plugin and OpenAPI file per app', () => {
  const files = buildAgent(input());
  assert.deepEqual(Object.keys(files).sort(), [
    'ai-plugin-assets.json', 'ai-plugin-core.json', 'ai-plugin-directory.json', 'ai-plugin-helpdesk.json', 'declarativeAgent.json', 'manifest.json',
    'openapi-assets.json', 'openapi-core.json', 'openapi-directory.json', 'openapi-helpdesk.json'
  ]);
  const agent = JSON.parse(files['declarativeAgent.json']);
  assert.deepEqual(agent.actions.map((a) => a.file), ['ai-plugin-core.json', 'ai-plugin-helpdesk.json', 'ai-plugin-assets.json', 'ai-plugin-directory.json']);
  assert.ok(agent.conversation_starters.length <= 6);
  assert.deepEqual(agent.capabilities[0].items_by_url.map((u) => u.url), ['https://contoso.sharepoint.com/sites/helpdesk', 'https://contoso.sharepoint.com/sites/assets', 'https://contoso.sharepoint.com/sites/directory']);
});

test('generated files validate against Microsoft\'s published JSON schemas', () => {
  const files = buildAgent(input());
  const checks = [
    ['manifest.json', 'teams/v1.19/MicrosoftTeams.schema.json'],
    ['declarativeAgent.json', 'copilot/declarative-agent/v1.3/schema.json'],
    ...['core', 'helpdesk', 'assets', 'directory'].map((k) => [`ai-plugin-${k}.json`, 'copilot/plugin/v2.2/schema.json'])
  ];
  for (const [file, schema] of checks) {
    const validate = validator(schema);
    const ok = validate(JSON.parse(files[file]));
    assert.ok(ok, `${file}: ${JSON.stringify(validate.errors, null, 1)}`);
  }
});

test('every plugin function has a matching operation in its OpenAPI file, on Graph v1.0, with no Prefer header', () => {
  const files = buildAgent(input());
  for (const k of ['core', 'helpdesk', 'assets', 'directory']) {
    const plugin = JSON.parse(files[`ai-plugin-${k}.json`]);
    const spec = JSON.parse(files[plugin.runtimes[0].spec.url]);
    assert.equal(spec.servers[0].url, 'https://graph.microsoft.com/v1.0');
    const ops = [];
    for (const p of Object.values(spec.paths)) for (const op of Object.values(p)) ops.push(op);
    const ids = ops.map((o) => o.operationId);
    assert.deepEqual(plugin.functions.map((f) => f.name).sort(), [...ids].sort(), k);
    assert.deepEqual(plugin.runtimes[0].run_for_functions.sort(), [...ids].sort());
    assert.equal(plugin.runtimes[0].auth.reference_id, 'reg-123');
    for (const o of ops) assert.ok(!(o.parameters || []).some((x) => x.in === 'header'), `${o.operationId} sends a header`);
  }
});

test('function names are unique across the whole agent', () => {
  const names = ['getCurrentUser'].concat(...Object.values(MODULES).map((m) => m.functions.map((f) => f.name)));
  assert.equal(new Set(names).size, names.length);
});

test('instructions fit Microsoft\'s 8,000-character limit for every combination of apps', () => {
  const keys = ['helpdesk', 'assets', 'directory'];
  for (let mask = 1; mask < 8; mask++) {
    const apps = keys.filter((_, i) => mask & (1 << i));
    const len = buildInstructions(input(apps)).length;
    assert.ok(len <= INSTRUCTIONS_LIMIT, `${apps.join('+')}: ${len}`);
  }
});

test('safety rules are in the instructions', () => {
  const text = buildInstructions(input());
  assert.match(text, /For non-agents ALWAYS add fields\/RequesterEmail eq/);
  assert.match(text, /Never show comments with IsInternal true to non-agents/);
  assert.match(text, /Never show anyone with HiddenFromDirectory true/);
  assert.match(text, /fields\/EmpStatus eq 'Active'/);
  assert.match(text, /Confirm before creating or changing anything/);
});

test('every write asks the user to confirm first', () => {
  const files = buildAgent(input());
  for (const k of ['helpdesk', 'assets', 'directory']) {
    const plugin = JSON.parse(files[`ai-plugin-${k}.json`]);
    const spec = JSON.parse(files[`openapi-${k}.json`]);
    const writes = [];
    for (const p of Object.values(spec.paths)) for (const [verb, op] of Object.entries(p)) if (verb !== 'get') writes.push(op.operationId);
    for (const w of writes) assert.ok(plugin.functions.find((f) => f.name === w).capabilities?.confirmation, `${w} has no confirmation`);
    if (k === 'directory') assert.deepEqual(writes, [], 'the directory is read-only');
  }
});

test('Directory never requests private fields', () => {
  for (const f of ['MobilePhone', 'HireDate', 'EmployeeNumber', 'GraphExtras', 'Overrides', 'BirthdayMD']) assert.ok(!PEOPLE_FIELDS.includes(f), f);
  const spec = JSON.parse(buildAgent(input(['directory']))['openapi-directory.json']);
  assert.doesNotMatch(JSON.stringify(spec), /MobilePhone|HireDate|EmployeeNumber/);
});

test('validation catches what would break the agent', () => {
  assert.deepEqual(validateAgentInput(input()), []);
  const off = input(['helpdesk']);
  off.apps[0].options.copilotEnabled = 'false';
  assert.match(validateAgentInput(off).join(' '), /Process Copilot requests/);
  const missing = input(['assets']);
  delete missing.apps[0].listIds.requests;
  assert.match(validateAgentInput(missing).join(' '), /AM_Requests was not found/);
  assert.match(validateAgentInput(input([], {})).join(' '), /No Brillienta app was found/);
  assert.match(validateAgentInput(input(['assets'], { appId: 'x', agentName: 'A name that is far too long for Teams', oauthRegistrationId: '' })).join(' '), /GUID.*OAuth.*30 characters/);
  const twice = input(['assets']);
  twice.apps.push(target('assets'));
  assert.match(validateAgentInput(twice).join(' '), /added twice/);
  const badSite = input(['directory']);
  badSite.apps[0].graphSiteId = 'contoso';
  assert.match(validateAgentInput(badSite).join(' '), /not a Graph site id/);
});

test('package zip holds every file plus icons', () => {
  const zip = buildPackageZip(buildAgent(input()), { color: new Uint8Array([1, 2, 3]), outline: new Uint8Array([4]) });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'bc-'));
  const file = path.join(dir, 'agent.zip');
  fs.writeFileSync(file, Buffer.from(zip));
  const listing = execFileSync('unzip', ['-Z1', file]).toString().trim().split('\n').sort();
  assert.deepEqual(listing, [
    'ai-plugin-assets.json', 'ai-plugin-core.json', 'ai-plugin-directory.json', 'ai-plugin-helpdesk.json', 'color.png', 'declarativeAgent.json', 'manifest.json',
    'openapi-assets.json', 'openapi-core.json', 'openapi-directory.json', 'openapi-helpdesk.json', 'outline.png'
  ]);
  const manifest = JSON.parse(execFileSync('unzip', ['-p', file, 'manifest.json']).toString());
  assert.equal(manifest.id, '0bb3470d-1450-4033-9c7e-38e3749dec6a');
});

// ------------------------------------------------------------------ drift guards against the product repos

const REPOS = path.join(__dirname, '..', '..');
const schemaOf = (repo, rel) => {
  const p = path.join(REPOS, repo, rel);
  return fs.existsSync(p) ? fs.readFileSync(p, 'utf8') : null;
};
/** The text of one list definition in a product schema.ts, from "key: '<key>'" to the next list. */
function listBlock(src, key) {
  const start = src.indexOf(`key: '${key}'`);
  assert.ok(start >= 0, `list ${key} not found`);
  const next = src.indexOf("\n    key: '", start + 10);
  return src.slice(start, next < 0 ? undefined : next);
}
function assertIndexed(src, key, fields) {
  const block = listBlock(src, key);
  for (const f of fields) assert.match(block, new RegExp(`name: '${f}'[^}\\n]*indexed: true`), `${key}.${f} must be indexed`);
}
function assertColumns(src, key, fields) {
  const block = listBlock(src, key);
  for (const f of fields.filter((x) => !['Title', 'Created', 'Modified'].includes(x))) assert.match(block, new RegExp(`name: '${f}'`), `${key}.${f} missing`);
}

const HD = schemaOf('brillientahelpdeskwebpart', 'src/webparts/helpdesk/schema/schema.ts');
test('Helpdesk: filter columns are indexed, list suffixes exist and queued actions match the app', { skip: !HD && 'Helpdesk repo not checked out next to this one' }, () => {
  assertIndexed(HD, 'tickets', TICKET_FILTER_FIELDS.filter((f) => f !== 'ApprovalStatus'));
  for (const l of MODULES.helpdesk.lists) assert.match(HD, new RegExp(`suffix: '${l.suffix}'`), l.suffix);
  const actionsSrc = fs.readFileSync(path.join(REPOS, 'brillientahelpdeskwebpart/src/webparts/helpdesk/services/copilot/actions.ts'), 'utf8');
  const appActions = [...actionsSrc.matchAll(/\{ name: '(\w+)', description:/g)].map((m) => m[1]);
  assert.deepEqual(HELPDESK_ACTIONS.map((a) => a.name), appActions);
});

const AM = schemaOf('brillienta-assetmanagement', 'src/webparts/assetManagement/schema/schema.ts');
test('Asset Management: filter columns are indexed and requested columns exist', { skip: !AM && 'Asset Management repo not checked out next to this one' }, () => {
  assertIndexed(AM, 'assets', ASSET_FILTER_FIELDS);
  assertIndexed(AM, 'requests', REQUEST_FILTER_FIELDS);
  for (const l of MODULES.assets.lists) assert.match(AM, new RegExp(`suffix: '${l.suffix}'`), l.suffix);
});

const ED = schemaOf('brillientaemployeedirectory', 'src/webparts/employeeDirectory/schema/schema.ts');
test('Employee Directory: filter columns are indexed and requested columns exist', { skip: !ED && 'Employee Directory repo not checked out next to this one' }, () => {
  assertIndexed(ED, 'employees', PEOPLE_FILTER_FIELDS);
  assertColumns(ED, 'employees', PEOPLE_FIELDS);
  for (const l of MODULES.directory.lists) assert.match(ED, new RegExp(`suffix: '${l.suffix}'`), l.suffix);
});

test('Assets: requests and receipt confirmations when the site supports them', () => {
  const files = buildAgent(input(['assets']));
  const plugin = JSON.parse(files['ai-plugin-assets.json']);
  assert.deepEqual(plugin.functions.map((f) => f.name), ['searchAssets', 'getAsset', 'searchAssetRequests', 'listAssetTypes', 'queueAssetAction', 'getAssetActionResult', 'confirmReceipt']);
  const spec = JSON.parse(files['openapi-assets.json']);
  const queue = spec.paths['/sites/' + SITE + '/lists/assets-copilotActions-id/items'].post;
  assert.deepEqual(queue.requestBody.content['application/json'].schema.properties.fields.properties.ActionName.enum, ['requestAsset', 'reportIssue', 'returnAsset', 'cancelRequest']);
  const ack = spec.paths['/sites/' + SITE + '/lists/assets-acknowledgements-id/items'].post;
  assert.deepEqual(ack.requestBody.content['application/json'].schema.properties.fields.required, ['Title', 'AssetId', 'AckResponse', 'AckStatus']);
  const text = buildInstructions(input(['assets']));
  assert.match(text, /queueAssetAction/);
  assert.match(text, /Only for the signed-in person and their own assets and requests/);
  assert.deepEqual(MODULES.assets.notes(input(['assets']).apps[0]), []);
});

test('Assets: read-only (and still valid) on older Asset Management or with Copilot requests off', () => {
  const older = input(['assets']);
  delete older.apps[0].listIds.copilotActions;
  const off = input(['assets']);
  off.apps[0].options.copilotEnabled = 'false';
  for (const [i, note] of [[older, /update Asset Management/], [off, /turn on Settings > General > Features/]]) {
    assert.deepEqual(validateAgentInput(i), [], 'optional lists and settings never block the download');
    const plugin = JSON.parse(buildAgent(i)['ai-plugin-assets.json']);
    assert.ok(!plugin.functions.some((f) => f.name === 'queueAssetAction'));
    assert.ok(plugin.functions.some((f) => f.name === 'confirmReceipt'), 'receipt confirmation does not need the queue');
    assert.doesNotMatch(buildInstructions(i), /queueAssetAction/);
    assert.match(MODULES.assets.notes(i.apps[0]).join(' '), note);
  }
  const noAcks = input(['assets']);
  noAcks.apps[0].options.acknowledgementsEnabled = 'false';
  assert.ok(!JSON.parse(buildAgent(noAcks)['ai-plugin-assets.json']).functions.some((f) => f.name === 'confirmReceipt'));
});

if (AM) {
  test('Asset Management: queued actions, the queue list and receipt fields match the app', () => {
    const src = fs.readFileSync(path.join(REPOS, 'brillienta-assetmanagement/src/webparts/assetManagement/services/copilotActions.ts'), 'utf8');
    const names = /COPILOT_ACTION_NAMES = \[([^\]]+)\]/.exec(src)[1].match(/'(\w+)'/g).map((x) => x.slice(1, -1));
    assert.deepEqual(ASSET_ACTIONS.map((a) => a.name), names);
    const queue = listBlock(AM, 'copilotActions');
    assert.match(queue, /suffix: 'CopilotActions'/);
    for (const n of names) assert.match(queue, new RegExp(`'${n}'`));
    assertColumns(AM, 'copilotActions', ['ActionName', 'Payload', 'ActionStatus', 'ResultText', 'RequestId', 'ProcessedAt']);
    assertColumns(AM, 'acknowledgements', ['AssetId', 'AckResponse', 'AckComment', 'AckStatus']);
    assertColumns(AM, 'assetTypes', ['Details', 'Requestable', 'RequiresApproval', 'AllowLoans', 'IsActive']);
  });
}
