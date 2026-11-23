// Builds one Microsoft 365 Copilot declarative agent for every Brillienta app found: a Teams app manifest, the
// declarative agent, one API plugin + OpenAPI description per app, and a small core plugin (who is asking).
// Each plugin calls Microsoft Graph as the signed-in person, so SharePoint permissions decide what they see.

import { IAgentInput, IAppModule, IAppTarget, AgentFiles, AppKey, Json } from './types';
import { GRAPH_SCOPES, fixed, openApiDocument } from './openapi';
import { helpdeskModule } from './apps/helpdesk';
import { assetsModule } from './apps/assets';
import { directoryModule } from './apps/directory';

export const MODULES: Record<AppKey, IAppModule> = { helpdesk: helpdeskModule, assets: assetsModule, directory: directoryModule };
export const APP_ORDER: AppKey[] = ['helpdesk', 'assets', 'directory'];

/** Microsoft's limit for declarative agent instructions. */
export const INSTRUCTIONS_LIMIT = 8000;
const STARTERS_LIMIT = 6;

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function ordered(apps: IAppTarget[]): IAppTarget[] {
  return APP_ORDER.map((k) => apps.filter((a) => a.key === k)[0]).filter((a): a is IAppTarget => !!a);
}

export function buildInstructions(i: IAgentInput): string {
  const apps = ordered(i.apps);
  const names = apps.map((a) => MODULES[a.key].label).join(', ');
  const head = `You are ${i.agentName}, the ${i.orgName} workplace assistant for: ${names}. All data lives in the organisation's SharePoint lists and you reach it only through your actions. Never invent records, numbers, statuses or people.

# Always
- Call core.getCurrentUser first when the answer depends on who is asking. Their email is mail in lower case (userPrincipalName if mail is empty); their directory id is id.
- Every list action takes an OData $filter on item fields (fields/<Name>). Filter only on the fields each action lists: SharePoint rejects others. Never use $orderby. Use $top 25 unless asked for more.
- Confirm before creating or changing anything. If an action fails or access is denied, say so plainly and give the app's portal link. Answer questions that span apps by calling each app's actions.`;
  const style = `# Style
Concise, friendly, plain words. Use tables for lists of records.`;
  return [head].concat(apps.map((a) => MODULES[a.key].instructions(a)), [style]).join('\n\n');
}

function corePaths(): Json {
  return {
    '/me': {
      get: {
        operationId: 'getCurrentUser',
        summary: 'The signed-in user.',
        description: 'Returns the signed-in user: id, name, email, job title and department.',
        parameters: [fixed('$select', 'id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation')],
        responses: {
          '200': {
            description: 'User',
            content: { 'application/json': { schema: { type: 'object', properties: { id: { type: 'string' }, displayName: { type: 'string' }, mail: { type: 'string' }, userPrincipalName: { type: 'string' } } } } }
          }
        }
      }
    }
  };
}

function plugin(i: IAgentInput, namespace: string, human: string, model: string, functions: Array<{ name: string; description: string; capabilities?: Json }>, specFile: string): Json {
  const names = functions.map((f) => f.name);
  return {
    $schema: 'https://developer.microsoft.com/json-schemas/copilot/plugin/v2.2/schema.json',
    schema_version: 'v2.2',
    name_for_human: `${i.agentName} ${namespace}`.substring(0, 20),
    description_for_human: human.substring(0, 100),
    description_for_model: model.substring(0, 2048),
    namespace,
    functions: functions.map((f) => ({ name: f.name, description: f.description, ...(f.capabilities ? { capabilities: f.capabilities } : {}) })),
    runtimes: [
      {
        type: 'OpenApi',
        auth: { type: 'OAuthPluginVault', reference_id: i.oauthRegistrationId || '${{OAUTH_REGISTRATION_ID}}' },
        spec: { url: specFile },
        run_for_functions: names
      }
    ]
  };
}

export function buildAgent(i: IAgentInput): AgentFiles {
  const apps = ordered(i.apps);
  const files: AgentFiles = {};
  const actions: Array<{ id: string; file: string }> = [];

  files['openapi-core.json'] = JSON.stringify(openApiDocument(`${i.orgName} ${i.agentName}: you`, 'The signed-in user.', i.appVersion, i.tenantId, corePaths()), null, 2);
  files['ai-plugin-core.json'] = JSON.stringify(
    plugin(i, 'core', 'Who is asking.', 'Get the signed-in user. Call first when the answer depends on who is asking.', [{ name: 'getCurrentUser', description: 'Get the signed-in user (id, name, email, department).' }], 'openapi-core.json'),
    null,
    2
  );
  actions.push({ id: 'coreActions', file: 'ai-plugin-core.json' });

  apps.forEach((a) => {
    const m = MODULES[a.key];
    const d = m.pluginDescription(i.orgName);
    files[`openapi-${a.key}.json`] = JSON.stringify(openApiDocument(`${i.orgName} ${m.label}`, d.human, i.appVersion, i.tenantId, m.paths(a)), null, 2);
    files[`ai-plugin-${a.key}.json`] = JSON.stringify(plugin(i, m.namespace, d.human, d.model, m.functions, `openapi-${a.key}.json`), null, 2);
    actions.push({ id: `${a.key}Actions`, file: `ai-plugin-${a.key}.json` });
  });

  const starters = apps
    .map((a) => MODULES[a.key].starters)
    .reduce((all, s) => all.concat(s), [] as Array<{ title: string; text: string }>)
    .concat(apps.length > 1 ? [{ title: 'My day', text: 'What needs my attention today across my requests, equipment and team?' }] : [])
    .slice(0, STARTERS_LIMIT);
  const siteUrls = apps.map((a) => a.siteUrl).filter((u, idx, all) => all.indexOf(u) === idx);
  const labels = apps.map((a) => MODULES[a.key].label.toLowerCase());

  files['declarativeAgent.json'] = JSON.stringify(
    {
      $schema: 'https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.3/schema.json',
      version: 'v1.3',
      name: i.agentName.substring(0, 100),
      description: `Ask about your ${labels.join(', ')} at ${i.orgName}, in one place.`.substring(0, 1000),
      instructions: buildInstructions(i),
      capabilities: [{ name: 'OneDriveAndSharePoint', items_by_url: siteUrls.map((url) => ({ url })) }, { name: 'People' }],
      conversation_starters: starters,
      actions
    },
    null,
    2
  );

  const semver = (/^\d+\.\d+\.\d+/.exec(i.appVersion) || ['1.0.0'])[0];
  files['manifest.json'] = JSON.stringify(
    {
      $schema: 'https://developer.microsoft.com/json-schemas/teams/v1.19/MicrosoftTeams.schema.json',
      manifestVersion: '1.19',
      version: semver,
      id: i.appId,
      developer: {
        name: i.developerName || i.orgName,
        websiteUrl: i.websiteUrl || apps[0]?.siteUrl || '',
        privacyUrl: i.privacyUrl || apps[0]?.siteUrl || '',
        termsOfUseUrl: i.termsUrl || apps[0]?.siteUrl || ''
      },
      icons: { color: 'color.png', outline: 'outline.png' },
      name: { short: i.agentName.substring(0, 30), full: `${i.orgName} ${i.agentName}`.substring(0, 100) },
      description: {
        short: `Your ${i.orgName} workplace assistant in Copilot.`.substring(0, 80),
        full: `Ask Copilot about your ${labels.join(', ')}. Data stays in ${i.orgName}'s SharePoint and everyone sees only what their SharePoint permissions allow.`.substring(0, 4000)
      },
      accentColor: /^#[0-9a-f]{6}$/i.test(i.brandColor) ? i.brandColor : '#008D62',
      copilotAgents: { declarativeAgents: [{ id: 'brillientaAgent', file: 'declarativeAgent.json' }] },
      permissions: ['identity'],
      validDomains: []
    },
    null,
    2
  );
  return files;
}

/** Problems that would make the package fail to install or behave wrongly. Empty = ready. */
export function validateAgentInput(i: IAgentInput): string[] {
  const p: string[] = [];
  if (!GUID.test(i.appId)) p.push('App id must be a GUID.');
  if (!i.oauthRegistrationId) p.push('OAuth client registration ID is missing (Teams Developer Portal > Tools > OAuth client registration).');
  if (!i.agentName.trim()) p.push('Agent name is required.');
  if (i.agentName.length > 30) p.push('Agent name must be 30 characters or fewer.');
  [i.websiteUrl, i.privacyUrl, i.termsUrl].forEach((u) => {
    if (u && !/^https:\/\//i.test(u)) p.push(`${u} must start with https://`);
  });
  if (!i.apps.length) p.push('No Brillienta app was found. Add the site of at least one app.');
  const seen: Record<string, boolean> = {};
  i.apps.forEach((a) => {
    const m = MODULES[a.key];
    if (seen[a.key]) p.push(`${m.label} is added twice; keep one site.`);
    seen[a.key] = true;
    if (!/^[^,]+,[0-9a-f-]{36},[0-9a-f-]{36}$/i.test(a.graphSiteId)) p.push(`${m.label}: the site id "${a.graphSiteId}" is not a Graph site id.`);
    m.lists.forEach((l) => {
      if (!a.listIds[l.key]) p.push(`${m.label}: list ${a.prefix}${l.suffix} was not found on ${a.siteUrl}.`);
    });
    (m.check ? m.check(a) : []).forEach((x) => p.push(`${m.label}: ${x}`));
  });
  const names = ['getCurrentUser'].concat(...i.apps.map((a) => MODULES[a.key].functions.map((f) => f.name)));
  names.forEach((n, idx) => {
    if (names.indexOf(n) !== idx) p.push(`Two actions are both called ${n}.`);
  });
  if (i.apps.length && buildInstructions(i).length > INSTRUCTIONS_LIMIT) p.push(`The agent instructions are longer than ${INSTRUCTIONS_LIMIT} characters.`);
  return p;
}

export { GRAPH_SCOPES };
