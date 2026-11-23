// Small builders for the OpenAPI descriptions of Microsoft Graph list calls. Same shapes as the Helpdesk agent,
// so behaviour in Copilot is identical.

import { Json } from './types';

/** Delegated Microsoft Graph permissions every Brillienta agent signs in with (same as Brillienta-Helpdesk, so the
 * publisher's existing OAuth client registration works for this agent too). */
export const GRAPH_SCOPES = ['https://graph.microsoft.com/Sites.ReadWrite.All', 'https://graph.microsoft.com/User.Read', 'offline_access'];

export const GRAPH_V1 = 'https://graph.microsoft.com/v1.0';

export function authority(tenantId?: string): string {
  return `https://login.microsoftonline.com/${tenantId || 'organizations'}/oauth2/v2.0`;
}

export const fixed = (name: string, value: string, description?: string): Json => ({
  name,
  in: 'query',
  required: true,
  description: description || 'Always send exactly this value.',
  schema: { type: 'string', enum: [value], default: value }
});

export const filterParam = (description: string, required: boolean = true): Json => ({
  name: '$filter',
  in: 'query',
  required,
  description,
  schema: { type: 'string' }
});

export const topParam: Json = { name: '$top', in: 'query', required: false, description: 'Maximum rows (default 25).', schema: { type: 'integer', default: 25 } };

export const idParam = (description: string): Json => ({ name: 'id', in: 'path', required: true, description, schema: { type: 'string' } });

export const itemSchema: Json = {
  type: 'object',
  properties: {
    id: { type: 'string' },
    createdDateTime: { type: 'string' },
    createdBy: { type: 'object', properties: { user: { type: 'object', properties: { displayName: { type: 'string' }, email: { type: 'string' } } } } },
    fields: { type: 'object', additionalProperties: true }
  }
};

export const listResponse = (description: string): Json => ({
  '200': { description, content: { 'application/json': { schema: { type: 'object', properties: { value: { type: 'array', items: itemSchema } } } } } }
});

export const itemResponse = (description: string): Json => ({
  '200': { description, content: { 'application/json': { schema: itemSchema } } },
  '201': { description, content: { 'application/json': { schema: itemSchema } } }
});

export const body = (description: string, fields: Json, required: string[]): Json => ({
  required: true,
  description,
  content: {
    'application/json': {
      schema: { type: 'object', required: ['fields'], properties: { fields: { type: 'object', required, properties: fields } } }
    }
  }
});

export const itemsPath = (graphSiteId: string, listId: string): string => `/sites/${graphSiteId}/lists/${listId}/items`;

export const expand = (fields: string[]): string => `fields($select=${fields.join(',')})`;

export function listGet(operationId: string, summary: string, expandValue: string, extra: Json[] = []): Json {
  return { get: { operationId, summary, description: summary, parameters: [fixed('$expand', expandValue)].concat(extra), responses: listResponse(summary) } };
}

/** Wraps a module's paths into a full OpenAPI 3 document with the Graph OAuth scheme. */
export function openApiDocument(title: string, description: string, version: string, tenantId: string | undefined, paths: Json): Json {
  return {
    openapi: '3.0.1',
    info: { title, description, version },
    servers: [{ url: GRAPH_V1 }],
    components: {
      securitySchemes: {
        graphDelegated: {
          type: 'oauth2',
          flows: {
            authorizationCode: {
              authorizationUrl: `${authority(tenantId)}/authorize`,
              tokenUrl: `${authority(tenantId)}/token`,
              refreshUrl: `${authority(tenantId)}/token`,
              scopes: GRAPH_SCOPES.reduce(
                (a, s) => ({ ...a, [s]: s === 'offline_access' ? 'Stay signed in' : `Microsoft Graph ${s.split('/').pop()}` }),
                {} as Record<string, string>
              )
            }
          }
        }
      }
    },
    security: [{ graphDelegated: GRAPH_SCOPES }],
    paths
  };
}

export const confirm = (title: string, text: string): Json => ({ confirmation: { type: 'AdaptiveCard', title, body: text } });

/** Adaptive card for list results: a title line and a fact set from item fields. */
export function card(titleExpr: string, facts: Array<[string, string]>): Json {
  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [
      { type: 'TextBlock', text: titleExpr, weight: 'Bolder', wrap: true },
      { type: 'FactSet', facts: facts.map(([title, value]) => ({ title, value })) }
    ]
  };
}
