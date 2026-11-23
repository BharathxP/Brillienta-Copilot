// Brillienta Asset Management in the suite agent. Phase 1 is read-only: who has what, my equipment, my requests.
// Requesting or returning equipment needs Asset Management's own rules (approvals, stock, notifications), so it
// waits for an action queue in that app (phase 2), like Helpdesk's.

import { IAppModule } from '../types';
import { card, expand, filterParam, fixed, idParam, itemResponse, itemsPath, listGet, listResponse, topParam } from '../openapi';

/** Indexed in Asset Management's schema; Graph rejects filters on anything else. */
export const ASSET_FILTER_FIELDS = ['AssetTag', 'SerialNumber', 'AssignedToEmail', 'StatusKind', 'AssetTypeId', 'LocationId'];
export const REQUEST_FILTER_FIELDS = ['RequestNumber', 'RequesterEmail', 'RequestStatus', 'ApprovalStatus', 'AssetId'];

const ASSET = [
  'Title', 'AssetTag', 'AssetTypeName', 'StatusName', 'StatusKind', 'Manufacturer', 'ModelName', 'SerialNumber', 'AssignedToName', 'AssignedToEmail',
  'AssignedAt', 'AckState', 'LocationName', 'Department', 'Condition', 'DeviceName', 'OperatingSystem', 'ComplianceState', 'LastAuditAt'
];
const REQUEST = [
  'Title', 'RequestNumber', 'RequestKind', 'AssetTypeName', 'AssetTag', 'RequesterName', 'RequesterEmail', 'RequestStatus', 'ApprovalStatus',
  'ApproverName', 'Urgency', 'LocationName', 'Created', 'FulfilledAt', 'ResolutionNote'
];

export const assetsModule: IAppModule = {
  key: 'assets',
  label: 'Asset Management',
  defaultPrefix: 'AM_',
  namespace: 'assets',
  lists: [
    { key: 'assets', suffix: 'Assets' },
    { key: 'requests', suffix: 'Requests' },
    { key: 'assetTypes', suffix: 'AssetTypes' }
  ],

  instructions(t) {
    return `## Asset Management (equipment, read-only)
- searchAssets takes an OData $filter on fields, only on ${ASSET_FILTER_FIELDS.join(', ')}. My equipment: fields/AssignedToEmail eq '<my email, lower case>'. Who has a tag: fields/AssetTag eq 'LT-0142'. StatusKind is Available, Assigned, Maintenance, Retired, Disposed or Lost.
- searchAssetRequests: fields/RequesterEmail eq '<email>'; RequestStatus is Submitted, PendingApproval, Approved, Rejected, InProgress, Fulfilled, Cancelled or Closed.
- Show equipment as a table: tag, type, model, status, assigned to. AckState Pending means the person has not confirmed receipt yet.
- You cannot request, return or reassign equipment yet: send the user to the Asset Management portal: ${t.siteUrl}. To report a broken device, raise a helpdesk ticket if Helpdesk is available.`;
  },

  paths(t) {
    const g = t.graphSiteId;
    const l = t.listIds;
    const assets = itemsPath(g, l.assets);
    return {
      [assets]: {
        get: {
          operationId: 'searchAssets',
          summary: 'Search equipment with an OData filter on fields.',
          description: "Filter examples: fields/AssignedToEmail eq 'a@b.com'; fields/AssetTag eq 'LT-0142'.",
          parameters: [filterParam(`OData filter on fields. Only these fields: ${ASSET_FILTER_FIELDS.join(', ')}.`), fixed('$expand', expand(ASSET)), topParam],
          responses: listResponse('Assets')
        }
      },
      [`${assets}/{id}`]: {
        get: {
          operationId: 'getAsset',
          summary: 'One asset with details.',
          description: 'Get an asset by its list item id.',
          parameters: [idParam('Asset item id'), fixed('$expand', expand(ASSET.concat(['PurchaseCost', 'VendorName', 'TagsText'])))],
          responses: itemResponse('Asset')
        }
      },
      [itemsPath(g, l.requests)]: {
        get: {
          operationId: 'searchAssetRequests',
          summary: 'Search equipment requests with an OData filter on fields.',
          description: "Filter example: fields/RequesterEmail eq 'a@b.com'.",
          parameters: [filterParam(`OData filter on fields. Only these fields: ${REQUEST_FILTER_FIELDS.join(', ')}.`), fixed('$expand', expand(REQUEST)), topParam],
          responses: listResponse('Requests')
        }
      },
      [itemsPath(g, l.assetTypes)]: listGet('listAssetTypes', 'Asset types and whether employees can request them.', expand(['Title', 'Details', 'Requestable', 'RequiresApproval', 'IsActive']))
    };
  },

  functions: [
    {
      name: 'searchAssets',
      description: 'Search equipment (laptops, phones, monitors...) by who has it, tag, serial number or status.',
      capabilities: { response_semantics: { data_path: '$.value', properties: { title: '$.fields.Title', subtitle: '$.fields.AssetTag' }, static_template: card('${fields.AssetTag} · ${fields.Title}', [['Type', '${fields.AssetTypeName}'], ['Status', '${fields.StatusName}'], ['Assigned to', '${fields.AssignedToName}']]) } }
    },
    { name: 'getAsset', description: 'Get one asset with details by id.' },
    { name: 'searchAssetRequests', description: 'Search equipment requests by requester or status.' },
    { name: 'listAssetTypes', description: 'List asset types and which ones can be requested.' }
  ],

  pluginDescription: (org) => ({
    human: `Equipment and equipment requests from ${org} Asset Management.`,
    model: `Use for questions about ${org} IT equipment and assets: laptops, phones, monitors, who has which device, asset tags, serial numbers and equipment requests.`
  }),

  starters: [{ title: 'My equipment', text: 'What equipment is assigned to me?' }]
};
